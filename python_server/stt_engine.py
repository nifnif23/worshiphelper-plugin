# ============================================================================
# stt_engine.py  --  v7 "church-tuned" Faster-Whisper wrapper
#
# v6 → v7 changes:
#
#   1. CUDA detection rewritten.
#      Old: ctypes.CDLL("nvcuda.dll") — silently fails on many Windows setups
#           (driver path issues, WDDM vs TCC, etc.), forcing CPU even when a
#           GPU is present. This was the primary cause of bad output.
#      New: torch.cuda.is_available() — the canonical check, same one
#           faster-whisper uses internally.
#
#   2. Default model changed: small.en → distil-large-v3.
#      distil-large-v3 runs at ~medium speed on an RTX 3050 (float16, ~1.5 GB
#      VRAM) but delivers accuracy close to large-v3. On CPU we keep small.en
#      because large models are unusable at CPU speed.
#
#   3. Church words removed from the hallucination blacklist.
#      "amen", "hallelujah", "god bless", "praise the lord" are things a
#      preacher actually says — often immediately before citing a verse.
#      Silencing them was causing legitimate utterances to be dropped.
#      The small.en model hallucinated these on silence; distil-large-v3
#      does not need the workaround.
#
#   4. LOGPROB_MIN relaxed: -1.00 → -1.20.
#      West African / Caribbean / South Asian preachers (common in UK
#      churches) produce lower avg_logprob on English-only models. The old
#      threshold was binning real speech as low-confidence.
#
#   5. NO_SPEECH_MAX relaxed: 0.50 → 0.60 to match the new VAD threshold.
#
#   6. Bible initial_prompt added.
#      Tells Whisper it is listening to scripture readings so book names,
#      "chapter", "verse", and ordinal numbers are decoded correctly.
#      This is the single biggest accuracy improvement for explicit references.
#
#   7. compute_type defaults to float16 on CUDA (not "auto").
#      CTranslate2's "auto" picks int8 on some driver versions, which cuts
#      accuracy noticeably. float16 is the right choice for RTX 30-series.
# ============================================================================

from __future__ import annotations

import logging
import re
import threading
import time
from dataclasses import dataclass
from typing import List, Optional, Sequence

import numpy as np
from faster_whisper import WhisperModel

log = logging.getLogger("stt_engine")

# ---------------------------------------------------------------------------
# Post-decode quality thresholds (Whisper's own signals).
# ---------------------------------------------------------------------------
# Relaxed from v6 to handle accented English preachers (lower logprob) and
# church background noise (higher no_speech on borderline frames).
NO_SPEECH_MAX   = 0.60   # drop if Whisper is highly confident it heard nothing
LOGPROB_MIN     = -1.20  # drop if token quality is very low (relaxed for accents)
COMPRESSION_MAX = 2.40   # drop if output has heavy token repetition

MAX_DURATION_S  = 20.0   # OOM guard — VAD already chunks real speech far below this
DEDUP_WINDOW_S  = 6.0    # suppress exact-repeat / counter-increment outputs

# ---------------------------------------------------------------------------
# Bible-domain initial prompt.
# Primes the Whisper decoder with the vocabulary and patterns it will hear.
# This dramatically improves book-name transcription and ordinal number
# handling (e.g. "twenty-three" vs "23" vs "twenty three").
# ---------------------------------------------------------------------------
BIBLE_INITIAL_PROMPT = (
    "Scripture reading. Bible verse reference. "
    "Genesis, Exodus, Leviticus, Numbers, Deuteronomy, Joshua, Judges, Ruth, "
    "First Samuel, Second Samuel, First Kings, Second Kings, "
    "First Chronicles, Second Chronicles, Ezra, Nehemiah, Esther, Job, Psalms, "
    "Proverbs, Ecclesiastes, Song of Solomon, Isaiah, Jeremiah, Lamentations, "
    "Ezekiel, Daniel, Hosea, Joel, Amos, Obadiah, Jonah, Micah, Nahum, "
    "Habakkuk, Zephaniah, Haggai, Zechariah, Malachi, "
    "Matthew, Mark, Luke, John, Acts, Romans, "
    "First Corinthians, Second Corinthians, Galatians, Ephesians, Philippians, "
    "Colossians, First Thessalonians, Second Thessalonians, "
    "First Timothy, Second Timothy, Titus, Philemon, Hebrews, James, "
    "First Peter, Second Peter, First John, Second John, Third John, Jude, "
    "Revelation. "
    "Chapter verse. John three sixteen. Romans eight twenty-eight. "
    "Psalm twenty-three verse one. Genesis one one."
)

# ---------------------------------------------------------------------------
# Hallucination phrases: common Whisper outputs on silence / room tone.
# NOTE: Church-specific words (amen, hallelujah, etc.) are intentionally
# ABSENT. A preacher saying "Amen" before citing a verse is real speech.
# The distil-large-v3 model does not hallucinate these on silence the way
# small.en did. If you switch back to small.en you may want to re-add them.
# ---------------------------------------------------------------------------
HALLUCINATION_PHRASES = frozenset({
    "thanks for watching", "thank you for watching",
    "thanks for watching!", "thank you.",
    "please subscribe", "subscribe to my channel",
    "you", "bye", "bye.", "bye bye",
    ".", "...", ". . .",
    "music", "[music]", "(music)",
    "silence", "[silence]",
    "transcribed by", "transcription by",
    "please like and subscribe",
    "see you next time", "see you in the next video",
    "don't forget to subscribe",
    "hmm", "hmm.", "um", "uh",
    "yeah", "okay", "ok",
})


@dataclass
class Transcript:
    text: str
    is_final: bool
    avg_logprob: float
    no_speech_prob: float
    compression_ratio: float
    duration_s: float
    reason_dropped: Optional[str] = None

    @property
    def confidence(self) -> float:
        score = max(0.0, min(1.0, self.avg_logprob + 1.0))
        score *= max(0.0, 1.0 - self.no_speech_prob)
        return round(score, 3)


class FasterWhisperEngine:
    """Thread-safe wrapper around a single faster-whisper model.

    Callers feed it already-endpointed utterances (from UtteranceAggregator).
    Whisper's built-in Silero VAD handles silence/segmentation; post-decode
    quality gates handle hallucinations in the model output.
    """

    def __init__(
        self,
        model_size: str = "distil-large-v3",
        device: str = "auto",
        compute_type: str = "auto",
        models_dir: Optional[str] = None,
        hotwords: Optional[Sequence[str]] = None,
        vad_threshold: float = 0.45,
        vad_min_silence_ms: int = 600,
        vad_speech_pad_ms: int = 300,
    ):
        if device == "auto":
            device = self._detect_device()
        if compute_type == "auto":
            # float16 is the right choice for RTX 30-series.
            # int8_float16 is a fallback for tight VRAM budgets.
            compute_type = "float16" if device == "cuda" else "int8"

        if device == "cpu" and model_size not in (
            "tiny", "tiny.en", "base", "base.en", "small", "small.en",
            "distil-small.en",
        ):
            log.warning(
                "CPU detected — falling back from %s to small.en "
                "(large models are unusable at CPU speed).",
                model_size,
            )
            model_size = "small.en"

        log.info(
            "Loading Faster-Whisper model=%s device=%s compute=%s",
            model_size, device, compute_type,
        )
        t0 = time.time()
        self.model = WhisperModel(
            model_size,
            device=device,
            compute_type=compute_type,
            download_root=models_dir,
        )
        log.info("Model loaded in %.1fs.", time.time() - t0)

        self._lock = threading.Lock()
        self.sample_rate = 16_000
        self.device = device
        self.model_size = model_size

        self._vad_params = dict(
            threshold=vad_threshold,
            min_silence_duration_ms=vad_min_silence_ms,
            speech_pad_ms=vad_speech_pad_ms,
        )

        self._last_text: str = ""
        self._last_emit_ts: float = 0.0
        self._hotwords = list(hotwords) if hotwords else []

    # ------------------------------------------------------------------
    @staticmethod
    def _detect_device() -> str:
        """Use torch's canonical CUDA check — not ctypes, not assumptions."""
        try:
            import torch
            if torch.cuda.is_available():
                name = torch.cuda.get_device_name(0)
                log.info("CUDA available: %s", name)
                return "cuda"
        except Exception as exc:
            log.debug("torch CUDA check failed (%s), trying ctypes fallback.", exc)
        # Ctypes fallback for environments where torch is not installed yet.
        try:
            import ctypes
            import os
            lib = "nvcuda.dll" if os.name == "nt" else "libcuda.so.1"
            ctypes.CDLL(lib)
            log.info("CUDA available (ctypes fallback).")
            return "cuda"
        except OSError:
            pass
        log.info("No CUDA found — using CPU.")
        return "cpu"

    @staticmethod
    def _normalise(text: str) -> str:
        t = text.lower().strip()
        t = re.sub(r"[^a-z0-9\s:]", " ", t)
        t = re.sub(r"\s+", " ", t).strip()
        return t

    @classmethod
    def _is_hallucination_phrase(cls, text: str) -> bool:
        n = cls._normalise(text)
        if not n or len(n) <= 2:
            return True
        return n in HALLUCINATION_PHRASES

    def _is_near_duplicate(self, text: str, now_ts: float) -> bool:
        if not self._last_text:
            return False
        if now_ts - self._last_emit_ts > DEDUP_WINDOW_S:
            return False
        a = self._normalise(text)
        b = self._last_text
        if a == b:
            return True
        # Suppress "John 3" → "John 4" → "John 5" counter-increment loops.
        tokens_a = a.split()
        tokens_b = b.split()
        if (
            len(tokens_a) == len(tokens_b) == 2
            and tokens_a[0] == tokens_b[0]
            and tokens_a[1].isdigit() and tokens_b[1].isdigit()
            and abs(int(tokens_a[1]) - int(tokens_b[1])) <= 2
        ):
            return True
        return False

    # ------------------------------------------------------------------
    def set_hotwords(self, words: Sequence[str]) -> None:
        with self._lock:
            self._hotwords = list(words or [])

    # ------------------------------------------------------------------
    def transcribe(
        self,
        pcm_int16: bytes,
        initial_prompt: Optional[str] = None,
    ) -> List[Transcript]:
        """Transcribe one already-endpointed utterance."""
        if not pcm_int16:
            return []

        audio = np.frombuffer(pcm_int16, dtype=np.int16).astype(np.float32) / 32768.0
        duration = len(audio) / self.sample_rate

        if duration > MAX_DURATION_S:
            log.debug("clamp: buffer %.2fs > %.1fs — trimming", duration, MAX_DURATION_S)
            audio = audio[-int(MAX_DURATION_S * self.sample_rate):]
            duration = MAX_DURATION_S

        # On GPU: beam_size=5 for best accuracy.
        # On CPU: beam_size=1 to keep latency bearable.
        beam = 5 if self.device == "cuda" else 1

        prompt = initial_prompt or BIBLE_INITIAL_PROMPT

        with self._lock:
            segments, _info = self.model.transcribe(
                audio,
                language="en",
                task="transcribe",

                # Decoding quality
                beam_size=beam,
                best_of=beam,
                patience=1.0,
                temperature=[0.0, 0.2, 0.4, 0.6, 0.8],
                compression_ratio_threshold=COMPRESSION_MAX,
                log_prob_threshold=LOGPROB_MIN,
                no_speech_threshold=NO_SPEECH_MAX,

                # No cross-utterance carry-over (primary hallucination loop driver).
                condition_on_previous_text=False,

                # Whisper's own silence guard — catches what Silero VAD misses.
                hallucination_silence_threshold=2.0,

                # Built-in Silero VAD — strips non-speech before decoding.
                vad_filter=True,
                vad_parameters=self._vad_params,

                # Domain bias: prime the decoder with Bible vocabulary.
                initial_prompt=prompt,
                hotwords=" ".join(self._hotwords) if self._hotwords else None,
            )
            raw_segments = list(segments)

        now_ts = time.time()
        out: List[Transcript] = []

        for seg in raw_segments:
            txt = (seg.text or "").strip()

            t = Transcript(
                text=txt,
                is_final=True,
                avg_logprob=float(seg.avg_logprob),
                no_speech_prob=float(seg.no_speech_prob),
                compression_ratio=float(seg.compression_ratio),
                duration_s=float(seg.end - seg.start),
            )

            if t.no_speech_prob > NO_SPEECH_MAX:
                t.reason_dropped = f"no_speech_prob={t.no_speech_prob:.2f}"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue
            if t.avg_logprob < LOGPROB_MIN:
                t.reason_dropped = f"avg_logprob={t.avg_logprob:.2f}"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue
            if t.compression_ratio > COMPRESSION_MAX:
                t.reason_dropped = f"compression_ratio={t.compression_ratio:.2f}"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue
            if self._is_hallucination_phrase(txt):
                t.reason_dropped = "hallucination_phrase"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue
            if self._is_near_duplicate(txt, now_ts):
                t.reason_dropped = "dedup"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue

            with self._lock:
                self._last_text = self._normalise(txt)
                self._last_emit_ts = now_ts
            out.append(t)

        if raw_segments and not out:
            log.info(
                "all %d segment(s) dropped (dur=%.2fs) — silence/hallucination",
                len(raw_segments), duration,
            )
        elif out:
            log.info(
                "ok: %s (dur=%.2fs conf=%.2f logprob=%.2f cr=%.2f)",
                out[0].text, duration, out[0].confidence,
                out[0].avg_logprob, out[0].compression_ratio,
            )

        return out

    def reset_context(self) -> None:
        """Clear dedup state on session disconnect or explicit reset."""
        with self._lock:
            self._last_text = ""
            self._last_emit_ts = 0.0
