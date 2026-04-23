# ============================================================================
# stt_engine.py  --  v6 "VAD-native" Faster-Whisper wrapper
#
# v5 → v6 change:
#   Removed manual RMS energy gate and duration clamp as pre-decode guards.
#   faster-whisper's built-in Silero VAD (vad_filter=True) already does both:
#     * It strips non-speech regions before decoding, so a buffer that is
#       entirely silence produces no segments -- no decode happens at all.
#     * It naturally handles short/long utterances by segmenting the audio
#       into speech chunks before passing them to Whisper.
#   Keeping a redundant manual energy gate on top creates two problems:
#     1. False drops: soft-spoken words near the RMS threshold get silently
#        discarded before the VAD ever sees them.
#     2. Inconsistency: the manual threshold is a different model to the
#        Silero VAD, so they disagree on edge cases.
#   We still keep the post-decode guards (Gates 3-5) because those operate
#   on Whisper's own confidence signals -- no redundancy there.
#
#   The only pre-decode gate we keep is a hard maximum duration check, purely
#   to prevent OOM on runaway audio buffers. This is infeasible for the VAD
#   to guard against (it segments within the buffer, not before accepting it).
# ============================================================================

from __future__ import annotations

import logging
import os
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
NO_SPEECH_MAX   = 0.50   # drop segment if Whisper thinks it's silence
LOGPROB_MIN     = -1.00  # drop segment if avg token log-prob is too low
COMPRESSION_MAX = 2.40   # drop segment if token repetition ratio is too high

# Hard maximum buffer size fed to the model in one call. Purely an OOM guard;
# the VAD will already have chunked real speech well below this.
MAX_DURATION_S  = 20.0

# Dedup window: suppress exact-repeat or single-counter-increment outputs
# (the "John 3 / John 4 / John 5" hallucination loop) within this window.
DEDUP_WINDOW_S  = 6.0


# ---------------------------------------------------------------------------
# Common Whisper hallucinations seen in silence / noise.
# ---------------------------------------------------------------------------
HALLUCINATION_PHRASES = frozenset({
    "thanks for watching", "thank you for watching",
    "thanks for watching!", "thank you.", "thank you",
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
    "amen",
    "god bless", "god bless you",
    "hallelujah", "praise the lord",
})


# ---------------------------------------------------------------------------
# Return type.
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Main engine.
# ---------------------------------------------------------------------------
class FasterWhisperEngine:
    """Thread-safe wrapper around a single faster-whisper model.

    Callers feed it already-endpointed utterances. Whisper's built-in Silero
    VAD handles silence detection and speech segmentation before decoding;
    post-decode quality guards handle hallucinations in the model output.
    """

    def __init__(
        self,
        model_size: str = "small.en",
        device: str = "auto",
        compute_type: str = "auto",
        models_dir: Optional[str] = None,
        hotwords: Optional[Sequence[str]] = None,
        # VAD tuning -- exposed here so callers can adjust for their
        # acoustic environment without touching the transcribe() call.
        vad_threshold: float = 0.5,
        vad_min_silence_ms: int = 500,
        vad_speech_pad_ms: int = 200,
    ):
        if device == "auto":
            device = self._detect_device()
        if compute_type == "auto":
            compute_type = "float16" if device == "cuda" else "int8"

        if device == "cpu" and model_size in ("large-v3", "large-v2", "large"):
            log.warning(
                "CPU detected -- falling back from %s to small.en "
                "(large models are too slow and hallucinate more on short utterances).",
                model_size,
            )
            model_size = "small.en"

        log.info("Loading Faster-Whisper model=%s device=%s compute=%s",
                 model_size, device, compute_type)
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

        # VAD parameters stored so transcribe() can pass them consistently.
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
        try:
            import ctypes
            ctypes.CDLL("nvcuda.dll" if os.name == "nt" else "libcuda.so")
            return "cuda"
        except OSError:
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
        tokens_a = a.split()
        tokens_b = b.split()
        if (
            len(tokens_a) == len(tokens_b) == 2
            and tokens_a[0] == tokens_b[0]
            and tokens_a[1].isdigit() and tokens_b[1].isdigit()
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
        """Transcribe one already-endpointed utterance.

        The only pre-decode guard is a hard duration cap to prevent OOM on
        runaway buffers. Everything else -- silence detection, short-utterance
        filtering, speech segmentation -- is handled by Whisper's built-in
        Silero VAD (vad_filter=True).
        """
        if not pcm_int16:
            return []

        audio = np.frombuffer(pcm_int16, dtype=np.int16).astype(np.float32) / 32768.0
        duration = len(audio) / self.sample_rate

        # Only pre-decode gate: hard OOM guard on buffer length.
        # The VAD handles silence; this only catches pathological inputs.
        if duration > MAX_DURATION_S:
            log.debug(
                "clamp: buffer %.2fs > MAX_DURATION_S=%.1fs -- trimming tail",
                duration, MAX_DURATION_S,
            )
            audio = audio[-int(MAX_DURATION_S * self.sample_rate):]
            duration = MAX_DURATION_S

        beam = 1 if self.device == "cpu" else 5

        with self._lock:
            segments, _info = self.model.transcribe(
                audio,
                language="en",
                task="transcribe",

                # Decoding
                beam_size=beam,
                best_of=beam,
                patience=1.0,
                temperature=[0.0, 0.2, 0.4, 0.6, 0.8],
                compression_ratio_threshold=COMPRESSION_MAX,
                log_prob_threshold=LOGPROB_MIN,
                no_speech_threshold=NO_SPEECH_MAX,

                # No cross-utterance carry-over (the primary hallucination
                # loop driver in v4).
                condition_on_previous_text=False,

                # Whisper's own hallucination guard -- flags silence-padded
                # outputs that Silero VAD missed.
                hallucination_silence_threshold=2.0,

                # Built-in VAD: Silero segments the audio before decoding,
                # discarding non-speech regions entirely. This replaces the
                # manual RMS energy gate and MIN_DURATION_S guard from v5.
                vad_filter=True,
                vad_parameters=self._vad_params,

                # Biasing
                initial_prompt=initial_prompt,
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

            # Gate 1: Whisper's own per-segment quality signals.
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

            # Gate 2: common-phrase blacklist.
            if self._is_hallucination_phrase(txt):
                t.reason_dropped = "hallucination_phrase"
                log.debug("drop: %s text=%r", t.reason_dropped, txt)
                continue

            # Gate 3: near-duplicate suppression.
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
                "all %d segment(s) dropped (dur=%.2fs) -- likely silence/hallucination",
                len(raw_segments), duration,
            )
        elif out:
            log.info(
                "ok: %s (dur=%.2fs conf=%.2f logprob=%.2f cr=%.2f)",
                out[0].text, duration, out[0].confidence,
                out[0].avg_logprob, out[0].compression_ratio,
            )

        return out

    # ------------------------------------------------------------------
    def reset_context(self) -> None:
        """Clear dedup state on session disconnect or explicit reset."""
        with self._lock:
            self._last_text = ""
            self._last_emit_ts = 0.0