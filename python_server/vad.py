# ============================================================================
# vad.py  --  Voice-activity detection + utterance aggregator
#
# v7 church-environment tuning:
#
#   SPEECH_PROB_THRESHOLD  0.55 → 0.45
#     The old threshold was cutting soft-spoken phrases (common in reflective
#     preaching styles) and any speech against a constant background
#     (organ, congregation murmur, HVAC). 0.45 matches what Silero
#     recommends for challenging acoustic environments.
#
#   VAD_END_SILENCE_MS  700 → 800
#     Preachers in the UK/African church tradition often have longer,
#     more deliberate pauses mid-sentence ("Let's turn to... John three
#     sixteen"). The old 700ms was chopping those into separate utterances
#     which confused the detector into thinking they were unrelated phrases.
#
#   PRE_SPEECH_PAD_MS  250 → 300
#     Extra 50ms pre-pad catches the plosive onset of "P-salm" and "Ph-ilippians"
#     which are commonly the leading consonant that gets cut.
#
#   VAD_SPEECH_MIN_MS  160 → 120
#     Lowered slightly so short-but-confident references ("John 3:16")
#     don't get rejected at the entry gate.
# ============================================================================

from __future__ import annotations

import logging
from collections import deque
from dataclasses import dataclass
from typing import Deque, List, Optional

import numpy as np

log = logging.getLogger("vad")

# ---------------------------------------------------------------------------
# Tunables
# ---------------------------------------------------------------------------
SAMPLE_RATE              = 16_000
FRAME_MS                 = 32          # Silero natively operates on 32ms frames
FRAME_SAMPLES            = SAMPLE_RATE * FRAME_MS // 1000  # 512
FRAME_BYTES              = FRAME_SAMPLES * 2               # int16

# Church-tuned: lower than 0.55 default to handle background noise without
# missing soft-spoken scripture references.
SPEECH_PROB_THRESHOLD    = 0.45

# Shorter minimum — short references like "John 3" should not be gate-killed
# at the entry check.
VAD_SPEECH_MIN_MS        = 120     # ~4 frames

# Wider end-silence to handle deliberate pauses in preaching cadence.
# "Let us read... [pause]... Psalm 23... [pause]... verse 4."
VAD_END_SILENCE_MS       = 800

MAX_UTTERANCE_MS         = 18_000

# Wider pre-speech pad: catches plosive onsets ("P-salm", "B-ook").
PRE_SPEECH_PAD_MS        = 300

# Trailing pad — gives Whisper's VAD a clean ending boundary.
POST_SPEECH_PAD_MS       = 250


# ---------------------------------------------------------------------------
@dataclass
class Utterance:
    pcm: bytes                       # int16 mono 16kHz
    duration_s: float
    peak_rms: float                  # peak frame RMS in float32 units


# ---------------------------------------------------------------------------
class SileroVAD:
    """Thin wrapper around torch-hub Silero. Lazy-loaded on first use."""

    def __init__(self):
        self._model = None
        self._utils = None

    def _ensure_loaded(self) -> None:
        if self._model is not None:
            return
        import torch
        torch.set_num_threads(1)
        log.info("Loading Silero VAD...")
        model, utils = torch.hub.load(
            repo_or_dir="snakers4/silero-vad",
            model="silero_vad",
            trust_repo=True,
            onnx=False,
        )
        self._model = model
        self._utils = utils
        self._torch = torch
        log.info("Silero VAD loaded.")

    def speech_prob(self, frame: np.ndarray) -> float:
        """frame: float32 mono, 512 samples at 16kHz. Returns P(speech) 0..1."""
        self._ensure_loaded()
        t = self._torch.from_numpy(frame).float()
        with self._torch.no_grad():
            p = self._model(t, SAMPLE_RATE).item()
        return float(p)

    def reset(self) -> None:
        if self._model is not None and hasattr(self._model, "reset_states"):
            self._model.reset_states()


# ---------------------------------------------------------------------------
class UtteranceAggregator:
    """
    Feed PCM bytes in; get complete utterances out.

    Typical usage:
        agg = UtteranceAggregator()
        for pcm_bytes in stream:
            for utt in agg.feed(pcm_bytes):
                yield utt        # ready to transcribe

    Thread-safety: NOT thread safe. One aggregator per session.
    """

    def __init__(self, vad: Optional[SileroVAD] = None):
        self._vad = vad or SileroVAD()

        self._pending: bytearray = bytearray()

        pad_frames = max(1, PRE_SPEECH_PAD_MS // FRAME_MS)
        self._pre_pad: Deque[bytes] = deque(maxlen=pad_frames)

        self._in_utterance: bool = False
        self._utterance_bytes: bytearray = bytearray()
        self._utterance_ms: int = 0
        self._voiced_run_ms: int = 0
        self._silent_run_ms: int = 0
        self._peak_rms: float = 0.0

    # ------------------------------------------------------------------
    def reset(self) -> None:
        self._pending.clear()
        self._pre_pad.clear()
        self._in_utterance = False
        self._utterance_bytes = bytearray()
        self._utterance_ms = 0
        self._voiced_run_ms = 0
        self._silent_run_ms = 0
        self._peak_rms = 0.0
        self._vad.reset()

    # ------------------------------------------------------------------
    def feed(self, pcm_int16: bytes) -> List[Utterance]:
        """Consume raw PCM bytes. Returns zero or more complete utterances."""
        if not pcm_int16:
            return []
        self._pending.extend(pcm_int16)

        out: List[Utterance] = []
        while len(self._pending) >= FRAME_BYTES:
            frame_bytes = bytes(self._pending[:FRAME_BYTES])
            del self._pending[:FRAME_BYTES]
            done = self._process_frame(frame_bytes)
            if done is not None:
                out.append(done)
        return out

    # ------------------------------------------------------------------
    def flush(self) -> Optional[Utterance]:
        """Force-close any open utterance (e.g. on client disconnect)."""
        if not self._in_utterance:
            return None
        return self._close_utterance(reason="flush")

    # ------------------------------------------------------------------
    def _process_frame(self, frame_bytes: bytes) -> Optional[Utterance]:
        pcm_i16 = np.frombuffer(frame_bytes, dtype=np.int16)
        pcm_f32 = pcm_i16.astype(np.float32) / 32768.0
        rms = float(np.sqrt(np.mean(pcm_f32 * pcm_f32))) if pcm_f32.size else 0.0

        try:
            p_speech = self._vad.speech_prob(pcm_f32)
        except Exception as exc:  # noqa: BLE001
            log.warning("Silero VAD failed (%s) — falling back to RMS gate.", exc)
            # Slightly looser RMS fallback to match the relaxed VAD threshold.
            p_speech = 1.0 if rms > 0.010 else 0.0

        is_speech = p_speech >= SPEECH_PROB_THRESHOLD

        if not self._in_utterance:
            self._pre_pad.append(frame_bytes)
            if is_speech:
                self._voiced_run_ms += FRAME_MS
                if self._voiced_run_ms >= VAD_SPEECH_MIN_MS:
                    # Commit: start an utterance with pre-pad so we don't
                    # clip the opening consonant.
                    self._in_utterance = True
                    self._utterance_bytes = bytearray(b"".join(self._pre_pad))
                    self._utterance_ms = (
                        len(self._utterance_bytes) // 2 * 1000 // SAMPLE_RATE
                    )
                    self._silent_run_ms = 0
                    self._peak_rms = rms
                    self._pre_pad.clear()
            else:
                self._voiced_run_ms = 0
            return None

        # Inside an utterance.
        self._utterance_bytes.extend(frame_bytes)
        self._utterance_ms += FRAME_MS
        self._peak_rms = max(self._peak_rms, rms)

        if is_speech:
            self._silent_run_ms = 0
        else:
            self._silent_run_ms += FRAME_MS

        if self._silent_run_ms >= VAD_END_SILENCE_MS:
            return self._close_utterance(reason="silence")

        if self._utterance_ms >= MAX_UTTERANCE_MS:
            return self._close_utterance(reason="max_duration")

        return None

    # ------------------------------------------------------------------
    def _close_utterance(self, reason: str) -> Utterance:
        pad_bytes = POST_SPEECH_PAD_MS // FRAME_MS * FRAME_BYTES
        self._utterance_bytes.extend(b"\x00" * pad_bytes)

        utt = Utterance(
            pcm=bytes(self._utterance_bytes),
            duration_s=len(self._utterance_bytes) / 2 / SAMPLE_RATE,
            peak_rms=self._peak_rms,
        )
        log.debug(
            "endpoint: reason=%s dur=%.2fs peak_rms=%.3f",
            reason, utt.duration_s, utt.peak_rms,
        )

        self._in_utterance = False
        self._utterance_bytes = bytearray()
        self._utterance_ms = 0
        self._voiced_run_ms = 0
        self._silent_run_ms = 0
        self._peak_rms = 0.0
        self._pre_pad.clear()
        return utt
