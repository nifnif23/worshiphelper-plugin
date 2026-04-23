# ============================================================================
# vad.py  --  Voice-activity detection + utterance aggregator
#
# Purpose:
#   The v4 server transcribed every 1.5s of audio whether or not it contained
#   speech. That was the main structural cause of hallucination loops -- the
#   model was constantly asked "what's in this room tone?" and it would make
#   something up.
#
#   This module replaces that wall-clock windowing with proper endpointing.
#   Audio bytes stream in at whatever cadence the client feels like, and
#   we buffer until we detect:
#       * voiced frames (> VAD_SPEECH_MIN_MS of continuous speech)
#       * THEN a trailing silence of >= VAD_END_SILENCE_MS
#   At that point we emit exactly one utterance to the transcriber.
#
# Why Silero and not WebRTC-VAD?
#   WebRTC-VAD is CPU-cheap but prone to false-positives on steady hums
#   (projector fans, HVAC). Silero is a small (~1.8MB) neural VAD that's
#   rock-solid on music-room environments. It's the same VAD faster-whisper
#   uses internally, so we get a consistent answer end-to-end.
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

# Speech probability threshold. Silero returns 0..1; 0.5 is the canonical
# break-point. We use 0.55 so HVAC/projector-fan noise doesn't trip us.
SPEECH_PROB_THRESHOLD    = 0.55

# How long voiced audio must be continuously seen before we commit to "yes,
# this is an utterance, keep buffering".
VAD_SPEECH_MIN_MS        = 160     # ~5 frames

# Once we're in an utterance, how long of trailing silence ends it.
# 700ms is a good compromise:
#   * short enough that "John 3 verse seven" feels instant
#   * long enough to not chop "Psalm 23 ... verse four"
VAD_END_SILENCE_MS       = 700

# Hard maximum on a single utterance. Prevents a stuck-open door from
# growing the buffer forever (the transcribe call will clamp anyway,
# but this saves RAM and keeps latency bounded).
MAX_UTTERANCE_MS         = 18_000

# Pre-speech pad -- keep this many ms of audio BEFORE the first voiced frame
# so we don't clip the speaker's "L-" on "Let's turn to".
PRE_SPEECH_PAD_MS        = 250

# Post-speech pad -- include this much trailing silence after the endpoint
# so Whisper's own VAD is happy.
POST_SPEECH_PAD_MS       = 200


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
        # The silero-vad package publishes its weights via torch.hub. We pin
        # onnx=False because ONNX runtime pulls extra deps on Windows.
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
        # Silero's forward returns a 1-element tensor for a 512-sample chunk.
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
                yield utt        # <-- ready to transcribe

    Thread-safety: NOT thread safe. One aggregator per session.
    """

    def __init__(self, vad: Optional[SileroVAD] = None):
        self._vad = vad or SileroVAD()

        # Byte-level input buffer -- consumed in 32ms frames.
        self._pending: bytearray = bytearray()

        # Rolling pre-speech pad (int16 bytes).
        pad_frames = max(1, PRE_SPEECH_PAD_MS // FRAME_MS)
        self._pre_pad: Deque[bytes] = deque(maxlen=pad_frames)

        # State machine
        self._in_utterance: bool = False
        self._utterance_bytes: bytearray = bytearray()
        self._utterance_ms: int = 0
        self._voiced_run_ms: int = 0        # how much speech we've seen while entering
        self._silent_run_ms: int = 0        # how much silence since last voice
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
        """
        Consume raw PCM bytes. Returns zero or more complete utterances.
        """
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
        # Compute speech probability and RMS.
        pcm_i16 = np.frombuffer(frame_bytes, dtype=np.int16)
        pcm_f32 = pcm_i16.astype(np.float32) / 32768.0
        rms = float(np.sqrt(np.mean(pcm_f32 * pcm_f32))) if pcm_f32.size else 0.0

        try:
            p_speech = self._vad.speech_prob(pcm_f32)
        except Exception as exc:   # noqa: BLE001
            log.warning("Silero VAD failed (%s) -- falling back to RMS gate.", exc)
            p_speech = 1.0 if rms > 0.012 else 0.0

        is_speech = p_speech >= SPEECH_PROB_THRESHOLD

        if not self._in_utterance:
            # We're waiting for an utterance to start.
            self._pre_pad.append(frame_bytes)
            if is_speech:
                self._voiced_run_ms += FRAME_MS
                if self._voiced_run_ms >= VAD_SPEECH_MIN_MS:
                    # Commit: start an utterance. Seed it with the pre-pad
                    # so we don't lose the opening consonant.
                    self._in_utterance = True
                    self._utterance_bytes = bytearray(b"".join(self._pre_pad))
                    self._utterance_ms = len(self._utterance_bytes) // 2 * 1000 // SAMPLE_RATE
                    self._silent_run_ms = 0
                    self._peak_rms = rms
                    self._pre_pad.clear()
            else:
                self._voiced_run_ms = 0
            return None

        # We're inside an utterance.
        self._utterance_bytes.extend(frame_bytes)
        self._utterance_ms += FRAME_MS
        self._peak_rms = max(self._peak_rms, rms)

        if is_speech:
            self._silent_run_ms = 0
        else:
            self._silent_run_ms += FRAME_MS

        # Endpoint?
        if self._silent_run_ms >= VAD_END_SILENCE_MS:
            return self._close_utterance(reason="silence")

        # Hard cap.
        if self._utterance_ms >= MAX_UTTERANCE_MS:
            return self._close_utterance(reason="max_duration")

        return None

    # ------------------------------------------------------------------
    def _close_utterance(self, reason: str) -> Utterance:
        # Tack on a small trailing pad of silence so Whisper sees the end.
        pad_bytes = PRE_SPEECH_PAD_MS // FRAME_MS * FRAME_BYTES  # reuse size
        self._utterance_bytes.extend(b"\x00" * pad_bytes)

        utt = Utterance(
            pcm=bytes(self._utterance_bytes),
            duration_s=len(self._utterance_bytes) / 2 / SAMPLE_RATE,
            peak_rms=self._peak_rms,
        )
        log.debug("endpoint: reason=%s dur=%.2fs peak_rms=%.3f",
                  reason, utt.duration_s, utt.peak_rms)

        # Reset state but keep the VAD hot (speaker's still here).
        self._in_utterance = False
        self._utterance_bytes = bytearray()
        self._utterance_ms = 0
        self._voiced_run_ms = 0
        self._silent_run_ms = 0
        self._peak_rms = 0.0
        self._pre_pad.clear()
        return utt
