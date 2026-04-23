# ============================================================================
# stt_engine.py
# Faster-Whisper wrapper. Loaded once at server startup; handles audio chunks
# (int16 PCM @ 16 kHz mono) and returns partial + final transcripts.
#
# Key design choices:
#   * Model is loaded ONCE (large-v3 on GPU, medium.en on CPU fallback).
#   * We keep a rolling PCM buffer per session; whenever we have >= 1.0 s of
#     audio we run a transcription pass with `condition_on_previous_text=True`
#     to keep context across chunks.
#   * Voice activity detection (Silero via faster_whisper) suppresses silence
#     so we don't waste GPU cycles on empty mic.
# ============================================================================

from __future__ import annotations

import io
import logging
import os
import threading
import time
from dataclasses import dataclass
from typing import List, Optional

import numpy as np
from faster_whisper import WhisperModel

log = logging.getLogger("stt_engine")


@dataclass
class Transcript:
    text: str
    is_final: bool
    avg_logprob: float
    no_speech_prob: float
    duration_s: float

    @property
    def confidence(self) -> float:
        """Map avg_logprob (approx -1..0) to a 0..1 confidence."""
        score = max(0.0, min(1.0, (self.avg_logprob + 1.0)))
        score *= max(0.0, 1.0 - self.no_speech_prob)
        return round(score, 3)


class FasterWhisperEngine:
    """
    Wraps a single faster-whisper model and serialises access with a lock.
    Thread-safe -- multiple WebSocket sessions can call `transcribe()`.
    """

    def __init__(
        self,
        model_size: str = "large-v3",
        device: str = "auto",
        compute_type: str = "auto",
        models_dir: Optional[str] = None,
    ):
        if device == "auto":
            device = self._detect_device()
        if compute_type == "auto":
            compute_type = "float16" if device == "cuda" else "int8"

        if device == "cpu" and model_size == "large-v3":
            log.warning("CPU detected -- falling back to medium.en (large-v3 too slow on CPU).")
            model_size = "medium.en"

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

    @staticmethod
    def _detect_device() -> str:
        try:
            import ctypes
            ctypes.CDLL("nvcuda.dll" if os.name == "nt" else "libcuda.so")
            return "cuda"
        except OSError:
            return "cpu"

    def transcribe(
        self,
        pcm_int16: bytes,
        initial_prompt: Optional[str] = None,
    ) -> List[Transcript]:
        """
        Run transcription on a raw 16-kHz mono int16 PCM blob.
        Returns a list of Transcript segments (usually 1).
        """
        if not pcm_int16:
            return []

        audio = np.frombuffer(pcm_int16, dtype=np.int16).astype(np.float32) / 32768.0
        duration = len(audio) / self.sample_rate
        if duration < 0.3:
            return []

        with self._lock:
            segments, info = self.model.transcribe(
                audio,
                language="en",
                task="transcribe",
                beam_size=5,
                vad_filter=True,
                vad_parameters=dict(min_silence_duration_ms=300),
                condition_on_previous_text=True,
                initial_prompt=initial_prompt or (
                    "A Christian minister announces Bible references such as "
                    "John 3:16, First Corinthians 13, Psalm twenty three, "
                    "Genesis chapter one verse one."
                ),
                temperature=[0.0, 0.2],
                no_speech_threshold=0.8,
            )
            out: List[Transcript] = []
            for seg in segments:
                out.append(Transcript(
                    text=seg.text.strip(),
                    is_final=True,
                    avg_logprob=seg.avg_logprob,
                    no_speech_prob=seg.no_speech_prob,
                    duration_s=seg.end - seg.start,
                ))
            log.info("DEBUG: segments=%d lang=%s", len(out), info.language)
            return out