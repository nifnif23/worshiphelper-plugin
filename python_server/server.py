# ============================================================================
# server.py  --  v7 WorshipHelper STT server
#
# v6 → v7 changes:
#   * Default model: small.en → distil-large-v3
#     (runs at ~medium speed on RTX 30-series, large-v3 accuracy for English)
#   * Default compute: auto → float16 on CUDA, int8 on CPU
#   * BIBLE_INITIAL_PROMPT is now passed to every engine.transcribe() call
#     so Whisper always has Bible vocabulary context.
#   * VAD parameters exposed as CLI args so tuning doesn't require code edits.
#   * Embed endpoint wired into /health so the C# side can check both.
#
# Protocol (unchanged from v5/v6):
#   Client → Server
#     <binary PCM chunk>
#     {"type":"flush"}
#     {"type":"reset"}
#     {"type":"ping"}
#     {"type":"hotwords","words":["John","Romans",...]}
#
#   Server → Client
#     {"type":"status","message":"ready","model":"...","device":"..."}
#     {"type":"transcript","text":"...","confidence":0.91,"final":true,...}
#     {"type":"dropped","reason":"...","duration":1.2}
#     {"type":"error","message":"..."}
#     {"type":"pong"}
#
# Endpoints:
#   ws://127.0.0.1:8765/        streaming STT
#   GET http://127.0.0.1:8765/health
# ============================================================================

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import signal
import time
from typing import Optional

import websockets
from websockets.server import WebSocketServerProtocol

from stt_engine import BIBLE_INITIAL_PROMPT, FasterWhisperEngine
from vad import SileroVAD, UtteranceAggregator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
log = logging.getLogger("server")

ENGINE: Optional[FasterWhisperEngine] = None
SHARED_VAD: Optional[SileroVAD] = None

_metrics = {
    "connections": 0,
    "active": 0,
    "utterances_total": 0,
    "utterances_emitted": 0,
    "utterances_dropped": 0,
}

# 66 canonical book names sent to Whisper as hotwords so the decoder weights
# them as real English words. Per-session custom words are merged on top.
DEFAULT_HOTWORDS = [
    "Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy",
    "Joshua", "Judges", "Ruth", "Samuel", "Kings", "Chronicles",
    "Ezra", "Nehemiah", "Esther", "Job", "Psalm", "Psalms",
    "Proverbs", "Ecclesiastes", "Isaiah", "Jeremiah",
    "Lamentations", "Ezekiel", "Daniel", "Hosea", "Joel", "Amos",
    "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah",
    "Haggai", "Zechariah", "Malachi",
    "Matthew", "Mark", "Luke", "John", "Acts", "Romans",
    "Corinthians", "Galatians", "Ephesians", "Philippians",
    "Colossians", "Thessalonians", "Timothy", "Titus", "Philemon",
    "Hebrews", "James", "Peter", "Jude", "Revelation",
    "chapter", "verse", "scripture",
]


# ---------------------------------------------------------------------------
async def handle_session(ws: WebSocketServerProtocol) -> None:
    peer = ws.remote_address
    _metrics["connections"] += 1
    _metrics["active"] += 1
    log.info("Client connected: %s (active=%d)", peer, _metrics["active"])

    aggregator = UtteranceAggregator(vad=SHARED_VAD)
    per_session_hotwords = list(DEFAULT_HOTWORDS)

    await ws.send(json.dumps({
        "type": "status", "message": "ready",
        "model": ENGINE.model_size, "device": ENGINE.device,
        "protocol": "v7",
    }))

    try:
        async for msg in ws:
            # --- Audio (binary) -------------------------------------------
            if isinstance(msg, (bytes, bytearray)):
                try:
                    utterances = aggregator.feed(bytes(msg))
                except Exception as exc:
                    log.exception("aggregator.feed failed: %s", exc)
                    continue

                for utt in utterances:
                    _metrics["utterances_total"] += 1
                    ENGINE.set_hotwords(per_session_hotwords)

                    # Pass BIBLE_INITIAL_PROMPT so every call gets domain context.
                    segments = await asyncio.to_thread(
                        ENGINE.transcribe, utt.pcm, BIBLE_INITIAL_PROMPT
                    )

                    if not segments:
                        _metrics["utterances_dropped"] += 1
                        await _safe_send(ws, {
                            "type": "dropped",
                            "reason": "engine_guardrails",
                            "duration": round(utt.duration_s, 2),
                            "peak_rms": round(utt.peak_rms, 4),
                        })
                        continue

                    for seg in segments:
                        _metrics["utterances_emitted"] += 1
                        await _safe_send(ws, {
                            "type": "transcript",
                            "text": seg.text,
                            "confidence": seg.confidence,
                            "final": True,
                            "duration": round(seg.duration_s, 2),
                            "avg_logprob": round(seg.avg_logprob, 3),
                            "compression_ratio": round(seg.compression_ratio, 3),
                            "no_speech_prob": round(seg.no_speech_prob, 3),
                        })
                continue

            # --- Control (text) -------------------------------------------
            try:
                ctrl = json.loads(msg)
            except json.JSONDecodeError:
                continue
            if not isinstance(ctrl, dict):
                continue

            ctype = ctrl.get("type")
            if ctype == "flush":
                utt = aggregator.flush()
                if utt:
                    segments = await asyncio.to_thread(
                        ENGINE.transcribe, utt.pcm, BIBLE_INITIAL_PROMPT
                    )
                    for seg in segments:
                        await _safe_send(ws, {
                            "type": "transcript",
                            "text": seg.text,
                            "confidence": seg.confidence,
                            "final": True,
                            "duration": round(seg.duration_s, 2),
                        })
            elif ctype == "reset":
                aggregator.reset()
                ENGINE.reset_context()
                await _safe_send(ws, {"type": "status", "message": "reset"})
            elif ctype == "hotwords":
                words = ctrl.get("words") or []
                if isinstance(words, list):
                    merged = list(DEFAULT_HOTWORDS)
                    for w in words:
                        if isinstance(w, str) and w and w not in merged:
                            merged.append(w)
                    per_session_hotwords = merged
                    log.debug("hotwords updated: +%d → %d", len(words), len(merged))
            elif ctype == "ping":
                await _safe_send(ws, {"type": "pong", "t": time.time()})

    except websockets.ConnectionClosed:
        pass
    except Exception as exc:
        log.exception("Session error: %s", exc)
        await _safe_send(ws, {"type": "error", "message": str(exc)})
    finally:
        _metrics["active"] -= 1
        log.info("Client disconnected: %s (active=%d)", peer, _metrics["active"])


async def _safe_send(ws, payload: dict) -> None:
    try:
        await ws.send(json.dumps(payload))
    except Exception as exc:
        log.debug("send failed: %s", exc)


# ---------------------------------------------------------------------------
async def process_request(path, headers):
    if path == "/health":
        body = json.dumps({
            "status": "ok",
            "model": ENGINE.model_size if ENGINE else None,
            "device": ENGINE.device if ENGINE else None,
            **_metrics,
        }).encode()
        return (
            200,
            [("Content-Type", "application/json"),
             ("Content-Length", str(len(body)))],
            body,
        )
    return None


# ---------------------------------------------------------------------------
async def main():
    parser = argparse.ArgumentParser(description="WorshipHelper STT server v7")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument(
        "--model", default="distil-large-v3",
        help=(
            "faster-whisper model size. Options by VRAM usage on RTX 30-series:\n"
            "  distil-large-v3 (~1.5 GB float16) — RECOMMENDED for RTX 3050+\n"
            "  medium.en       (~0.6 GB float16) — faster, slightly less accurate\n"
            "  large-v3        (~3.0 GB float16) — best quality, needs 4GB+ VRAM\n"
            "  small.en        (~0.2 GB)         — CPU fallback only"
        ),
    )
    parser.add_argument("--device",  default="auto", choices=["auto", "cpu", "cuda"])
    parser.add_argument("--compute", default="auto",
                        help="float16 / int8_float16 / int8 / auto")
    parser.add_argument("--models-dir", default=None)
    # VAD tuning flags (church-optimised defaults already in vad.py)
    parser.add_argument("--vad-threshold", type=float, default=0.45,
                        help="Silero speech probability threshold (0-1, default 0.45)")
    parser.add_argument("--vad-silence-ms", type=int, default=800,
                        help="Trailing silence before utterance endpoint (ms, default 800)")
    parser.add_argument("--vad-pad-ms", type=int, default=300,
                        help="Speech padding added around detected speech (ms, default 300)")
    args = parser.parse_args()

    global ENGINE, SHARED_VAD
    ENGINE = FasterWhisperEngine(
        model_size=args.model,
        device=args.device,
        compute_type=args.compute,
        models_dir=args.models_dir,
        vad_threshold=args.vad_threshold,
        vad_min_silence_ms=args.vad_silence_ms,
        vad_speech_pad_ms=args.vad_pad_ms,
    )
    SHARED_VAD = SileroVAD()
    SHARED_VAD._ensure_loaded()

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, stop.set)
        except NotImplementedError:
            pass  # Windows

    async with websockets.serve(
        handle_session,
        args.host,
        args.port,
        max_size=8 * 1024 * 1024,
        ping_interval=20,
        ping_timeout=20,
        process_request=process_request,
    ):
        log.info(
            "WorshipHelper STT server v7 on ws://%s:%d/  model=%s device=%s",
            args.host, args.port, args.model, args.device,
        )
        log.info("Health: http://%s:%d/health", args.host, args.port)
        await stop.wait()
        log.info("Shutting down.")


if __name__ == "__main__":
    asyncio.run(main())
