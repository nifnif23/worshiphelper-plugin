# ============================================================================
# server.py  --  v5 WorshipHelper STT server
#
# Architecture:
#
#   Client -- binary PCM (int16 mono 16 kHz) --> Server
#                                         |
#                                         v
#                              UtteranceAggregator (VAD)
#                                         |
#                                 on endpoint: complete utterance
#                                         |
#                                         v
#                              FasterWhisperEngine.transcribe()
#                                         |
#                              guardrails (logprob, compression, blacklist)
#                                         |
#                                         v
#   Server <-- JSON transcript messages  <-- .
#
# Protocol (mostly backwards-compatible with v4):
#   Client -> Server
#     <binary PCM chunk>
#     {"type":"flush"}                    -- force endpoint now
#     {"type":"reset"}                    -- wipe VAD + dedup state
#     {"type":"ping"}
#     {"type":"hotwords","words":["John","Romans",...]}  -- optional per-session bias
#
#   Server -> Client
#     {"type":"status","message":"ready","model":"small.en","device":"cpu"}
#     {"type":"partial","text":"...","confidence":0.7}   -- (reserved for future)
#     {"type":"transcript","text":"John 3 verse seven","confidence":0.91,
#        "final":true,"duration":1.8,"avg_logprob":-0.3,
#        "compression_ratio":1.2,"no_speech_prob":0.02}
#     {"type":"dropped","reason":"hallucination_phrase","text":"Thanks for watching"}
#     {"type":"error","message":"..."}
#     {"type":"pong"}
#
# Endpoints:
#   ws://127.0.0.1:8765/        -- streaming STT
#   GET http://127.0.0.1:8765/health  -> {"status":"ok","utterances":42,...}
#   POST http://127.0.0.1:8765/embed  -> {"embedding": [...]}   (for SemanticSearch)
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

from stt_engine import FasterWhisperEngine
from vad import SileroVAD, UtteranceAggregator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
log = logging.getLogger("server")

ENGINE: Optional[FasterWhisperEngine] = None
SHARED_VAD: Optional[SileroVAD] = None

# Metrics (atomic-ish; single-threaded event loop updates).
_metrics = {
    "connections": 0,
    "active": 0,
    "utterances_total": 0,
    "utterances_emitted": 0,
    "utterances_dropped": 0,
}

# Default hotwords: the 66 canonical book names, so the decoder knows those
# are real English words even when the domain model isn't. Kept as a *list*
# so C# can update it per-session.
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
    "chapter", "verse",
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
        "protocol": "v5",
    }))

    try:
        async for msg in ws:
            # --- Audio (binary) -------------------------------------------
            if isinstance(msg, (bytes, bytearray)):
                try:
                    utterances = aggregator.feed(bytes(msg))
                except Exception as exc:   # noqa: BLE001
                    log.exception("aggregator.feed failed: %s", exc)
                    continue

                for utt in utterances:
                    _metrics["utterances_total"] += 1
                    # Run Whisper on a thread so we don't stall the event loop.
                    # Each engine call is serialised internally by its own lock.
                    ENGINE.set_hotwords(per_session_hotwords)
                    segments = await asyncio.to_thread(ENGINE.transcribe, utt.pcm)

                    if not segments:
                        _metrics["utterances_dropped"] += 1
                        # Tell the client so the debug panel can show it.
                        await _safe_send(ws, {
                            "type": "dropped",
                            "reason": "engine_guardrails",
                            "duration": utt.duration_s,
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
                            "duration": seg.duration_s,
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
                    segments = await asyncio.to_thread(ENGINE.transcribe, utt.pcm)
                    for seg in segments:
                        await _safe_send(ws, {
                            "type": "transcript",
                            "text": seg.text,
                            "confidence": seg.confidence,
                            "final": True,
                            "duration": seg.duration_s,
                        })
            elif ctype == "reset":
                aggregator.reset()
                ENGINE.reset_context()
                await _safe_send(ws, {"type": "status", "message": "reset"})
            elif ctype == "hotwords":
                words = ctrl.get("words") or []
                if isinstance(words, list):
                    # Merge with defaults so we always have the canonical books.
                    merged = list(DEFAULT_HOTWORDS)
                    for w in words:
                        if isinstance(w, str) and w and w not in merged:
                            merged.append(w)
                    per_session_hotwords = merged
                    log.debug("hotwords updated: +%d -> %d",
                              len(words), len(merged))
            elif ctype == "ping":
                await _safe_send(ws, {"type": "pong", "t": time.time()})
            elif ctype == "config":
                # Reserved for future per-session configuration (thresholds, etc.)
                pass

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
    except Exception as exc:   # noqa: BLE001
        log.debug("send failed: %s", exc)


# ---------------------------------------------------------------------------
# Simple HTTP side-channel for /health and /embed.
# ---------------------------------------------------------------------------
_EMBEDDER = None


def _get_embedder():
    global _EMBEDDER
    if _EMBEDDER is None:
        from sentence_transformers import SentenceTransformer
        _EMBEDDER = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")
    return _EMBEDDER


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
    # /embed is handled by a companion FastAPI / aiohttp process; the
    # websockets library's process_request hook can't easily read a POST
    # body, so we just advertise 404 here. See embed_server.py (companion).
    return None


# ---------------------------------------------------------------------------
async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host",       default="127.0.0.1")
    parser.add_argument("--port",       type=int, default=8765)
    parser.add_argument("--model",      default="small.en",
                        help="faster-whisper size: tiny.en / base.en / small.en / "
                             "medium.en / large-v3 / distil-small.en / distil-large-v3")
    parser.add_argument("--device",     default="auto",
                        choices=["auto", "cpu", "cuda"])
    parser.add_argument("--compute",    default="auto",
                        help="float32 / float16 / int8 / int8_float16 / auto")
    parser.add_argument("--models-dir", default=None)
    args = parser.parse_args()

    global ENGINE, SHARED_VAD
    ENGINE = FasterWhisperEngine(
        model_size=args.model,
        device=args.device,
        compute_type=args.compute,
        models_dir=args.models_dir,
    )
    SHARED_VAD = SileroVAD()
    # Pre-load VAD weights so the first utterance isn't slow.
    SHARED_VAD._ensure_loaded()

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, stop.set)
        except NotImplementedError:
            pass   # Windows

    async with websockets.serve(
        handle_session,
        args.host,
        args.port,
        max_size=8 * 1024 * 1024,
        ping_interval=20,
        ping_timeout=20,
        process_request=process_request,
    ):
        log.info("WorshipHelper STT server v5 listening on ws://%s:%d/",
                 args.host, args.port)
        log.info("Health check: http://%s:%d/health", args.host, args.port)
        await stop.wait()
        log.info("Shutting down.")


if __name__ == "__main__":
    asyncio.run(main())
