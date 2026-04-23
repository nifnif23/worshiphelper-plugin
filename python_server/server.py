# ============================================================================
# server.py  --  Faster-Whisper WebSocket server
#
# Protocol (JSON text frames + binary audio frames):
#
#   Client -- binary frame (int16 PCM @ 16 kHz mono) --> Server   # audio chunk
#   Client -- {"type":"flush"}                        --> Server   # force transcribe
#   Client -- {"type":"reset"}                        --> Server   # clear context
#
#   Server -- {"type":"transcript","text":"...","confidence":0.83,"final":true}
#   Server -- {"type":"status","message":"ready"}
#   Server -- {"type":"error","message":"..."}
#
# Endpoints:
#   ws://127.0.0.1:8765/stt      # streaming STT
#   http://127.0.0.1:8765/health # GET -> {"status":"ok"}
# ============================================================================

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import signal
from collections import deque
from typing import Deque

import websockets
from websockets.server import WebSocketServerProtocol

from stt_engine import FasterWhisperEngine

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
log = logging.getLogger("server")

WINDOW_MS = 1500
SAMPLE_RATE = 16_000
BYTES_PER_WINDOW = int(SAMPLE_RATE * (WINDOW_MS / 1000.0)) * 2  # int16 = 2 bytes

ENGINE: FasterWhisperEngine | None = None

# Lazy-load the sentence-transformer (same model as build_verse_embeddings)
_embedder = None


def _get_embedder():
    global _embedder
    if _embedder is None:
        from sentence_transformers import SentenceTransformer
        _embedder = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")
    return _embedder


async def handle_session(ws: WebSocketServerProtocol) -> None:
    """One client = one session = one rolling buffer."""
    peer = ws.remote_address
    log.info("Client connected: %s", peer)

    buf: Deque[bytes] = deque()
    buf_bytes = 0
    last_sent = ""

    await ws.send(json.dumps({"type": "status", "message": "ready",
                              "model": ENGINE.model_size, "device": ENGINE.device}))

    try:
        async for msg in ws:
            if isinstance(msg, (bytes, bytearray)):
                buf.append(bytes(msg))
                buf_bytes += len(msg)

                if buf_bytes >= BYTES_PER_WINDOW:
                    pcm = b"".join(buf)
                    buf.clear()
                    buf_bytes = 0

                    segments = await asyncio.to_thread(ENGINE.transcribe, pcm)
                    for seg in segments:
                        if not seg.text:
                            continue
                        last_sent = seg.text
                        await ws.send(json.dumps({
                            "type": "transcript",
                            "text": seg.text,
                            "confidence": seg.confidence,
                            "final": True,
                            "duration": seg.duration_s,
                        }))
                continue

            try:
                ctrl = json.loads(msg)
            except json.JSONDecodeError:
                continue

            if ctrl.get("type") == "flush":
                if buf_bytes > 0:
                    pcm = b"".join(buf); buf.clear(); buf_bytes = 0
                    segments = await asyncio.to_thread(ENGINE.transcribe, pcm)
                    for seg in segments:
                        await ws.send(json.dumps({
                            "type": "transcript", "text": seg.text,
                            "confidence": seg.confidence, "final": True,
                            "duration": seg.duration_s,
                        }))
            elif ctrl.get("type") == "reset":
                buf.clear(); buf_bytes = 0; last_sent = ""
                await ws.send(json.dumps({"type": "status", "message": "reset"}))
            elif ctrl.get("type") == "ping":
                await ws.send(json.dumps({"type": "pong"}))

    except websockets.ConnectionClosed:
        pass
    except Exception as exc:
        log.exception("Session error: %s", exc)
        try:
            await ws.send(json.dumps({"type": "error", "message": str(exc)}))
        except Exception:
            pass
    finally:
        log.info("Client disconnected: %s", peer)


async def process_request(path, headers):
    """Tiny HTTP responder for liveness probes and embed endpoint."""
    if path == "/health":
        body = b'{"status":"ok"}'
        return 200, [("Content-Type", "application/json"),
                     ("Content-Length", str(len(body)))], body
    if path == "/embed":
        # Read body from headers (workaround: pass text as query param for simplicity)
        # For production, use a separate aiohttp/FastAPI process on port 8766.
        pass
    return None


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--model", default="large-v3")
    parser.add_argument("--device", default="auto")
    parser.add_argument("--models-dir", default=None)
    args = parser.parse_args()

    global ENGINE
    ENGINE = FasterWhisperEngine(
        model_size=args.model,
        device=args.device,
        models_dir=args.models_dir,
    )

    stop = asyncio.Event()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            asyncio.get_running_loop().add_signal_handler(sig, stop.set)
        except NotImplementedError:
            pass  # Windows

    async with websockets.serve(
        handle_session,
        args.host,
        args.port,
        max_size=4 * 1024 * 1024,
        ping_interval=20,
        ping_timeout=20,
        process_request=process_request,
    ):
        log.info("WorshipHelper STT server listening on ws://%s:%d/", args.host, args.port)
        await stop.wait()
        log.info("Shutting down.")


if __name__ == "__main__":
    asyncio.run(main())
