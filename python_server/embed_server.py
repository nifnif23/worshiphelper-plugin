# ============================================================================
# embed_server.py  --  Companion HTTP server for semantic-search embeddings.
#
# Runs on port 8766 by default. The C# SemanticSearch class calls POST /embed
# with {"text": "..."} and gets back {"embedding": [384 floats]}.
#
# Kept separate from the WebSocket STT server because:
#   * The websockets package's process_request hook can't read POST bodies.
#   * Keeps the STT loop hot and CPU-bound work on the embed thread.
#
# Minimal dependencies -- just aiohttp.
# ============================================================================
from __future__ import annotations

import argparse
import asyncio
import logging

from aiohttp import web

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
log = logging.getLogger("embed_server")

_model = None


def _load_model():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        log.info("Loading sentence-transformers/all-MiniLM-L6-v2 ...")
        _model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")
        log.info("Model loaded. dim=%d", _model.get_sentence_embedding_dimension())
    return _model


async def embed(request: web.Request) -> web.Response:
    try:
        data = await request.json()
    except Exception:
        return web.json_response({"error": "invalid_json"}, status=400)

    text = (data or {}).get("text")
    if not isinstance(text, str) or not text.strip():
        return web.json_response({"error": "text_required"}, status=400)

    model = _load_model()
    vec = model.encode([text], normalize_embeddings=True)[0]
    return web.json_response({"embedding": [float(x) for x in vec.tolist()]})


async def health(_: web.Request) -> web.Response:
    return web.json_response({"status": "ok"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8766)
    ap.add_argument("--preload", action="store_true",
                    help="Load the embedder at startup instead of on first request")
    args = ap.parse_args()

    if args.preload:
        _load_model()

    app = web.Application()
    app.router.add_post("/embed", embed)
    app.router.add_get("/health", health)
    log.info("Embed server listening on http://%s:%d/", args.host, args.port)
    web.run_app(app, host=args.host, port=args.port, print=None)


if __name__ == "__main__":
    main()
