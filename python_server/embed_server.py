# ============================================================================
# embed_server.py  --  v7 companion HTTP server for semantic-search embeddings
#
# v6 → v7 changes:
#   * Model upgraded: all-MiniLM-L6-v2 (384d) → BAAI/bge-base-en-v1.5 (768d)
#     bge-base outperforms MiniLM significantly on semantic similarity
#     benchmarks and handles paraphrased scripture much better.
#     Memory: ~430 MB VRAM on CUDA (vs ~90 MB for MiniLM).
#   * Model loaded on CUDA by default when available.
#   * --preload is now the default (True). The embed server is worthless
#     on the first request if it has to load a 430 MB model on demand.
#
# Runs on port 8766 by default.
# C# SemanticSearch calls POST /embed with {"text": "..."}.
# Returns {"embedding": [768 floats], "model": "BAAI/bge-base-en-v1.5"}.
#
# NOTE: After switching models, you MUST rebuild verses.sqlite with the
#       new model. The embedding dimensions change (384 → 768) and old
#       SQLite blobs are incompatible. Run:
#
#         python build_verse_embeddings.py --bible path/to/NASB.xmm \
#             --out path/to/verses.sqlite
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

# BGE-base-en-v1.5: 109M params, 768 dims, significantly better than MiniLM
# for paraphrase / semantic similarity, especially on domain-specific text.
MODEL_NAME = "BAAI/bge-base-en-v1.5"

_model = None


def _detect_device() -> str:
    try:
        import torch
        if torch.cuda.is_available():
            log.info("CUDA available for embedder.")
            return "cuda"
    except Exception:
        pass
    return "cpu"


def _load_model(device: str = "auto"):
    global _model
    if _model is not None:
        return _model
    if device == "auto":
        device = _detect_device()
    from sentence_transformers import SentenceTransformer
    log.info("Loading %s on %s ...", MODEL_NAME, device)
    _model = SentenceTransformer(MODEL_NAME, device=device)
    log.info(
        "Embedder ready. dim=%d device=%s",
        _model.get_sentence_embedding_dimension(),
        device,
    )
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
    return web.json_response({
        "embedding": [float(x) for x in vec.tolist()],
        "model": MODEL_NAME,
        "dim": len(vec),
    })


async def health(_: web.Request) -> web.Response:
    status = "ready" if _model is not None else "loading"
    return web.json_response({"status": status, "model": MODEL_NAME})


def main():
    ap = argparse.ArgumentParser(description="WorshipHelper embed server v7")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8766)
    ap.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    ap.add_argument("--no-preload", action="store_true",
                    help="Skip loading model at startup (loads on first request)")
    args = ap.parse_args()

    # Default: preload at startup so first embed request is instant.
    if not args.no_preload:
        _load_model(args.device)

    app = web.Application()
    app.router.add_post("/embed", embed)
    app.router.add_get("/health", health)
    log.info("Embed server on http://%s:%d/", args.host, args.port)
    web.run_app(app, host=args.host, port=args.port, print=None)


if __name__ == "__main__":
    main()
