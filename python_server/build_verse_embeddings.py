# ============================================================================
# build_verse_embeddings.py  --  v7
#
# One-shot offline job: read OpenSong XMM Bible, embed every verse with
# sentence-transformers, write a SQLite DB the C# side reads at startup.
#
# v6 → v7:
#   * Default model changed: all-MiniLM-L6-v2 → BAAI/bge-base-en-v1.5
#     Embedding dim changes from 384 → 768. If you already have a verses.sqlite
#     built with MiniLM, delete it and re-run — the dim mismatch will cause
#     SemanticSearch to silently no-op in C#.
#   * --device arg added: model runs on CUDA when available (RTX 3050 encodes
#     31k verses in ~60s vs ~15 min on CPU).
#   * Batch size increased: 256 on GPU, 64 on CPU.
#
# Usage:
#   python build_verse_embeddings.py \
#       --bible "..\WorshipHelperVSTO\data\Bibles\NASB.xmm" \
#       --out   "..\WorshipHelperVSTO\data\verses.sqlite"
#
# Schema:
#   CREATE TABLE verses (
#     id INTEGER PRIMARY KEY,
#     book TEXT NOT NULL, chapter INTEGER NOT NULL, verse INTEGER NOT NULL,
#     text TEXT NOT NULL,
#     embedding BLOB NOT NULL    -- float32 little-endian, dim=768
#   );
#   CREATE INDEX idx_ref ON verses(book, chapter, verse);
#   CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
#     -- 'dim'   : embedding dimension as string
#     -- 'model' : model name used to build this DB
# ============================================================================

import argparse
import sqlite3
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
from sentence_transformers import SentenceTransformer

DEFAULT_MODEL = "BAAI/bge-base-en-v1.5"


def _detect_device() -> str:
    try:
        import torch
        if torch.cuda.is_available():
            print(f"CUDA available: {torch.cuda.get_device_name(0)}")
            return "cuda"
    except Exception:
        pass
    print("No CUDA — using CPU (this will take ~15 min).")
    return "cpu"


def iter_verses(xmm_path: Path):
    tree = ET.parse(xmm_path)
    root = tree.getroot()
    for b in root.iter("b"):
        book = b.get("n")
        for c in b.iter("c"):
            chapter = int(c.get("n"))
            for v in c.iter("v"):
                verse = int(v.get("n"))
                text = (v.text or "").strip()
                if text:
                    yield book, chapter, verse, text


def main():
    ap = argparse.ArgumentParser(description="Build WorshipHelper verse embedding DB")
    ap.add_argument("--bible", required=True, help="Path to OpenSong XMM bible")
    ap.add_argument("--out",   required=True, help="Output SQLite path")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"sentence-transformers model (default: {DEFAULT_MODEL})")
    ap.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    args = ap.parse_args()

    device = _detect_device() if args.device == "auto" else args.device

    print(f"Loading embedder: {args.model} on {device}")
    model = SentenceTransformer(args.model, device=device)
    dim = model.get_sentence_embedding_dimension()
    print(f"Embedding dim: {dim}")

    verses = list(iter_verses(Path(args.bible)))
    print(f"Read {len(verses):,} verses from {args.bible}")

    texts = [t for (_, _, _, t) in verses]
    batch_size = 256 if device == "cuda" else 64
    print(f"Encoding {len(texts):,} verses (batch={batch_size}, this may take a few minutes)...")
    embs = model.encode(
        texts,
        batch_size=batch_size,
        show_progress_bar=True,
        normalize_embeddings=True,  # cosine = dot product on normalised vectors
    )

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()

    conn = sqlite3.connect(out)
    conn.execute("""
        CREATE TABLE verses (
            id       INTEGER PRIMARY KEY,
            book     TEXT    NOT NULL,
            chapter  INTEGER NOT NULL,
            verse    INTEGER NOT NULL,
            text     TEXT    NOT NULL,
            embedding BLOB   NOT NULL
        );""")
    conn.execute("CREATE INDEX idx_ref ON verses(book, chapter, verse);")
    conn.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);")
    conn.execute("INSERT INTO meta VALUES ('dim',   ?)", (str(dim),))
    conn.execute("INSERT INTO meta VALUES ('model', ?)", (args.model,))

    rows = []
    for (book, ch, vs, text), emb in zip(verses, embs):
        blob = emb.astype(np.float32).tobytes()
        rows.append((book, ch, vs, text, blob))

    conn.executemany(
        "INSERT INTO verses(book,chapter,verse,text,embedding) VALUES (?,?,?,?,?)",
        rows,
    )
    conn.commit()
    conn.close()

    size_mb = out.stat().st_size / 1e6
    print(f"\nDone. Wrote {out}  ({size_mb:.1f} MB, {dim}-dim embeddings)")
    print(
        "\nNOTE: The C# SemanticSearch.cs checks the 'dim' field in meta.\n"
        "If you changed models, the old verses.sqlite is incompatible — "
        "this run replaced it correctly."
    )


if __name__ == "__main__":
    main()
