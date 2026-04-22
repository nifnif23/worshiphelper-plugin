# ============================================================================
# build_verse_embeddings.py
# One-shot offline job: read the OpenSong XMM Bible, embed every verse with
# sentence-transformers, and write a SQLite DB the C# side reads.
#
# Schema:
#   CREATE TABLE verses (
#     id INTEGER PRIMARY KEY,
#     book TEXT NOT NULL,
#     chapter INTEGER NOT NULL,
#     verse INTEGER NOT NULL,
#     text TEXT NOT NULL,
#     embedding BLOB NOT NULL        -- float32 little-endian, dim=384
#   );
#   CREATE INDEX idx_ref ON verses(book, chapter, verse);
# ============================================================================

import argparse
import sqlite3
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
from sentence_transformers import SentenceTransformer


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
    ap = argparse.ArgumentParser()
    ap.add_argument("--bible", required=True, help="Path to OpenSong XMM bible")
    ap.add_argument("--out",   required=True, help="Output SQLite path")
    ap.add_argument("--model", default="sentence-transformers/all-MiniLM-L6-v2")
    args = ap.parse_args()

    print(f"Loading embedder: {args.model}")
    model = SentenceTransformer(args.model)
    dim = model.get_sentence_embedding_dimension()
    print(f"Embedding dim: {dim}")

    verses = list(iter_verses(Path(args.bible)))
    print(f"Read {len(verses):,} verses from {args.bible}")

    texts = [t for (_, _, _, t) in verses]
    print("Encoding (this may take a few minutes)...")
    embs = model.encode(texts, batch_size=64, show_progress_bar=True,
                        normalize_embeddings=True)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()

    conn = sqlite3.connect(out)
    conn.execute("""
        CREATE TABLE verses (
            id INTEGER PRIMARY KEY,
            book TEXT NOT NULL,
            chapter INTEGER NOT NULL,
            verse INTEGER NOT NULL,
            text TEXT NOT NULL,
            embedding BLOB NOT NULL
        );""")
    conn.execute("CREATE INDEX idx_ref ON verses(book, chapter, verse);")
    conn.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);")
    conn.execute("INSERT INTO meta VALUES ('dim', ?)", (str(dim),))
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
    print(f"Wrote {out} ({out.stat().st_size/1e6:.1f} MB)")


if __name__ == "__main__":
    main()
