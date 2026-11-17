"""
view_production_collection.py — pipelines/sistematika/

Baca isi koleksi ChromaDB PRODUKSI dan dump ke JSON tanpa ingest/embed ulang.
Melaporkan ruang jarak, model embed, norma vektor (harus ~1.0), dan jumlah
chunk per chunk_type.

Pakai:
  python view_production_collection.py --out production_snapshot.json
"""

import sys, os, json, argparse
from collections import Counter

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import numpy as np

from config import PATHS
from vector_store import get_collection, _collection_space


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--chroma-dir", default=PATHS["chroma_dir"])
    ap.add_argument("--collection", default=PATHS["chroma_collection"])
    ap.add_argument("--out", default="production_snapshot.json")
    args = ap.parse_args()

    print(f"Membaca koleksi '{args.collection}' di {args.chroma_dir} ...")
    try:
        # read_only: nama salah -> error, bukan koleksi kosong baru.
        # check_space=False: koleksi lama (L2) tetap bisa dibuka.
        collection = get_collection(
            args.chroma_dir, args.collection, check_space=False, read_only=True
        )
    except Exception as e:
        print(f"Koleksi tidak bisa dibuka: {e}")
        return 1
    stored = collection.get(include=["documents", "metadatas", "embeddings"])

    meta = collection.metadata or {}
    space = _collection_space(collection)
    model = meta.get("embed_model", "tidak tercatat")
    embs = stored.get("embeddings")
    norm_info = None
    if embs is not None and len(embs) > 0:
        norms = np.linalg.norm(np.asarray(embs), axis=1)
        norm_info = {
            "min": float(norms.min()),
            "max": float(norms.max()),
            "mean": float(norms.mean()),
        }
    types = Counter(m.get("chunk_type", "(tanpa chunk_type)") for m in stored["metadatas"])
    print(f"  ruang jarak : {space}")
    print(f"  model embed : {model}")
    print(f"  norma vektor: {norm_info}")
    print(f"  chunk_type  : {dict(types)}")
    if space not in ("cosine", "unknown"):
        print("  !! koleksi ini BUKAN cosine -> ingest ulang dengan --reset")
    if "(tanpa chunk_type)" in types:
        print("  !! ada chunk tanpa chunk_type -> ingest ulang dengan --reset")

    dump = {
        "collection_name": args.collection,
        "chroma_dir": args.chroma_dir,
        "space": space,
        "embed_model": model,
        "vector_norm": norm_info,
        "chunk_types": dict(types),
        "total_stored": len(stored["ids"]),
        "items": [
            {
                "id": stored["ids"][i],
                "document_text": stored["documents"][i],
                "metadata": stored["metadatas"][i],
                "embedding_dim": len(embs[i]) if embs is not None else None,
            }
            for i in range(len(stored["ids"]))
        ],
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"Total {dump['total_stored']} chunk ditemukan. Ditulis -> {args.out}")


if __name__ == "__main__":
    sys.exit(main())
