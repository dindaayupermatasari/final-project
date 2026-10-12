"""
ingest_guideline.py — pipelines/sistematika/

CLI: parser -> chunker -> vector_store. Dijalankan sekali di awal atau
tiap panduan direvisi.

Perubahan: embedding dilakukan SEBELUM koleksi lama dihapus (--reset
diteruskan ke upsert_chunks), jadi kegagalan Ollama tidak mengosongkan
koleksi produksi. Chunk template ("Bagian 1/2/3") dibuang; contoh disimpan
dengan chunk_type="example" dan bisa difilter saat query.

Pakai:
  pip install pymupdf chromadb requests --break-system-packages
  ollama pull nomic-embed-text:v1.5

  python ingest_guideline.py --pdf panduan.pdf --chroma-dir ./chroma_store \
                             --collection panduan_sistematika --reset [--dry-run]
"""

import sys, os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import argparse
from collections import Counter

from config import PATHS
from parser import extract_text_by_page, split_into_sections
from chunker import filter_section, build_chunks
from vector_store import upsert_chunks, build_embed_text, EMBED_MODEL


def _preview(c: dict, n: int = 250) -> str:
    t = build_embed_text(c["text"], c["metadata"])
    return t[:n].replace("\n", " | ") + ("..." if len(t) > n else "")


def main():
    ap = argparse.ArgumentParser(
        description="Ingest panduan penulisan PENS ke ChromaDB"
    )
    ap.add_argument("--pdf", default=PATHS["guideline_pdf"])
    ap.add_argument("--chroma-dir", default=PATHS["chroma_dir"])
    ap.add_argument("--collection", default=PATHS["chroma_collection"])
    ap.add_argument(
        "--reset",
        action="store_true",
        help="Ganti koleksi lama (wajib saat ganti model/metrik). Koleksi baru "
        "baru menimpa yang lama SETELAH semua embedding berhasil.",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="Print hasil chunking (termasuk yang dibuang) tanpa embed/simpan",
    )
    args = ap.parse_args()

    print(f"Membaca {args.pdf} ...")
    full_text = extract_text_by_page(args.pdf)

    print("Memecah jadi BAB > sub-bab ...")
    sections = split_into_sections(full_text)
    print(f"  ditemukan {len(sections)} sub-bab")

    print("Menyaring kalimat format visual dari tiap sub-bab ...")
    for sec in sections:
        filter_section(sec)

    all_chunks = build_chunks(sections, include_templates=True)
    chunks = [c for c in all_chunks if c["metadata"]["chunk_type"] != "template"]
    dropped = [c for c in all_chunks if c["metadata"]["chunk_type"] == "template"]
    counts = Counter(c["metadata"]["chunk_type"] for c in chunks)
    print(
        f"  {len(chunks)} chunk siap diproses {dict(counts)}, "
        f"{len(dropped)} chunk template dibuang"
    )

    if args.dry_run:
        for c in all_chunks:
            m = c["metadata"]
            tag = "DIBUANG" if m["chunk_type"] == "template" else m["chunk_type"]
            print(f"\n[{c['id']}] ({tag}) {m['subbab_title']}")
            print("  (di-embed) " + _preview(c))
        return

    print(f"Embedding ({EMBED_MODEL}) & menyimpan ke ChromaDB ...")
    n = upsert_chunks(chunks, args.chroma_dir, args.collection, reset=args.reset)
    print(f"Selesai. {n} chunk masuk ke koleksi '{args.collection}' di {args.chroma_dir}")


if __name__ == "__main__":
    sys.exit(main())
