"""
ingest_guideline.py — pipelines/sistematika/

Script CLI (dijalankan sekali di awal, atau tiap panduan direvisi) yang
mengorkestrasi: parser -> chunker -> vector_store. Tidak ada logic berat
di sini -- semua ada di modul masing-masing supaya bisa dites & dipakai
ulang terpisah.

Pakai:
  pip install pymupdf chromadb requests --break-system-packages
  ollama pull nomic-embed-text

  python ingest_guideline.py --pdf panduan.pdf --chroma-dir ./chroma_store \
                             --collection panduan_sistematika [--dry-run]
"""

import sys, os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import argparse

from config import PATHS
from parser import extract_text_by_page, split_into_sections
from chunker import filter_section, build_chunks
from vector_store import upsert_chunks


def main():
    ap = argparse.ArgumentParser(
        description="Ingest panduan penulisan PENS ke ChromaDB"
    )
    ap.add_argument(
        "--pdf",
        default=PATHS["guideline_pdf"],
        help="Default diambil dari config.PATHS['guideline_pdf']",
    )
    ap.add_argument("--chroma-dir", default=PATHS["chroma_dir"])
    ap.add_argument("--collection", default=PATHS["chroma_collection"])
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="Cuma print hasil chunking, tanpa embed/simpan",
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

    chunks = build_chunks(sections)
    print(f"  {len(chunks)} chunk siap diproses")

    if args.dry_run:
        for c in chunks:
            print(f"\n[{c['id']}] {c['metadata']['subbab_title']}")
            print(c["text"][:200] + ("..." if len(c["text"]) > 200 else ""))
        return

    print("Embedding & menyimpan ke ChromaDB ...")
    n = upsert_chunks(chunks, args.chroma_dir, args.collection)
    print(
        f"Selesai. {n} chunk masuk ke koleksi '{args.collection}' di {args.chroma_dir}"
    )


if __name__ == "__main__":
    sys.exit(main())
