"""
inspect_pipeline.py — pipelines/sistematika/

Debug tool: jalanin pipeline RAG (parsing -> chunking -> embedding ->
vector store) TAHAP PER TAHAP, dan tiap tahap nge-dump output-nya ke file
JSON terpisah supaya bisa diperiksa manual. Tidak mengubah parser.py,
chunker.py, atau vector_store.py sama sekali -- cuma memanggil fungsi
yang sudah ada dan menyimpan hasil antaranya.

Pakai:
  python inspect_pipeline.py --pdf panduan.pdf --out ./debug_output \
      [--chroma-dir ./chroma_store] [--collection panduan_sistematika] \
      [--skip-embedding]

Kalau --skip-embedding dipasang, stage 3 & 4 dilewati (berguna kalau
Ollama belum jalan, jadi tetap bisa cek hasil parsing & chunking dulu).

Output:
  stage1_parsed_sections.json  -- hasil parser.py (sebelum difilter)
  stage2_chunks.json           -- hasil chunker.py (sesudah difilter + dibentuk chunk)
  stage3_embeddings.json       -- preview vektor per chunk (dari vector_store.embed_text)
  stage4_vector_store.json     -- isi persis yang tersimpan di koleksi ChromaDB
"""

import sys, os, json, argparse

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from parser import extract_text_by_page, split_into_sections
from chunker import split_into_sentences, filter_section, build_chunks, FORMAT_KEYWORDS


def stage1_parsing(pdf_path: str, out_dir: str):
    print(f"[stage1] membaca {pdf_path} ...")
    full_text = extract_text_by_page(pdf_path)
    sections = split_into_sections(full_text)
    print(f"[stage1] {len(sections)} section (bab/sub-bab) ditemukan")

    dump = [
        {
            "bab_num": s.bab_num,
            "bab_title": s.bab_title,
            "subbab_num": s.subbab_num,
            "subbab_title": s.subbab_title,
            "raw_text_char_count": len(s.raw_text),
            "raw_text": s.raw_text,
        }
        for s in sections
    ]
    path = os.path.join(out_dir, "stage1_parsed_sections.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"[stage1] ditulis -> {path}")
    return sections


def stage2_chunking(sections, out_dir: str):
    print("[stage2] memfilter kalimat format visual per section ...")
    per_section_debug = []
    for sec in sections:
        all_sentences = split_into_sentences(sec.raw_text)
        filter_section(sec)  # ini yang mengisi sec.kept_sentences
        kept_set = set(sec.kept_sentences)
        dropped = [s for s in all_sentences if s not in kept_set]
        per_section_debug.append(
            {
                "bab_num": sec.bab_num,
                "subbab_num": sec.subbab_num,
                "subbab_title": sec.subbab_title,
                "total_sentences": len(all_sentences),
                "kept_count": len(sec.kept_sentences),
                "dropped_count": len(dropped),
                "kept_sentences": sec.kept_sentences,
                "dropped_sentences": dropped,
            }
        )

    chunks = build_chunks(sections)
    print(f"[stage2] {len(chunks)} chunk final siap embed (dari {len(sections)} section)")

    dump = {
        "format_keywords_used_for_filtering": FORMAT_KEYWORDS,
        "per_section_filter_detail": per_section_debug,
        "final_chunks": chunks,
    }
    path = os.path.join(out_dir, "stage2_chunks.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"[stage2] ditulis -> {path}")
    return chunks


def stage3_embedding(chunks, out_dir: str):
    from vector_store import embed_text

    print(f"[stage3] embedding {len(chunks)} chunk lewat Ollama (nomic-embed-text) ...")
    dump = []
    for c in chunks:
        vec = embed_text(c["text"])
        norm = sum(v * v for v in vec) ** 0.5
        dump.append(
            {
                "id": c["id"],
                "subbab_title": c["metadata"]["subbab_title"],
                "embedding_dim": len(vec),
                "embedding_norm": norm,
                "embedding_preview_first_8": vec[:8],
            }
        )
        print(f"    - {c['id']}: dim={len(vec)}")
    path = os.path.join(out_dir, "stage3_embeddings.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"[stage3] ditulis -> {path}  (preview saja, bukan full vector -- lihat stage4 utk yg tersimpan)")


def stage4_vector_store(chunks, chroma_dir: str, collection_name: str, out_dir: str):
    from vector_store import upsert_chunks, get_collection

    print(f"[stage4] upsert {len(chunks)} chunk ke ChromaDB ({chroma_dir}, koleksi={collection_name}) ...")
    n = upsert_chunks(chunks, chroma_dir, collection_name)
    print(f"[stage4] {n} chunk ter-upsert")

    collection = get_collection(chroma_dir, collection_name)
    stored = collection.get(include=["documents", "metadatas", "embeddings"])

    dump = {
        "collection_name": collection_name,
        "chroma_dir": chroma_dir,
        "total_stored": len(stored["ids"]),
        "items": [
            {
                "id": stored["ids"][i],
                "document_text": stored["documents"][i],
                "metadata": stored["metadatas"][i],
                "embedding_dim": len(stored["embeddings"][i]) if stored.get("embeddings") is not None else None,
            }
            for i in range(len(stored["ids"]))
        ],
    }
    path = os.path.join(out_dir, "stage4_vector_store.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"[stage4] ditulis -> {path}  (ini persis isi koleksi Chroma sekarang)")


def main():
    ap = argparse.ArgumentParser(description="Inspeksi tiap tahap pipeline RAG sistematika")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", default="./debug_output")
    ap.add_argument("--chroma-dir", default="./chroma_store_debug")
    ap.add_argument("--collection", default="panduan_sistematika_debug")
    ap.add_argument(
        "--skip-embedding",
        action="store_true",
        help="Lewati stage 3 & 4 (kalau Ollama/Chroma belum siap)",
    )
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)

    sections = stage1_parsing(args.pdf, args.out)
    chunks = stage2_chunking(sections, args.out)

    if args.skip_embedding:
        print("[skip] --skip-embedding dipasang, stage 3 & 4 dilewati")
        return

    stage3_embedding(chunks, args.out)
    stage4_vector_store(chunks, args.chroma_dir, args.collection, args.out)


if __name__ == "__main__":
    main()
