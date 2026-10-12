"""
test_retrieval.py — pipelines/sistematika/

Uji RETRIEVAL beneran dari koleksi produksi: masukkan potongan teks
(misal draft bab 1 mahasiswa), lihat sub-bab panduan mana yang paling
mirip. Ini yang paling relevan buat didemoin ke dosen, karena
nunjukkin RAG-nya "nyambung" -- bukan cuma bukti tiap tahap pipeline
jalan, tapi bukti hasil akhirnya (retrieval) masuk akal.

Tidak mengubah parser/chunker/vector_store -- cuma manggil
query_similar() yang sudah ada.

Pakai -- satu query manual:
  python test_retrieval.py --query "Penelitian ini membahas tentang..." --top-k 3

Pakai -- banyak query sekaligus (buat demo, tiap query dikasih label):
  python test_retrieval.py --queries-file contoh_queries.json --top-k 3

Format contoh_queries.json:
  {
    "Cuplikan Bab 1 - Latar Belakang (draft mahasiswa)": "teks draft di sini...",
    "Cuplikan Bab 3 - Perancangan Sistem (draft mahasiswa)": "teks draft di sini..."
  }

Output:
  - dicetak ke terminal dengan format rapi (siap discreenshot/didemoin)
  - disimpan juga ke JSON (--out, default retrieval_test_results.json)
"""

import sys, os, json, argparse

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from config import PATHS
from vector_store import query_similar


def run_one_query(label: str, query_text: str, chroma_dir: str, collection: str, top_k: int) -> dict:
    result = query_similar(query_text, chroma_dir, collection, top_k=top_k)

    ids = result.get("ids", [[]])[0]
    documents = result.get("documents", [[]])[0]
    metadatas = result.get("metadatas", [[]])[0]
    distances = result.get("distances", [[]])[0]

    matches = []
    for i in range(len(ids)):
        matches.append({
            "rank": i + 1,
            "chunk_id": ids[i],
            "bab_title": metadatas[i].get("bab_title"),
            "subbab_num": metadatas[i].get("subbab_num"),
            "subbab_title": metadatas[i].get("subbab_title"),
            "distance": distances[i],
            "matched_text": documents[i],
        })

    return {
        "query_label": label,
        "query_text": query_text,
        "top_k": top_k,
        "matches": matches,
    }


def print_result(res: dict):
    print("=" * 70)
    print(f"QUERY: {res['query_label']}")
    print(f"  teks: {res['query_text'][:150]}{'...' if len(res['query_text']) > 150 else ''}")
    print("-" * 70)
    for m in res["matches"]:
        print(f"  #{m['rank']}  [BAB {m['bab_title']} > {m['subbab_num']} {m['subbab_title']}]  "
              f"(distance={m['distance']:.4f})")
        preview = m["matched_text"][:150]
        print(f"       \"{preview}{'...' if len(m['matched_text']) > 150 else ''}\"")
    print()


def main():
    ap = argparse.ArgumentParser(description="Tes retrieval RAG terhadap koleksi panduan produksi")
    ap.add_argument("--query", help="Satu query teks langsung dari command line")
    ap.add_argument("--queries-file", help="Path ke JSON {label: query_text} buat banyak query sekaligus")
    ap.add_argument("--top-k", type=int, default=3)
    ap.add_argument("--chroma-dir", default=PATHS["chroma_dir"])
    ap.add_argument("--collection", default=PATHS["chroma_collection"])
    ap.add_argument("--out", default="retrieval_test_results.json")
    args = ap.parse_args()

    if not args.query and not args.queries_file:
        ap.error("Isi salah satu: --query \"teks\" atau --queries-file path.json")

    queries = {}
    if args.query:
        queries["Query manual"] = args.query
    if args.queries_file:
        with open(args.queries_file, "r", encoding="utf-8") as f:
            queries.update(json.load(f))

    print(f"Koleksi: {args.collection}  |  dir: {args.chroma_dir}  |  top_k={args.top_k}")
    print()

    all_results = []
    for label, query_text in queries.items():
        res = run_one_query(label, query_text, args.chroma_dir, args.collection, args.top_k)
        print_result(res)
        all_results.append(res)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"Semua hasil ditulis -> {args.out}")


if __name__ == "__main__":
    main()
