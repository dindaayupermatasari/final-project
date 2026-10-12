"""
test_retrieval_from_student_doc.py — pipelines/sistematika/

Mengukur tahap retrieval (dan opsional re-ranking) pada sistem: tiap sub-bab
dokumen mahasiswa dijadikan query, lalu dicek apakah sub-bab panduan yang
benar berhasil ditemukan.

Metrik yang dipakai sengaja disesuaikan dengan desain sistem
(dense retrieval -> cross-encoder -> LLM):

  * recall@K (retriever): apakah jawaban benar ada di dalam K kandidat yang
    diteruskan ke re-ranker. Ini BATAS ATAS: chunk yang tidak masuk kandidat
    tidak bisa diselamatkan re-ranker maupun LLM.
  * hit@1 / recall@k / MRR setelah re-rank (--rerank): apakah cross-encoder
    berhasil menaikkan jawaban benar ke atas di antara kandidat yang sama.

Mode query (--mode): label | content | label+content | all
  label          -> judul sub-bab saja (sebagian "mencontek": header indeks
                    memuat judul sub-bab)
  content        -> isi dibersihkan saja (ujian paling jujur)
  label+content  -> judul + isi

Pakai:
  python test_retrieval_from_student_doc.py --doc draft.docx --mode all
  python test_retrieval_from_student_doc.py --doc draft.docx --mode all --rerank
  python test_retrieval_from_student_doc.py --doc draft.docx --mode content --rerank \
         --rerank-no-header --candidates-k 8
"""

import sys, os, re, json, argparse, difflib
from typing import Dict, List, Optional

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from config import PATHS
from student_sectionizer import sectionize_student_doc
from vector_store import (
    query_collection,
    get_collection,
    normalize_title as _norm,
    PREFIXES,
)

MODES = ["label", "content", "label+content"]

_CAPTION = re.compile(r"^(tabel|gambar)\s+\d", re.IGNORECASE)
_SENT_END = re.compile(r"[.:;!?]$")


# ---------------------------------------------------------------------
# Pembersihan & penyusunan query
# ---------------------------------------------------------------------


def clean_query_text(text: str) -> str:
    """Buang nomor halaman, caption, baris tabel/angka, fragmen sel; gabung baris wrap."""
    kept = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if re.fullmatch(r"\d{1,3}", s):
            continue
        if _CAPTION.match(s):
            continue
        letters = sum(ch.isalpha() for ch in s)
        digits = sum(ch.isdigit() for ch in s)
        if digits and digits >= letters:
            continue
        if len(s) < 25 and not _SENT_END.search(s):
            continue
        kept.append(s)
    return " ".join(" ".join(kept).split())


def truncate_at_sentence(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars]
    i = max(cut.rfind(". "), cut.rfind("? "), cut.rfind("! "))
    return cut[: i + 1] if i > max_chars * 0.5 else cut


def build_query(mode: str, title: str, cleaned: str, max_chars: int) -> str:
    body = truncate_at_sentence(cleaned, max_chars)
    if mode == "label":
        return title
    if mode == "content":
        return body
    return f"{title}. {body}".strip()


# ---------------------------------------------------------------------
# Evaluasi
# ---------------------------------------------------------------------


def map_to_guideline_title(title: str, guideline_titles: set) -> Optional[str]:
    t = _norm(title)
    if t in guideline_titles:
        return t
    close = difflib.get_close_matches(t, list(guideline_titles), n=1, cutoff=0.88)
    return close[0] if close else None


def rank_of(matches: Optional[List[dict]], expected_norm: str) -> Optional[int]:
    if not matches:
        return None
    for m in matches:
        if _norm(m["subbab_title"] or "") == expected_norm:
            return m["rank"]
    return None


def summarize_ranks(ranks: List[Optional[int]], ks: List[int]) -> dict:
    """ranks: peringkat jawaban benar per item evaluable (None = tidak ditemukan)."""
    n = len(ranks)
    if n == 0:
        return {"evaluable": 0}
    out = {"evaluable": n}
    for k in ks:
        out[f"recall@{k}"] = round(sum(r is not None and r <= k for r in ranks) / n, 3)
    out["mrr"] = round(sum(1 / r for r in ranks if r) / n, 3)
    return out


def print_top(matches: List[dict], top_k: int, reranked: bool):
    for m in matches[:top_k]:
        extra = (
            f"  rerank {m['rerank_score']}  (retriever #{m['retriever_rank']})"
            if reranked
            else f"  (dist {m['distance']})"
        )
        print(f"     #{m['rank']} {m['subbab_num']} {m['subbab_title']}{extra}")
    print()


def slim(matches: List[dict]) -> List[dict]:
    """Versi ringkas untuk JSON (tanpa teks chunk)."""
    keys = [
        "rank",
        "retriever_rank",
        "rerank_score",
        "chunk_id",
        "subbab_num",
        "subbab_title",
        "chunk_type",
        "distance",
    ]
    return [{k: m[k] for k in keys if k in m} for m in matches]


def run_mode(mode, sections, args, guideline_titles, reranker):
    from reranker import rerank_matches  # impor lazy: hanya bila --rerank

    chunk_type = None if args.chunk_type == "all" else args.chunk_type
    results, skipped = [], []
    ranks_ret: List[Optional[int]] = []
    ranks_rr: List[Optional[int]] = []

    for sec in sections:
        title = sec.subbab_title
        label = f"BAB {sec.bab_num} > {sec.subbab_num} {title} (dari dokumen mahasiswa)"
        expected = map_to_guideline_title(title, guideline_titles)
        cleaned = clean_query_text(sec.raw_text)

        if mode != "label" and not cleaned:
            skipped.append(
                {"mode": mode, "label": label, "error": "kosong setelah dibersihkan"}
            )
            if expected:  # dihitung MISS supaya penyebut sama antar mode
                ranks_ret.append(None)
                ranks_rr.append(None)
            print(f"  [{mode}] {label} -> DILEWATI (isi kosong setelah dibersihkan)\n")
            continue

        query = build_query(mode, title, cleaned, args.max_chars)
        try:
            cands = query_collection(
                query,
                args.chroma_dir,
                args.collection,
                top_k=args.candidates_k,
                kind=args.prefix,
                chunk_type=chunk_type,
            )
            ranked = (
                rerank_matches(
                    query, cands, reranker, use_header=not args.rerank_no_header
                )
                if reranker
                else None
            )
        except Exception as e:
            print(f"     GAGAL [{label}]: {e}\n")
            skipped.append(
                {"mode": mode, "label": label, "error": str(e), "preview": query[:300]}
            )
            continue

        r_ret = rank_of(cands, expected) if expected else None
        r_rr = rank_of(ranked, expected) if (expected and ranked) else None
        if expected:
            ranks_ret.append(r_ret)
            ranks_rr.append(r_rr)

        if expected is None:
            flag = "-"
        elif reranker:
            flag = f"retriever #{r_ret or 'x'} -> rerank #{r_rr or 'x'}"
        else:
            flag = f"retriever #{r_ret or 'x'}"
        print(f"  [{mode}] {label}  ({len(query)} char) -> {flag}")
        print_top(ranked if ranked else cands, args.top_k, reranked=bool(ranked))

        results.append(
            {
                "query_label": label,
                "query_text": query,
                "expected_guideline_title": expected,
                "retriever_rank": r_ret,
                "rerank_rank": r_rr,
                "candidates": slim(ranked if ranked else cands),
            }
        )

    ks = sorted({k for k in (1, 3, 5, args.candidates_k) if k <= args.candidates_k})
    summary = {
        "total_sections": len(sections),
        "retriever": summarize_ranks(ranks_ret, ks),
        "reranked": summarize_ranks(ranks_rr, ks) if reranker else None,
    }
    return results, summary, skipped


def main():
    ap = argparse.ArgumentParser(
        description="Tes retrieval (+ re-ranking) pakai dokumen asli mahasiswa"
    )
    ap.add_argument("--doc", required=True, help="Dokumen mahasiswa (.docx / .pdf)")
    ap.add_argument("--top-k", type=int, default=3, help="Yang ditampilkan per query")
    ap.add_argument(
        "--candidates-k",
        type=int,
        default=8,
        help="Jumlah kandidat dari retriever yang diteruskan ke re-ranker "
        "(samakan dengan setting produksi). recall@K dihitung di sini.",
    )
    ap.add_argument("--chroma-dir", default=PATHS["chroma_dir"])
    ap.add_argument("--collection", default=PATHS["chroma_collection"])
    ap.add_argument("--out", default="retrieval_test_from_student_doc.json")
    ap.add_argument("--mode", choices=MODES + ["all"], default="label+content")
    ap.add_argument("--max-chars", type=int, default=1000)
    ap.add_argument("--prefix", choices=list(PREFIXES), default="query")
    ap.add_argument(
        "--chunk-type", choices=["instruction", "example", "all"], default="instruction"
    )
    ap.add_argument("--rerank", action="store_true", help="Aktifkan cross-encoder")
    ap.add_argument(
        "--reranker-model",
        default=None,
        help="Nama model HuggingFace (default: BAAI/bge-reranker-v2-m3)",
    )
    ap.add_argument("--rerank-max-length", type=int, default=1024)
    ap.add_argument(
        "--rerank-no-header",
        action="store_true",
        help="Reranker hanya melihat isi chunk (tanpa judul sub-bab panduan)",
    )
    args = ap.parse_args()
    args.candidates_k = max(args.candidates_k, args.top_k)

    print(f"Membaca {args.doc} ...")
    sections, rejected = sectionize_student_doc(args.doc)
    print(
        f"  ditemukan {len(sections)} sub-bab, {len(rejected)} kandidat ditolak (kemungkinan tabel/data)\n"
    )
    if not sections:
        print("PERINGATAN: 0 sub-bab kedeteksi. Cek format heading dokumen ini.")
        return

    col = get_collection(args.chroma_dir, args.collection, read_only=True)
    stored = col.get(include=["metadatas"])
    guideline_titles = {_norm(m.get("subbab_title", "")) for m in stored["metadatas"]}

    reranker = None
    if args.rerank:
        from reranker import CrossEncoderReranker, DEFAULT_MODEL

        model = args.reranker_model or DEFAULT_MODEL
        print(f"Memuat reranker {model} (pertama kali: mengunduh model) ...")
        reranker = CrossEncoderReranker(model, max_length=args.rerank_max_length)
        print("  siap\n")

    modes = MODES if args.mode == "all" else [args.mode]
    out: Dict[str, dict] = {}
    all_skipped = []
    for mode in modes:
        print(
            f"===== MODE: {mode} (prefix={args.prefix}, chunk_type={args.chunk_type}, "
            f"K={args.candidates_k}{', rerank' if reranker else ''}) ====="
        )
        results, summary, skipped = run_mode(
            mode, sections, args, guideline_titles, reranker
        )
        out[mode] = {"summary": summary, "results": results}
        all_skipped += skipped

    print(f"\n===== RINGKASAN (K kandidat = {args.candidates_k}) =====")
    for mode, data in out.items():
        s = data["summary"]
        print(f"  {mode:15s} retriever {s['retriever']}")
        if s["reranked"]:
            print(f"  {'':15s} +rerank   {s['reranked']}")
    print(
        "\nCatatan: recall@K retriever = batas atas; re-ranker hanya bisa mengurutkan "
        "ulang kandidat yang sama, jadi recall@K-nya tidak akan melebihi retriever."
    )

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(
            {
                "settings": {
                    "prefix": args.prefix,
                    "chunk_type": args.chunk_type,
                    "top_k": args.top_k,
                    "candidates_k": args.candidates_k,
                    "max_chars": args.max_chars,
                    "rerank": bool(reranker),
                    "reranker_model": (reranker.model_name if reranker else None),
                    "rerank_header": (not args.rerank_no_header) if reranker else None,
                },
                "modes": out,
                "skipped": all_skipped,
                "rejected_sectionizer_candidates": [
                    {
                        "bab_num": r.bab_num,
                        "matched_num": r.matched_num,
                        "matched_title": r.matched_title,
                        "reason": r.reason,
                    }
                    for r in rejected
                ],
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    print(f"\nSemua hasil ditulis -> {args.out}")


if __name__ == "__main__":
    main()
