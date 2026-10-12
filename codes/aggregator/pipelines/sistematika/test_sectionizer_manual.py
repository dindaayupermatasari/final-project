"""
test_sectionizer_manual.py — pipelines/sistematika/

Jalankan student_sectionizer.py ke satu dokumen mahasiswa dan dump
hasilnya ke JSON: section yang berhasil dikenali DAN kandidat yang
DIBUANG (rejected_candidates) -- biar bisa diaudit manual apakah
aturan filternya kebablasan (nolak sub-bab yang beneran) atau sudah pas.

Pakai:
  python test_sectionizer_manual.py --doc "path/draft.pdf" --out hasil.json
  python test_sectionizer_manual.py --doc "path/draft.docx" --out hasil.json
"""

import argparse, json
from student_sectionizer import sectionize_student_doc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", required=True)
    ap.add_argument("--out", default="sectionizer_result.json")
    args = ap.parse_args()

    print(f"Membaca {args.doc} ...")
    sections, rejected = sectionize_student_doc(args.doc)

    print(f"\n{len(sections)} section diterima:")
    for s in sections:
        print(f"  [{s.source}] BAB {s.bab_num} > {s.subbab_num}  {s.subbab_title}  ({len(s.raw_text)} char)")

    print(f"\n{len(rejected)} kandidat DITOLAK (kemungkinan baris tabel/data):")
    for r in rejected:
        print(f"  BAB {r.bab_num} > '{r.matched_num} {r.matched_title}'  -- {r.reason}")

    dump = {
        "total_sections": len(sections),
        "total_rejected": len(rejected),
        "sections": [
            {
                "source": s.source,
                "bab_num": s.bab_num,
                "bab_title": s.bab_title,
                "subbab_num": s.subbab_num,
                "subbab_title": s.subbab_title,
                "raw_text_char_count": len(s.raw_text),
                "raw_text": s.raw_text,
            }
            for s in sections
        ],
        "rejected_candidates": [
            {
                "bab_num": r.bab_num,
                "matched_num": r.matched_num,
                "matched_title": r.matched_title,
                "reason": r.reason,
            }
            for r in rejected
        ],
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2)
    print(f"\nDitulis -> {args.out}")


if __name__ == "__main__":
    main()
