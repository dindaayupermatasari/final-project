"""Tes alur lengkap: judul -> problem domain/problem/uniqueness -> nilai satu subbab Bab 1.

Jalankan dari folder llm/:
    python test_judge.py "<judul proyek akhir>" [LATAR BELAKANG|PERMASALAHAN|TUJUAN] [file_teks.txt]
Sesuaikan dua path di bawah dengan lokasi file di proyekmu.
"""
import json
import sys
from pathlib import Path

from criteria import get_criteria
from judge import judge_subbab
from title_decomposer import decompose_title

BASE = Path(__file__).resolve().parent.parent  # folder sistematika/


def find_file(name: str) -> Path:
    """Cari file di bawah sistematika/ supaya tidak tergantung lokasi pastinya."""
    hits = sorted(BASE.rglob(name))
    if not hits:
        sys.exit(f"!! {name} tidak ditemukan di bawah {BASE}. Jalankan dulu skrip yang membuatnya.")
    return hits[0]


SECTIONIZER_JSON = find_file("hasil_sectionizer.json")
SNAPSHOT_JSON = find_file("production_snapshot.json")

judul = sys.argv[1]
target = (sys.argv[2] if len(sys.argv) > 2 else "TUJUAN").upper()

decomp = decompose_title(judul)
print("== Hasil pemecahan judul")
print(json.dumps(decomp, ensure_ascii=False, indent=2))
if decomp["perlu_dicek"]:
    print("!! Pemecahan judul perlu dicek manual sebelum dipercaya.")

sections = json.load(open(SECTIONIZER_JSON, encoding="utf-8"))["sections"]
print(f"\n(sectionizer: {SECTIONIZER_JSON})")
bab1 = [s for s in sections if str(s.get("bab_num")) == "1"]
sec = next((s for s in bab1 if s["subbab_title"].strip().upper() == target), None)
if sec is None:
    print(f"!! Subbab '{target}' tidak ada di Bab 1 pada file itu. Subbab Bab 1 yang ada:")
    for s in bab1:
        print(f"   {s['subbab_num']} {s['subbab_title']}")
    sys.exit("Cek file sectionizer yang dipakai, atau salah satu judul di atas.")

spec = get_criteria(target)
snap = json.load(open(SNAPSHOT_JSON, encoding="utf-8"))["items"]
guideline = next((i["document_text"] for i in snap if i["id"] == spec["guideline_id"]), "")
if not guideline:
    print(f"!! Chunk panduan {spec['guideline_id']} tidak ditemukan di snapshot.")

teks = sec["raw_text"]
if len(sys.argv) > 3:  # opsional: nilai teks dari file .txt (untuk uji versi yang sengaja dirusak)
    teks = Path(sys.argv[3]).read_text(encoding="utf-8")
    print(f"(memakai teks dari {sys.argv[3]})")

print(f"\n== Penilaian subbab {target}")
out = judge_subbab(decomp, sec["subbab_title"], teks, guideline)
print(json.dumps(out, ensure_ascii=False, indent=2))
