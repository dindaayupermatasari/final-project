"""
main.py
=======
ORKESTRASI PIPELINE (sesuai diagram desain sistem):
Parsing -> Extract Parameter -> Rule-based Comparator
        -> Calculate Score -> Output Format Score

CHANGELOG (revisi -- format input adaptif):
- run_pipeline() SEKARANG mendeteksi ekstensi file & memilih jalur
  extractor yang sesuai, bukan cuma menerima .docx:
    .docx        -> jalur LAMA (extractor.py, baca XML docx langsung
                     -- paling presisi). Tetap di-render ke PDF SEKALI
                     (render_to_pdf) khusus untuk cek fisik halaman
                     (ganjil/genap, footer), SAMA seperti sebelumnya.
    .pdf         -> jalur BARU (pdf_extractor.py). Ini jalur UTAMA
                     sekarang karena mayoritas mahasiswa upload PDF
                     langsung, dan pembimbing eksplisit minta PDF
                     diutamakan. Tidak perlu render_to_pdf lagi --
                     filenya sendiri sudah PDF.
    lainnya      -> (.doc, .odt, dst) convert dulu ke PDF via
                     LibreOffice (render_to_pdf() SUDAH generic,
                     bukan cuma nerima docx -- cuma dokumentasinya
                     yang menyebut "docx", implementasinya cuma
                     manggil `soffice --convert-to pdf`), lalu ikut
                     jalur .pdf di atas.
- Kedua jalur extractor (docx & pdf) mengembalikan bentuk `params`
  yang SAMA PERSIS ({"paragraphs","sections","tables","mirror_margins"}),
  jadi compare_all/check_bibliography/calculate_score di bawahnya
  TIDAK PERLU tahu jalur mana yang dipakai -- lihat pdf_extractor.py
  untuk penjelasan lengkap kenapa ini bisa dijamin sama.

BUG DIPERBAIKI (2026-08-20 -- "Umum (Penomoran Halaman)" & margin
selalu 0% di jalur PDF): _extract_params() jalur .pdf/lainnya dulu
BERHENTI di extract_all_pdf() apa adanya -- params["sections"] balik 1
section ber-pgnum_fmt/pgnum_start/footer_has_page_field/footer_alignment
SEMUANYA None (lihat komentar di extract_section_params_pdf(),
pdf_extractor.py, yang MEMANG sengaja mendelegasikan 4 field itu ke
page_layout_checker.py), dan SEMUA paragraf["section"] hardcoded 0.
Tidak ada satupun kode yang benar-benar mengisi 4 field itu ATAU
membetulkan paragraf["section"] -- compare_page_numbering() di
comparator.py jadi PASTI gagal di cek 0/1/2 apapun isi dokumennya, dan
find_section_start_pages() di bawah cuma pernah nemu 1 section (karena
semua paragraf section=0), jadi Section Break sebelum BAB 1 SELALU
dilaporkan "tidak ditemukan" walau dokumennya benar.
FIX: page_layout_checker.finalize_page_numbering() dipanggil di
_extract_params() SETELAH extract_all_pdf() (baca footer PDF asli,
bangun section penomoran beneran, betulkan paragraf["section"]) --
SEBELUM params dikembalikan ke run_pipeline(), supaya find_section_start_pages()
& compare_all() & check_page_layout() di bawah semua otomatis dapat
data yang benar tanpa perlu tahu detail jalur mana yang dipakai. Jalur
.docx TIDAK disentuh -- section & pgnum_fmt/dst di jalur itu sudah
akurat langsung dari XML (w:pgNumType), tidak butuh fix ini.
"""

import os
import sys
import re
import json
from collections import defaultdict

from aggregator.pipelines.visual.parser import parse_docx
from aggregator.pipelines.visual.extractor import extract_all
from aggregator.pipelines.visual.pdf_extractor import extract_all_pdf
from aggregator.pipelines.visual.comparator import compare_all
from aggregator.pipelines.visual.page_layout_checker import (
    render_to_pdf,
    extract_page_texts,
    extract_footer_numbers,
    check_page_layout,
    find_section_start_pages,
    finalize_page_numbering,
)
from aggregator.pipelines.visual.scorer import calculate_score

# DINONAKTIFKAN (atas permintaan user, 2026-08-21): cek format DAFTAR
# PUSTAKA di-nonaktifkan sementara -- lihat blok yang dikomentari di
# run_pipeline() di bawah. Import-nya ikut dikomentari supaya tidak ada
# import yang menggantung tak terpakai.
# from aggregator.pipelines.visual.bibliography_checker import check_bibliography

# Ekstensi yang dianggap "docx asli" -- jalur extractor XML presisi
# tinggi. Selain ini & .pdf, semua dianggap "perlu dikonversi dulu".
DOCX_EXTENSIONS = {".docx"}
PDF_EXTENSIONS = {".pdf"}


def _chapter_sort_key(chapter: str):
    """
    Urutan tampil yang natural: Pra-Bab dulu, lalu BAB 1/2/3/... urut
    angka, baru grup "Umum (...)" dan DAFTAR PUSTAKA di akhir.
    """
    if chapter == "Pra-Bab":
        return (0, 0, chapter)
    m = re.match(r"^BAB\s+(\d+)", chapter)
    if m:
        return (1, int(m.group(1)), chapter)
    if chapter == "DAFTAR PUSTAKA":
        return (3, 0, chapter)
    return (2, 0, chapter)


def group_findings_by_chapter(findings: list[dict]) -> dict:
    grouped = defaultdict(list)
    for f in findings:
        grouped[f["chapter"]].append({k: v for k, v in f.items() if k != "chapter"})
    return {
        chapter: grouped[chapter] for chapter in sorted(grouped, key=_chapter_sort_key)
    }


def merge_checks(base: dict, extra: dict) -> dict:
    """Gabung dua dict checks ({chapter: {checked, failed}}) dengan menjumlahkan."""
    merged = defaultdict(lambda: {"checked": 0, "failed": 0})
    for source in (base, extra):
        for chapter, c in source.items():
            merged[chapter]["checked"] += c["checked"]
            merged[chapter]["failed"] += c["failed"]
    return dict(merged)


def _extract_params(filepath: str) -> tuple[dict, str, str]:
    """
    Pilih jalur extractor sesuai ekstensi file, balik
    (params, pdf_path, source). pdf_path SELALU balik path PDF yang
    valid (baik file aslinya sendiri, atau hasil convert) -- dipakai
    bareng oleh check_page_layout & find_section_start_pages di bawah,
    SAMA seperti alur lama, cuma sekarang sumbernya bisa dari salah
    satu dari dua jalur.

    source ("docx" | "pdf") ditambahkan (2026-08-24) supaya
    run_pipeline() bisa meneruskannya ke compare_all() -> compare_sections(),
    yang butuh tahu jalur mana yang dipakai buat nentuin toleransi
    margin yang tepat (lihat docstring compare_sections di comparator.py):
    margin dari .docx dibaca EXACT dari XML, sedangkan margin dari PDF
    (baik file .pdf asli maupun hasil convert .doc/.odt) cuma ESTIMASI
    geometris dari posisi teks -- jadi wajar lebih longgar. Makanya
    cabang .docx balik source="docx", dua cabang lainnya balik
    source="pdf" (cabang "convert dulu" tetap "pdf" karena ujungnya
    tetap lewat extract_all_pdf() juga, sama-sama estimasi geometris,
    bukan baca XML).

    Jalur .pdf & "lainnya" (convert dulu) SEKARANG WAJIB lewat
    finalize_page_numbering() sebelum return -- lihat CHANGELOG di atas
    modul ini untuk root-cause lengkap kenapa ini wajib ada di sini
    (bukan cukup ditambahkan di pdf_extractor.py/page_layout_checker.py
    saja): tanpa panggilan ini, params["sections"] & paragraf["section"]
    tetap versi stub lama, dan SEMUA fungsi di bawah run_pipeline()
    (find_section_start_pages, compare_all, check_page_layout) diam-diam
    kembali ke perilaku lama yang selalu salah, walau kode di
    page_layout_checker.py sendiri sudah benar.
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext in DOCX_EXTENSIONS:
        doc = parse_docx(filepath)
        params = extract_all(doc)
        pdf_path = render_to_pdf(filepath)
        return params, pdf_path, "docx"

    if ext in PDF_EXTENSIONS:
        params = extract_all_pdf(filepath)
        params = finalize_page_numbering(params, filepath)
        return params, filepath, "pdf"

    # Format lain (.doc, .odt, dst) -- convert dulu ke PDF, lalu ikut
    # jalur PDF. render_to_pdf() namanya menyebut "docx" di docstring
    # lama, tapi implementasinya (_render_via_libreoffice) cuma
    # menjalankan `soffice --convert-to pdf` apa adanya -- generic
    # untuk format apapun yang didukung LibreOffice, jadi aman dipakai
    # di sini tanpa perlu diubah.
    pdf_path = render_to_pdf(filepath)
    params = extract_all_pdf(pdf_path)
    params = finalize_page_numbering(params, pdf_path)
    return params, pdf_path, "pdf"


def run_pipeline(filepath: str) -> dict:
    # 1 & 2. Parsing + Extract Parameter (jalur dipilih berdasar format input)
    params, pdf_path, source = _extract_params(filepath)

    # 2b. page_texts & footer_numbers SELALU dari PDF (baik file aslinya
    #     PDF, atau hasil render/convert) -- dipakai bareng oleh
    #     check_page_layout() (ganjil/genap awal bab) DAN
    #     find_section_start_pages() (lokasi halaman fisik di finding
    #     penomoran halaman). Untuk jalur PDF, params["paragraphs"]
    #     sudah punya elem["section"] yang BENAR di titik ini (dibetulkan
    #     oleh finalize_page_numbering() di _extract_params()), jadi
    #     find_section_start_pages() otomatis nemu SEMUA section (bukan
    #     cuma section 0) tanpa perlu tahu jalur mana yang dipakai.
    page_texts = extract_page_texts(pdf_path)
    footer_numbers = extract_footer_numbers(pdf_path)
    section_pages = find_section_start_pages(page_texts, params["paragraphs"])

    # 3. Rule-based Comparator
    # source diteruskan supaya compare_sections() tahu apakah margin di
    # `params["sections"]` ini dibaca EXACT dari XML docx atau cuma
    # ESTIMASI geometris dari PDF -- lihat docstring compare_sections()
    # di comparator.py & _extract_params() di atas.
    findings, checks = compare_all(params, section_pages, source=source)

    # 3b. Render & Locate: cek halaman ganjil/genap awal Bab
    layout_findings, layout_checks = check_page_layout(
        page_texts, params["sections"], section_pages, footer_numbers
    )
    findings = findings + layout_findings
    checks = merge_checks(checks, layout_checks)

    # DINONAKTIFKAN (atas permintaan user, 2026-08-21): cek format
    # DAFTAR PUSTAKA (check_bibliography, dari bibliography_checker.py)
    # di-nonaktifkan sementara -- findings & checks dari rule ini tidak
    # lagi digabung ke hasil akhir. Aktifkan kembali dengan un-comment
    # blok ini + import check_bibliography di atas modul ini kalau rule
    # ini mau dipakai lagi.
    # bib_findings, bib_checks = check_bibliography(params["paragraphs"])
    # findings = findings + bib_findings
    # checks = merge_checks(checks, bib_checks)

    # 4. Calculate Score (per bab)
    summary = calculate_score(checks)
    summary["per_chapter"] = {
        chapter: summary["per_chapter"][chapter]
        for chapter in sorted(summary["per_chapter"], key=_chapter_sort_key)
    }

    # 5. Output Format Score
    return {
        "findings_per_chapter": group_findings_by_chapter(findings),
        "summary": summary,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python main.py path/ke/dokumen.[docx|pdf|doc|odt] [output.json]")
        sys.exit(1)

    filepath = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "CheckRun/output/hasil.json"

    result = run_pipeline(filepath)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    summary = result["summary"]
    print("=" * 55)
    print(f"Skor keseluruhan : {summary['overall_score']}")
    print(f"Total temuan     : {summary['total_findings']}")
    print("-" * 55)
    for chapter, c in summary["per_chapter"].items():
        print(
            f"  {chapter:<28} {c['score']:>6.2f}  ({c['failed']}/{c['checked']} gagal)"
        )
    print("=" * 55)
