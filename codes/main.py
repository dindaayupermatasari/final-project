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
    .rtf / .doc  -> jalur BARU (rtf_converter.py). Convert dulu ke
                     .docx via LibreOffice (BUKAN ke PDF -- lihat
                     CHANGELOG "FIX RTF" di bawah untuk alasannya),
                     lalu ikut jalur .docx di atas apa adanya, dengan
                     fidelity yang sama seperti dokumen .docx asli.
    lainnya      -> (.odt, dst) convert dulu ke PDF via LibreOffice
                     (render_to_pdf() SUDAH generic, bukan cuma
                     nerima docx -- cuma dokumentasinya yang menyebut
                     "docx", implementasinya cuma manggil
                     `soffice --convert-to pdf`), lalu ikut jalur
                     .pdf di atas.
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

BUG DIPERBAIKI (2026-09-08 -- "FIX RTF": .rtf salah masuk jalur PDF):
_extract_params() SEBELUMNYA tidak punya cabang khusus untuk .rtf/.doc
sama sekali -- ekstensi ini jatuh ke cabang "format lain" paling
bawah, yang mengonversi file ke PDF (render_to_pdf) lalu memprosesnya
lewat extract_all_pdf(). Akibatnya file .rtf dicek pakai MESIN YANG
SALAH: estimasi geometris dari koordinat PDF, bukan pembacaan XML
docx yang presisi -- padahal modul rtf_converter.py (convert_to_docx,
via LibreOffice --convert-to docx) SUDAH ADA di project ini persis
untuk mencegah ini, tapi tidak pernah di-import/dipanggil di sini.
Gejala yang diamati: hasil cek .rtf berbeda jauh dari .docx sumber
yang identik -- jumlah paragraf ter-parsing beda drastis (PyMuPDF
memecah teks per blok visual, bukan per elemen XML <w:p>), temuan
alignment gambar "rata tengah" hilang (cek ini memang sengaja
dinonaktifkan di jalur PDF karena tidak reliable dari koordinat),
beberapa temuan font baru muncul (PyMuPDF baca font per span, beda
sensitivitas dari python-docx yang baca dari style Word), dan margin
section berubah pola (7 section jadi 2, margin ganjil/genap asimetris
ala mirror-margin PDF) karena python-docx & PyMuPDF membaca section
break dengan cara yang sama sekali berbeda.
FIX: tambah cabang RTF_DOC_EXTENSIONS di _extract_params() yang
memanggil rtf_converter.convert_to_docx() lebih dulu, lalu meneruskan
hasilnya ke jalur .docx yang sudah ada (parse_docx + extract_all +
render_to_pdf), PERSIS seperti kalau mahasiswa upload .docx asli --
bukan membuat jalur baru. source yang dikembalikan juga "docx" (bukan
"pdf"), supaya compare_sections() di comparator.py memakai toleransi
margin yang ketat (exact dari XML), sesuai kenyataan bahwa margin
sudah dibaca dari file .docx hasil konversi, bukan diestimasi dari
koordinat PDF.

BUG DIPERBAIKI (2026-09-09 -- finding halaman ganjil/genap disajikan
seolah pasti benar walau font render tidak terverifikasi): render_to_pdf()
sekarang balik tuple (pdf_path, font_status) -- BUKAN cuma pdf_path
polos lagi (lihat CHANGELOG page_layout_checker.py). _extract_params()
di bawah SEKARANG ikut balik font_status sebagai elemen ke-4
(params, pdf_path, source, font_status), dan run_pipeline() WAJIB
meneruskannya ke check_page_layout() lewat parameter font_status= --
kalau tidak, seluruh manfaat perbaikan itu (finding ditandai
confidence="rendah" saat font "Times New Roman" tidak terverifikasi
terpasang di server render) hilang percuma. Jalur .pdf (mahasiswa
upload PDF asli, TIDAK lewat render_to_pdf()/LibreOffice sama sekali)
selalu diberi font_status="verified" -- tidak ada substitusi font yang
mungkin terjadi kalau PDF-nya memang sudah jadi dari sono-nya (mis.
export langsung dari Word oleh mahasiswa sendiri).
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
from aggregator.pipelines.visual.rtf_converter import convert_to_docx, ConversionError

# DINONAKTIFKAN (atas permintaan user, 2026-08-21): cek format DAFTAR
# PUSTAKA di-nonaktifkan sementara -- lihat blok yang dikomentari di
# run_pipeline() di bawah. Import-nya ikut dikomentari supaya tidak ada
# import yang menggantung tak terpakai.
# from aggregator.pipelines.visual.bibliography_checker import check_bibliography

# Ekstensi yang dianggap "docx asli" -- jalur extractor XML presisi
# tinggi.
DOCX_EXTENSIONS = {".docx"}
PDF_EXTENSIONS = {".pdf"}
# Ekstensi yang dikonversi ke .docx DULU (via rtf_converter.py), lalu
# ikut jalur .docx apa adanya -- BUKAN dikonversi ke PDF. Lihat
# CHANGELOG "FIX RTF" di atas modul ini untuk alasan lengkapnya.
RTF_DOC_EXTENSIONS = {".rtf", ".doc"}


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


def _extract_params(filepath: str) -> tuple[dict, str, str, str]:
    """
    Pilih jalur extractor sesuai ekstensi file, balik
    (params, pdf_path, source, font_status). pdf_path SELALU balik path
    PDF yang valid (baik file aslinya sendiri, atau hasil convert) --
    dipakai bareng oleh check_page_layout & find_section_start_pages di
    bawah, SAMA seperti alur lama, cuma sekarang sumbernya bisa dari
    salah satu dari beberapa jalur.

    font_status ("verified"/"skipped"/"unconfirmed", ditambahkan
    2026-09-09) adalah status verifikasi font "Times New Roman" untuk
    PDF yang dipakai di atas -- lihat CHANGELOG modul ini & CHANGELOG
    page_layout_checker.py untuk alasan lengkap. run_pipeline() WAJIB
    meneruskan ini ke check_page_layout(font_status=...) supaya finding
    ganjil/genap ikut ditandai confidence-nya. Jalur .pdf (tidak lewat
    render_to_pdf()/LibreOffice sama sekali -- filenya sudah PDF asli
    dari mahasiswa) selalu "verified", karena tidak ada substitusi font
    yang mungkin terjadi di jalur itu.

    source ("docx" | "pdf") ditambahkan (2026-08-24) supaya
    run_pipeline() bisa meneruskannya ke compare_all() -> compare_sections(),
    yang butuh tahu jalur mana yang dipakai buat nentuin toleransi
    margin yang tepat (lihat docstring compare_sections di comparator.py):
    margin dari .docx dibaca EXACT dari XML, sedangkan margin dari PDF
    (baik file .pdf asli maupun hasil convert .odt/dst) cuma ESTIMASI
    geometris dari posisi teks -- jadi wajar lebih longgar.

    Cabang .rtf/.doc (2026-09-08) SENGAJA dikembalikan sebagai
    source="docx" juga -- BUKAN "pdf" -- karena setelah convert_to_docx()
    berhasil, params di bawah ini dibaca dari file .docx hasil konversi
    lewat extractor.py yang SAMA PERSIS dipakai untuk .docx asli (bukan
    lewat extract_all_pdf()), jadi margin & struktur lain yang
    dihasilkan memang EXACT dari XML, bukan estimasi geometris --
    menandainya "pdf" di sini justru akan salah membuat comparator.py
    memakai toleransi yang tidak perlu.

    Jalur .pdf & "lainnya" (convert dulu ke PDF) SEKARANG WAJIB lewat
    finalize_page_numbering() sebelum return -- lihat CHANGELOG di atas
    modul ini untuk root-cause lengkap kenapa ini wajib ada di sini
    (bukan cukup ditambahkan di pdf_extractor.py/page_layout_checker.py
    saja): tanpa panggilan ini, params["sections"] & paragraf["section"]
    tetap versi stub lama, dan SEMUA fungsi di bawah run_pipeline()
    (find_section_start_pages, compare_all, check_page_layout) diam-diam
    kembali ke perilaku lama yang selalu salah, walau kode di
    page_layout_checker.py sendiri sudah benar. Jalur .rtf/.doc TIDAK
    lewat sini sama sekali karena ujungnya jadi .docx, dan jalur .docx
    memang tidak butuh finalize_page_numbering() (sudah akurat dari
    XML w:pgNumType).
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext in DOCX_EXTENSIONS:
        doc = parse_docx(filepath)
        params = extract_all(doc)
        pdf_path, font_status = render_to_pdf(filepath)
        return params, pdf_path, "docx", font_status

    if ext in PDF_EXTENSIONS:
        params = extract_all_pdf(filepath)
        params = finalize_page_numbering(params, filepath)
        # Tidak lewat render_to_pdf()/LibreOffice sama sekali -- file
        # ini SUDAH PDF asli (mis. export langsung dari Word oleh
        # mahasiswa), jadi tidak ada risiko substitusi font sama sekali.
        return params, filepath, "pdf", "verified"

    if ext in RTF_DOC_EXTENSIONS:
        # Convert RTF/DOC -> DOCX dulu (BUKAN -> PDF), lalu ikut jalur
        # .docx di atas apa adanya -- lihat CHANGELOG "FIX RTF" untuk
        # alasan lengkap kenapa ini wajib lewat sini, bukan lewat
        # render_to_pdf()/extract_all_pdf() seperti cabang "format
        # lain" di bawah.
        try:
            docx_path = convert_to_docx(filepath)
        except ConversionError as e:
            raise RuntimeError(
                f"Gagal mengonversi {filepath} ke .docx sebelum " f"pengecekan: {e}"
            ) from e
        doc = parse_docx(docx_path)
        params = extract_all(doc)
        pdf_path, font_status = render_to_pdf(docx_path)
        return params, pdf_path, "docx", font_status

    # Format lain (.odt, dst) -- convert dulu ke PDF, lalu ikut jalur
    # .pdf. render_to_pdf() namanya menyebut "docx" di docstring lama,
    # tapi implementasinya (_render_via_libreoffice) cuma menjalankan
    # `soffice --convert-to pdf` apa adanya -- generic untuk format
    # apapun yang didukung LibreOffice, jadi aman dipakai di sini
    # tanpa perlu diubah. .rtf/.doc SUDAH ditangani di cabang khusus
    # di atas dan tidak akan pernah sampai ke sini.
    pdf_path, font_status = render_to_pdf(filepath)
    params = extract_all_pdf(pdf_path)
    params = finalize_page_numbering(params, pdf_path)
    return params, pdf_path, "pdf", font_status


def run_pipeline(filepath: str) -> dict:
    # 1 & 2. Parsing + Extract Parameter (jalur dipilih berdasar format input)
    params, pdf_path, source, font_status = _extract_params(filepath)

    # 2b. page_texts & footer_numbers SELALU dari PDF (baik file aslinya
    #     PDF, atau hasil render/convert) -- dipakai bareng oleh
    #     check_page_layout() (ganjil/genap awal bab) DAN
    #     find_section_start_pages() (lokasi halaman fisik di finding
    #     penomoran halaman). Untuk jalur PDF, params["paragraphs"]
    #     sudah punya elem["section"] yang BENAR di titik ini (dibetulkan
    #     oleh finalize_page_numbering() di _extract_params()), jadi
    #     find_section_start_pages() otomatis nemu SEMUA section (bukan
    #     cuma section 0) tanpa perlu tahu jalur mana yang dipakai.
    #     Untuk jalur .docx (termasuk .rtf/.doc yang sudah dikonversi),
    #     section sudah akurat langsung dari XML.
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
    # font_status diteruskan supaya finding di sini ditandai
    # confidence="rendah" kalau font "Times New Roman" tidak
    # terverifikasi terpasang di server render (lihat CHANGELOG modul
    # ini & page_layout_checker.py).
    layout_findings, layout_checks = check_page_layout(
        page_texts,
        params["sections"],
        section_pages,
        footer_numbers,
        font_status=font_status,
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
        print(
            "Usage: python main.py path/ke/dokumen.[docx|pdf|rtf|doc|odt] [output.json]"
        )
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
