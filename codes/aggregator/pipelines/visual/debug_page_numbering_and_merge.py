"""
debug_page_numbering_and_merge.py
==================================
Script SEKALI PAKAI buat diagnosa KENAPA laporan masih sama persis
setelah fix "Umum (Penomoran Halaman)" & fix margin -- jalankan ini di
laptop Dinda ke PDF ASLI (bukan hasil convert baru, pakai PDF yang sama
persis yang dipakai buat generate laporan JSON kemarin), lalu paste
SELURUH output-nya (bukan PDF-nya) balik ke chat.

CARA PAKAI
----------
1. Taruh file ini SATU FOLDER sama page_layout_checker.py &
   pdf_extractor.py (atau sesuaikan import di bawah ke path project
   kamu -- aggregator.pipelines.visual.*).
2. Ganti PDF_PATH ke path PDF buku TA asli.
3. python debug_page_numbering_and_merge.py
4. Paste seluruh outputnya ke chat.

APA YANG DICEK
--------------
BAGIAN 1 -- Footer/penomoran halaman:
  Panggil extract_footer_details() LANGSUNG (bukan lewat
  finalize_page_numbering()) dan print hasilnya per halaman. Kalau
  dict-nya KOSONG atau isinya jauh lebih sedikit dari total halaman,
  itu tandanya deteksi footer gagal di dokumen ini -- entah karena
  FOOTER_ZONE_PT (70pt) gak nutup posisi footer di dokumen ini, atau
  formatnya beda dari yang diasumsikan (mis. "Halaman 5" bukan cuma
  "5", atau nomor halaman digabung sama teks lain di baris yang sama).

BAGIAN 2 -- Merge paragraf:
  Hitung, dari SATU halaman body-text biasa (bukan cover/TOC), berapa
  blok yang PyMuPDF deteksi MENTAH (sebelum merge) vs berapa yang
  tersisa SETELAH _merge_wrapped_paragraph_blocks() jalan. Kalau
  angkanya SAMA (gak berkurang), berarti merge gak berhasil gabung
  apa-apa -- ini akar penyebab margin kanan nyasar ke 14.93cm (filter
  margin cuma pakai blok multi-baris, kalau paragraf gak pernah
  kegabung jadi multi-baris, kandidatnya kosong/salah).
"""

import sys

# --- SESUAIKAN import ini ke lokasi project kamu kalau perlu ---
from aggregator.pipelines.visual.page_layout_checker import (
    extract_footer_details,
    FOOTER_ZONE_PT,
)
from aggregator.pipelines.visual.pdf_extractor import (
    _extract_blocks,
    _merge_wrapped_paragraph_blocks,
    extract_section_params_pdf,
)

import fitz  # pip install pymupdf

PDF_PATH = "GANTI_DENGAN_PATH_PDF_ASLI.pdf"

# Halaman body-text BIASA (bukan cover/TOC/daftar gambar) buat dites
# merge-nya -- ganti ke halaman yang KAMU TAHU isinya beberapa paragraf
# body normal.
BODY_TEST_PAGE = 20  # 1-indexed


def main():
    print("=" * 70)
    print("BAGIAN 1: extract_footer_details()")
    print("=" * 70)
    details = extract_footer_details(PDF_PATH)
    with fitz.open(PDF_PATH) as pdf:
        total_pages = pdf.page_count
    print(f"Total halaman fisik: {total_pages}")
    print(f"Halaman yang footernya BERHASIL dibaca meyakinkan: {len(details)}")
    if not details:
        print(">> KOSONG SAMA SEKALI. Ini konfirmasi akar masalahnya:")
        print("   extract_footer_details() gagal nemu SATU PUN footer yang")
        print("   meyakinkan di dokumen ini. Kemungkinan penyebab:")
        print("   a) FOOTER_ZONE_PT (70pt) gak nutup posisi footer asli --")
        print("      cek 'BAGIAN 1b' di bawah buat lihat posisi y token")
        print("      paling bawah tiap halaman.")
        print("   b) Nomor halaman digabung sama teks lain di baris yang")
        print("      sama (mis. 'Halaman 5' atau '- 5 -'), gak match regex")
        print("      angka/romawi polos.")
    else:
        for p in sorted(details)[:15]:
            d = details[p]
            print(
                f"  hal.{p}: value={d['value']} fmt={d['fmt']} centered={d['centered']}"
            )
        if len(details) > 15:
            print(f"  ... ({len(details) - 15} halaman lainnya)")

    print()
    print("BAGIAN 1b: posisi token PALING BAWAH di tiap halaman (10 halaman")
    print("pertama) -- buat cek apa FOOTER_ZONE_PT (70pt dari bawah) nutup")
    print("posisi footer asli atau ENGGAK.")
    print("-" * 70)
    with fitz.open(PDF_PATH) as pdf:
        for i, page in enumerate(pdf):
            if i >= 10:
                break
            page_height = page.rect.height
            words = page.get_text("words")
            if not words:
                print(f"  hal.{i+1}: (tidak ada teks sama sekali)")
                continue
            lowest = max(words, key=lambda w: w[1])
            dist_from_bottom = page_height - lowest[1]
            in_zone = dist_from_bottom <= FOOTER_ZONE_PT
            print(
                f"  hal.{i+1}: token terbawah={lowest[4]!r} "
                f"jarak_dari_tepi_bawah={dist_from_bottom:.1f}pt "
                f"{'(MASUK zona 70pt)' if in_zone else '(DI LUAR zona 70pt!)'}"
            )

    print()
    print("=" * 70)
    print(f"BAGIAN 2: merge paragraf di halaman {BODY_TEST_PAGE}")
    print("=" * 70)
    sections = extract_section_params_pdf(PDF_PATH)
    left_margin_pt = sections[0]["left_twips"] / 20
    right_margin_pt = sections[0]["width_twips"] / 20 - sections[0]["right_twips"] / 20

    with fitz.open(PDF_PATH) as pdf:
        page = pdf[BODY_TEST_PAGE - 1]
        raw_blocks = _extract_blocks(page)
        merged_blocks = _merge_wrapped_paragraph_blocks(
            list(raw_blocks), left_margin_pt, right_margin_pt
        )

    print(f"Blok MENTAH dari PyMuPDF (sebelum merge): {len(raw_blocks)}")
    print(f"Blok SETELAH _merge_wrapped_paragraph_blocks(): {len(merged_blocks)}")
    if len(merged_blocks) == len(raw_blocks):
        print(">> TIDAK ADA satupun blok yang kegabung. Merge gagal total di")
        print("   halaman ini -- ini kemungkinan besar penyebab margin kanan")
        print("   nyasar ke nilai aneh (14.93cm), karena extract_section_params_pdf()")
        print("   cuma percaya blok >=2 baris sebagai kandidat margin, dan")
        print("   kalau semua paragraf tetap 1-baris, kandidatnya jadi salah/kosong.")
    else:
        print(f">> {len(raw_blocks) - len(merged_blocks)} blok berhasil kegabung.")
    print()
    print("Detail tiap blok MENTAH di halaman ini (index, jumlah baris, preview teks):")
    for i, b in enumerate(raw_blocks):
        preview = b["full_text"][:60].replace("\n", " ")
        print(f"  [{i}] {len(b['line_bboxes'])} baris -- {preview!r}")


if __name__ == "__main__":
    main()
