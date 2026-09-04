"""
config.py
=========
Rule config (threshold) sesuai Panduan Buku PA Sarjana Terapan PENS.
Dipisah biar mudah di-update kalau ada revisi panduan, tanpa
menyentuh logika parsing/comparator.
"""

import os

from docx.shared import Cm


def cm_to_twips(cm_value: float) -> int:
    """
    Konversi cm -> twips (1440 twips = 1 inch), dibulatkan ke integer
    terdekat.

    Dipakai sebagai basis PERBANDINGAN margin & ukuran kertas (bukan
    cm), karena twips itu unit ASLI yang tersimpan di XML docx
    (w:pgMar, w:pgSz semuanya twips) -- membandingkan langsung di
    twips menghindari noise floating-point dari konversi EMU->cm
    (mis. 3.9999998 vs 4.0 cm) dan lebih dekat ke representasi
    sumbernya. RULES (dalam cm) tetap dipertahankan apa adanya untuk
    keterbacaan config & buat nyusun pesan error ke user; RULES_TWIPS
    di bawah yang dipakai comparator.py untuk logika perbandingan
    aktualnya. Lihat CHANGELOG di comparator.compare_sections().
    """
    return round(Cm(cm_value).twips)


RULES = {
    "font_name": "Times New Roman",
    "body_font_size": 12,
    "chapter_title_size": 14,
    "subchapter_title_size": 12,
    "caption_font_size": 11,
    "line_spacing": 1.5,
    "margin_left_cm": 4.0,
    "margin_other_cm": 3.0,  # top, bottom, right
    "margin_tolerance_cm": 0.1,
    "page_width_cm": 21.0,  # A4, potrait: 21 x 29.7. Dicek orientasi-agnostik
    "page_height_cm": 29.7,  # (short/long side), jadi landscape tetap valid.
    "page_size_tolerance_cm": 0.2,
    "mirror_margins": True,  # cetak bolak-balik (mirror margin kiri/kanan)
    # --- Abstrak/Abstract (cek FORMAT VISUAL saja) ---
    # Panduan juga mensyaratkan isi tertentu (latar belakang, permasalahan,
    # solusi, kinerja) -- itu SENGAJA tidak dicek di sini karena butuh
    # pemahaman semantik/konten, bukan hitungan/format visual. Bagian
    # semacam itu ranahnya pipeline sistematika (RAG), bukan comparator
    # rule-based ini.
    "abstract_min_words": 300,
    "abstract_max_words": 500,
    "abstract_max_keywords": 5,  # Kata Kunci / Keywords maksimum 5 frasa
    # --- Kutipan langsung (direct quotation) ---
    # Hanya berlaku untuk kutipan langsung PANJANG (>2 baris, ditulis
    # sebagai paragraf terpisah) -- lihat quotation_checker.py &
    # compare_quotations() di comparator.py. Kutipan langsung PENDEK
    # (menyatu di paragraf) tidak punya aturan font/indent sendiri,
    # cuma wajib italic.
    "direct_quote_long_font_size": 10,
    "direct_quote_indent_cm": 1.0,  # 10mm masuk dari batas kiri & kanan
    "direct_quote_indent_tolerance_cm": 0.2,
}

# Versi TWIPS dari rule margin & ukuran kertas -- dihitung SEKALI di sini
# (saat modul ini pertama di-import), bukan berulang tiap kali
# compare_sections() dipanggil. RULES tetap jadi single source of truth
# (nilai dalam cm, gampang di-update kalau panduan revisi); RULES_TWIPS
# cuma turunannya buat perbandingan, comparator.py TIDAK PERNAH hardcode
# angka twips sendiri.
RULES_TWIPS = {
    "margin_left": cm_to_twips(RULES["margin_left_cm"]),
    "margin_other": cm_to_twips(RULES["margin_other_cm"]),
    "margin_tolerance": cm_to_twips(RULES["margin_tolerance_cm"]),
    "page_width": cm_to_twips(RULES["page_width_cm"]),
    "page_height": cm_to_twips(RULES["page_height_cm"]),
    "page_size_tolerance": cm_to_twips(RULES["page_size_tolerance_cm"]),
}

# Path data untuk pipeline sistematika (RAG). Dipisah dari RULES karena ini
# bukan aturan pemeriksaan, melainkan lokasi file/folder data. Pakai
# os.getenv supaya bisa di-override tanpa ubah kode (misal beda path saat
# testing vs production, atau saat panduan direvisi ke Rev06 dst).
#
# BASE_DIR = folder tempat config.py ini berada (aggregator/). Dipakai
# sebagai basis path absolut supaya path data selalu ketemu, gak peduli
# script dijalankan dari folder mana pun (cwd tidak berpengaruh).
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CHECKRUN_DIR = os.path.join(BASE_DIR, "CheckRun")

PATHS = {
    "guideline_pdf": os.getenv(
        "GUIDELINE_PDF_PATH",
        os.path.join(
            CHECKRUN_DIR,
            "data",
            "guidelines",
            "Panduan_Buku_PA_SarjanaTerapan_A4_Rev05 (1).pdf",
        ),
    ),
    "chroma_dir": os.getenv(
        "CHROMA_DIR", os.path.join(CHECKRUN_DIR, "data", "chroma_store")
    ),
    "chroma_collection": os.getenv("CHROMA_COLLECTION", "panduan_sistematika"),
}
