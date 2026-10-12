"""
comparator.py
=============
TAHAP 3: RULE-BASED COMPARATOR
--------------------------------
Menerima parameter mentah dari extractor.py, membandingkan dengan
RULES dari config.py, dan menghasilkan findings.

CATATAN terkait revisi extractor.py:
- ptype "subsubchapter_title" sekarang mencakup DUA sumber: level
  "X.Y.Z" bernomor (baik manual maupun via numbering Word) DAN heading
  tanpa nomor (mis. "Ekstraksi Parameter"). Comparator tidak perlu tahu
  bedanya -- keduanya memang harus tunduk pada aturan format yang SAMA:
  Title Case (elem["is_titlecase"]) + bold + ukuran font sama dengan
  subchapter_title_size. Ini sesuai instruksi panduan: "Setiap bagian
  inti (latar belakang, dst) ditulis dengan huruf kapital di awal tiap
  kata dan dicetak tebal."
- font_name & font_size_pt dari extractor.py sekarang sudah ditelusuri
  sampai ke style (tidak None lagi kecuali benar-benar tidak ada
  definisinya di mana pun), sehingga pengecekan di bawah ini otomatis
  mencakup kasus yang dulu ke-skip tanpa perlu ubah logika di sini.
- Ditambah compare_abstract(): cek format visual Abstrak (ID) & Abstract
  (EN) -- jumlah kata abstrak (300-500), jumlah frasa Kata Kunci/Keywords
  (maks 5), dan wajib-italic khusus Abstract (EN). Berbasis field
  elem["abstract_part"] baru dari extractor.py. Sengaja TIDAK mengecek
  isi/konten (latar belakang, permasalahan, dst) -- itu ranah pipeline
  sistematika (RAG) yang butuh pemahaman semantik, bukan comparator
  rule-based ini.
- compare_visual_objects(): GAMBAR sekarang JUGA dicek wajib punya
  caption ("Gambar N.N ...") menempel di salah satu sisinya (atas/
  bawah), sama seperti TABEL yang sudah lebih dulu dicek begini --
  sebelumnya gambar cuma dicek rata-tengahnya, caption-nya sendiri
  cuma divalidasi format PENOMORANNYA (compare_caption_numbering())
  KALAU ada, tapi tidak pernah dicek WAJIB ADA-nya.
- Ditambah compare_quotations(): cek format kutipan LANGSUNG (pendek &
  panjang) berbasis quotation_checker.check_paragraph() yang baru --
  lihat modul itu untuk detail deteksi & disambiguasi dari istilah/
  dialog biasa yang kebetulan diberi tanda kutip. Kutipan TIDAK
  LANGSUNG (parafrase) SENGAJA TIDAK dicek -- lihat docstring
  compare_quotations() untuk alasannya.
"""

import re
import bisect
from collections import defaultdict
from aggregator.config import RULES, RULES_TWIPS

# DINONAKTIFKAN (atas permintaan user, 2026-08-21): cek "istilah asing
# wajib italic" di-nonaktifkan sementara -- lihat blok yang dikomentari
# di compare_paragraphs() di bawah. Import-nya ikut dikomentari supaya
# tidak ada import yang menggantung tak terpakai.
# from aggregator.pipelines.visual.foreign_term_checker import check_paragraph
from aggregator.pipelines.visual.quotation_checker import (
    check_paragraph as check_quotations,
)
from aggregator.pipelines.visual.extractor import (
    EQUATION_NUMBER_IN_TEXT_RE,
    normalize_equation_text,
)


def text_preview(text: str, length: int = 40) -> str:
    text = text.strip()
    return text[:length] + ("..." if len(text) > length else "")


# PENYEBAB INKONSISTENSI ANTAR-FORMAT (ditemukan 2026-09-07, dari
# perbandingan hasil docx vs pdf vs rtf pada dokumen yang SAMA persis):
# jalur RTF/PDF sering melapor font "Liberation Serif" / "Liberation
# Sans" / "Liberation Mono" padahal di file .docx aslinya (dan di XML
# style-nya) font yang di-set sudah benar "Times New Roman" / "Arial" /
# "Courier New". Ini BUKAN kesalahan mahasiswa -- ini artefak proses
# konversi: LibreOffice (dipakai untuk render/convert ke RTF & PDF di
# server) mengganti otomatis font Windows yang tidak ter-install di
# server dengan pengganti metric-compatible dari paket "Liberation
# Fonts" (dibuat memang supaya lebar karakter identik, cuma nama
# family-nya beda). _normalize_font_name() di pdf_extractor.py sudah
# menangani variasi nama PostScript ("TimesNewRomanPSMT" dst), tapi itu
# tidak menyentuh kasus ini karena "Liberation Serif" bukan varian nama
# PostScript dari "Times New Roman" -- itu FAMILY FONT YANG BENAR-BENAR
# BEDA secara nama. Tanpa pemetaan ini, dokumen yang formatnya sudah
# 100% benar bisa jatuh skornya HANYA gara-gara font apa yang kebetulan
# terpasang di mesin yang menjalankan konversi, bukan gara-gara isi
# dokumennya -- inilah salah satu penyebab utama skor RTF vs DOCX beda
# padahal sumbernya sama.
FONT_SUBSTITUTE_ALIASES = {
    "times new roman": {"liberation serif", "tinos", "nimbus roman"},
    "arial": {"liberation sans", "arimo", "nimbus sans"},
    "courier new": {"liberation mono", "cousine", "nimbus mono", "nimbus mono ps"},
    "calibri": {"carlito"},
    "cambria": {"caladea"},
}


def _font_matches_expected(actual_font: str, expected_font: str) -> bool:
    """
    True kalau actual_font == expected_font ATAU actual_font adalah
    pengganti metric-compatible yang DIKENAL untuk expected_font (lihat
    FONT_SUBSTITUTE_ALIASES). Case-insensitive karena nama font yang
    dibaca dari PDF/RTF kadang beda kapitalisasi dengan yang di docx.
    """
    if actual_font == expected_font:
        return True
    a = actual_font.strip().lower()
    e = expected_font.strip().lower()
    if a == e:
        return True
    return a in FONT_SUBSTITUTE_ALIASES.get(e, set())


EXPECTED_SIZE_BY_TYPE = {
    "chapter_title": RULES["chapter_title_size"],
    "subchapter_title": RULES["subchapter_title_size"],
    # panduan tidak menetapkan ukuran font terpisah untuk level "X.Y.Z"
    # (bernomor maupun heading tanpa nomor), jadi disamakan dengan
    # level "X.Y" (subchapter_title_size)
    "subsubchapter_title": RULES["subchapter_title_size"],
    # Nama lampiran (baris setelah "LAMPIRAN <huruf>") -- level
    # hierarkinya setara chapter_title (lihat LAMPIRAN_TITLE_RE di
    # extractor.py), jadi ukuran fontnya sama, TAPI ptype-nya dipisah
    # supaya bisa dikecualikan dari cek ALL CAPS di bawah (pedoman
    # PENS tidak mensyaratkan nama lampiran ditulis kapital semua).
    "lampiran_title": RULES["chapter_title_size"],
    "caption": RULES["caption_font_size"],
    "body": RULES["body_font_size"],
}


def _new_counter():
    return {"checked": 0, "failed": 0}


# DIKECUALIKAN TOTAL dari semua pengecekan format (atas permintaan
# user, 2026-09-17): LAMPIRAN (semua lampiran, "LAMPIRAN A", "LAMPIRAN
# B", dst -- current_chapter sudah dinormalisasi ke bentuk ini oleh
# extractor.py/pdf_extractor.py, lihat LAMPIRAN_TITLE_RE) dan BIODATA
# PENULIS. Isi lampiran umumnya dokumen pendukung (source code, hasil
# kuesioner, dsb) yang formatnya di luar kendali aturan skripsi baku,
# dan BIODATA PENULIS adalah halaman penutup singkat yang juga bebas
# format -- keduanya tidak lagi relevan dikenai aturan font/ukuran/
# bold/perataan/caption/kutipan. Dipakai di semua compare_* di modul
# ini yang mengiterasi per elemen/chapter, supaya cakupannya konsisten
# untuk jalur docx maupun pdf sekaligus (keduanya sudah dinormalisasi
# ke field "chapter" yang sama oleh extractor.py/pdf_extractor.py,
# comparator.py ini dipakai bersama oleh kedua jalur).
def _is_excluded_chapter(chapter: str | None) -> bool:
    if not chapter:
        return False
    if chapter == "BIODATA PENULIS":
        return True
    if chapter.startswith("LAMPIRAN"):
        return True
    return False


def compare_paragraphs(paragraphs: list[dict], source: str = "docx"):
    """
    source: "docx" | "pdf" -- default "docx" supaya pemanggil lama yang
    belum diupdate kasih argumen ini perilakunya TIDAK BERUBAH. Dipakai
    KHUSUS untuk membatasi cek line-spacing ke jalur docx saja -- lihat
    catatan "DINONAKTIFKAN KHUSUS PDF" di bawah.

    Paragraf yang chapter-nya LAMPIRAN (mana pun) atau BIODATA PENULIS
    dilewati total (lihat _is_excluded_chapter()), berlaku sama untuk
    jalur docx maupun pdf.
    """
    findings = []
    checks = defaultdict(_new_counter)

    for elem in paragraphs:
        ptype = elem["type"]
        if ptype == "cover":
            continue  # halaman sampul tidak diperiksa
        if ptype == "toc_entry":
            # Baris di halaman Daftar Isi/Gambar/Tabel (heading-nya
            # sendiri maupun entrinya) -- lihat extractor.py untuk
            # alasan lengkap kenapa ini dikecualikan total, bukan cuma
            # di-demote ke "body".
            continue
        if ptype == "notice":
            # "Halaman ini sengaja dikosongkan" -- bukan body text.
            continue
        if ptype in ("image", "equation", "source_citation"):
            # Gambar/persamaan/sumber-kutipan tidak tunduk ke aturan
            # font/size/bold judul maupun body biasa -- ditangani
            # dedicated checker compare_visual_objects() di bawah.
            continue

        chapter = elem["chapter"]
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        loc = f"Paragraf #{elem['index']} ('{text_preview(elem['full_text'])}')"
        expected_size = EXPECTED_SIZE_BY_TYPE[ptype]

        # --- font name: 1 paragraf = 1 unit cek, bukan 1 run = 1 unit cek.
        #     Paragraf panjang dengan banyak run (mis. banyak istilah asing
        #     yang di-italic, atau teks hasil copy-paste yang membawa font
        #     tersembunyi berbeda) tidak boleh "membanjiri" angka checked
        #     dibanding paragraf pendek yang cuma 1 run. Kesalahan tetap
        #     dilaporkan spesifik (nilai font apa yang salah), cuma
        #     dikumpulkan dulu per paragraf sebelum dicatat sebagai findings.
        # Run yang berisi variabel/simbol matematis (mis. huruf italic
        # matematis "n" berfont Cambria Math yang diketik sebagai teks
        # biasa, BUKAN objek OMML utuh sehingga tidak tertangkap sebagai
        # ptype "equation") dikecualikan dari cek font di sini -- lihat
        # is_math/_is_math_run_text()/MATH_FONT_NAMES di extractor.py.
        # Dikecualikan PER-RUN (bukan per-paragraf/ptype) supaya sisa
        # kalimat body di paragraf yang sama tetap kena cek font normal.
        #
        # BUG DIPERBAIKI (jalur PDF -- "Font 'Arial' tidak sesuai" /
        # "Ukuran font 11.0pt tidak sesuai" padahal dokumen aslinya 100%
        # Times New Roman 12pt): run yang isinya CUMA whitespace (satu
        # spasi kosong) TIDAK dikecualikan dari cek ini. PyMuPDF/LibreOffice
        # sering merender artefak spasi tunggal di antara nomor subbab dan
        # judulnya (mis. antara "1.1" dan "LATAR BELAKANG") memakai font
        # /ukuran BERBEDA dari font asli paragraf (terverifikasi empiris:
        # SELURUH run non-whitespace di dokumen konsisten Times New Roman,
        # font "Arial"/"Calibri" & size 11pt HANYA pernah muncul di run
        # whitespace-only) -- sisa artefak numbering list Word yang ikut
        # terekspor, bukan font/ukuran teks yang sesungguhnya diketik
        # mahasiswa. Tanpa filter ini, satu spasi artefak itu membuat
        # SELURUH paragraf salah tandai "font/ukuran tidak sesuai" padahal
        # isinya benar -- inilah penyebab skor Font & Ukuran Huruf jatuh
        # drastis di jalur PDF dibanding jalur docx. Filter whitespace-only
        # ini SENGAJA tidak diterapkan di level ekstraksi (pdf_extractor.py)
        # supaya "".join(runs) == full_text tetap terjaga untuk checker lain
        # (quotation_checker.py, foreign_term_checker.py) yang butuh
        # rekonstruksi teks utuh -- cukup dikecualikan di titik pemakaian
        # (cek font/size) di comparator.py ini.
        runs_with_font_name = [
            r
            for r in elem["runs"]
            if r["font_name"] is not None
            and not r.get("is_math")
            and r["text"].strip() != ""
        ]
        # Sama seperti font size di bawah (lihat CHANGELOG di sana): pakai
        # font MAYORITAS paragraf (character-weighted) sebagai representasi
        # paragraf, bukan "any run fails" -- artefak pendek (nomor list
        # otomatis, nomor SEQ caption) yang ke-render pakai font beda tidak
        # boleh membuat SELURUH paragraf yang sudah benar ikut ke-flag.
        if runs_with_font_name:
            checks[chapter]["checked"] += 1
            font_char_weight = defaultdict(int)
            for r in runs_with_font_name:
                font_char_weight[r["font_name"]] += len(r["text"].strip())
            dominant_font = max(font_char_weight, key=font_char_weight.get)
            if not _font_matches_expected(dominant_font, RULES["font_name"]):
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": f"Font '{dominant_font}' tidak sesuai, seharusnya '{RULES['font_name']}'.",
                    }
                )

        # --- font size: sama prinsipnya, 1 paragraf = 1 unit cek ---
        # Filter whitespace-only run yang sama seperti runs_with_font_name
        # di atas -- lihat CHANGELOG-nya untuk root-cause lengkap (run
        # spasi artefak numbering list ikut bawa ukuran yang salah, mis.
        # 11pt nyempil di antara run 12pt yang benar).
        #
        # BUG DIPERBAIKI (2026-09-20, permintaan user): run matematis
        # (is_math -- lihat filter runs_with_font_name di atas & catatan
        # panjang MATH_FONT_NAMES/_is_math_run_text di extractor.py)
        # SUDAH dikecualikan dari cek NAMA font sejak awal, tapi belum
        # ikut dikecualikan dari cek UKURAN font di sini. Akibatnya
        # rumus yang mestinya bebas ukuran (mis. subscript/superscript
        # notasi matematis yang wajar lebih kecil dari body text, atau
        # simbol yang di-render Word/Cambria Math dengan metric beda
        # dari body) tetap ke-flag "Ukuran font X pt tidak sesuai...
        # seharusnya 12pt" walau font-nya sendiri sudah tidak lagi
        # disalahkan. Filter is_math ditambahkan di sini juga supaya
        # perlakuannya KONSISTEN: run matematis bebas dari aturan
        # font DAN ukuran body sekaligus, bukan cuma salah satunya.
        # BUG DIPERBAIKI (jalur PDF -- "Ukuran font 10.0pt/12.0pt tidak
        # sesuai" padahal seluruh isi paragraf di dokumen asli SUDAH benar
        # sesuai aturan): sebelumnya SATU run mana pun yang ukurannya beda
        # langsung membuat SELURUH paragraf ditandai gagal ("any run
        # fails"). Ini kena FALSE POSITIVE oleh artefak Word yang
        # auto-generated -- BUKAN teks yang benar-benar diketik mahasiswa:
        #   - nomor list bernomor otomatis ("1.", "2.", "3." di awal item
        #     list Word) kadang di-render PDF dengan ukuran font BEDA dari
        #     body text-nya sendiri (mis. paragraf "3. Perintah Kendali
        #     Manual ROV (7): ..." isinya 100% 12pt, tapi angka "3."-nya
        #     ke-render 10pt).
        #   - nomor caption hasil SEQ field Word ("Gambar 3. 7 ...") kadang
        #     ke-render dengan ukuran beda dari teks caption di sebelahnya
        #     (mis. isi caption sudah benar 11pt, tapi angka nomornya
        #     ke-render 12pt).
        # Sama seperti filter whitespace-only run yang sudah ada, artefak
        # ini bukan representasi ukuran font yang sungguhan diketik. Karena
        # panjangnya cuma beberapa karakter (bukan mayoritas isi paragraf),
        # solusinya: pakai ukuran MAYORITAS paragraf (character-weighted,
        # bukan run-weighted -- run pendek tidak boleh setara run panjang)
        # sebagai representasi paragraf, BUKAN "any run fails". Paragraf
        # yang genuinely salah ukuran (mayoritas isinya memang bukan
        # expected_size) tetap ke-flag seperti biasa.
        runs_with_size = [
            r
            for r in elem["runs"]
            if r["font_size_pt"] is not None
            and not r.get("is_math")
            and r["text"].strip() != ""
        ]
        if runs_with_size:
            checks[chapter]["checked"] += 1
            size_char_weight = defaultdict(int)
            for r in runs_with_size:
                size_char_weight[r["font_size_pt"]] += len(r["text"].strip())
            dominant_size = max(size_char_weight, key=size_char_weight.get)
            if abs(dominant_size - expected_size) > 0.1:
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": f"Ukuran font {dominant_size}pt tidak sesuai untuk tipe '{ptype}', seharusnya {expected_size}pt.",
                    }
                )

        # --- bold untuk judul bab/subbab: 1 paragraf = 1 unit cek ---
        # BUG DIPERBAIKI (jalur PDF -- judul yang di dokumen aslinya SUDAH
        # bold penuh tetap ditandai "seharusnya bold" salah): sebelumnya
        # `any(run["bold"] is not True for run in elem["runs"])` diuji ke
        # SEMUA run tanpa kecuali, termasuk run spasi/whitespace-only
        # (baik spasi sintetis hasil merge baris di pdf_extractor.py,
        # maupun span spasi asli dari PyMuPDF). Sama seperti bug font/size
        # whitespace-run yang sudah diperbaiki di atas: PyMuPDF/LibreOffice
        # kerap merender span spasi tunggal di antara kata dengan flag bold
        # TIDAK ikut ke-set (glyph spasi tidak "mewarisi" flag style run
        # tetangganya secara konsisten), padahal seluruh KARAKTER tercetak
        # di judul tsb sudah bold. Akibatnya satu spasi artefak membuat
        # SELURUH judul salah ditandai "belum bold" walau sudah benar.
        # Filter run whitespace-only di sini -- konsisten dengan
        # runs_with_font_name/runs_with_size di atas -- supaya cek bold
        # cuma menilai karakter yang benar-benar tercetak.
        if ptype in (
            "chapter_title",
            "subchapter_title",
            "subsubchapter_title",
            "lampiran_title",
        ):
            runs_with_text_for_bold = [
                r for r in elem["runs"] if r["text"].strip() != ""
            ]
            if runs_with_text_for_bold:
                checks[chapter]["checked"] += 1
                if any(run["bold"] is not True for run in runs_with_text_for_bold):
                    checks[chapter]["failed"] += 1
                    findings.append(
                        {
                            "category": "Format",
                            "chapter": chapter,
                            "location": loc,
                            "description": f"Judul tipe '{ptype}' seharusnya dicetak tebal (bold) secara keseluruhan.",
                        }
                    )

        # --- rata tengah untuk judul setingkat bab (BAB N, DAFTAR ISI,
        #     DAFTAR GAMBAR, DAFTAR TABEL, DAFTAR PUSTAKA, KATA PENGANTAR,
        #     ABSTRAK/ABSTRACT, dst) ---
        # SENGAJA cuma "chapter_title", TIDAK termasuk subchapter_title/
        # subsubchapter_title -- judul level "1.1"/"1.1.1" konvensinya
        # rata KIRI, bukan rata tengah, jadi tidak boleh ikut kena cek
        # ini. Sebelumnya tidak ada cek center sama sekali untuk ptype
        # ini, jadi judul yang tidak center (mis. "DAFTAR TABEL" pada
        # gambar yang dilaporkan user) lolos tanpa pernah dicek.
        #
        # DIAKTIFKAN KEMBALI UNTUK JALUR PDF (2026-09-09, permintaan
        # user eksplisit). RIWAYAT: sempat DIMATIKAN TOTAL untuk
        # source=="pdf" karena elem["alignment"] versi PDF cuma
        # ESTIMASI dari posisi bbox relatif terhadap margin kolom hasil
        # inferensi statistik (median jarak baris body ke tepi kanan/
        # kiri di SELURUH dokumen -- lihat _infer_alignment()/
        # extract_section_params_pdf di pdf_extractor.py), BUKAN dibaca
        # langsung dari atribut seperti w:jc di docx, dan estimasi itu
        # terbukti bisa meleset cukup untuk over-flag heading yang
        # sebenarnya SUDAH rata tengah. Batasan itu MASIH ADA (belum
        # ada perbaikan baru di sisi estimasi margin) -- kalau muncul
        # over-flagging massal lagi (heading yang jelas rata tengah di
        # dokumen tapi ke-flag di hampir semua bab sekaligus), itu
        # gejala keterbatasan yang sama, bukan bug baru. Opsi
        # mitigasinya kalau itu terjadi: (a) matikan lagi khusus
        # source=="pdf" seperti semula, atau (b) turunkan severity-nya
        # jadi "Info" (bukan "Format") untuk PDF supaya tidak memotong
        # skor walau tetap dilaporkan.
        # DIHAPUS (atas permintaan user eksplisit, 2026-09-12): cek rata
        # tengah untuk chapter_title dimatikan lagi, berlaku sama untuk
        # docx maupun pdf. Alasan: false-positive pada dokumen yang
        # judulnya SUDAH rata tengah di dokumen asli, tapi salah
        # terdeteksi lewat jalur PDF (lihat penjelasan panjang di atas
        # soal elem["alignment"] versi PDF cuma ESTIMASI dari bbox,
        # bukan atribut w:jc yang dibaca langsung) -- kasus nyata:
        # "ABSTRACT" & "DAFTAR TABEL" ke-flag padahal sudah rata tengah
        # di dokumen aslinya. Blok lama dikomentari (bukan dihapus
        # fisik) untuk referensi/gampang diaktifkan lagi kalau estimasi
        # alignment PDF-nya diperbaiki di masa depan.
        # if ptype == "chapter_title":
        #     checks[chapter]["checked"] += 1
        #     if elem["alignment"] != "center":
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": f"Judul tipe '{ptype}' seharusnya ditulis rata tengah.",
        #             }
        #         )

        # --- rata tengah untuk caption gambar/tabel ---
        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): sama seperti
        # cek persamaan di compare_visual_objects(), dimatikan tapi
        # dipertahankan (dikomentari) supaya gampang diaktifkan lagi.
        # if ptype == "caption":
        #     checks[chapter]["checked"] += 1
        #     if elem["alignment"] != "center":
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": (
        #                     "Nomor identitas dan keterangan gambar/tabel "
        #                     "seharusnya ditulis rata tengah."
        #                 ),
        #             }
        #         )

        # --- kapitalisasi judul: ALL CAPS untuk level "X.Y" ---
        if ptype in ("chapter_title", "subchapter_title"):
            checks[chapter]["checked"] += 1
            if not elem["is_upper"]:
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": f"Judul tipe '{ptype}' seharusnya ditulis kapital semua.",
                    }
                )

        # --- kapitalisasi judul: Title Case untuk level "X.Y.Z" (bernomor
        #     maupun heading tanpa nomor yang disatukan ke ptype ini) ---
        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): dikomentari,
        # bukan dihapus -- sub-subbab (contoh "1.1.1 HTML5") diberi
        # kelonggaran, tidak wajib Title Case.
        # if ptype == "subsubchapter_title":
        #     checks[chapter]["checked"] += 1
        #     if not elem["is_titlecase"]:
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": (
        #                     f"Judul tipe '{ptype}' seharusnya ditulis kapital di awal "
        #                     f"setiap kata (Title Case), bukan ALL CAPS ataupun huruf "
        #                     f"kecil semua."
        #                 ),
        #             }
        #         )

        # --- line spacing untuk body text ---
        # DINONAKTIFKAN LAGI (atas permintaan user, 2026-09-07, hari yang
        # sama sempat diaktifkan): sempat diaktifkan penuh, lalu dibatasi
        # cuma jalan di docx (source != "pdf") karena elem["line_spacing"]
        # dari pdf_extractor.py banyak false-positive -- BLOK HASIL MERGE
        # yang sebenarnya BUKAN paragraf body 2-baris asli, melainkan dua
        # elemen visual terpisah yang kebetulan berdekatan (mis. baris
        # sitasi/URL yang ke-merge dengan caption gambar di sebelahnya,
        # atau judul sub-subbab yang ke-merge dengan body di bawahnya)
        # kebaca sbg "line spacing 1.0" padahal bukan spasi baris body
        # text sungguhan. Tapi user lalu menilai itu TIDAK ADIL kalau
        # cuma docx yang kena cek sementara pdf tidak -- mahasiswa yang
        # submit docx & pdf jadi dinilai dgn standar berbeda utk aturan
        # yang sama. Jadi dinonaktifkan TOTAL lagi utk KEDUA source,
        # bukan cuma pdf -- dikomentari, bukan dihapus, pola yang sama
        # seperti cek-cek lain yang dimatikan di modul ini.
        #
        # TODO kalau mau diaktifkan lagi: HARUS utk docx & pdf SEKALIGUS
        # (supaya adil), jadi perbaiki dulu akar masalah di
        # pdf_extractor.py -- blok hasil _merge_wrapped_paragraph_
        # blocks() perlu ditandai/dipisahkan kalau sebenarnya gabungan
        # elemen heterogen (sitasi+caption, heading+body dst), bukan
        # paragraf body asli, SEBELUM line_spacing dihitung dari situ.
        # if ptype == "body" and elem["line_spacing"] is not None:
        #     checks[chapter]["checked"] += 1
        #
        #     if not elem["line_spacing_is_multiple"]:
        #         # Rule EXACTLY/AT_LEAST: paragraf pakai spacing tetap (pt),
        #         # bukan mode kelipatan (multiple) sama sekali -- ini salah
        #         # secara MODE, terlepas dari berapa pun nilai pt-nya, jadi
        #         # gak perlu (dan gak masuk akal) dibandingin sebagai angka.
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": (
        #                     f"Spasi baris pakai mode tetap/fixed ({elem['line_spacing']:.1f}pt, "
        #                     f"rule={elem['line_spacing_rule']}), seharusnya mode kelipatan "
        #                     f"(multiple) {RULES['line_spacing']}."
        #                 ),
        #             }
        #         )
        #     elif abs(elem["line_spacing"] - RULES["line_spacing"]) > 0.01:
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": f"Spasi baris {elem['line_spacing']} tidak sesuai, seharusnya {RULES['line_spacing']}.",
        #             }
        #         )

        # --- istilah asing wajib italic: khusus body text, karena judul/
        #     caption punya aturan kapitalisasi sendiri yang mendominasi
        #     tampilannya. DAFTAR PUSTAKA sengaja DIKECUALIKAN TOTAL --
        #     entri di situ adalah kutipan judul artikel/jurnal berbahasa
        #     Inggris apa adanya (dalam tanda kutip), yang secara konvensi
        #     sitasi TIDAK di-italic per kata; kalau tidak dikecualikan,
        #     kata sambung/fungsi di judul kutipan ("and", "for", "the",
        #     dst) akan salah dianggap istilah asing yang perlu italic.
        #     check_paragraph() merekonstruksi paragraf jadi satu string
        #     utuh dulu (bukan run per run) supaya kata yang kebetulan
        #     terpotong di batas run (format beda di tengah kata, biasa
        #     dari copy-paste) tetap terbaca sebagai satu kata -- lihat
        #     foreign_term_checker.py untuk detail & daftar exclusion
        #     (stopword Inggris, proper noun/nama tool, dst).
        #
        #     Paragraf Abstract (EN) -- abstract_part "abstract_body"/
        #     "keyword_en" -- JUGA DIKECUALIKAN dari cek per-kata ini,
        #     walau bukan DAFTAR PUSTAKA. Alasannya beda dari DAFTAR
        #     PUSTAKA: paragraf ini memang teks Inggris UTUH yang wajib
        #     italic SELURUHNYA, dan itu sudah dicek sebagai SATU unit
        #     (per paragraf) di compare_abstract() -- lihat cek 3 di
        #     bawah ("Teks Abstract ... seharusnya dicetak italic secara
        #     keseluruhan"). Kalau tidak dikecualikan di sini, cek
        #     per-kata ini ikut jalan di paragraf yang sama dan hampir
        #     SEMUA katanya (karena memang bahasa Inggris) ke-flag satu
        #     per satu sebagai "istilah asing" terpisah -- ini bukan
        #     recall tambahan yang berguna, cuma duplikat/derivatif dari
        #     finding "seluruh paragraf belum italic" yang sudah ada,
        #     dan membanjiri description dengan puluhan kata yang
        #     membingungkan (mis. "previous", "generations", "and", dst
        #     ikut ke-list padahal levelnya seharusnya paragraf, bukan
        #     kata). Abstrak (ID) TIDAK ikut dikecualikan -- istilah
        #     asing yang menyelip di teks Indonesia tetap relevan dicek
        #     per-kata seperti body text biasa.
        # DINONAKTIFKAN (atas permintaan user, 2026-08-21): cek "istilah
        # asing wajib italic" di-nonaktifkan sementara -- seluruh blok
        # if di bawah (dan check_paragraph() dari foreign_term_checker,
        # lihat import yang dikomentari di atas modul ini) TIDAK
        # dieksekusi lagi, jadi checks[chapter]["checked"]/["failed"]
        # untuk rule ini juga tidak lagi bertambah. Aktifkan kembali
        # dengan un-comment blok ini + import check_paragraph di atas
        # kalau rule ini mau dipakai lagi.
        # if (
        #     ptype == "body"
        #     and chapter != "DAFTAR PUSTAKA"
        #     and elem["abstract_part"] not in ("abstract_body", "keyword_en")
        # ):
        #     found_any_candidate, bad_words = check_paragraph(elem["runs"])
        #
        #     if found_any_candidate:
        #         checks[chapter]["checked"] += 1
        #         if bad_words:
        #             checks[chapter]["failed"] += 1
        #             joined = ", ".join(f"'{w}'" for w in bad_words)
        #             findings.append(
        #                 {
        #                     "category": "Format",
        #                     "chapter": chapter,
        #                     "location": loc,
        #                     "description": (
        #                         f"Istilah asing {joined} terdeteksi tidak "
        #                         f"dicetak italic."
        #                     ),
        #                 }
        #             )

    return findings, checks


def compare_sections(sections: list[dict], mirror_margins: bool, source: str = "docx"):
    """
    Margin di Word berlaku sebagai satu pengaturan untuk seluruh dokumen
    (Page Setup), bukan properti yang wajar berbeda-beda per section --
    section break biasanya dipakai mahasiswa untuk kebutuhan lain (mis.
    halaman kosong sisipan, ganti orientasi tabel lebar), bukan untuk
    sengaja mengubah margin. Kalau tiap section dihitung sebagai unit
    cek terpisah (versi lama), mahasiswa yang lebih banyak insert section
    break otomatis dapat "checked" lebih besar untuk kesalahan MARGIN
    YANG SAMA PERSIS dibanding mahasiswa yang jarang insert section break
    -- padahal substansi kesalahannya identik, dan bisa membuat skor
    antar-mahasiswa timpang cuma gara-gara kebiasaan pakai section break,
    bukan gara-gara kualitas dokumennya. Supaya jumlah section break
    tidak mempengaruhi bobot penalti, tiap SISI margin (kiri/atas/bawah/
    kanan) sekarang cuma 1 unit "checked" untuk SELURUH dokumen, terlepas
    dari ada berapa section. Kalau ternyata ada section yang nilainya
    beda sendiri (kasus sah: margin memang sengaja dibedakan per section,
    mis. halaman landscape), itu tetap dilaporkan spesifik di description
    supaya tidak hilang informasinya buat ditelusuri.

    UNIT PERBANDINGAN: kiri/atas/bawah/kanan & ukuran kertas dibandingkan
    dalam TWIPS (RULES_TWIPS), bukan cm. Twips itu unit asli yang
    tersimpan di w:pgMar/w:pgSz pada XML docx, jadi integer EXACT --
    membandingkan langsung di situ menghindari noise floating-point dari
    konversi EMU->cm (mis. section.left_margin.cm bisa balik 3.9999998
    padahal aslinya persis 4.0 cm di XML). *_cm dari extractor.py TETAP
    dipakai, tapi HANYA untuk format pesan error yang dibaca mahasiswa/
    dosen (lihat sections[i]["left_cm"] dst), tidak untuk logika
    abs(a - b) > tolerance lagi.

    TOLERANSI KHUSUS JALUR PDF (ditambahkan 2026-08-24)
    -----------------------------------------------------
    sections dari extract_all(doc) (docx) berisi angka margin EXACT,
    dibaca langsung dari w:pgMar -- tol_twips dasar dari RULES_TWIPS
    sudah cukup ketat untuk itu, memang seharusnya begitu.

    sections dari extract_all_pdf() (PDF) SEBALIKNYA berisi ESTIMASI
    geometris dari posisi bounding-box teks (lihat docstring
    extract_section_params_pdf), bukan properti yang dibaca literal --
    jadi wajar ada noise pengukuran, TIDAK sama presisinya dengan docx.
    Noise ini juga TIDAK SIMETRIS antar sisi margin:
    - kiri/kanan diambil dari baris BODY MULTI-BARIS SAJA (sudah
      difilter dari heading/caption/center) -- estimasinya relatif
      stabil, jadi toleransi ekstra kecil saja cukup.
    - atas/bawah diambil dari blok PALING ATAS/BAWAH di tiap halaman
      TANPA filter apapun (lihat extract_section_params_pdf) -- kalau
      blok itu kebetulan judul bab yang punya "space before" dari
      style paragraf, jaraknya ikut kehitung sebagai bagian margin
      padahal itu spacing paragraf, bukan margin halaman murni. Sisi
      ini butuh toleransi ekstra yang lebih longgar.

    source="pdf" menambahkan allowance ekstra ini DI ATAS tol_twips
    dasar (bukan menggantikannya) -- supaya deviasi kecil yang mungkin
    cuma bias pengukuran tidak ikut menjatuhkan skor, sementara deviasi
    besar (mis. margin kiri meleset ~1cm) tetap kefilter dan dilaporkan
    sebagai temuan asli. Jalur docx (source="docx", default) TIDAK
    terpengaruh sama sekali -- allowance ekstra ini 0 untuk docx.
    """
    findings = []
    checks = defaultdict(_new_counter)
    tol_twips = RULES_TWIPS["margin_tolerance"]

    # Allowance ekstra HANYA untuk jalur PDF, per sisi margin -- lihat
    # penjelasan lengkap di docstring fungsi ini. Nilainya dalam twips
    # (1cm = 567 twips): kiri/kanan +0.15cm (~85 twips), atas/bawah
    # +0.35cm (~198 twips) karena lebih rawan bias "space before" judul.
    # Titik awal yang wajar, bukan angka final -- boleh dikalibrasi
    # ulang kalau ternyata masih kebanyakan/kekurangan false positive
    # setelah dicoba ke lebih banyak dokumen sungguhan.
    PDF_EXTRA_TOLERANCE_TWIPS = {
        "kiri": 85,
        "kanan": 85,
        "atas": 198,
        "bawah": 198,
    }
    group = "Umum (Margin Halaman)"
    expected_mirror = RULES["mirror_margins"]

    # atas/bawah: TIDAK pernah berselang-seling ganjil/genap, mirror-
    # margin cuma soal kiri/kanan -- jadi selalu dicek sebagai satu nilai
    # gabungan utk seluruh dokumen, sama seperti sebelumnya, terlepas
    # dari source atau status mirror-margin.
    margin_defs = [
        (
            "atas",
            "top_cm",
            "top_twips",
            RULES_TWIPS["margin_other"],
            RULES["margin_other_cm"],
        ),
        (
            "bawah",
            "bottom_cm",
            "bottom_twips",
            RULES_TWIPS["margin_other"],
            RULES["margin_other_cm"],
        ),
    ]

    # atas/bawah dulu (selalu 1 nilai gabungan, tidak ada parity) --
    # kiri/kanan ditangani terpisah di blok `pairs` di bawah, karena
    # kiri/kanan butuh logika parity-aware utk pdf mirror-margin
    # (lihat komentar panjang di blok itu).
    for margin_def in margin_defs:
        # atas/bawah -- tidak ada allowance PDF per-sisi terpisah dgn
        # nama beda dari label tampilan, tol_label = label asli.
        label, cm_key, twips_key, expected_twips, expected_cm = margin_def
        tol_label = label

        checks[group]["checked"] += 1  # 1x per sisi margin, untuk seluruh dokumen

        effective_tol = tol_twips
        if source == "pdf":
            effective_tol += PDF_EXTRA_TOLERANCE_TWIPS[tol_label]

        offending = [
            (sec["index"], sec[cm_key])  # cm_key cuma buat ditampilkan
            for sec in sections
            if abs(sec[twips_key] - expected_twips) > effective_tol
        ]
        if offending:
            checks[group]["failed"] += 1
            detail = ", ".join(f"#{idx} ({val:.2f}cm)" for idx, val in offending)
            plural = "section" if len(offending) == 1 else "beberapa section"
            findings.append(
                {
                    "category": "Format",
                    "chapter": group,
                    "location": f"Section {detail}",
                    "description": (
                        f"Margin {label} tidak sesuai di {plural} tsb, "
                        f"seharusnya {expected_cm}cm."
                    ),
                }
            )

    # kiri/kanan: DIPISAH secara PENGUKURAN kalau source pdf & dokumen
    # mirror-margin (kiri/kanan SEHARUSNYA beda antara halaman ganjil &
    # genap -- lihat penjelasan panjang di atas), TAPI tetap DIHITUNG
    # sebagai SATU unit "checked" per sisi (bukan dua/empat), SAMA
    # PERSIS seperti jalur docx.
    #
    # BUG DIPERBAIKI (2026-09-07, ditemukan dari perbandingan skor tiga
    # format pada dokumen yang sama): sebelum ini, jalur pdf/rtf mirror-
    # margin menghasilkan margin_defs 2x lebih banyak dari jalur docx
    # (6 unit checked -- atas, bawah, kiri-ganjil, kanan-ganjil, kiri-
    # genap, kanan-genap -- dibanding 4 unit docx: atas, bawah, kiri,
    # kanan). Karena skor per-kategori dihitung sebagai rasio
    # (checked-failed)/checked, dua dokumen dengan KESALAHAN MARGIN YANG
    # SUBSTANSINYA SAMA bisa dapat skor "Umum (Margin Halaman)" yang
    # jauh berbeda HANYA gara-gara source-nya beda (pdf vs docx), bukan
    # gara-gara kualitas dokumennya -- ini salah satu penyebab utama
    # skor margin docx 0%, pdf 16.67%, rtf 50% padahal ketiganya dari
    # dokumen fisik yang sama. Fix: kiri (ganjil+genap digabung jadi
    # SATU keputusan pass/fail) dan kanan (idem) masing-masing tetap
    # cuma 1 unit checked, persis seperti docx -- detail ganjil/genap
    # tetap dilaporkan di description kalau gagal, informasinya tidak
    # hilang, cuma tidak lagi dobel-hitung ke skor.
    if source == "pdf" and expected_mirror:
        pairs = [
            (
                "kiri",
                [
                    ("halaman ganjil", "left_odd_cm", "left_odd_twips", "margin_left"),
                    (
                        "halaman genap",
                        "left_even_cm",
                        "left_even_twips",
                        "margin_other",
                    ),
                ],
            ),
            (
                "kanan",
                [
                    (
                        "halaman ganjil",
                        "right_odd_cm",
                        "right_odd_twips",
                        "margin_other",
                    ),
                    (
                        "halaman genap",
                        "right_even_cm",
                        "right_even_twips",
                        "margin_left",
                    ),
                ],
            ),
        ]
    else:
        pairs = [
            ("kiri", [(None, "left_cm", "left_twips", "margin_left")]),
            ("kanan", [(None, "right_cm", "right_twips", "margin_other")]),
        ]

    for tol_label, sub_defs in pairs:
        checks[group]["checked"] += 1  # 1x per sisi (kiri/kanan), untuk seluruh dokumen
        effective_tol = tol_twips
        if source == "pdf":
            effective_tol += PDF_EXTRA_TOLERANCE_TWIPS[tol_label]

        offending_parts = []
        any_offending = False
        for parity_label, cm_key, twips_key, rule_key in sub_defs:
            expected_twips = RULES_TWIPS[rule_key]
            expected_cm = RULES[
                {"margin_left": "margin_left_cm", "margin_other": "margin_other_cm"}[
                    rule_key
                ]
            ]
            bad = [
                (sec["index"], sec[cm_key])
                for sec in sections
                if abs(sec[twips_key] - expected_twips) > effective_tol
            ]
            if bad:
                any_offending = True
                loc_suffix = f" ({parity_label})" if parity_label else ""
                for idx, val in bad:
                    offending_parts.append(
                        f"#{idx}{loc_suffix} ({val:.2f}cm, harus {expected_cm}cm)"
                    )

        if any_offending:
            checks[group]["failed"] += 1
            detail = ", ".join(offending_parts)
            plural = "section" if len(offending_parts) == 1 else "beberapa section"
            findings.append(
                {
                    "category": "Format",
                    "chapter": group,
                    "location": f"Section {detail}",
                    "description": f"Margin {tol_label} tidak sesuai di {plural} tsb.",
                }
            )

    # --- ukuran kertas: A4, orientasi-agnostik (potrait/landscape) ---
    # Word menyimpan width/height literal sesuai orientasi saat itu, jadi
    # landscape A4 punya width & height yang KETUKER dibanding potrait
    # (29.7 x 21, bukan 21 x 29.7). Supaya section landscape yang sengaja
    # dipakai utk tabel lebar tidak salah dianggap "bukan A4", short/long
    # side-nya dibandingkan (bukan width/height mentah), jadi orientasi
    # apa pun tetap lolos selama kertasnya memang A4.
    # Sama seperti margin: 1 unit "checked" untuk SELURUH dokumen, tidak
    # per section, supaya jumlah section break tidak mempengaruhi bobot.
    size_group = "Umum (Ukuran Kertas)"
    tol_size_twips = RULES_TWIPS["page_size_tolerance"]
    expected_short = min(RULES_TWIPS["page_width"], RULES_TWIPS["page_height"])
    expected_long = max(RULES_TWIPS["page_width"], RULES_TWIPS["page_height"])

    checks[size_group]["checked"] += 1

    offending = []
    for sec in sections:
        short = min(sec["width_twips"], sec["height_twips"])
        long = max(sec["width_twips"], sec["height_twips"])
        if (
            abs(short - expected_short) > tol_size_twips
            or abs(long - expected_long) > tol_size_twips
        ):
            # width_cm/height_cm (bukan twips) yang disimpan ke offending
            # -- cuma buat format pesan error di bawah, bukan logika cek.
            offending.append((sec["index"], sec["width_cm"], sec["height_cm"]))

    if offending:
        checks[size_group]["failed"] += 1
        detail = ", ".join(f"#{idx} ({w:.2f}cm x {h:.2f}cm)" for idx, w, h in offending)
        plural = "section" if len(offending) == 1 else "beberapa section"
        findings.append(
            {
                "category": "Format",
                "chapter": size_group,
                "location": f"Section {detail}",
                "description": (
                    f"Ukuran kertas di {plural} tsb tidak sesuai A4 "
                    f"({RULES['page_width_cm']}cm x {RULES['page_height_cm']}cm, "
                    f"potrait maupun landscape)."
                ),
            }
        )

    # --- mirror-margin: wajib ON, level DOKUMEN (bukan per section) ---
    # KOREKSI (sebelumnya salah): w:mirrorMargins itu di settings.xml,
    # SATU untuk seluruh dokumen -- beda dari margin/ukuran kertas yang
    # memang atribut tiap section. Jadi TIDAK ada "section yang
    # menyimpang" seperti dua cek di atas -- cuma 1 flag boolean untuk
    # seluruh dokumen, checked=1/failed=0 kalau sesuai, 1/1 kalau tidak.
    mirror_group = "Umum (Cetak Bolak-Balik)"
    # expected_mirror sudah dihitung di atas (dipakai jg utk milih jalur
    # kiri/kanan parity-aware vs gabungan) -- tidak perlu dihitung ulang.

    checks[mirror_group]["checked"] += 1
    if mirror_margins != expected_mirror:
        checks[mirror_group]["failed"] += 1
        findings.append(
            {
                "category": "Format",
                "chapter": mirror_group,
                "location": "Pengaturan Dokumen (Page Setup)",
                "description": (
                    "Mirror-margin tidak aktif di pengaturan dokumen, padahal "
                    "dokumen dicetak bolak-balik seharusnya mengaktifkan "
                    "mirror-margin (Layout > Margins > Mirror Margins di Word)."
                ),
            }
        )

    return findings, checks


# Toleransi spasi opsional setelah titik, sinkron dengan CAPTION_RE di
# extractor.py -- kalau dua regex ini tidak sinkron, caption yang lolos
# klasifikasi "caption" di extractor tapi gagal di-parse di sini akan
# silently di-skip (lihat guard "if not m: continue" di bawah).
#
# BUG DIPERBAIKI (2026-09-12, laporan user): caption di dalam LAMPIRAN
# pakai HURUF ("Tabel A.1", "Gambar B.1") sebagai pengganti angka bab
# -- grup ke-2 di sini sekarang menerima huruf JUGA (dulu cuma \d+),
# sinkron dengan perluasan CAPTION_RE di extractor.py. BAB_NUMBER_RE
# ditambah padanannya untuk "LAMPIRAN <huruf>" (LAMPIRAN_NUMBER_RE),
# dipakai di compare_caption_numbering() supaya cek 1 ("angka/huruf
# bab di nomor caption wajib cocok chapter tempat caption itu ada")
# jalan juga untuk lampiran, bukan cuma BAB.
CAPTION_PARSE_RE = re.compile(r"^(Gambar|Tabel)\s+([A-Za-z0-9]+)\.\s*(\d+)")
BAB_NUMBER_RE = re.compile(r"^BAB\s+(\d+)$")
LAMPIRAN_NUMBER_RE = re.compile(r"^LAMPIRAN\s+([A-Z]+)$")

# Buang prefix label "Kata Kunci"/"Keywords" (beserta separator
# opsional : / . / -) dari baris keyword, sisanya baru dipecah per
# frasa berdasarkan koma di compare_abstract().
KEYWORD_STRIP_RE = re.compile(r"^(Kata\s+Kunci|Keywords?)\s*[:.\-]?\s*", re.IGNORECASE)


def compare_caption_numbering(paragraphs: list[dict]):
    """
    Cek penomoran Gambar/Tabel PER BAB. CAPTION_RE di extractor.py cuma
    memvalidasi BENTUK angka ("Gambar/Tabel \\d+\\.\\d+"), bukan
    KEBENARAN angkanya -- fungsi ini yang mengisi bagian itu, jalan
    setelah tahu current_chapter tiap paragraf. Dua aturan:

      1. Angka bab di depan nomor caption (mis. "2" di "Gambar 2.3")
         wajib cocok dengan bab tempat caption itu ada.
      2. Nomor urut (mis. "3" di "Gambar 2.3") wajib berurutan tanpa
         lompat/duplikat -- counter dihitung terpisah per (bab, jenis),
         karena Gambar dan Tabel masing-masing punya urutan sendiri dan
         reset ke 1 tiap ganti bab.

    Caption di dalam LAMPIRAN (mana pun) atau BIODATA PENULIS dilewati
    total, termasuk cek 1 di atas -- lihat _is_excluded_chapter().
    """
    findings = []
    checks = defaultdict(_new_counter)
    next_seq = {}  # {(chapter, jenis): nomor urut yang diharapkan berikutnya}

    for elem in paragraphs:
        if elem["type"] != "caption":
            continue

        chapter = elem["chapter"]
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        text = elem["full_text"]
        loc = f"Paragraf #{elem['index']} ('{text_preview(text)}')"

        m = CAPTION_PARSE_RE.match(text)
        if not m:
            # Harusnya tidak terjadi -- ptype "caption" sudah lolos
            # CAPTION_RE (bentuk regex sama) di extractor.py. Dijaga di
            # sini supaya tidak crash kalau kedua regex ini di masa
            # depan tidak sinkron lagi.
            continue

        jenis, cap_bab_str, cap_seq_str = m.groups()
        cap_seq = int(cap_seq_str)
        bab_match = BAB_NUMBER_RE.match(chapter)
        lampiran_match = LAMPIRAN_NUMBER_RE.match(chapter)

        # --- cek 1: angka/huruf bab di nomor caption harus cocok chapter
        # aktual. Dua jenis chapter dibandingkan beda cara: "BAB N" ->
        # cap_bab_str harus ANGKA yang sama dengan N (bandingkan int,
        # tahan variasi "04" vs "4"); "LAMPIRAN X" -> cap_bab_str harus
        # HURUF yang sama dengan X (bandingkan string, case-insensitive
        # -- CAPTION_RE menerima huruf apa adanya, mahasiswa bisa saja
        # ketik "tabel a.1" huruf kecil meski jarang).
        checks[chapter]["checked"] += 1
        cocok = False
        if bab_match is not None and cap_bab_str.isdigit():
            cocok = int(cap_bab_str) == int(bab_match.group(1))
        elif lampiran_match is not None and cap_bab_str.isalpha():
            cocok = cap_bab_str.upper() == lampiran_match.group(1)

        if not cocok:
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"Nomor {jenis.lower()} '{cap_bab_str}.{cap_seq_str}' "
                        f"tidak cocok dengan {chapter} tempat paragraf ini "
                        f"berada."
                    ),
                }
            )

        # --- cek 2: nomor urut harus berurutan, tidak lompat/duplikat ---
        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): dikomentari,
        # bukan dihapus. `next_seq[key]` tetap di-update apa adanya
        # (ikut nomor yang DITULIS mahasiswa, bukan yang "seharusnya")
        # supaya kalau nanti diaktifkan lagi urutan hitungnya tetap benar.
        key = (chapter, jenis)
        # expected_seq = next_seq.get(key, 1)
        #
        # checks[chapter]["checked"] += 1
        # if cap_seq != expected_seq:
        #     checks[chapter]["failed"] += 1
        #     findings.append(
        #         {
        #             "category": "Format",
        #             "chapter": chapter,
        #             "location": loc,
        #             "description": (
        #                 f"Nomor urut {jenis.lower()} '{cap_bab_str}.{cap_seq_str}' "
        #                 f"tidak berurutan, seharusnya '{cap_bab_str}.{expected_seq}' "
        #                 f"({jenis.lower()} ke-{expected_seq} di {chapter})."
        #             ),
        #         }
        #     )
        #     # Lanjutkan hitungan dari angka yang SEHARUSNYA (bukan ikut
        #     # angka yang salah ketik mahasiswa) -- supaya satu typo tidak
        #     # bikin semua caption berikutnya di bab yang sama ikut
        #     # dianggap salah berantai (cascading false positives).
        #     next_seq[key] = expected_seq + 1
        # else:
        #     next_seq[key] = cap_seq + 1
        next_seq[key] = cap_seq + 1

    return findings, dict(checks)


def _nearby_paragraphs(
    paragraphs_by_index: dict, idx, direction: str, max_hops: int = 3
):
    """
    Sampai `max_hops` paragraf TERDEKAT dari index fisik `idx` ke arah
    `direction` ("prev": index <= idx, menjauh; "next": index >= idx,
    menjauh). Dipakai khusus buat cek caption tabel -- caption yang
    "menempel" tidak selalu persis paragraf pertama yang ditemui:
    kadang ada 1+ baris kosong (di-skip total dari paragraphs_by_index)
    ATAU baris "Sumber: ..." yang justru duduk PERSIS di antara tabel
    dan caption-nya. Kalau cuma dicek satu paragraf terdekat saja,
    caption yang sebenarnya ada -- cuma satu-dua paragraf lebih jauh --
    keliru dianggap tidak ada.
    """
    if idx is None or not paragraphs_by_index:
        return []
    sorted_indices = sorted(paragraphs_by_index)
    if direction == "prev":
        start = bisect.bisect_right(sorted_indices, idx) - 1
        step = -1
    else:
        start = bisect.bisect_left(sorted_indices, idx)
        step = 1
    result = []
    pos = start
    while 0 <= pos < len(sorted_indices) and len(result) < max_hops:
        result.append(paragraphs_by_index[sorted_indices[pos]])
        pos += step
    return result


def _nearest_paragraph(paragraphs_by_index: dict, idx, direction: str):
    """
    Paragraf non-kosong TERDEKAT dari index fisik `idx` (bukan cuma
    yang persis di index itu) -- paragraf kosong sengaja di-skip total
    di extractor.py, jadi index yang persis bersebelahan dengan
    gambar/tabel bisa jadi kosong walau ada paragraf isi (mis. caption)
    tidak jauh dari situ. direction: "prev" cari index <= idx, "next"
    cari index >= idx.
    """
    nearby = _nearby_paragraphs(paragraphs_by_index, idx, direction, max_hops=1)
    return nearby[0] if nearby else None


def _chapter_of_table(paragraphs_by_index: dict, tbl: dict) -> str:
    """Bab tempat sebuah tabel berada, ditebak dari paragraf tetangganya
    (prev diutamakan, next kalau prev tidak ada)."""
    prev_elem = _nearest_paragraph(paragraphs_by_index, tbl["prev_para_idx"], "prev")
    if prev_elem is not None:
        return prev_elem["chapter"]
    next_elem = _nearest_paragraph(paragraphs_by_index, tbl["next_para_idx"], "next")
    if next_elem is not None:
        return next_elem["chapter"]
    return "Umum (Tabel)"


def compare_visual_objects(
    paragraphs: list[dict], tables: list[dict], source: str = "docx"
):
    """
    source: "docx" | "pdf" -- dipakai untuk toleransi estimasi PDF di
    beberapa cek visual di modul ini (lihat catatan tiap blok).

    Cek aturan tata letak Gambar, Tabel, dan Persamaan:

      - Gambar: rata tengah, DAN wajib punya caption ("Gambar N.N ...")
        menempel di salah satu sisinya (atas/bawah) -- simetris dengan
        aturan Tabel di bawah.
      - Tabel: caption ("Tabel N.N ...") menempel di salah satu sisinya
        (atas/bawah). Rata tengahnya sendiri berlaku untuk CAPTION-nya
        (ptype "caption", dicek di compare_paragraphs), BUKAN untuk
        tabelnya -- tabel Word lazim diregangkan penuh ke lebar margin,
        jadi tidak relevan dipaksa "rata tengah" seperti objek gambar.
        DITAMBAH: baris header (baris pertama) tabel dicek rata tengah,
        tebal, dan berlatar-belakang gelap -- sesuai anjuran panduan
        ("Judul pada tabel DAPAT dituliskan rata tengah, tebal dan
        berlatar-belakang agak gelap"). Dilewati (None) untuk tabel
        non-data (n_rows<=1, mis. trik layout nomor persamaan).
      - Persamaan: rata kiri (mentok, TIDAK dipaksa indentasi 10mm
        persis -- atas permintaan user, 2026-09-07, cuma "rumusnya rata
        kiri" yang wajib), 1 baris kosong di atas & bawah (jalur docx
        dibaca langsung dari struktur dokumen, jalur PDF diESTIMASI dari
        jarak vertikal antar elemen -- lihat
        _infer_equation_blank_lines() di pdf_extractor.py), dan wajib
        punya nomor identitas mis. "(2.1)" ATAU "(Persamaan 2.1)" yang
        dicetak tebal dan rata kanan (dideteksi lewat
        EQUATION_NUMBER_IN_TEXT_RE -- lihat docstring regex tsb di
        extractor.py untuk cara deteksinya).
      - Sumber kutipan ("Sumber: ..."): OPSIONAL (panduan bilang
        "penulis DAPAT menyebutkan"), jadi cuma diperiksa formatnya
        (rata tengah, 11pt) KALAU baris itu memang ada di dokumen --
        tidak pernah dianggap wajib ada.

    SENGAJA TIDAK dicek (supaya tidak overly strict / banyak false
    positive):
      - Bold untuk caption gambar/tabel -- teks panduan MEMANG bilang
        "...dituliskan rata tengah dan tebal" (baik untuk gambar
        maupun tabel), TAPI contoh visual di panduan (Gambar 3.1,
        Gambar 3.2, Tabel 3.1, dst) captionnya TIDAK bold. Karena teks
        & contoh saling bertentangan, cek bold sengaja dilonggarkan
        (tidak dipaksakan) -- rata tengahnya sendiri tetap dicek
        karena teks & contoh konsisten soal itu.
      - 1 baris kosong di atas/bawah GAMBAR & TABEL -- DIHAPUS dari cek
        (beda dari persamaan, yang tetap dicek KHUSUS jalur docx).
        Alasan: mahasiswa yang pakai "Spacing Before/After Paragraph"
        (bukan tombol Enter/baris kosong beneran) untuk memberi jarak
        visual di sekitar gambar/tabel akan salah ditandai kalau ini
        dipaksakan, padahal secara visual hasilnya sama-sama benar.
      - Indentasi persis 10mm untuk persamaan -- DILONGGARKAN (atas
        permintaan user, 2026-09-07) jadi cuma "rata kiri", tidak
        dipaksa jarak indentasi yang eksak.

    KETERBATASAN YANG DISADARI (bukan dicek di sini, butuh verifikasi
    manual):
      - Posisi caption Gambar HARUS di bawah gambar per contoh visual
        panduan, tapi redaksi panduan untuk Tabel *literally* juga
        bilang "dituliskan pada 1 baris DIBAWAH tabel" -- padahal
        contoh visualnya (Tabel 3.1) caption-nya di ATAS tabel.
        Kemungkinan salah ketik redaksi (copy-paste dari bagian
        Gambar). Karena ambigu, cek posisi caption Tabel di sini
        SENGAJA menerima caption di atas MAUPUN di bawah tabel.
      - Tabel tidak boleh lebih dari 1 halaman: butuh info halaman
        fisik hasil render PDF per baris tabel, di luar jangkauan
        modul ini (python-docx tidak tahu pemotongan halaman).

    Gambar/Persamaan/Tabel/Sumber kutipan yang berada di LAMPIRAN (mana
    pun) atau BIODATA PENULIS dilewati total dari semua cek di atas --
    lihat _is_excluded_chapter().
    """
    findings = []
    checks = defaultdict(_new_counter)
    paragraphs_by_index = {e["index"]: e for e in paragraphs}

    # --- GAMBAR ---
    for elem in paragraphs:
        if elem["type"] != "image":
            continue
        chapter = elem["chapter"]
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        loc = f"Paragraf #{elem['index']} (gambar)"

        # --- DINONAKTIFKAN (atas permintaan user, 2026-09-09): cek
        #     "Gambar seharusnya diletakkan rata tengah." -- dihapus
        #     dari pemeriksaan, berlaku untuk jalur docx maupun pdf
        #     karena fungsi ini dipakai bersama oleh keduanya.
        # checks[chapter]["checked"] += 1
        # if elem["alignment"] != "center":
        #     checks[chapter]["failed"] += 1
        #     findings.append(
        #         {
        #             "category": "Format",
        #             "chapter": chapter,
        #             "location": loc,
        #             "description": "Gambar seharusnya diletakkan rata tengah.",
        #         }
        #     )

        # --- DINONAKTIFKAN (atas permintaan user, 2026-09-09): cek
        #     caption ("Gambar N.N ...") wajib ada, menempel di salah
        #     satu sisi (atas/bawah) -- dihapus dari pemeriksaan,
        #     berlaku untuk jalur docx maupun pdf.
        # nearby_elems = _nearby_paragraphs(
        #     paragraphs_by_index, elem["index"] - 1, "prev"
        # ) + _nearby_paragraphs(paragraphs_by_index, elem["index"] + 1, "next")
        # has_caption_adjacent = any(
        #     e["type"] == "caption" and e["full_text"].startswith("Gambar")
        #     for e in nearby_elems
        # )
        # checks[chapter]["checked"] += 1
        # if not has_caption_adjacent:
        #     checks[chapter]["failed"] += 1
        #     findings.append(
        #         {
        #             "category": "Format",
        #             "chapter": chapter,
        #             "location": loc,
        #             "description": (
        #                 "Gambar tidak punya nomor identitas & keterangan "
        #                 "('Gambar N.N ...') tepat di atas/bawahnya."
        #             ),
        #         }
        #     )

    # --- PERSAMAAN ---
    # DIAKTIFKAN KEMBALI (atas permintaan user, 2026-09-07) dengan 2
    # penyesuaian dari versi lama yang sempat dinonaktifkan 2026-08-23:
    #   1. Cek indentasi 10mm persis DIHAPUS -- diganti cek "rata kiri"
    #      (elem["alignment"] == "left") saja, sesuai instruksi user:
    #      "ga harus 10mm yg penting rumusnya rata kiri mentok".
    #   2. Cek nomor identitas sekarang JUGA menerima "(Persamaan 2.1)"
    #      selain "(2.1)" -- lihat EQUATION_NUMBER_IN_TEXT_RE di
    #      extractor.py.
    # Cek "1 baris kosong di atas/bawah" AKTIF untuk docx & PDF (nilai
    # PDF-nya hasil estimasi dari pdf_extractor.py, lihat docstring
    # _infer_equation_blank_lines() untuk keterbatasannya).
    for elem in paragraphs:
        if elem["type"] != "equation":
            continue
        chapter = elem["chapter"]
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        loc = f"Paragraf #{elem['index']} (persamaan)"

        # --- rata kiri (RUMUSNYA, bukan cuma paragraf pembungkusnya) ---
        # BUG DIPERBAIKI (2026-09-08, laporan mahasiswa -- "rumus rata
        # tengah tapi tidak pernah ke-flag rata kiri"): Word menyimpan
        # justifikasi VISUAL sebuah Equation di m:oMathParaPr/m:jc,
        # TERPISAH dari w:pPr/w:jc paragraf pembungkusnya (elem
        # ["alignment"]) -- lihat docstring _get_equation_alignment()
        # di extractor.py. Equation yang di-set "Display" oleh Word
        # defaultnya "centerGroup" (rata tengah) di property INI,
        # walau w:pPr/w:jc paragrafnya sendiri kebetulan "left" (atau
        # malah None/warisan). Cek versi lama HANYA baca elem
        # ["alignment"], jadi kasus ini lolos tanpa ke-flag padahal
        # rumusnya SECARA VISUAL rata tengah, bukan rata kiri.
        # Prioritas sumber: equation_alignment (kalau ada) > alignment
        # paragraf biasa (fallback, utk equation inline kecil yang
        # tidak pernah di-"Display"-kan sehingga tidak punya
        # m:oMathParaPr/m:jc sama sekali).
        checks[chapter]["checked"] += 1
        effective_alignment = elem.get("equation_alignment") or elem["alignment"]
        if effective_alignment != "left":
            checks[chapter]["failed"] += 1
            detected = effective_alignment or "tidak diset"
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"Rumus (persamaan) seharusnya rata kiri (mentok "
                        f"kiri), terdeteksi: {detected}. Cek justifikasi "
                        f"objek Equation-nya sendiri (klik panah kecil di "
                        f"pojok kanan-bawah kotak rumus -> Justification -> "
                        f"Left), BUKAN cuma perataan paragraf -- Word "
                        f"menyimpan keduanya secara terpisah."
                    ),
                }
            )

        # --- 1 baris kosong di atas & bawah ---
        # DIAKTIFKAN UNTUK PDF JUGA (2026-09-20, laporan user --
        # ditemukan dokumen mahasiswa nyata yang persamaannya TIDAK
        # diberi baris kosong, padahal cek ini sebelumnya SELALU
        # dilewati untuk source=="pdf" karena blank_before/blank_after
        # dari pdf_extractor.py dulu cuma placeholder False, tidak
        # pernah benar-benar dihitung). pdf_extractor.py sekarang
        # mengisi kedua field ini via _infer_equation_blank_lines()
        # (estimasi dari jarak vertikal antar elemen, BUKAN pembacaan
        # struktur dokumen langsung seperti docx) -- lihat docstring
        # fungsi itu untuk detail & keterbatasan estimasinya. Karena
        # ini estimasi (bukan docx yang baca w:p kosong langsung),
        # kalau nanti terbukti sering false-positive di jalur PDF,
        # opsi mitigasinya sama seperti cek PDF lain yang lebih dulu
        # dimatikan khusus source=="pdf" -- BUKAN menghapus baris
        # kode di bawah ini, cukup kembalikan guard "if source ==
        # 'docx'" yang sempat ada di sini sebelum perubahan ini.
        checks[chapter]["checked"] += 1
        if not (elem["blank_before"] and elem["blank_after"]):
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        "Persamaan seharusnya menyisakan 1 baris kosong di "
                        "atas dan di bawahnya."
                    ),
                }
            )

        # --- nomor identitas, rata kanan (terdeteksi via posisi di
        #     akhir baris, lihat docstring EQUATION_NUMBER_IN_TEXT_RE),
        #     dan bold ---
        checks[chapter]["checked"] += 1
        m = EQUATION_NUMBER_IN_TEXT_RE.search(elem["full_text"])
        # FALLBACK (2026-09-08, laporan mahasiswa -- "nomor identitas
        # sudah ada tapi dibilang tidak ada"): kalau tidak ketemu di
        # full_text (w:t biasa), coba cari lagi di equation_math_text
        # (<m:t>, teks yang diketik masih di dalam math zone -- lihat
        # docstring _equation_math_text() di extractor.py). Ini KHUSUS
        # mengonfirmasi ADA/TIDAKNYA nomor -- cek bold di bawah SENGAJA
        # tidak dijalankan untuk kasus fallback ini karena offset
        # karakter <m:t> tidak comparable ke elem["runs"] (yang cuma
        # berisi <w:r> biasa), jadi tidak bisa dipetakan run mana yang
        # perlu dicek bold-nya.
        found_in_math_zone = False
        if not m:
            # Urutan fallback (2026-09-21, laporan user -- nomor "(3.1)",
            # "(3.2)" sudah ada tapi tetap dibilang tidak ada):
            #   1. equation_full_text -- teks LENGKAP urutan dokumen
            #      termasuk kurung objek <m:d> ("#(3.1)" otomatis Word
            #      cuma menyimpan "3.1" di <m:t>, kurungnya properti
            #      <m:d>) dan hasil field SEQ (<w:fldSimple>) yang tidak
            #      terbaca para.text. Lihat _equation_full_text().
            #   2. equation_math_text -- perilaku lama, dipertahankan.
            #   3. full_text yang dinormalkan -- jalur PDF: teks hasil
            #      ekstraksi rumus sering membawa spasi tipis/NBSP/
            #      zero-width di sekitar nomor.
            # Semua tetap wajib nomor jadi karakter TERAKHIR (aturan
            # posisi tidak dilonggarkan) -- yang dilonggarkan cuma
            # cara membaca teksnya.
            for key in ("equation_full_text", "equation_math_text", "full_text"):
                candidate = normalize_equation_text(elem.get(key) or "")
                if candidate and EQUATION_NUMBER_IN_TEXT_RE.search(candidate):
                    found_in_math_zone = True
                    break

        if not m and not found_in_math_zone:
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        "Persamaan tidak memiliki nomor identitas rata kanan, "
                        "mis. '(2.1)' atau '(Persamaan 2.1)'."
                    ),
                }
            )
        elif m:
            checks[chapter]["checked"] += 1
            # Cek bold HANYA pada run yang jadi bagian span nomor
            # identitas (hasil match "m" di atas), BUKAN seluruh run
            # paragraf.
            #
            # BUG DIPERBAIKI (2026-09-07): versi lama ("runs_with_text
            # = [r for r in elem['runs'] if r['text'].strip()]") malah
            # mengecek SEMUA run di paragraf persamaan, termasuk badan
            # rumusnya sendiri (huruf italic matematis dst, yang WAJAR
            # tidak bold) -- akibatnya cek ini nyaris SELALU gagal
            # walau nomornya sendiri sudah bold dengan benar, karena
            # badan rumus yang non-bold ikut kehitung. Panduan cuma
            # mewajibkan NOMOR-nya yang tebal, bukan seluruh persamaan.
            # Dipetakan lewat offset karakter kumulatif di full_text
            # (span hasil match), konsisten dengan invarian
            # "".join(r["text"] for r in runs) == full_text yang sudah
            # dijaga di extractor.py/pdf_extractor.py.
            span_start, span_end = m.span()
            offset = 0
            number_runs = []
            for r in elem["runs"]:
                r_start, r_end = offset, offset + len(r["text"])
                if r_start < span_end and r_end > span_start and r["text"].strip():
                    number_runs.append(r)
                offset = r_end
            if number_runs and any(r["bold"] is not True for r in number_runs):
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": (
                            "Nomor identitas persamaan seharusnya dicetak "
                            "tebal (bold)."
                        ),
                    }
                )

    # --- TABEL ---
    # Rata tengah TIDAK dicek di sini -- itu aturan untuk CAPTION-nya
    # (ptype "caption", sudah dicek di compare_paragraphs), bukan untuk
    # tabelnya sendiri. Baris kosong di atas/bawah juga tidak dicek,
    # lihat alasan di docstring atas.
    for tbl in tables:
        chapter = _chapter_of_table(paragraphs_by_index, tbl)
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        loc = f"Tabel ke-{tbl['seq']} dalam dokumen"

        # Cek sampai 3 paragraf terdekat di tiap arah (bukan cuma yang
        # persis bersebelahan) -- lihat _nearby_paragraphs() untuk
        # alasan (baris kosong dan/atau "Sumber: ..." di antara tabel
        # dan caption-nya).
        nearby_elems = _nearby_paragraphs(
            paragraphs_by_index, tbl["prev_para_idx"], "prev"
        ) + _nearby_paragraphs(paragraphs_by_index, tbl["next_para_idx"], "next")

        # Tabel di halaman template tetap (sampul, halaman judul
        # dalam, halaman pengesahan, surat pernyataan keaslian) BUKAN
        # tabel data -- lazimnya tabel di halaman-halaman ini dipakai
        # murni buat TATA LETAK (mis. blok tanda tangan di halaman
        # pengesahan), jadi wajar tidak punya caption "Tabel N.N ...".
        # Paragraf tetangga tabel semacam ini sudah ber-ptype "cover"
        # (lihat extractor.py) -- dipakai sebagai sinyal buat
        # mengecualikan tabel ini TOTAL dari cek caption, bukan cuma
        # dari cek caption tapi tetap kena "checked" (yang bikin skor
        # bab Pra-Bab turun padahal bukan pelanggaran beneran).
        if any(e["type"] == "cover" for e in nearby_elems):
            continue

        # Tabel 1 baris BUKAN tabel data -- dalam praktiknya selalu
        # dipakai sebagai trik layout (nomor persamaan sejajar rumus,
        # atau "kotak" pembungkus teks/kode/prompt), bukan buat
        # nyajiin data yang wajar diberi caption "Tabel N.N ...".
        # Tabel data sungguhan (perbandingan penelitian, metadata,
        # dst) selalu >1 baris (header + minimal 1 baris data), jadi
        # exclude di sini TOTAL (skip dari "checked" juga) supaya
        # tidak ikut menurunkan skor karena pelanggaran yang bukan
        # pelanggaran beneran -- sama seperti alasan exclude "cover".
        if tbl.get("n_rows", 0) <= 1:
            continue

        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): cek "tabel
        # wajib punya caption 'Tabel N.N ...' tepat di atas/bawahnya"
        # dikomentari, bukan dihapus.
        # has_caption_adjacent = any(
        #     e["type"] == "caption" and e["full_text"].startswith("Tabel")
        #     for e in nearby_elems
        # )
        # checks[chapter]["checked"] += 1
        # if not has_caption_adjacent:
        #     checks[chapter]["failed"] += 1
        #     findings.append(
        #         {
        #             "category": "Format",
        #             "chapter": chapter,
        #             "location": loc,
        #             "description": (
        #                 "Tabel tidak punya nomor identitas & keterangan "
        #                 "('Tabel N.N ...') tepat di atas/bawahnya."
        #             ),
        #         }
        #     )

        # --- judul kolom (baris header) tabel: rata tengah, tebal,
        #     berlatar-belakang gelap -- sesuai anjuran panduan
        #     ("Judul pada tabel DAPAT dituliskan rata tengah, tebal
        #     dan berlatar-belakang agak gelap"). header_format sudah
        #     None untuk tabel non-data (n_rows<=1) dari extractor.py,
        #     jadi tidak perlu dicek ulang di sini -- kalau None berarti
        #     memang tidak relevan, bukan berarti gagal.
        # DINONAKTIFKAN SELURUHNYA (atas permintaan user, 2026-08-24): cek
        # format baris header tabel ("rata tengah, tebal, berlatar-belakang
        # gelap") dinonaktifkan total -- sebelumnya cuma syarat "tebal"
        # yang masih aktif (2026-08-23), tapi itu pun terlalu sering false
        # positive di jalur PDF (header_all_bold dari tabel hasil render PDF
        # tidak selalu akurat mendeteksi bold asli), jadi seluruh blok cek
        # ini dimatikan. Dikomentari, bukan dihapus, supaya bisa diaktifkan
        # lagi kalau deteksi bold tabel dari PDF sudah lebih andal.
        # hf = tbl.get("header_format")
        # if hf is not None:
        #     checks[chapter]["checked"] += 1
        #     bad = []
        #     # if not hf["header_all_center"]:
        #     #     bad.append("rata tengah")
        #     if not hf["header_all_bold"]:
        #         bad.append("tebal")
        #     # if not hf["header_has_shading"]:
        #     #     bad.append("berlatar-belakang gelap")
        #     if bad:
        #         checks[chapter]["failed"] += 1
        #         findings.append(
        #             {
        #                 "category": "Format",
        #                 "chapter": chapter,
        #                 "location": loc,
        #                 "description": (
        #                     f"Judul kolom (baris header) tabel belum "
        #                     f"{', '.join(bad)}, sesuai anjuran panduan."
        #                 ),
        #             }
        #         )

    # --- SUMBER KUTIPAN (opsional, dicek HANYA kalau ada) ---
    for elem in paragraphs:
        if elem["type"] != "source_citation":
            continue
        chapter = elem["chapter"]
        if _is_excluded_chapter(chapter):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue
        loc = f"Paragraf #{elem['index']} ('{text_preview(elem['full_text'])}')"

        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): dikomentari,
        # bukan dihapus -- syarat rata tengah utk baris "Sumber:"
        # dilonggarkan.
        # checks[chapter]["checked"] += 1
        # if elem["alignment"] != "center":
        #     checks[chapter]["failed"] += 1
        #     findings.append(
        #         {
        #             "category": "Format",
        #             "chapter": chapter,
        #             "location": loc,
        #             "description": "Baris 'Sumber:' seharusnya ditulis rata tengah.",
        #         }
        #     )

        runs_with_size = [r for r in elem["runs"] if r["font_size_pt"] is not None]
        if runs_with_size:
            checks[chapter]["checked"] += 1
            bad = any(
                abs(r["font_size_pt"] - RULES["caption_font_size"]) > 0.1
                for r in runs_with_size
            )
            if bad:
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": (
                            f"Baris 'Sumber:' seharusnya berukuran "
                            f"{RULES['caption_font_size']}pt."
                        ),
                    }
                )

    return findings, dict(checks)


def compare_abstract(paragraphs: list[dict]):
    """
    Cek FORMAT VISUAL Abstrak (ID) & Abstract (EN), berbasis field
    elem["abstract_part"] yang di-set extractor.py selama paragraf
    berada di region Abstrak/Abstract:

      1. Abstrak (ID) wajib 300-500 kata (dihitung dari SELURUH
         paragraf ber-abstract_part=="abstrak_body", bukan cuma satu
         paragraf pertama -- abstrak mahasiswa kadang ditulis lebih
         dari satu paragraf).
      2. Kata Kunci / Keywords (ID maupun EN, kalau ada) wajib
         maksimum 5 frasa, dipisahkan tanda koma.
      3. Abstract (EN) -- terjemahan abstrak -- wajib dicetak italic
         (miring) SELURUHNYA.

    SENGAJA TIDAK mengecek isi/konten (latar belakang, permasalahan,
    solusi, kinerja) -- itu butuh pemahaman semantik, bukan format
    visual/hitungan, jadi ranahnya pipeline sistematika (RAG), bukan
    comparator rule-based ini.
    """
    findings = []
    checks = defaultdict(_new_counter)

    id_body = [e for e in paragraphs if e.get("abstract_part") == "abstrak_body"]
    en_body = [e for e in paragraphs if e.get("abstract_part") == "abstract_body"]
    keyword_elems = [
        e for e in paragraphs if e.get("abstract_part") in ("keyword_id", "keyword_en")
    ]

    # --- cek 1: jumlah kata Abstrak (ID) ---
    if id_body:
        chapter = id_body[0]["chapter"]
        word_count = sum(len(e["full_text"].split()) for e in id_body)
        lo_idx, hi_idx = id_body[0]["index"], id_body[-1]["index"]
        loc = (
            f"Abstrak (Paragraf #{lo_idx})"
            if lo_idx == hi_idx
            else f"Abstrak (Paragraf #{lo_idx}-#{hi_idx})"
        )
        lo, hi = RULES["abstract_min_words"], RULES["abstract_max_words"]

        checks[chapter]["checked"] += 1
        if not (lo <= word_count <= hi):
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"Abstrak berisi {word_count} kata, seharusnya "
                        f"{lo}-{hi} kata."
                    ),
                }
            )

    # --- cek 2: Kata Kunci / Keywords maksimum N frasa ---
    max_kw = RULES["abstract_max_keywords"]
    for elem in keyword_elems:
        chapter = elem["chapter"]
        loc = f"Paragraf #{elem['index']} ('{text_preview(elem['full_text'])}')"
        label = "Kata Kunci" if elem["abstract_part"] == "keyword_id" else "Keywords"
        body_text = KEYWORD_STRIP_RE.sub("", elem["full_text"], count=1)
        phrases = [p.strip() for p in body_text.split(",") if p.strip()]

        checks[chapter]["checked"] += 1
        if len(phrases) > max_kw:
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"{label} berjumlah {len(phrases)} frasa, seharusnya "
                        f"maksimum {max_kw} frasa (dipisahkan tanda koma)."
                    ),
                }
            )

    # --- cek 3: Abstract (EN) wajib italic seluruhnya ---
    # BUG DIPERBAIKI (jalur PDF -- satu Abstract (EN) yang SUDAH italic
    # 100% muncul sebagai 3-4 finding terpisah "seharusnya italic"):
    # sebelumnya loop ini menjadikan TIAP elemen di en_body (1 elemen PDF =
    # 1 "block" hasil deteksi PyMuPDF, lihat keterbatasan #5 di
    # pdf_extractor.py) sebagai UNIT CEK TERPISAH. Abstract (EN) yang
    # panjangnya satu paragraf utuh sering ke-pecah PyMuPDF jadi beberapa
    # block (jarak antar baris yang sedikit tidak rata di PDF hasil
    # render), padahal secara logis itu SATU paragraf abstrak yang sama.
    # Akibatnya finding "seharusnya italic" muncul berkali-kali utk
    # paragraf yang sama, walau isinya sudah 100% italic.
    # Fix: gabungkan SEMUA elemen en_body jadi SATU unit cek -- persis pola
    # yang sudah dipakai cek 1 (jumlah kata Abstrak ID) di atas, yang
    # sengaja menjumlahkan seluruh id_body jadi satu angka/satu location.
    if en_body:
        chapter = en_body[0]["chapter"]
        lo_idx, hi_idx = en_body[0]["index"], en_body[-1]["index"]
        loc = (
            f"Abstract (Paragraf #{lo_idx})"
            if lo_idx == hi_idx
            else f"Abstract (Paragraf #{lo_idx}-#{hi_idx})"
        )
        runs_with_text = [r for e in en_body for r in e["runs"] if r["text"].strip()]

        if runs_with_text:
            checks[chapter]["checked"] += 1
            if any(r["italic"] is not True for r in runs_with_text):
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": (
                            "Teks Abstract (terjemahan Bahasa Inggris) "
                            "seharusnya dicetak italic (miring) secara "
                            "keseluruhan."
                        ),
                    }
                )

    return findings, dict(checks)


def compare_quotations(paragraphs: list[dict]):
    """
    Cek format KUTIPAN LANGSUNG (direct quotation) berbasis
    quotation_checker.check_paragraph() -- lihat modul itu untuk detail
    lengkap cara deteksi & disambiguasi kandidat kutipan dari istilah/
    dialog biasa yang kebetulan diberi tanda kutip.

      1. Kutipan langsung PENDEK (span kutipan cuma sebagian dari
         paragraf): wajib italic seluruhnya.
      2. Kutipan langsung PANJANG (span kutipan mendominasi/mencakup
         hampir seluruh paragraf -- proxy untuk "ditulis sebagai
         paragraf terpisah"): wajib font 10pt DAN indentasi kiri+kanan
         masuk 10mm dari batas tulisan, dicek di level PARAGRAF
         (bukan span) karena "font 10pt" & "indentasi" adalah atribut
         paragraf/run penuh, bukan cuma karakter di dalam tanda kutip.

    Cakupan paragraf yang diperiksa: HANYA ptype "body", dengan
    pengecualian SAMA PERSIS seperti cek istilah asing wajib italic di
    compare_paragraphs() (lihat komentar di sana untuk alasan
    lengkapnya):
      - DAFTAR PUSTAKA dikecualikan total -- entri bibliografi sering
        memuat judul artikel/jurnal dalam tanda kutip yang BUKAN
        kutipan langsung isi tulisan.
      - Paragraf Abstract (EN) (abstract_part "abstract_body"/
        "keyword_en") dikecualikan -- paragraf itu sudah wajib italic
        SELURUHNYA lewat cek terpisah di compare_abstract(), bukan
        konteks kutipan langsung body biasa.
      - LAMPIRAN & BIODATA PENULIS dikecualikan total (lihat
        _is_excluded_chapter()).

    SENGAJA TIDAK DICEK:
      - Kutipan TIDAK LANGSUNG (parafrase) -- panduan cuma mensyaratkan
        "sitasi ditulis setelah kalimat selesai", tapi tidak ada
        penanda tekstual yang bisa dipakai untuk membedakan kalimat
        parafrase yang butuh sitasi dari kalimat body biasa yang
        memang tidak perlu sitasi (beda dengan kutipan langsung, yang
        penanda tanda-kutipnya jelas). Memaksakan cek ini berisiko
        sangat overclaim false positive (banyak kalimat body valid
        yang tidak butuh sitasi akan ikut tertandai), jadi sengaja
        dilewati -- perlu verifikasi manual.
      - Jumlah baris fisik PERSIS (1 vs 2 vs >2 baris) -- python-docx
        tidak tahu pembungkusan baris hasil render, dipakai proxy
        cakupan span terhadap paragraf (lihat quotation_checker.py).
    """
    findings = []
    checks = defaultdict(_new_counter)
    quote_font = RULES["direct_quote_long_font_size"]
    quote_indent = RULES["direct_quote_indent_cm"]
    quote_tol = RULES["direct_quote_indent_tolerance_cm"]

    for elem in paragraphs:
        if elem["type"] != "body":
            continue
        if elem["chapter"] == "DAFTAR PUSTAKA":
            continue
        if elem.get("abstract_part") in ("abstract_body", "keyword_en"):
            continue
        if _is_excluded_chapter(elem["chapter"]):
            # LAMPIRAN & BIODATA PENULIS dikecualikan total dari
            # pengecekan format -- lihat docstring _is_excluded_chapter().
            continue

        chapter = elem["chapter"]
        loc = f"Paragraf #{elem['index']} ('{text_preview(elem['full_text'])}')"

        for quote in check_quotations(elem["runs"]):
            preview = text_preview(quote["quoted_text"])

            if quote["is_long"]:
                # --- kutipan langsung PANJANG: font 10pt + indentasi
                #     kiri & kanan 10mm, dicek level paragraf ---
                bad = []

                runs_with_size = [
                    r for r in elem["runs"] if r["font_size_pt"] is not None
                ]
                if runs_with_size and any(
                    abs(r["font_size_pt"] - quote_font) > 0.1 for r in runs_with_size
                ):
                    bad.append(f"ukuran font seharusnya {quote_font}pt")

                li = elem.get("left_indent_cm")
                if li is None or abs(li - quote_indent) > quote_tol:
                    bad.append(f"indentasi kiri seharusnya {quote_indent * 10:.0f}mm")

                ri = elem.get("right_indent_cm")
                if ri is None or abs(ri - quote_indent) > quote_tol:
                    bad.append(f"indentasi kanan seharusnya {quote_indent * 10:.0f}mm")

                checks[chapter]["checked"] += 1
                if bad:
                    checks[chapter]["failed"] += 1
                    findings.append(
                        {
                            "category": "Format",
                            "chapter": chapter,
                            "location": loc,
                            "description": (
                                f"Kutipan langsung panjang ('{preview}') tidak "
                                f"sesuai format: {'; '.join(bad)}."
                            ),
                        }
                    )
            else:
                # --- kutipan langsung PENDEK: wajib italic seluruhnya ---
                checks[chapter]["checked"] += 1
                if quote["problems"]:
                    checks[chapter]["failed"] += 1
                    findings.append(
                        {
                            "category": "Format",
                            "chapter": chapter,
                            "location": loc,
                            "description": (
                                f"Kutipan langsung pendek ('{preview}') "
                                f"{', '.join(quote['problems'])}."
                            ),
                        }
                    )

    return findings, dict(checks)


def compare_page_numbering(
    paragraphs: list[dict],
    sections: list[dict],
    section_pages: dict[int, int] | None = None,
):
    """
    Cek aturan penomoran halaman, SENGAJA DIKOSONGKAN untuk Pra-Bab
    (permintaan user eksplisit): mulai dihitung dari BAB 1 saja.
      1. Section tempat "BAB 1" dimulai wajib format Arab (pgNumType
         fmt="decimal") DAN restart dari 1 -- supaya BAB 1 memang
         halaman "1", bukan lanjutan romawi pra-bab.
      2. Semua section dari BAB 1 dan seterusnya wajib punya field
         nomor halaman (PAGE) di footer, rata tengah.

    Cek 0 (section SEBELUM BAB 1 wajib format Romawi kecil) DIHAPUS
    atas permintaan user eksplisit -- aturan nomor halaman di Pra-Bab
    (dari setelah cover sampai sebelum BAB 1) sengaja dikosongkan,
    berlaku sama untuk docx maupun pdf. Blok lama dikomentari di bawah
    (bukan dihapus fisik) supaya gampang ditelusuri/diaktifkan lagi
    kalau kebijakan ini berubah.

    Cek keberadaan Section Break sebelum BAB 1 juga DIHAPUS atas
    permintaan user (berlaku untuk docx maupun pdf).
    """
    section_pages = section_pages or {}

    def _loc(idx: int) -> str:
        page = section_pages.get(idx)
        return f"Halaman {page} (Section #{idx})" if page else f"Section #{idx}"

    findings = []
    checks = defaultdict(_new_counter)
    group = "Umum (Penomoran Halaman)"

    # Cocokkan PREFIX "BAB 1", bukan exact-equality ke "BAB 1" -- teks
    # paragraf aslinya bisa "BAB 1" polos ATAU "BAB 1 PENDAHULUAN"
    # (nama bab menyatu di baris yang sama). exact-equality lama gagal
    # total untuk format kedua, bikin bab1_section_idx selalu None dan
    # SELURUH blok cek penomoran halaman di bawah ini silently ke-skip
    # (bukan dianggap salah, tapi memang tidak pernah dicek sama sekali).
    BAB1_TITLE_RE = re.compile(r"^BAB\s+1(\s+\S.*)?$")
    bab1_section_idx = next(
        (
            elem["section"]
            for elem in paragraphs
            if elem["type"] == "chapter_title"
            and BAB1_TITLE_RE.match(elem["full_text"])
        ),
        None,
    )
    if bab1_section_idx is None:
        return findings, checks

    bab1_section = next((s for s in sections if s["index"] == bab1_section_idx), None)
    if bab1_section is None:
        return findings, checks

    # --- cek 0: semua section SEBELUM BAB 1 wajib romawi kecil ---
    # Dicek pakai pgnum_fmt_effective (hasil warisan), bukan pgnum_fmt
    # mentah -- kalau cuma section 0 yang eksplisit set "lowerRoman" dan
    # section 1, 2 dst di Pra-Bab tidak override apa-apa, mereka tetap
    # harus dianggap romawi (warisan), bukan ke-skip / dianggap salah.
    # BUG DIPERBAIKI (2026-09-07): pgnum_fmt jalur PDF/RTF diisi None di
    # pdf_extractor.py sendiri (lihat komentar di extract_section_params_
    # pdf) -- nilai SEBENARNYA baru ditambahkan belakangan oleh
    # page_layout_checker.py lewat pembacaan angka footer beneran, lalu
    # digabung ke dict section ini di main.py. Kalau pembacaan footer
    # untuk SATU halaman gagal (mis. angka halaman tidak kebaca jelas
    # oleh render), pgnum_fmt section itu bisa TETAP None sampai ke sini.
    # Kode LAMA membandingkan "None != 'lowerRoman'" -> otomatis True,
    # jadi section yang formatnya SEBENARNYA sudah benar (tapi gagal
    # dibaca) ikut dilaporkan SALAH -- False positive yang HANYA bisa
    # terjadi di jalur PDF/RTF (docx tidak pernah punya pgnum_fmt None),
    # sehingga secara sistematis merugikan skor kedua jalur itu
    # dibanding docx. None sekarang dianggap "belum bisa diverifikasi",
    # BUKAN "salah" -- section dengan pgnum_fmt None dilewati dari
    # checked/failed sama sekali, supaya tidak menjatuhkan skor karena
    # keterbatasan pembacaan, tapi tetap dicatat sebagai catatan
    # informatif (tidak masuk hitungan skor) supaya tidak hilang dari
    # laporan begitu saja.
    # DIHAPUS (atas permintaan user eksplisit): aturan format Romawi
    # kecil untuk section sebelum BAB 1 sengaja dikosongkan -- Pra-Bab
    # (dari setelah cover sampai sebelum BAB 1) TIDAK LAGI dicek format
    # penomoran halamannya sama sekali, baik docx maupun pdf. Hitungan
    # aturan nomor halaman sekarang mulai dari BAB 1 saja (cek 1 & cek 2
    # di bawah). Blok lama dikomentari (bukan dihapus fisik) untuk
    # referensi/gampang diaktifkan lagi kalau kebijakan ini berubah.
    #
    # pre_bab1_sections = [s for s in sections if s["index"] < bab1_section_idx]
    # if pre_bab1_sections:
    #     verifiable = [s for s in pre_bab1_sections if s["pgnum_fmt"] is not None]
    #     unverifiable = [s for s in pre_bab1_sections if s["pgnum_fmt"] is None]
    #     if unverifiable:
    #         findings.append(
    #             {
    #                 "category": "Info",
    #                 "chapter": group,
    #                 "location": ", ".join(_loc(s["index"]) for s in unverifiable),
    #                 "description": (
    #                     "Format penomoran halaman di section tsb tidak bisa "
    #                     "diverifikasi otomatis (angka footer tidak terbaca "
    #                     "jelas) -- mohon cek manual, tidak ikut dihitung ke "
    #                     "skor."
    #                 ),
    #             }
    #         )
    #     if verifiable:
    #         checks[group]["checked"] += 1
    #         offending = [
    #             sec["index"] for sec in verifiable if sec["pgnum_fmt"] != "lowerRoman"
    #         ]
    #         if offending:
    #             checks[group]["failed"] += 1
    #             detail = ", ".join(_loc(idx) for idx in offending)
    #             findings.append(
    #                 {
    #                     "category": "Format",
    #                     "chapter": group,
    #                     "location": detail,
    #                     "description": (
    #                         "Penomoran halaman di section tsb (sebelum BAB 1) "
    #                         "harus format angka Romawi kecil (i, ii, iii, ...)."
    #                     ),
    #                 }
    #             )

    # --- cek 1: format Arab + restart dari 1 ---
    # Sama seperti cek 0 di atas: pgnum_fmt/pgnum_start None (jalur
    # PDF/RTF, gagal dibaca dari footer) berarti "belum bisa
    # diverifikasi", bukan otomatis "salah" -- lihat catatan lengkap di
    # cek 0.
    if bab1_section["pgnum_fmt"] is None or bab1_section["pgnum_start"] is None:
        findings.append(
            {
                "category": "Info",
                "chapter": group,
                "location": f"Section {bab1_section_idx} (mulai BAB 1)",
                "description": (
                    "Format penomoran halaman di section awal BAB 1 tidak "
                    "bisa diverifikasi otomatis (angka footer tidak terbaca "
                    "jelas) -- mohon cek manual, tidak ikut dihitung ke skor."
                ),
            }
        )
    else:
        checks[group]["checked"] += 1
        if not (
            bab1_section["pgnum_fmt"] == "decimal" and bab1_section["pgnum_start"] == 1
        ):
            checks[group]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": group,
                    "location": f"Section {bab1_section_idx} (mulai BAB 1)",
                    "description": (
                        "Penomoran halaman di section awal BAB 1 harus format angka "
                        "Arab dan dimulai ulang dari 1, terdeteksi "
                        f"fmt={bab1_section['pgnum_fmt']!r}, "
                        f"start={bab1_section['pgnum_start']!r}."
                    ),
                }
            )

    # --- cek 2: field PAGE di footer, dari BAB 1 dst ---
    # DILONGGARKAN (atas permintaan user, 2026-08-23): syarat
    # "footer_alignment == center" dihapus dari kondisi -- yang masih
    # wajib cuma ADANYA field nomor halaman (footer_has_page_field),
    # posisi rata tengah tidak lagi dipaksakan. Baris lama dikomentari,
    # bukan dihapus.
    # Sama seperti cek 0 & cek 1: footer_has_page_field bisa None
    # (belum berhasil dibaca), yang HARUS dibedakan dari False
    # (berhasil dibaca, terbukti memang tidak ada field-nya) -- "not
    # None" == True secara Python, jadi tanpa pemisahan ini section
    # yang belum sempat dibaca akan ikut ke-flag "tidak punya nomor
    # halaman" walau sebenarnya belum tentu begitu.
    post_bab1_sections = [s for s in sections if s["index"] >= bab1_section_idx]
    verifiable = [
        s for s in post_bab1_sections if s["footer_has_page_field"] is not None
    ]
    unverifiable = [s for s in post_bab1_sections if s["footer_has_page_field"] is None]
    if unverifiable:
        findings.append(
            {
                "category": "Info",
                "chapter": group,
                "location": ", ".join(_loc(s["index"]) for s in unverifiable),
                "description": (
                    "Keberadaan field nomor halaman di footer section tsb "
                    "tidak bisa diverifikasi otomatis -- mohon cek manual, "
                    "tidak ikut dihitung ke skor."
                ),
            }
        )
    if verifiable:
        checks[group]["checked"] += 1
        offending = [
            sec["index"]
            for sec in verifiable
            # and not (sec["footer_has_page_field"] and sec["footer_alignment"] == "center")
            if not sec["footer_has_page_field"]
        ]
        if offending:
            checks[group]["failed"] += 1
            detail = ", ".join(_loc(idx) for idx in offending)
            findings.append(
                {
                    "category": "Format",
                    "chapter": group,
                    "location": detail,
                    "description": (
                        f"Footer di section tsb tidak punya nomor halaman, "
                        f"padahal sudah masuk BAB 1 atau setelahnya."
                    ),
                }
            )

    return findings, checks


def deduplicate_findings(findings: list[dict]) -> list[dict]:
    seen = set()
    deduped = []
    for f in findings:
        key = (f["location"], f["description"])
        if key not in seen:
            seen.add(key)
            deduped.append(f)
    return deduped


# ===================== HALAMAN SAMPUL (COVER) =====================
# BARU (2026-09-17, permintaan user, berdasar lembar aturan cover resmi
# PENS): sebelumnya ptype "cover" DIKECUALIKAN TOTAL dari comparator
# (lihat "if ptype == 'cover': continue" di compare_paragraphs() di
# atas) -- itu TIDAK berubah. compare_cover_page() ini FUNGSI TERPISAH
# yang secara eksplisit memproses elemen ber-ptype "cover", dipanggil
# sendiri oleh compare_all(), supaya halaman sampul yang tadinya "bebas
# sesuai template" sekarang tetap dicek sesuai lembar aturan berikut:
#   - "PROYEK AKHIR"                         : ALL CAPS, 20pt, bold
#   - Judul Proyek Akhir                     : ALL CAPS, 18pt, bold
#   - Nama Mahasiswa & NRP                   : Title Case (nama), 16pt, bold
#   - "DOSEN PEMBIMBING"                     : ALL CAPS, 16pt, bold
#   - Nama Pembimbing 1/2 & NIP              : Title Case (nama), 16pt, bold
#   - "PROGRAM STUDI SARJANA TERAPAN" & isian: ALL CAPS, 16pt, bold
#   - "POLITEKNIK ELEKTRONIKA NEGERI SURABAYA" & tahun: 16pt, bold (ALL
#     CAPS untuk baris nama politeknik; baris tahun cuma angka)
#   - Warna latar belakang halaman sampul    : #002FA7 (RGB 0, 47, 167)
#
# TANTANGAN UTAMA dibanding checker lain di modul ini: HAMPIR SEMUA
# baris di halaman sampul dokumen SUNGGUHAN mahasiswa adalah data
# isian (nama, NRP, NIP, judul beneran) yang MENGGANTI placeholder
# template ("Nama Mahasiswa", "NRP. ....", dst) -- jadi baris-baris
# itu TIDAK BISA dikenali lewat pencocokan teks literal seperti
# CHAPTER_TITLE_RE dkk. Deteksi di sini pakai STATE MACHINE berbasis
# LABEL TETAP yang memang TIDAK PERNAH diganti mahasiswa ("DOSEN
# PEMBIMBING", "PROGRAM STUDI SARJANA TERAPAN", "POLITEKNIK ELEKTRONIKA
# NEGERI SURABAYA") sebagai penanda transisi antar blok -- baris DI
# ANTARA dua label tetap itulah yang diperlakukan sebagai data isian,
# dicek FORMATnya (ukuran, bold, kapitalisasi) TANPA mensyaratkan
# isinya sama persis dengan placeholder.
#
# ATURAN RESMI (lembar aturan cover buku Proyek Akhir PENS) yang
# diimplementasikan oleh state machine + COVER_RULES di bawah:
#   "Halaman judul sama dengan halaman kulit luar, berisi: tulisan
#   'PROYEK AKHIR' (ukuran huruf 20, huruf kapital semua), judul
#   Proyek Akhir (ukuran huruf 18, huruf kapital semua); lambang PENS,
#   nama lengkap penulis dengan nomor pokok (NRP) dan nama lengkap
#   dosen pembimbing beserta gelar (ukuran huruf 16); nama program
#   studi, departemen, politeknik (ditulis lengkap), serta tahun
#   pembuatan (ukuran huruf 16, huruf kapital semua). Judul Proyek
#   Akhir harus diketik seluruhnya dengan huruf kapital dan tidak
#   boleh disingkat, kecuali singkatan yang sudah baku. Nama penulis
#   dan dosen pembimbing harus ditulis dengan huruf kapital di awal
#   kata dan tidak disingkat. Halaman sampul buku proyek akhir
#   berwarna biru dengan kode hexa '#002FA7' (RGB 0, 47, 167)."
#
# CATATAN PENTING (bug diperbaiki 2026-09-17): ptype "cover" di
# extractor.py/pdf_extractor.py MELIPUTI SELURUH halaman template
# tetap bagian depan buku (sampul, halaman judul dalam, lembar
# pengesahan, pernyataan orisinalitas, dst) sebagai satu kelompok --
# BUKAN cuma halaman sampul/judul yang aturan di atas berlaku
# untuknya. compare_cover_page() cuma boleh MENERAPKAN aturan di atas
# untuk BLOK PERTAMA (label "PROYEK AKHIR" s.d. baris tahun setelah
# blok politeknik) lalu WAJIB BERHENTI (state "done") -- halaman
# template lain (pengesahan, pernyataan orisinalitas) punya
# formatnya sendiri yang TIDAK diatur oleh lembar aturan cover ini,
# jadi tidak boleh ikut dicek pakai aturan cover. SEBELUM perbaikan
# ini, state machine tidak pernah berhenti (state "tahun_block" lama
# mencocokkan SETIAP baris sisanya sebagai "detail_size", termasuk
# semua baris di halaman pengesahan/orisinalitas berikutnya) --
# itulah penyebab utama findings palsu yang membanjiri "Pra-Bab".
#
# CATATAN JUDUL MULTI-BARIS (bug diperbaiki 2026-09-17): kalau judul
# Proyek Akhir dipisah jadi lebih dari satu paragraf/baris (word-wrap
# otomatis Word atau Enter manual per baris), SEMUA baris itu
# tetap ALL CAPS (mengikuti aturan resmi di atas) -- beda dari baris
# nama mahasiswa sesudahnya yang HARUS Title Case, bukan ALL CAPS.
# Perbedaan kapitalisasi inilah yang dipakai sebagai penanda transisi
# antara "masih judul" vs "sudah masuk nama mahasiswa": selama baris
# demi baris ALL CAPS, tetap diperlakukan sebagai lanjutan judul
# (dicek title_size); begitu ketemu baris yang BUKAN ALL CAPS,
# baris itu (dan seterusnya) diperlakukan sebagai blok Nama Mahasiswa.
COVER_PROYEK_AKHIR_LABEL_RE = re.compile(r"^PROYEK\s+AKHIR$")
COVER_DOSEN_PEMBIMBING_RE = re.compile(r"^DOSEN\s+PEMBIMBING$")
COVER_PROGRAM_STUDI_RE = re.compile(r"^PROGRAM\s+STUDI\s+SARJANA\s+TERAPAN\b")
COVER_POLITEKNIK_RE = re.compile(r"^POLITEKNIK\s+ELEKTRONIKA\s+NEGERI\s+SURABAYA\b")
# "NRP. ..........." / "NIP ..................." -- baris nomor, bukan
# nama, jadi DIKECUALIKAN dari cek Title Case (nomor tidak punya
# "kapital di awal kata").
COVER_NRP_NIP_RE = re.compile(r"^(NRP|NIP)\b", re.IGNORECASE)

COVER_RULES = {
    "label_size": 20,  # tulisan "PROYEK AKHIR"
    "title_size": 18,  # judul Proyek Akhir (boleh >1 baris, semua ALL CAPS)
    "detail_size": 16,  # nama+NRP, dosen pembimbing+gelar, prodi/departemen/politeknik, tahun
    "background_hex": "002FA7",  # RGB 0, 47, 167
}


# Toleransi jarak warna (skala 0-255 per kanal RGB) buat dokumen PDF:
# warna latar sampul PDF didekati lewat vector rect ATAU sampling piksel
# render halaman (lihat _get_cover_background_color_hex_pdf() di
# pdf_extractor.py) -- kompresi gambar (mis. sampul ditempel sebagai
# JPEG) atau anti-aliasing rendering bisa menggeser nilai RGB sedikit
# dari warna aslinya walau SECARA VISUAL identik (contoh nyata:
# "#002AA6" ke-detect padahal sampul aslinya "#002FA7" -- selisih cuma
# beberapa satuan per kanal, tidak terlihat mata, bukan kesalahan
# mahasiswa). Threshold 24 dipilih supaya toleran ke pergeseran wajar
# akibat kompresi/anti-aliasing, tapi tetap menangkap warna yang
# SUNGGUH beda (mis. biru terlalu terang/gelap, atau warna lain sama
# sekali) -- jarak Euclidean RGB untuk kesalahan sungguhan biasanya
# jauh lebih besar dari ini.
COVER_BACKGROUND_COLOR_TOLERANCE = 24


def _hex_color_close_enough(actual_hex: str, expected_hex: str) -> bool:
    """True kalau actual_hex "cukup dekat" dengan expected_hex (lihat
    COVER_BACKGROUND_COLOR_TOLERANCE)."""
    try:
        a = actual_hex.strip().lstrip("#")
        e = expected_hex.strip().lstrip("#")
        ar, ag, ab = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
        er, eg, eb = int(e[0:2], 16), int(e[2:4], 16), int(e[4:6], 16)
    except (ValueError, IndexError):
        # Hex tidak valid/tidak lengkap -- dianggap TIDAK cocok (bukan
        # dianggap cocok secara diam-diam) supaya tetap kelihatan.
        return False
    dist = ((ar - er) ** 2 + (ag - eg) ** 2 + (ab - eb) ** 2) ** 0.5
    return dist <= COVER_BACKGROUND_COLOR_TOLERANCE


def _cover_dominant_size(elem):
    """Sama seperti pola font-size mayoritas di compare_paragraphs(), versi ringkas untuk sampul."""
    runs = [
        r for r in elem["runs"] if r["font_size_pt"] is not None and r["text"].strip()
    ]
    if not runs:
        return None
    weight = defaultdict(int)
    for r in runs:
        weight[r["font_size_pt"]] += len(r["text"].strip())
    return max(weight, key=weight.get)


def _cover_fully_bold(elem) -> bool:
    runs = [r for r in elem["runs"] if r["text"].strip()]
    return bool(runs) and all(r["bold"] is True for r in runs)


def _cover_is_caps(elem) -> bool:
    """Kapital semua SECARA VISUAL (teks kapital ATAU font "All caps").
    Jalur PDF tidak punya is_visual_upper -- teks hasil ekstraksinya
    sudah apa yang tercetak, cukup is_upper."""
    return bool(elem.get("is_visual_upper", elem["is_upper"]))


COVER_NRP_NIP_ANYWHERE_RE = re.compile(r"\b(NRP|NIP)\b", re.IGNORECASE)


def _cover_is_name_line(cover_elems, idx, has_nrp_line: bool) -> bool:
    """
    True kalau baris cover ke-idx (setelah label "PROYEK AKHIR") sudah
    masuk blok Nama Mahasiswa, BUKAN lanjutan judul.

    BUG YANG DIPERBAIKI (2026-09-21, laporan user -- "judul di sampul
    sudah 18pt sesuai aturan tapi disalahkan, seharusnya 16pt"):
    heuristik lama menganggap judul sebagai "baris ALL CAPS" dan baris
    pertama yang BUKAN ALL CAPS sebagai nama mahasiswa. Judul yang
    ditulis huruf biasa ("Contextual Analysis untuk Opinion Mining ...")
    -- atau tampil kapital hanya lewat font "All caps" sehingga
    teksnya sendiri huruf biasa -- jadi salah dikira nama dan dicek
    dengan ukuran 16pt.

    Penanda yang lebih andal: NAMA mahasiswa SELALU langsung diikuti
    baris "NRP ..." (dan nama pembimbing oleh "NIP ..."). Jadi judul =
    semua baris SEBELUM baris yang langsung mendahului baris NRP
    pertama, tanpa peduli huruf besar/kecilnya. Kalau sampul sama
    sekali tidak punya baris NRP/NIP, jatuh ke heuristik lama
    (ALL CAPS) ditambah petunjuk ukuran font judul.
    """
    elem = cover_elems[idx]
    text = elem["full_text"]
    if COVER_NRP_NIP_ANYWHERE_RE.search(text):
        return True
    if has_nrp_line:
        nxt = cover_elems[idx + 1]["full_text"] if idx + 1 < len(cover_elems) else ""
        return bool(COVER_NRP_NIP_RE.match(nxt))
    if _cover_is_caps(elem):
        return False
    size = _cover_dominant_size(elem)
    return not (size is not None and abs(size - COVER_RULES["title_size"]) <= 0.1)


def _check_cover_line(elem, expected_size, require_all_caps, findings, checks, chapter):
    """1 baris sampul = sampai 3 unit cek: ukuran font, bold, (opsional) ALL CAPS."""
    loc = f"Paragraf #{elem['index']} ('{text_preview(elem['full_text'])}')"

    size = _cover_dominant_size(elem)
    if size is not None:
        checks[chapter]["checked"] += 1
        if abs(size - expected_size) > 0.1:
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"Ukuran font {size}pt pada halaman sampul tidak "
                        f"sesuai, seharusnya {expected_size}pt."
                    ),
                }
            )

    checks[chapter]["checked"] += 1
    if not _cover_fully_bold(elem):
        checks[chapter]["failed"] += 1
        findings.append(
            {
                "category": "Format",
                "chapter": chapter,
                "location": loc,
                "description": (
                    "Teks pada halaman sampul seharusnya dicetak tebal "
                    "(bold) seluruhnya."
                ),
            }
        )

    if require_all_caps:
        checks[chapter]["checked"] += 1
        if not _cover_is_caps(elem):
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        "Teks pada halaman sampul seharusnya ditulis huruf "
                        "kapital semua."
                    ),
                }
            )


def _check_cover_name_titlecase(elem, findings, checks, chapter, label: str):
    """Baris nama (Nama Mahasiswa / Nama Pembimbing) wajib Title Case, bukan NRP/NIP."""
    checks[chapter]["checked"] += 1
    if not elem["is_titlecase"]:
        checks[chapter]["failed"] += 1
        findings.append(
            {
                "category": "Format",
                "chapter": chapter,
                "location": (
                    f"Paragraf #{elem['index']} "
                    f"('{text_preview(elem['full_text'])}')"
                ),
                "description": (
                    f"{label} pada halaman sampul seharusnya ditulis huruf "
                    "kapital di awal tiap kata (Title Case), tidak disingkat."
                ),
            }
        )


def compare_cover_page(paragraphs: list[dict], background_hex: str | None = None):
    """
    Cek FORMAT VISUAL halaman sampul (ptype "cover") sesuai lembar
    aturan cover resmi -- lihat blok komentar di atas fungsi ini untuk
    daftar lengkap aturan & alasan pendekatan state-machine yang
    dipakai. background_hex: hasil _get_document_background_color_hex()
    di extractor.py (docx) -- None kalau tidak terdeteksi / jalur PDF,
    dan None DIPERLAKUKAN sebagai "tidak bisa dipastikan", bukan salah.
    """
    findings = []
    checks = defaultdict(_new_counter)

    cover_elems = [e for e in paragraphs if e["type"] == "cover" and e["full_text"]]
    if not cover_elems:
        return findings, dict(checks)

    # DIUBAH (2026-09-17, atas permintaan user): SEBELUMNYA seluruh
    # kelompok ptype "cover" ikut diperiksa lengkap dengan aturan
    # sampul ini, dan begitu satu blok sampul selesai ("done"), state
    # akan RESTART kalau ketemu label "PROYEK AKHIR" lagi (asumsi
    # halaman judul dalam = salinan kedua sampul luar, sesuai bunyi
    # pedoman). User secara eksplisit minta pengecekan aturan sampul
    # ini HANYA berlaku untuk halaman sampul PERTAMA (yang berwarna
    # biru & kuning) -- halaman sesudahnya (halaman judul dalam,
    # lembar pengesahan, pernyataan orisinalitas, dst) TIDAK PERLU
    # diperiksa pakai aturan ini sama sekali. Fix: batasi cover_elems
    # ke HALAMAN FISIK PERTAMA saja (page_index sama dengan elemen
    # cover pertama) kalau field itu ada (jalur PDF) -- jalur docx
    # tidak punya field ini, jadi tetap diproses semua seperti
    # sebelumnya TAPI state "done" sekarang TERMINAL (lihat di bawah),
    # bukan restart, supaya blok kedua & seterusnya tidak ikut kena
    # aturan ini juga.
    if "page_index" in cover_elems[0]:
        first_page = cover_elems[0]["page_index"]
        cover_elems = [e for e in cover_elems if e.get("page_index") == first_page]

    chapter = cover_elems[0]["chapter"]

    # --- warna latar belakang sampul ---
    expected_hex = COVER_RULES["background_hex"]
    if background_hex is not None:
        checks[chapter]["checked"] += 1
        if not _hex_color_close_enough(background_hex, expected_hex):
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": "Halaman sampul",
                    "description": (
                        f"Warna latar belakang halaman sampul #{background_hex} "
                        f"tidak sesuai, seharusnya #{expected_hex} "
                        "(RGB 0, 47, 167)."
                    ),
                }
            )
    # background_hex None: SENGAJA tidak dilaporkan sama sekali -- lihat
    # docstring _get_document_background_color_hex() di extractor.py.

    # --- state machine urutan blok sampul ---
    # State "done" MENANDAI satu blok sampul resmi (label "PROYEK AKHIR"
    # s.d. baris tahun) sudah selesai diperiksa, dan SEKARANG TERMINAL
    # (diubah 2026-09-17, atas permintaan user) -- sebelumnya "done"
    # akan restart lagi ke "before_label" begitu ketemu label "PROYEK
    # AKHIR" kedua (dengan alasan pedoman PENS menyebut "Halaman judul
    # SAMA DENGAN halaman kulit luar", jadi keduanya dianggap wajib
    # tunduk ke aturan yang sama). User secara eksplisit minta
    # sebaliknya: cukup periksa halaman sampul PERTAMA (yang berwarna
    # biru & kuning) saja -- halaman judul dalam, lembar pengesahan,
    # pernyataan orisinalitas, dst TIDAK PERLU diperiksa pakai aturan
    # sampul ini sama sekali. cover_elems di atas sudah dibatasi ke
    # halaman fisik pertama (jalur PDF, via page_index); "done" dibuat
    # terminal di sini sebagai lapis kedua yang berlaku SAMA untuk
    # jalur docx (tidak punya page_index) -- begitu blok sampul
    # pertama selesai, baris apa pun sesudahnya (kalaupun ptype-nya
    # masih "cover") dilewati tanpa dicek.
    state = "before_label"
    # Sampul punya baris "NRP ..."/"NIP ..."? Dipakai _cover_is_name_line()
    # buat memisahkan judul dari nama TANPA bergantung huruf kapital.
    has_nrp_line = any(COVER_NRP_NIP_RE.match(e["full_text"]) for e in cover_elems)
    for idx, elem in enumerate(cover_elems):
        text = elem["full_text"]
        upper = text.upper()

        if state == "done":
            # TERMINAL (lihat catatan perubahan 2026-09-17 di atas) --
            # begitu satu blok sampul resmi selesai diperiksa, SEMUA
            # baris cover sesudahnya (kalaupun masih ada, mis. jalur
            # docx yang tidak difilter per-halaman di atas) dilewati
            # tanpa dicek, TIDAK PERNAH restart lagi.
            continue

        if state == "before_label":
            if COVER_PROYEK_AKHIR_LABEL_RE.match(upper):
                _check_cover_line(
                    elem, COVER_RULES["label_size"], True, findings, checks, chapter
                )
                state = "title_block"
            continue

        if state == "title_block":
            # Judul boleh terdiri lebih dari satu baris (word-wrap
            # otomatis atau Enter manual per baris) -- selama barisnya
            # ALL CAPS, tetap diperlakukan sebagai lanjutan judul
            # (title_size). Baris pertama yang BUKAN ALL CAPS menandai
            # judul sudah selesai dan ini sudah masuk baris Nama
            # Mahasiswa -- lihat catatan "CATATAN JUDUL MULTI-BARIS" di
            # atas COVER_RULES untuk alasan lengkap heuristik ini.
            # DIUBAH (2026-09-21): batas judul/nama TIDAK LAGI ditentukan
            # oleh "baris ini ALL CAPS atau bukan" -- judul huruf biasa
            # (mis. "Contextual Analysis untuk Opinion Mining ...")
            # salah dikira nama mahasiswa dan diperiksa 16pt padahal
            # 18pt sudah benar. Sekarang lewat _cover_is_name_line().
            # Judul yang memang tidak kapital semua tetap dilaporkan --
            # tapi lewat cek "huruf kapital semua" yang benar, bukan
            # lewat cek ukuran font yang salah sasaran.
            if not _cover_is_name_line(cover_elems, idx, has_nrp_line):
                _check_cover_line(
                    elem, COVER_RULES["title_size"], True, findings, checks, chapter
                )
                continue
            state = "nama_mahasiswa_block"
            # sengaja TIDAK "continue" -- baris ini sendiri sudah baris
            # Nama Mahasiswa pertama, harus diproses lewat cabang di
            # bawah pada iterasi yang SAMA, bukan dilewati.

        if state == "nama_mahasiswa_block":
            if COVER_DOSEN_PEMBIMBING_RE.match(upper):
                _check_cover_line(
                    elem, COVER_RULES["detail_size"], True, findings, checks, chapter
                )
                state = "pembimbing_block"
                continue
            _check_cover_line(
                elem, COVER_RULES["detail_size"], False, findings, checks, chapter
            )
            if not COVER_NRP_NIP_RE.match(text):
                _check_cover_name_titlecase(
                    elem, findings, checks, chapter, "Nama mahasiswa"
                )
            continue

        if state == "pembimbing_block":
            if COVER_PROGRAM_STUDI_RE.match(upper):
                _check_cover_line(
                    elem, COVER_RULES["detail_size"], True, findings, checks, chapter
                )
                state = "prodi_block"
                continue
            _check_cover_line(
                elem, COVER_RULES["detail_size"], False, findings, checks, chapter
            )
            if not COVER_NRP_NIP_RE.match(text):
                _check_cover_name_titlecase(
                    elem, findings, checks, chapter, "Nama dosen pembimbing"
                )
            continue

        if state == "prodi_block":
            if COVER_POLITEKNIK_RE.match(upper):
                _check_cover_line(
                    elem, COVER_RULES["detail_size"], True, findings, checks, chapter
                )
                state = "tahun_block"
                continue
            # Baris [NAMA PRODI] / [NAMA DEPARTEMEN] -- nama institusi,
            # ALL CAPS sama seperti "PROGRAM STUDI SARJANA TERAPAN".
            _check_cover_line(
                elem, COVER_RULES["detail_size"], True, findings, checks, chapter
            )
            continue

        if state == "tahun_block":
            # Baris tahun (mis. "2025") -- angka, cuma cek ukuran+bold.
            # Cuma SATU baris (bukan loop tanpa akhir seperti sebelumnya):
            # begitu baris tahun ini diproses, blok sampul resmi sudah
            # lengkap -- pindah ke state terminal "done" supaya halaman
            # template lain sesudahnya tidak ikut kena aturan cover.
            _check_cover_line(
                elem, COVER_RULES["detail_size"], False, findings, checks, chapter
            )
            state = "done"
            continue

    return findings, dict(checks)


def compare_all(
    extracted_params: dict,
    section_pages: dict[int, int] | None = None,
    source: str = "docx",
):
    """
    source: "docx" | "pdf" -- diteruskan apa adanya ke compare_sections()
    supaya toleransi margin PDF (lihat docstring compare_sections) cuma
    nyala untuk dokumen yang memang lewat jalur extract_all_pdf() di
    main.py. Default "docx" supaya pemanggil lama yang belum di-update
    untuk kasih argumen ini perilakunya TIDAK BERUBAH (tetap toleransi
    ketat seperti sebelum perubahan ini ada).
    """
    p_findings, p_checks = compare_paragraphs(
        extracted_params["paragraphs"], source=source
    )
    s_findings, s_checks = compare_sections(
        extracted_params["sections"],
        extracted_params["mirror_margins"],
        source=source,
    )
    n_findings, n_checks = compare_page_numbering(
        extracted_params["paragraphs"], extracted_params["sections"], section_pages
    )
    c_findings, c_checks = compare_caption_numbering(extracted_params["paragraphs"])
    ab_findings, ab_checks = compare_abstract(extracted_params["paragraphs"])
    o_findings, o_checks = compare_visual_objects(
        extracted_params["paragraphs"],
        extracted_params.get("tables", []),
        source=source,
    )
    q_findings, q_checks = compare_quotations(extracted_params["paragraphs"])
    # extracted_params.get(...) (bukan [...]) -- "cover_background_hex"
    # cuma diisi jalur docx (lihat extractor.py); jalur PDF belum
    # mengimplementasikan deteksi warna latar halaman, jadi defaultnya
    # None di sini, dan None DIPERLAKUKAN "tidak bisa dipastikan" oleh
    # compare_cover_page(), bukan "warna salah".
    cv_findings, cv_checks = compare_cover_page(
        extracted_params["paragraphs"],
        background_hex=extracted_params.get("cover_background_hex"),
    )

    findings = deduplicate_findings(
        p_findings
        + s_findings
        + n_findings
        + c_findings
        + ab_findings
        + o_findings
        + q_findings
        + cv_findings
    )

    checks = defaultdict(_new_counter)
    for group_checks in (
        p_checks,
        s_checks,
        n_checks,
        c_checks,
        ab_checks,
        o_checks,
        q_checks,
        cv_checks,
    ):
        for chapter, c in group_checks.items():
            checks[chapter]["checked"] += c["checked"]

    for chapter in checks:
        checks[chapter]["failed"] = 0
    for f in findings:
        checks[f["chapter"]]["failed"] += 1

    return findings, dict(checks)
