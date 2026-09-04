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

# EQUATION_NUMBER_IN_TEXT_RE: dipakai HANYA oleh blok cek persamaan di
# compare_visual_objects() yang sekarang dinonaktifkan (lihat komentar
# "DINONAKTIFKAN" di bawah) -- import ikut dikomentari supaya tidak ada
# import menggantung tak terpakai, konsisten dengan pola
# foreign_term_checker di atas. Un-comment bareng blok persamaan kalau
# nanti diaktifkan lagi.
# from aggregator.pipelines.visual.extractor import EQUATION_NUMBER_IN_TEXT_RE


def text_preview(text: str, length: int = 40) -> str:
    text = text.strip()
    return text[:length] + ("..." if len(text) > length else "")


EXPECTED_SIZE_BY_TYPE = {
    "chapter_title": RULES["chapter_title_size"],
    "subchapter_title": RULES["subchapter_title_size"],
    # panduan tidak menetapkan ukuran font terpisah untuk level "X.Y.Z"
    # (bernomor maupun heading tanpa nomor), jadi disamakan dengan
    # level "X.Y" (subchapter_title_size)
    "subsubchapter_title": RULES["subchapter_title_size"],
    "caption": RULES["caption_font_size"],
    "body": RULES["body_font_size"],
}


def _new_counter():
    return {"checked": 0, "failed": 0}


def compare_paragraphs(paragraphs: list[dict]):
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
            if dominant_font != RULES["font_name"]:
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
            if r["font_size_pt"] is not None and r["text"].strip() != ""
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
        if ptype in ("chapter_title", "subchapter_title", "subsubchapter_title"):
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
        if ptype == "chapter_title":
            checks[chapter]["checked"] += 1
            if elem["alignment"] != "center":
                checks[chapter]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": chapter,
                        "location": loc,
                        "description": f"Judul tipe '{ptype}' seharusnya ditulis rata tengah.",
                    }
                )

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
        # DINONAKTIFKAN (atas permintaan user, 2026-08-23): dikomentari,
        # bukan dihapus -- pola yang sama seperti cek-cek lain yang
        # dimatikan di modul ini.
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

    # (label, cm_key, twips_key, expected_twips, expected_cm_untuk_pesan)
    margin_defs = [
        (
            "kiri",
            "left_cm",
            "left_twips",
            RULES_TWIPS["margin_left"],
            RULES["margin_left_cm"],
        ),
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
        (
            "kanan",
            "right_cm",
            "right_twips",
            RULES_TWIPS["margin_other"],
            RULES["margin_other_cm"],
        ),
    ]

    for label, cm_key, twips_key, expected_twips, expected_cm in margin_defs:
        checks[group]["checked"] += 1  # 1x per sisi margin, untuk seluruh dokumen

        effective_tol = tol_twips
        if source == "pdf":
            effective_tol += PDF_EXTRA_TOLERANCE_TWIPS[label]

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
    expected_mirror = RULES["mirror_margins"]

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
CAPTION_PARSE_RE = re.compile(r"^(Gambar|Tabel)\s+(\d+)\.\s*(\d+)")
BAB_NUMBER_RE = re.compile(r"^BAB\s+(\d+)$")

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
    """
    findings = []
    checks = defaultdict(_new_counter)
    next_seq = {}  # {(chapter, jenis): nomor urut yang diharapkan berikutnya}

    for elem in paragraphs:
        if elem["type"] != "caption":
            continue

        chapter = elem["chapter"]
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
        cap_bab = int(cap_bab_str)
        cap_seq = int(cap_seq_str)
        bab_match = BAB_NUMBER_RE.match(chapter)

        # --- cek 1: angka bab di nomor caption harus cocok bab aktual ---
        checks[chapter]["checked"] += 1
        if bab_match is None or cap_bab != int(bab_match.group(1)):
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        f"Nomor {jenis.lower()} '{cap_bab_str}.{cap_seq_str}' "
                        f"memakai angka bab {cap_bab_str}, padahal paragraf ini "
                        f"berada di {chapter}."
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


def compare_visual_objects(paragraphs: list[dict], tables: list[dict]):
    """
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
      - Persamaan: masuk 10mm dari batas kiri, 1 baris kosong di atas &
        bawah, dan wajib punya nomor identitas mis. "(2.1)" yang
        dicetak tebal.
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
        (beda dari persamaan, yang tetap dicek). Alasan: mahasiswa yang
        pakai "Spacing Before/After Paragraph" (bukan tombol Enter/
        baris kosong beneran) untuk memberi jarak visual di sekitar
        gambar/tabel akan salah ditandai kalau ini dipaksakan, padahal
        secara visual hasilnya sama-sama benar.

    KETERBATASAN YANG DISADARI (bukan dicek di sini, butuh verifikasi
    manual):
      - Posisi caption Gambar HARUS di bawah gambar per contoh visual
        panduan, tapi redaksi panduan untuk Tabel *literally* juga
        bilang "dituliskan pada 1 baris DIBAWAH tabel" -- padahal
        contoh visualnya (Tabel 3.1) caption-nya di ATAS tabel.
        Kemungkinan salah ketik redaksi (copy-paste dari bagian
        Gambar). Karena ambigu, cek posisi caption Tabel di sini
        SENGAJA menerima caption di atas MAUPUN di bawah tabel.
      - Nomor persamaan "rata kanan": TIDAK dicek eksplisit. Nomor
        biasanya satu paragraf yang sama dengan rumusnya (dipisah tab
        stop rata-kanan), bukan atribut rata-paragraf (w:jc) yang bisa
        dibaca langsung -- mengecek w:jc paragraf persamaan tidak akan
        merepresentasikan posisi nomornya dengan benar, jadi sengaja
        tidak dipaksakan supaya tidak salah-tandai dokumen yang benar.
      - Tabel tidak boleh lebih dari 1 halaman: butuh info halaman
        fisik hasil render PDF per baris tabel, di luar jangkauan
        modul ini (python-docx tidak tahu pemotongan halaman).
    """
    findings = []
    checks = defaultdict(_new_counter)
    paragraphs_by_index = {e["index"]: e for e in paragraphs}

    # --- GAMBAR ---
    for elem in paragraphs:
        if elem["type"] != "image":
            continue
        chapter = elem["chapter"]
        loc = f"Paragraf #{elem['index']} (gambar)"

        checks[chapter]["checked"] += 1
        if elem["alignment"] != "center":
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": "Gambar seharusnya diletakkan rata tengah.",
                }
            )

        # --- caption ("Gambar N.N ...") wajib ada, menempel di salah
        #     satu sisi (atas/bawah) -- sebelumnya TIDAK PERNAH dicek
        #     sama sekali (cuma format PENOMORAN caption yang dicek
        #     KALAU ada, lihat compare_caption_numbering()), padahal
        #     panduan mewajibkan tiap gambar punya nomor identitas &
        #     keterangan. Simetris dengan cek yang sama untuk TABEL di
        #     bawah -- pakai `elem["index"] - 1` / `+ 1` (bukan
        #     elem["index"] itu sendiri) sebagai titik mulai pencarian
        #     supaya paragraf gambar itu sendiri tidak ikut kehitung
        #     sebagai "tetangganya sendiri" (beda dari tabel, yang
        #     memang bukan entri di paragraphs_by_index sama sekali).
        nearby_elems = _nearby_paragraphs(
            paragraphs_by_index, elem["index"] - 1, "prev"
        ) + _nearby_paragraphs(paragraphs_by_index, elem["index"] + 1, "next")
        has_caption_adjacent = any(
            e["type"] == "caption" and e["full_text"].startswith("Gambar")
            for e in nearby_elems
        )
        checks[chapter]["checked"] += 1
        if not has_caption_adjacent:
            checks[chapter]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": chapter,
                    "location": loc,
                    "description": (
                        "Gambar tidak punya nomor identitas & keterangan "
                        "('Gambar N.N ...') tepat di atas/bawahnya."
                    ),
                }
            )

    # --- PERSAMAAN ---
    # DINONAKTIFKAN (atas permintaan user, 2026-08-23): keempat cek
    # persamaan (indentasi 10mm, 1 baris kosong di atas/bawah, wajib
    # punya nomor identitas "(2.1)", dan bold-nya nomor tsb) sengaja
    # dimatikan -- user eksplisit bilang ini tidak perlu dicek. Kode
    # dipertahankan (dikomentari, bukan dihapus) mengikuti pola yang
    # sama dengan foreign_term_checker/bibliography_checker yang juga
    # pernah dinonaktifkan sebelumnya -- supaya gampang diaktifkan lagi
    # kalau suatu saat dibutuhkan, tinggal un-comment blok di bawah.
    #
    # left_indent_target_cm = 1.0  # 10mm
    # tol_cm = 0.2
    # for elem in paragraphs:
    #     if elem["type"] != "equation":
    #         continue
    #     chapter = elem["chapter"]
    #     loc = f"Paragraf #{elem['index']} (persamaan)"
    #
    #     checks[chapter]["checked"] += 1
    #     li = elem["left_indent_cm"]
    #     if li is None or abs(li - left_indent_target_cm) > tol_cm:
    #         checks[chapter]["failed"] += 1
    #         detected = "tidak diset" if li is None else f"{li:.2f}cm"
    #         findings.append(
    #             {
    #                 "category": "Format",
    #                 "chapter": chapter,
    #                 "location": loc,
    #                 "description": (
    #                     f"Persamaan seharusnya masuk 10mm dari batas kiri "
    #                     f"(terdeteksi: {detected})."
    #                 ),
    #             }
    #         )
    #
    #     checks[chapter]["checked"] += 1
    #     if not (elem["blank_before"] and elem["blank_after"]):
    #         checks[chapter]["failed"] += 1
    #         findings.append(
    #             {
    #                 "category": "Format",
    #                 "chapter": chapter,
    #                 "location": loc,
    #                 "description": (
    #                     "Persamaan seharusnya menyisakan 1 baris kosong di "
    #                     "atas dan di bawahnya."
    #                 ),
    #             }
    #         )
    #
    #     checks[chapter]["checked"] += 1
    #     m = EQUATION_NUMBER_IN_TEXT_RE.search(elem["full_text"])
    #     if not m:
    #         checks[chapter]["failed"] += 1
    #         findings.append(
    #             {
    #                 "category": "Format",
    #                 "chapter": chapter,
    #                 "location": loc,
    #                 "description": (
    #                     "Persamaan tidak memiliki nomor identitas, " "mis. '(2.1)'."
    #                 ),
    #             }
    #         )
    #     else:
    #         checks[chapter]["checked"] += 1
    #         runs_with_text = [r for r in elem["runs"] if r["text"].strip()]
    #         if runs_with_text and any(r["bold"] is not True for r in runs_with_text):
    #             checks[chapter]["failed"] += 1
    #             findings.append(
    #                 {
    #                     "category": "Format",
    #                     "chapter": chapter,
    #                     "location": loc,
    #                     "description": (
    #                         "Nomor identitas persamaan seharusnya dicetak "
    #                         "tebal (bold)."
    #                     ),
    #                 }
    #             )

    # --- TABEL ---
    # Rata tengah TIDAK dicek di sini -- itu aturan untuk CAPTION-nya
    # (ptype "caption", sudah dicek di compare_paragraphs), bukan untuk
    # tabelnya sendiri. Baris kosong di atas/bawah juga tidak dicek,
    # lihat alasan di docstring atas.
    for tbl in tables:
        chapter = _chapter_of_table(paragraphs_by_index, tbl)
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
    Cek 3 aturan penomoran halaman:
      0. Section SEBELUM BAB 1 (cover, kata pengantar, daftar isi, dst)
         wajib format Romawi kecil.
      1. Section tempat "BAB 1" dimulai wajib format Arab (pgNumType
         fmt="decimal") DAN restart dari 1 -- supaya BAB 1 memang
         halaman "1", bukan lanjutan romawi pra-bab.
      2. Semua section dari BAB 1 dan seterusnya wajib punya field
         nomor halaman (PAGE) di footer, rata tengah.
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

    # --- cek -1: pastikan SECTION BREAK memang ada sebelum BAB 1 ---
    # MASALAH YANG DIPERBAIKI: kalau dokumen tidak punya Section Break
    # sama sekali sebelum BAB 1, front matter (cover, kata pengantar,
    # daftar isi, dst) ikut menyatu di section YANG SAMA dengan BAB 1.
    # Efeknya pre_bab1_sections (dicek di bawah) jadi KOSONG -- bukan
    # karena memang tidak ada front matter, tapi karena front matter-nya
    # tidak punya section index sendiri yang lebih kecil. Kode LAMA
    # menganggap list kosong = "tidak ada yang perlu dicek" (diam-diam
    # lolos), padahal list kosong itu AMBIGU: bisa juga berarti "section
    # break-nya lupa ditaruh". Cek ini membedakan dua kemungkinan itu
    # secara eksplisit dengan bukti yang pasti: bandingkan posisi
    # PARAGRAF (bukan posisi section) -- kalau ada paragraf SEBELUM
    # paragraf "BAB 1" yang section-nya SAMA dengan section BAB 1,
    # berarti PASTI tidak ada section break di antara mereka (soalnya
    # kalau section break beneran ada, paragraf sebelum BAB 1 pasti
    # section index-nya lebih kecil, tidak mungkin sama).
    bab1_para_index = next(
        (
            elem["index"]
            for elem in paragraphs
            if elem["type"] == "chapter_title"
            and BAB1_TITLE_RE.match(elem["full_text"])
        ),
        None,
    )
    front_matter_shares_section = bab1_para_index is not None and any(
        elem["index"] < bab1_para_index and elem["section"] == bab1_section_idx
        for elem in paragraphs
    )
    if front_matter_shares_section:
        checks[group]["checked"] += 1
        checks[group]["failed"] += 1
        findings.append(
            {
                "category": "Format",
                "chapter": group,
                "location": "Sebelum BAB 1",
                "description": (
                    "Tidak ditemukan Section Break sebelum BAB 1 -- bagian "
                    "sebelum BAB 1 (cover, kata pengantar, daftar isi, dst) "
                    "tergabung dalam SATU section yang sama dengan BAB 1. "
                    "Format nomor halaman (romawi/Arab) itu properti PER-"
                    "SECTION di Word, jadi TIDAK MUNGKIN bagian sebelum BAB "
                    "1 romawi sementara BAB 1 dst Arab-restart-1 tanpa "
                    "Section Break (Next Page) eksplisit di antara "
                    "keduanya. Tambahkan Section Break sebelum halaman "
                    "BAB 1."
                ),
            }
        )
        # Cek 0 di bawah SENGAJA tetap dilanjutkan (bukan return di sini)
        # -- kalaupun front matter menyatu section, pgnum_fmt SECTION itu
        # sendiri (dibagi bersama BAB 1) tetap bisa saja diperiksa & bisa
        # jadi kebetulan "lowerRoman" (walau ganjil), jadi tidak ada
        # ruginya cek 0 tetap jalan sebagai lapis validasi tambahan --
        # cuma finding di atas ini yang PASTI muncul duluan untuk kasus
        # ini, memberi tahu akar masalahnya secara eksplisit.

    # --- cek 0: semua section SEBELUM BAB 1 wajib romawi kecil ---
    # Dicek pakai pgnum_fmt_effective (hasil warisan), bukan pgnum_fmt
    # mentah -- kalau cuma section 0 yang eksplisit set "lowerRoman" dan
    # section 1, 2 dst di Pra-Bab tidak override apa-apa, mereka tetap
    # harus dianggap romawi (warisan), bukan ke-skip / dianggap salah.
    pre_bab1_sections = [s for s in sections if s["index"] < bab1_section_idx]
    if pre_bab1_sections:
        checks[group]["checked"] += 1
        offending = [
            sec["index"]
            for sec in pre_bab1_sections
            if sec["pgnum_fmt"] != "lowerRoman"
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
                        "Penomoran halaman di section tsb (sebelum BAB 1) "
                        "harus format angka Romawi kecil (i, ii, iii, ...)."
                    ),
                }
            )

    # --- cek 1: format Arab + restart dari 1 ---
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
    checks[group]["checked"] += 1
    offending = [
        sec["index"]
        for sec in sections
        if sec["index"] >= bab1_section_idx
        # and not (sec["footer_has_page_field"] and sec["footer_alignment"] == "center")
        and not sec["footer_has_page_field"]
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
    p_findings, p_checks = compare_paragraphs(extracted_params["paragraphs"])
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
        extracted_params["paragraphs"], extracted_params.get("tables", [])
    )
    q_findings, q_checks = compare_quotations(extracted_params["paragraphs"])

    findings = deduplicate_findings(
        p_findings
        + s_findings
        + n_findings
        + c_findings
        + ab_findings
        + o_findings
        + q_findings
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
    ):
        for chapter, c in group_checks.items():
            checks[chapter]["checked"] += c["checked"]

    for chapter in checks:
        checks[chapter]["failed"] = 0
    for f in findings:
        checks[f["chapter"]]["failed"] += 1

    return findings, dict(checks)
