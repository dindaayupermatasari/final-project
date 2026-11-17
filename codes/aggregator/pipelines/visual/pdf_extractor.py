"""
pdf_extractor.py
=================
TAHAP 2 (Extract Parameter) versi PDF-native.

KENAPA MODUL INI ADA
---------------------
extractor.py (versi lama) membaca parameter LANGSUNG dari XML docx
(w:pgMar, w:jc, style, run.bold, dst). Itu paling akurat, TAPI cuma
bisa jalan kalau file sumbernya .docx. Mayoritas mahasiswa upload PDF
(bukan docx), dan pembimbing eksplisit minta PDF diutamakan sebagai
sumber pengecekan -- jadi modul ini dibuat sebagai PENGGANTI/PIPELINE
UTAMA, bukan cuma pelengkap check_page_layout() seperti render PDF yang
sudah ada sebelumnya.

PRINSIP DESAIN (PALING PENTING)
---------------------------------
extract_all() di sini mengembalikan bentuk yang SAMA PERSIS dengan
extract_all() versi docx: {"paragraphs", "sections", "tables",
"mirror_margins"}, dan tiap paragraf tetap punya "runs": list of
{"text","font_name","font_size_pt","bold","italic"}.

Konsekuensinya: comparator.py, bibliography_checker.py,
foreign_term_checker.py, quotation_checker.py TIDAK PERLU diubah SAMA
SEKALI -- semua modul itu sudah generic, cuma konsumsi dict/list biasa,
tidak pernah pegang objek docx langsung. Yang beda HANYA cara isi dict
itu didapat.

Fungsi classify_paragraph(), is_titlecase(), dan semua REGEX heading/
caption/dst DIIMPOR ULANG dari extractor.py (bukan diduplikasi) --
logic itu murni berbasis STRING, sama sekali tidak menyentuh objek
docx, jadi 100% valid dipakai ulang di sini.

KETERBATASAN DIBANDING VERSI DOCX (baca sebelum pakai sebagai sumber
utama produksi -- ini bukan pengganti 1:1, tapi pendekatan terbaik yang
bisa didapat dari PDF)
----------------------------------------------------------------------
1. Heading via w:numPr/ilvl (Jalur 2 classify_paragraph, "bab yang
   nomornya digenerate murni oleh multilevel-list Word tanpa teks
   'BAB N' literal") TIDAK bisa dideteksi dari PDF -- info itu sama
   sekali tidak ada lagi begitu dokumen dirender jadi PDF. Dampaknya
   kecil karena mayoritas dokumen menulis "BAB N" literal.
2. is_unnumbered_heading() versi PDF (_is_unnumbered_heading_runs di
   bawah) pakai `runs` hasil ekstraksi sendiri (bukan para.runs docx) --
   secara prinsip setara (bold penuh + pendek + tanpa tanda baca akhir),
   bukan cuma didekati.
3. Rumus matematika OMML (m:oMath) tidak terdeteksi presisi -- OMML
   sudah "diratakan" jadi glyph biasa saat render ke PDF. Proxy yang
   dipakai SAMA seperti _is_math_run_text() di extractor.py: karakter
   Unicode matematis & nama font matematis (Cambria Math) -- ini tetap
   valid karena berbasis teks/font run, bukan struktur XML.
4. alignment, indentasi kiri/kanan, line_spacing, margin section semua
   DIDEKATI dari posisi (bbox) tiap span/blok teks terhadap tepi
   halaman -- BUKAN dibaca langsung dari atribut seperti di docx.
   Toleransi tiap fungsi dijelaskan di docstring masing-masing.
5. Satu paragraf PDF didekati sebagai SATU "block" hasil deteksi
   PyMuPDF (page.get_text("dict")["blocks"]). Ini heuristik bawaan
   PyMuPDF (dikelompokkan dari jarak antar baris) -- pada dokumen rapi
   umumnya sudah cukup akurat, tapi BUKAN jaminan 1:1 sama seperti
   docx.paragraphs (yang eksplisit dari w:p). Kasus tepi yang mungkin
   meleset: dua paragraf pendek berurutan tanpa jarak antar-baris
   ekstra bisa ke-gabung jadi satu "paragraf".
6. Section docx (tiap section bisa beda margin/orientasi) disederhanakan
   jadi SATU section untuk seluruh isi dokumen (margin diambil median
   dari semua halaman body, BUKAN per-section) -- dokumen dengan margin
   berbeda antar section (mis. lampiran landscape) tidak dibedakan.
"""

import re
import statistics
from collections import defaultdict

import fitz  # PyMuPDF

from aggregator.pipelines.visual.extractor import (
    CHAPTER_TITLE_RE,
    CHAPTER_TITLE_ONLY_RE,
    LAMPIRAN_TITLE_RE,
    LAMPIRAN_TITLE_ONLY_RE,
    SPECIAL_FRONT_PAGE_MARKERS,
    NORMAL_FRONT_MATTER_MARKERS,
    BLANK_PAGE_NOTICE_RE,
    ABSTRACT_KEYWORD_RE,
    SOURCE_CITATION_RE,
    EQUATION_NUMBER_BARE_RE,
    EQUATION_NUMBER_IN_TEXT_RE,
    EQUATION_NUMBER_RE,
    normalize_equation_text,
    CAPTION_RE,
    is_caption_title,
    FRONT_HEADINGS,
    MATH_FONT_NAMES,
    _is_math_run_text,
    _MATH_ALPHANUMERIC_RE,
    classify_paragraph,
    is_titlecase,
)

PT_TO_CM = 2.54 / 72.0

# Blok di zona ini (dari tepi atas/bawah halaman) dianggap header/footer
# -- BUKAN badan dokumen. 70pt konsisten dengan FOOTER_ZONE_PT yang
# sudah diverifikasi empiris di page_layout_checker.py.
HEADER_FOOTER_ZONE_PT = 70

# Zona "dekat tepi" yang LEBIH LEBAR dari HEADER_FOOTER_ZONE_PT, dipakai
# KHUSUS untuk cek nomor-halaman-badan-yang-nyasar di bawah (lihat
# catatan bug 2026-09-17 di pemanggilnya) -- supaya baris pendek berisi
# angka di TENGAH halaman (mis. baris tahun "2025" di halaman sampul)
# tidak ikut terbuang, tapi nomor halaman yang cuma nyasar sedikit dari
# 70pt (margin bawah kecil) tetap tertangkap.
NEAR_EDGE_ZONE_PT = 85

# Blok pendek (nomor halaman, dsb) di zona footer -- pola longgar,
# angka biasa atau angka romawi saja.
_PAGE_NUM_ONLY_RE = re.compile(r"^[ivxlcdm\d]+$", re.IGNORECASE)

# Ciri baris entri Daftar Isi/Daftar Gambar/Daftar Tabel: diakhiri
# dot-leader (>=3 titik berurutan) + opsional nomor halaman di ujung.
# Ini SINYAL YANG TIDAK ADA di docx (dot-leader di sana cuma properti
# tab-stop Word, bukan karakter sungguhan di w:t) -- tapi begitu
# dokumen dirender jadi PDF, dot-leader itu IKUT TERBACA sebagai teks
# biasa. Sinyal ini dipakai buat membedakan entri TOC (mis. "BAB 1
# PENDAHULUAN ...................... 5") dari heading ASLI (mis. "BAB
# 1 PENDAHULUAN" tanpa embel-embel dot-leader), yang teksnya bisa
# identik persis dan SAMA-SAMA lolos classify_paragraph() sebagai
# "chapter_title" -- lihat penjelasan lengkap di blok penanganan
# in_toc/in_figure_table_list pada extract_paragraph_params_pdf().
_TOC_LEADER_RE = re.compile(r"\.{3,}\s*[ivxlcdm\d]*\s*$", re.IGNORECASE)

# BUG DIPERBAIKI (2026-09-17, laporan user -- baris institusi ALL CAPS
# di halaman sampul, mis. "PROGRAM STUDI SARJANA TERAPAN", salah
# ditandai "Nama mahasiswa seharusnya Title Case"): label sampul tetap
# "DOSEN PEMBIMBING" (dan "PROYEK AKHIR") WAJIB berdiri sendiri sebagai
# satu baris menurut template (lihat COVER_DOSEN_PEMBIMBING_RE /
# COVER_PROYEK_AKHIR_LABEL_RE di comparator.py, keduanya anchor exact
# match "^...$"). Kalau mahasiswa menulis baris nama pembimbing TEPAT
# di baris berikutnya TANPA baris kosong ekstra (wajar, memang tidak
# diwajibkan pedoman), jaraknya jadi cukup rapat sehingga PyMuPDF
# menggabung "DOSEN PEMBIMBING" dengan baris nama pembimbing di
# bawahnya jadi SATU block -- dibuktikan langsung ke PDF nyata (block
# berisi "DOSEN PEMBIMBING Ali Ridho Barakbah ,S.Kom., Ph.D." sebagai
# SATU teks gabungan). Akibatnya COVER_DOSEN_PEMBIMBING_RE (yang
# mensyaratkan baris PERSIS "DOSEN PEMBIMBING" tanpa embel apa pun)
# GAGAL match, compare_cover_page() gagal pindah state ke
# "pembimbing_block" dan MACET PERMANEN di "nama_mahasiswa_block" --
# SEMUA baris sampul sesudahnya (termasuk baris institusi ALL CAPS)
# ikut salah dicek pakai aturan "nama mahasiswa harus Title Case".
# Fix: paksa baris label sampul ini jadi grup PyMuPDF tersendiri, sama
# seperti perlakuan FRONT_HEADINGS/CAPTION_RE/SOURCE_CITATION_RE di
# bawah -- supaya baris nama pembimbing di bawahnya SELALU jadi block
# terpisah, apa pun jarak baris di dokumen aslinya.
_COVER_STANDALONE_LABEL_RE = re.compile(r"^(PROYEK\s+AKHIR|DOSEN\s+PEMBIMBING)$")


# --------------------------------------------------------------------
# Span-level helpers
# --------------------------------------------------------------------


_SUBSET_PREFIX_RE = re.compile(r"^[A-Z]{6}\+")
_FONT_SUFFIX_RE = re.compile(
    r"-?(BoldItalic|BoldItal|Bold|Italic|Oblique|Regular)$", re.IGNORECASE
)


def _normalize_font_name(name: str | None) -> str | None:
    """
    PDF (baik hasil export Word langsung maupun render LibreOffice)
    menyimpan nama font sebagai PostScript name internal, BUKAN nama
    family yang manusia baca -- mis. "Times New Roman" tersimpan
    sebagai "TimesNewRomanPSMT" (regular), "TimesNewRomanPS-BoldMT"
    (bold), "TimesNewRomanPS-ItalicMT" (italic). docx XML sebaliknya
    langsung simpan "Times New Roman" apa adanya.

    Tanpa normalisasi ini, RULES["font_name"] ("Times New Roman") di
    config.py TIDAK PERNAH match nama PostScript-nya -- comparator.py
    jadi salah tandai HAMPIR SEMUA paragraf sebagai "font tidak
    sesuai", padahal font-nya sudah benar. Ini bukan kasus tepi, tapi
    penyebab UTAMA skor jalur PDF jatuh drastis dibanding jalur docx.

    Strategi: lucuti awalan subset ("ABCDEF+..."), lucuti akhiran
    "MT"/style ("PS", "-Bold", "-Italic", "-BoldItalic", dst) satu per
    satu dari belakang, lalu pisahkan sisa nama camelCase jadi kata per
    kata ("TimesNewRoman" -> "Times New Roman"). Nama yang BUKAN pola
    PostScript umum (mis. "Calibri", "Arial" polos) tetap apa adanya --
    itu memang bukan Times New Roman, dan harus tetap ke-flag.
    """
    if not name:
        return name
    n = _SUBSET_PREFIX_RE.sub("", name)
    # Lucuti "MT" di akhir (penanda umum PostScript TrueType), lalu
    # lucuti akhiran gaya (Bold/Italic/dst, dengan/tanpa "-"), lalu
    # lucuti sisa "PS" penanda PostScript -- ulangi karena bisa
    # bertumpuk (mis. "PS-BoldMT" = PS + "-Bold" + "MT").
    for _ in range(3):
        n = re.sub(r"MT$", "", n)
        n = _FONT_SUFFIX_RE.sub("", n)
        n = re.sub(r"-?PS$", "", n)
    n = n.strip("-").strip()
    if n and " " not in n and re.search(r"[a-z][A-Z]", n):
        n = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", n)
    return n or name


def _span_bold(span: dict) -> bool:
    """
    Flag bit 2**4 di span["flags"] PyMuPDF = bold. Ditambah fallback ke
    nama font ("...-Bold...") karena sebagian PDF hasil LibreOffice
    tidak selalu set flag itu dengan benar untuk font tertentu.
    """
    if span.get("flags", 0) & (1 << 4):
        return True
    return "bold" in (span.get("font") or "").lower()


def _span_italic(span: dict) -> bool:
    """Simetris dengan _span_bold(), flag bit 2**1 = italic."""
    if span.get("flags", 0) & (1 << 1):
        return True
    fname = (span.get("font") or "").lower()
    return "italic" in fname or "oblique" in fname


def _span_is_math(span: dict) -> bool:
    """
    True kalau span ini rumus/simbol matematis -- font matematis
    (Cambria Math, dst) ATAU karakter Unicode matematis di teksnya.

    BUG DIPERBAIKI (2026-09-20, laporan user -- rumus yang di dokumen
    ASLINYA sudah benar berfont Cambria Math tetap ke-flag "Font
    'Cambria Math' tidak sesuai, seharusnya 'Times New Roman'", kadang
    plus ukurannya ikut disalahkan kalau beda dari body_font_size):
    sebelumnya fungsi ini membandingkan span.get("font") MENTAH --
    nama PostScript internal PDF, mis. "ABCDEF+CambriaMath" (ada
    awalan subset prefix 6-huruf), "CambriaMath-Italic", atau
    "CambriaMathItalic" (SAMA SEKALI TANPA SPASI) -- langsung terhadap
    MATH_FONT_NAMES = {"Cambria Math"} (pakai spasi). Ini TIDAK PERNAH
    match -- persis masalah yang sama yang sudah diperbaiki untuk font
    body biasa lewat _normalize_font_name() (lihat docstring-nya di
    atas), cuma normalisasi itu belum ikut dipakai di sini. Akibatnya
    SEMUA run rumus matematika (yang diketik sebagai teks biasa
    berfont Cambria Math, bukan OMML) gagal dikenali sebagai is_math
    di jalur PDF, lolos ke pengecekan font & ukuran body biasa seolah
    teks biasa -- padahal font ini WAJAR dipakai untuk notasi
    matematis dan tidak seharusnya dituntut Times New Roman 12pt.

    Fix: pakai nama font yang SUDAH dinormalisasi lewat proses yang
    SAMA seperti font body text (_normalize_font_name -- lucuti subset
    prefix, lucuti akhiran gaya, pisah camelCase jadi berspasi), lalu
    dibandingkan case-insensitive (konsisten dengan
    _font_matches_expected() di comparator.py yang juga
    case-insensitive) -- BUKAN exact-match string mentah.
    """
    normalized = (_normalize_font_name(span.get("font")) or "").strip().lower()
    if normalized in {n.lower() for n in MATH_FONT_NAMES}:
        return True
    return _is_math_run_text(span.get("text", ""))


# --------------------------------------------------------------------
# Paragraph (block) reconstruction per halaman
# --------------------------------------------------------------------


def _looks_like_page_number(text: str) -> bool:
    t = text.strip()
    return bool(t) and len(t) <= 6 and bool(_PAGE_NUM_ONLY_RE.match(t))


def _overlaps_table(
    bbox: tuple, table_bboxes: list[tuple], threshold: float = 0.5
) -> bool:
    """
    True kalau proporsi area `bbox` yang beririsan dengan salah satu
    `table_bboxes` >= threshold. Dipakai buat MEMBUANG blok teks yang
    sebenarnya isi sel tabel (bukan paragraf body) dari daftar paragraf
    -- tanpa ini, teks tabel ikut nyasar jadi "paragraf" body biasa,
    mengacaukan klasifikasi & index paragraf (yang dipakai tables_data
    buat mencari prev/next paragraf tetangga tabel).
    """
    bx0, by0, bx1, by1 = bbox
    area = max(bx1 - bx0, 0) * max(by1 - by0, 0)
    if area <= 0:
        return False
    for tx0, ty0, tx1, ty1 in table_bboxes:
        ix0, iy0 = max(bx0, tx0), max(by0, ty0)
        ix1, iy1 = min(bx1, tx1), min(by1, ty1)
        inter = max(ix1 - ix0, 0) * max(iy1 - iy0, 0)
        if inter / area >= threshold:
            return True
    return False


def _line_tail_token(line: dict) -> tuple[str, float, float] | None:
    """
    Token paling KANAN di satu baris PyMuPDF: span-span di ujung kanan yang
    saling berdekatan (celah <= 6pt) digabung jadi satu token, berhenti di
    celah lebar pertama. Dipakai membaca nomor identitas persamaan lewat
    POSISI (mis. "(3.4)" atau "4.1" tanpa kurung di pojok kanan), bukan lewat
    urutan teks. Return (teks, x0, x1) atau None.
    """
    spans = [sp for sp in (line.get("spans") or []) if sp.get("text", "").strip()]
    if not spans:
        return None
    spans.sort(key=lambda sp: sp["bbox"][0])
    group = [spans[-1]]
    for sp in reversed(spans[:-1]):
        if group[0]["bbox"][0] - sp["bbox"][2] <= 6.0:
            group.insert(0, sp)
        else:
            break
    text = "".join(sp.get("text", "") for sp in group)
    return text, group[0]["bbox"][0], group[-1]["bbox"][2]


def _extract_blocks(
    page: "fitz.Page", table_bboxes: list[tuple] | None = None
) -> list[dict]:
    """
    Ambil semua blok teks badan halaman (header/footer/nomor-halaman/
    isi-tabel sudah difilter). Tiap blok balik: {"full_text", "runs",
    "bbox", "line_bboxes"}.

    "runs" dibangun dengan MENYISIPKAN run spasi sintetis di antar
    baris yang di-wrap -- supaya "".join(r["text"] for r in runs) ==
    full_text, PERSIS asumsi yang dipakai foreign_term_checker.py &
    quotation_checker.py saat merekonstruksi paragraf dari runs.
    """
    page_height = page.rect.height
    table_bboxes = table_bboxes or []
    raw = page.get_text("dict")["blocks"]
    blocks = []

    for b in raw:
        if b.get("type") != 0:  # 0 = teks, 1 = gambar
            continue
        bbox = b["bbox"]
        y0, y1 = bbox[1], bbox[3]
        # Buang blok yang SELURUHNYA ada di zona header/footer.
        if y1 <= HEADER_FOOTER_ZONE_PT or y0 >= page_height - HEADER_FOOTER_ZONE_PT:
            continue
        # Buang blok yang isinya sel tabel (lihat docstring _overlaps_table).
        #
        # BUG DIPERBAIKI (2026-09-09, laporan user -- caption gambar
        # "hilang total" dari hasil ekstraksi PDF, bukan cuma salah
        # klasifikasi, sehingga compare_caption_numbering() tidak
        # pernah sempat mengecek kecocokan angka babnya): screenshot
        # tampilan UI (mis. "Gambar 3.2 Tampilan Validasi") sering
        # punya elemen kotak/border/grid yang bikin heuristik
        # _detect_tables_on_page() (page.find_tables()) salah
        # mengira area itu tabel. Kalau baris caption yang menempel
        # PERSIS di tepi gambar itu kebetulan >=50% areanya jatuh di
        # dalam bbox tabel palsu tsb, blok itu dibuang ATAS NAMA "isi
        # sel tabel" -- padahal captionnya sendiri jelas BUKAN sel
        # tabel. Akibatnya caption itu tidak pernah masuk daftar
        # paragraf sama sekali, jadi tidak pernah dicek apa pun
        # (ukuran font, rata tengah, DAN kecocokan angka bab di
        # compare_caption_numbering()) -- bukan salah lolos, tapi
        # tidak pernah dicek sama sekali. Fix: sebelum membuang, intip
        # cepat teks mentah blok ini (gabung semua span, belum lewat
        # pipeline merge/split penuh) -- kalau match CAPTION_RE
        # ("Gambar/Tabel N.N ..."), JANGAN dibuang meski overlap
        # tabel, karena baris yang berpola begitu tidak pernah
        # merupakan isi sel tabel sungguhan.
        raw_block_text = "".join(
            span.get("text", "")
            for line in b.get("lines", [])
            for span in line.get("spans", [])
        ).strip()
        looks_like_caption = is_caption_title(raw_block_text)
        if _overlaps_table(bbox, table_bboxes) and not looks_like_caption:
            continue

        # Gabung dulu "line" PyMuPDF yang baseline-nya nyaris identik
        # (selisih <=1pt) SEBELUM diproses jadi runs/line_bboxes.
        #
        # BUG (terverifikasi empiris via debug_line_spacing.py, lihat
        # temuan "Paragraf #429... Spasi baris 0.0"): pada baris yang
        # rata kiri-kanan (justify), PyMuPDF kadang memecah SATU baris
        # visual jadi BEBERAPA "line" internal di titik antar-kata --
        # spasi antar-kata yang melebar akibat justify membuat heuristik
        # clustering-nya salah baca celah itu sebagai baris baru.
        # Baseline semua fragmen itu IDENTIK (karena memang satu baris
        # yang sama), jadi kalau dibiarkan apa adanya, gap baseline
        # antar "baris" ini kehitung 0 -> line_spacing_infer() balik 0.0
        # padahal aslinya paragraf itu bisa saja 1.5 spasi.
        raw_lines = b["lines"]
        merged_pymupdf_lines = []
        for line in raw_lines:
            line_spans = line.get("spans") or []
            baseline = line_spans[0]["origin"][1] if line_spans else line["bbox"][3]
            if merged_pymupdf_lines:
                prev_baseline, prev_line = merged_pymupdf_lines[-1]
                if abs(baseline - prev_baseline) <= 1.0:
                    # Gabung ke line sebelumnya: sambung spans (urutan
                    # asli PyMuPDF sudah left-to-right dalam blok, jadi
                    # concat langsung aman) + union bbox.
                    prev_line["spans"] = prev_line["spans"] + line_spans
                    pb, lb = prev_line["bbox"], line["bbox"]
                    prev_line["bbox"] = (
                        min(pb[0], lb[0]),
                        min(pb[1], lb[1]),
                        max(pb[2], lb[2]),
                        max(pb[3], lb[3]),
                    )
                    continue
            merged_pymupdf_lines.append((baseline, dict(line)))
        pdf_lines = [ml[1] for ml in merged_pymupdf_lines]

        # BUG (baru ditemukan 2026-08-24): PyMuPDF kadang menyisipkan
        # "baris hantu" -- line dengan span yang teksnya kosong/spasi
        # saja -- terutama di halaman dengan banyak baris kosong sisa
        # (mis. halaman "DAFTAR TABEL" yang sebagian besar blank, atau
        # baris kosong bekas paragraf kosong hasil export Word->PDF).
        # Baris begini TIDAK menyumbang karakter apa pun ke full_text
        # (spans-nya toh di-skip saat membangun runs, lihat filter
        # `text == ""` di bawah), tapi kalau bbox-nya tetap dibiarkan
        # ikut serta di line_bboxes/group_bbox, union bbox blok jadi
        # jauh melenceng dari bbox teks yang BENERAN tercetak.
        # Efeknya fatal untuk _infer_alignment(): judul yang aslinya
        # rata tengah (mis. "DAFTAR GAMBAR", "DAFTAR TABEL") keliru
        # terbaca "rata kiri" karena bbox gabungannya "ditarik" sampai
        # ke posisi baris hantu yang kebetulan nempel margin kiri.
        # Verifikasi empiris: baris "DAFTAR GAMBAR" (14pt bold, benar2
        # center secara visual) punya line_bboxes asli [(242.0, ..,
        # 371.1, ..), (85.1, .., 88.1, ..)] -- baris kedua itu 3pt
        # doang lebarnya dan spans-nya kosong, tapi cukup buat menarik
        # x0 gabungan dari 242.0 turun ke 85.1, mengubah titik tengah
        # blok dari ~306pt (SANGAT dekat titik tengah kolom, harusnya
        # lolos sebagai "center") jadi ~228pt (jauh dari titik tengah,
        # kebaca "left"). Fix: buang baris yang sama sekali tidak
        # punya teks tercetak SEBELUM dipakai di mana pun (bbox atau
        # full_text) -- baris begini tidak pernah dibutuhkan untuk
        # keduanya.
        pdf_lines = [
            line
            for line in pdf_lines
            if "".join(s.get("text", "") for s in (line.get("spans") or [])).strip()
        ]

        # BUG DIPERBAIKI (jalur PDF -- "Ukuran font 12.0pt tidak sesuai
        # untuk tipe 'caption', seharusnya 11pt" padahal captionnya
        # sendiri 100% 11pt): PyMuPDF sering menggabungkan baris caption
        # gambar ("Gambar N.N ...") dengan baris atribusi "Sumber: ..."
        # tepat di bawahnya jadi SATU block PyMuPDF, karena jaraknya
        # rapat (cuma selisih satu baris single-spacing). Padahal
        # keduanya jenis paragraf yang beda (caption vs source_citation)
        # dengan aturan font BERBEDA (11pt vs 12pt) -- kalau dibiarkan
        # jadi satu elemen gabungan, size_char_weight di comparator.py
        # milih ukuran MAYORITAS karakter, dan teks "Sumber: <url>"
        # (seringnya lebih panjang dari teks caption-nya sendiri karena
        # URL panjang) menang jadi "dominant_size" -- caption yang
        # aslinya 100% 11pt pun ke-flag "12pt tidak sesuai". Fix: pecah
        # blok PyMuPDF di titik mana pun baris berikutnya diawali
        # "Sumber:"/"Sumber ." (SOURCE_CITATION_RE) -- baris itu SELALU
        # memulai paragraf/elemen baru sendiri, tidak peduli seberapa
        # rapat jaraknya dengan baris di atasnya.
        # BUG DIPERBAIKI (2026-09-09, laporan user -- "DAFTAR TABEL"
        # hilang total dari daftar chapter_title jalur PDF, padahal
        # "DAFTAR GAMBAR"/"DAFTAR ISI" di dokumen yang SAMA terdeteksi
        # normal): dibuktikan via debug langsung ke PDF asli -- pada
        # halaman "DAFTAR TABEL" yang isinya cuma pendek (judul + 2
        # baris entri tabel), SISA HALAMAN yang kosong (baris-baris
        # paragraf kosong bekas Word) tidak cukup lebar jaraknya buat
        # PyMuPDF mendeteksi batas blok baru -- akibatnya "DAFTAR
        # TABEL" dan KEDUA baris entri tabel di bawahnya ke-gabung
        # jadi SATU blok PyMuPDF (full_text = "DAFTAR TABEL Tabel 2.1
        # ... Tabel 3.1 ..."). classify_paragraph() mensyaratkan teks
        # PERSIS SAMA DENGAN "DAFTAR TABEL" (match FRONT_HEADINGS) atau
        # PERSIS "BAB N ..." (match CHAPTER_TITLE_RE) untuk mengenali
        # heading -- begitu ketiga baris itu ketercampur jadi satu
        # string panjang, ptype "chapter_title" TIDAK PERNAH kena,
        # jadi paragraf ini lolos tanpa dicek bold/center SAMA SEKALI
        # (halaman "DAFTAR GAMBAR" kebetulan selamat karena py MuPDF
        # sudah memisahkannya sendiri sebagai blok tersendiri di
        # dokumen ini -- bukan berarti fix di SOURCE_CITATION_RE di
        # atas ini cukup, karena masalahnya beda: bukan caption
        # nempel ke sitasi, tapi HEADING nempel ke entri daftar
        # isinya sendiri). Fix: sama seperti pemisahan di
        # SOURCE_CITATION_RE di atas, tambah titik pisah paksa kalau
        # sebuah baris PERSIS SAMA (setelah upper+strip) dengan salah
        # satu FRONT_HEADINGS ("DAFTAR TABEL", "DAFTAR GAMBAR", dst)
        # atau match penuh CHAPTER_TITLE_RE ("BAB N ...") -- baris
        # begini SELALU judul heading berdiri sendiri, jadi dipisah
        # baik SEBELUM (supaya tidak menelan baris badan sebelumnya)
        # MAUPUN SESUDAHNYA (supaya tidak menelan baris badan
        # sesudahnya, beda dari SOURCE_CITATION_RE yang cuma perlu
        # dipisah SEBELUM baris "Sumber:"-nya).
        # BUG DIPERBAIKI (2026-09-09, laporan user -- "Gambar 3.2
        # Tampilan Validasi" masih tidak kena cek nomor bab di
        # compare_caption_numbering() walau fix caption-vs-tabel-palsu
        # di _extract_blocks() sudah aktif): dibuktikan LANGSUNG ke PDF
        # asli -- di halaman ini SAMA SEKALI TIDAK ADA tabel terdeteksi
        # (table_bboxes kosong), jadi bug-nya BUKAN soal tabel palsu.
        # Root cause-nya beda lagi: baris SAMBUNGAN (wrap ke baris ke-2)
        # dari kutipan "Sumber: ..." SEBELUMNYA ("Artikel Ilmiah [21]")
        # ke-gabung PyMuPDF jadi SATU blok dengan baris caption
        # "Gambar 3.2 Tampilan Validasi" yang datang tepat sesudahnya
        # (jaraknya cukup rapat, ~13.8pt, satu spasi baris biasa).
        # Aturan pemisah SOURCE_CITATION_RE di atas HANYA mengenali
        # baris yang DIAWALI "Sumber:" -- baris sambungannya sendiri
        # ("Artikel Ilmiah [21]") sudah tidak diawali "Sumber:" lagi
        # (itu baris ke-2 dari kutipan yang sama), jadi caption yang
        # nempel di belakangnya lolos tanpa pernah kepisah. Fix:
        # tambahkan CAPTION_RE ("Gambar/Tabel N.N ...") sebagai pemicu
        # split-SEBELUM juga, sama seperti SOURCE_CITATION_RE -- begitu
        # ketemu baris yang PERSIS diawali pola caption, paksa mulai
        # grup baru di situ, TIDAK PEDULI baris sebelumnya isinya apa.
        # Sengaja TIDAK ditambah split-SESUDAH (beda dari
        # is_standalone_heading_line di atas) -- caption kadang wajar
        # nyambung ke baris berikutnya sendiri (judul caption panjang
        # yang wrap ke baris ke-2), itu HARUS tetap satu paragraf.
        split_pdf_line_groups = [[]]
        for line in pdf_lines:
            line_spans = line.get("spans") or []
            line_text = "".join(s.get("text", "") for s in line_spans).strip()
            line_text_upper = line_text.upper()
            is_standalone_heading_line = bool(line_text) and (
                line_text_upper in FRONT_HEADINGS
                or (
                    CHAPTER_TITLE_RE.match(line_text_upper)
                    and line_text == line_text_upper
                )
                or (
                    # LAMPIRAN setara BAB (lihat LAMPIRAN_TITLE_RE) --
                    # tanpa ini, "LAMPIRAN B" gampang ke-gabung PyMuPDF
                    # dengan baris badan di sekitarnya jadi satu blok,
                    # sama seperti bug lama "DAFTAR TABEL"/"BAB N" yang
                    # didokumentasikan di atas sebelum guard ini ada.
                    LAMPIRAN_TITLE_RE.match(line_text_upper)
                    and line_text == line_text_upper
                )
                or (
                    # Label sampul tetap ("PROYEK AKHIR"/"DOSEN
                    # PEMBIMBING") -- lihat catatan bug di
                    # _COVER_STANDALONE_LABEL_RE di atas.
                    _COVER_STANDALONE_LABEL_RE.match(line_text_upper)
                    and line_text == line_text_upper
                )
            )
            is_caption_start_line = is_caption_title(line_text)
            if (
                (
                    line_text
                    and SOURCE_CITATION_RE.match(line_text)
                    and split_pdf_line_groups[-1]
                )
                or (is_standalone_heading_line and split_pdf_line_groups[-1])
                or (is_caption_start_line and split_pdf_line_groups[-1])
            ):
                split_pdf_line_groups.append([])
            split_pdf_line_groups[-1].append(line)
            if is_standalone_heading_line:
                split_pdf_line_groups.append([])

        for pdf_lines_group in split_pdf_line_groups:
            if not pdf_lines_group:
                continue

            # BUG DIPERBAIKI (2026-09-08, laporan mahasiswa -- "nomor
            # identitas sudah ada tapi dibilang tidak ada" utk persamaan
            # PECAHAN/fraction, mis. rumus rata-rata bergerak "(A1+A2+
            # ...+n)/n"): PyMuPDF mengurutkan "line" DI DALAM SATU blok
            # murni berdasar posisi vertikal (atas->bawah). Untuk
            # persamaan pecahan, baris pembilang ada di atas & penyebut
            # di bawah, sedangkan nomor identitasnya ("(2.1)") secara
            # visual rata kanan TERHADAP KESELURUHAN persamaan tapi
            # posisi vertikalnya ada DI TENGAH (sejajar garis pembagi)
            # -- BUKAN sejajar baris penyebut (baris paling bawah).
            # Akibatnya urutan baca natural PyMuPDF jadi [pembilang,
            # nomor, penyebut], dan setelah digabung (lihat penggabungan
            # antar-baris di bawah) full_text berakhir "... (2.1) n" --
            # nomornya ada DI TENGAH string, bukan di akhir. Padahal
            # EQUATION_NUMBER_IN_TEXT_RE mewajibkan pola itu ada PERSIS
            # di akhir string ("$"), jadi gagal cocok -- comparator.py
            # salah lapor "tidak memiliki nomor identitas" padahal
            # sudah ditulis mahasiswa dengan benar & rata kanan.
            # Fix: kalau PERSIS SATU baris dalam grup ini isinya CUMA
            # nomor identitas (dicek via _BARE_EQUATION_NUMBER_RE,
            # full match -- bukan cuma "diakhiri" pola itu, supaya
            # baris body biasa yang kebetulan diakhiri "(1)" tidak ikut
            # kena) dan baris itu BUKAN baris terakhir, pindahkan ke
            # akhir urutan SEBELUM digabung -- konsisten dengan makna
            # visualnya sebagai identitas PENUTUP keseluruhan persamaan.
            bare_number_line_idxs = [
                li
                for li, line in enumerate(pdf_lines_group)
                if _BARE_EQUATION_NUMBER_RE.match(
                    normalize_equation_text(
                        "".join(s.get("text", "") for s in (line.get("spans") or []))
                    )
                )
            ]
            if (
                len(bare_number_line_idxs) == 1
                and bare_number_line_idxs[0] != len(pdf_lines_group) - 1
            ):
                number_line = pdf_lines_group.pop(bare_number_line_idxs[0])
                pdf_lines_group.append(number_line)
            elif not bare_number_line_idxs and len(pdf_lines_group) >= 2:
                # BUG DIPERBAIKI (laporan user -- persamaan pecahan
                # "Optimization Score = (Time+Power+Energy)/3 (3.4)" dibilang
                # tidak bernomor): nomor identitas sejajar baris "Optimization
                # Score =" (garis pembagi), jadi PyMuPDF menyatukannya dalam
                # SATU baris ("Optimization Score =(3.4)") yang urutannya di
                # TENGAH grup, sebelum penyebut ("3"). Kasus bare-number di atas
                # tidak kena karena barisnya bukan cuma nomor. Pindahkan baris
                # yang DIAKHIRI nomor ke akhir kalau semua baris sesudahnya
                # cuma serpihan pendek (penyebut/pangkat), bukan kalimat.
                tails = []
                for li, line in enumerate(pdf_lines_group):
                    ltxt = normalize_equation_text(
                        "".join(sp.get("text", "") for sp in (line.get("spans") or []))
                    )
                    if EQUATION_NUMBER_IN_TEXT_RE.search(ltxt):
                        tails.append(li)
                if len(tails) == 1 and tails[0] != len(pdf_lines_group) - 1:
                    after = pdf_lines_group[tails[0] + 1 :]
                    if all(
                        len(
                            "".join(
                                sp.get("text", "") for sp in (ln.get("spans") or [])
                            ).strip()
                        )
                        <= 8
                        for ln in after
                    ):
                        pdf_lines_group.append(pdf_lines_group.pop(tails[0]))

            runs = []
            line_bboxes = []
            line_baselines = []
            line_tails = []
            for li, line in enumerate(pdf_lines_group):
                line_bboxes.append(line["bbox"])
                line_tails.append(_line_tail_token(line))
                # Baseline ASLI baris (koordinat "origin" span pertama), BUKAN
                # top-of-bbox -- lihat docstring _infer_line_spacing() untuk
                # alasan kenapa ini penting.
                line_spans = line.get("spans") or []
                line_baselines.append(
                    line_spans[0]["origin"][1] if line_spans else line["bbox"][3]
                )
                for span in line["spans"]:
                    text = span.get("text", "")
                    if text == "":
                        continue
                    runs.append(
                        {
                            "text": text,
                            "font_name": _normalize_font_name(span.get("font")),
                            "font_size_pt": round(span.get("size", 0.0), 1),
                            "bold": _span_bold(span),
                            "italic": _span_italic(span),
                            "is_math": _span_is_math(span),
                        }
                    )
                if li < len(pdf_lines_group) - 1:
                    # Sambungan antar baris yang di-wrap -- spasi sintetis,
                    # format-nya ikut run TERAKHIR baris ini (paling wajar
                    # untuk keperluan italic_map di checker downstream).
                    inherit = runs[-1] if runs else None
                    runs.append(
                        {
                            "text": " ",
                            "font_name": inherit["font_name"] if inherit else None,
                            "font_size_pt": (
                                inherit["font_size_pt"] if inherit else None
                            ),
                            "bold": inherit["bold"] if inherit else False,
                            "italic": inherit["italic"] if inherit else False,
                            "is_math": False,
                        }
                    )

            full_text = "".join(r["text"] for r in runs)
            if not full_text.strip():
                continue

            group_bbox = (
                min(l["bbox"][0] for l in pdf_lines_group),
                min(l["bbox"][1] for l in pdf_lines_group),
                max(l["bbox"][2] for l in pdf_lines_group),
                max(l["bbox"][3] for l in pdf_lines_group),
            )

            # BUG DIPERBAIKI (2026-09-17, laporan user -- baris tahun pada
            # halaman sampul, mis. "2025", hilang total dari hasil
            # ekstraksi PDF): cek "nomor halaman badan yang nyasar ke luar
            # zona footer" di bawah ini SEBELUMNYA tidak diberi batas jarak
            # SAMA SEKALI dari tepi halaman -- ia menguji pola teks saja
            # (pendek + cuma digit/angka romawi), jadi SEMBARANG baris
            # pendek berisi angka DI MANA PUN di badan halaman (termasuk
            # "2025" di tengah halaman sampul) ikut terbuang, bukan cuma
            # nomor halaman yang SUNGGUH ada di dekat footer/header
            # (di luar HEADER_FOOTER_ZONE_PT=70pt tapi tetap dekat, sesuai
            # niat komentar aslinya). Akibatnya baris tahun sampul selalu
            # hilang, compare_cover_page() (comparator.py) tidak pernah
            # dapat kesempatan menutup state "tahun_block" jadi "done" --
            # baris SAMPUL BERIKUTNYA (baik pengulangan blok sampul kedua
            # yang sah maupun halaman lain sesudahnya) ikut kena aturan
            # sampul yang SALAH, karena state machine menyangka baris itu
            # baris tahun. Fix: tambahkan syarat jarak -- HANYA buang kalau
            # blok ini juga berada di zona yang LEBIH LEBAR dari
            # HEADER_FOOTER_ZONE_PT (mis. gara-gara margin bawah dokumen
            # kecil, angka page number bisa sedikit lolos dari 70pt) tapi
            # TETAP dekat tepi atas/bawah halaman -- BUKAN di mana pun di
            # badan halaman seperti sebelumnya.
            near_edge = (
                group_bbox[1] <= NEAR_EDGE_ZONE_PT
                or group_bbox[3] >= page_height - NEAR_EDGE_ZONE_PT
            )
            if (
                _looks_like_page_number(full_text)
                and len(pdf_lines_group) == 1
                and near_edge
            ):
                continue

            blocks.append(
                {
                    "full_text": full_text,
                    "runs": runs,
                    "bbox": group_bbox,
                    "line_bboxes": line_bboxes,
                    "line_baselines": line_baselines,
                    "line_tails": line_tails,
                }
            )

    return blocks


_SENTENCE_END_RE = re.compile(r"""["')\]]*[.!?:;]["')\]]*$""")


def _merge_wrapped_paragraph_blocks(
    blocks: list[dict], left_margin_pt: float, right_margin_pt: float
) -> list[dict]:
    """
    Gabung kembali blok-blok satu-baris (atau multi-baris pendek) yang
    SEBENARNYA satu paragraf yang sama, tapi kepecah jadi blok terpisah
    oleh heuristik clustering internal PyMuPDF -- kejadian PALING sering
    di paragraf ber-spasi lebar (>=1.5x), terverifikasi lewat
    debug_line_spacing.py di file hasil konversi LibreOffice.

    Tanpa fungsi ini, dampaknya BUKAN cuma line_spacing balik None
    (blok 1-baris tidak ada gap buat dihitung) -- font/alignment/indent
    per paragraf ikut salah karena satu paragraf asli dihitung sebagai
    banyak paragraf kecil terpisah, dan index paragraf (dipakai
    comparator.py buat referensi lokasi temuan) jadi tidak match jumlah
    paragraf yang sebenarnya.

    Syarat gabung blocks[i] -> blocks[i+1] (SEMUA harus terpenuhi):
      1. Baris terakhir blocks[i] TIDAK diakhiri tanda baca penutup
         kalimat (. ! ? : ; -- boleh diikuti kutip/kurung tutup).
         Paragraf yang beneran sudah selesai hampir selalu diakhiri
         tanda baca ini; kalau tidak, kalimatnya kemungkinan besar
         masih berlanjut ke blok berikutnya.
      2. Baris terakhir blocks[i] dan baris pertama blocks[i+1] sama-
         sama rata ke left_margin_pt (toleransi noise render). Format
         PENS tidak pakai indentasi first-line, jadi baris lanjutan
         paragraf yang di-wrap SELALU rata margin kiri kolom -- sama
         kayak baris pertama paragraf body biasa (makanya syarat ini
         SENDIRIAN tidak cukup, harus dibarengi #1 dan #3).
      3. Gap baseline (baris pertama blocks[i+1] - baris terakhir
         blocks[i]) masuk akal sebagai jarak antar-baris NORMAL
         (<= ~2.1x baseline single-line-height font tsb), bukan gap
         besar yang menandakan baris kosong/spasi-sebelum paragraf
         baru beneran.
      4. Ukuran font median kedua blok sama (toleransi 0.5pt) --
         paragraf lanjutan hampir pasti font sama; ganti ukuran font
         biasanya menandakan elemen lain (caption, heading, dst).

    CATATAN PENTING: ini heuristik yang BELUM diverifikasi terhadap PDF
    asli mahasiswa (baru diuji lewat simulasi/mock), beda dari fix
    baseline/1.17 yang sudah diverifikasi empiris. Kalau nanti dicoba
    di dokumen asli dan ada gejala:
      - Paragraf body ke-gabung sama heading/caption di sebelahnya
        (over-merge) -> longgarkan/perketat syarat #4, atau tambah cek
        bold sebagai syarat kelima.
      - Paragraf panjang yang seharusnya satu tetap kepecah
        (under-merge) -> gap aslinya di luar rentang 2.1x yang
        diasumsikan di syarat #3, perlu angka gap mentah dari
        debug_line_spacing.py di paragraf yang kepecah itu.

    FIX (2026-09-08 -- laporan user + debug_paragraph_split.py: 195
    pasangan block sehalaman gagal merge, gagal SEMUA di syarat #2
    ('left_ok'), termasuk kasus yang jelas-jelas satu kalimat yang
    di-wrap, mis. "...Proyek akhir merupakan karya ilmiah terapan yang
    wajib diselesaikan oleh setiap " lanjut ke blok "mahasiswa program
    sarjana terapan..." -- not_sentence_end/gap_ok/same_size semuanya
    OK, cuma left_ok gagal dengan selisih ~27pt yang KONSISTEN di
    banyak kasus serupa): dokumen mahasiswa ini pakai indentasi baris
    pertama (first-line indent) di paragraf body, gaya akademik umum
    yang TIDAK diasumsikan oleh syarat #2 versi lama ("format PENS
    tidak pakai indentasi first-line" -- asumsi itu tidak berlaku
    general). Kalau blocks[i] kebetulan cuma satu baris (baris pertama
    paragraf yang ter-cluster sendirian oleh PyMuPDF), baris itu WAJAR
    menjorok dari left_margin_pt sejauh indentasi standar, sedangkan
    baris LANJUTAN di blocks[i+1] rata di left_margin_pt biasa --
    dibandingkan satu sama lain (posisi x baris prev vs baris cur)
    keduanya beda ~27pt padahal itu satu paragraf yang sama. Fix:
    kalau blocks[i] cuma 1 baris DAN baris itu menjorok dari
    left_margin_pt sejauh indentasi wajar (antara noise floor ~0.1cm
    dan ~2cm) DAN baris pertama blocks[i+1] rata di left_margin_pt,
    anggap syarat #2 terpenuhi juga (paragraf ber-first-line-indent).
    Kalau blocks[i] lebih dari 1 baris, perilaku lama (rata baris-ke-
    baris) tetap dipakai apa adanya -- itu kasus baris LANJUTAN yang
    memang harus sama-sama rata, bukan baris pertama.
    """
    if len(blocks) < 2:
        return blocks

    tol = 0.25 / PT_TO_CM  # ~7pt, toleransi rata margin kiri
    noise_floor_pt = 0.1 / PT_TO_CM  # ~0.28pt/mm noise floor, selaras _infer_indent_cm
    first_line_indent_max_pt = (
        2.0 / PT_TO_CM
    )  # ~2cm, batas atas indentasi first-line wajar
    merged: list[dict] = []
    for block in blocks:
        if merged:
            prev = merged[-1]
            prev_text = prev["full_text"].rstrip()
            prev_last_line = prev["line_bboxes"][-1]
            cur_first_line = block["line_bboxes"][0]
            prev_last_baseline = prev["line_baselines"][-1]
            cur_first_baseline = block["line_baselines"][0]

            prev_sizes = [r["font_size_pt"] for r in prev["runs"] if r["font_size_pt"]]
            cur_sizes = [r["font_size_pt"] for r in block["runs"] if r["font_size_pt"]]
            font_size = statistics.median(prev_sizes) if prev_sizes else None
            same_size = bool(
                prev_sizes
                and cur_sizes
                and font_size is not None
                and abs(font_size - statistics.median(cur_sizes)) <= 0.5
            )

            not_sentence_end = bool(prev_text) and not _SENTENCE_END_RE.search(
                prev_text
            )
            # FIX (baru ditemukan 2026-08-24, kasus nyata: paragraf
            # "...sebagaimana ditunjukkan pada" [baris terakhir] gagal
            # gabung dengan "Gambar 3.10." [baris berikutnya] walau
            # keduanya SAMA-SAMA rata pada x=70.9pt -- PERSIS sama satu
            # sama lain, cuma beda dari left_margin_pt GLOBAL (85.1pt,
            # median seluruh dokumen). Sebagian halaman body text
            # dokumen ini ternyata rata pada margin kiri yang sedikit
            # berbeda dari median global (indikasi indentasi/margin
            # tidak 100% seragam antar bagian dokumen) -- versi lama
            # yang mensyaratkan KEDUA baris cocok ke left_margin_pt
            # GLOBAL gagal total di kasus begini, blok "Gambar 3.10."
            # (fragmen akhir kalimat, BUKAN caption sungguhan) jadi
            # tidak pernah ke-merge dan lolos ke classify_paragraph()
            # sebagai ptype "caption" murni karena teksnya kebetulan
            # cocok pola CAPTION_RE. Fix: syarat rata-kiri cukup
            # dicek antara DUA BARIS ITU SENDIRI (harus sama-sama
            # nempel di posisi x yang SAMA), bukan terhadap referensi
            # eksternal yang bisa saja tidak match margin lokal
            # halaman itu.
            left_ok = abs(prev_last_line[0] - cur_first_line[0]) <= tol
            if not left_ok and len(prev["line_bboxes"]) == 1:
                # Kemungkinan prev = baris pertama paragraf dengan gaya
                # indentasi yang beda dari baris lanjutannya -- ada DUA
                # pola yang mungkin, tergantung arah indentasi:
                prev_indent = prev_last_line[0] - left_margin_pt
                cur_indent = cur_first_line[0] - left_margin_pt

                # Kasus A: FIRST-LINE INDENT (paragraf body akademik
                # biasa) -- baris pertama (prev) menjorok, baris
                # lanjutan (cur) rata margin. (lihat catatan FIX
                # 2026-09-08 di atas.)
                first_line_indent_case = (
                    noise_floor_pt < prev_indent <= first_line_indent_max_pt
                    and abs(cur_indent) <= tol
                )

                # Kasus B: HANGING INDENT (format umum Daftar Pustaka
                # IEEE/APA) -- KEBALIKAN dari kasus A: baris pertama
                # (prev) rata margin, baris lanjutan (cur) yang
                # menjorok ke kanan. Tanpa ini, entri Daftar Pustaka
                # yang wrap ke >1 baris kepecah jadi beberapa elemen
                # "body" terpisah (baris-1 gagal gabung ke baris-2
                # karena kedua baris dianggap "tidak rata kiri sama
                # lain"), sehingga bibliography_checker.py menghitung
                # SATU entri fisik sebagai BEBERAPA entri logis --
                # inilah penyebab utama checked count DAFTAR PUSTAKA
                # jalur PDF bisa ~2x lipat dibanding jalur docx pada
                # dokumen yang identik (ditemukan 2026-09-09, dari
                # perbandingan hasil docx vs pdf: 52 vs 106 checked).
                #
                # REGRESI DIPERBAIKI (2026-09-09, ditemukan langsung
                # setelah kasus B di atas diaktifkan -- laporan: heading
                # "1.4 MANFAAT" [1 baris, rata margin] ke-merge dengan
                # kalimat body di bawahnya ["Sistem yang dikembangkan
                # ..."], jadi 1 elemen campuran yang salah diklasifikasi
                # & gagal cek bold/uppercase heading): pola posisi
                # "baris flush -> baris berikutnya indent" ITU SENDIRI
                # ternyata AMBIGU -- bisa berarti sambungan hanging-
                # indent (yang memang harus digabung), TAPI BISA JUGA
                # cuma kebetulan "heading pendek 1-baris" diikuti
                # "paragraf body baru ber-first-line-indent" (yang
                # justru harus TETAP terpisah, dua elemen berbeda).
                # `same_size` di atas tidak cukup membedakan dua kasus
                # ini kalau heading & body kebetulan sama-sama 12pt
                # (umum di template PA Indonesia). Pembeda yang lebih
                # andal: heading akademik (BAB/subbab) SELALU dicetak
                # bold PENUH, sedangkan body/entri Daftar Pustaka pada
                # umumnya TIDAK -- jadi kasus B cuma valid kalau baris
                # batas gabungan (akhir prev & awal cur) SAMA-SAMA
                # tidak bold. Kalau prev-nya full-bold (indikasi kuat
                # itu heading, bukan entri pustaka), JANGAN gabung.
                prev_content_runs = [r for r in prev["runs"] if r["text"].strip()]
                prev_all_bold = bool(prev_content_runs) and all(
                    r["bold"] for r in prev_content_runs
                )
                cur_content_runs = [r for r in block["runs"] if r["text"].strip()]
                cur_first_bold = bool(cur_content_runs) and cur_content_runs[0]["bold"]

                hanging_indent_case = (
                    abs(prev_indent) <= tol
                    and noise_floor_pt < cur_indent <= first_line_indent_max_pt
                    and not prev_all_bold
                    and not cur_first_bold
                )

                if first_line_indent_case or hanging_indent_case:
                    left_ok = True

            # BUG DIPERBAIKI (jalur PDF -- laporan user: baris kedua
            # caption gambar yang wrap, mis. "Gambar 4.2 Antarmuka ...
            # Panel Kendali" lanjut ke baris "Parameterisasi Prompting"
            # di bawahnya, gagal ke-merge dan jatuh ke ptype "body"
            # sendiri -- akibatnya baris wrap itu (yang sebenarnya
            # bagian caption, 11pt) ke-flag "font 11pt tidak sesuai
            # untuk tipe body, seharusnya 12pt"): caption GAMBAR/TABEL
            # di template ini dicetak RATA TENGAH, bukan rata kiri --
            # tiap baris caption diposisikan tengah SENDIRI-SENDIRI
            # berdasarkan lebar tulisannya masing-masing, jadi baris
            # ke-2 caption ("Parameterisasi Prompting") x0-nya jauh
            # dari x0 baris ke-1 ("Gambar 4.2 ...") walau keduanya
            # sebenarnya satu caption yang sama -- syarat left_ok di
            # atas (dirancang utk paragraf rata-kiri) SELALU gagal
            # untuk kasus ini. Verifikasi langsung ke PDF: titik tengah
            # kedua baris HAMPIR SAMA PERSIS (mis. x-center baris 1
            # 331.1pt vs baris 2 332.5pt) walau x0/x1 masing-masing
            # beda jauh -- ciri khas teks rata tengah. Fix: KHUSUS
            # kalau blok prev memang caption ("Gambar N.N ..."/
            # "Tabel N.N ..." -- CAPTION_RE), izinkan gabung juga kalau
            # titik tengah baris terakhir prev & baris pertama cur
            # SAMA (toleransi sama seperti left_ok), tanpa mensyaratkan
            # left_ok sama sekali -- caption memang boleh rata tengah,
            # beda dari body/pustaka yang harus rata kiri. Sengaja
            # dibatasi HANYA untuk prev yang caption (bukan bikin
            # syarat "rata tengah = boleh gabung" berlaku umum ke semua
            # paragraf) supaya tidak menimbulkan over-merge baru di
            # elemen center-aligned lain (mis. judul bab) yang justru
            # HARUS tetap 1 baris/elemen terpisah.
            if not left_ok and is_caption_title(prev_text.strip()):
                prev_center = (prev_last_line[0] + prev_last_line[2]) / 2
                cur_center = (cur_first_line[0] + cur_first_line[2]) / 2
                if abs(prev_center - cur_center) <= tol:
                    left_ok = True

            gap = cur_first_baseline - prev_last_baseline
            gap_ok = bool(font_size) and 0 < gap <= font_size * 1.17 * 2.1

            # BUG DIPERBAIKI (laporan user, jalur PDF -- heading "4.7 ANALISIS
            # HASIL EKSPERIMEN" ke-flag "seharusnya ditulis kapital semua"
            # padahal sudah kapital semua di dokumen asli): heading yang
            # bold PENUH dan 1 baris, diikuti item list/body yang RATA
            # KIRI di x yang sama (mis. "a. Analisis Model Sentimen"),
            # lolos SEMUA syarat di atas (tidak berakhir tanda baca, left_ok,
            # gap normal, ukuran sama 12pt) dan ke-merge jadi satu elemen
            # "4.7 ANALISIS HASIL EKSPERIMEN a.Analisis Model Sentimen".
            # Teks campuran itu lantas dinilai is_upper=False walau
            # judulnya sendiri sudah benar. Syarat kelima (sudah
            # dijanjikan di catatan "CATATAN PENTING" docstring): kalau
            # blok prev bold PENUH tapi teks pertama blok berikutnya
            # TIDAK bold, itu pola "heading lalu isi", bukan baris
            # lanjutan paragraf yang sama -- jangan digabung. Heading
            # panjang yang wrap 2 baris tetap tergabung (kedua bloknya
            # sama-sama bold).
            heading_then_body = False
            _prev_runs_txt = [r for r in prev["runs"] if r["text"].strip()]
            _cur_runs_txt = [r for r in block["runs"] if r["text"].strip()]
            if _prev_runs_txt and _cur_runs_txt:
                heading_then_body = all(r["bold"] for r in _prev_runs_txt) and not (
                    _cur_runs_txt[0]["bold"]
                )

            if (
                not_sentence_end
                and left_ok
                and gap_ok
                and same_size
                and not heading_then_body
            ):
                inherit = prev["runs"][-1] if prev["runs"] else None
                prev["runs"].append(
                    {
                        "text": " ",
                        "font_name": inherit["font_name"] if inherit else None,
                        "font_size_pt": inherit["font_size_pt"] if inherit else None,
                        "bold": inherit["bold"] if inherit else False,
                        "italic": inherit["italic"] if inherit else False,
                        "is_math": False,
                    }
                )
                prev["runs"].extend(block["runs"])
                prev["full_text"] = prev["full_text"] + " " + block["full_text"]
                pb, cb = prev["bbox"], block["bbox"]
                prev["bbox"] = (
                    min(pb[0], cb[0]),
                    min(pb[1], cb[1]),
                    max(pb[2], cb[2]),
                    max(pb[3], cb[3]),
                )
                prev["line_bboxes"].extend(block["line_bboxes"])
                prev["line_baselines"].extend(block["line_baselines"])
                prev.setdefault("line_tails", []).extend(block.get("line_tails", []))
                continue

        merged.append(
            {
                "full_text": block["full_text"],
                "runs": list(block["runs"]),
                "bbox": block["bbox"],
                "line_bboxes": list(block["line_bboxes"]),
                "line_baselines": list(block["line_baselines"]),
                "line_tails": list(block.get("line_tails", [])),
            }
        )
    return merged


# --------------------------------------------------------------------
# Heuristik pengganti atribut docx (alignment, indent, dsb)
# --------------------------------------------------------------------


def _infer_alignment(
    block: dict, left_margin_pt: float, right_margin_pt: float
) -> str | None:
    """
    Dekati w:jc dari posisi baris:
      - center : titik tengah bbox blok ~ titik tengah kolom teks
      - right  : sisi kanan blok nempel batas kanan, sisi kiri menjorok
      - both (justify) : SEMUA baris (kecuali baris terakhir) nyaris
        menyentuh batas kiri MAUPUN kanan kolom
      - left   : default kalau tidak match pola di atas
    Toleransi 0.45 cm (dikonversi ke pt) -- dilonggarkan dari 0.3cm
    (2026-08-23) supaya cek "center" tetap tahan terhadap noise residual
    di left_margin_pt/right_margin_pt (nilai itu ESTIMASI per-dokumen
    dari extract_section_params_pdf(), bukan dibaca presis seperti w:jc
    di docx) -- tanpa longgar secukupnya, judul yang SUDAH center di
    dokumen asli masih bisa salah ditandai "belum rata tengah" kalau
    estimasi margin dokumen itu meleset beberapa milimeter dari nilai
    sebenarnya. 0.45cm masih cukup ketat untuk tetap menangkap judul
    yang BENERAN rata kiri/kanan (yang titik tengahnya biasanya bergeser
    >1cm dari titik tengah kolom).
    """
    tol = 0.45 / PT_TO_CM  # ~12.75pt
    col_left, col_right = left_margin_pt, right_margin_pt
    col_center = (col_left + col_right) / 2
    bbox = block["bbox"]
    x0, x1 = bbox[0], bbox[2]
    center = (x0 + x1) / 2

    # NOTE (fixed): versi sebelumnya juga mensyaratkan (col_right -
    # col_left) - (x1 - x0) > tol, dengan maksud "blok tidak boleh
    # selebar kolom" biar gak salah anggap paragraf body rata-kiri
    # sebagai center. Tapi ini gagal untuk baris CENTER yang kebetulan
    # hampir selebar kolom (kasus nyata: baris "Sumber: ..." di bawah
    # gambar/tabel) -- baris itu pas di tengah kolom TAPI lebarnya
    # nyaris = lebar kolom, jadi syarat kedua ini menolaknya walau
    # syarat pertama (titik tengah blok == titik tengah kolom) sudah
    # cukup buat bedain dari rata-kiri (yang titik tengahnya BAKAL
    # geser ke kiri kalau baris itu tidak selebar kolom). Titik tengah
    # sendirian sudah cukup diskriminatif; syarat kedua dibuang.
    if abs(center - col_center) <= tol:
        return "center"

    if abs(x1 - col_right) <= tol and (x0 - col_left) > tol:
        return "right"

    lines = block["line_bboxes"]
    if len(lines) >= 2:
        body_lines = lines[:-1]  # baris terakhir paragraf wajar tidak penuh
        left_ok = all(abs(l[0] - col_left) <= tol for l in body_lines)
        right_ok = all(abs(l[2] - col_right) <= tol for l in body_lines)
        if left_ok and right_ok:
            return "both"

    return "left"


def _infer_equation_alignment(
    block: dict,
    left_margin_pt: float,
    right_margin_pt: float,
    extra_left_margins_pt: tuple = (),
) -> str | None:
    """
    Rata KHUSUS untuk blok ptype "equation", TERPISAH dari
    _infer_alignment() generik di atas.

    BUG YANG DIPERBAIKI (2026-09-08, laporan mahasiswa -- "rumus rata
    tengah tapi tidak pernah ke-flag rata kiri" di jalur PDF): perhatikan
    baik-baik alur _infer_alignment() generik -- fungsi itu cuma
    mendeteksi POSITIF pola "center"/"right"/"both (justify)"; begitu
    TIDAK ADA satu pun yang cocok, dia BALIK "left" SEBAGAI DEFAULT,
    bukan karena benar-benar mendeteksi sisi kiri blok nempel margin
    kiri. Untuk paragraf body/heading biasa, default ini praktis selalu
    benar (kalau bukan center/right/justify, ya memang rata kiri).

    TAPI untuk blok persamaan yang menggabungkan RUMUS (sering
    diposisikan Word rata TENGAH kolom lewat m:oMathParaPr, independen
    dari w:pPr/w:jc paragraf pembungkusnya -- lihat catatan setara di
    extractor.py versi docx) DENGAN nomor identitasnya (dipepet ke
    margin KANAN), bbox gabungan keduanya jadi bentuk yang JANGGAL:
    sisi kiri ada di tengah kolom (awal rumus), sisi kanan nempel
    margin kanan (nomor identitas). Titik tengah bbox gabungan ini
    TERTARIK ke kanan oleh nomor -- gagal cocok pola "center" (yang
    minta titik tengah blok == titik tengah kolom) MAUPUN "right" (yang
    minta sisi KIRI *tidak* nempel margin kiri sambil sisi kanan
    nempel -- ini justru YANG TERJADI DI SINI, harusnya kena, tapi
    kalau toleransinya pas-pasan meleset dikit, blok ini jatuh ke
    default "left" padahal SISI KIRI RUMUSNYA JELAS TIDAK PERNAH nempel
    margin kiri sama sekali).

    Fix: JANGAN pakai "left" sebagai hasil residual/default. Di sini
    "left" HANYA dikembalikan kalau sisi KIRI blok memang positif
    terverifikasi dekat left_margin_pt. Kalau tidak, balikin "center"
    atau "right" sesuai bentuknya, atau None (dianggap "tidak diset",
    tetap gagal cek rata-kiri di comparator.py) kalau bentuknya
    ambigu -- yang penting TIDAK PERNAH diam-diam dianggap "left" tanpa
    bukti positif.
    """
    tol = 0.45 / PT_TO_CM  # konsisten dengan toleransi _infer_alignment()
    x0, x1 = block["bbox"][0], block["bbox"][2]

    # BUG DIPERBAIKI (laporan user -- persamaan yang JELAS sudah mentok kiri
    # tetap dibilang "belum diatur perataannya"): margin kiri halaman ini
    # diambil dari paritas halaman FISIK (ganjil/genap). Di dokumen
    # mirror-margin yang penomoran halamannya restart (BAB 1 mulai dari 1),
    # paritas nomor TERCETAK bisa berlawanan dengan paritas fisik, sehingga
    # margin kiri yang dipakai tertukar (selisih ~1cm > toleransi 0.45cm)
    # dan rumus yang rata kiri sungguhan jatuh ke "None". Cek rata-kiri
    # sekarang menerima margin kiri MANA PUN dari set margin dokumen
    # (ganjil/genap/gabungan). Rumus rata tengah tetap tidak lolos karena
    # x0-nya berjarak jauh (>3cm) dari semua margin itu.
    left_candidates = (left_margin_pt, *extra_left_margins_pt)
    if any(abs(x0 - m) <= tol for m in left_candidates):
        return "left"

    col_center = (left_margin_pt + right_margin_pt) / 2
    center = (x0 + x1) / 2
    if abs(center - col_center) <= tol:
        return "center"

    if abs(x1 - right_margin_pt) <= tol:
        return "right"

    return None


# Sama dengan pola di extractor.py (DILONGGARKAN 2026-09-21: spasi di
# dalam kurung, "(3-1)"/"(3,1)", "Pers.", kurung fullwidth) -- dipakai
# bersama supaya docx & PDF tidak beda tafsir soal "apa itu nomor
# persamaan".
_BARE_EQUATION_NUMBER_RE = EQUATION_NUMBER_RE


_BARE_NUMBER_TOKEN_RE = EQUATION_NUMBER_BARE_RE


def _has_visual_equation_number(block: dict, right_margin_pt: float) -> bool:
    """
    True kalau SALAH SATU baris blok punya token di ujung KANAN (dekat margin
    kanan, <= 2.5cm) yang berupa nomor identitas: "(3.4)", "(Persamaan 3.4)"
    ATAU angka bab-urut telanjang "4.1" (tanpa kurung). Dibaca dari posisi,
    jadi tidak terpengaruh urutan teks rumus bertingkat.
    """
    near = 2.5 / PT_TO_CM
    for tail in block.get("line_tails") or []:
        if not tail:
            continue
        text, _x0, x1 = tail
        if x1 < right_margin_pt - near:
            continue
        t = normalize_equation_text(text)
        if (
            EQUATION_NUMBER_RE.match(t)
            or _BARE_NUMBER_TOKEN_RE.match(t)
            or EQUATION_NUMBER_IN_TEXT_RE.search(t)
        ):
            return True
    return False


def _slice_runs(runs: list[dict], start: int, end: int) -> list[dict]:
    """Potong daftar run menurut offset karakter [start, end) di full_text."""
    out = []
    offset = 0
    for r in runs:
        r_start, r_end = offset, offset + len(r["text"])
        offset = r_end
        lo, hi = max(r_start, start), min(r_end, end)
        if lo >= hi:
            continue
        piece = dict(r)
        piece["text"] = r["text"][lo - r_start : hi - r_start]
        out.append(piece)
    return out


def _move_equation_number_to_end(elem: dict) -> None:
    """
    Pindahkan nomor "(3.4)" yang terjebak di TENGAH full_text (karena urutan
    baca PyMuPDF: pembilang, baris garis-pembagi+nomor, penyebut) ke akhir,
    selama sisa teks sesudahnya cuma serpihan pendek (penyebut/pangkat).
    full_text & runs diubah bersama supaya invarian "".join(runs) == full_text
    tetap terjaga.
    """
    text = elem["full_text"]
    if EQUATION_NUMBER_IN_TEXT_RE.search(normalize_equation_text(text)):
        return
    matches = [
        m
        for m in re.finditer(r"[(\uff08][^()\uff08\uff09]{1,24}[)\uff09]", text)
        if EQUATION_NUMBER_RE.match(normalize_equation_text(m.group(0)))
    ]
    if len(matches) != 1:
        return
    m = matches[0]
    tail = text[m.end() :]
    if not tail.strip() or len(tail.strip()) > 12:
        return
    runs = elem["runs"]
    if "".join(r["text"] for r in runs) != text:
        return
    joiner = {
        "text": " ",
        "font_name": None,
        "font_size_pt": None,
        "bold": False,
        "italic": False,
        "is_math": False,
    }
    new_runs = (
        _slice_runs(runs, 0, m.start())
        + _slice_runs(runs, m.end(), len(text))
        + [joiner]
        + _slice_runs(runs, m.start(), m.end())
    )
    elem["runs"] = new_runs
    elem["full_text"] = "".join(r["text"] for r in new_runs)
    elem["equation_number_visual"] = True


def _merge_equation_fragments(elements: list[dict]) -> list[dict]:
    """
    Gabungkan serpihan SATU persamaan yang oleh PyMuPDF dipecah jadi
    beberapa blok (pembilang, "Optimization Score =", penyebut, nomor).

    BUG DIPERBAIKI (laporan user -- satu persamaan dilaporkan 2-3 kali dengan
    pesan yang sama, mis. temuan "Penomoran persamaan" nomor 12 dan 14 di
    halaman yang sama): tiap serpihan lolos jadi elemen ptype "equation"
    sendiri, lalu masing-masing dicek rata kiri, baris kosong, dan nomor
    -- serpihan tanpa nomor selalu gagal walau persamaannya utuh sudah
    bernomor. Sekarang serpihan yang berdempetan secara vertikal (celah
    <= 8pt atau saling tumpang) di halaman yang sama digabung jadi satu
    elemen. Serpihan non-equation hanya ikut digabung kalau SANGAT pendek
    (<= 12 karakter, mis. penyebut "3") dan berada di dalam rentang
    horizontal persamaan, supaya paragraf body tidak pernah ikut tertelan.
    """
    if not elements:
        return elements

    gap_tol = 8.0
    consumed: set[int] = set()
    for i, elem in enumerate(elements):
        if i in consumed or elem["type"] != "equation":
            continue
        for j in range(i + 1, min(i + 6, len(elements))):
            if j in consumed:
                continue
            cand = elements[j]
            if cand.get("page_index") != elem.get("page_index"):
                break
            eb, cb = elem["bbox"], cand["bbox"]
            vertical_gap = max(cb[1] - eb[3], eb[1] - cb[3])
            if vertical_gap > gap_tol:
                break  # elemen berikutnya makin jauh di bawah
            if cand["type"] == "equation":
                pass
            else:
                ctext = cand["full_text"].strip()
                inside_x = cb[0] >= eb[0] - 30 and cb[2] <= eb[2] + 30
                if not (
                    cand["type"] == "body"
                    and 0 < len(ctext) <= 12
                    and inside_x
                ):
                    continue
            elem["full_text"] = (
                elem["full_text"].rstrip() + " " + cand["full_text"].strip()
            ).strip()
            elem["runs"] = (
                elem["runs"]
                + [
                    {
                        "text": " ",
                        "font_name": None,
                        "font_size_pt": None,
                        "bold": False,
                        "italic": False,
                        "is_math": False,
                    }
                ]
                + cand["runs"]
            )
            elem["bbox"] = (
                min(eb[0], cb[0]),
                min(eb[1], cb[1]),
                max(eb[2], cb[2]),
                max(eb[3], cb[3]),
            )
            elem["equation_number_visual"] = bool(
                elem.get("equation_number_visual")
                or cand.get("equation_number_visual")
            )
            if cand["type"] == "equation" and cand.get("alignment") == "left":
                elem["alignment"] = "left"
            consumed.add(j)

    if not consumed:
        return elements
    return [e for k, e in enumerate(elements) if k not in consumed]


def _merge_split_equation_numbers(elements: list[dict]) -> list[dict]:
    """
    Gabungkan blok "equation" yang nomor identitasnya KEPISAH jadi blok
    PyMuPDF TERSENDIRI, kembali ke blok equation induknya.

    BUG YANG DIPERBAIKI (2026-09-08, laporan mahasiswa -- "nomor
    identitas sudah ada tapi dibilang tidak ada" di jalur PDF): rumus
    yang diposisikan rata tengah kolom dan nomor identitasnya dipepet
    ke margin kanan SERING punya jarak horizontal yang jauh (isi rumus
    di tengah, nomor jauh di kanan) -- PyMuPDF (page.get_text("dict"))
    mengelompokkan teks jadi "block" berdasarkan kedekatan spasial,
    sehingga jarak sejauh itu kerap membuat rumus dan nomor identitas
    dianggap DUA BLOK TERPISAH, walau keduanya visualnya satu baris
    yang sama. Akibatnya elemen ptype "equation" (blok rumus saja)
    full_text-nya TIDAK PERNAH memuat nomor identitasnya sendiri (yang
    nyasar jadi elemen "body" terpisah tak jauh setelahnya) --
    comparator.py pun salah lapor "tidak memiliki nomor identitas"
    padahal mahasiswa SUDAH menulisnya, cuma programnya yang gagal
    menyatukan dua blok ini kembali.

    HEURISTIK PENGGABUNGAN (sengaja konservatif, supaya tidak salah
    gabung elemen lain yang kebetulan mirip):
      - Elemen kandidat harus PERSIS berupa nomor identitas SENDIRIAN
        (full_text SELURUHNYA cuma "(2.1)"/"(Persamaan 2.1)", dicek via
        fullmatch -- BUKAN cuma diakhiri pola itu) -- badan paragraf
        biasa yang KEBETULAN diakhiri kurung angka tidak ikut kena.
      - Hanya dicari BEBERAPA elemen berikutnya (default 2) di
        HALAMAN YANG SAMA -- equation & nomornya tidak mungkin
        terpisah lebih jauh dari itu dalam urutan baca PyMuPDF.
      - Posisi vertikal (bbox tengah) kandidat harus DEKAT (toleransi
        1 baris, ~1.5x tinggi blok equation) dengan posisi vertikal
        blok equation -- memastikan memang "satu baris yang sama",
        bukan paragraf lain yang kebetulan cuma berisi angka berkurung
        (mis. poin bernomor "(1)" di body text biasa).
      - Kandidat harus berada di SEBELAH KANAN blok equation (bbox
        kiri kandidat >= bbox kanan equation dikurangi sedikit
        toleransi overlap) -- equation & nomornya tidak mungkin
        tertukar posisi.
      - equation YANG SUDAH PUNYA nomor identitas di full_text-nya
        sendiri (regex sudah match) TIDAK disentuh -- fungsi ini
        murni perbaikan untuk kasus yang literally kepisah blok.

    Kalau ketemu, teks & runs kandidat DIGABUNG ke elemen equation
    (supaya EQUATION_NUMBER_IN_TEXT_RE di comparator.py bisa menemukan
    nomornya), bbox equation diperluas mencakup kandidat, dan elemen
    kandidat DIHAPUS TOTAL dari daftar -- dia bukan paragraf mandiri,
    cuma "pecahan" dari equation yang sama.
    """
    consumed_indices = set()

    for i, elem in enumerate(elements):
        if elem["type"] != "equation":
            continue
        if EQUATION_NUMBER_IN_TEXT_RE.search(
            normalize_equation_text(elem["full_text"])
        ):
            continue  # sudah punya nomor di blok-nya sendiri, tidak perlu gabung

        eq_bbox = elem["bbox"]
        eq_center_y = (eq_bbox[1] + eq_bbox[3]) / 2
        eq_height = max(eq_bbox[3] - eq_bbox[1], 1.0)
        v_tol = eq_height * 1.5

        # Jendela pencarian dilebarkan 2 -> 4 elemen (2026-09-21): rumus
        # bertingkat sering dipecah PyMuPDF jadi beberapa blok kecil
        # (pembilang, penyebut, simbol) sebelum blok nomornya muncul.
        # Syarat halaman sama + sejajar vertikal + di sebelah kanan
        # tetap menjaga supaya tidak salah gabung.
        for j in range(i + 1, min(i + 5, len(elements))):
            if j in consumed_indices:
                continue
            cand = elements[j]
            if cand.get("page_index") != elem.get("page_index"):
                break  # sudah lewat halaman, tidak mungkin ketemu lagi
            cand_text = normalize_equation_text(cand["full_text"])

            is_bare_number = bool(_BARE_EQUATION_NUMBER_RE.match(cand_text))
            trailing_match = None
            if not is_bare_number:
                # KASUS TAMBAHAN (laporan mahasiswa, rumus PECAHAN --
                # mis. rumus rata-rata bergerak "(A1+A2+...+n)/n"):
                # untuk pecahan, PENYEBUT (denominator, mis. "n") posisi
                # vertikalnya SAMA dengan baris nomor identitas (nomor
                # dipepet rata kanan sejajar baris penyebut, bukan
                # sejajar pembilang), sehingga PyMuPDF sering
                # mengelompokkan penyebut & nomor identitas jadi SATU
                # block yang sama (mis. full_text-nya "n (2.1)") --
                # BUKAN dua block terpisah. Block gabungan ini juga
                # classify_paragraph()/rasio is_math gagal dianggap
                # "equation" (isinya cuma satu variabel + nomor,
                # mayoritas karakter BUKAN glyph matematis), sehingga
                # jatuh sebagai ptype "body" biasa dan tidak pernah
                # dicek oleh _BARE_EQUATION_NUMBER_RE (yang mewajibkan
                # SELURUH teks cuma nomor, tanpa "n" di depannya).
                # Akibatnya persamaan (block pembilang, ptype
                # "equation") tidak pernah ketemu nomornya --
                # comparator.py salah lapor "tidak memiliki nomor
                # identitas" padahal mahasiswa sudah menulisnya persis
                # di ujung baris penyebut.
                # Fix: terima juga kandidat yang CUMA DIAKHIRI pola
                # nomor identitas (bukan fullmatch), asalkan sisa teks
                # SEBELUM nomor itu pendek (<=6 karakter -- cukup untuk
                # variabel/denominator tunggal seperti "n", "x2", tapi
                # menolak paragraf body sungguhan yang kebetulan
                # diakhiri kurung angka).
                trailing_match = EQUATION_NUMBER_IN_TEXT_RE.search(cand_text)
                if not trailing_match:
                    continue
                remainder = cand_text[: trailing_match.start()].strip()
                if len(remainder) > 6:
                    continue

            cand_bbox = cand["bbox"]
            cand_center_y = (cand_bbox[1] + cand_bbox[3]) / 2
            # Kandidat "campuran" (penyebut + nomor) posisinya DI BAWAH
            # pembilang (bukan sejajar tengah blok equation induk kalau
            # equation induk cuma berisi pembilang), jadi toleransi
            # vertikalnya dilonggarkan sedikit dibanding kandidat nomor
            # murni.
            v_tol_use = v_tol if is_bare_number else v_tol * 1.6
            if abs(cand_center_y - eq_center_y) > v_tol_use:
                continue

            if is_bare_number:
                if cand_bbox[0] < eq_bbox[2] - 2.0:  # toleransi overlap kecil
                    continue
            else:
                # Kandidat campuran boleh MULAI lebih ke kiri dari sisi
                # kanan equation induk (penyebut pecahan lazimnya lebih
                # sempit & lebih ke tengah dibanding pembilang) -- yang
                # jadi syarat di sini justru sisi KANAN kandidat harus
                # melewati sisi kanan equation induk, sebagai bukti
                # kandidat memuat sesuatu (nomor identitas) yang posisi
                # rata kanannya lebih jauh ke kanan daripada equation
                # itu sendiri.
                if cand_bbox[2] <= eq_bbox[2] + 2.0:
                    continue

            # --- gabungkan ---
            elem["full_text"] = (
                elem["full_text"] + " " + cand["full_text"].strip()
            ).strip()
            # Run spasi sintetis penyambung -- skema field SAMA seperti
            # run sintetis penyambung baris wrap di _extract_blocks(),
            # supaya konsumen elem["runs"] lain (mis. cek bold nomor
            # identitas di compare_visual_objects()) tidak KeyError.
            joiner_run = {
                "text": " ",
                "font_name": None,
                "font_size_pt": None,
                "bold": False,
                "italic": False,
                "is_math": False,
            }
            elem["runs"] = elem["runs"] + [joiner_run] + cand["runs"]
            elem["bbox"] = (
                min(eq_bbox[0], cand_bbox[0]),
                min(eq_bbox[1], cand_bbox[1]),
                max(eq_bbox[2], cand_bbox[2]),
                max(eq_bbox[3], cand_bbox[3]),
            )
            elem["equation_number_visual"] = bool(
                elem.get("equation_number_visual")
                or cand.get("equation_number_visual")
                or is_bare_number
                or trailing_match
            )
            consumed_indices.add(j)
            break  # satu equation cuma butuh satu nomor

    if not consumed_indices:
        return elements
    return [e for i, e in enumerate(elements) if i not in consumed_indices]


# Rasio gap vertikal (dinormalisasi ke satuan "tinggi 1 baris tunggal",
# size*1.15 -- konstanta kalibrasi SAMA seperti _infer_line_spacing(),
# lihat docstring-nya) yang jadi ambang batas "ADA baris kosong" di
# _has_blank_gap() di bawah. Transisi antar-paragraf TANPA baris kosong
# (asumsi "space before/after paragraph" default 0) jaraknya kira-kira
# SAMA dengan jarak antar-baris DALAM paragraf yang sama (~1.5x baris
# tunggal untuk body text spasi 1.5). Kalau ADA 1 baris kosong
# disisipkan, jarak itu kira-kira DUA KALI LIPAT (~3.0x baris tunggal)
# -- baris kosong itu sendiri "memakan" satu baris tambahan. Ambang
# 2.2 dipilih di TENGAH antara ~1.5 (tidak ada) dan ~3.0 (ada), supaya
# cukup jauh dari keduanya untuk menyerap noise estimasi posisi.
_BLANK_LINE_GAP_RATIO_THRESHOLD = 2.2


def _has_blank_gap(before: dict | None, after: dict | None) -> bool:
    """
    True kalau gap vertikal antara elemen `before` & `after` (dua
    elemen BERURUTAN di `elements`, mestinya sudah dipastikan
    bersebelahan langsung oleh pemanggil) cukup lebar untuk
    mengindikasikan ada 1 baris kosong di antaranya -- lihat komentar
    _BLANK_LINE_GAP_RATIO_THRESHOLD di atas untuk penjelasan rasionya.

    Return True (dianggap ADA baris kosong / lolos cek) kalau salah
    satu elemen None (batas dokumen), ATAU keduanya beda halaman fisik
    (page_index beda -- pergantian halaman, gap vertikalnya tidak
    representatif "baris kosong" sungguhan lagi), ATAU datanya tidak
    masuk akal untuk dihitung (gap <= 0, tidak ada info ukuran font).
    SENGAJA True (bukan False) di semua kasus tidak-bisa-dipastikan
    ini -- supaya batas dokumen/halaman/data tidak lengkap tidak ikut
    salah tandai "seharusnya baris kosong tapi tidak ada" padahal itu
    cuma efek posisi/keterbatasan data, bukan format yang sungguhan
    salah. Konsisten dengan sikap modul ini di tempat lain (mis.
    _infer_alignment dkk) yang lebih memilih tidak menandai daripada
    salah tandai kalau estimasinya tidak bisa diandalkan.
    """
    if before is None or after is None:
        return True
    if before.get("page_index") != after.get("page_index"):
        return True

    gap_pt = after["bbox"][1] - before["bbox"][3]
    if gap_pt <= 0:
        # Overlap/negatif -- posisi tidak masuk akal untuk dihitung
        # (kemungkinan blok salah urut atau tumpang tindih); jangan
        # dipaksa "tidak ada baris kosong" dari data yang jelas aneh.
        return True

    sizes = [
        r["font_size_pt"]
        for elem in (before, after)
        for r in elem["runs"]
        if r["font_size_pt"]
    ]
    if not sizes:
        return True
    avg_size = statistics.median(sizes)
    if avg_size <= 0:
        return True

    ratio = gap_pt / (avg_size * 1.15)
    return ratio > _BLANK_LINE_GAP_RATIO_THRESHOLD


def _infer_equation_blank_lines(elements: list[dict]) -> None:
    """
    Mengisi blank_before/blank_after untuk elemen ptype "equation"
    secara IN-PLACE, berdasar jarak vertikal (bbox) terhadap elemen
    tetangganya di `elements` -- PDF tidak menyimpan "paragraf kosong"
    sebagai objek tersendiri seperti docx (yang baca langsung
    doc.paragraphs[i-1]/[i+1], lihat extractor.py), jadi keberadaan 1
    baris kosong di sini HARUS didekati dari seberapa lebar gap-nya
    dibanding gap antar-paragraf normal (tanpa baris kosong) -- lihat
    _has_blank_gap()/_BLANK_LINE_GAP_RATIO_THRESHOLD untuk detail
    rasio & alasan pemilihan ambangnya.

    DITAMBAHKAN 2026-09-20 (laporan user -- ditemukan dokumen
    mahasiswa nyata yang persamaannya TIDAK diberi baris kosong di
    atas/bawahnya, padahal cek ini sebelumnya SELALU dilewati untuk
    jalur PDF karena blank_before/blank_after belum pernah dihitung
    sama sekali, cuma placeholder False -- lihat comparator.py untuk
    sisi "if source == 'docx'" yang ikut dihapus).

    KETERBATASAN (baca sebelum mengandalkan ini sebagai sumber utama):
    ini estimasi statistik dari POSISI, bukan pembacaan struktur
    dokumen langsung seperti docx -- rentan meleset kalau mahasiswa
    memakai "space before/after paragraph" custom (bukan 0 default).
    Ini kali PERTAMA cek ini aktif untuk jalur PDF, jadi belum ada
    bukti empiris seberapa sering false-positive/false-negative-nya
    di lapangan -- kalau nanti muncul over-flagging massal (banyak
    equation yang di dokumen aslinya SUDAH punya baris kosong tapi
    tetap ke-flag), opsi mitigasinya: (a) naikkan
    _BLANK_LINE_GAP_RATIO_THRESHOLD dulu (jangan langsung matikan
    total), atau (b) kalau tetap tidak stabil, matikan lagi khusus
    source=="pdf" di comparator.py seperti pola cek lain yang sudah
    dimatikan (rata tengah caption, dst).
    """
    for i, elem in enumerate(elements):
        if elem["type"] != "equation":
            continue
        prev_elem = elements[i - 1] if i > 0 else None
        next_elem = elements[i + 1] if i < len(elements) - 1 else None
        elem["blank_before"] = _has_blank_gap(prev_elem, elem)
        elem["blank_after"] = _has_blank_gap(elem, next_elem)


def _infer_indent_cm(
    block: dict, left_margin_pt: float, right_margin_pt: float
) -> tuple[float | None, float | None]:
    """
    Indentasi kiri/kanan paragraf dalam cm, didekati dari selisih bbox
    blok terhadap batas kolom teks. None-kan (bukan 0.0) kalau selisih
    di bawah noise floor (~0.1cm) -- konsisten dengan semantik docx
    "tidak di-set" vs "di-set 0".
    """
    noise_floor_pt = 0.1 / PT_TO_CM
    bbox = block["bbox"]
    left_delta = bbox[0] - left_margin_pt
    right_delta = right_margin_pt - bbox[2]
    left_cm = left_delta * PT_TO_CM if left_delta > noise_floor_pt else None
    right_cm = right_delta * PT_TO_CM if right_delta > noise_floor_pt else None
    return left_cm, right_cm


def _infer_line_spacing(block: dict) -> float | None:
    """
    Rasio jarak-antar-baseline terhadap ukuran font. Cuma valid kalau
    blok >= 2 baris -- balik None kalau cuma 1 baris (tidak ada gap
    untuk dihitung, BUKAN berarti spacing-nya salah).

    PENTING (bug yang diperbaiki 2026-08-19): versi sebelumnya memakai
    top-of-bbox tiap baris (line["bbox"][1]) buat hitung gap. Itu BUKAN
    ukuran yang stabil -- tinggi bbox satu baris ikut naik-turun
    tergantung glyph yang ADA di baris itu (baris dengan huruf berekor
    seperti g/y/p naik sedikit lebih rendah, baris dengan aksen/huruf
    kapital semua naik lebih tinggi), jadi gap top-to-top antar baris
    bisa jauh meleset dari jarak baris sebenarnya (diuji empiris:
    paragraf 1.5 spasi yang SAMA bisa kebaca rasio 0.6x-1.45x tergantung
    isi barisnya). Baseline (line["spans"][0]["origin"][1] dari
    PyMuPDF) adalah titik jangkar vertikal yang presisi -- konsisten
    berapa pun bentuk glyph di baris tsb -- jadi dipakai di sini,
    bukan bbox.

    KALIBRASI KONSTANTA (diperbaiki lagi 2026-08-19, setelah fix merge
    paragraf di atas bikin bug ini KELIHATAN): konstanta 1.17 yang
    dipakai sebelumnya TERBUKTI sedikit kebesaran -- setelah paragraf
    yang tadinya kepecah jadi blok satu-baris berhasil digabung
    (_merge_wrapped_paragraph_blocks), puluhan paragraf yang jelas-
    jelas 1.5 spasi di dokumen asli malah konsisten kebaca 1.47-1.48,
    BUKAN acak/noise tapi BIAS SISTEMATIS ke satu arah. Dihitung balik
    dari situ (gap/size = hasil_lama * 1.17): rasio mentah gap/size
    ~1.72-1.73, dan 1.72/1.5 ~ 1.15 -- bukan 1.17. Konstanta diganti
    jadi 1.15.

    TOLERANSI (masalah kedua, terpisah dari kalibrasi): comparator.py
    membandingkan line_spacing ke RULES["line_spacing"] dengan toleransi
    CUMA 0.01 (abs(x - 1.5) > 0.01 -> gagal). Toleransi seketat itu
    masuk akal buat jalur DOCX (nilai persis dari XML, bukan estimasi),
    tapi buat jalur PDF nilai ini SELALU estimasi dari metrik render
    font -- noise pt-level dari hinting/DPI bisa geser hasil 0.01-0.03
    walau konstanta sudah benar. Daripada melonggarkan toleransi
    comparator.py (yang jadi kelonggaran untuk docx juga, padahal disitu
    memang harus presisi), noise itu diserap DI SINI: kalau rasio mentah
    dekat (toleransi 0.05) ke salah satu nilai lazim di Word (1.0, 1.15,
    2.0, 2.5, 3.0, dan 1.5 -- preset "Single/1.5 lines/Double" dst),
    di-snap PAS ke nilai itu. Paragraf yang beneran salah spasi (mis.
    hasil mentah 0.98x atau 1.3x, di luar jangkauan 0.05 dari semua
    preset) TETAP dilaporkan apa adanya, tidak ke-snap ke 1.5 -- jadi
    ini bukan melonggarkan deteksi, cuma membersihkan noise di sekitar
    nilai yang MEMANG dimaksud.

    CATATAN UNTUK VERIFIKASI LANJUTAN: konstanta 1.15 ini didapat dari
    reverse-engineering nilai 1.47/1.48 yang muncul di laporan real
    (bukan dari pengukuran font-metric langsung) -- cukup solid karena
    biasnya sistematis & sempit (1.47-1.48, bukan tersebar random), tapi
    kalau nanti dicoba ke PDF lain dan masih ada bias residual, cek dulu
    apa font_size yang kepakai (Arial vs TNR bisa beda dikit metriknya)
    sebelum utak-atik konstanta lagi.
    """
    baselines = block.get("line_baselines") or [
        bbox[1] for bbox in block["line_bboxes"]
    ]
    if len(baselines) < 2:
        return None
    sizes = [r["font_size_pt"] for r in block["runs"] if r["font_size_pt"]]
    if not sizes:
        return None
    avg_size = statistics.median(sizes)
    if avg_size <= 0:
        return None
    gaps = [baselines[i + 1] - baselines[i] for i in range(len(baselines) - 1)]
    avg_gap = statistics.median(gaps)
    # Baseline-to-baseline "single spacing" wajar ~1.15x font size
    # (line-height bawaan font, bukan tepat 1.0x) -- dibagi konstanta
    # ini supaya hasilnya kalibrasi ke skala "1.0 = single" yang sama
    # dengan RULES["line_spacing"] (1.5) di config.py.
    raw_ratio = avg_gap / (avg_size * 1.15)
    for preset in _LINE_SPACING_PRESETS:
        if abs(raw_ratio - preset) <= _LINE_SPACING_SNAP_TOL:
            return preset
    return round(raw_ratio, 2)


# Preset spasi baris lazim di Word (Single/1.5 lines/Double, dst) --
# lihat docstring _infer_line_spacing() bagian TOLERANSI.
_LINE_SPACING_PRESETS = (1.0, 1.15, 1.5, 2.0, 2.5, 3.0)
_LINE_SPACING_SNAP_TOL = 0.05


def _is_unnumbered_heading_runs(runs: list[dict], text: str) -> bool:
    """
    Padanan is_unnumbered_heading(para) di extractor.py, tapi berbasis
    `runs` (list dict generik) alih-alih para.runs docx -- prinsipnya
    identik: seluruh run berisi teks harus bold, teks pendek (<=8 kata),
    dan tidak diakhiri tanda baca kalimat.
    """
    text_runs = [r for r in runs if r["text"].strip()]
    if not text_runs:
        return False
    all_bold = all(r["bold"] for r in text_runs)
    is_short = len(text.split()) <= 8
    no_sentence_ending = not text.rstrip().endswith((".", ",", ";"))
    return all_bold and is_short and no_sentence_ending


# --------------------------------------------------------------------
# Section params (margin & ukuran halaman) dari geometri PDF
# --------------------------------------------------------------------


def extract_section_params_pdf(pdf_path: str) -> list[dict]:
    """
    Padanan extract_section_params(doc) versi docx, TAPI disederhanakan
    jadi SATU section untuk seluruh dokumen (lihat keterbatasan #6 di
    docstring modul). page_width/height diambil LANGSUNG dari
    page.rect -- ini EXACT (bukan inferensi) karena PDF menyimpan
    ukuran halaman literal, malah lebih presisi dari cm-rounding docx.

    BUG (diperbaiki 2026-08-19): versi sebelumnya menghitung margin
    kiri/kanan dari min(bbox[0])/max(bbox[2]) SEMUA blok di tiap
    halaman (lalu median antar-halaman) -- ini TERMASUK judul center
    (DAFTAR ISI, BAB N JUDUL), cover, caption, "Sumber:", dst yang
    MEMANG tidak menyentuh margin kiri/kanan asli (center/indent
    sengaja). Kalau di suatu halaman kebetulan tidak ada body text yang
    mepet margin asli (mis. halaman isinya cuma judul bab + gambar),
    margin halaman itu kebaca ngaco dan menarik median keseluruhan
    dokumen jadi salah. Dampaknya BUKAN cuma temuan margin sendiri
    (skor "Umum (Margin Halaman)" 0%) -- left_margin_pt/right_margin_pt
    yang salah ini dipakai juga oleh _infer_alignment() buat menghitung
    titik tengah kolom, jadi JUDUL YANG ASLINYA SUDAH CENTER (mis.
    "DAFTAR ISI", "BAB 1 PENDAHULUAN") ikut salah dilaporkan "seharusnya
    rata tengah" -- terverifikasi persis pola ini di laporan nyata,
    findings-nya SELALU 3 rangkap (font+bold+center) utk semua
    chapter_title, termasuk yang kemungkinan besar sudah benar di
    dokumen asli.

    Fix: filter kandidat margin ke paragraf BODY multi-baris SAJA
    (>=2 baris) -- heading/caption/cover/"Sumber:" hampir selalu cuma
    1 baris, jadi filter ini otomatis menyingkirkan mereka TANPA perlu
    klasifikasi penuh (classify_paragraph baru bisa jalan SETELAH
    margin section ada -- circular dependency kalau mau pakai
    ptype=="body" beneran, makanya dipakai proxy struktural "multi-
    baris" di sini). Konvensi PENS tidak pakai indentasi first-line,
    jadi SEMUA baris paragraf body (termasuk baris pertama) rata ke
    margin kiri -- dipakai buat estimasi kiri. Margin kanan dari SEMUA
    baris KECUALI baris terakhir tiap paragraf (baris terakhir wajar
    tidak penuh/tidak mepet margin kanan).

    BUG LANJUTAN (diperbaiki 2026-08-23, sebelumnya tercatat sebagai
    "margin kanan overcorrection" yang belum tuntas): fungsi ini
    memanggil `_extract_blocks(page)` MENTAH -- TANPA lewat
    `_merge_wrapped_paragraph_blocks()` -- padahal PyMuPDF sering
    memecah SATU paragraf body ber-spasi 1.5 jadi BEBERAPA blok
    terpisah (masing-masing seringnya cuma 1 baris) akibat heuristik
    clustering internalnya sendiri (root cause yang SAMA dengan yang
    dijelaskan panjang lebar di docstring `_merge_wrapped_paragraph_
    blocks`). Filter ">=2 baris" di atas otomatis MEMBUANG mayoritas
    blok pecahan ini (karena tiap pecahannya cuma 1 baris), sehingga
    kandidat margin kanan yang TERSISA didominasi sisa blok pendek/tak
    representatif (mis. potongan baris terakhir paragraf, sub-blok
    hasil pecahan yang kebetulan 2 baris tapi baris terakhirnya
    pendek) -- inilah kenapa median-nya bisa melenceng jauh (mis.
    terbaca ~14.9cm padahal margin asli ~3cm).

    Fix: jalankan `_merge_wrapped_paragraph_blocks()` DULU per halaman
    sebelum memfilter ">=2 baris", supaya paragraf body yang aslinya
    banyak baris tapi kepecah oleh PyMuPDF tergabung kembali dulu.
    Karena fungsi merge itu sendiri butuh estimasi margin kiri/kanan
    (buat cek "rata ke margin kiri" di syarat #2-nya) -- circular
    dependency yang sama seperti klasifikasi paragraf -- dipakai
    pendekatan DUA TAHAP:
      1. Kumpulkan estimasi KASAR (dari SEMUA blok mentah, tanpa
         filter jumlah baris) hanya untuk menyeed proses merge.
         Blok satu-baris (heading dst) tetap ikut di tahap ini, tapi
         karena cuma dipakai sebagai TITIK ACUAN merge (bukan hasil
         akhir), noise dari situ tidak fatal -- merge cuma butuh
         margin kiri yang "cukup dekat", bukan presisi.
      2. Jalankan merge per halaman pakai estimasi kasar itu, lalu
         hitung margin PRESISI dari blok hasil merge yang >=2 baris
         (logika sama seperti sebelumnya, cuma sumber bloknya beda).
    """
    with fitz.open(pdf_path) as pdf:
        if pdf.page_count == 0:
            raise ValueError("PDF tidak punya halaman sama sekali.")

        first_page = pdf[0]
        width_pt, height_pt = first_page.rect.width, first_page.rect.height

        # --- Tahap 1: estimasi KASAR margin kiri/kanan, cuma buat
        #     menyeed _merge_wrapped_paragraph_blocks() (lihat catatan
        #     di docstring). Dikumpulkan dari SEMUA blok mentah di
        #     SELURUH halaman (bukan per-halaman), lalu diambil MEDIAN --
        #     bukan min() seperti versi sebelumnya. min() per halaman
        #     rapuh terhadap SATU blok outlier (mis. tabel/gambar lebar
        #     penuh yang bbox-nya menclok sampai nyaris tepi fisik
        #     halaman) -- satu outlier begini cukup buat menarik seluruh
        #     estimasi kasar meleset jauh dari margin sebenarnya, padahal
        #     tujuannya cuma jadi "kira-kira" buat menyeed merge, jadi
        #     harus tahan-outlier. Median atas SEMUA blok (bukan cuma
        #     yang >=2 baris) tetap valid buat tujuan ini karena mayoritas
        #     blok di dokumen memang teks body yang rata margin kiri/
        #     kanan -- minoritas blok center/indent tidak cukup untuk
        #     menggeser median.
        rough_pages_blocks: list[list[dict]] = []
        rough_left: list[float] = []
        rough_right: list[float] = []
        for page in pdf:
            blocks = _extract_blocks(page)
            rough_pages_blocks.append(blocks)
            for b in blocks:
                rough_left.append(b["bbox"][0])
                rough_right.append(page.rect.width - b["bbox"][2])
        rough_left_pt = statistics.median(rough_left) if rough_left else 0.0
        rough_right_pt = statistics.median(rough_right) if rough_right else 0.0

        # --- Tahap 2: gabungkan blok yang terpecah pakai estimasi
        #     kasar di atas, baru filter ">=2 baris" & hitung margin
        #     presisi dari hasil gabungan itu.
        #
        # DITAMBAH (2026-09-07, atas permintaan user): kiri/kanan SEKARANG
        # JUGA dikumpulkan terpisah per PARITAS halaman fisik (ganjil vs
        # genap), bukan cuma digabung jadi satu median seperti sebelumnya.
        # Alasannya: kalau dokumen pakai mirror-margin, margin kiri
        # halaman ganjil & genap MEMANG seharusnya beda (mis. ganjil 4cm
        # sisi-dalam, genap 3cm sisi-luar) -- median gabungan keduanya
        # jadi angka "campuran" yang tidak merepresentasikan salah satu
        # aturan manapun (bukan 4, bukan juga 3), sehingga cek margin per
        # sisi jadi tidak valid utk dokumen mirror-margin. left_cm/
        # right_cm gabungan TETAP dihitung & dipertahankan apa adanya
        # (dipakai comparator.py utk dokumen yang TIDAK pakai mirror-
        # margin, di mana kiri/kanan memang seragam di semua halaman).
        # Paritas dipakai APA ADANYA dari nomor halaman FISIK (page_idx+1)
        # -- pendekatan sama seperti _infer_mirror_margins() di bawah,
        # BUKAN nomor halaman TERCETAK -- supaya konsisten dgn fungsi itu
        # dan tidak butuh data footer number (yang baru tersedia belakangan
        # dari page_layout_checker.py, sumber terpisah).
        left_candidates: list[float] = []
        right_candidates: list[float] = []
        left_odd_candidates: list[float] = []
        left_even_candidates: list[float] = []
        right_odd_candidates: list[float] = []
        right_even_candidates: list[float] = []
        tops, bottoms = [], []
        for page_idx, (page, raw_blocks) in enumerate(zip(pdf, rough_pages_blocks)):
            if not raw_blocks:
                continue
            blocks = _merge_wrapped_paragraph_blocks(
                raw_blocks, rough_left_pt, rough_right_pt
            )
            is_odd_page = (page_idx + 1) % 2 == 1
            for b in blocks:
                lines = b["line_bboxes"]
                if len(lines) < 2:
                    # Blok satu-baris (heading/caption/cover/"Sumber:")
                    # sering SENGAJA tidak mepet margin -- disingkirkan
                    # dari kandidat, lihat docstring di atas.
                    continue
                for l in lines:
                    left_candidates.append(l[0])
                    (
                        left_odd_candidates if is_odd_page else left_even_candidates
                    ).append(l[0])
                for l in lines[:-1]:
                    right_val = page.rect.width - l[2]
                    right_candidates.append(right_val)
                    (
                        right_odd_candidates if is_odd_page else right_even_candidates
                    ).append(right_val)
            tops.append(min(b["bbox"][1] for b in blocks))
            bottoms.append(page.rect.height - max(b["bbox"][3] for b in blocks))

        if not left_candidates:
            # Fallback kalau dokumen aneh (tidak ada body multi-baris
            # sama sekali, bahkan setelah merge) -- balik ke pendekatan
            # paling kasar biar tidak crash. Kandidat per-paritas SENGAJA
            # tidak direkonstruksi di jalur fallback ini (dokumen sudah
            # "aneh" & datanya minim) -- nanti otomatis jatuh ke fallback
            # gabungan di bawah (margin_left_pt/margin_right_pt).
            left_candidates, right_candidates = [], []
            for blocks in rough_pages_blocks:
                if not blocks:
                    continue
                left_candidates.append(min(b["bbox"][0] for b in blocks))
                right_candidates.append(min(width_pt - b["bbox"][2] for b in blocks))
        if not left_candidates:
            raise ValueError("Tidak ada teks badan terdeteksi di PDF ini sama sekali.")

        # CATATAN (2026-09-09): sempat dicoba pakai _robust_min (rata-
        # rata beberapa nilai TERKECIL) untuk margin kanan juga, dengan
        # alasan yang sama seperti margin atas/bawah di bawah. TERNYATA
        # SALAH di dokumen nyata -- REVERT, jangan diulangi lagi:
        # margin atas/bawah cuma punya SATU kandidat per HALAMAN (jumlah
        # kandidat kecil, rapuh terhadap satu outlier tapi outlier itu
        # sendiri jarang separah kandidat baris). Margin kanan pakai
        # kandidat per BARIS di SELURUH dokumen (bisa ribuan baris) --
        # di populasi sebesar itu HAMPIR PASTI ada segelintir baris yang
        # "kebetulan" jaraknya ke margin jauh lebih kecil dari margin
        # asli (kata/URL panjang yang mepet, sel tabel yang ke-anggap
        # baris body, dst). _robust_min (k terkecil) rapuh ketarik
        # segelintir kasus itu, BUKAN mendekati margin asli -- akibatnya
        # margin kanan malah UNDERESTIMATE parah (di data nyata:
        # terbaca ~1.4cm/2.4cm padahal margin asli 3-4cm), yang balas
        # menggeser col_center ke KANAN dan membuat heading yang
        # SEBENARNYA sudah center malah ikut ke-flag "belum rata
        # tengah" -- over-flagging, kebalikan dari bug yang mau
        # diperbaiki. Kembali ke median seperti semula: median atas
        # populasi sebesar itu jauh lebih tahan terhadap outlier
        # individual dibanding _robust_min(k=3).
        margin_left_pt = statistics.median(left_candidates)
        margin_right_pt = (
            statistics.median(right_candidates) if right_candidates else margin_left_pt
        )
        # Median per paritas -- fallback ke median GABUNGAN kalau salah
        # satu grup kosong (mis. dokumen halaman ganjil semua "aneh" &
        # tersaring habis oleh filter >=2 baris), supaya tetap ada angka
        # yang masuk akal dipakai, bukan None/crash.
        margin_left_odd_pt = (
            statistics.median(left_odd_candidates)
            if left_odd_candidates
            else margin_left_pt
        )
        margin_left_even_pt = (
            statistics.median(left_even_candidates)
            if left_even_candidates
            else margin_left_pt
        )
        margin_right_odd_pt = (
            statistics.median(right_odd_candidates)
            if right_odd_candidates
            else margin_right_pt
        )
        margin_right_even_pt = (
            statistics.median(right_even_candidates)
            if right_even_candidates
            else margin_right_pt
        )

        def _robust_min(values: list[float], k: int = 3) -> float:
            if not values:
                return 0.0
            s = sorted(values)
            k = min(k, len(s))
            return sum(s[:k]) / k

        # BUG DIPERBAIKI (2026-09-07, ditemukan dari perbandingan margin
        # bawah docx vs pdf pada dokumen yang SAMA: docx persis 2.50cm,
        # pdf lolos toleransi seolah-olah ~3cm): margin ATAS/BAWAH
        # sebelumnya diestimasi pakai MEDIAN dari `tops`/`bottoms`
        # (masing-masing sudah min-per-halaman). Median menyimpan bias
        # SATU ARAH untuk sisi bawah -- teks TIDAK PERNAH bisa dirender
        # melewati margin (jarak ke tepi kertas SELALU >= margin
        # sebenarnya), tapi HALAMAN YANG TIDAK PENUH (baris terakhir
        # berhenti sebelum menyentuh batas margin, wajar krn Word
        # memutus halaman begitu baris berikutnya TIDAK MUAT lagi)
        # membuat jarak TERUKUR ikut LEBIH BESAR dari margin sebenarnya.
        # Ini terjadi di HAMPIR SEMUA halaman (cuma halaman yang
        # kebetulan penuh pas sampai baris terakhir yang mendekati
        # margin asli), jadi median dari seluruh halaman condong
        # melebih-lebihkan margin bawah (dan, dengan alasan sama tapi
        # jauh lebih kecil dampaknya, margin atas -- karena halaman baru
        # SELALU mulai tepat setelah margin atas, jadi bias di sisi ini
        # jauh lebih kecil, tapi tetap ikut dibetulkan untuk konsisten).
        #
        # Fix: pakai estimator yang TAHAN-OUTLIER tapi condong ke NILAI
        # TERKECIL yang teramati (bukan median) -- rata-rata dari
        # beberapa nilai TERKECIL (bukan cuma satu min() mentah, supaya
        # tidak rapuh kalau kebetulan ada satu halaman dengan gangguan
        # deteksi/artefak). Di antara SEMUA halaman, halaman yang
        # kontennya paling PENUH (baris terakhir paling dekat ke margin
        # asli) akan menghasilkan nilai TERKECIL -- itulah pendekatan
        # terbaik yang tersedia untuk margin sebenarnya tanpa akses ke
        # w:pgMar langsung seperti docx.
        margin_top_pt = _robust_min(tops)
        margin_bottom_pt = _robust_min(bottoms)

    def twips(pt: float) -> int:
        return round(pt * 20)  # 1pt = 20 twips

    return [
        {
            "index": 0,
            "left_cm": margin_left_pt * PT_TO_CM,
            "left_twips": twips(margin_left_pt),
            "top_cm": margin_top_pt * PT_TO_CM,
            "top_twips": twips(margin_top_pt),
            "bottom_cm": margin_bottom_pt * PT_TO_CM,
            "bottom_twips": twips(margin_bottom_pt),
            "right_cm": margin_right_pt * PT_TO_CM,
            "right_twips": twips(margin_right_pt),
            # left_odd/even & right_odd/even: median per PARITAS halaman
            # fisik -- dipakai comparator.py KHUSUS utk memvalidasi
            # dokumen mirror-margin (kiri/kanan wajib berselang-seling
            # antar halaman ganjil/genap, bukan seragam). Dokumen yang
            # TIDAK mirror-margin cukup pakai left_cm/right_cm gabungan
            # di atas seperti sebelumnya -- field ini tidak mengubah
            # perilaku lama, cuma menambah data baru.
            "left_odd_cm": margin_left_odd_pt * PT_TO_CM,
            "left_odd_twips": twips(margin_left_odd_pt),
            "left_even_cm": margin_left_even_pt * PT_TO_CM,
            "left_even_twips": twips(margin_left_even_pt),
            "right_odd_cm": margin_right_odd_pt * PT_TO_CM,
            "right_odd_twips": twips(margin_right_odd_pt),
            "right_even_cm": margin_right_even_pt * PT_TO_CM,
            "right_even_twips": twips(margin_right_even_pt),
            "width_cm": width_pt * PT_TO_CM,
            "width_twips": twips(width_pt),
            "height_cm": height_pt * PT_TO_CM,
            "height_twips": twips(height_pt),
            # pgnum_fmt/start & footer_* TIDAK diisi di sini -- itu
            # sudah dihitung lebih akurat oleh extract_footer_numbers()
            # (baca angka BENERAN tercetak) di page_layout_checker.py,
            # yang dipanggil terpisah di main.py. Diisi None supaya
            # skema dict tetap sama, bukan dihitung dobel/tidak akurat.
            "pgnum_fmt": None,
            "pgnum_start": None,
            "footer_has_page_field": None,
            "footer_alignment": None,
        }
    ]


def _infer_mirror_margins(pdf_path: str) -> bool:
    """
    Padanan _get_mirror_margins(doc). Didekati dari POLA: margin kiri
    halaman ganjil vs genap berbeda signifikan (>0.3cm) DAN konsisten
    berselang-seling -- itu ciri cetak bolak-balik/mirror margin.
    Kalau dokumen cuma sedikit halaman atau tidak konsisten, default
    False (asumsi paling umum/aman).
    """
    tol_pt = 0.3 / PT_TO_CM
    odd_lefts, even_lefts = [], []
    with fitz.open(pdf_path) as pdf:
        for i, page in enumerate(pdf):
            blocks = _extract_blocks(page)
            if not blocks:
                continue
            left = min(b["bbox"][0] for b in blocks)
            physical_page_no = i + 1
            (odd_lefts if physical_page_no % 2 == 1 else even_lefts).append(left)

    if len(odd_lefts) < 3 or len(even_lefts) < 3:
        return False
    return abs(statistics.median(odd_lefts) - statistics.median(even_lefts)) > tol_pt


# --------------------------------------------------------------------
# Table params (best-effort, pakai fitur deteksi tabel bawaan PyMuPDF)
# --------------------------------------------------------------------


def _row_has_shading(page: "fitz.Page", row_bbox: tuple) -> bool:
    """
    Deteksi latar-belakang gelap baris header, dari VECTOR DRAWING
    (rect terisi warna) yang bbox-nya beririsan signifikan dengan
    row_bbox -- PDF tidak simpan "shading cell" sebagai atribut
    langsung kayak w:shd di docx, jadi ini didekati dari objek gambar
    (kotak warna) yang dirender DI BELAKANG teks header. BEST EFFORT:
    kalau gagal terbaca / rendernya bukan vector rect (mis. shading di-
    rasterize jadi gambar), balik False (dianggap tidak shading),
    bukan error -- konsisten dengan filosofi "gagal ke arah aman/
    presisi" yang sudah dipakai checker lain di kodebase ini.
    """
    try:
        drawings = page.get_drawings()
    except Exception:
        return False
    rx0, ry0, rx1, ry1 = row_bbox
    row_area = max(rx1 - rx0, 0) * max(ry1 - ry0, 0)
    if row_area <= 0:
        return False
    for d in drawings:
        fill = d.get("fill")
        if not fill:
            continue
        # Putih/nyaris putih dianggap "tidak ada shading" (kertas
        # polos), bukan warna latar sungguhan.
        if all(c >= 0.95 for c in fill[:3]):
            continue
        dx0, dy0, dx1, dy1 = d.get("rect", (0, 0, 0, 0))
        ix0, iy0 = max(rx0, dx0), max(ry0, dy0)
        ix1, iy1 = min(rx1, dx1), min(ry1, dy1)
        inter = max(ix1 - ix0, 0) * max(iy1 - iy0, 0)
        if inter / row_area >= 0.5:
            return True
    return False


def _rgb01_to_hex(rgb: tuple) -> str:
    """(r, g, b) float 0..1 (skala fitz) -> string hex 'RRGGBB' huruf besar."""
    r, g, b = (max(0, min(255, round(c * 255))) for c in rgb[:3])
    return f"{r:02X}{g:02X}{b:02X}"


def _get_cover_background_color_hex_pdf(page: "fitz.Page") -> str | None:
    """
    Padanan _get_document_background_color_hex() (extractor.py/docx) untuk
    jalur PDF. Dipakai comparator.py buat cek warna sampul "#002FA7" (RGB
    0, 47, 167) sesuai lembar aturan cover -- lihat COVER_RULES di
    comparator.py.

    docx punya elemen <w:background> eksplisit (Page Color), tapi PDF
    TIDAK PUNYA konsep "warna latar halaman" sebagai atribut -- setelah
    dirender/dikonversi (mis. lewat LibreOffice), Page Color ATAUPUN
    shape/rectangle full-page yang di-fill manual SAMA-SAMA cuma jadi
    objek visual biasa di halaman. Jadi warna latar di sini didekati
    dari dua sumber, dicoba berurutan:

    1. VECTOR RECTANGLE (page.get_drawings()): kalau warna sampul dibuat
       dari shape/rectangle vektor yang di-fill (termasuk hasil konversi
       Page Color Word -> PDF, yang lazimnya JUGA jadi rect vektor),
       cari rect berisi warna (fill, bukan cuma outline) dengan luas
       paling besar dan meliputi sebagian besar halaman (>= 60% luas
       halaman) -- itu nyaris pasti latar belakang, bukan elemen kecil
       seperti logo/kotak teks.
    2. FALLBACK RASTER (page.get_pixmap()): kalau sampul ditempel
       sebagai satu gambar utuh (mis. dari elemen background-image),
       tidak akan ada rect vektor sama sekali. Di sini halaman dirender
       ke bitmap resolusi rendah lalu di-sampling di titik-titik dekat
       tepi halaman (menghindari area tengah yang biasanya berisi teks
       putih) -- kalau MAYORITAS titik sampel sepakat pada satu warna
       (bukan putih/hampir putih kertas polos), warna itu yang dipakai.

    BEST EFFORT dengan filosofi "gagal ke arah aman": kalau kedua cara
    di atas tidak menemukan kandidat yang meyakinkan, balikin None --
    comparator.py WAJIB memperlakukan None sebagai "tidak bisa
    dipastikan otomatis", BUKAN "warna salah" (sampul yang visualnya
    sudah benar tidak boleh ikut disalahkan cuma karena caranya
    membuat warna latar tidak terbaca otomatis).
    """
    page_w, page_h = page.rect.width, page.rect.height
    page_area = page_w * page_h
    if page_area <= 0:
        return None

    # --- 1. vector rectangle terbesar yang isi warna ---
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []

    best_area = 0.0
    best_fill = None
    for d in drawings:
        fill = d.get("fill")
        if not fill:
            continue
        dx0, dy0, dx1, dy1 = d.get("rect", (0, 0, 0, 0))
        area = max(dx1 - dx0, 0) * max(dy1 - dy0, 0)
        if area / page_area < 0.6:
            continue
        if area > best_area:
            best_area = area
            best_fill = fill
    if best_fill is not None:
        return _rgb01_to_hex(best_fill)

    # --- 2. fallback: sampling piksel dari raster halaman ---
    # Zoom dinaikkan (1.5x, bukan 0.3x) supaya titik sampel jatuh di
    # tengah blok warna solid, bukan di piksel hasil downsampling yang
    # sudah tercampur (blur) dengan warna tetangga -- itu salah satu
    # penyebab hasil sampling sebelumnya sedikit meleset dari warna
    # aslinya (mis. "#002AA6" padahal sampul aslinya "#002FA7").
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    except Exception:
        return None
    if pix.n < 3:
        return None
    w, h = pix.width, pix.height
    if w == 0 or h == 0:
        return None
    # Titik-titik dekat tepi (8% dari lebar/tinggi, BUKAN 5%) -- cukup
    # jauh dari tepi fisik halaman (menghindari anti-aliasing di garis
    # batas antar elemen, mis. pita kuning di sampul PENS) tapi tetap
    # menghindari area tengah yang lazim berisi teks putih.
    margin_x, margin_y = max(int(w * 0.08), 1), max(int(h * 0.08), 1)
    sample_points = [
        (margin_x, margin_y),
        (w - margin_x, margin_y),
        (margin_x, h - margin_y),
        (w - margin_x, h - margin_y),
        (margin_x, h // 2),
        (w - margin_x, h // 2),
        (w // 2, margin_y),
        (w // 2, h - margin_y),
    ]

    def _patch_median(cx, cy):
        """Median warna dari patch 5x5 piksel di sekitar (cx, cy) --
        meredam noise satu piksel (anti-aliasing, artefak kompresi)
        dibanding baca satu piksel mentah-mentah."""
        rs, gs, bs = [], [], []
        for dx in range(-2, 3):
            for dy in range(-2, 3):
                x, y = min(max(cx + dx, 0), w - 1), min(max(cy + dy, 0), h - 1)
                px = pix.pixel(x, y)
                rs.append(px[0])
                gs.append(px[1])
                bs.append(px[2])
        return (
            statistics.median(rs),
            statistics.median(gs),
            statistics.median(bs),
        )

    counts = defaultdict(int)
    for x, y in sample_points:
        x, y = min(x, w - 1), min(y, h - 1)
        r, g, b = _patch_median(x, y)
        if r >= 245 and g >= 245 and b >= 245:
            continue  # kertas putih polos, bukan warna latar sungguhan
        counts[(r, g, b)] += 1
    if not counts:
        return None
    (r, g, b), n = max(counts.items(), key=lambda kv: kv[1])
    if n / len(sample_points) < 0.5:
        # Tidak ada warna yang dominan di mayoritas titik sampel --
        # kemungkinan besar bukan latar belakang penuh satu warna.
        return None
    return _rgb01_to_hex((r / 255, g / 255, b / 255))


def _detect_tables_on_page(page: "fitz.Page") -> list[dict]:
    """
    Deteksi tabel di SATU halaman pakai page.find_tables() bawaan
    PyMuPDF (heuristik garis/kolom) -- BEST EFFORT, tidak seakurat baca
    w:tbl langsung dari XML docx, terutama untuk tabel tanpa garis
    (borderless) yang cukup umum dipakai di format buku TA. Balik list
    dict MENTAH (belum ada prev_para_idx/next_para_idx/seq -- itu diisi
    belakangan di _extract_paragraphs_and_tables_pdf() yang tahu posisi
    tabel ini relatif ke paragraf sekitarnya), diurutkan dari atas ke
    bawah halaman (dibutuhkan supaya penggabungan dengan paragraf per
    halaman urut benar).
    """
    try:
        found = page.find_tables()
    except Exception:
        return []

    results = []
    for t in found.tables:
        bbox = t.bbox
        rows = list(t.rows) if t.rows else []
        n_rows = len(rows)

        header_format = None
        if n_rows > 1:
            first_row_bbox = rows[0].bbox
            spans_in_row = []
            text_dict = page.get_text("dict")
            for b in text_dict["blocks"]:
                if b.get("type") != 0:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sx0, sy0, sx1, sy1 = span["bbox"]
                        if (
                            sy0 >= first_row_bbox[1] - 1
                            and sy1 <= first_row_bbox[3] + 1
                            and sx0 >= first_row_bbox[0] - 1
                            and sx1 <= first_row_bbox[2] + 1
                            and span.get("text", "").strip()
                        ):
                            spans_in_row.append(span)
            if spans_in_row:
                row_center = (first_row_bbox[0] + first_row_bbox[2]) / 2
                text_x0 = min(s["bbox"][0] for s in spans_in_row)
                text_x1 = max(s["bbox"][2] for s in spans_in_row)
                text_center = (text_x0 + text_x1) / 2
                header_format = {
                    "header_all_bold": all(_span_bold(s) for s in spans_in_row),
                    "header_all_center": abs(text_center - row_center) <= 10,
                    "header_has_shading": _row_has_shading(page, first_row_bbox),
                }

        alignment = None
        page_center = page.rect.width / 2
        tbl_center = (bbox[0] + bbox[2]) / 2
        if abs(tbl_center - page_center) <= 15:
            alignment = "center"

        results.append(
            {
                "bbox": bbox,
                "n_rows": n_rows,
                "alignment": alignment,
                "header_format": header_format,
            }
        )

    results.sort(key=lambda t: t["bbox"][1])
    return results


def _find_prev_para_idx(
    paragraphs: list[dict], page_index: int, table_bbox: tuple
) -> int | None:
    """Index paragraf TERAKHIR yang posisinya di atas/sebelum tabel ini."""
    best = None
    for p in paragraphs:
        before = p["page_index"] < page_index or (
            p["page_index"] == page_index and p["bbox"][3] <= table_bbox[1] + 1
        )
        if before and (best is None or p["index"] > best):
            best = p["index"]
    return best


def _find_next_para_idx(
    paragraphs: list[dict], page_index: int, table_bbox: tuple
) -> int | None:
    """Index paragraf PERTAMA yang posisinya di bawah/setelah tabel ini."""
    best = None
    for p in paragraphs:
        after = p["page_index"] > page_index or (
            p["page_index"] == page_index and p["bbox"][1] >= table_bbox[3] - 1
        )
        if after and (best is None or p["index"] < best):
            best = p["index"]
    return best


def extract_table_params_pdf(pdf_path: str, paragraphs: list[dict]) -> list[dict]:
    """
    Padanan extract_table_params(doc) versi docx -- skema output SAMA
    PERSIS (seq/alignment/prev_para_idx/next_para_idx/n_rows/
    header_format) supaya compare_visual_objects() di comparator.py
    bisa dipakai tanpa perubahan sama sekali.

    WAJIB dipanggil SETELAH extract_paragraph_params_pdf() -- prev/next
    para idx dicari dari posisi (page_index + bbox) paragraf yang sudah
    diekstrak lebih dulu, bukan dihitung independen.
    """
    tables_data = []
    seq = 0
    with fitz.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf):
            for t in _detect_tables_on_page(page):
                seq += 1
                tables_data.append(
                    {
                        "seq": seq,
                        "alignment": t["alignment"],
                        "prev_para_idx": _find_prev_para_idx(
                            paragraphs, page_index, t["bbox"]
                        ),
                        "next_para_idx": _find_next_para_idx(
                            paragraphs, page_index, t["bbox"]
                        ),
                        "n_rows": t["n_rows"],
                        "header_format": t["header_format"],
                    }
                )
    return tables_data


# --------------------------------------------------------------------
# Paragraph params -- state machine chapter tracking (DIPORTING dari
# extract_paragraph_params() di extractor.py, bagian yang murni
# berbasis TEKS dipertahankan APA ADANYA; bagian yang butuh objek docx
# diganti versi PDF-nya, lihat keterbatasan #1-#2 di docstring modul)
# --------------------------------------------------------------------


def extract_paragraph_params_pdf(pdf_path: str) -> list[dict]:
    sections = extract_section_params_pdf(pdf_path)
    left_margin_pt = sections[0]["left_twips"] / 20
    right_margin_pt = sections[0]["width_twips"] / 20 - sections[0]["right_twips"] / 20

    # BUG DIPERBAIKI (2026-09-07 -- laporan user: judul "BAB 1" yang di
    # dokumen aslinya SUDAH rata tengah tetap ditandai "seharusnya rata
    # tengah"): dokumen mirror-margin (cetak bolak-balik) punya kolom
    # teks yang BERGESER tiap ganti paritas halaman (mis. ganjil kiri
    # 4.0cm/kanan 2.9cm, genap kiri 3.0cm/kanan 3.9cm -- lihat
    # left_odd/left_even/right_odd/right_even di
    # extract_section_params_pdf()). Sebelumnya left_margin_pt/
    # right_margin_pt GENERIK di atas (median GABUNGAN semua halaman,
    # tidak dipisah paritas) dipakai untuk _infer_alignment()/
    # _infer_equation_alignment()/_infer_indent_cm() di SEMUA halaman
    # tanpa kecuali, termasuk dokumen mirror-margin. Akibatnya titik
    # tengah kolom yang dihitung meleset ~0.5cm untuk separuh halaman
    # (paritas yang berbeda dari campuran generik itu) -- verifikasi
    # empiris pada laporan nyata: judul "BAB 1" (halaman ganjil) meleset
    # ~21pt dari toleransi 12.75pt kalau dipakaikan margin generik,
    # padahal cuma ~7pt (LOLOS) kalau dipakaikan margin ganjil yang
    # benar. Fix: kalau dokumen terdeteksi mirror-margin, pakai margin
    # PER-PARITAS halaman fisik (gantian tiap iterasi berdasar
    # is_odd_page) -- kalau bukan mirror-margin, tetap pakai
    # left_margin_pt/right_margin_pt gabungan seperti sebelumnya (fallback
    # eksplisit, bukan diam-diam berubah perilaku utk dokumen yang sudah
    # benar).
    mirror_margins = _infer_mirror_margins(pdf_path)
    left_margin_odd_pt = sections[0]["left_odd_twips"] / 20
    left_margin_even_pt = sections[0]["left_even_twips"] / 20
    right_margin_odd_pt = (
        sections[0]["width_twips"] / 20 - sections[0]["right_odd_twips"] / 20
    )
    right_margin_even_pt = (
        sections[0]["width_twips"] / 20 - sections[0]["right_even_twips"] / 20
    )

    elements = []

    current_chapter = "Pra-Bab"
    expect_chapter_name = False
    # Padanan expect_lampiran_name di extractor.py -- lihat komentar
    # deklarasinya di sana untuk alasan lengkap (nama lampiran tidak
    # wajib ALL CAPS per pedoman PENS, beda dari nama BAB).
    expect_lampiran_name = False
    in_figure_table_list = False
    in_toc = False
    seen_bab = False
    abstract_state = None
    in_special_page = True

    idx = 0
    with fitz.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf):
            # Margin efektif utk halaman ini -- lihat catatan bug di atas.
            is_odd_page = (page_index + 1) % 2 == 1
            if mirror_margins:
                page_left_pt = (
                    left_margin_odd_pt if is_odd_page else left_margin_even_pt
                )
                page_right_pt = (
                    right_margin_odd_pt if is_odd_page else right_margin_even_pt
                )
            else:
                page_left_pt = left_margin_pt
                page_right_pt = right_margin_pt

            table_bboxes = [t["bbox"] for t in _detect_tables_on_page(page)]
            blocks = _extract_blocks(page, table_bboxes)
            blocks = _merge_wrapped_paragraph_blocks(
                blocks, page_left_pt, page_right_pt
            )
            for block in blocks:
                stripped = block["full_text"].strip()
                if not stripped:
                    continue

                ptype = classify_paragraph(stripped, para=None)
                upper_now = stripped.upper()

                if upper_now in NORMAL_FRONT_MATTER_MARKERS:
                    in_special_page = False
                elif (
                    upper_now in SPECIAL_FRONT_PAGE_MARKERS
                    and not in_toc
                    and not in_figure_table_list
                ):
                    in_special_page = True

                if BLANK_PAGE_NOTICE_RE.match(stripped):
                    ptype = "notice"
                    expect_chapter_name = False
                elif in_special_page and not seen_bab:
                    ptype = "cover"
                    expect_chapter_name = False
                else:
                    upper_stripped = stripped.upper()
                    # BUG UTAMA (fixed): classify_paragraph() murni
                    # berbasis string, jadi TIDAK TAHU dia lagi ada di
                    # dalam region Daftar Isi/Daftar Gambar/Daftar
                    # Tabel atau bukan. Entri TOC yang teksnya "BAB 1
                    # PENDAHULUAN .......... 5" tetap match
                    # CHAPTER_TITLE_RE sama seperti heading ASLI "BAB 1
                    # PENDAHULUAN" -- versi lama percaya begitu saja ke
                    # ptype=="chapter_title" dan MENIMPA current_chapter/
                    # seen_bab/in_toc di baris TOC itu sendiri, padahal
                    # ini bukan heading sungguhan. Akibatnya domino:
                    # in_toc keburu di-reset False di entri TOC pertama
                    # yang match, jadi SEMUA entri TOC berikutnya
                    # (subbab, DAFTAR GAMBAR, dst) ikut lolos jadi
                    # heading beneran juga, minta font 14pt/bold/center
                    # yang seharusnya tidak berlaku buat baris daftar
                    # isi. Fix: kalau lagi di dalam TOC/daftar
                    # gambar/tabel DAN baris ini masih berciri entri
                    # (dot-leader di ujung -- lihat _TOC_LEADER_RE -- atau
                    # classify_paragraph() TIDAK bilang chapter_title
                    # sama sekali), demosi paksa ke "toc_entry" TANPA
                    # menyentuh current_chapter/seen_bab/abstract_state/
                    # expect_chapter_name apapun. Heading ASLI (tanpa
                    # dot-leader) tetap lolos ke cabang normal di bawah,
                    # termasuk yang menutup region TOC (bukan didemosi).
                    is_toc_leader_line = bool(_TOC_LEADER_RE.search(stripped))
                    if (in_toc or in_figure_table_list) and (
                        is_toc_leader_line or ptype != "chapter_title"
                    ):
                        ptype = "toc_entry"
                    elif is_toc_leader_line and not seen_bab:
                        # BUG DIPERBAIKI (jalur PDF -- entri "Daftar
                        # Gambar"/"Daftar Tabel" individual ke-flag salah
                        # ukuran font, mis. "Ukuran font 12.0pt tidak
                        # sesuai untuk tipe 'caption', seharusnya 11pt"
                        # padahal baris itu SAMA SEKALI BUKAN caption
                        # gambar -- itu entri daftar isi di halaman
                        # "Daftar Gambar" yang isinya memang body text
                        # 12pt, bukan caption 11pt): in_toc/
                        # in_figure_table_list di atas cuma di-set True
                        # lewat heading literal "DAFTAR ISI"/"DAFTAR
                        # GAMBAR"/"DAFTAR TABEL", dan gampang ke-reset
                        # False lagi kalau ADA baris di antaranya yang
                        # (keliru) lolos classify_paragraph() sebagai
                        # ptype=="chapter_title" tanpa dot-leader --
                        # domino yang sama seperti bug "BUG UTAMA" di atas,
                        # cuma triggernya baris lain di tengah daftar,
                        # bukan baris pembuka daftarnya sendiri. Daripada
                        # terus menerus tambal state tracking yang rapuh
                        # ini, dipasang jaring pengaman independen: baris
                        # apa pun yang berpola dot-leader + nomor halaman
                        # di ujung (is_toc_leader_line) DAN muncul SEBELUM
                        # heading "BAB 1 ..." pertama (not seen_bab) HAMPIR
                        # PASTI entri Daftar Isi/Daftar Gambar/Daftar Tabel
                        # -- pola ini tidak pernah muncul di body text
                        # sungguhan pada bagian depan buku (cover,
                        # abstrak, dst.), jadi aman didemosi ke "toc_entry"
                        # tanpa bergantung sama sekali ke status
                        # in_toc/in_figure_table_list.
                        ptype = "toc_entry"
                    elif ptype == "chapter_title":
                        if upper_stripped == "ABSTRAK":
                            abstract_state = "id"
                        elif upper_stripped == "ABSTRACT":
                            abstract_state = "en"
                        else:
                            abstract_state = None

                        if (
                            CHAPTER_TITLE_RE.match(upper_stripped)
                            and stripped == upper_stripped
                        ):
                            current_chapter = re.match(
                                r"^BAB\s+\d+", upper_stripped
                            ).group(0)
                            seen_bab = True
                            expect_chapter_name = bool(
                                CHAPTER_TITLE_ONLY_RE.match(upper_stripped)
                            )
                        elif (
                            LAMPIRAN_TITLE_RE.match(upper_stripped)
                            and stripped == upper_stripped
                        ):
                            # Padanan cabang LAMPIRAN di extractor.py --
                            # current_chapter dinormalisasi ke "LAMPIRAN
                            # <huruf>" murni (bukan ikut cabang generik
                            # "elif seen_bab" di bawah yang akan
                            # menyimpan teks lengkap apa adanya, mis.
                            # "LAMPIRAN B DOKUMENTASI EKSPERIMEN" kalau
                            # nama lampirannya menyatu satu baris) --
                            # supaya caption "Tabel A.1"/"Gambar B.1" di
                            # dalamnya cocok dengan chapter yang dicek
                            # compare_caption_numbering() di
                            # comparator.py.
                            current_chapter = re.match(
                                r"^LAMPIRAN(?:\s+[A-Z]+)?", upper_stripped
                            ).group(0)
                            seen_bab = True
                            # PAKAI expect_lampiran_name (bukan
                            # expect_chapter_name) -- lihat komentar
                            # deklarasinya di atas.
                            expect_chapter_name = False
                            expect_lampiran_name = bool(
                                LAMPIRAN_TITLE_ONLY_RE.match(upper_stripped)
                            )
                        elif seen_bab:
                            # NOTE (keterbatasan #1): cabang "heading
                            # Heading-1-style tanpa teks BAB N literal"
                            # (jalur bab_counter di extractor.py) TIDAK
                            # ada padanannya di sini -- info numbering
                            # Word sudah hilang begitu jadi PDF. Kalau
                            # dokumen konsisten menulis "BAB N" literal
                            # (mayoritas kasus), baris ini tidak pernah
                            # kepakai.
                            current_chapter = upper_stripped
                            expect_chapter_name = False
                        else:
                            expect_chapter_name = False

                        in_figure_table_list = upper_stripped in (
                            "DAFTAR GAMBAR",
                            "DAFTAR TABEL",
                        )
                        in_toc = upper_stripped == "DAFTAR ISI"
                    elif expect_chapter_name:
                        ptype = "chapter_title"
                        expect_chapter_name = False
                    elif expect_lampiran_name:
                        # Padanan expect_chapter_name di atas -- lihat
                        # komentar lampiran_title di extractor.py.
                        ptype = "lampiran_title"
                        expect_lampiran_name = False
                    elif in_toc or in_figure_table_list:
                        # Jaga-jaga: entri TOC yang classify_paragraph()
                        # sudah benar (subchapter/subsubchapter/body,
                        # ptype != "chapter_title") TAPI entah kenapa
                        # tidak match _TOC_LEADER_RE (mis. dot-leader
                        # tidak ke-render sebagai karakter di beberapa
                        # PDF) -- tetap didemosi di sini sebagai
                        # fallback kedua, konsisten dengan perilaku lama.
                        ptype = "toc_entry"
                    elif ptype == "body" and _is_unnumbered_heading_runs(
                        block["runs"], stripped
                    ):
                        # Padanan Jalur 3 classify_paragraph() (heading
                        # tanpa nomor) -- di extractor.py ini dicoba DI
                        # DALAM classify_paragraph() lewat objek para;
                        # di sini classify_paragraph() dipanggil dengan
                        # para=None (jalur itu otomatis skip), jadi
                        # ditambal manual di sini pakai runs PDF.
                        ptype = "subsubchapter_title"

                if ptype != "cover":
                    if any(r.get("is_math") for r in block["runs"]):
                        # Proxy has_equation -- BUKAN deteksi OMML asli
                        # (lihat keterbatasan #3), cuma menandai run
                        # matematis di level PARAGRAF kalau MAYORITAS
                        # blok ini memang simbol matematis (bukan cuma
                        # satu variabel di tengah kalimat body).
                        #
                        # BUG DIPERBAIKI (2026-09-21, laporan mahasiswa
                        # -- baris keterangan/jadwal biasa yang cuma
                        # menyisipkan SATU simbol/karakter matematis di
                        # tengah kalimat body panjang, mis. "= nilai ∂
                        # terbesar dari seluruh valley yang ada" (∂ di
                        # tengah, sisanya 45+ karakter narasi biasa),
                        # ke-flag ptype "equation" utuh: versi lama
                        # menghitung math_chars dengan menjumlah
                        # PANJANG PENUH tiap run yang is_math-nya True
                        # ("all-or-nothing" per run) -- is_math sendiri
                        # cuma butuh SATU karakter matematis ADA DI
                        # MANA SAJA dalam run itu (lihat
                        # _is_math_run_text() di extractor.py), jadi
                        # kalau PyMuPDF kebetulan menggabungkan satu
                        # simbol matematis dengan puluhan karakter
                        # kalimat body biasa jadi SATU span/run (font
                        # sama, tidak ada alasan PyMuPDF memecahnya),
                        # seluruh panjang run itu ikut kehitung sebagai
                        # "math", membuat rasio meleset jauh ke atas.
                        # Ini beda dari kasus run BERFONT matematis
                        # (mis. Cambria Math) -- run seperti itu memang
                        # SUNGGUHAN seluruhnya notasi matematis (Word
                        # tidak pernah mencampur body text biasa ke
                        # dalam satu run berfont Cambria Math), jadi
                        # tetap wajar dihitung penuh.
                        #
                        # Fix: run berfont matematis (font_name normal
                        # ada di MATH_FONT_NAMES) tetap dihitung PENUH,
                        # tapi run BUKAN font matematis yang cuma
                        # ke-flag is_math via karakter Unicode di
                        # tengah teks (fallback _is_math_run_text)
                        # HANYA dihitung sepanjang KARAKTER matematis
                        # yang benar-benar cocok regex-nya
                        # (_MATH_ALPHANUMERIC_RE), bukan seluruh run.
                        math_chars = 0
                        for r in block["runs"]:
                            if not r.get("is_math"):
                                continue
                            font_name = (r.get("font_name") or "").strip().lower()
                            if font_name in {n.lower() for n in MATH_FONT_NAMES}:
                                math_chars += len(r["text"])
                            else:
                                math_chars += sum(
                                    len(m)
                                    for m in _MATH_ALPHANUMERIC_RE.findall(r["text"])
                                )
                        if math_chars / max(len(stripped), 1) > 0.5:
                            ptype = "equation"
                    if ptype != "equation" and SOURCE_CITATION_RE.match(stripped):
                        ptype = "source_citation"

                runs_data = [
                    {
                        "text": r["text"],
                        "font_name": r["font_name"],
                        "font_size_pt": r["font_size_pt"],
                        "bold": r["bold"],
                        "italic": r["italic"],
                        "is_math": r.get("is_math", False),
                    }
                    for r in block["runs"]
                ]

                abstract_part = None
                if abstract_state is not None and ptype == "body":
                    is_keyword_line = bool(ABSTRACT_KEYWORD_RE.match(stripped))
                    if abstract_state == "id":
                        abstract_part = (
                            "keyword_id" if is_keyword_line else "abstrak_body"
                        )
                    else:
                        abstract_part = (
                            "keyword_en" if is_keyword_line else "abstract_body"
                        )

                left_indent_cm, right_indent_cm = _infer_indent_cm(
                    block, page_left_pt, page_right_pt
                )

                elements.append(
                    {
                        "index": idx,
                        "page_index": page_index,
                        "bbox": block["bbox"],
                        "type": ptype,
                        "full_text": stripped,
                        "is_upper": stripped == stripped.upper(),
                        "is_titlecase": is_titlecase(stripped),
                        "line_spacing": _infer_line_spacing(block),
                        # PDF tidak punya konsep "rule" (EXACTLY/AT_LEAST
                        # vs MULTIPLE) seperti docx -- line_spacing di
                        # atas SELALU angka kelipatan hasil estimasi.
                        "line_spacing_is_multiple": True,
                        "line_spacing_rule": None,
                        "runs": runs_data,
                        "chapter": current_chapter,
                        "section": 0,
                        "abstract_part": abstract_part,
                        "alignment": (
                            _infer_equation_alignment(
                                block,
                                page_left_pt,
                                page_right_pt,
                                (left_margin_odd_pt, left_margin_even_pt),
                            )
                            if ptype == "equation"
                            else _infer_alignment(block, page_left_pt, page_right_pt)
                        ),
                        "left_indent_cm": left_indent_cm,
                        "right_indent_cm": right_indent_cm,
                        # blank_before/blank_after: placeholder False di
                        # sini utk SEMUA tipe elemen (PDF tidak punya
                        # objek "paragraf kosong" tersendiri seperti
                        # docx yang baca langsung doc.paragraphs[i-1]/
                        # [i+1]) -- KHUSUS ptype "equation", nilai ini
                        # ditimpa (diestimasi dari jarak vertikal antar
                        # elemen) oleh _infer_equation_blank_lines() di
                        # akhir fungsi ini, lihat docstring-nya untuk
                        # detail & keterbatasan. Tipe elemen lain (mis.
                        # gambar) TIDAK ikut ditimpa -- cek "1 baris
                        # kosong" comparator.py cuma dipakai utk
                        # equation saat ini, jadi belum diperlukan.
                        "blank_before": False,
                        "blank_after": False,
                        "equation_number_visual": (
                            _has_visual_equation_number(block, page_right_pt)
                            if ptype == "equation"
                            else False
                        ),
                    }
                )
                idx += 1

    # Perbaiki equation yang nomor identitasnya kepisah jadi blok
    # PyMuPDF tersendiri -- lihat docstring _merge_split_equation_numbers()
    # untuk alasan lengkap kenapa ini diperlukan KHUSUS jalur PDF.
    elements = _merge_equation_fragments(elements)
    for _e in elements:
        if _e["type"] == "equation":
            _move_equation_number_to_end(_e)
    elements = _merge_split_equation_numbers(elements)

    # Estimasi blank_before/blank_after utk equation dari jarak
    # vertikal antar elemen -- WAJIB dijalankan SETELAH merge di atas
    # (bbox equation yang sudah gabung nomor identitasnya harus jadi
    # acuan gap, bukan bbox blok nomor yang sudah tidak ada lagi di
    # `elements`). Lihat docstring _infer_equation_blank_lines().
    _infer_equation_blank_lines(elements)

    return elements


def _get_cover_background_hex_pdf(pdf_path: str) -> str | None:
    """
    Buka ulang halaman PERTAMA dokumen (halaman sampul/judul, sesuai
    urutan tetap template PENS -- selalu halaman pertama) khusus untuk
    deteksi warna latar. Sengaja dipisah dari
    extract_paragraph_params_pdf() (bukan digabung ke loop utama di
    sana) supaya BENTUK RETURN fungsi itu tidak berubah -- menghindari
    pecah di pemanggil lain yang mungkin masih bergantung ke bentuk
    lama (list[dict] polos). Biayanya cuma buka 1 halaman tambahan,
    murah dibanding manfaat isolasi ini.
    """
    try:
        with fitz.open(pdf_path) as pdf:
            if len(pdf) == 0:
                return None
            return _get_cover_background_color_hex_pdf(pdf[0])
    except Exception:
        return None


def extract_all_pdf(pdf_path: str) -> dict:
    """
    Entry point tahap Extract Parameter versi PDF -- padanan extract_all(doc)
    di extractor.py. Bentuk return SAMA PERSIS, supaya main.py &
    comparator.py bisa pakai salah satu (docx atau PDF) tanpa tahu
    bedanya.
    """
    paragraphs = extract_paragraph_params_pdf(pdf_path)
    sections = extract_section_params_pdf(pdf_path)
    tables = extract_table_params_pdf(pdf_path, paragraphs)
    mirror_margins = _infer_mirror_margins(pdf_path)
    return {
        "paragraphs": paragraphs,
        "sections": sections,
        "tables": tables,
        "mirror_margins": mirror_margins,
        # Sebelumnya SELALU None untuk jalur PDF (belum diimplementasikan
        # sama sekali) -- comparator.py jadi selalu melewatkan cek warna
        # sampul "#002FA7" untuk dokumen PDF. Sekarang didekati lewat
        # _get_cover_background_hex_pdf() (vector rect / sampling piksel
        # -- lihat docstring _get_cover_background_color_hex_pdf() untuk
        # detail & keterbatasannya). Tetap None kalau tidak terdeteksi;
        # comparator.py memperlakukan None sebagai "tidak bisa
        # dipastikan", BUKAN "warna salah".
        "cover_background_hex": _get_cover_background_hex_pdf(pdf_path),
    }
