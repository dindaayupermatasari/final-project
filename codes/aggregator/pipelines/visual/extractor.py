"""
extractor.py
============
TAHAP 2: EXTRACT PARAMETER
---------------------------
Mengubah objek Document mentah menjadi kumpulan parameter terstruktur
(font, size, bold, spacing, margin, dsb) PER ELEMEN. Modul ini TIDAK
tahu apa itu "benar" atau "salah" — cuma menangkap fakta mentah dari
dokumen. Perbandingan terhadap rule sepenuhnya jadi tugas comparator.py.

CHANGELOG (revisi terbaru):
- SUBCHAPTER_L1_RE/L2_RE (hardcode 2 & 3 level angka) DIHAPUS, diganti
  SUBCHAPTER_NUMBER_RE generik yang menangkap SEMUA level angka
  bertingkat ("1.1", "2.2.1", "3.2.2.1", dst) sekaligus menghitung
  kedalamannya. Sebelumnya "3.2.2.1" (4 level) tidak pernah match sama
  sekali karena L2_RE cuma menerima persis 3 level.
- Ditambah jalur deteksi KEDUA yang independen dari teks: nomor hasil
  automatic numbering Word (w:numPr/w:ilvl) TIDAK PERNAH masuk ke
  para.text -- ini kenapa subbab yang angkanya digenerate otomatis oleh
  Word (bukan diketik manual) sebelumnya gagal terdeteksi sama sekali
  dan jatuh ke default "body". Fungsi _get_numbering_ilvl() membaca
  ilvl langsung dari XML paragraf (atau dari style kalau numPr
  didefinisikan di style, bukan di paragraf).
  PENTING: ilvl TIDAK bergantung pada nama Style paragraf ("Heading
  2", "Normal", dst) sama sekali -- jadi tetap terdeteksi walau
  mahasiswa tidak pakai built-in heading style Word.
- Ditambah jalur deteksi KETIGA: heading level di bawah "X.Y.Z" yang
  memang TIDAK dinomori sama sekali oleh panduan (mis. "Ekstraksi
  Parameter", "Rule-Based Comparator"). Karena tidak ada angka untuk
  dicari sama sekali (bukan tersembunyi, tapi memang tidak ada),
  deteksi pakai ciri visual: bold penuh, pendek, tidak diakhiri tanda
  baca kalimat (is_unnumbered_heading). Level ini disatukan ke ptype
  "subsubchapter_title" karena aturan formatnya SAMA PERSIS dengan
  level "X.Y.Z": wajib Title Case + bold (lihat comparator.py).
- Ditambah _resolve_effective_font_size() dan
  _resolve_effective_font_name(), simetris dengan
  _resolve_effective_bold() yang sudah ada sebelumnya. Sebelumnya kalau
  font size/name di-set lewat style (bukan langsung di run), nilainya
  None dan comparator men-skip pengecekan itu sepenuhnya -- bukan
  dianggap salah, tapi juga tidak pernah dicek. Sekarang ditelusuri
  sampai ke base_style seperti bold.
- is_titlecase() TIDAK berubah -- tetap dipakai apa adanya untuk
  subsubchapter_title, termasuk untuk heading tanpa nomor (karena
  fungsinya sudah aman menoleransi teks tanpa prefix angka).
- Ditambah _resolve_effective_italic(), simetris dengan
  _resolve_effective_bold(). Dipakai foreign_term_checker.py /
  comparator.py untuk cek "istilah asing wajib italic" -- run yang
  italic-nya None (warisan style, bukan di-set langsung) sebelumnya
  akan salah dianggap "tidak italic" kalau dibaca run.italic mentah.
- Ditambah _get_right_indent_cm(), simetris dengan _get_left_indent_cm()
  yang sudah ada. left_indent_cm/right_indent_cm SEKARANG diekstrak
  untuk SEMUA paragraf (sebelumnya cuma untuk ptype "equation") --
  dibutuhkan quotation_checker.py/compare_quotations() di comparator.py
  untuk cek indentasi kiri & kanan kutipan langsung PANJANG, yang
  ptype-nya tetap "body" biasa.
"""

import re
from docx.oxml.ns import qn
from docx.shared import Length

CHAPTER_TITLE_ONLY_RE = re.compile(r"^BAB\s+\d+$")
# Menerima DUA format judul bab: "BAB 1" berdiri sendiri (nama bab di
# paragraf terpisah setelahnya) ATAU "BAB 1 PENDAHULUAN" digabung satu
# paragraf. Sebelumnya regex ini HANYA menerima format pertama --
# dokumen yang menulis "BAB N NAMA" dalam satu baris (format yang sama
# validnya) gagal dikenali sebagai chapter_title sama sekali, jatuh ke
# fallback is_unnumbered_heading() dan salah diklasifikasi sebagai
# subsubchapter_title (expected size 12pt, bukan 14pt bold; expected
# Title Case, bukan ALL CAPS). Efek berantainya lebih parah: karena
# ptype-nya bukan "chapter_title", current_chapter TIDAK PERNAH pindah
# dari bab sebelumnya -- seluruh isi bab tsb ikut numpuk ke bab
# sebelumnya di laporan.
CHAPTER_TITLE_RE = re.compile(r"^BAB\s+\d+(\s+\S.*)?$")

# Label nomor persamaan berdiri sendiri, mis. "(2.1)" di akhir baris
# rumus matematika. Pendek + bold -> kalau tidak dikecualikan di sini,
# ke-tangkap is_unnumbered_heading() sebagai heading tanpa nomor.
EQUATION_NUMBER_RE = re.compile(r"^\(\d+(\.\d+)*\)$")

# Marker teks untuk halaman-halaman "template tetap" di bagian depan
# dokumen (sampul, halaman judul dalam, halaman pengesahan, surat
# pernyataan keaslian) -- font/ukuran/posisi di halaman-halaman ini
# memang bebas sesuai template resmi, BUKAN mengikuti aturan body/
# judul bab biasa. Dipakai gantiin asumsi lama "section index == 0",
# yang gagal begitu mahasiswa punya halaman kosong/section break
# tambahan (mis. untuk cetak bolak-balik) sebelum cover benar-benar
# muncul.
SPECIAL_FRONT_PAGE_MARKERS = {
    "PROPOSAL PROYEK AKHIR",
    "LAPORAN PROYEK AKHIR",
    "TUGAS AKHIR",
    "HALAMAN SAMPUL",
    "HALAMAN JUDUL",
    "HALAMAN PENGESAHAN",
    "SURAT PERNYATAAN KEASLIAN",
}

# Marker yang menandakan kita SUDAH KELUAR dari halaman template tetap
# di atas, masuk ke front matter yang justru WAJIB ikut aturan body
# normal (font 12pt TNR, spasi 1.5, dst).
NORMAL_FRONT_MATTER_MARKERS = {
    "KATA PENGANTAR",
    "ABSTRAK",
    "ABSTRACT",
    "DAFTAR ISI",
    "DAFTAR GAMBAR",
    "DAFTAR TABEL",
    "DAFTAR LAMPIRAN",
}

# Generik: menangkap SEMUA level angka bertingkat manual di teks
# ("1.1 ...", "2.2.1 ...", "3.2.2.1 ...", dst), tidak dihardcode ke
# jumlah level tertentu seperti versi lama (SUBCHAPTER_L1_RE/L2_RE).
SUBCHAPTER_NUMBER_RE = re.compile(r"^((?:\d+\.)+\d+)\s+\S+")

# Baris "Kata Kunci: ..." / "Keywords: ..." di dalam Abstrak/Abstract --
# dipakai buat MEMISAHKAN baris keyword dari baris body abstrak saat
# nge-tag abstract_part di extract_paragraph_params(). Sengaja dibuat
# longgar (boleh tanpa titik dua) karena mahasiswa suka nulis
# "Kata Kunci - ..." atau "Kata Kunci ...\t..." juga.
ABSTRACT_KEYWORD_RE = re.compile(r"^(Kata\s+Kunci|Keywords?)\b", re.IGNORECASE)

# Toleransi spasi OPSIONAL setelah titik ("Gambar 2. 1" ATAU "Gambar
# 2.1") -- format auto-numbering caption Word yang umum dipakai
# kampus literally menyisipkan spasi di situ (bukan salah ketik
# mahasiswa). Sebelumnya regex ini kaku ("\d+\.\d+" tanpa spasi),
# jadi SEMUA caption yang formatnya "N. N" gagal match dan jatuh ke
# ptype "body" -- ikut kena cek line-spacing/font-size/istilah-asing
# ala body text (padahal caption punya aturan sendiri), DAN otomatis
# lolos tanpa pernah divalidasi oleh compare_caption_numbering() di
# comparator.py (yang cuma memproses elemen ber-ptype "caption").
CAPTION_RE = re.compile(r"^(Gambar|Tabel)\s+\d+\.\s*\d+")

# Baris "Sumber: ..." / "Sumber : ..." di bawah gambar/tabel kutipan --
# opsional (panduan: "penulis DAPAT menyebutkan sumber referensi"),
# jadi cuma dicek FORMATnya (rata tengah, 11pt) KALAU baris ini ada,
# tidak pernah dianggap wajib ada.
SOURCE_CITATION_RE = re.compile(r"^Sumber\s*[:.]", re.IGNORECASE)

# Notice baku "Halaman ini sengaja dikosongkan" yang ditaruh di halaman
# genap kosong (buat cetak bolak-balik). Ini BUKAN body text -- tidak
# ada aturan font/spasi/istilah-asing yang relevan buat baris ini, jadi
# dikecualikan total (ptype "notice") supaya tidak salah-tandai dengan
# aturan body biasa. Longgar terhadap tanda kutip lurus/pintar (", “”,
# ‘’) di kedua sisi dan variasi kapitalisasi.
BLANK_PAGE_NOTICE_RE = re.compile(
    r"^[\"'“”‘’]*\s*halaman\s+ini\s+sengaja\s+dikosongkan\s*[\"'“”‘’]*$",
    re.IGNORECASE,
)

# Nomor identitas persamaan, mis. "(2.1)" -- dicari di AKHIR teks
# paragraf persamaan (bukan cuma exact-match seluruh baris seperti
# EQUATION_NUMBER_RE), karena OMML (rumus matematikanya sendiri) tidak
# pernah muncul di para.text sama sekali -- yang tersisa di teks
# paragraf cuma tab + label nomornya.
EQUATION_NUMBER_IN_TEXT_RE = re.compile(r"\(\d+(?:\.\d+)*\)\s*$")

# Kata sambung/preposisi pendek yang boleh huruf kecil di tengah judul
# Title Case (bukan kata pertama/terakhir). Dipakai is_titlecase().
SMALL_WORDS = {
    # indonesia
    "yang",
    "dan",
    "atau",
    "di",
    "ke",
    "dari",
    "untuk",
    "dengan",
    "pada",
    "dalam",
    "sebagai",
    "bagi",
    "oleh",
    "atas",
    "antara",
    "tanpa",
    "serta",
    "per",
    "akan",
    "adalah",
    # english
    "of",
    "in",
    "on",
    "for",
    "with",
    "and",
    "or",
    "the",
    "a",
    "an",
    "to",
    "as",
    "by",
    "at",
    "vs",
}

# Judul setingkat bab yang bukan "BAB N" tapi tetap pakai aturan
# ukuran/bold/kapital yang sama (14pt, bold, all-caps).
FRONT_HEADINGS = {
    "DAFTAR ISI",
    "DAFTAR GAMBAR",
    "DAFTAR TABEL",
    "DAFTAR PUSTAKA",
    "DAFTAR LAMPIRAN",
    "KATA PENGANTAR",
    "ABSTRAK",
    "ABSTRACT",
}


def _map_section_indices(doc) -> dict:
    """
    Peta {index_paragraf: index_section}, dibaca dari elemen w:sectPr asli
    docx (bukan tebakan teks). Section berganti setelah paragraf yang
    menyimpan w:sectPr di dalam w:pPr-nya (ini pola standar python-docx
    untuk section break di tengah body).
    """
    section_idx = 0
    mapping = {}
    for i, para in enumerate(doc.paragraphs):
        mapping[i] = section_idx
        pPr = para._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            section_idx += 1
    return mapping


def _get_numbering_ilvl(para) -> int | None:
    """
    Ambil ilvl (level list, 0-indexed) langsung dari w:numPr paragraf.
    Ini mekanisme XML yang SAMA yang dipakai Word untuk merender angka
    "3.2.1" di layar -- kalau paragraf tidak punya ini, Word memang
    tidak menampilkan angka otomatis apa pun untuknya. Sengaja TIDAK
    bergantung pada nama Style paragraf sama sekali, supaya tetap
    terdeteksi walau mahasiswa pakai style "Normal" biasa.

    Return None kalau paragraf ini tidak punya numbering otomatis
    (baik langsung di paragraf maupun diwariskan dari style).
    """
    pPr = para._p.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr")) if pPr is not None else None

    if numPr is None:
        # cek juga apakah numPr didefinisikan di level style, bukan
        # langsung di paragraf
        style = para.style
        while style is not None:
            style_pPr = style.element.find(qn("w:pPr"))
            style_numPr = (
                style_pPr.find(qn("w:numPr")) if style_pPr is not None else None
            )
            if style_numPr is not None:
                numPr = style_numPr
                break
            style = style.base_style

    if numPr is None:
        return None

    ilvl_el = numPr.find(qn("w:ilvl"))
    return int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0


def _has_outline_level(para) -> bool:
    """
    True kalau paragraf ini (langsung atau lewat style) punya
    w:outlineLvl -- ini yang benar-benar menandakan "nomor ini heading
    di outline dokumen" (1.1, 1.2, ...), BEDA dari w:numPr/w:ilvl yang
    juga dipakai Word untuk bullet/numbered list BIASA di badan teks
    (mis. daftar "a. b. c." menjelaskan jenis-jenis ROV).

    Sebelumnya kode ini nganggep SEMUA paragraf ber-ilvl otomatis =
    heading, jadi list biasa yang kebetulan diberi indent level 1
    (ilvl=1) ikut ke-klasifikasi "subchapter_title" -- padahal itu
    cuma bullet point di dalam body text.
    """
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None and pPr.find(qn("w:outlineLvl")) is not None:
        return True

    style = para.style
    while style is not None:
        style_pPr = style.element.find(qn("w:pPr"))
        if style_pPr is not None and style_pPr.find(qn("w:outlineLvl")) is not None:
            return True
        style = style.base_style

    return False


def is_unnumbered_heading(para) -> bool:
    """
    Deteksi heading level di bawah "X.Y.Z" yang memang TIDAK dinomori
    sama sekali oleh panduan (mis. "Parsing", "Ekstraksi Parameter",
    "Rule-Based Comparator"). Karena tidak ada angka untuk dicari sama
    sekali (bukan tersembunyi -- memang tidak ada), deteksi pakai ciri
    visual: seluruh baris bold, pendek, dan tidak diakhiri tanda baca
    kalimat (heading berdiri sendiri, bukan kalimat penuh).

    Ini tebakan berbasis pola, bukan fakta pasti seperti ilvl/nomor
    manual -- berisiko false positive/negative pada kasus tepi, tapi
    ini satu-satunya sinyal yang tersedia untuk level tanpa nomor.
    """
    text = para.text.strip()
    if not text:
        return False

    runs_with_text = [r for r in para.runs if r.text.strip()]
    if not runs_with_text:
        return False

    all_bold = all(_resolve_effective_bold(r, para) for r in runs_with_text)
    is_short = len(text.split()) <= 8
    no_sentence_ending = not text.rstrip().endswith((".", ",", ";"))

    return all_bold and is_short and no_sentence_ending


def _looks_like_heading_text(remainder: str) -> bool:
    """
    Heuristik pembeda "X.Y ..." yang BENERAN judul sub-bab (mis. "1.1
    Latar Belakang") dari angka yang KEBETULAN match pola
    SUBCHAPTER_NUMBER_RE tapi sebenarnya bagian dari kalimat body biasa.

    BUG YANG DIPERBAIKI: notasi angka ribuan gaya Indonesia MEMAKAI
    TITIK sebagai pemisah (bukan koma seperti format Inggris), mis.
    "1.234.567 penduduk mengalami ...", "2.500 responden mengisi
    survei ...", "10.000 data training digunakan ...". Semua ini
    match SUBCHAPTER_NUMBER_RE (angka bertitik diikuti spasi + kata)
    dan SEBELUM PERBAIKAN INI langsung diklasifikasi sebagai
    subchapter_title/subsubchapter_title -- comparator.py lantas
    salah menuntut paragraf statistik semacam itu ditulis ALL
    CAPS/Title Case + bold + 12pt, padahal itu body text biasa.

    Pembeda yang dipakai (sama semangatnya dengan is_unnumbered_
    heading() di atas -- ciri visual/tekstual heading, bukan makna
    semantik): judul sub-bab ASLI selalu (1) PENDEK -- bukan kalimat
    penuh, dan (2) kata SETELAH angkanya diawali HURUF KAPITAL sesuai
    konvensi ALL CAPS/Title Case panduan. Kalimat statistik yang
    kebetulan diawali angka ribuan TIDAK PUNYA dua ciri itu -- kata
    setelah angkanya nyaris selalu kata benda huruf kecil ("penduduk",
    "responden", "data", dst), dan kalimatnya panjang.
    """
    remainder = remainder.strip()
    if not remainder:
        return False

    is_short = len(remainder.split()) <= 10
    no_sentence_ending = not remainder.rstrip().endswith((".", ",", ";"))

    first_letter = re.search(r"[A-Za-z]", remainder)
    starts_capital = bool(first_letter) and first_letter.group(0).isupper()

    return is_short and no_sentence_ending and starts_capital


def is_titlecase(text: str) -> bool:
    """
    Cek apakah teks Title Case: kata utama diawali huruf kapital, kata
    sambung/preposisi pendek (SMALL_WORDS) boleh huruf kecil kecuali
    jadi kata pertama/terakhir. Nomor bagian di depan (mis. "2.2.1 ")
    dan token sitasi (mis. "[19]") diabaikan dulu sebelum dicek. Aman
    dipakai juga untuk heading tanpa nomor (regex prefix-nya memang
    tidak match kalau tidak ada nomor, tidak masalah).

    Teks yang ALL CAPS sengaja dianggap BUKAN title case -- itu domain
    aturan level "X.Y", bukan level ini.
    """
    cleaned = re.sub(r"^\d+(\.\d+)*\s+", "", text.strip())
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    if not cleaned:
        return False

    words = re.findall(r"[A-Za-z][A-Za-z'\-]*", cleaned)
    if not words:
        return False

    if all(w.isupper() and len(w) > 1 for w in words):
        return False

    for idx, w in enumerate(words):
        is_edge = idx == 0 or idx == len(words) - 1
        if not is_edge and w.lower() in SMALL_WORDS:
            continue
        if not w[0].isupper():
            return False

    return True


def classify_paragraph(text: str, para=None) -> str:
    """
    Menentukan tipe elemen. Tiga jalur deteksi dicoba berurutan:
      1. Nomor diketik manual di teks (regex generik, semua level)
      2. Nomor digenerate otomatis oleh Word (ilvl di w:numPr)
      3. Heading tanpa nomor sama sekali (ciri visual: bold + pendek)
    Jalur 2 & 3 butuh objek `para` (bukan cuma string teks), karena
    informasinya ada di XML paragraf, bukan di teksnya.
    """
    stripped = text.strip()
    if not stripped:
        return "empty"

    upper = stripped.upper()

    if CHAPTER_TITLE_RE.match(upper) and stripped == upper:
        return "chapter_title"

    if upper in FRONT_HEADINGS and stripped == upper:
        return "chapter_title"

    # Label nomor persamaan ("(2.1)") -- pendek & sering di-bold, tapi
    # BUKAN heading. Harus dicek SEBELUM jalur is_unnumbered_heading()
    # supaya tidak ke-tangkap sebagai subsubchapter_title.
    if EQUATION_NUMBER_RE.match(stripped):
        return "body"

    # --- Jalur 1: nomor diketik manual, kedalaman dihitung dari jumlah titik ---
    m = SUBCHAPTER_NUMBER_RE.match(stripped)
    if m:
        # PENJAGA (lihat _looks_like_heading_text() untuk detail lengkap):
        # angka ribuan gaya Indonesia ("1.234.567 penduduk ...", "2.500
        # responden ...") JUGA match pola di atas -- tanpa penjagaan ini,
        # kalimat body yang kebetulan diawali angka ribuan seperti itu
        # akan salah diklasifikasi jadi heading. HANYA dipercaya sebagai
        # heading kalau teks setelah angkanya juga terlihat seperti judul
        # (pendek + diawali huruf kapital), bukan sekadar berhasil match
        # regex-nya.
        remainder = stripped[m.end(1) :]
        if _looks_like_heading_text(remainder):
            depth = m.group(1).count(".") + 1  # "1.1"->2, "2.2.1"->3, "3.2.2.1"->4
            if depth == 2:
                return "subchapter_title"
            elif depth >= 3:
                return "subsubchapter_title"
        # Kalau gagal heuristik di atas, JANGAN langsung return "body" --
        # sengaja dibiarkan lanjut ke Jalur 2/3 di bawah (siapa tahu
        # paragraf ini toh memang heading lewat numbering Word yang
        # angkanya kebetulan tidak diketik manual di teks), baru jatuh ke
        # default "body" di akhir fungsi kalau semua jalur gagal juga.

    # --- Jalur 2: nomor digenerate otomatis via numbering Word ---
    # HANYA dipercaya kalau paragraf ini beneran linked ke outline
    # level Word (w:outlineLvl) -- itu satu-satunya sinyal yang
    # membedakan "nomor heading" dari "nomor list biasa di body text"
    # (keduanya sama-sama pakai w:numPr/w:ilvl di XML).
    if para is not None and _has_outline_level(para):
        ilvl = _get_numbering_ilvl(para)
        if ilvl == 0:
            # Heading level TERTINGGI (mis. style "Heading 1") --
            # setara "BAB N", TERLEPAS teksnya diawali kata "BAB"
            # atau tidak. BUG SEBELUMNYA: cabang ilvl==0 ini TIDAK
            # ditangani sama sekali (cuma ilvl==1 & ilvl>=2 yang
            # return), jadi kontrol jatuh ke Jalur 3 (is_unnumbered_
            # heading()) dan keliru diklasifikasi "subsubchapter_
            # title" -- level TERDALAM, padahal ini level TERTINGGI.
            # Ini terjadi terutama di dokumen yang nomor babnya
            # digenerate murni oleh multilevel-list Word yang nempel
            # ke style Heading 1 (mis. teks paragraf cuma
            # "PENDAHULUAN", bukan "BAB 1 PENDAHULUAN" -- nomor "BAB
            # 1"-nya cuma dirender visual oleh Word, tidak pernah ada
            # di w:t). Akibat salah klasifikasi ini, extract_paragraph_
            # params() di bawah tidak pernah mendeteksi heading ini
            # sebagai "chapter_title" -> current_chapter tidak pernah
            # pindah dari "Pra-Bab", DAN flag in_figure_table_list/
            # in_toc yang sempat ON di halaman DAFTAR GAMBAR/DAFTAR
            # TABEL/DAFTAR ISI tidak pernah di-reset lagi untuk SISA
            # seluruh dokumen -- semua paragraf sesudahnya (termasuk
            # caption "Gambar N.N ..." yang sudah benar diketik siswa)
            # keliru didemosi jadi ptype "toc_entry" dan dilewati dari
            # cek caption, sehingga gambar yang SUDAH punya caption
            # tetap dilaporkan "tidak punya caption".
            return "chapter_title"
        elif ilvl == 1:
            return "subchapter_title"
        elif ilvl is not None and ilvl >= 2:
            return "subsubchapter_title"

        # --- Jalur 3: heading tanpa nomor sama sekali ---
        # Disatukan ke "subsubchapter_title" karena aturan formatnya
        # (Title Case + bold) identik dengan level "X.Y.Z".
        if is_unnumbered_heading(para):
            return "subsubchapter_title"

    if CAPTION_RE.match(stripped):
        return "caption"

    return "body"


def _resolve_effective_bold(run, paragraph) -> bool:
    """
    run.bold bisa bernilai None kalau bold TIDAK di-set langsung di run,
    melainkan diwarisi dari paragraph style. None di sini bukan berarti
    "tidak bold", tapi "ikut warisan style" -- jadi harus ditelusuri ke
    style dan base_style-nya sebelum dianggap tidak bold.
    """
    if run.bold is not None:
        return run.bold

    style = paragraph.style
    while style is not None:
        if style.font.bold is not None:
            return style.font.bold
        style = style.base_style

    return False


def _resolve_effective_italic(run, paragraph) -> bool:
    """
    Simetris dengan _resolve_effective_bold(). run.italic bisa None
    kalau italic tidak di-set langsung di run, melainkan diwarisi dari
    paragraph style -- ditelusuri ke style dan base_style-nya sebelum
    dianggap tidak italic.
    """
    if run.italic is not None:
        return run.italic

    style = paragraph.style
    while style is not None:
        if style.font.italic is not None:
            return style.font.italic
        style = style.base_style

    return False


def _resolve_effective_font_size(run, paragraph) -> float | None:
    """
    Simetris dengan _resolve_effective_bold(). Kalau run.font.size
    None (tidak di-set langsung di run), telusuri ke style dan
    base_style-nya. Sebelumnya nilai None di sini membuat comparator
    MEN-SKIP pengecekan ukuran font sepenuhnya (bukan dianggap salah),
    sehingga mahasiswa yang set ukuran font lewat modifikasi style
    (bukan pilih ukuran langsung di teks) lolos tanpa pernah dicek.
    """
    if run.font.size is not None:
        return run.font.size.pt

    style = paragraph.style
    while style is not None:
        if style.font.size is not None:
            return style.font.size.pt
        style = style.base_style

    return None


def _resolve_effective_font_name(run, paragraph) -> str | None:
    """Simetris dengan _resolve_effective_font_size(), untuk nama font."""
    if run.font.name is not None:
        return run.font.name

    style = paragraph.style
    while style is not None:
        if style.font.name is not None:
            return style.font.name
        style = style.base_style

    return None


def _get_pgnum_type(section, previous_fmt: str | None) -> dict:
    """
    Ambil format & start value nomor halaman dari w:pgNumType di sectPr.

    KOREKSI (setelah verifikasi render langsung ke PDF -- LibreOffice
    ternyata membedakan DUA KASUS, bukan satu aturan seragam):

      1. Elemen w:pgNumType SAMA SEKALI TIDAK ADA di sectPr -> section
         ini betul-betul tidak nyentuh pengaturan penomoran apa pun,
         fmt MEWARISI dari section sebelumnya (dibuktikan lewat render:
         halaman fisik 7 di dokumen test, section tanpa pgNumType sama
         sekali, tetap muncul "v" di footer -- lanjut romawi dari
         section sebelumnya, bukan reset ke decimal).
      2. Elemen w:pgNumType ADA tapi w:fmt-nya TIDAK disebutkan (mis.
         cuma w:start yang di-set buat restart ke 1) -> fmt-nya DEFAULT
         ke "decimal", BUKAN warisan (dibuktikan lewat render: halaman
         fisik 8 / BAB 1 di dokumen test, section-nya punya
         <w:pgNumType w:start="1"/> tanpa w:fmt, tetap muncul "1" di
         footer -- bukan "i", walau section sebelumnya romawi). Ini
         cocok sama perilaku dialog "Format Page Numbers" di Word:
         begitu dialog itu disentuh buat set start, format dropdown-nya
         default ke "1,2,3..." kalau gak dipilih eksplisit ke romawi.

    w:start TETAP selalu ikut aturan lama (None = lanjut dari section
    sebelumnya) -- itu gak berubah, cuma w:fmt yang beda perilaku
    tergantung elemen pgNumType-nya ada atau tidak.

    previous_fmt: fmt EFEKTIF (hasil resolusi, bukan mentah) dari
    section sebelumnya, None kalau ini section pertama di dokumen.
    """
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))

    if pgNumType is None:
        # Kasus 1: elemen sama sekali gak ada -> full warisan.
        inherited_fmt = previous_fmt if previous_fmt is not None else "decimal"
        return {"fmt": inherited_fmt, "start": None}

    # Kasus 2: elemen ADA -> fmt default "decimal" kalau gak disebutkan
    # (BUKAN warisan dari section sebelumnya).
    fmt = pgNumType.get(qn("w:fmt")) or "decimal"
    start = pgNumType.get(qn("w:start"))
    return {"fmt": fmt, "start": int(start) if start is not None else None}


def _get_mirror_margins(doc) -> bool:
    """
    Ambil status mirror-margin, dari w:mirrorMargins di settings.xml.

    KOREKSI (sebelumnya salah): w:mirrorMargins BUKAN child dari
    w:sectPr (beda dari margin/pgNumType yang memang atribut section).
    Elemen ini ada di w:settings (settings.xml) -- SATU untuk SELURUH
    DOKUMEN, hasil centang "Mirror margins" di Page Setup > Multiple
    Pages. Versi lama nyari di section._sectPr, yang gak akan PERNAH
    ketemu elemennya di sana sama sekali -- selalu balik False walau
    dokumen sudah benar diaktifkan mirror-margin-nya.

    w:mirrorMargins itu elemen "presence flag": kalau elemen ADA dan
    w:val TIDAK eksplisit "0"/"false"/"off", artinya ON. Kalau elemen
    itu sama sekali TIDAK ADA, artinya OFF (default Word).

    Parameter `doc` di sini Document (bukan Section), karena settingnya
    memang level dokumen.
    """
    settings_el = doc.settings.element
    el = settings_el.find(qn("w:mirrorMargins"))
    if el is None:
        return False
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return val.lower() not in ("0", "false", "off")


def _footer_page_field_info(section) -> dict:
    """
    Cek apakah footer section ini (footer MILIK SENDIRI, bukan warisan)
    punya field nomor halaman (PAGE), dan kalau ada, apakah paragrafnya
    center-aligned.

    PENTING -- paragraf field PAGE yang di-insert lewat menu Word
    "Insert > Page Number" dibungkus w:sdt/w:sdtContent (content
    control/Quick Part), BUKAN w:p sebagai anak langsung footer part.
    section.footer.paragraphs cuma expose w:p yang jadi anak LANGSUNG
    body footer, jadi paragraf di dalam sdt tidak pernah kelihatan
    lewat situ walau field-nya beneran ada -- selalu keliru dilaporkan
    "tidak ada field". Makanya di sini nyari langsung ke SEMUA elemen
    w:p di XML footer part pakai .iter(), supaya tetap ketemu walau
    dibungkus sdt (atau tabel, kasus struktur lain yang serupa).
    """
    footer_el = section.footer._element
    for p in footer_el.iter(qn("w:p")):
        has_field = any(
            "PAGE" in (fld.get(qn("w:instr")) or "").upper()
            for fld in p.iter(qn("w:fldSimple"))
        ) or any(
            instr.text and "PAGE" in instr.text.upper()
            for instr in p.iter(qn("w:instrText"))
        )
        if has_field:
            pPr = p.find(qn("w:pPr"))
            jc_el = pPr.find(qn("w:jc")) if pPr is not None else None
            alignment = jc_el.get(qn("w:val")) if jc_el is not None else None
            return {"has_page_field": True, "alignment": alignment}
    return {"has_page_field": False, "alignment": None}


def _get_paragraph_alignment(para) -> str | None:
    """
    Ambil rata paragraf (w:jc) sebagai string mentah ("left", "center",
    "right", "both"=justify), ditelusuri sampai ke style/base_style
    kalau tidak di-set langsung di paragraf -- simetris dengan cara
    field lain (bold/italic/font) sudah ditelusuri di modul ini. Dibaca
    dari XML mentah (bukan enum python-docx) supaya konsisten dengan
    cara footer alignment sudah dibaca di _footer_page_field_info().
    """
    pPr = para._p.find(qn("w:pPr"))
    jc_el = pPr.find(qn("w:jc")) if pPr is not None else None
    if jc_el is not None:
        return jc_el.get(qn("w:val"))

    style = para.style
    while style is not None:
        style_pPr = style.element.find(qn("w:pPr"))
        style_jc = style_pPr.find(qn("w:jc")) if style_pPr is not None else None
        if style_jc is not None:
            return style_jc.get(qn("w:val"))
        style = style.base_style

    return None


def _has_drawing(para) -> bool:
    """
    True kalau paragraf memuat objek gambar (w:drawing -- mencakup
    gambar inline maupun anchored/floating, keduanya sama-sama muncul
    sebagai elemen w:drawing di dalam w:r). Dicari lewat .iter() karena
    w:drawing adalah CUCU paragraf (w:p > w:r > w:drawing), bukan anak
    langsung.
    """
    return next(para._p.iter(qn("w:drawing")), None) is not None


def _has_equation(para) -> bool:
    """
    True kalau paragraf memuat rumus matematika OMML (m:oMath). OMML
    TIDAK PERNAH muncul di para.text (python-docx cuma membaca w:t) --
    jadi paragraf persamaan yang tidak diberi nomor identitas apa pun
    akan terlihat seperti paragraf kosong tanpa deteksi ini, dan lolos
    tanpa pernah dicek sama sekali.
    """
    return next(para._p.iter(qn("m:oMath")), None) is not None


# Rentang Unicode "Mathematical Alphanumeric Symbols" (U+1D400-U+1D7FF)
# -- dipakai mahasiswa yang mengetik variabel/simbol rumus (mis. huruf
# italic matematis "n") langsung sebagai teks biasa dengan font Cambria
# Math per-run, TANPA membungkusnya sebagai objek OMML asli (m:oMath).
# _has_equation() di atas tidak menangkap kasus ini karena tidak ada
# m:oMath sama sekali -- makanya butuh deteksi terpisah di level
# karakter/run (bukan di level paragraf), supaya paragraf yang cuma
# menyisipkan SATU variabel matematis di tengah kalimat body biasa
# tidak ikut ke-reclassify total jadi ptype "equation" (yang akan
# membuat SISA kalimatnya ikut lolos dari cek font/size body juga).
_MATH_ALPHANUMERIC_RE = re.compile(
    r"[\U0001D400-\U0001D7FF]|[\u2200-\u22FF\u2202\u2207\u221A\u221E]"
)

# Nama font yang lazim dipakai Word untuk rumus/simbol matematis (di-set
# otomatis oleh Word saat pakai fitur Insert Equation/Math AutoCorrect,
# atau diset manual mahasiswa saat mengetik variabel matematis sebagai
# teks biasa). Font-font ini WAJAR beda dari Times New Roman body text
# dan tidak seharusnya ditandai sebagai kesalahan font.
MATH_FONT_NAMES = {"Cambria Math"}


def _is_math_run_text(text: str) -> bool:
    """
    True kalau teks run mengandung karakter matematis Unicode (huruf
    italic/bold matematis, operator himpunan, akar, dsb) -- indikasi
    run ini adalah variabel/simbol rumus yang diketik sebagai teks
    biasa, bukan body text sungguhan.
    """
    return bool(_MATH_ALPHANUMERIC_RE.search(text))


def _get_left_indent_cm(para) -> float | None:
    """Indentasi kiri paragraf dalam cm, None kalau tidak di-set."""
    ind = para.paragraph_format.left_indent
    return ind.cm if ind is not None else None


def _get_right_indent_cm(para) -> float | None:
    """
    Indentasi kanan paragraf dalam cm, None kalau tidak di-set. Simetris
    dengan _get_left_indent_cm() -- dibutuhkan untuk cek kutipan
    langsung PANJANG (quotation_checker.py / compare_quotations() di
    comparator.py), yang mensyaratkan indentasi KIRI *DAN* KANAN masuk
    10mm dari batas tulisan (beda dari persamaan yang cuma indentasi
    kiri).
    """
    ind = para.paragraph_format.right_indent
    return ind.cm if ind is not None else None


def extract_paragraph_params(doc) -> list[dict]:
    """
    Ekstrak parameter mentah tiap paragraf: tipe, teks, chapter, dan
    parameter font/spacing dari tiap run di dalamnya.
    """
    elements = []
    section_map = _map_section_indices(doc)

    current_chapter = "Pra-Bab"  # sampul, kata pengantar, daftar isi, dst.
    expect_chapter_name = False  # true tepat setelah baris "BAB N" ditemukan
    bab_counter = 0  # penomor "BAB N" mandiri, dipakai KHUSUS untuk heading
    # level-0 (ilvl==0 / style "Heading 1") yang teksnya TIDAK diawali
    # "BAB N" sama sekali -- lihat cabang baru di bawah untuk detail.
    # HANYA dipakai kalau dokumen TIDAK PERNAH menulis "BAB N" secara
    # literal di teks sama sekali (lihat literal_bab_used di bawah) --
    # kalau dokumen SUDAH punya bab literal, cabang bab_counter ini
    # dimatikan supaya heading Heading-1-style lain (paragraf judul
    # paper di tinjauan pustaka, subbab yang stylenya kebetulan
    # Heading 1, dst) tidak ikut dianggap bab baru.
    literal_bab_used = False  # true begitu ketemu heading "BAB N" literal pertama
    in_figure_table_list = False  # true selama di halaman Daftar Gambar/Tabel
    in_toc = False  # true selama di halaman Daftar Isi
    seen_bab = False  # true setelah "BAB N" pertama ditemukan
    # None | "id" | "en" -- lagi di dalam region Abstrak (ID) atau
    # Abstract (EN)? Dipakai buat nge-tag elem["abstract_part"], yang
    # jadi input compare_abstract() di comparator.py. Berhenti (balik
    # ke None) begitu ketemu heading LAIN (ABSTRACT sendiri jadi
    # pengecualian: dia MENGGANTI state dari "id" ke "en", bukan
    # menutup region).
    abstract_state = None

    # Halaman template tetap (sampul, halaman judul dalam, halaman
    # pengesahan, surat pernyataan keaslian) -- dokumen SELALU dimulai
    # di dalam salah satu halaman ini. Dilacak lewat isi teks (marker),
    # BUKAN index section Word, karena jumlah section/halaman kosong
    # di bagian depan beda-beda per dokumen (lihat SPECIAL_FRONT_PAGE_MARKERS
    # di atas).
    in_special_page = True

    for i, para in enumerate(doc.paragraphs):
        stripped = para.text.strip()
        has_image = _has_drawing(para)
        has_equation = _has_equation(para)
        # Paragraf gambar/persamaan sering TIDAK punya teks sama sekali
        # (objeknya OMML/drawing, bukan w:t) -- kalau di-skip seperti
        # paragraf kosong biasa, gambar/persamaan itu jadi tidak pernah
        # tertangkap atau dicek format tata letaknya sama sekali.
        if not stripped and not has_image and not has_equation:
            continue

        section_idx = section_map[i]
        ptype = classify_paragraph(para.text, para)
        upper_now = stripped.upper()

        # Keluar dari halaman template tetap begitu ketemu front matter
        # yang justru WAJIB ikut aturan body normal (Abstrak, Kata
        # Pengantar, Daftar Isi, dst). Sebaliknya, masuk (lagi) ke
        # halaman template tetap begitu ketemu marker sampul/pengesahan
        # -- ini yang membuatnya kebal terhadap berapa pun banyaknya
        # halaman kosong/section break di antaranya.
        if upper_now in NORMAL_FRONT_MATTER_MARKERS:
            in_special_page = False
        elif (
            upper_now in SPECIAL_FRONT_PAGE_MARKERS
            and not in_toc
            and not in_figure_table_list
        ):
            # Guard in_toc/in_figure_table_list: teks "HALAMAN JUDUL"/
            # "HALAMAN PENGESAHAN" juga muncul sebagai ENTRI di Daftar
            # Isi (bukan halaman aslinya) -- tanpa guard ini, entri
            # tsb keliru men-trigger in_special_page lagi dan lolos
            # lewat jalur "cover" alih-alih jalur demosi in_toc di bawah.
            in_special_page = True

        if BLANK_PAGE_NOTICE_RE.match(stripped):
            # "Halaman ini sengaja dikosongkan" -- bukan body text,
            # tidak ada aturan format yang relevan untuk baris ini.
            ptype = "notice"
            expect_chapter_name = False
        elif in_special_page and not seen_bab:
            # Halaman sampul/judul/pengesahan: dikecualikan total dari
            # rule body/caption/judul (formatnya ikut template resmi,
            # bukan aturan body 12pt/bold/kapitalisasi biasa).
            ptype = "cover"
            expect_chapter_name = False
        else:
            if ptype == "chapter_title":
                upper_stripped = stripped.upper()
                # Region Abstrak (ID) dibuka oleh heading "ABSTRAK", lalu
                # DIGANTI (bukan ditutup) oleh heading "ABSTRACT" jadi
                # region Abstract (EN). Heading LAIN manapun (BAB 1, KATA
                # PENGANTAR, dst) menutup kedua region ini.
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
                    # current_chapter SELALU dinormalisasi ke "BAB N"
                    # murni (mis. "BAB 1"), terlepas nama babnya
                    # menyatu di baris yang sama ("BAB 1 PENDAHULUAN")
                    # atau menyusul di paragraf berikutnya. Tanpa ini,
                    # dua dokumen dengan format penulisan judul bab
                    # yang beda akan menghasilkan label bab yang beda
                    # juga ("BAB 1" vs "BAB 1 PENDAHULUAN"), padahal
                    # secara substansi itu bab yang sama.
                    current_chapter = re.match(r"^BAB\s+\d+", upper_stripped).group(0)
                    seen_bab = True
                    literal_bab_used = True
                    # expect_chapter_name HANYA kalau baris ini murni
                    # angka bab tanpa nama -- nama babnya menyusul di
                    # paragraf berikutnya. Kalau nama babnya sudah
                    # menyatu di baris yang sama, paragraf berikutnya
                    # adalah konten baru (subbab/body), BUKAN nama
                    # bab -- jangan sampai ikut dipaksa jadi
                    # chapter_title.
                    expect_chapter_name = bool(
                        CHAPTER_TITLE_ONLY_RE.match(upper_stripped)
                    )
                elif (
                    not literal_bab_used
                    and upper_stripped not in FRONT_HEADINGS
                    and _has_outline_level(para)
                    and _get_numbering_ilvl(para) == 0
                ):
                    # Heading level TERTINGGI (ilvl==0, mis. style
                    # "Heading 1") TAPI teksnya TIDAK diawali "BAB N"
                    # -- pola umum kalau nomor babnya digenerate MURNI
                    # oleh multilevel-list Word yang nempel ke style
                    # Heading 1 (nomornya cuma dirender visual oleh
                    # Word, TIDAK PERNAH ada di w:t/paragraf.text).
                    # current_chapter untuk kasus ini di-generate
                    # sendiri via bab_counter berurutan ("BAB 1", "BAB
                    # 2", dst mengikuti URUTAN heading level-0 yang
                    # ditemui) -- BUKAN dari teks headingnya sendiri
                    # ("PENDAHULUAN", "KAJIAN PUSTAKA", dst), supaya
                    # tetap konsisten dengan label "BAB N" yang dipakai
                    # skema chapter di seluruh pipeline (findings/
                    # summary per_chapter, dll). Cabang ini JUGA yang
                    # men-trigger reset in_figure_table_list/in_toc di
                    # bawah -- tanpa ini, flag itu nyangkut ON dari
                    # DAFTAR GAMBAR/DAFTAR TABEL/DAFTAR ISI untuk SISA
                    # dokumen, karena heading BAB sungguhan di sini
                    # sebelumnya tidak pernah dikenali sebagai
                    # "chapter_title" sama sekali (lihat classify_
                    # paragraph()).
                    #
                    # GUARD BARU (not literal_bab_used): cabang ini
                    # HANYA boleh jalan sebelum dokumen pernah menulis
                    # "BAB N" literal. Begitu dokumen terbukti memakai
                    # konvensi bab literal (BAB 1-5 dst di dokumen
                    # nyata), style "Heading 1" yang dipakai LAGI di
                    # tempat lain (mis. daftar judul paper di tinjauan
                    # pustaka, subbab yang stylenya salah pilih Heading
                    # 1 alih-alih Heading 2) BUKAN penanda bab baru --
                    # tanpa guard ini, tiap paragraf semacam itu
                    # menaikkan bab_counter dan menghasilkan "BAB 6",
                    # "BAB 7", ..., "BAB 14" palsu yang tidak pernah
                    # ada di dokumen aslinya, sekaligus salah menuntut
                    # paragraf itu 14pt/rata-tengah/ALL CAPS (aturan
                    # chapter_title) padahal ini cuma heading subbab.
                    bab_counter += 1
                    current_chapter = f"BAB {bab_counter}"
                    seen_bab = True
                    expect_chapter_name = False
                elif (
                    literal_bab_used
                    and upper_stripped not in FRONT_HEADINGS
                    and _has_outline_level(para)
                    and _get_numbering_ilvl(para) == 0
                ):
                    # Sama seperti kasus di atas (Heading-1-style tanpa
                    # teks "BAB N"), TAPI dokumen ini SUDAH terbukti
                    # memakai bab literal -- jadi heading ini bukan bab
                    # baru, cuma subbab/heading biasa yang stylenya
                    # kebetulan Heading 1. current_chapter TIDAK
                    # berubah (tetap ikut bab yang sedang berjalan),
                    # dan ptype diturunkan ke "subsubchapter_title"
                    # supaya dicek pakai aturan subbab (Title Case +
                    # bold) alih-alih dipaksa 14pt/center/ALL CAPS ala
                    # judul bab.
                    ptype = "subsubchapter_title"
                    expect_chapter_name = False
                elif seen_bab:
                    # FRONT_HEADINGS (DAFTAR PUSTAKA, DAFTAR LAMPIRAN, dst)
                    # yang muncul SETELAH minimal satu BAB N -- beri
                    # current_chapter sendiri (bukan ikut nempel ke BAB
                    # terakhir seperti sebelumnya), supaya findings/checks
                    # untuk section ini dilaporkan sebagai grup sendiri,
                    # bukan numpuk membengkakkan angka "checked" BAB
                    # terakhir. FRONT_HEADINGS SEBELUM BAB 1 (mis. DAFTAR
                    # ISI, KATA PENGANTAR, ABSTRAK) TIDAK masuk sini --
                    # tetap sengaja digabung ke "Pra-Bab" seperti semula.
                    current_chapter = upper_stripped
                    expect_chapter_name = False
                else:
                    expect_chapter_name = False
                in_figure_table_list = upper_stripped in (
                    "DAFTAR GAMBAR",
                    "DAFTAR TABEL",
                )
                in_toc = upper_stripped == "DAFTAR ISI"
                # PENTING: flag in_figure_table_list/in_toc di atas cuma
                # dipakai buat baris-baris SETELAH heading ini (lihat
                # cabang elif in_toc/in_figure_table_list di bawah).
                # Heading "DAFTAR ISI"/"DAFTAR GAMBAR"/"DAFTAR TABEL" itu
                # SENDIRI TIDAK didemosi ke "toc_entry" -- dia tetap
                # "chapter_title" dan tetap wajib kena cek bold/size/
                # all-caps/center yang sama seperti "KATA PENGANTAR" atau
                # "BAB N" lainnya. Versi sebelumnya keliru mendemosi baris
                # heading-nya sendiri juga (bukan cuma entri di bawahnya),
                # jadi heading "DAFTAR TABEL" yang tidak bold/tidak center
                # lolos tanpa pernah dicek sama sekali.
            elif expect_chapter_name:
                ptype = "chapter_title"
                expect_chapter_name = False
            elif in_toc or in_figure_table_list:
                # SEMUA baris di halaman Daftar Isi/Gambar/Tabel --
                # entri caption ("Tabel 2.1 Tabel Perbandingan ... 12"),
                # entri judul bab ("BAB 1 PENDAHULUAN ....... 5"), nomor
                # halaman, dst -- DIKECUALIKAN TOTAL lewat ptype khusus
                # "toc_entry" (dilewati semua cek di comparator.py, sama
                # seperti "cover"). SEBELUMNYA cuma di-demote ke "body",
                # yang ternyata tetap kena cek spacing 1.5 & istilah
                # asing wajib italic milik body text biasa -- padahal
                # baris daftar isi/gambar/tabel secara konvensi lazim
                # pakai format berbeda (dot leader, dll), bukan body
                # text sungguhan. Baris ber-ptype "body" polos (mis.
                # nomor halaman sendirian) yang sebelumnya tidak
                # tersentuh cabang manapun sekarang juga ikut tercakup.
                ptype = "toc_entry"

        # --- override: gambar / persamaan / sumber kutipan ---
        # Dicek SETELAH klasifikasi heading/body/caption/cover biasa
        # (bukan menggantinya di awal) supaya halaman sampul/template
        # (ptype "cover") tetap dikecualikan total seperti sebelumnya --
        # gambar/persamaan di halaman itu memang tidak tunduk ke aturan
        # tata letak BAB biasa.
        if ptype != "cover":
            if has_equation:
                ptype = "equation"
            elif has_image:
                ptype = "image"
            elif SOURCE_CITATION_RE.match(stripped):
                ptype = "source_citation"

        runs_data = []
        for run in para.runs:
            if not run.text.strip():
                continue
            run_font_name = _resolve_effective_font_name(run, para)
            runs_data.append(
                {
                    "text": run.text,
                    "font_name": run_font_name,
                    "font_size_pt": _resolve_effective_font_size(run, para),
                    "bold": _resolve_effective_bold(run, para),
                    "italic": _resolve_effective_italic(run, para),
                    # True kalau run ini variabel/simbol matematis yang
                    # diketik sebagai teks biasa (bukan objek OMML) --
                    # dipakai comparator.py buat skip cek font PER-RUN
                    # (bukan per-paragraf), lihat _is_math_run_text() &
                    # MATH_FONT_NAMES di atas untuk alasan lengkapnya.
                    "is_math": (
                        run_font_name in MATH_FONT_NAMES or _is_math_run_text(run.text)
                    ),
                }
            )

        # abstract_part: dipakai compare_abstract() di comparator.py buat
        # cek jumlah kata Abstrak/Kata Kunci & wajib-italic Abstract (EN).
        # Cuma paragraf ptype "body" yang ditandai (bukan heading
        # "ABSTRAK"/"ABSTRACT" itu sendiri) -- baris "Kata Kunci"/
        # "Keywords" dipisah dari baris body biasa via ABSTRACT_KEYWORD_RE.
        abstract_part = None
        if abstract_state is not None and ptype == "body":
            is_keyword_line = bool(ABSTRACT_KEYWORD_RE.match(stripped))
            if abstract_state == "id":
                abstract_part = "keyword_id" if is_keyword_line else "abstrak_body"
            else:  # "en"
                abstract_part = "keyword_en" if is_keyword_line else "abstract_body"

        raw_line_spacing = para.paragraph_format.line_spacing
        line_spacing_rule = para.paragraph_format.line_spacing_rule
        # Kalau rule-nya EXACTLY/AT_LEAST, python-docx balikin Length (satuan
        # fisik: pt/cm/EMU) -- bukan angka kelipatan (multiple) kayak "1.5".
        # Membandingkan Length mentah langsung ke RULES["line_spacing"] (1.5)
        # gak masuk akal, itu dua besaran yang beda dimensi sama sekali.
        # Ditandai terpisah di sini (line_spacing_is_multiple) supaya
        # comparator tau harus membandingkan APA, bukan cuma ANGKA berapa.
        line_spacing_is_multiple = not isinstance(raw_line_spacing, Length)
        line_spacing_value = (
            raw_line_spacing.pt
            if isinstance(raw_line_spacing, Length)
            else raw_line_spacing
        )

        elements.append(
            {
                "index": i,
                "type": ptype,
                "full_text": stripped,
                "is_upper": stripped == stripped.upper(),
                "is_titlecase": is_titlecase(stripped),
                "line_spacing": line_spacing_value,
                "line_spacing_is_multiple": line_spacing_is_multiple,
                "line_spacing_rule": (
                    line_spacing_rule.name if line_spacing_rule is not None else None
                ),
                "runs": runs_data,
                "chapter": current_chapter,
                "section": section_idx,
                "abstract_part": abstract_part,
                "alignment": _get_paragraph_alignment(para),
                # Sebelumnya cuma diekstrak untuk ptype "equation" (satu-
                # satunya pemakai saat itu). Sekarang diekstrak untuk
                # SEMUA paragraf -- compare_quotations() di comparator.py
                # butuh left_indent_cm & right_indent_cm milik paragraf
                # kutipan langsung panjang, yang secara ptype tetap
                # "body" (bukan tipe khusus tersendiri).
                "left_indent_cm": _get_left_indent_cm(para),
                "right_indent_cm": _get_right_indent_cm(para),
                # Baris kosong TEPAT sebelum/sesudah paragraf ini (dicek
                # langsung dari doc.paragraphs, BUKAN dari `elements`
                # yang sudah dibangun -- paragraf kosong sengaja tidak
                # pernah masuk ke `elements`, jadi harus dicek di sumber
                # aslinya). Dipakai buat aturan "1 baris kosong di atas
                # & bawah" gambar/persamaan.
                "blank_before": i > 0 and not doc.paragraphs[i - 1].text.strip(),
                "blank_after": (
                    i < len(doc.paragraphs) - 1
                    and not doc.paragraphs[i + 1].text.strip()
                ),
            }
        )

    return elements


def extract_section_params(doc) -> list[dict]:
    """Ekstrak parameter margin, ukuran kertas, & penomoran halaman per section."""
    sections_data = []
    last_footer_info = {"has_page_field": False, "alignment": None}
    last_pgnum_fmt = None  # None di section pertama -- tidak ada yang diwarisi

    for i, section in enumerate(doc.sections):
        pgnum = _get_pgnum_type(section, last_pgnum_fmt)
        last_pgnum_fmt = pgnum[
            "fmt"
        ]  # fmt EFEKTIF (sudah warisan), buat section berikutnya

        if section.footer.is_linked_to_previous:
            footer_info = last_footer_info
        else:
            footer_info = _footer_page_field_info(section)
            last_footer_info = footer_info

        # Tiap margin/ukuran-kertas disimpan DUA representasi:
        #   *_twips -> basis PERBANDINGAN (comparator.py). Ini unit asli
        #     yang tersimpan di w:pgMar/w:pgSz di XML docx (python-docx
        #     Length.twips baca langsung dari situ), jadi integer EXACT,
        #     bebas noise floating-point konversi EMU->cm.
        #   *_cm    -> cuma buat DITAMPILKAN (pesan error ke user,
        #     display page_layout_checker, dsb). JANGAN dipakai buat
        #     perbandingan abs(a - b) > tol lagi -- itu tugas *_twips.
        sections_data.append(
            {
                "index": i,
                "left_cm": section.left_margin.cm,
                "left_twips": section.left_margin.twips,
                "top_cm": section.top_margin.cm,
                "top_twips": section.top_margin.twips,
                "bottom_cm": section.bottom_margin.cm,
                "bottom_twips": section.bottom_margin.twips,
                "right_cm": section.right_margin.cm,
                "right_twips": section.right_margin.twips,
                "width_cm": section.page_width.cm,
                "width_twips": section.page_width.twips,
                "height_cm": section.page_height.cm,
                "height_twips": section.page_height.twips,
                "pgnum_fmt": pgnum["fmt"],
                "pgnum_start": pgnum["start"],
                "footer_has_page_field": footer_info["has_page_field"],
                "footer_alignment": footer_info["alignment"],
            }
        )
    return sections_data


def _cell_bold_center_text(tc_element) -> tuple[bool, bool, bool]:
    """
    Untuk SATU cell tabel (w:tc): (fully_bold, is_center, has_text).
    fully_bold True hanya kalau SEMUA run berisi teks di cell itu bold.
    is_center dibaca dari w:jc tiap paragraf yang punya isi (cell
    header biasanya cuma 1 paragraf, tapi dijaga aman kalau lebih).
    has_text False -> cell kosong, diabaikan dari agregasi "semua cell
    header" supaya kolom kosong (jarang, tapi mungkin dipakai buat
    spasi visual) tidak salah menjatuhkan cek.
    """
    paras = tc_element.findall(qn("w:p"))
    bold_flags = []
    center_ok = True
    has_text = False
    for p in paras:
        pPr = p.find(qn("w:pPr"))
        jc = pPr.find(qn("w:jc")) if pPr is not None else None
        align_val = jc.get(qn("w:val")) if jc is not None else None
        p_has_text = False
        for r in p.findall(qn("w:r")):
            t = r.find(qn("w:t"))
            text = t.text if t is not None and t.text else ""
            if not text.strip():
                continue
            p_has_text = True
            has_text = True
            rPr = r.find(qn("w:rPr"))
            b = rPr.find(qn("w:b")) if rPr is not None else None
            is_bold = b is None or b.get(qn("w:val")) not in ("false", "0")
            bold_flags.append(is_bold)
        if p_has_text and align_val != "center":
            center_ok = False
    fully_bold = bool(bold_flags) and all(bold_flags)
    return fully_bold, center_ok, has_text


def _cell_has_shading(tc_element) -> bool:
    """True kalau cell punya w:shd dengan fill selain kosong/auto/putih."""
    tcPr = tc_element.find(qn("w:tcPr"))
    if tcPr is None:
        return False
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return False
    fill = shd.get(qn("w:fill"))
    return bool(fill) and fill.upper() not in ("AUTO", "FFFFFF")


def _extract_table_header_format(tbl_element) -> dict | None:
    """
    Format baris HEADER (baris pertama) tabel: bold, rata tengah,
    berlatar-belakang gelap -- sesuai anjuran panduan ("Judul pada
    tabel DAPAT dituliskan rata tengah, tebal dan berlatar-belakang
    agak gelap"). Return None kalau baris pertama tabel tidak punya
    cell berisi teks sama sekali (supaya comparator.py tahu untuk
    skip, bukan salah menganggap semua False).
    """
    rows = tbl_element.findall(qn("w:tr"))
    if not rows:
        return None
    cells = rows[0].findall(qn("w:tc"))
    if not cells:
        return None

    all_bold = True
    all_center = True
    any_shaded = False
    any_text = False
    for tc in cells:
        fully_bold, is_center, has_text = _cell_bold_center_text(tc)
        if has_text:
            any_text = True
            if not fully_bold:
                all_bold = False
            if not is_center:
                all_center = False
        if _cell_has_shading(tc):
            any_shaded = True

    if not any_text:
        return None

    return {
        "header_all_bold": all_bold,
        "header_all_center": all_center,
        "header_has_shading": any_shaded,
    }


def extract_table_params(doc) -> list[dict]:
    """
    Ekstrak parameter tabel Word (rata & posisi), ditelusuri lewat
    URUTAN ASLI elemen w:p / w:tbl di body XML -- doc.tables SENDIRI
    TIDAK dalam urutan dokumen dan tidak membawa info paragraf
    tetangganya, padahal itu dibutuhkan untuk cek posisi caption
    ("Tabel N.N ...") dan baris kosong sebelum/sesudah tabel.

    Ditambah "header_format": info bold/rata-tengah/shading baris
    pertama tabel, diambil LANGSUNG dari elemen w:tbl yang sama (bukan
    lewat doc.tables) -- konsisten dengan alasan kenapa modul ini sudah
    tidak pakai doc.tables sama sekali (lihat paragraf di atas).
    """
    body = doc.element.body

    # PENTING: index paragraf di sini ditelusuri lewat COUNTER berjalan
    # (bertambah tiap ketemu w:p LANGSUNG di body), BUKAN lewat
    # id(element) yang tadinya dipakai buat mencocokkan ke doc.paragraphs.
    # id(element) TERNYATA TIDAK RELIABLE: lxml membuat proxy Python
    # BARU tiap kali sebuah elemen diakses (mis. tiap iterasi
    # doc.paragraphs vs tiap iterasi body.iterchildren()), dan begitu
    # proxy lama di-garbage-collect, id()-nya (alamat memori) bisa
    # DIPAKAI ULANG oleh proxy baru yang lain -- collision, hasil
    # matching salah/acak. doc.paragraphs SENDIRI adalah persis daftar
    # w:p level-body yang sama urutannya dengan yang ditemui
    # body.iterchildren() di sini (paragraf DI DALAM tabel/textbox tidak
    # ikut doc.paragraphs), jadi counter sederhana ini sudah cukup dan
    # jauh lebih aman daripada mengandalkan identitas objek.
    tables_data = []
    prev_para_idx = None
    para_counter = -1
    table_seq = 0

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            para_counter += 1
            prev_para_idx = para_counter
        elif child.tag == qn("w:tbl"):
            table_seq += 1
            tblPr = child.find(qn("w:tblPr"))
            jc_el = tblPr.find(qn("w:jc")) if tblPr is not None else None
            alignment = jc_el.get(qn("w:val")) if jc_el is not None else None

            # Jumlah baris (w:tr) -- dipakai comparator.py buat MEMBEDAKAN
            # tabel data sungguhan dari tabel yang cuma dipakai sebagai
            # trik layout (bukan buat nyajiin data), makanya wajar TIDAK
            # punya caption "Tabel N.N ...":
            #   - tabel 1 baris x 2 kolom tanpa border, dipakai buat
            #     naruh nomor persamaan sejajar rumusnya (mis. "(2.2)")
            #     -- trik umum karena tab stop suka geser.
            #   - tabel 1 baris x 1 kolom BERBORDER, dipakai sebagai
            #     "kotak" pembungkus teks (prompt/kutipan/kode), bukan
            #     tabel data.
            # Tabel data beneran (perbandingan penelitian, metadata, dst)
            # SELALU py>1 baris (header + minimal 1 baris data), jadi
            # n_rows == 1 adalah sinyal kuat "ini bukan tabel data".
            n_rows = len(child.findall(qn("w:tr")))

            # header_format cuma relevan buat tabel DATA sungguhan
            # (n_rows > 1, sama seperti kriteria yang sudah dipakai
            # comparator.py buat exclude tabel trik-layout dari cek
            # caption) -- tabel 1-baris (mis. trik naruh nomor
            # persamaan) tidak punya konsep "baris header" sama
            # sekali, jadi sengaja dibiarkan None di sini juga.
            header_format = _extract_table_header_format(child) if n_rows > 1 else None

            # Paragraf pertama SETELAH tabel ini (dicari lewat sibling
            # XML langsung, bukan cuma "index paragraf berikutnya" --
            # supaya tetap benar walau tabel lain menyusul tanpa
            # paragraf pemisah di antaranya).
            next_idx = None
            sib = child.getnext()
            lookahead = para_counter
            while sib is not None:
                if sib.tag == qn("w:p"):
                    lookahead += 1
                    next_idx = lookahead
                    break
                if sib.tag == qn("w:tbl"):
                    break
                sib = sib.getnext()

            tables_data.append(
                {
                    "seq": table_seq,
                    "alignment": alignment,
                    "prev_para_idx": prev_para_idx,
                    "next_para_idx": next_idx,
                    "n_rows": n_rows,
                    "header_format": header_format,
                }
            )
            # Reset -- paragraf berikutnya (kalau ada) jadi "prev" milik
            # tabel SELANJUTNYA, bukan ikut nyangkut ke tabel ini lagi.
            prev_para_idx = None

    return tables_data


def extract_all(doc) -> dict:
    """Entry point tahap Extract Parameter: gabungkan semua parameter mentah."""
    return {
        "paragraphs": extract_paragraph_params(doc),
        "sections": extract_section_params(doc),
        "tables": extract_table_params(doc),
        # Level DOKUMEN, bukan per-section -- lihat docstring _get_mirror_margins().
        "mirror_margins": _get_mirror_margins(doc),
    }
