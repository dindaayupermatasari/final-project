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
- BUG DIPERBAIKI (laporan user, Buku_PA_Masyitha_Fahra.docx):
  (1) _resolve_effective_font_size() sekarang jatuh ke <w:docDefaults>
      (dan character style) kalau ukuran tidak di-set di run/paragraph
      style. Sebelumnya paragraf 12pt dengan satu run "." 10pt dilaporkan
      "10.0pt" karena run 12pt (mewarisi dari docDefaults) dianggap None
      dan dibuang dari perhitungan ukuran dominan.
  (2) Level heading dari numbering Word sekarang dihitung dari NOMOR YANG
      TAMPIL (_get_numbering_depth, membaca lvlText di numbering.xml),
      bukan dari ilvl mentah. Template dengan nomor bab literal
      ("2.%1" / "2.%1.%2") sebelumnya membuat "2.1.1 Judul" (ilvl 1)
      dibaca sebagai level "X.Y" dan salah dituntut KAPITAL SEMUA.
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

import bisect
import re
import types
from docx.oxml.ns import qn
from docx.shared import Length
from docx.table import Table

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

# BUG DIPERBAIKI (2026-09-12, laporan user): heading "LAMPIRAN A" /
# "LAMPIRAN B" dst level hierarkinya SAMA dengan "BAB N" -- dicetak
# 14pt bold center ALL CAPS sama seperti judul bab, dan caption
# Gambar/Tabel di dalamnya memakai HURUF lampiran sebagai pengganti
# angka bab (mis. "Tabel A.1", "Gambar B.1"), persis seperti caption
# "Tabel 4.1"/"Gambar 4.1" di dalam "BAB 4" -- itu sendiri buktinya
# lampiran memang level yang setara BAB, bukan subbab. Sebelum fix
# ini, "LAMPIRAN B" TIDAK dikenali di classify_paragraph() sama sekali
# (bukan match CHAPTER_TITLE_RE, bukan di FRONT_HEADINGS), jatuh ke
# jalur "heading tanpa nomor" dan keliru diklasifikasi
# "subsubchapter_title" (expected 12pt) padahal harusnya
# "chapter_title" (expected 14pt), sama seperti bug lama CHAPTER_
# TITLE_RE untuk BAB N di atas. Dua varian format diterima sama
# seperti BAB: "LAMPIRAN A" berdiri sendiri (nama lampiran di paragraf
# terpisah setelahnya) ATAU "LAMPIRAN A DOKUMENTASI ..." digabung satu
# baris.
LAMPIRAN_TITLE_ONLY_RE = re.compile(r"^LAMPIRAN(?:\s+[A-Z]+)?$")
# BUG DIPERBAIKI (2026-09-21, laporan user): dokumen dengan LAMPIRAN
# TUNGGAL (tidak dipecah per huruf A/B/dst, cuma satu heading
# "LAMPIRAN" polos) sebelumnya tidak match REGEX manapun di atas (baik
# LAMPIRAN_TITLE_ONLY_RE maupun LAMPIRAN_TITLE_RE versi lama, keduanya
# MEWAJIBKAN huruf lampiran setelah "LAMPIRAN"), jatuh ke jalur
# "heading tanpa nomor" -> keliru diklasifikasi "subsubchapter_title"
# (expected 12pt, padahal ini judul setingkat BAB, expected 14pt) DAN
# current_chapter TIDAK PERNAH pindah dari bab/daftar terakhir sebelum
# heading ini (mis. "DAFTAR PUSTAKA") -- akibatnya SELURUH isi lampiran
# (screenshot UI, caption "Gambar N. ...", dst) tetap tercatat di bawah
# chapter lama itu, LOLOS dari pengecualian _is_excluded_chapter() di
# comparator.py (yang cuma mengenali chapter berawalan literal
# "LAMPIRAN"), dan ikut kena SEMUA aturan format body/caption biasa --
# inilah penyebab utama banyaknya false-positive di bawah chapter
# "DAFTAR PUSTAKA" pada laporan nyata. Fix: huruf lampiran dibuat
# OPSIONAL di kedua regex -- "LAMPIRAN" polos kini juga dikenali
# setara level BAB, current_chapter dinormalisasi jadi "LAMPIRAN" saja
# (tetap lolos startswith("LAMPIRAN") di _is_excluded_chapter()).
LAMPIRAN_TITLE_RE = re.compile(r"^LAMPIRAN(?:\s+[A-Z]+)?(\s+\S.*)?$")

# Label nomor persamaan berdiri sendiri, mis. "(2.1)" di akhir baris
# rumus matematika. Pendek + bold -> kalau tidak dikecualikan di sini,
# ke-tangkap is_unnumbered_heading() sebagai heading tanpa nomor.
# Pola nomor identitas persamaan yang dipakai BERSAMA oleh jalur docx & PDF
# (EQUATION_NUMBER_RE, EQUATION_NUMBER_IN_TEXT_RE di bawah, dan
# _BARE_EQUATION_NUMBER_RE di pdf_extractor.py).
#
# DILONGGARKAN (2026-09-21, laporan user -- "nomor persamaan 3.1, 3.2
# sudah ada tapi tetap dibilang tidak ada"): pola lama cuma menerima
# "(2.1)" / "(Persamaan 2.1)" PERSIS. Nomor yang hasil ketikan Word/PDF
# sungguhan sering sedikit beda dari itu walau secara visual sama:
#   - spasi di dalam kurung: "( 3.1 )", "(3. 1)" (sering muncul di teks
#     PDF hasil ekspor rumus, atau spasi tipis U+2009 dari math zone);
#   - pemisah bab-urut selain titik: "(3-1)", "(3,1)";
#   - kata "Pers." selain "Persamaan";
#   - kurung fullwidth "（3.1）" dari IME.
# Semua varian itu tetap satu nomor identitas yang sah. Aturan posisinya
# TIDAK berubah: nomor wajib jadi karakter paling akhir di baris.
_EQ_NUM_LABEL = r"(?:(?:Persamaan|Pers\.?)\s*)?"
_EQ_NUM_DIGITS = r"\d+(?:\s*[.\-\u2010-\u2013,]\s*\d+)*"
_EQ_NUM_PAREN = r"[(\uff08]\s*" + _EQ_NUM_LABEL + _EQ_NUM_DIGITS + r"\s*[)\uff09]"
EQUATION_NUMBER_RE = re.compile(r"^" + _EQ_NUM_PAREN + r"$", re.IGNORECASE)

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
#
# BUG DIPERBAIKI (2026-09-12, laporan user): caption di dalam LAMPIRAN
# pakai HURUF sebagai pengganti angka bab (mis. "Tabel A.1", "Gambar
# B.1" -- lihat LAMPIRAN_TITLE_RE di atas). Sebelumnya regex ini HANYA
# menerima angka ("\d+\.\s*\d+"), jadi caption lampiran yang sudah
# 100% benar formatnya gagal match CAPTION_RE, jatuh ke ptype "body"
# default (expected 12pt) alih-alih "caption" (expected 11pt) --
# caption 11pt yang sudah benar keliru dilaporkan "font tidak sesuai
# untuk tipe body, seharusnya 12pt".
#
# BUG DIPERBAIKI (2026-09-20, laporan user): regex lama cuma mengecek
# AWALAN "Gambar/Tabel N.N", tanpa peduli apa yang muncul SESUDAHNYA --
# jadi kalimat body biasa yang KEBETULAN menyebut nomor gambar/tabel di
# awal kalimat ("Gambar 3.2 menampilkan mockup untuk para pengguna...",
# "Tabel 4.1 merupakan hasil pengujian...") ikut ke-match dan salah
# diklasifikasi ptype "caption" (expected 11pt) padahal ini body text
# biasa (expected 12pt) -- laporan comparator jadi nunjuk ke KALIMAT DI
# TENGAH PARAGRAF, bukan caption gambar/tabel yang sesungguhnya.
# Caption asli SELALU berupa judul singkat ("Gambar 3.2 Mockup Halaman
# Login"), bukan kalimat penuh -- kata pertama setelah nomornya selalu
# diawali huruf kapital (ikut konvensi judul), sedangkan kalimat body
# yang menyambung dari nomor gambar/tabel selalu diawali kata kerja
# rujukan huruf kecil ("menampilkan", "merupakan", "menunjukkan",
# "memperlihatkan", dst -- karena secara gramatikal itu predikat
# kalimat, bukan judul). Fix: wajibkan huruf kapital tepat setelah
# nomornya (dengan toleransi tanda baca pemisah opsional ":"/"-"/"."
# yang umum dipakai sebelum judul caption) supaya kalimat body yang
# hanya numpang lewat nomor gambar/tabel tidak lagi ke-match.
CAPTION_RE = re.compile(
    r"^(Gambar|Tabel)\s+(?:\d+|[A-Z]+)\.\s*\d+\s*[:.\-]?\s*[A-Za-z]"
)

# LAPIS PERTAHANAN KEDUA (2026-09-21, laporan user): syarat "huruf
# kapital setelah nomor" di CAPTION_RE di atas TIDAK CUKUP kalau
# mahasiswa (tidak sengaja/tidak konsisten) menuliskan kata kerja
# predikat rujukan itu dengan huruf kapital di awal, mis. "Gambar 4.14
# Menggambarkan perubahan daya..." -- ini tetap kalimat body biasa
# yang menyambung dari nomor gambar (predikat "menggambarkan"),
# BUKAN judul caption, walau kebetulan huruf pertamanya kapital.
# CAPTION_REFERRAL_WORDS berisi kata kerja/predikat rujukan yang
# lazim dipakai mahasiswa untuk MENYEBUT gambar/tabel di awal kalimat
# body (baik bentuk aktif "menampilkan" maupun pasif "ditampilkan"),
# dicocokkan case-insensitive terhadap kata PERTAMA setelah nomornya.
# Caption asli tidak pernah dibuka dengan kata-kata ini -- caption
# adalah judul benda ("Gambar 4.14 Grafik Perubahan Daya"), bukan
# kalimat berpredikat.
CAPTION_REFERRAL_WORDS = {
    "menampilkan",
    "merupakan",
    "menunjukkan",
    "memperlihatkan",
    "menggambarkan",
    "menyajikan",
    "memberikan",
    "menjelaskan",
    "menerangkan",
    "menjabarkan",
    "mengilustrasikan",
    "menampakkan",
    "diperlihatkan",
    "ditampilkan",
    "ditunjukkan",
    "disajikan",
    "digambarkan",
    "dijelaskan",
    "berisi",
    "terdiri",
    "adalah",
    "yaitu",
    "berikut",
}

# Caption asli SELALU judul singkat (nama gambar/tabel), tidak pernah
# sepanjang kalimat penuh. Kalimat body yang lolos dua filter lain
# (kapital + bukan kata rujukan di atas) tapi keseluruhan teksnya jauh
# lebih panjang dari sebuah judul hampir pasti tetap kalimat biasa,
# bukan caption -- dipakai sebagai jaring pengaman terakhir.
CAPTION_MAX_WORDS = 20

# Kata benda berawalan me-/di-/ter-/ber-/ke- yang lazim membuka judul
# caption berhuruf kecil, dikecualikan dari filter "awalan kata kerja".
CAPTION_LOWERCASE_NOUN_EXCEPTIONS = {
    "diagram",
    "dimensi",
    "desain",
    "direktori",
    "berkas",
    "beranda",
    "mekanisme",
    "metode",
    "metodologi",
    "medan",
    "media",
    "memori",
    "mesin",
    "menu",
    "meter",
    "kerangka",
    "keluaran",
    "keseluruhan",
    "kesalahan",
    "kecepatan",
    "ketinggian",
    "kemiringan",
    "keadaan",
    "kondisi",
    "konfigurasi",
    "kurva",
}


def is_caption_title(text: str) -> bool:
    """
    Wrapper lebih ketat di atas CAPTION_RE mentah -- dipakai di SEMUA
    titik keputusan yang butuh tahu "apakah baris ini caption
    Gambar/Tabel sungguhan" (klasifikasi ptype "caption" di
    classify_paragraph(), dan heuristik penggabungan baris caption
    bersambung di pdf_extractor.py), supaya definisi "caption" KONSISTEN
    di semua titik, bukan cuma di ptype akhir. Kalau prefix "Gambar/
    Tabel N.N [huruf kapital]" cocok tapi kata PERTAMA setelah nomornya
    ada di CAPTION_REFERRAL_WORDS (case-insensitive), atau keseluruhan
    teksnya melebihi CAPTION_MAX_WORDS, tetap ditolak -- lihat
    CHANGELOG di atas untuk alasan lengkap kenapa syarat kapital saja
    tidak cukup.
    """
    m = CAPTION_RE.match(text)
    if not m:
        return False
    # m.end() - 1 adalah index huruf kapital yang sudah ke-match --
    # sertakan kembali supaya first_word tidak kehilangan huruf itu.
    remainder = text[m.end() - 1 :].strip()
    first_word_match = re.match(r"[A-Za-z]+", remainder)
    if first_word_match and first_word_match.group(0).lower() in CAPTION_REFERRAL_WORDS:
        return False
    if len(text.split()) > CAPTION_MAX_WORDS:
        return False
    # BUG DIPERBAIKI (2026-09-21, laporan user): caption yang kata
    # pertama judulnya huruf kecil (mis. "Gambar 4. 47 grafik arus
    # listrik simulasi statis") tertolak oleh syarat kapital dan jatuh
    # ke "body" (expected 12pt), padahal fontnya 11pt sudah benar.
    # CAPTION_RE sekarang menerima huruf kecil, tapi khusus kasus
    # huruf kecil kita perketat supaya kalimat body tetap tidak lolos:
    # (1) tidak boleh berakhir tanda titik/koma/titik dua (kalimat
    # penuh), (2) kata pertama tidak boleh berimbuhan kata kerja
    # (me-/di-/ter-/ber-) kecuali noun yang umum di judul caption.
    if remainder[:1].islower():
        if re.search(r"[.,:;]\s*$", text):
            return False
        fw = first_word_match.group(0).lower() if first_word_match else ""
        if (
            re.match(r"^(me|mem|men|meng|meny|di|ter|ber|ke)[a-z]{4,}$", fw)
            and fw not in CAPTION_LOWERCASE_NOUN_EXCEPTIONS
        ):
            return False
    return True


# Baris "Sumber: ..." / "Sumber : ..." di bawah gambar/tabel kutipan --
# opsional (panduan: "penulis DAPAT menyebutkan sumber referensi"),
# jadi cuma dicek FORMATnya (rata tengah, 11pt) KALAU baris ini ada,
# tidak pernah dianggap wajib ada.
# BUG DIPERBAIKI (jalur PDF -- laporan user: paragraf body biasa yang
# KEBETULAN diawali kata "sumber." huruf kecil (sisa kalimat yang
# terpotong baris/halaman, mis. "...berasal dari berbagai\nsumber.
# Risiko terjadinya kesalahan saat...") ke-klasifikasi keliru sebagai
# ptype "source_citation" dan dicek pakai aturan ukuran font caption
# (11pt) padahal ini body text biasa (12pt): re.IGNORECASE di regex
# lama membuatnya cocok ke "sumber." huruf kecil manapun, termasuk kata
# "sumber" biasa di tengah/awal kalimat. Baris kutipan "Sumber: ..."/
# "Sumber. ..." yang SUNGGUHAN selalu ditulis dengan S kapital
# (konvensi baku akademik) -- fix: buang re.IGNORECASE supaya regex
# ini HANYA cocok ke "Sumber" berhuruf S kapital di awal baris.
SOURCE_CITATION_RE = re.compile(r"^Sumber\s*[:.]")

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
#
# ATURAN DARI PAK ALI (dosen pembimbing, klarifikasi lisan): nomor
# persamaan HARUS jadi karakter paling akhir di baris tsb -- begitu
# tutup kurung ")", langsung ganti baris, TIDAK BOLEH ada spasi/tab
# sisa setelahnya. Ini yang dipakai buat MENDETEKSI apakah nomornya
# benar-benar rata kanan (bukan cuma "rata kanan visual" krn dipepet
# banyak spasi manual, yang teksnya tetap akan ada residu spasi
# setelah karakter ")").
#
# BUG DIPERBAIKI: pola lama pakai "\s*$" (spasi trailing OPSIONAL
# sebelum akhir teks) -- ini SALAH karena tetap match walau ada sisa
# spasi/tab setelah ")", padahal itu justru kasus yang harus DITOLAK
# menurut aturan di atas. Diganti jadi anchor "$" langsung tanpa
# toleransi whitespace apapun setelah tutup kurung.
#
# LONGGARAN (atas permintaan user, 2026-09-07): label boleh ditulis
# "(2.1)" ATAU "(Persamaan 2.1)" -- kata "Persamaan" di dalam kurung
# label dianggap sah, tidak wajib dibuang. Case-insensitive supaya
# "persamaan"/"PERSAMAAN" juga diterima.
#
# DILONGGARKAN LAGI (laporan user -- persamaan bernomor "4.1" TANPA
# kurung tetap dibilang tidak bernomor): nomor identitas sekarang juga
# boleh berupa angka bab-urut telanjang ("4.1", "3-2"), TAPI hanya kalau
# posisinya jelas terpisah sebagai label di ujung kanan -- yaitu tepat
# sesudah tab, sesudah >= 2 spasi, atau berdiri sendiri di awal teks.
# Syarat posisi ini yang mencegah ujung rumus biasa seperti "... = 0.5"
# salah dianggap nomor persamaan (di situ cuma ada SATU spasi).
_EQ_NUM_BARE = r"\d+\s*[.\-\u2010-\u2013,]\s*\d+"
EQUATION_NUMBER_BARE_RE = re.compile(r"^" + _EQ_NUM_BARE + r"$")
EQUATION_NUMBER_IN_TEXT_RE = re.compile(
    r"(?:" + _EQ_NUM_PAREN + r"|(?:^|(?<=\t)|(?<=\s\s))" + _EQ_NUM_BARE + r")$",
    re.IGNORECASE,
)

_EQ_SPACE_CHARS_RE = re.compile(r"[\u00a0\u2000-\u200a\u202f\u205f\u3000]")
_EQ_INVISIBLE_CHARS_RE = re.compile(r"[\u200b-\u200d\u2060\ufeff]")


def normalize_equation_text(text: str) -> str:
    """
    Rapikan teks untuk pencarian nomor identitas persamaan: spasi
    "aneh" (NBSP, spasi tipis, dll) jadi spasi biasa, karakter tak
    terlihat (zero-width) dibuang, lalu trim ujung-ujungnya.

    HANYA dipakai untuk teks hasil ekstraksi rumus (math zone, field,
    teks PDF) yang memang sering membawa karakter semacam itu -- BUKAN
    untuk full_text w:t biasa, yang tetap dicek KETAT (aturan Pak Ali:
    tidak boleh ada sisa spasi/tab setelah tutup kurung).
    """
    text = _EQ_SPACE_CHARS_RE.sub(" ", text or "")
    text = _EQ_INVISIBLE_CHARS_RE.sub("", text)
    return text.strip()


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
    # BUG DIPERBAIKI (2026-09-12, laporan user): "BIODATA PENULIS" --
    # halaman penutup standalone di akhir dokumen (setelah semua
    # LAMPIRAN), BUKAN bagian dari isi lampiran/bab manapun sehingga
    # tidak seharusnya diperlakukan sebagai subbab. Namanya "FRONT_
    # HEADINGS" menyesatkan (kebanyakan isinya memang heading bagian
    # depan), tapi secara fungsi set ini dipakai untuk SEMUA heading
    # standalone tanpa nomor yang levelnya setara BAB, di posisi
    # manapun di dokumen -- "DAFTAR PUSTAKA" sendiri juga baru muncul
    # di BELAKANG semua BAB, bukan di depan, dan sudah diperlakukan
    # begitu sebelum fix ini. Tanpa entry ini, "BIODATA PENULIS" tidak
    # match apa pun di classify_paragraph() (bukan BAB, bukan LAMPIRAN,
    # tidak ada di set ini) sehingga jatuh ke jalur "heading tanpa
    # nomor" dan keliru diklasifikasi subsubchapter_title (12pt),
    # persis kasus yang sama seperti bug LAMPIRAN sebelum diperbaiki.
    "BIODATA PENULIS",
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


def _find_numbering_level(part, num_id, ilvl):
    """
    Cari elemen <w:lvl> milik (numId, ilvl) di numbering.xml. Memperhitungkan
    <w:lvlOverride> di <w:num> kalau ada, lalu jatuh ke definisi
    <w:abstractNum> aslinya. Return None kalau tidak ketemu / dokumen tidak
    punya numbering part.
    """
    try:
        numbering = part.numbering_part.element
    except Exception:
        return None
    num = next(
        (
            n
            for n in numbering.findall(qn("w:num"))
            if n.get(qn("w:numId")) == str(num_id)
        ),
        None,
    )
    if num is None:
        return None
    for ov in num.findall(qn("w:lvlOverride")):
        if ov.get(qn("w:ilvl")) == str(ilvl):
            lvl = ov.find(qn("w:lvl"))
            if lvl is not None:
                return lvl
    abs_id_el = num.find(qn("w:abstractNumId"))
    if abs_id_el is None:
        return None
    abs_id = abs_id_el.get(qn("w:val"))
    for an in numbering.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == abs_id:
            for lvl in an.findall(qn("w:lvl")):
                if lvl.get(qn("w:ilvl")) == str(ilvl):
                    return lvl
    return None


def _ilvl_from_style_link(part, num_id, style_id):
    """
    Kalau numPr di STYLE cuma menyebut numId tanpa w:ilvl, level-nya
    ditentukan oleh <w:pStyle> di dalam definisi level list (tautan
    style <-> level). Return None kalau tidak ada tautan.
    """
    try:
        numbering = part.numbering_part.element
    except Exception:
        return None
    num = next(
        (
            n
            for n in numbering.findall(qn("w:num"))
            if n.get(qn("w:numId")) == str(num_id)
        ),
        None,
    )
    if num is None:
        return None
    abs_id_el = num.find(qn("w:abstractNumId"))
    if abs_id_el is None:
        return None
    for an in numbering.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == abs_id_el.get(qn("w:val")):
            for lvl in an.findall(qn("w:lvl")):
                ps = lvl.find(qn("w:pStyle"))
                if ps is not None and ps.get(qn("w:val")) == style_id:
                    return int(lvl.get(qn("w:ilvl")))
    return None


def _get_numbering_depth(para) -> int | None:
    """
    KEDALAMAN heading (1 = setingkat BAB, 2 = "X.Y", 3+ = "X.Y.Z" dst)
    berdasarkan NOMOR YANG BENAR-BENAR DITAMPILKAN Word, bukan sekadar
    w:ilvl mentah.

    BUG DIPERBAIKI (laporan user -- "2.1.1 Internet Service Provider"
    dituntut KAPITAL SEMUA seperti judul "2.1 ..."): kode lama menganggap
    ilvl 0 = BAB, ilvl 1 = "X.Y", ilvl 2+ = "X.Y.Z". Asumsi itu hanya
    benar kalau nomor BAB ikut jadi bagian dari list. Di banyak template
    (termasuk dokumen ini) nomor bab ditulis LITERAL di lvlText per bab:
        ilvl 0 -> lvlText "2.%1"     (tampil "2.1"   = level X.Y)
        ilvl 1 -> lvlText "2.%1.%2"  (tampil "2.1.1" = level X.Y.Z)
    sehingga judul "2.1.1 ..." (ilvl 1) keliru dibaca sebagai "X.Y" dan
    wajib ALL CAPS, padahal panduan hanya mewajibkan ALL CAPS di level
    "X.Y"; level "X.Y.Z" tidak wajib.

    Kedalaman dihitung dari jumlah komponen bernomor di lvlText (dipisah
    titik/strip; komponen literal "2" dan placeholder "%1" sama-sama
    dihitung), lalu diambil max() dengan ilvl+1 supaya lvlText yang tidak
    bernomor (mis. numFmt none / "(%4)") tidak turun di bawah perilaku
    lama. Return None kalau paragraf tidak punya numbering.
    """
    pPr = para._p.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
    linked_style_id = None
    if numPr is None:
        style = para.style
        while style is not None:
            style_pPr = style.element.find(qn("w:pPr"))
            style_numPr = (
                style_pPr.find(qn("w:numPr")) if style_pPr is not None else None
            )
            if style_numPr is not None:
                numPr = style_numPr
                linked_style_id = style.style_id
                break
            style = style.base_style
    if numPr is None:
        return None

    num_id_el = numPr.find(qn("w:numId"))
    num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else None
    if num_id == "0":
        return None  # numbering dimatikan eksplisit

    ilvl_el = numPr.find(qn("w:ilvl"))
    if ilvl_el is not None:
        ilvl = int(ilvl_el.get(qn("w:val")))
    else:
        ilvl = None
        if num_id is not None and linked_style_id is not None:
            ilvl = _ilvl_from_style_link(para.part, num_id, linked_style_id)
        if ilvl is None:
            ilvl = 0

    lvl_text = ""
    if num_id is not None:
        lvl = _find_numbering_level(para.part, num_id, ilvl)
        if lvl is not None:
            lt = lvl.find(qn("w:lvlText"))
            lvl_text = (lt.get(qn("w:val")) if lt is not None else "") or ""

    parts = [p for p in re.split(r"[.\-]", lvl_text) if re.search(r"%\d|\d", p)]
    return max(len(parts), ilvl + 1)


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

    # LAMPIRAN setara BAB (lihat komentar LAMPIRAN_TITLE_RE di atas) --
    # dicek di jalur yang sama, SEBELUM jalur heading tanpa nomor,
    # supaya tidak ke-tangkap sebagai subsubchapter_title.
    if LAMPIRAN_TITLE_RE.match(upper) and stripped == upper:
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
        # Pakai KEDALAMAN nomor yang tampil (lihat _get_numbering_depth),
        # BUKAN ilvl mentah: template dengan nomor bab literal di lvlText
        # ("2.%1" / "2.%1.%2") membuat ilvl 0 = "X.Y" dan ilvl 1 = "X.Y.Z".
        depth = _get_numbering_depth(para)
        if depth == 1:
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
        elif depth == 2:
            return "subchapter_title"
        elif depth is not None and depth >= 3:
            return "subsubchapter_title"

        # --- Jalur 3: heading tanpa nomor sama sekali ---
        # Disatukan ke "subsubchapter_title" karena aturan formatnya
        # (Title Case + bold) identik dengan level "X.Y.Z".
        if is_unnumbered_heading(para):
            return "subsubchapter_title"

    if is_caption_title(stripped):
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


def _resolve_effective_all_caps(run, paragraph) -> bool:
    """
    True kalau run ditampilkan HURUF KAPITAL SEMUA oleh Word lewat
    properti font "All caps" (w:caps) -- teks aslinya di XML bisa saja
    huruf biasa ("Contextual Analysis ..."), tapi yang tercetak/terlihat
    KAPITAL. Ditelusuri dari run -> character style -> paragraph style
    -> base style, sama seperti _resolve_effective_bold().

    Dipakai KHUSUS cek "huruf kapital semua" di halaman sampul (lihat
    elem["is_visual_upper"]) -- tanpa ini, judul yang tampil ALL CAPS
    lewat style/w:caps dianggap "bukan kapital" karena para.text-nya
    huruf biasa.
    """
    if run.font.all_caps is not None:
        return bool(run.font.all_caps)
    style = run.style
    while style is not None:
        if style.font.all_caps is not None:
            return bool(style.font.all_caps)
        style = style.base_style
    style = paragraph.style
    while style is not None:
        if style.font.all_caps is not None:
            return bool(style.font.all_caps)
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


def _get_doc_default_font_size_pt(paragraph) -> float | None:
    """
    Ukuran font default dokumen dari <w:docDefaults>/<w:rPrDefault>/<w:rPr>/
    <w:sz> di styles.xml (half-points -> pt). None kalau tidak didefinisikan.
    """
    try:
        styles_el = paragraph.part.styles.element
    except Exception:
        return None
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        return None
    rpr_default = dd.find(qn("w:rPrDefault"))
    rpr = rpr_default.find(qn("w:rPr")) if rpr_default is not None else None
    sz = rpr.find(qn("w:sz")) if rpr is not None else None
    if sz is None:
        return None
    try:
        return int(sz.get(qn("w:val"))) / 2.0
    except (TypeError, ValueError):
        return None


def _resolve_effective_font_size(run, paragraph) -> float | None:
    """
    Simetris dengan _resolve_effective_bold(). Urutan pewarisan Word:
    run langsung -> character style -> paragraph style (+ base_style)
    -> <w:docDefaults>.

    BUG DIPERBAIKI (laporan user -- "Ukuran font 10.0pt tidak sesuai"
    padahal paragraf aslinya 12pt): run yang ukurannya TIDAK di-set
    langsung dan tidak ada di style chain (mewarisi 12pt dari
    docDefaults) dulu dikembalikan None lalu di-SKIP comparator. Kalau di
    paragraf yang sama ada satu run pendek yang ukurannya eksplisit beda
    (mis. tanda titik "." 10pt di akhir kalimat), run itulah SATU-SATUNYA
    yang dihitung sehingga "ukuran dominan" paragraf salah jadi 10pt.
    Dengan fallback docDefaults, run tanpa ukuran eksplisit ikut dihitung
    12pt dan ukuran dominan (character-weighted) menjadi benar.
    """
    if run.font.size is not None:
        return run.font.size.pt

    char_style = run.style
    while char_style is not None:
        if char_style.font.size is not None:
            return char_style.font.size.pt
        char_style = char_style.base_style

    style = paragraph.style
    while style is not None:
        if style.font.size is not None:
            return style.font.size.pt
        style = style.base_style

    return _get_doc_default_font_size_pt(paragraph)


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
    True kalau paragraf memuat rumus matematika OMML (m:oMath) DI MANA
    PUN -- termasuk equation INLINE kecil yang cuma jadi bagian dari
    kalimat prosa biasa (mis. "...cluster (delta_i^2), yang..."). OMML
    TIDAK PERNAH muncul di para.text (python-docx cuma membaca w:t) --
    jadi paragraf persamaan yang tidak diberi nomor identitas apa pun
    akan terlihat seperti paragraf kosong tanpa deteksi ini, dan lolos
    tanpa pernah dicek sama sekali.

    DIPAKAI HANYA untuk keputusan "paragraf ini kosong atau tidak" (di
    extract_all, supaya paragraf yang isinya cuma OMML -- inline maupun
    display -- tidak ikut ke-skip sebagai paragraf kosong) dan untuk
    mengisi field equation_alignment/equation_math_text/equation_full_text
    (yang aman tetap dihitung longgar, karena field itu cuma dipakai
    comparator.py KALAU ptype elemennya "equation"). BUKAN dipakai untuk
    keputusan ptype "equation" itu sendiri -- lihat _has_display_equation()
    di bawah untuk itu.
    """
    return next(para._p.iter(qn("m:oMath")), None) is not None


def _has_display_equation(para) -> bool:
    """
    True kalau paragraf memuat rumus matematika dalam mode "Display"
    Word -- yaitu dibungkus <m:oMathPara> (equation yang di-set berdiri
    sendiri di baris sendiri, lewat panah kecil di equation -> "Change
    to Display", atau otomatis kalau equation-nya besar/mandiri) --
    BUKAN sekadar mengandung <m:oMath> di mana pun.

    BUG DIPERBAIKI (2026-09-22, laporan user -- kalimat prosa biasa yang
    cuma menyisipkan SATU simbol matematis kecil di tengah kalimat, mis.
    "menghitung variansi masing-masing cluster (delta_i^2), yang
    menggambarkan..." -- 219 karakter, MAYORITAS teks biasa -- salah
    diklasifikasikan SELURUH paragrafnya jadi ptype "equation" hanya
    karena ada SATU <m:oMath> inline di dalamnya. Akibatnya paragraf
    prosa itu dipaksa tunduk ke aturan display-equation (rata kiri, 1
    baris kosong atas-bawah, wajib nomor identitas) yang sama sekali
    tidak relevan buat dia, dan SELALU gagal karena memang bukan
    persamaan display.

    Pembeda yang dipakai di sini SAMA PERSIS dengan yang sudah dipakai
    _get_equation_alignment(): Word HANYA membungkus OMML dengan
    <m:oMathPara> (dan mengisi <m:oMathParaPr>) kalau equation-nya
    di-set mode Display. Equation inline yang menyatu di tengah kalimat
    TIDAK PERNAH punya <m:oMathPara> -- OMML-nya langsung anak <w:r>
    tanpa pembungkus itu. Verified langsung dari dokumen mahasiswa
    nyata: kalimat prosa ber-simbol inline di atas -> False; persamaan
    display sungguhan (mis. "(3.1)") -> True.

    Dipakai KHUSUS untuk keputusan ptype "equation" di extract_all --
    supaya cuma persamaan display sungguhan yang tunduk ke aturan
    format persamaan, dan paragraf prosa yang kebetulan menyisipkan
    simbol matematis inline tetap diperlakukan (dan dicek) sebagai
    body text biasa.
    """
    return next(para._p.iter(qn("m:oMathPara")), None) is not None


def _prose_text(el) -> str:
    """
    Teks <w:t> BIASA (di luar rumus OMML) dari elemen XML apa pun
    (paragraf, sel tabel, dst), termasuk yang dibungkus w:fldSimple/
    w:hyperlink/w:ins -- yang TIDAK terbaca para.text python-docx.
    <m:t> (teks di dalam math zone) sengaja tidak ikut: itu bagian rumus,
    bukan kalimat prosa. Teks w:del dilewati.
    """
    skip = qn("w:delText")
    return "".join(
        t.text or ""
        for t in el.iter(qn("w:t"))
        if t.tag != skip
    )


def _is_standalone_inline_equation(para) -> bool:
    """
    True kalau paragraf berisi rumus OMML INLINE (tanpa pembungkus
    <m:oMathPara>) yang BERDIRI SENDIRI, yaitu satu-satunya isi
    paragraf selain (opsional) tab dan nomor identitas "(3.1)".

    KENAPA PERLU (2026-11-09, laporan user -- pengecekan rumus di PDF
    jalan, di DOCX tidak): _has_display_equation() HANYA True untuk
    rumus mode Display (<m:oMathPara>). Padahal cara sangat umum
    mahasiswa menulis persamaan + nomor di Word adalah:
        [rumus inline] <TAB kanan> (3.1)
    Rumus semacam itu TIDAK pernah dibungkus <m:oMathPara> (Word cuma
    membungkus kalau equation di-set Display / berdiri di baris
    sendiri tanpa teks lain), sehingga seluruh paragrafnya jatuh ke
    ptype "body" dan LOLOS TANPA cek persamaan sama sekali (rata kiri,
    nomor identitas, tebal, baris kosong atas-bawah). Jalur PDF tidak
    kena masalah ini karena mendeteksi persamaan dari tata letak
    visual, bukan dari struktur XML.

    Pembeda dari kalimat prosa yang menyisipkan simbol inline (kasus
    yang diperbaiki 2026-09-22 di _has_display_equation): di sini
    prosa SAMA SEKALI TIDAK ADA -- teks <w:t> paragraf kosong atau
    persis sebuah nomor identitas. Kalimat prosa pasti punya teks lain
    sehingga tetap ptype "body".
    """
    if _has_display_equation(para) or not _has_equation(para):
        return False
    if _has_drawing(para):
        return False
    rest = normalize_equation_text(_prose_text(para._p))
    return rest == "" or bool(EQUATION_NUMBER_RE.match(rest))


def _math_run_is_bold(m_r) -> bool:
    """
    True kalau <m:r> (math run) dicetak tebal. Word menyimpan tebal di
    math zone lewat <m:rPr><m:sty m:val="b"|"bi"/> ATAU lewat <w:rPr>
    <w:b/> di dalam <m:r> -- keduanya dibaca di sini.
    """
    rpr = m_r.find(qn("m:rPr"))
    if rpr is not None:
        sty = rpr.find(qn("m:sty"))
        if sty is not None and sty.get(qn("m:val")) in ("b", "bi"):
            return True
    wrpr = m_r.find(qn("w:rPr"))
    if wrpr is not None:
        b = wrpr.find(qn("w:b"))
        if b is not None and b.get(qn("w:val")) not in ("0", "false", "off"):
            return True
    return False


def _equation_number_bold(para) -> bool | None:
    """
    Apakah NOMOR identitas persamaan yang terjebak di math zone
    (mis. "rumus#(3.1)" otomatis Word -> "3.1" di <m:r> dalam <m:d>)
    dicetak tebal?

    Return True / False, atau None kalau tidak bisa dinilai (nomor tidak
    ditemukan di akhir teks, atau nomor ada di <w:r> biasa -- yang itu
    sudah dicek lewat elem["runs"] di comparator.py).

    Sebelumnya cek tebal SENGAJA dilewati untuk nomor di math zone
    (offset <m:t> tidak sejajar dengan elem["runs"]), sehingga di DOCX
    nomor "(3.1)" yang tidak tebal tidak pernah ke-flag, padahal di
    jalur PDF ke-flag. Di sini dibangun teks + penanda tebal PER KARAKTER
    dengan urutan jalan yang sama seperti _equation_full_text(), lalu
    span nomor dicari dengan EQUATION_NUMBER_IN_TEXT_RE.
    """
    _equation_full_text(para)  # memastikan _EQ_TEXT_SKIP_TAGS terisi
    chars: list[str] = []
    flags: list[bool | None] = []

    def chr_prop(dpr, name, default):
        if dpr is None:
            return default
        el = dpr.find(qn(f"m:{name}"))
        if el is None:
            return default
        val = el.get(qn("m:val"))
        return default if val is None else val

    def emit(text: str, flag):
        chars.append(text)
        flags.extend([flag] * len(text))

    def walk(el, flag):
        tag = el.tag
        if tag in (qn("w:t"), qn("m:t")):
            emit(el.text or "", flag)
            return
        if tag == qn("w:tab"):
            emit("\t", None)
            return
        if tag in _EQ_TEXT_SKIP_TAGS:
            return
        if tag == qn("m:r"):
            flag = _math_run_is_bold(el)
        if tag == qn("m:d"):
            dpr = el.find(qn("m:dPr"))
            emit(chr_prop(dpr, "begChr", "("), None)
            sep = chr_prop(dpr, "sepChr", "|")
            for k, e in enumerate(el.findall(qn("m:e"))):
                if k:
                    emit(sep, None)
                walk(e, flag)
            emit(chr_prop(dpr, "endChr", ")"), None)
            return
        for child in el:
            walk(child, flag)

    walk(para._p, None)
    text = "".join(chars)
    m = EQUATION_NUMBER_IN_TEXT_RE.search(text)
    if not m:
        return None
    vals = [
        flags[i]
        for i in range(*m.span())
        if flags[i] is not None and text[i].strip()
    ]
    if not vals:
        return None
    return all(vals)


def _get_equation_alignment(para) -> str | None:
    """
    Rata (justification) KHUSUS milik objek Equation Word, dibaca dari
    m:oMathParaPr/m:jc -- BUKAN w:pPr/w:jc paragraf biasa yang dibaca
    _get_paragraph_alignment().

    BUG YANG DIPERBAIKI (2026-09-08, laporan mahasiswa): Word menyimpan
    rata visual sebuah Equation di tempat TERPISAH dari rata paragraf
    pembungkusnya. Begitu mahasiswa insert equation lalu equation itu
    di-set mode "Display" (klik panah kecil di equation -> "Change to
    Display" -- ini otomatis terjadi kalau equation-nya besar/mandiri
    di baris sendiri, seperti rumus ber-pecahan pada laporan ini), Word
    menyimpan justifikasi VISUAL rumus itu di m:oMathParaPr/m:jc DI
    DALAM elemen m:oMathPara, dengan default "centerGroup" (rumus
    diposisikan rata TENGAH halaman) -- SAMA SEKALI TIDAK MENYENTUH
    w:pPr/w:jc paragrafnya sendiri, yang bisa saja tetap "left" (atau
    malah warisan dari style, bukan di-set eksplisit).

    Akibatnya sebelum fix ini: cek "rata kiri" comparator.py yang cuma
    membaca elem["alignment"] (asal dari w:pPr/w:jc) LOLOS padahal
    rumusnya SECARA VISUAL rata tengah -- persis studi kasus laporan
    mahasiswa yang rumusnya rata tengah tapi tidak pernah ke-flag.

    Return None kalau paragraf ini tidak (atau belum) punya
    m:oMathParaPr/m:jc sama sekali (mis. equation inline kecil yang
    tidak pernah di-"Display"-kan, ikut w:pPr/w:jc paragraf biasa apa
    adanya) -- pemanggil (extract_all -> elements[i]) harus fallback ke
    _get_paragraph_alignment() dalam kasus ini.

    Nilai m:jc yang mungkin: "left", "right", "center", "centerGroup".
    "centerGroup" (rumus & nomor identitas diposisikan sbg SATU grup di
    tengah horizontal -- default Word) DINORMALISASI jadi "center" di
    sini, karena efeknya SAMA-SAMA bukan rata kiri dari sudut pandang
    comparator.py.
    """
    om_para_pr = next(para._p.iter(qn("m:oMathParaPr")), None)
    if om_para_pr is None:
        return None
    jc_el = om_para_pr.find(qn("m:jc"))
    if jc_el is None:
        return None
    val = jc_el.get(qn("m:val"))
    return "center" if val == "centerGroup" else val


_EQ_TEXT_SKIP_TAGS = None


def _equation_full_text(para) -> str:
    """
    Teks paragraf persamaan dalam URUTAN DOKUMEN yang meniru apa yang
    dilihat pembaca di Word: gabungan <w:t>, tab, <m:t>, DAN kurung
    pembungkus objek <m:d> (delimiter).

    BUG YANG DIPERBAIKI (2026-09-21, laporan user -- nomor "(3.1)",
    "(3.2)" sudah ada tapi dibilang tidak ada). Dua penyebab nyata,
    dua-duanya sudah direproduksi di dokumen uji:

    1. Nomor persamaan otomatis Word ("rumus#(3.1)" -> Enter) disimpan
       sebagai <m:eqArr> berisi "#" lalu objek <m:d> yang isinya cuma
       "3.1". Kurung "(" dan ")" TIDAK ADA di <m:t> mana pun -- itu
       properti bawaan <m:d> (begChr/endChr default), sehingga
       _equation_math_text() cuma menghasilkan "...#3.1" dan regex
       nomor (yang mewajibkan kurung) gagal cocok.
    2. Nomor yang dibuat dengan field (mis. "(3." + SEQ Persamaan +
       ")"). Hasil field tersimpan sebagai <w:fldSimple><w:r><w:t>,
       yang BUKAN anak langsung <w:p>, jadi para.text/para.runs
       python-docx melewatinya -- full_text cuma "(3.)".

    Fungsi ini berjalan rekursif di seluruh anak paragraf, sehingga
    kedua kasus (dan pembungkus lain seperti w:ins, w:hyperlink,
    w:sdt) terbaca lengkap. Teks yang dihapus (w:del), instruksi field
    (w:instrText), dan properti (rPr/pPr/...) dilewati.

    Dipakai comparator.py HANYA sebagai fallback mencari nomor
    identitas, sama seperti _equation_math_text().
    """
    global _EQ_TEXT_SKIP_TAGS
    if _EQ_TEXT_SKIP_TAGS is None:
        _EQ_TEXT_SKIP_TAGS = {
            qn(t)
            for t in (
                "w:pPr",
                "w:rPr",
                "w:del",
                "w:delText",
                "w:instrText",
                "m:ctrlPr",
                "m:rPr",
                "m:dPr",
                "m:eqArrPr",
                "m:oMathParaPr",
            )
        }

    def _chr_prop(dpr, name, default):
        if dpr is None:
            return default
        el = dpr.find(qn(f"m:{name}"))
        if el is None:
            return default
        val = el.get(qn("m:val"))
        return default if val is None else val

    def walk(el) -> str:
        tag = el.tag
        if tag in (qn("w:t"), qn("m:t")):
            return el.text or ""
        if tag == qn("w:tab"):
            return "\t"
        if tag in _EQ_TEXT_SKIP_TAGS:
            return ""
        if tag == qn("m:d"):
            dpr = el.find(qn("m:dPr"))
            beg = _chr_prop(dpr, "begChr", "(")
            end = _chr_prop(dpr, "endChr", ")")
            sep = _chr_prop(dpr, "sepChr", "|")
            inner = sep.join(walk(e) for e in el.findall(qn("m:e")))
            return beg + inner + end
        return "".join(walk(child) for child in el)

    return walk(para._p)


def _equation_math_text(para) -> str:
    """
    Gabungan literal teks dari SEMUA elemen <m:t> (math text run) di
    dalam OMML rumus milik paragraf ini.

    BUG YANG DIPERBAIKI (2026-09-08, laporan mahasiswa): kalau nomor
    identitas persamaan (mis. "(Persamaan 2.2)") diketik SAMBIL MASIH
    berada di dalam math zone Word (kursor belum di-klik keluar dari
    mode Equation sebelum mengetik nomornya), teks itu ikut terbungkus
    sebagai bagian dari OMML dan tersimpan sebagai <m:t>, BUKAN sebagai
    <w:t> run biasa. para.text/para.runs (python-docx) SAMA SEKALI
    TIDAK MEMBACA <m:t> -- cuma <w:t>/<w:tab>/<w:br> yang terbaca --
    jadi nomor identitas yang "kejebak" di math zone begini HILANG
    TOTAL dari elem["full_text"] dan comparator.py keliru melaporkan
    "tidak memiliki nomor identitas", padahal mahasiswa SUDAH menulis
    nomornya, cuma posisi ketikannya yang belum keluar dari mode
    Equation.

    Dipakai comparator.py HANYA sebagai fallback pencarian pola nomor
    identitas (EQUATION_NUMBER_IN_TEXT_RE) kalau full_text (w:t biasa)
    tidak ketemu -- BUKAN untuk cek font/bold nomornya, karena offset
    karakter <m:t> tidak sejajar/comparable dengan elem["runs"] (yang
    cuma berisi <w:r> biasa, bukan <m:r> math run).
    """
    return "".join(m_t.text for m_t in para._p.iter(qn("m:t")) if m_t.text)


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


def _indent_fallback_cm(para, attr_name: str) -> float | None:
    """
    Baca w:ind/@w:start atau @w:end langsung dari XML mentah -- fallback
    untuk dokumen yang atribut indent-nya ditulis pakai w:start/w:end
    (varian bidi-aware OOXML) alih-alih w:left/w:right klasik.

    KENAPA INI PERLU: python-docx (paragraph_format.left_indent/
    right_indent) HANYA membaca w:ind/@w:left dan @w:right -- tidak
    pernah fallback ke w:start/w:end walau keduanya sah menurut skema
    OOXML (ECMA-376) dan merepresentasikan nilai yang SAMA PERSIS untuk
    dokumen left-to-right seperti tesis Bahasa Indonesia. LibreOffice
    diketahui menulis w:start/w:end (bukan w:left/w:right) saat
    meng-IMPOR file .rtf lalu simpan ulang jadi .docx -- konsekuensinya,
    dokumen yang mahasiswa submit dalam bentuk RTF dan dikonversi lewat
    LibreOffice akan SELALU balik None di sini walau indentasi
    sebenarnya ADA di file, dan CEK KUTIPAN PANJANG (indentasi
    kiri-kanan 10mm) jadi tidak pernah ter-deteksi/gagal secara keliru.
    Diverifikasi lewat pengujian round-trip docx -> rtf -> docx nyata:
    nilai w:start="720"/w:end="720" (720 twips = 1.27cm) HILANG total
    kalau cuma baca left_indent/right_indent bawaan python-docx, dan
    balik BENAR (1.27cm) begitu fallback ini ditambahkan.
    """
    pPr = para._p.pPr
    if pPr is None:
        return None
    ind_el = pPr.find(qn("w:ind"))
    if ind_el is None:
        return None
    raw = ind_el.get(qn(attr_name))
    if raw is None:
        return None
    try:
        return Length(int(raw) * 635).cm  # 1 twip = 635 EMU
    except (TypeError, ValueError):
        return None


def _get_left_indent_cm(para) -> float | None:
    """
    Indentasi kiri paragraf dalam cm, None kalau tidak di-set. Fallback
    ke w:ind/@w:start kalau w:left kosong -- lihat docstring
    _indent_fallback_cm() untuk kenapa ini penting (kasus dokumen hasil
    konversi RTF->DOCX lewat LibreOffice).
    """
    ind = para.paragraph_format.left_indent
    if ind is not None:
        return ind.cm
    return _indent_fallback_cm(para, "w:start")


def _get_right_indent_cm(para) -> float | None:
    """
    Indentasi kanan paragraf dalam cm, None kalau tidak di-set. Simetris
    dengan _get_left_indent_cm() -- dibutuhkan untuk cek kutipan
    langsung PANJANG (quotation_checker.py / compare_quotations() di
    comparator.py), yang mensyaratkan indentasi KIRI *DAN* KANAN masuk
    10mm dari batas tulisan (beda dari persamaan yang cuma indentasi
    kiri). Fallback ke w:ind/@w:end -- lihat _indent_fallback_cm().
    """
    ind = para.paragraph_format.right_indent
    if ind is not None:
        return ind.cm
    return _indent_fallback_cm(para, "w:end")


def extract_paragraph_params(doc) -> list[dict]:
    """
    Ekstrak parameter mentah tiap paragraf: tipe, teks, chapter, dan
    parameter font/spacing dari tiap run di dalamnya.
    """
    elements = []
    section_map = _map_section_indices(doc)

    current_chapter = "Pra-Bab"  # sampul, kata pengantar, daftar isi, dst.
    expect_chapter_name = False  # true tepat setelah baris "BAB N" ditemukan
    # BUG DIPERBAIKI (2026-09-12, laporan user + konfirmasi pedoman PENS):
    # flag TERPISAH dari expect_chapter_name, khusus untuk baris nama
    # SETELAH heading "LAMPIRAN <huruf>" berdiri sendiri (mis. baris
    # "Arsitektur English-Language Prompting" setelah "LAMPIRAN A").
    # Pedoman PENS (dicek langsung oleh user) HANYA menjelaskan isi
    # umum bagian LAMPIRAN -- tidak pernah mensyaratkan nama tiap
    # lampiran ditulis ALL CAPS seperti nama BAB. Kalau baris ini ikut
    # dipaksa jadi ptype "chapter_title" biasa (disatukan dengan
    # expect_chapter_name), dia akan ikut kena cek ALL CAPS di
    # comparator.py yang sebenarnya HANYA berlaku untuk nama BAB --
    # jadi dipisah ke ptype sendiri "lampiran_title" yang tetap kena
    # cek ukuran font/bold setingkat chapter_title, TAPI TIDAK kena
    # cek ALL CAPS (lihat comparator.py).
    expect_lampiran_name = False
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
        # True HANYA untuk equation mode "Display" (m:oMathPara) --
        # dipakai KHUSUS buat keputusan ptype "equation" di bawah, beda
        # dari has_equation (mendeteksi m:oMath di mana pun, termasuk
        # inline di tengah kalimat prosa) yang tetap dipakai apa adanya
        # buat keputusan "paragraf kosong" & pengisian field equation_*.
        # Lihat docstring _has_display_equation() untuk alasan lengkap.
        has_display_equation = _has_display_equation(para)
        # Rumus inline berdiri sendiri (+ tab + nomor) -- lihat docstring
        # _is_standalone_inline_equation(). Diperlakukan sebagai persamaan
        # juga supaya aturan rata kiri/nomor/tebal berlaku di DOCX seperti
        # di PDF.
        standalone_inline_equation = (
            not has_display_equation and _is_standalone_inline_equation(para)
        )
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
                    LAMPIRAN_TITLE_RE.match(upper_stripped)
                    and stripped == upper_stripped
                ):
                    # Sama seperti normalisasi "BAB N" di atas, tapi untuk
                    # LAMPIRAN (lihat komentar LAMPIRAN_TITLE_RE) --
                    # current_chapter dinormalisasi ke "LAMPIRAN <huruf>"
                    # murni supaya caption Gambar/Tabel di dalamnya (mis.
                    # "Tabel A.1") bisa dicocokkan ke chapter yang benar
                    # oleh compare_caption_numbering() di comparator.py.
                    # Dicek SEBELUM cabang ilvl==0 di bawah (harus menang
                    # duluan) supaya "LAMPIRAN B" yang stylenya kebetulan
                    # Heading 1 tidak salah masuk jalur bab_counter atau
                    # jalur demosi subsubchapter_title -- teksnya sendiri
                    # sudah eksplisit bilang LAMPIRAN, tidak perlu ditebak
                    # dari ilvl Word.
                    current_chapter = re.match(
                        r"^LAMPIRAN(?:\s+[A-Z]+)?", upper_stripped
                    ).group(0)
                    seen_bab = True
                    literal_bab_used = True
                    # PAKAI expect_lampiran_name (bukan expect_chapter_name)
                    # -- lihat komentar deklarasinya di atas: baris nama
                    # setelah "LAMPIRAN <huruf>" tidak wajib ALL CAPS.
                    expect_chapter_name = False
                    expect_lampiran_name = bool(
                        LAMPIRAN_TITLE_ONLY_RE.match(upper_stripped)
                    )
                elif (
                    not literal_bab_used
                    and upper_stripped not in FRONT_HEADINGS
                    and _has_outline_level(para)
                    and _get_numbering_depth(para) == 1
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
                    and _get_numbering_depth(para) == 1
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
            elif expect_lampiran_name:
                # Padanan expect_chapter_name di atas, tapi ptype-nya
                # "lampiran_title" (bukan "chapter_title") supaya lolos
                # dari cek ALL CAPS di comparator.py -- lihat komentar
                # deklarasi expect_lampiran_name.
                ptype = "lampiran_title"
                expect_lampiran_name = False
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
            if has_display_equation or standalone_inline_equation:
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
                    "all_caps": _resolve_effective_all_caps(run, para),
                    # True kalau run ini variabel/simbol matematis yang
                    # diketik sebagai teks biasa (bukan objek OMML) --
                    # dipakai comparator.py buat skip cek font & UKURAN
                    # PER-RUN (bukan per-paragraf), lihat
                    # _is_math_run_text() & MATH_FONT_NAMES di atas
                    # untuk alasan lengkapnya.
                    # Perbandingan case-insensitive (2026-09-20) --
                    # konsisten dengan _span_is_math() versi PDF di
                    # pdf_extractor.py yang juga case-insensitive
                    # (perbaikan bug jalur PDF, lihat docstring-nya):
                    # nama font dari docx XML normalnya sudah persis
                    # "Cambria Math", tapi ini jaga-jaga kalau ada
                    # variasi kapitalisasi yang lolos dari style chain.
                    "is_math": (
                        (run_font_name or "").strip().lower()
                        in {n.lower() for n in MATH_FONT_NAMES}
                        or _is_math_run_text(run.text)
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
                # Tampil KAPITAL SEMUA secara visual: teksnya sudah
                # kapital, ATAU setiap run berhuruf dipasang "All caps"
                # (w:caps). Dipakai cek sampul -- lihat
                # _resolve_effective_all_caps().
                "is_visual_upper": stripped == stripped.upper()
                or (
                    any(any(ch.isalpha() for ch in r["text"]) for r in runs_data)
                    and all(
                        r.get("all_caps") or r["text"].upper() == r["text"]
                        for r in runs_data
                        if any(ch.isalpha() for ch in r["text"])
                    )
                ),
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
                # KHUSUS ptype "equation": rata VISUAL rumus itu sendiri
                # (m:oMathParaPr/m:jc), TERPISAH dari "alignment" di atas
                # (w:pPr/w:jc paragraf pembungkus) -- lihat docstring
                # _get_equation_alignment() untuk alasan lengkap kenapa
                # dua-duanya WAJIB dibaca terpisah (equation Word punya
                # justifikasinya sendiri yang independen dari paragraf).
                # None kalau bukan equation, atau equation-nya tidak
                # (belum) punya m:oMathParaPr/m:jc eksplisit sama sekali.
                "equation_alignment": (
                    _get_equation_alignment(para) if has_equation else None
                ),
                # KHUSUS ptype "equation": teks dari <m:t> (math text run)
                # -- fallback pencarian nomor identitas kalau nomornya
                # kejebak diketik masih di dalam math zone, lihat
                # docstring _equation_math_text() untuk detail lengkap.
                "equation_math_text": (
                    _equation_math_text(para).strip() if has_equation else ""
                ),
                # KHUSUS ptype "equation": teks LENGKAP dalam urutan
                # dokumen (w:t + tab + m:t + kurung <m:d> + hasil
                # field) -- fallback nomor identitas yang tidak
                # kelihatan di full_text. Lihat _equation_full_text().
                "equation_full_text": (
                    _equation_full_text(para).strip() if has_equation else ""
                ),
                # "display" (<m:oMathPara>) | "inline_tab" (rumus inline +
                # tab + nomor, tanpa <m:oMathPara>) | None. Dipakai
                # comparator.py untuk menafsirkan rata kiri (lihat catatan
                # di compare_visual_objects).
                "equation_layout": (
                    "display"
                    if has_display_equation
                    else ("inline_tab" if standalone_inline_equation else None)
                ),
                # Tebal/tidaknya nomor yang terjebak di math zone (lihat
                # _equation_number_bold). None = tidak bisa dinilai.
                "equation_number_bold": (
                    _equation_number_bold(para) if has_equation else None
                ),
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


def _is_blank_body_p(el) -> bool:
    """True kalau elemen XML adalah <w:p> kosong sungguhan (tanpa teks,
    gambar, maupun rumus) -- dipakai utk cek 1 baris kosong di sekitar
    persamaan yang dibungkus tabel."""
    if el is None or el.tag != qn("w:p"):
        return False
    if next(el.iter(qn("m:oMath")), None) is not None:
        return False
    if (
        next(el.iter(qn("w:drawing")), None) is not None
        or next(el.iter(qn("w:pict")), None) is not None
    ):
        return False
    return not _prose_text(el).strip()


def _table_row_as_equation(tbl_el, doc):
    """
    Kalau tabel ini adalah TRIK LAYOUT persamaan -- tepat 1 baris, 1-3
    sel, tepat satu sel berisi rumus OMML murni (tanpa teks prosa), sel
    lain kosong atau cuma nomor identitas "(3.1)" -- kembalikan
    (runs, formula_paragraph); selain itu None.
    """
    tbl = Table(tbl_el, doc._body)
    if len(tbl.rows) != 1:
        return None
    cells = []
    for c in tbl.rows[0].cells:
        if not any(c._tc is x._tc for x in cells):  # sel merge: jangan dobel
            cells.append(c)
    if not 1 <= len(cells) <= 3:
        return None

    formula_cells = [
        c
        for c in cells
        if next(c._tc.iter(qn("m:oMath")), None) is not None
    ]
    if len(formula_cells) != 1:
        return None
    formula_para = None
    for p in formula_cells[0].paragraphs:
        if next(p._p.iter(qn("m:oMath")), None) is not None:
            formula_para = p
            break
    if formula_para is None:
        return None

    runs = []
    for c in cells:
        prose = normalize_equation_text(_prose_text(c._tc))
        if c is formula_cells[0]:
            if prose:  # sel rumus bercampur prosa -> bukan trik persamaan
                return None
        elif prose and not EQUATION_NUMBER_RE.match(prose):
            return None  # sel lain berisi teks selain nomor (mis. kotak teks)
        for p in c.paragraphs:
            for run in p.runs:
                if not run.text.strip():
                    continue
                runs.append(
                    {
                        "text": run.text,
                        "font_name": _resolve_effective_font_name(run, p),
                        "font_size_pt": _resolve_effective_font_size(run, p),
                        "bold": _resolve_effective_bold(run, p),
                        "italic": _resolve_effective_italic(run, p),
                    }
                )
    return runs, formula_para


def extract_table_equations(doc, paragraphs: list[dict]) -> list[dict]:
    """
    Persamaan yang diletakkan di dalam TABEL LAYOUT 1 baris (rumus di
    satu sel, nomor "(3.1)" di sel sebelahnya) -- trik yang sangat umum
    karena tab stop gampang geser (lihat catatan n_rows di
    extract_table_params).

    KENAPA PERLU (2026-11-09): extract_paragraph_params() hanya membaca
    doc.paragraphs, yaitu paragraf level-body saja; paragraf DI DALAM
    sel tabel tidak ikut sama sekali. Persamaan ber-tabel jadi tak
    terlihat oleh cek persamaan, dan tabelnya sendiri dikecualikan dari
    cek tabel (n_rows<=1). Jalur PDF tidak kena karena membaca tata
    letak visual.

    Hasilnya elemen berbentuk sama dengan elemen paragraf ber-ptype
    "equation" (dipakai compare_visual_objects(table_equations=...)),
    ditambah "location_label". Sengaja DISIMPAN TERPISAH dari
    elem["paragraphs"] supaya index paragraf (yang jadi acuan
    extract_table_params & cek lain) tidak bergeser.
    """
    body = doc.element.body
    idxs = [p["index"] for p in paragraphs]
    results = []
    para_counter = -1
    table_seq = 0
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            para_counter += 1
            continue
        if child.tag != qn("w:tbl"):
            continue
        table_seq += 1
        parsed = _table_row_as_equation(child, doc)
        if parsed is None:
            continue
        runs, fp = parsed

        # Bab/chapter: elemen paragraf terdekat SEBELUM tabel ini.
        pos = bisect.bisect_right(idxs, para_counter) - 1
        chapter = paragraphs[pos]["chapter"] if pos >= 0 else "Pra-Bab"

        eq_align = _get_equation_alignment(fp)
        if eq_align is None and _has_display_equation(fp):
            # Default OMML (ECMA-376) untuk equation Display tanpa
            # m:jc eksplisit = centerGroup (rata tengah).
            eq_align = "center"

        shim = types.SimpleNamespace(_p=child)  # walker rekursif generik
        results.append(
            {
                "index": para_counter,
                "type": "equation",
                "chapter": chapter,
                "location_label": (
                    f"Persamaan dalam tabel (tabel ke-{table_seq} dalam "
                    f"dokumen)"
                ),
                "full_text": "".join(r["text"] for r in runs),
                "runs": runs,
                "alignment": _get_paragraph_alignment(fp),
                "equation_alignment": eq_align,
                "equation_math_text": _equation_math_text(shim).strip(),
                "equation_full_text": _equation_full_text(shim).strip(),
                "equation_number_bold": _equation_number_bold(shim),
                "equation_layout": "table",
                "blank_before": _is_blank_body_p(child.getprevious()),
                "blank_after": _is_blank_body_p(child.getnext()),
            }
        )
    return results


def _get_document_background_color_hex(doc) -> str | None:
    """
    Warna latar belakang halaman (Word: Design > Page Color), disimpan
    di XML sebagai elemen <w:background w:color="RRGGBB"/> tepat di
    bawah root <w:document> -- SEJAJAR dengan <w:body>, BUKAN di
    dalamnya -- makanya diambil langsung dari doc.element, beda dari
    elemen lain di modul ini yang selalu lewat doc.element.body.

    Dipakai comparator.py untuk cek warna sampul "#002FA7" (RGB 0, 47,
    167) sesuai panduan.

    KETERBATASAN PENTING: cek ini HANYA menangkap warna yang diset
    lewat fitur Page Color bawaan Word (elemen <w:background> di atas).
    Kalau mahasiswa mewarnai sampul dengan cara lain -- misalnya shape/
    rectangle selebar 1 halaman penuh yang di-fill manual, atau
    menempelkan gambar sampul biru sebagai background image -- warna
    itu TIDAK PERNAH tertangkap fungsi ini, walau secara visual sampul
    sudah 100% benar. comparator.py HARUS memperlakukan return None di
    sini sebagai "tidak bisa dipastikan otomatis", BUKAN sebagai
    "warna salah" -- supaya sampul yang sebenarnya sudah benar (dibuat
    lewat shape, bukan Page Color) tidak ikut salah ditandai.
    """
    bg = doc.element.find(qn("w:background"))
    if bg is None:
        return None
    color = bg.get(qn("w:color"))
    if not color or color.upper() == "AUTO":
        return None
    return color.upper()


def extract_all(doc) -> dict:
    """Entry point tahap Extract Parameter: gabungkan semua parameter mentah."""
    paragraphs = extract_paragraph_params(doc)
    return {
        "paragraphs": paragraphs,
        # Persamaan yang dibungkus tabel layout 1 baris (lihat
        # extract_table_equations). Terpisah dari "paragraphs".
        "table_equations": extract_table_equations(doc, paragraphs),
        "sections": extract_section_params(doc),
        "tables": extract_table_params(doc),
        # Level DOKUMEN, bukan per-section -- lihat docstring _get_mirror_margins().
        "mirror_margins": _get_mirror_margins(doc),
        # None kalau tidak terdeteksi -- lihat keterbatasan di docstring
        # _get_document_background_color_hex() di atas. Cuma tersedia di
        # jalur docx; jalur PDF (pdf_extractor.py) belum mengimplementasikan
        # deteksi warna latar halaman sama sekali, comparator.py akan
        # menerima None lewat extracted_params.get(...) untuk jalur itu.
        "cover_background_hex": _get_document_background_color_hex(doc),
    }
