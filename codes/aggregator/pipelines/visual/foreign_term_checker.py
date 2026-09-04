"""
foreign_term_checker.py
========================
Deteksi "istilah asing" yang wajib dicetak italic, berbasis DUA kamus
(Indonesia & Inggris), bukan cuma satu -- supaya presisinya cukup
tinggi untuk ikut dihitung normal ke skor (checked/failed), sama
seperti pengecekan font/bold/size lain di comparator.py, bukan
dijadikan kategori "catatan review manual" terpisah.

CARA KERJA
----------
Pakai library `wordfreq` (data frekuensi kata di-bundle offline di
package-nya sendiri, TIDAK perlu fetch API/internet tiap kali dokumen
dicek) sebagai proxy dua kamus:
    - zipf_frequency(word, "id")  ~ proxy "ada di KBBI / lazim dipakai
      dalam Bahasa Indonesia" (termasuk kata serapan yang sudah lazim
      dipakai TANPA italic, mis. "internet", "aplikasi", "digital").
    - zipf_frequency(word, "en")  ~ proxy "kata Inggris asli".

Sebuah kata dianggap KANDIDAT istilah asing kalau SELISIH-nya
signifikan: jauh lebih "Inggris" daripada "Indonesia". Threshold
default (margin=1.5, min_en_zipf=1.5) dikalibrasi dari sampel kata
umum di skripsi bidang data science/NLP.

CHANGELOG (revisi setelah uji coba dokumen asli):
- STOPWORDS ditambah: kata sambung/fungsi Inggris ("and", "for",
  "the", "with", "from", dst) SEBELUMNYA ikut lolos threshold zipf
  (karena memang kata Inggris asli yang sangat umum, jadi id_zipf
  rendah & en_zipf tinggi -- persis pola "istilah asing"). Tapi kata
  fungsi bukan "istilah" yang berdiri sendiri secara leksikal, jadi
  sekarang di-exclude eksplisit di luar mekanisme zipf.
- Deteksi proper noun DIPERLUAS: sebelumnya cuma cek huruf pertama
  kapital ("Qwen" ke-exclude, tapi "arXiv" LOLOS karena huruf pertama
  'a' kecil). Sekarang juga cek huruf kapital DI TENGAH kata
  (camelCase-like: "arXiv", "GitHub", "iPhone") sebagai sinyal nama
  diri/merek.
- BUG DIPERBAIKI: sebelumnya find_foreign_words() dipanggil PER RUN
  secara terpisah dari comparator.py. Kalau satu kata kebetulan
  terpotong jadi 2 run (mis. hasil copy-paste yang bawa format beda
  di tengah kata -- "C" di run A, "hecking" di run B), kata itu
  KELIRU dibaca sebagai dua token terpisah ("hecking" jadi kandidat
  sendiri). Sekarang paragraf direkonstruksi jadi SATU string penuh
  dulu (check_paragraph()), tokenisasi kata dilakukan di string utuh
  itu, baru status italic-nya dicek per-karakter dari run mana kata
  itu berasal. Kata dianggap "sudah italic" hanya kalau SELURUH
  karakternya italic (bukan cuma sebagian run yang menyusunnya).
- find_foreign_words(text) versi lama (input string lepas, bukan list
  run) DIPERTAHANKAN untuk kompatibilitas/testing cepat, tapi
  comparator.py SEKARANG pakai check_paragraph() yang baru.
- WHITELIST DIPINDAH KE FILE EKSTERNAL (foreign_term_glossary.json,
  sejajar dengan file ini). Sebelumnya KNOWN_PROPER_NOUNS dan
  KNOWN_ABSORBED_WORDS di-hardcode di kode -- tiap nemu kata baru yang
  salah ke-flag harus edit .py + redeploy. Sekarang dua daftar itu
  dimuat dari JSON saat modul di-import (_load_glossary()), jadi bisa
  direvisi oleh pembimbing/prodi tanpa sentuh kode. Kalau file JSON
  tidak ada / gagal di-parse, fallback ke daftar default di
  _DEFAULT_PROPER_NOUNS / _DEFAULT_ABSORBED_WORDS di bawah supaya
  checker tetap jalan (fail-safe, bukan fail-crash). STOPWORDS
  (kata sambung Inggris) SENGAJA TETAP di kode -- itu closed-class
  linguistik yang jarang berubah, beda alasan dengan glossary istilah
  domain-specific yang open-class dan terus bertambah.
- DITAMBAH forced_foreign DI GLOSSARY: proper_nouns & absorbed_words
  menambal arah FALSE POSITIVE (kata salah ke-flag padahal bukan
  istilah asing). Tapi ada arah SEBALIKNYA yang belum tertambal:
  FALSE NEGATIVE -- istilah yang SEHARUSNYA di-flag wajib italic,
  tapi LOLOS tidak terdeteksi karena skor zipf-nya kebetulan tidak
  cukup tinggi (umum terjadi untuk istilah niche/domain-specific yang
  jarang muncul di korpus umum wordfreq, mis. istilah RAG/NLP yang
  cukup spesifik). forced_foreign menutup celah ini: kata di daftar
  ini SELALU dianggap kandidat istilah asing, melewati SEMUA
  exclusion lain (_is_excluded, threshold zipf) -- lihat
  check_paragraph() & find_foreign_words().
- FALLBACK KE DEFAULT diperbaiki supaya HANYA terjadi kalau file
  benar-benar gagal dibaca/di-parse (exception), BUKAN berdasar isi
  kosong-atau-tidak. Sebelumnya kalau proper_nouns & absorbed_words
  sama-sama kosong (misal SENGAJA dikosongkan pembimbing/prodi untuk
  testing), kode salah mengira itu "file rusak" dan maksa balik ke
  default -- padahal itu keadaan valid yang harus dihormati.
"""

import json
import re
from pathlib import Path

from wordfreq import zipf_frequency

WORD_RE = re.compile(r"[A-Za-z][A-Za-z'\-]*")

# Selisih minimum (en_zipf - id_zipf) supaya kata dianggap "cukup asing".
DEFAULT_MARGIN = 1.5
# Kata harus minimal "cukup dikenal" di korpus Inggris -- kata acak /
# typo / nama produk yang kebetulan ke-parse sebagai huruf latin tidak
# boleh lolos hanya karena id_zipf-nya kebetulan 0.
DEFAULT_MIN_EN_ZIPF = 1.5

# Kata fungsi/sambung Inggris -- SELALU dikecualikan terlepas dari
# skor zipf, karena ini bukan "istilah" (tidak punya makna leksikal
# berdiri sendiri), melainkan struktur kalimat.
STOPWORDS = {
    "a",
    "an",
    "the",
    "and",
    "or",
    "but",
    "nor",
    "so",
    "of",
    "in",
    "on",
    "at",
    "by",
    "for",
    "with",
    "from",
    "to",
    "as",
    "into",
    "onto",
    "over",
    "under",
    "via",
    "per",
    "is",
    "are",
    "was",
    "were",
    "be",
    "been",
    "being",
    "this",
    "that",
    "these",
    "those",
    "it",
    "its",
    "vs",
    "etc",
}

# --- Glossary eksternal (proper nouns & kata serapan) ---------------------
#
# Dua daftar ini SEBELUMNYA di-hardcode langsung sebagai set literal di
# sini. Sekarang dipindah ke foreign_term_glossary.json (sejajar dengan
# file ini) supaya bisa direvisi pembimbing/prodi tanpa edit kode.
# Alasan kenapa dua kategori ini dipisah (bukan digabung jadi satu
# daftar) tetap sama seperti sebelumnya:
#   - proper_nouns     : nama tool/library/produk -- BUKAN istilah,
#                         bukan kata umum sama sekali (mis. "arxiv",
#                         "python-docx").
#   - absorbed_words    : kata serapan/istilah baku yang lazim dipakai
#                         TANPA italic dalam tulisan ilmiah Indonesia,
#                         tapi korpus "id" wordfreq tipis untuk
#                         kosakata teknis ini (id_zipf sering 0.0)
#                         sehingga selisih (en_zipf - id_zipf) salah
#                         lolos threshold margin. Menaikkan threshold
#                         global TIDAK aman -- lihat contoh lama:
#                         diff "evaluator" (2.49) < diff "manipulator"
#                         (2.85) yang sudah harus lolos whitelist, jadi
#                         threshold berapa pun akan salah motong salah
#                         satu sisi. Whitelist eksplisit per-kata lebih
#                         presisi.
#
# _DEFAULT_* di bawah HANYA dipakai sebagai fallback kalau file JSON
# tidak ditemukan / gagal di-parse (fail-safe, checker tetap jalan
# dengan cakupan whitelist minimal, bukan crash).

_GLOSSARY_PATH = Path(__file__).parent / "foreign_term_glossary.json"

_DEFAULT_PROPER_NOUNS = {
    "python-docx",
    "pymupdf",
    "libreoffice",
    "arxiv",
}

_DEFAULT_ABSORBED_WORDS = {
    "kustom",
    "render",
    "rendering",
    "manipulator",
    "dataset",
    "pipeline",
    "metadata",
    "evaluator",
    "bilingual",
    "frontend",
    "backend",
    "multilingual",
}

# forced_foreign KOSONG secara default -- beda dari 2 daftar di atas,
# ini bukan "kondisi awal yang wajar untuk skripsi data science pada
# umumnya" (proper_nouns/absorbed_words sudah punya isi bawaan yang
# relevan buat semua dokumen). forced_foreign sifatnya SANGAT
# spesifik per dokumen (istilah niche yang kebetulan lolos threshold
# zipf), jadi wajar mulai kosong dan diisi manual sesuai temuan nyata.
_DEFAULT_FORCED_FOREIGN = set()


def _load_glossary(
    path: Path = _GLOSSARY_PATH,
) -> tuple[set[str], set[str], set[str]]:
    """
    Muat proper_nouns, absorbed_words, & forced_foreign dari JSON
    eksternal. Fallback ke default di atas HANYA kalau file gagal
    dibaca sama sekali atau JSON-nya tidak valid (exception) --
    BUKAN berdasar isi kosong-atau-tidak, supaya glossary yang
    SENGAJA dikosongkan (mis. buat testing) tetap dihormati, bukan
    dipaksa balik ke default.
    """
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        proper_nouns = {str(w).lower() for w in raw.get("proper_nouns", [])}
        absorbed_words = {str(w).lower() for w in raw.get("absorbed_words", {})}
        forced_foreign = {str(w).lower() for w in raw.get("forced_foreign", {})}
        return proper_nouns, absorbed_words, forced_foreign
    except (OSError, json.JSONDecodeError, AttributeError):
        return (
            set(_DEFAULT_PROPER_NOUNS),
            set(_DEFAULT_ABSORBED_WORDS),
            set(_DEFAULT_FORCED_FOREIGN),
        )


KNOWN_PROPER_NOUNS, KNOWN_ABSORBED_WORDS, FORCED_FOREIGN_WORDS = _load_glossary()


def _has_internal_capital(word: str) -> bool:
    """
    True kalau ada huruf kapital SELAIN di posisi pertama (mis.
    "arXiv", "GitHub", "iPhone", "ChromaDB"). Sinyal kuat nama
    diri/merek/produk, bukan istilah umum.
    """
    return any(c.isupper() for c in word[1:])


def _is_excluded(word: str) -> bool:
    """
    Kata yang SENGAJA tidak pernah dianggap kandidat istilah asing:
      - Panjang < 3 huruf.
      - Stopword/kata fungsi Inggris (lihat STOPWORDS).
      - ALL CAPS (akronim: RAG, API, PENS, BAB).
      - Diawali huruf kapital (mis. "Qwen") ATAU ada huruf kapital di
        tengah kata (mis. "arXiv", "GitHub") -- kemungkinan besar nama
        diri/merek/nama produk, bukan istilah asing biasa yang dipakai
        sebagai kata umum. Heuristik ini sengaja longgar (condong
        precision, bukan recall) -- risikonya istilah asing yang
        kebetulan di awal kalimat bisa lolos tidak ke-flag.
    """
    if len(word) < 3:
        return True
    if word.lower() in STOPWORDS:
        return True
    if word.lower() in KNOWN_PROPER_NOUNS:
        return True
    if word.lower() in KNOWN_ABSORBED_WORDS:
        return True
    if word.isupper():
        return True
    if word[0].isupper():
        return True
    if _has_internal_capital(word):
        return True
    return False


def is_foreign_candidate(
    word: str,
    margin: float = DEFAULT_MARGIN,
    min_en_zipf: float = DEFAULT_MIN_EN_ZIPF,
) -> bool:
    """Cek satu kata (sudah diasumsikan lolos _is_excluded)."""
    lw = word.lower()
    en_zipf = zipf_frequency(lw, "en")
    if en_zipf < min_en_zipf:
        return False
    id_zipf = zipf_frequency(lw, "id")
    return (en_zipf - id_zipf) >= margin


def find_foreign_words(
    text: str,
    margin: float = DEFAULT_MARGIN,
    min_en_zipf: float = DEFAULT_MIN_EN_ZIPF,
) -> list[str]:
    """
    Ambil semua kandidat istilah asing dalam SATU string teks lepas
    (dipertahankan untuk testing/kalibrasi cepat -- lihat __main__).
    Untuk pengecekan sungguhan di comparator.py, pakai check_paragraph()
    di bawah, yang aman terhadap kata yang terpotong di batas run.
    """
    seen = set()
    result = []
    for w in WORD_RE.findall(text):
        lw = w.lower()
        forced = lw in FORCED_FOREIGN_WORDS
        if not forced and _is_excluded(w):
            continue
        if lw in seen:
            continue
        if forced or is_foreign_candidate(w, margin, min_en_zipf):
            seen.add(lw)
            result.append(w)
    return result


def check_paragraph(
    runs: list[dict],
    margin: float = DEFAULT_MARGIN,
    min_en_zipf: float = DEFAULT_MIN_EN_ZIPF,
) -> tuple[bool, list[str]]:
    """
    Cek satu paragraf UTUH (list run dari extractor.py, tiap run:
    {"text", "font_name", "font_size_pt", "bold", "italic"}).

    Paragraf direkonstruksi jadi satu string penuh dulu (bukan run per
    run) supaya kata yang kebetulan terpotong di batas run (akibat
    copy-paste yang bawa format beda di tengah kata) tetap terbaca
    sebagai SATU kata utuh. Status italic tiap kata ditentukan dari
    SELURUH karakternya -- kalau ada satu karakter saja yang bukan
    italic (mis. cuma huruf pertama yang italic, sisanya tidak), kata
    itu dianggap BELUM italic dengan benar.

    Return: (ada_kandidat, list_kata_yang_belum_italic)
    """
    text_parts = []
    italic_map = []  # paralel dengan text_parts gabungan, per karakter
    for run in runs:
        t = run["text"]
        text_parts.append(t)
        italic_map.extend([run["italic"] is True] * len(t))
    full_text = "".join(text_parts)

    seen = set()
    offending = []
    any_candidate = False

    for m in WORD_RE.finditer(full_text):
        word = m.group(0)
        lw = word.lower()
        forced = lw in FORCED_FOREIGN_WORDS
        if not forced:
            if _is_excluded(word):
                continue
            if lw in seen:
                continue
            if not is_foreign_candidate(word, margin, min_en_zipf):
                continue
        elif lw in seen:
            continue
        seen.add(lw)
        any_candidate = True

        span_italic = italic_map[m.start() : m.end()]
        fully_italic = bool(span_italic) and all(span_italic)
        if not fully_italic:
            offending.append(word)

    return any_candidate, offending
