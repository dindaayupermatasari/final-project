"""
quotation_checker.py
=====================
Deteksi & cek format KUTIPAN LANGSUNG (direct quotation) di body text,
sesuai Panduan Buku PA Sarjana Terapan PENS:

  1. Kutipan langsung PENDEK (1-2 baris): diapit tanda petik dua,
     ditulis italic, diikuti sitasi.
  2. Kutipan langsung PANJANG (>2 baris): paragraf terpisah, tanda
     petik dua di awal-akhir, font 10pt, indentasi kiri-kanan masuk
     10mm dari batas tulisan, diikuti sitasi.

CARA KERJA
----------
1. Cari SEMUA span teks yang diapit tanda kutip dua (lurus " atau
   smart “ ”) di tiap paragraf lewat regex -- bukan NLP, karena tanda
   kutip sudah jadi penanda tekstual yang cukup jelas untuk batas
   kutipan.
2. DISAMBIGUASI dari istilah/dialog biasa yang kebetulan dikutip (mis.
   'method disebut "custom"', atau dialog wawancara) -- BUKAN kutipan
   langsung formal yang perlu tunduk ke aturan format akademik. Sebuah
   span HANYA dianggap KANDIDAT kutipan langsung kalau LOLOS DUA
   syarat:
     a. Minimal MIN_QUOTE_WORDS kata di dalam tanda kutip -- kutipan
        akademik wajar berupa kalimat/klausa, bukan cuma 1-2 kata
        istilah yang kebetulan diberi tanda kutip.
     b. Diikuti pola SITASI (mis. "(Nama, 2020)", "(Nama et al.,
        2020)", atau "[19]") tidak jauh setelah tanda kutip penutup --
        kutipan langsung akademik SELALU disertai sitasi tepat
        setelahnya. Kalau tidak ada sitasi sama sekali di dekat span
        itu, kemungkinan besar itu cuma istilah/dialog yang kebetulan
        dikutip, BUKAN kutipan langsung formal -- span ini di-SKIP
        TOTAL (tidak checked, tidak failed), supaya condong presisi
        (menghindari false positive) daripada recall.
   Ini SENGAJA longgar terhadap arah sebaliknya: kutipan yang sitasinya
   ditulis SEBELUM kutipan (mis. "Menurut Nama (2020), '...'") tidak
   akan terdeteksi sebagai kandidat (karena tidak ada sitasi PERSIS
   setelah tanda kutip penutup). Ini trade-off yang sama seperti
   caption tabel yang boleh di atas MAUPUN di bawah tabel di
   comparator.py -- daripada memaksakan pola yang belum tentu benar
   dan salah-tandai dokumen yang sebenarnya sudah benar, kandidat yang
   ambigu arah sitasinya sengaja tidak ikut dihitung.
3. Klasifikasi PENDEK vs PANJANG: python-docx tidak tahu pembungkusan
   baris fisik paragraf (butuh render halaman ke PDF, di luar jangkauan
   modul ini -- keterbatasan yang sama seperti yang sudah disadari di
   compare_visual_objects() untuk hal lain, mis. tabel >1 halaman).
   Sebagai PROXY: kalau span kutipan mencakup HAMPIR SELURUH teks
   paragraf (>= LONG_QUOTE_COVERAGE), paragraf itu memang didedikasikan
   utuh untuk kutipan -- ini representasi wajar dari "paragraf
   terpisah" yang jadi ciri kutipan PANJANG di panduan. Selain itu
   (span kutipan cuma sebagian dari paragraf, menyatu dengan kalimat
   pengantar/lanjutan penulis) dianggap kutipan PENDEK.

ATURAN YANG DICEK DI SINI (level SPAN kutipan)
------------------------------------------------
- Kutipan PENDEK: seluruh KARAKTER span kutipan wajib italic (kalau
  cuma sebagian yang italic, mis. cuma kata pertama, dianggap BELUM
  italic dengan benar -- sama seperti prinsip check_paragraph() di
  foreign_term_checker.py).
- Kutipan PANJANG: font 10pt & indentasi kiri-kanan 10mm dicek di
  LEVEL PARAGRAF (butuh elem["runs"]/elem["left_indent_cm"]/
  elem["right_indent_cm"] penuh, bukan cuma span) -- jadi dilakukan di
  compare_quotations() (comparator.py), BUKAN di modul ini. Modul ini
  cuma menandai span sebagai kandidat "is_long" dan modul comparator
  yang mengisi pengecekan format paragrafnya.
- Sitasi wajib ada: sudah jadi SYARAT KANDIDAT (poin 2 di atas), bukan
  cek terpisah -- span tanpa sitasi memang tidak pernah masuk hitungan
  checked/failed.

SENGAJA TIDAK DICEK DI SINI (lihat compare_quotations() di
comparator.py untuk penjelasan lengkap kenapa):
- Kutipan TIDAK LANGSUNG (parafrase) -- aturan panduan cuma bilang
  "sitasi ditulis setelah kalimat selesai", tapi tidak ada penanda
  tekstual pasti untuk membedakan kalimat parafrase yang butuh sitasi
  dari kalimat body biasa yang memang tidak perlu sitasi. Beda dari
  kutipan langsung yang punya penanda jelas (tanda kutip), memaksakan
  cek ini akan sangat overclaim false positive.
- Jumlah baris fisik PERSIS (1-2 vs >2 baris) -- pakai proxy cakupan
  span terhadap paragraf (lihat poin 3 di atas), bukan hitungan baris
  hasil render PDF.
"""

import re

# Tanda kutip pembuka/penutup, lurus (") maupun smart ("curly": “ ”).
# Non-greedy (.+?) supaya tidak "melompati" ke tanda kutip penutup
# span BERIKUTNYA kalau ada lebih dari satu kutipan di paragraf yang
# sama. DOTALL supaya kutipan panjang yang (jarang tapi mungkin)
# mengandung newline literal di w:t tetap tertangkap utuh.
QUOTE_SPAN_RE = re.compile(r'(["\u201c])(.+?)(["\u201d])', re.DOTALL)

# Pola sitasi yang dicari TEPAT SETELAH tanda kutip penutup (boleh
# diselingi tanda baca/spasi pendek, mis. "..." (Nama, 2020) atau
# "..." [19]). Menerima dua gaya umum:
#   - "(Nama, 2020)" / "(Nama et al., 2020)" / "(Nama, 2020, hlm. 12)"
#     -- parenthetical berisi tahun 4 digit.
#   - "[19]" / "[19], [20]" -- gaya sitasi numerik IEEE.
CITATION_RE = re.compile(
    r"^\s*[.,]?\s*(\([^()]{0,100}?\d{4}[a-z]?[^()]{0,30}\)"
    r"|\[\d+(?:\s*[-,]\s*\d+)*\])"
)

# Kutipan formal wajar berupa kalimat/klausa, bukan 1-2 kata istilah
# yang kebetulan diberi tanda kutip -- lihat poin 2a di docstring modul.
MIN_QUOTE_WORDS = 4

# Span dianggap "mendominasi" paragraf (proxy kutipan PANJANG/paragraf
# terpisah) kalau mencakup proporsi ini dari total panjang teks
# paragraf.
LONG_QUOTE_COVERAGE = 0.85

# Seberapa jauh (jumlah karakter) sesudah tanda kutip penutup, pola
# sitasi dicari -- cukup longgar untuk menampung "..." (Nama, 2020).
# tapi tidak sampai menangkap sitasi kalimat lain yang tidak terkait.
CITATION_LOOKAHEAD_CHARS = 40


def _word_count(s: str) -> int:
    return len(s.split())


def find_quote_candidates(full_text: str) -> list[dict]:
    """
    Ambil semua span kutipan langsung yang lolos jadi KANDIDAT (lihat
    syarat 2a/2b di docstring modul) dari SATU string teks paragraf
    utuh. Tiap kandidat: {"start", "end" (span TERMASUK tanda kutip),
    "quoted_start", "quoted_end" (isi DI DALAM tanda kutip saja),
    "quoted_text", "is_long"}.
    """
    candidates = []
    for m in QUOTE_SPAN_RE.finditer(full_text):
        quoted_text = m.group(2)
        if _word_count(quoted_text) < MIN_QUOTE_WORDS:
            continue

        after = full_text[m.end() : m.end() + CITATION_LOOKAHEAD_CHARS]
        if not CITATION_RE.match(after):
            continue  # tidak ada sitasi mengiringi -- bukan kandidat

        span_len = m.end() - m.start()
        coverage = span_len / max(len(full_text), 1)

        candidates.append(
            {
                "start": m.start(),
                "end": m.end(),
                "quoted_start": m.start(2),
                "quoted_end": m.end(2),
                "quoted_text": quoted_text,
                "is_long": coverage >= LONG_QUOTE_COVERAGE,
            }
        )
    return candidates


def check_paragraph(runs: list[dict]) -> list[dict]:
    """
    Cek kandidat kutipan langsung dalam SATU paragraf UTUH (list run
    mentah dari extractor.py, tiap run: {"text", "font_name",
    "font_size_pt", "bold", "italic"}).

    Paragraf direkonstruksi jadi SATU string penuh dulu (bukan run per
    run) -- alasan SAMA PERSIS dengan check_paragraph() di
    foreign_term_checker.py: kutipan yang kebetulan terpotong di batas
    run (mis. tanda kutip pembuka di run yang beda format dari isinya)
    tetap terbaca sebagai satu kutipan utuh, dan status italic-nya
    ditentukan per-karakter dari run mana tiap karakter itu berasal.

    Return: list of dict, satu per kandidat kutipan yang ditemukan:
        {"is_long": bool, "quoted_text": str, "problems": [str, ...]}
    "problems" HANYA diisi untuk kutipan PENDEK (cek italic span, level
    kutipan). Kutipan PANJANG selalu balik "problems": [] dari modul
    ini -- cek font/indent-nya level PARAGRAF, dilakukan di
    compare_quotations() (comparator.py), bukan di sini.
    """
    text_parts = []
    italic_map = []  # paralel dengan full_text gabungan, per karakter
    for run in runs:
        t = run["text"]
        text_parts.append(t)
        italic_map.extend([run["italic"] is True] * len(t))
    full_text = "".join(text_parts)

    results = []
    for cand in find_quote_candidates(full_text):
        problems = []
        if not cand["is_long"]:
            span_italic = italic_map[cand["quoted_start"] : cand["quoted_end"]]
            fully_italic = bool(span_italic) and all(span_italic)
            if not fully_italic:
                problems.append("belum dicetak italic seluruhnya")

        results.append(
            {
                "is_long": cand["is_long"],
                "quoted_text": cand["quoted_text"],
                "problems": problems,
            }
        )

    return results
