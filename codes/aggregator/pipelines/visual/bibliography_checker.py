"""
bibliography_checker.py
========================
SUB-TAHAP: DAFTAR PUSTAKA (format IEEE)
-----------------------------------------
Aturan panduan untuk Daftar Pustaka bersifat POLA/URUTAN TEKS yang
deterministik (bukan pemahaman makna), jadi modul ini masuk jalur
RULE-BASED -- setingkat dengan comparator.py & page_layout_checker.py,
BUKAN pipeline sistematika (RAG). Entri Daftar Pustaka tidak butuh LLM
untuk "mengerti" apa isinya; cukup dicocokkan ke pola field per jenis
sumber (jurnal/seminar/buku/TA/media publik/online).

Sumber entri: paragraphs dari extractor.py dengan chapter=="DAFTAR
PUSTAKA" (heading ini sudah ada di FRONT_HEADINGS extractor.py, jadi
current_chapter otomatis ke-set benar oleh extract_paragraph_params())
dan type=="body" (entri individual, bukan heading "DAFTAR PUSTAKA" itu
sendiri).

Aturan yang dicek (dari Panduan Buku PA, bagian DAFTAR PUSTAKA):
    1. Tiap entri diakhiri Tahun (field terakhir).
    2. Tiap entri diawali Nama pengarang (bukan kosong/angka).
    3. Field wajib per JENIS sumber (heuristik dari kata kunci):
       - Jurnal   : ada "Vol."/"No." -> "Hal." OPSIONAL, tidak dicek
         sama sekali (sesuai ketentuan Panduan Buku PA, khusus
         makalah/jurnal halaman tidak wajib dicantumkan).
       - Tugas Akhir (TA/Skripsi/Tesis/Disertasi): wajib ada nama
         perguruan tinggi.
       - Online   : wajib ada alamat internet (http/www).
    4. Kalau semua entri diberi nomor urut manual ("1.", "2.", ...),
       nomor itu wajib berurutan tanpa lompat/duplikat -- KALAU
       penomoran dibuat via Word auto-numbering (tidak muncul di
       para.text sama sekali), cek ini otomatis di-skip (bukan
       dianggap salah), karena tidak ada apa pun untuk dibandingkan.

CATATAN soal urutan entri (kenapa cek alfabetis TIDAK ADA lagi di
modul ini):
    Format bernomor "[1] [2] [3] ..." di header PA memang terlihat
    seperti gaya IEEE, TAPI IEEE asli mengurutkan Daftar Pustaka
    berdasarkan URUTAN KEMUNCULAN SITASI PERTAMA di badan teks --
    BUKAN alfabetis nama pengarang (alfabetis itu ciri khas
    APA/Harvard). Modul ini tidak (dan realistis tidak bisa, tanpa
    parsing sitasi in-text di seluruh Bab 1-4) mengetahui urutan
    sitasi yang benar, jadi memvalidasi "harus alfabetis" berisiko
    tinggi salah menghukum entri yang sebenarnya sudah benar (mis.
    [1] disitasi pertama di Bab 1 padahal namanya diawali huruf Z).
    Cek ini sudah DIHAPUS total dari modul (bukan sekadar dimatikan).

CATATAN soal entri yang terpecah beberapa paragraf (WRAPPED LINES):
    Kalau dokumen sumber berasal dari konversi PDF (atau penulis menekan
    Enter di tiap akhir baris), SATU entri Daftar Pustaka yang makan 2-5
    baris visual masuk ke extractor.py sebagai 2-5 "paragraf" terpisah.
    Tanpa penanganan khusus, setiap potongan baris dicek sebagai entri
    sendiri dan menghasilkan temuan palsu massal ("tidak ada tahun",
    "tidak bernomor", "nomor tidak berurutan" yang makin bergeser).
    Karena itu extract_bibliography_entries() sekarang MENGGABUNGKAN
    paragraf lanjutan ke entri sebelumnya (lihat merge_wrapped_entries())
    sebelum cek apa pun dijalankan.

CATATAN heuristik (bukan kepastian mutlak, beda dari cek font/margin
yang deterministik penuh):
    - Klasifikasi jenis sumber pakai kata kunci, jadi entri yang
      penulisannya tidak baku bisa salah diklasifikasi. Tapi tetap
      lebih baik daripada tidak dicek sama sekali -- kasus tepi
      dilaporkan spesifik di description, mahasiswa/dosen tinggal
      verifikasi manual.

Kategori semua finding di sini: "Format".
"""

import re
from collections import defaultdict

CHAPTER_LABEL = "DAFTAR PUSTAKA"
ORDER_GROUP = "Umum (Urutan Daftar Pustaka)"

ENTRY_NUMBER_RE = re.compile(r"^(?:\[(\d+)\]|(\d+)\.)\s*(.*)$", re.DOTALL)
# Entri IEEE modern lazim DITUTUP metadata akses/identifikasi SETELAH
# tahun, dan beberapa metadata itu bisa menumpuk dalam satu entri, mis.:
#   "..., 2017, doi: 10.xxx/yyy. [Online]. Available: https://....",
#   "..., Jun. 22, 2022. [Online]. Available: URL. [Accessed: May 31, 2025].",
#   "..., Jakarta, Indonesia, 2023, p. 38."
# Semua itu BUKAN identitas publikasi (nama jurnal/penerbit) yang
# seharusnya mendahului tahun, jadi dikupas dulu satu per satu dari
# belakang (strip_trailing_metadata) baru dicek "apakah berakhir tahun".
# Versi lama cuma mengenali SATU metadata di belakang tahun, sehingga
# entri dengan doi + [Online] + [Accessed] salah dianggap error.
_TAIL_PATTERNS = [
    re.compile(r"[,\s]*\[accessed\s*:[^\]]*\]\.?\s*$", re.I),
    re.compile(r"[,\s]*\[online\]\.?\s*available\s*:.*$", re.I | re.S),
    re.compile(r"[,\s]*available\s*:.*$", re.I | re.S),
    re.compile(r"[,\s]*accessed\s*:.*$", re.I | re.S),
    # doi = field penutup tanpa koma di dalamnya (doi bisa terpotong
    # antarbaris, jadi boleh ada spasi di tengah).
    re.compile(r"[,\s]*doi\s*:\s*[^,]*$", re.I),
    # halaman/nomor artikel SETELAH tahun: ", p. 38." / ", pp. 1-14."
    re.compile(r",\s*(?:pp?\.|art\.\s*no\.)\s*[^,]*$", re.I),
]
YEAR_AT_END_RE = re.compile(r"(?:19|20)\d{2}\.?\s*$")
ACCESSED_YEAR_RE = re.compile(r"\[accessed\s*:[^\]]*(?:19|20)\d{2}[^\]]*\]", re.I)
YEAR_ANYWHERE_RE = re.compile(r"(19|20)\d{2}")
# Penanda AWAL entri baru: nomor manual 1-3 digit ("1. ", "12.", "[3]").
# Sengaja maksimal 3 digit supaya potongan baris yang kebetulan diawali
# tahun ("2025.") atau rentang halaman ("1624-1633,") tidak dikira entri
# baru; "(?!\d)" menolak "1.5 ..." (angka desimal, bukan nomor entri).
ENTRY_START_RE = re.compile(r"^(?:\[\d{1,3}\]\s*|\d{1,3}\.(?!\d)\s*)\S")
# Tahun di akhir yang sah sebagai penutup entri: didahului koma (pemisah
# field) atau singkatan bulan ("Apr. 2022."). Tahun yang kebetulan jatuh
# di ujung baris tapi masih bagian judul ("...Proceedings of the 2014")
# TIDAK sah.
_YEAR_FIELD_END_RE = re.compile(
    r"(?:,|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?)"
    r"\s*(?:19|20)\d{2}\.?\s*$",
    re.I,
)
# Baris yang diawali penanda metadata IEEE PASTI lanjutan entri di atasnya
# (mis. "[Accessed: Nov. 13, 2025]." yang terpotong sendirian di baris
# baru setelah URL), apa pun status "lengkap"-nya entri sebelumnya.
CONTINUATION_START_RE = re.compile(
    r"^(?:\[accessed\b|accessed\s*:|doi\s*:|\[online\]|available\s*:)", re.I
)
# Saklar cek nomor urut Daftar Pustaka ("Umum (Urutan Daftar Pustaka)").
# OFF karena nilainya kecil (IEEE urut berdasarkan sitasi, lompat nomor
# nyaris tak pernah terjadi) sementara risiko temuan palsu tinggi saat
# entri terpecah/ber-auto-numbering. Set True untuk mengaktifkan lagi.
CHECK_ENTRY_NUMBERING = False
# Paragraf yang BUKAN entri/lanjutan entri: nomor halaman nyasar dan
# penanda halaman kosong (muncul tepat setelah entri terakhir).
NON_ENTRY_RE = re.compile(
    r"^(?:\d{1,4}|\W*halaman ini sengaja dikosongkan\W*)$", re.IGNORECASE
)
# Tipe paragraf yang boleh jadi bagian entri. Potongan baris bisa salah
# diklasifikasi extractor.py (mis. baris diawali "Tabel ..." jadi
# "caption"), padahal secara isi tetap lanjutan entri.
ENTRY_PARA_TYPES = ("body", "caption", "subchapter_title", "subsubchapter_title")
URL_RE = re.compile(r"(https?://|www\.)\S+", re.IGNORECASE)
INSTITUTION_RE = re.compile(r"POLITEKNIK|UNIVERSITAS|INSTITUT|SEKOLAH TINGGI|AKADEMI")
THESIS_KEYWORDS = ("PROYEK AKHIR", "SKRIPSI", "TESIS", "DISERTASI")
BULAN_RE = re.compile(
    r"\b(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|"
    r"SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b"
)


def _new_counter():
    return {"checked": 0, "failed": 0}


def _text_preview(text: str, length: int = 40) -> str:
    text = text.strip()
    return text[:length] + ("..." if len(text) > length else "")


def strip_leading_number(text: str) -> tuple[str, int | None]:
    """
    Kalau entri diawali nomor manual ("1. Nama, ..."), pisahkan nomor
    dari isinya. Return (isi_tanpa_nomor, nomor_atau_None). Nomor None
    berarti entri ini TIDAK ada penomoran manual di teksnya -- bisa
    jadi memang tidak dinomori, bisa jadi dinomori via Word
    auto-numbering (tidak pernah masuk ke para.text, sama seperti
    kasus subbab di extractor.py).
    """
    m = ENTRY_NUMBER_RE.match(text.strip())
    if m:
        num = int(m.group(1) if m.group(1) is not None else m.group(2))
        return m.group(3).strip(), num
    return text.strip(), None


def extract_first_author(body: str) -> str:
    """Ambil nama pengarang PERTAMA (segmen sebelum koma pertama)."""
    return body.split(",", 1)[0].strip()


def classify_entry_type(body: str) -> str:
    """
    Heuristik jenis sumber dari kata kunci, urutan dari yang paling
    spesifik ke paling generik (biar tidak salah ke-klasifikasi ke
    jenis yang lebih longgar duluan). Fallback "buku" karena field-nya
    (Nama, Judul, Penerbit, Tahun) paling generik / superset dari
    kolom minimal yang dibutuhkan semua jenis lain.
    """
    upper = body.upper()

    # Gaya Indonesia ("Diakses tanggal ...") maupun IEEE ("[Online].
    # Available: ... [Accessed: ...]") -- keduanya entri online.
    if URL_RE.search(body) and (
        "DIAKSES" in upper or "[ONLINE]" in upper or "AVAILABLE:" in upper
    ):
        return "online"

    # Kata kunci TA cuma dianggap penanda JENIS entri kalau berdiri
    # sendiri sebagai satu segmen comma-separated (field label
    # eksplisit, mis. "..., Proyek Akhir, Politeknik..., 2024."),
    # BUKAN kalau cuma nempel di dalam judul artikel yang topiknya
    # kebetulan membahas skripsi/TA (mis. artikel jurnal berjudul
    # "Kesalahan Penulisan Kalimat dalam Skripsi Mahasiswa").
    segments = [s.strip(" \t\"'\u201c\u201d.,") for s in body.split(",")]
    if any(seg.upper() in THESIS_KEYWORDS for seg in segments):
        return "tugas_akhir"
    if re.search(r"\bVOL\.?\s*\d", upper):
        return "jurnal"
    if re.search(r"\bHAL\.?\s*\d", upper):
        return "seminar"
    if BULAN_RE.search(upper):
        return "media_publik"
    return "buku"


def strip_trailing_metadata(body: str) -> str:
    """
    Kupas metadata penutup (doi, [Online]. Available: URL, [Accessed:
    ...], p./pp./art. no.) dari BELAKANG sampai tidak ada lagi yang bisa
    dikupas, lalu buang koma/spasi/titik sisa. Sisa teks = bagian inti
    entri yang harus berakhir tahun.
    """
    text = body.strip()
    changed = True
    while changed:
        changed = False
        for pat in _TAIL_PATTERNS:
            new = pat.sub("", text).rstrip(" \t,")
            if new != text:
                text, changed = new, True
    return text


def entry_ends_with_year(body: str) -> bool:
    """
    True kalau inti entri (setelah metadata penutup dikupas) berakhir
    tahun. Entri web tanpa tahun terbit (lazim di IEEE: "..., [Online].
    Available: URL. [Accessed: Nov. 13, 2025].") dianggap sah kalau ada
    tahun akses di [Accessed: ...] -- tahun terakhir yang tertulis.
    """
    if YEAR_AT_END_RE.search(strip_trailing_metadata(body)):
        return True
    return bool(ACCESSED_YEAR_RE.search(body))


def entry_is_complete(text: str) -> bool:
    """
    Apakah teks gabungan sejauh ini SUDAH satu entri utuh (tidak mungkin
    masih berlanjut di baris berikutnya)? Dipakai merge_wrapped_entries()
    untuk entri tanpa nomor. Sengaja KONSERVATIF: ragu -> dianggap belum
    lengkap (lanjutan ditempel). Tanda belum lengkap:
      - berakhir pemisah/potongan kata ("-", "/", ":", ",") -> URL/doi/kata
        terpotong di ujung baris;
      - berakhir label tanpa isi ("Available:", "doi:", "[Online].");
      - kurung siku belum tertutup ("[Accessed: Nov." lalu "13, 2025].");
      - entri online yang belum memuat URL.
    """
    t = text.strip()
    low = t.lower()
    if not t or t.endswith(("-", "/", ":", ",")):
        return False
    if re.search(r"(?:available|doi)\s*:\s*$", low) or re.search(
        r"\[online\]\.?\s*$", low
    ):
        return False
    if t.count("[") > t.count("]"):
        return False
    if "[online]" in low or re.search(r"available\s*:", low):
        return bool(URL_RE.search(t)) and t.endswith((".", "]"))
    return bool(_YEAR_FIELD_END_RE.search(strip_trailing_metadata(t)))


def _is_entry_start(text: str) -> bool:
    return bool(ENTRY_START_RE.match(text.strip()))


def merge_wrapped_entries(paragraphs: list[dict]) -> list[dict]:
    """
    Gabungkan paragraf-paragraf yang sebenarnya satu entri Daftar Pustaka
    (terpecah per baris visual) menjadi satu elemen per entri.

    Paragraf dianggap AWAL entri baru kalau:
      - diawali nomor manual ("1.", "[1]"), ATAU
      - entri sebelumnya sudah LENGKAP (lihat entry_is_complete()) DAN baris
        ini bukan penanda metadata lanjutan ("[Accessed:", "doi:", ...), ATAU
      - ada paragraf kosong di antaranya (batas entri yang disengaja).
    Selain itu paragraf dianggap LANJUTAN dan ditempel ke entri sebelumnya.

    Kenapa "entri sebelumnya sudah lengkap" jadi syarat: entri yang
    benar-benar lupa diberi nomor tetap terdeteksi sebagai entri
    terpisah (dan kena cek penomoran), selama entri di atasnya lengkap.
    Keterbatasan yang disengaja: kalau entri di atasnya SENDIRI tidak
    berakhir tahun (kesalahan asli) DAN entri di bawahnya tidak bernomor,
    keduanya tergabung -- tidak ada sinyal teks yang bisa membedakannya
    dari baris yang terpecah.

    Elemen hasil = salinan paragraf pertama, dengan "full_text" gabungan
    (dipisah spasi), "index" = paragraf pertama, "index_end" = terakhir.
    """
    merged: list[dict] = []
    prev_raw_index = None
    for p in paragraphs:
        text = p["full_text"].strip()
        if not text or NON_ENTRY_RE.match(text):
            continue

        gap_blank = (
            prev_raw_index is not None
            and p["index"] - prev_raw_index > 1
            and p.get("blank_before", False)
        )
        prev_raw_index = p["index"]

        starts_new = (
            not merged
            or _is_entry_start(text)
            or gap_blank
            or (
                not CONTINUATION_START_RE.match(text)
                and entry_is_complete(merged[-1]["full_text"])
            )
        )
        if starts_new:
            entry = dict(p)
            entry["full_text"] = text
            entry["index_end"] = p["index"]
            merged.append(entry)
        else:
            merged[-1]["full_text"] += " " + text
            merged[-1]["index_end"] = p["index"]
    return merged


def extract_bibliography_entries(paragraphs: list[dict]) -> list[dict]:
    """
    Ambil entri Daftar Pustaka (chapter=="DAFTAR PUSTAKA"), terurut sesuai
    urutan kemunculan asli, dengan potongan baris yang terpecah sudah
    DIGABUNG per entri (lihat merge_wrapped_entries()). Heading "DAFTAR
    PUSTAKA" itu sendiri (type=="chapter_title") tidak ikut.
    """
    raw = [
        p
        for p in paragraphs
        if p["chapter"] == CHAPTER_LABEL and p["type"] in ENTRY_PARA_TYPES
    ]
    raw.sort(key=lambda p: p["index"])
    return merge_wrapped_entries(raw)


def _entry_location(entry: dict) -> str:
    """Label lokasi: 'Paragraf #671' atau 'Paragraf #671-#673' kalau entri
    terdiri dari beberapa paragraf."""
    start, end = entry["index"], entry.get("index_end", entry["index"])
    span = f"#{start}" if end == start else f"#{start}-#{end}"
    return f"Paragraf {span} ('{_text_preview(entry['full_text'])}')"


def check_entry_fields(entries: list[dict]) -> tuple[list[dict], dict]:
    """
    Cek field wajib PER ENTRI: diakhiri tahun, diawali nama pengarang,
    dan field spesifik sesuai jenis sumber (jurnal/TA/online).
    """
    findings = []
    checks = defaultdict(_new_counter)

    for entry in entries:
        raw_text = entry["full_text"]
        loc = _entry_location(entry)
        body, _ = strip_leading_number(raw_text)
        if not body:
            continue

        # --- cek 1: entri diakhiri Tahun (field terakhir) ---
        checks[CHAPTER_LABEL]["checked"] += 1
        if not entry_ends_with_year(body):
            checks[CHAPTER_LABEL]["failed"] += 1
            if YEAR_ANYWHERE_RE.search(body):
                desc = (
                    "Entri Daftar Pustaka ini mengandung tahun, tapi tidak di "
                    "posisi akhir -- tahun seharusnya jadi field TERAKHIR "
                    "(setelah nama publikasi/penerbit)."
                )
            else:
                desc = (
                    "Entri Daftar Pustaka ini tidak mengandung tahun sama "
                    "sekali, padahal tahun wajib jadi field terakhir."
                )
            findings.append(
                {
                    "category": "Format",
                    "chapter": CHAPTER_LABEL,
                    "location": loc,
                    "description": desc,
                }
            )

        # --- cek 2: entri diawali Nama pengarang ---
        checks[CHAPTER_LABEL]["checked"] += 1
        first_author = extract_first_author(body)
        if not first_author or first_author[0].isdigit():
            checks[CHAPTER_LABEL]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": CHAPTER_LABEL,
                    "location": loc,
                    "description": (
                        "Entri Daftar Pustaka seharusnya diawali Nama "
                        "pengarang/peneliti sebagai field pertama."
                    ),
                }
            )

        # --- cek 3: field wajib sesuai jenis sumber ---
        entry_type = classify_entry_type(body)
        upper = body.upper()

        if entry_type == "jurnal":
            # Sesuai ketentuan Panduan Buku PA (lihat contoh sistematika
            # no. 1 di header dokumen): field "Hal. aa-bb" untuk entri
            # jurnal/makalah bersifat OPSIONAL (tidak wajib ada), beda
            # dari field lain (Nama_Pengarang, Judul, Nama Jurnal, Vol.,
            # No., Penerbit, Tahun) yang tetap wajib dan sudah dicek di
            # cek 1 & 2 di atas. Karena itu TIDAK ADA pengecekan
            # kelengkapan halaman/DOI/Art.no. sama sekali untuk jenis
            # ini -- baik sebagai failed maupun sebagai soft-notice.
            pass

        elif entry_type == "tugas_akhir":
            checks[CHAPTER_LABEL]["checked"] += 1
            if not INSTITUTION_RE.search(upper):
                checks[CHAPTER_LABEL]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": CHAPTER_LABEL,
                        "location": loc,
                        "description": (
                            "Entri Proyek Akhir/Skripsi/Tesis/Disertasi ini "
                            "tidak menyebutkan nama Perguruan Tinggi, padahal "
                            "wajib dicantumkan."
                        ),
                    }
                )

        elif entry_type == "online":
            checks[CHAPTER_LABEL]["checked"] += 1
            if not URL_RE.search(body):
                checks[CHAPTER_LABEL]["failed"] += 1
                findings.append(
                    {
                        "category": "Format",
                        "chapter": CHAPTER_LABEL,
                        "location": loc,
                        "description": (
                            "Entri media online ini tidak mencantumkan alamat "
                            "internet (URL), padahal wajib ada."
                        ),
                    }
                )

    return findings, dict(checks)


def check_entry_numbering(entries: list[dict]) -> tuple[list[dict], dict]:
    """
    Cek nomor urut manual ("1.", "2.", ...) kalau entri MEMANG diberi
    nomor manual di teksnya. Kalau tidak satu pun entri punya nomor
    manual (kemungkinan besar dinomori via Word auto-numbering, yang
    tidak pernah muncul di para.text -- sama seperti kasus subbab di
    extractor.py), cek ini di-skip total: tidak nambah checked/failed
    sama sekali, supaya mahasiswa yang pakai auto-numbering tidak
    dirugikan hanya karena keterbatasan deteksi teks mentah.
    """
    findings = []
    checks = defaultdict(_new_counter)

    numbers = []
    for entry in entries:
        _, num = strip_leading_number(entry["full_text"])
        numbers.append((entry, num))

    has_any_number = any(num is not None for _, num in numbers)
    if not has_any_number:
        return findings, dict(checks)

    for i, (entry, num) in enumerate(numbers, start=1):
        loc = _entry_location(entry)
        checks[ORDER_GROUP]["checked"] += 1

        if num is None:
            checks[ORDER_GROUP]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": ORDER_GROUP,
                    "location": loc,
                    "description": (
                        "Entri ini tidak diberi nomor urut, padahal entri "
                        "Daftar Pustaka lain di dokumen ini memakai penomoran "
                        "manual -- penomoran harus konsisten untuk semua "
                        "entri."
                    ),
                }
            )
        elif num != i:
            checks[ORDER_GROUP]["failed"] += 1
            findings.append(
                {
                    "category": "Format",
                    "chapter": ORDER_GROUP,
                    "location": loc,
                    "description": (
                        f"Nomor urut Daftar Pustaka '{num}' tidak berurutan, "
                        f"seharusnya '{i}' (entri ke-{i})."
                    ),
                }
            )

    return findings, dict(checks)


def check_bibliography(paragraphs: list[dict]) -> tuple[list[dict], dict]:
    """
    Entry point sub-tahap Daftar Pustaka. Return (findings, checks)
    dengan BENTUK SAMA seperti compare_all()/check_page_layout(),
    supaya main.py tinggal menggabungkan ke findings/checks utama
    lewat merge_checks() yang sudah ada.
    """
    entries = extract_bibliography_entries(paragraphs)
    if not entries:
        return [], {}

    f1, c1 = check_entry_fields(entries)
    # Cek urutan alfabetis sudah dihapus total dari modul ini --
    # urutan alfabetis bukan aturan IEEE yang sebenarnya (IEEE asli
    # urut berdasarkan sitasi pertama di teks, yang tidak bisa kita
    # verifikasi dari Daftar Pustaka saja). Lihat CATATAN di docstring
    # modul ini untuk penjelasan lengkap.
    # Cek nomor urut dimatikan lewat CHECK_ENTRY_NUMBERING (lihat
    # komentar di deklarasinya); kodenya dipertahankan supaya bisa
    # dinyalakan lagi tanpa menulis ulang.
    f2, c2 = check_entry_numbering(entries) if CHECK_ENTRY_NUMBERING else ([], {})

    findings = f1 + f2
    checks = defaultdict(_new_counter)
    for group_checks in (c1, c2):
        for chapter, c in group_checks.items():
            checks[chapter]["checked"] += c["checked"]
            checks[chapter]["failed"] += c["failed"]

    return findings, dict(checks)
