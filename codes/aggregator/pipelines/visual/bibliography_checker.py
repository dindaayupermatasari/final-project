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
       - Jurnal   : ada "Vol."/"No." -> wajib juga ada "Hal."
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
# Tahun boleh diikuti metadata penutup standar IEEE (DOI, [Online].
# Available:, Accessed:) -- ini BUKAN identitas publikasi (nama
# jurnal/penerbit) yang seharusnya mendahului tahun, jadi tidak
# melanggar aturan "tahun jadi field terakhir sebelum metadata akses".
# Tanpa ini, hampir semua entri jurnal IEEE modern yang menutup
# dengan "..., 2023, doi: 10.xxxx/yyyy." akan salah dianggap error,
# padahal itu justru format IEEE yang benar.
_TRAILING_META = (
    r"(?:,?\s*"
    r"(?:doi\s*:\s*\S+"
    r"|\[online\]\.?\s*available\s*:.*"
    r"|available\s*:.*"
    r"|accessed\s*:.*"
    r")"
    r")?"
)
YEAR_AT_END_RE = re.compile(
    r"(19|20)\d{2}\.?\s*" + _TRAILING_META + r"\.?\s*$",
    re.IGNORECASE | re.DOTALL,
)
YEAR_ANYWHERE_RE = re.compile(r"(19|20)\d{2}")
URL_RE = re.compile(r"(https?://|www\.)\S+", re.IGNORECASE)
DOI_OR_ARTNO_RE = re.compile(
    r"\bDOI\b|\b10\.\d{4,9}/\S+|\bART\.?\s*NO\.?\s*\d|\bARTICLE\s*NO\.?\s*\d",
    re.IGNORECASE,
)
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

    if "DIAKSES" in upper and URL_RE.search(body):
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


def extract_bibliography_entries(paragraphs: list[dict]) -> list[dict]:
    """
    Ambil paragraf entri Daftar Pustaka (chapter=="DAFTAR PUSTAKA",
    type=="body"), terurut sesuai urutan kemunculan asli di dokumen.
    """
    entries = [
        p for p in paragraphs if p["chapter"] == CHAPTER_LABEL and p["type"] == "body"
    ]
    entries.sort(key=lambda p: p["index"])
    return entries


def check_entry_fields(entries: list[dict]) -> tuple[list[dict], dict]:
    """
    Cek field wajib PER ENTRI: diakhiri tahun, diawali nama pengarang,
    dan field spesifik sesuai jenis sumber (jurnal/TA/online).
    """
    findings = []
    checks = defaultdict(_new_counter)

    for entry in entries:
        raw_text = entry["full_text"]
        loc = f"Paragraf #{entry['index']} ('{_text_preview(raw_text)}')"
        body, _ = strip_leading_number(raw_text)
        if not body:
            continue

        # --- cek 1: entri diakhiri Tahun (field terakhir) ---
        checks[CHAPTER_LABEL]["checked"] += 1
        if not YEAR_AT_END_RE.search(body):
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
            # Jurnal open-access/early-access modern (IEEE Access, MDPI,
            # dll) sering pakai DOI atau "Art. no." sebagai pengganti sah
            # rentang halaman -- itu format asli sumbernya, bukan
            # kesalahan penulisan mahasiswa. Jadi kedua bentuk ini
            # diterima sebagai pemenuhan syarat "identitas lokasi
            # artikel", tidak wajib literally "Hal. aa-bb".
            has_pages = re.search(r"\bHAL\.?\s*\d", upper)
            has_doi_or_artno = DOI_OR_ARTNO_RE.search(body)
            checks[CHAPTER_LABEL]["checked"] += 1
            if not has_pages and not has_doi_or_artno:
                # TIDAK dihitung sebagai "failed" -- ini soft-notice,
                # bukan penalti skor, karena kita tidak bisa pastikan
                # dari teks saja apakah jurnal sumbernya memang tidak
                # menerbitkan halaman/DOI/Art.no. sama sekali (early
                # access, dsb). Tetap dilaporkan supaya mahasiswa bisa
                # cek manual.
                findings.append(
                    {
                        "category": "Format",
                        "chapter": CHAPTER_LABEL,
                        "location": loc,
                        "description": (
                            "Entri jurnal ini punya 'Vol.' tapi tidak "
                            "ditemukan 'Hal. aa-bb', DOI, atau nomor "
                            "artikel. Kalau jurnal sumbernya memang "
                            "menerbitkan salah satu dari itu, sebaiknya "
                            "dicantumkan; kalau sumbernya memang tidak "
                            "punya (mis. early access), ini boleh "
                            "diabaikan."
                        ),
                    }
                )

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
        loc = f"Paragraf #{entry['index']} ('{_text_preview(entry['full_text'])}')"
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
    f2, c2 = check_entry_numbering(entries)

    findings = f1 + f2
    checks = defaultdict(_new_counter)
    for group_checks in (c1, c2):
        for chapter, c in group_checks.items():
            checks[chapter]["checked"] += c["checked"]
            checks[chapter]["failed"] += c["failed"]

    return findings, dict(checks)
