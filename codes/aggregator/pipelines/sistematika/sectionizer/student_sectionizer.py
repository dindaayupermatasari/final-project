"""
student_sectionizer.py — pipelines/sistematika/

Pecah dokumen ASLI mahasiswa (.docx / .pdf) jadi struktur BAB > sub-bab.
Dipisah dari parser.py karena parser.py didesain untuk PDF panduan yang
rapi; dokumen mahasiswa campur narasi + tabel, sehingga baris tabel seperti
"1.1  Jumlah Axle" atau "1.067  mm" bisa ke-parse sebagai heading.

  .docx -> PRIORITAS heading STYLE Word (Heading 1 = BAB, Heading 2/3 =
           sub-bab). document.paragraphs otomatis tidak menyertakan isi
           tabel. Kalau tidak ada heading style sama sekali -> fallback regex.

  .pdf  -> regex + 2 validasi:
             1. nomor sub-bab wajib diawali nomor bab aktif
             2. digit setelah titik dibatasi 1-2 angka

  Bagian belakang (DAFTAR PUSTAKA, LAMPIRAN) dipotong dari isi sub-bab.

Perubahan versi ini:
  * BUG FIX: penomoran fallback sub-bab di jalur heading style tidak lagi
    meleset (section "x.0" tidak menaikkan penghitung sub-bab).
  * BAB_PATTERN menerima "BAB 1", "BAB I" (romawi), dan "BAB I PENDAHULUAN"
    / "BAB 1 - PENDAHULUAN" dalam satu baris (judul same-line harus HURUF
    BESAR supaya kalimat "BAB 3 dan BAB 4 membahas ..." tidak ketipu).
  * Daftar isi & running header PDF: kemunculan "BAB n" berurutan dengan
    nomor sama digabung; kalau nomor bab muncul di beberapa tempat (daftar
    isi vs isi sebenarnya) dipilih blok terpanjang.
  * Karakter judul sub-bab menerima ( ) / &.
  * Nomor halaman di PDF dibuang (memakai fungsi dari parser.py).
  * Kandidat yang ditolak validasi tetap dicatat di `rejected_candidates`.
"""

import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

import fitz  # PyMuPDF

from parser import _strip_page_number_lines

# "BAB 1" / "BAB I" sendirian, atau diikuti judul HURUF BESAR di baris yang sama.
BAB_PATTERN = re.compile(
    r"^[ \t]*BAB[ \t]+(?P<num>\d+|[IVX]+)\b[ \t]*[-–:.]?[ \t]*"
    r"(?P<title>[A-Z0-9 ,&()/\-]{0,60}?)[ \t]*$",
    re.MULTILINE,
)
# Untuk teks heading docx (boleh "Bab I", "BAB 1 PENDAHULUAN", dst).
BAB_IN_TEXT = re.compile(r"\bBAB\s+(\d+|[IVX]+)\b", re.IGNORECASE)

SUBBAB_PATTERN = re.compile(
    r"^(\d+\.\d{1,2}(?:\.\d{1,2})?)\s+([A-Za-z][A-Za-z0-9 ,\-'()/&]{2,60})$",
    re.MULTILINE,
)

_BACKMATTER_BODY = r"DAFTAR PUSTAKA|LAMPIRAN(?:[ \t]+[A-Z0-9]{1,3})?"
BACKMATTER_PATTERN = re.compile(rf"^[ \t]*(?:{_BACKMATTER_BODY})[ \t]*$", re.MULTILINE)
_BACKMATTER_EXACT = re.compile(_BACKMATTER_BODY)
_BACKMATTER_EXACT_ANYCASE = re.compile(_BACKMATTER_BODY, re.IGNORECASE)

_ROMAN = {"I": 1, "V": 5, "X": 10}


def _bab_to_int(token: str) -> int:
    token = token.strip().upper()
    if token.isdigit():
        return int(token)
    total, prev = 0, 0
    for ch in reversed(token):
        v = _ROMAN[ch]
        if v < prev:
            total -= v
        else:
            total += v
            prev = v
    return total


@dataclass
class Section:
    bab_num: int
    bab_title: str
    subbab_num: str
    subbab_title: str
    raw_text: str
    source: str = "regex"  # "heading_style" (dari .docx) atau "regex"


@dataclass
class RejectedCandidate:
    bab_num: int
    matched_num: str
    matched_title: str
    reason: str


# ---------------------------------------------------------------------
# .docx -- lewat heading style Word asli
# ---------------------------------------------------------------------


def extract_paragraphs_docx(path: str) -> List[Tuple[str, str]]:
    """List (text, style_name) dalam URUTAN dokumen.

    Isi tabel ikut dibaca (satu baris tabel = satu entri bergaya "Table",
    sel dipisah tab, sel gabungan hanya dihitung sekali), karena sebagian
    mahasiswa menaruh isi Sistematika Penulisan di dalam tabel. Entri tabel
    tidak pernah dianggap heading, jadi baris seperti "1.1  Jumlah Axle"
    tidak akan memecah sub-bab.
    """
    import docx  # python-docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = docx.Document(path)
    out: List[Tuple[str, str]] = []

    def walk_table(tbl: "Table") -> None:
        for row in tbl.rows:
            seen, cells = set(), []
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                txt = " ".join(cell.text.split())
                if txt and txt not in cells:
                    cells.append(txt)
            if cells:
                out.append(("\t".join(cells), "Table"))

    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para = Paragraph(child, document)
            out.append((para.text, para.style.name if para.style else "Normal"))
        elif tag == "tbl":
            walk_table(Table(child, document))
    return out


def _is_backmatter_heading(text: str, style: str) -> bool:
    t = " ".join(text.split())
    if _BACKMATTER_EXACT.fullmatch(t):
        return True
    return style.lower().startswith("heading") and bool(
        _BACKMATTER_EXACT_ANYCASE.fullmatch(t)
    )


def truncate_paragraphs_at_backmatter(
    paragraphs: List[Tuple[str, str]],
) -> List[Tuple[str, str]]:
    """Potong tepat sebelum DAFTAR PUSTAKA / LAMPIRAN pertama, dicari SESUDAH
    heading isi pertama supaya daftar isi tidak ikut kena."""
    guard = None
    for i, (text, style) in enumerate(paragraphs):
        if style.lower().startswith("heading") and BAB_IN_TEXT.search(text):
            guard = i
            break
    if guard is None:
        for i, (text, style) in enumerate(paragraphs):
            if style.lower() == "heading 1" and text.strip():
                guard = i
                break
    if guard is None:
        return paragraphs
    for i in range(guard + 1, len(paragraphs)):
        text, style = paragraphs[i]
        if _is_backmatter_heading(text, style):
            return paragraphs[:i]
    return paragraphs


# Heading BAB diawali kata "BAB n" (n angka / romawi), mis. "BAB 1\nPENDAHULUAN".
BAB_START = re.compile(r"^\s*BAB\s+(\d+|[IVX]+)\b", re.IGNORECASE)
# Style kustom untuk sub-bab: "Heading Subab", "Heading Subab2", "Heading Subabs5",
# "Heading Sub-bab 3", dst. Angka di belakangnya = nomor bab, BUKAN level heading.
_CUSTOM_SUBBAB_STYLE = re.compile(r"^(?:heading|judul)\s*sub\s*-?\s*b?ab", re.IGNORECASE)
_CUSTOM_BAB_STYLE = re.compile(r"^(?:heading|judul)\s*bab\s*$", re.IGNORECASE)
_NUMBERED_STYLE = re.compile(r"^(?:heading|judul)\s*(\d)\s*$", re.IGNORECASE)


def _heading_level(text: str, style: str, uses_custom_subbab: bool) -> Optional[int]:
    """Level heading: 1 = BAB, 2 = sub-bab, 3+ = anak sub-bab. None = bukan heading.

    Level TIDAK diambil dari angka pertama di nama style, karena pada template
    kustom (\"Heading Subab3\") angka itu nomor bab, bukan level.
    """
    s = style.strip()
    if not s.lower().startswith(("heading", "judul")):
        return None
    if BAB_START.match(text):
        return 1
    if _CUSTOM_BAB_STYLE.match(s):
        return 1
    if _CUSTOM_SUBBAB_STYLE.match(s):
        return 2
    m = _NUMBERED_STYLE.match(s)
    if m:
        n = int(m.group(1))
        if n == 1:
            return 1
        # Kalau dokumen memakai style kustom sub-bab, \"Heading 2\" biasanya
        # dipakai untuk anak sub-bab (mis. 2.2.1), bukan sub-bab itu sendiri.
        return n + 1 if uses_custom_subbab else n
    return 2


def split_by_heading_style(
    paragraphs: List[Tuple[str, str]],
) -> Optional[List[Section]]:
    """Pecah berdasarkan heading style asli. None kalau tak ada heading style."""
    heading_idx = [
        i
        for i, (text, style) in enumerate(paragraphs)
        if style.lower().startswith(("heading", "judul")) and text.strip()
    ]
    if not heading_idx:
        return None

    uses_custom_subbab = any(
        _CUSTOM_SUBBAB_STYLE.match(paragraphs[i][1].strip()) for i in heading_idx
    )
    # Ada heading yang tertulis \"BAB n\"? Kalau ya, heading level-1 tanpa \"BAB n\"
    # (PERNYATAAN, ABSTRAK, KATA PENGANTAR, DAFTAR ISI, dst.) dianggap bagian
    # depan / belakang dan DILEWATI, bukan dianggap bab baru.
    has_bab_text = any(BAB_START.match(paragraphs[i][0]) for i in heading_idx)

    sections: List[Section] = []
    cur_bab_num = 0
    cur_bab_title = ""
    cur_sub_count = 0
    cur_subsub_count = 0
    cur_sub_num = ""
    last_bab_pos = -1  # posisi di heading_idx dari heading BAB terakhir

    def body_between(start: int, end: int) -> str:
        return "\n".join(t for t, _ in paragraphs[start:end] if t.strip()).strip()

    def add_intro(body: str) -> None:
        # teks pembuka bab (sebelum sub-bab pertama); TIDAK menaikkan
        # cur_sub_count supaya sub-bab pertama tetap "x.1"
        if body:
            sections.append(
                Section(
                    bab_num=cur_bab_num,
                    bab_title=cur_bab_title,
                    subbab_num=f"{cur_bab_num}.0",
                    subbab_title=cur_bab_title,
                    raw_text=body,
                    source="heading_style",
                )
            )

    for pos, idx in enumerate(heading_idx):
        text, style = paragraphs[idx]
        level = _heading_level(text, style, uses_custom_subbab)
        next_idx = (
            heading_idx[pos + 1] if pos + 1 < len(heading_idx) else len(paragraphs)
        )
        body = body_between(idx + 1, next_idx)
        clean = " ".join(text.split())

        if level == 1:
            bab_match = BAB_START.match(text)
            if bab_match:
                cur_bab_num = _bab_to_int(bab_match.group(1))
                cur_bab_title = clean
                cur_sub_count = cur_subsub_count = 0
                cur_sub_num = ""
                last_bab_pos = pos
                add_intro(body)
                continue

            if not has_bab_text:  # dokumen tanpa tulisan "BAB n": penomoran otomatis
                cur_bab_num += 1
                cur_bab_title = clean
                cur_sub_count = cur_subsub_count = 0
                cur_sub_num = ""
                last_bab_pos = pos
                add_intro(body)
                continue

            # Judul bab yang ditulis di paragraf heading terpisah:
            #   [Heading] BAB 4   +   [Heading] EKSPERIMEN DAN ANALISIS
            prev_empty = (
                last_bab_pos == pos - 1
                and cur_bab_num > 0
                and not body_between(heading_idx[last_bab_pos] + 1, idx)
            )
            if prev_empty:
                cur_bab_title = f"{cur_bab_title} {clean}"
                last_bab_pos = pos
                add_intro(body)
                continue

            # Heading level-1 lain (ABSTRAK, DAFTAR ..., BIODATA): bukan isi bab.
            cur_bab_num = 0
            cur_bab_title = ""
            continue

        if cur_bab_num == 0:  # masih di bagian depan / sesudah bab terakhir
            continue

        num_match = re.match(r"^(\d+(?:\.\d+)+)\s+(.*)", clean)

        if level >= 3 and cur_sub_num:
            # anak sub-bab -> nomor x.y.z, nanti digabung ke induknya (merge_subsub)
            cur_subsub_count += 1
            if num_match:
                subbab_num, subbab_title = num_match.group(1), num_match.group(2).strip()
            else:
                subbab_num = f"{cur_sub_num}.{cur_subsub_count}"
                subbab_title = clean
        else:
            cur_sub_count += 1
            cur_subsub_count = 0
            if num_match:
                subbab_num, subbab_title = num_match.group(1), num_match.group(2).strip()
            else:
                subbab_num = f"{cur_bab_num}.{cur_sub_count}"
                subbab_title = clean
            cur_sub_num = subbab_num

        sections.append(
            Section(
                bab_num=cur_bab_num,
                bab_title=cur_bab_title,
                subbab_num=subbab_num,
                subbab_title=subbab_title,
                raw_text=body,
                source="heading_style",
            )
        )

    return sections


# ---------------------------------------------------------------------
# .pdf (dan fallback .docx tanpa heading style) -- regex + validasi
# ---------------------------------------------------------------------


def extract_text_pdf(path: str) -> str:
    doc = fitz.open(path)
    full_text = [_strip_page_number_lines(page.get_text("text")) for page in doc]
    doc.close()
    return "\n".join(full_text)


def _is_valid_subbab(bab_num: int, subbab_num: str) -> bool:
    """Nomor sub-bab WAJIB diawali nomor bab yang sedang aktif."""
    return subbab_num.startswith(f"{bab_num}.")


def _collect_bab_blocks(full_text: str) -> List[dict]:
    """Cari blok isi tiap BAB.

    - Heading "BAB n" berurutan dengan nomor sama (running header) digabung.
    - Blok dipotong di DAFTAR PUSTAKA / LAMPIRAN; blok kosong dibuang.
    - Nomor bab yang muncul di beberapa tempat (daftar isi vs isi asli):
      dipilih blok terpanjang.
    """
    matches = list(BAB_PATTERN.finditer(full_text))
    runs: List[dict] = []
    for i, m in enumerate(matches):
        num = _bab_to_int(m.group("num"))
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        if runs and runs[-1]["num"] == num:
            runs[-1]["end"] = end
        else:
            runs.append(
                {
                    "num": num,
                    "title": m.group("title").strip(),
                    "start": m.end(),
                    "end": end,
                }
            )

    best = {}
    for r in runs:
        block = full_text[r["start"] : r["end"]]
        backmatter = BACKMATTER_PATTERN.search(block)
        if backmatter:
            block = block[: backmatter.start()]
        block = BAB_PATTERN.sub("", block)  # buang sisa running header
        if not block.strip():
            continue
        cand = {**r, "block": block}
        cur = best.get(r["num"])
        if cur is None or len(block.strip()) > len(cur["block"].strip()):
            best[r["num"]] = cand
    return sorted(best.values(), key=lambda b: b["start"])


def split_by_regex(full_text: str) -> Tuple[List[Section], List[RejectedCandidate]]:
    sections: List[Section] = []
    rejected: List[RejectedCandidate] = []

    for b in _collect_bab_blocks(full_text):
        bab_num, bab_block = b["num"], b["block"]

        first_line = bab_block.strip().split("\n", 1)[0].strip()
        bab_title = b["title"] or (
            first_line if first_line.isupper() else f"BAB {bab_num}"
        )

        valid_matches = []
        for sm in SUBBAB_PATTERN.finditer(bab_block):
            num, title = sm.group(1), sm.group(2).strip()
            if _is_valid_subbab(bab_num, num):
                valid_matches.append(sm)
            else:
                rejected.append(
                    RejectedCandidate(
                        bab_num=bab_num,
                        matched_num=num,
                        matched_title=title,
                        reason=f"nomor '{num}' tidak diawali '{bab_num}.' -- kemungkinan baris tabel/data, bukan heading sub-bab",
                    )
                )

        if not valid_matches:
            sections.append(
                Section(
                    bab_num=bab_num,
                    bab_title=bab_title,
                    subbab_num=f"{bab_num}.0",
                    subbab_title=bab_title,
                    raw_text=bab_block.strip(),
                    source="regex",
                )
            )
            continue

        for j, sub_match in enumerate(valid_matches):
            sub_end = (
                valid_matches[j + 1].start()
                if j + 1 < len(valid_matches)
                else len(bab_block)
            )
            sections.append(
                Section(
                    bab_num=bab_num,
                    bab_title=bab_title,
                    subbab_num=sub_match.group(1),
                    subbab_title=sub_match.group(2).strip(),
                    raw_text=bab_block[sub_match.end() : sub_end].strip(),
                    source="regex",
                )
            )

    return sections, rejected


# ---------------------------------------------------------------------
# Entry point publik
# ---------------------------------------------------------------------


def sectionize_student_doc(path: str) -> Tuple[List[Section], List[RejectedCandidate]]:
    ext = path.lower().rsplit(".", 1)[-1]

    if ext == "docx":
        paragraphs = extract_paragraphs_docx(path)
        paragraphs = truncate_paragraphs_at_backmatter(paragraphs)
        sections = split_by_heading_style(paragraphs)
        if sections is not None:
            return sections, []
        full_text = "\n".join(t for t, _ in paragraphs)
        return split_by_regex(full_text)

    elif ext == "pdf":
        return split_by_regex(extract_text_pdf(path))

    raise ValueError(f"Format belum didukung: .{ext} (cuma .docx / .pdf)")
