"""
parser.py — pipelines/sistematika/

Baca file PDF panduan penulisan dan pecah teksnya jadi struktur BAB > sub-bab.
Modul ini FOKUS baca & strukturkan teks mentah saja — tidak melakukan
filtering konten atau embedding (itu tugas chunker.py dan vector_store.py).
Dipisah begini supaya kalau nanti kamu juga butuh parser untuk buku PA
mahasiswa yang diupload, polanya konsisten: baca -> strukturkan -> proses.
"""

import re
from dataclasses import dataclass, field
from typing import List

import fitz  # PyMuPDF

BAB_PATTERN = re.compile(r"^BAB\s+(\d+)\s*$", re.MULTILINE)
SUBBAB_PATTERN = re.compile(
    r"^(\d+\.\d+(?:\.\d+)?)\s+([A-Za-z][A-Za-z0-9 ,\-]{2,60})$",
    re.MULTILINE,
)


@dataclass
class Section:
    bab_num: int
    bab_title: str
    subbab_num: str
    subbab_title: str
    raw_text: str
    kept_sentences: List[str] = field(default_factory=list)


def extract_text_by_page(pdf_path: str) -> str:
    """Baca seluruh halaman PDF, gabung jadi satu string teks mentah."""
    doc = fitz.open(pdf_path)
    full_text = [page.get_text("text") for page in doc]
    doc.close()
    return "\n".join(full_text)


def split_into_sections(full_text: str) -> List[Section]:
    """Pecah teks penuh jadi daftar Section berdasarkan heading BAB dan sub-bab."""
    bab_matches = list(BAB_PATTERN.finditer(full_text))
    sections: List[Section] = []

    for i, bab_match in enumerate(bab_matches):
        bab_num = int(bab_match.group(1))
        bab_start = bab_match.end()
        bab_end = (
            bab_matches[i + 1].start() if i + 1 < len(bab_matches) else len(full_text)
        )
        bab_block = full_text[bab_start:bab_end]

        first_line = bab_block.strip().split("\n", 1)[0].strip()
        bab_title = first_line if first_line.isupper() else f"BAB {bab_num}"

        subbab_matches = list(SUBBAB_PATTERN.finditer(bab_block))
        if not subbab_matches:
            sections.append(
                Section(
                    bab_num=bab_num,
                    bab_title=bab_title,
                    subbab_num=f"{bab_num}.0",
                    subbab_title=bab_title,
                    raw_text=bab_block.strip(),
                )
            )
            continue

        for j, sub_match in enumerate(subbab_matches):
            sub_start = sub_match.end()
            sub_end = (
                subbab_matches[j + 1].start()
                if j + 1 < len(subbab_matches)
                else len(bab_block)
            )
            sections.append(
                Section(
                    bab_num=bab_num,
                    bab_title=bab_title,
                    subbab_num=sub_match.group(1),
                    subbab_title=sub_match.group(2).strip(),
                    raw_text=bab_block[sub_start:sub_end].strip(),
                )
            )

    return sections
