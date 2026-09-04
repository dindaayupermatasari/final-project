"""
chunker.py — pipelines/sistematika/

Ambil Section hasil parser.py, saring kalimat instruksi format visual,
lalu bangun chunk final yang siap di-embed. Dipisah dari parser supaya
logic "apa yang layak masuk vector store" gampang di-tuning/di-tes sendiri
tanpa harus utak-atik logic baca PDF.
"""

import re
from typing import List

from parser import Section

FORMAT_KEYWORDS = [
    "times new roman",
    "font",
    "ukuran huruf",
    "huruf kapital",
    "cetak tebal",
    "italic",
    "cetak miring",
    "margin",
    "spasi",
    "kertas a4",
    "a4-80 gram",
    "nomor halaman",
    "penomoran halaman",
    "footer",
    "header",
    "hexa",
    "rgb",
    "mirror-margin",
    "poin",
    "point",
    " pt ",
    "cm)",
    "inch",
    "rata tengah",
    "rata kanan",
    "rata kiri",
    "bolak-balik",
    "warna biru",
    "kode hexa",
]


def split_into_sentences(raw_text: str) -> List[str]:
    """Ratakan baris lalu pecah per kalimat (bukan per baris kosong -- lihat catatan di README)."""
    joined = re.sub(r"\s+", " ", raw_text).strip()
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", joined) if s.strip()]


def filter_section(section: Section) -> None:
    """Isi section.kept_sentences dengan kalimat yang BUKAN instruksi format visual."""
    kept = []
    for sent in split_into_sentences(section.raw_text):
        if not any(kw in sent.lower() for kw in FORMAT_KEYWORDS):
            kept.append(sent)
    section.kept_sentences = kept


def build_chunks(sections: List[Section]) -> List[dict]:
    """Ubah tiap Section (yang sudah difilter) jadi chunk dict siap embed."""
    chunks = []
    for sec in sections:
        if not sec.kept_sentences:
            continue
        chunks.append(
            {
                "id": f"bab{sec.bab_num}_sub{sec.subbab_num.replace('.', '-')}",
                "text": " ".join(sec.kept_sentences),
                "metadata": {
                    "bab_num": sec.bab_num,
                    "bab_title": sec.bab_title,
                    "subbab_num": sec.subbab_num,
                    "subbab_title": sec.subbab_title,
                },
            }
        )
    return chunks
