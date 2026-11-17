"""
chunker.py — pipelines/sistematika/

Ambil Section hasil parser.py, saring kalimat instruksi format visual,
klasifikasikan sisa kalimat (instruction / example / template), lalu
bangun chunk final yang siap di-embed.

Perubahan:
  * Tiap kalimat diklasifikasikan:
      template    -> placeholder ("Aaaa... Bbbb", "Disini penulis dapat ...")
      example     -> kalimat contoh ("Misalkan penulis mengambil judul ...
                     kanker ...")
      instruction -> sisanya
  * Satu sub-bab bisa menghasilkan sampai 3 chunk (satu per tipe). Tipe
    disimpan di metadata["chunk_type"] supaya bisa difilter saat query
    (where={"chunk_type": "instruction"}) tanpa membuang contoh.
  * Sub-bab "Bagian 1/2/3" dan sub-bab yang mayoritas placeholder
    ditandai template seluruhnya.

Heuristik contoh/template ada di konstanta bagian atas; tuning-nya lewat
`python ingest_guideline.py --dry-run` (menampilkan tiap chunk + tipenya).
"""

import re
from typing import Dict, List

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

# --- Heuristik klasifikasi (sesuaikan setelah melihat hasil --dry-run) ---

# Kalimat contoh: dibuka kata pengantar contoh di AWAL kalimat ...
EXAMPLE_START = re.compile(
    r"^\W*(misalkan|misalnya|contoh|sebagai contoh|umpamanya)\b", re.IGNORECASE
)
# ... atau menyebut topik contoh yang dipakai panduan.
EXAMPLE_TOPIC = re.compile(r"\b(kanker|sistem pakar|fuzzy)\b", re.IGNORECASE)

# Placeholder template: huruf berulang ("Aaaa", "bbbb") atau kalimat boilerplate.
TEMPLATE_PLACEHOLDER = re.compile(r"\b([A-Za-z])\1{2,}\b", re.IGNORECASE)
TEMPLATE_BOILERPLATE = re.compile(r"^\W*disini penulis dapat", re.IGNORECASE)
TEMPLATE_TITLE = re.compile(r"^\s*bagian\s+\d+\b", re.IGNORECASE)
TEMPLATE_RATIO = 0.5  # > 50%  kalimat template -> seluruh sub-bab dianggap template

CHUNK_TYPES = ("instruction", "example", "template")


def split_into_sentences(raw_text: str) -> List[str]:
    """Ratakan baris lalu pecah per kalimat."""
    joined = re.sub(r"\s+", " ", raw_text).strip()
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", joined) if s.strip()]


def filter_section(section: Section) -> None:
    """Isi section.kept_sentences dengan kalimat yang BUKAN instruksi format visual."""
    kept = []
    for sent in split_into_sentences(section.raw_text):
        if not any(kw in sent.lower() for kw in FORMAT_KEYWORDS):
            kept.append(sent)
    section.kept_sentences = kept


def classify_sentence(sent: str) -> str:
    if TEMPLATE_PLACEHOLDER.search(sent) or TEMPLATE_BOILERPLATE.search(sent):
        return "template"
    if EXAMPLE_START.search(sent) or EXAMPLE_TOPIC.search(sent):
        return "example"
    return "instruction"


def group_sentences(section: Section) -> Dict[str, List[str]]:
    """Kelompokkan kalimat section per chunk_type."""
    if TEMPLATE_TITLE.match(section.subbab_title or ""):
        return {"template": list(section.kept_sentences)}

    groups: Dict[str, List[str]] = {t: [] for t in CHUNK_TYPES}
    for sent in section.kept_sentences:
        groups[classify_sentence(sent)].append(sent)

    total = len(section.kept_sentences)
    if total and len(groups["template"]) / total > TEMPLATE_RATIO:
        return {"template": list(section.kept_sentences)}
    return groups


def build_chunks(
    sections: List[Section], include_templates: bool = False
) -> List[dict]:
    """Ubah tiap Section (yang sudah difilter) jadi chunk dict siap embed.

    include_templates=True mengembalikan juga chunk template (dipakai
    dry-run untuk audit); default-nya dibuang.
    """
    chunks: List[dict] = []
    seen_ids = set()
    for sec in sections:
        if not sec.kept_sentences:
            continue
        groups = group_sentences(sec)
        for ctype in CHUNK_TYPES:
            sents = groups.get(ctype)
            if not sents:
                continue
            if ctype == "template" and not include_templates:
                continue

            base_id = f"bab{sec.bab_num}_sub{sec.subbab_num.replace('.', '-')}_{ctype}"
            cid, n = base_id, 2
            while cid in seen_ids:  # heading ganda / nomor kembar
                cid = f"{base_id}_{n}"
                n += 1
            seen_ids.add(cid)

            chunks.append(
                {
                    "id": cid,
                    "text": " ".join(sents),
                    "metadata": {
                        "bab_num": sec.bab_num,
                        "bab_title": sec.bab_title,
                        "subbab_num": sec.subbab_num,
                        "subbab_title": sec.subbab_title,
                        "chunk_type": ctype,
                    },
                }
            )
    return chunks
