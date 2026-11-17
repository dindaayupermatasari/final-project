"""Penilaian satu subbab dokumen mahasiswa terhadap checklist kriteria panduan."""
import json
import re
import unicodedata

from criteria import get_criteria
from llm_client import call_llm_json

SYSTEM = """Kamu adalah pemeriksa kesesuaian penulisan dokumen proyek akhir terhadap panduan.
Nilai HANYA berdasarkan kriteria yang diberikan dan teks mahasiswa. Jangan menilai gaya bahasa,
ejaan, atau hal lain di luar kriteria. Jangan mengarang: setiap penilaian "terpenuhi" harus
didukung kutipan persis dari teks mahasiswa. Kalau tidak ada bukti, nilainya false dan bukti
dikosongkan. Balas hanya dengan satu objek JSON, tanpa teks lain."""

USER_TEMPLATE = """JUDUL PROYEK AKHIR:
{judul}
(Problem Domain: {problem_domain} | Problem: {problem} | Uniqueness: {uniqueness})

SUBBAB YANG DINILAI: {subbab_title}

PANDUAN (acuan, bukan untuk disalin):
{guideline}

KRITERIA:
{kriteria}

TEKS MAHASISWA:
\"\"\"
{teks}
\"\"\"

Balas dengan JSON berformat:
{{
  "hasil": [
    {{"id": "<id kriteria>", "terpenuhi": true/false,
      "bukti": "<kalimat lengkap dari teks mahasiswa yang menjadi dasar, persis seperti tertulis; kosong jika tidak ada>",
      "alasan": "<satu kalimat>"}}
  ],
  "saran": "<maksimal dua kalimat perbaikan yang paling penting, atau string kosong>"
}}
Harus ada tepat satu entri hasil untuk setiap id kriteria."""


_CITE = re.compile(r"\[\s*\d+(?:\s*[,;\u2013-]\s*\d+)*\s*\]")  # penanda sitasi seperti [2] [8] [23]


def match_key(s: str) -> str:
    """Kunci pembanding kutipan vs teks: buang penanda sitasi [n], samakan tanda hubung/kutip,
    NFKC, abaikan semua spasi dan huruf besar/kecil. Model sering membuang sitasi saat mengutip."""
    s = _CITE.sub("", s or "")
    s = re.sub("[\u2010-\u2015\u2212]", "-", s)
    s = s.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")
    return "".join(unicodedata.normalize("NFKC", s).split()).lower()


def clean_text(t: str) -> str:
    """Rapikan teks hasil ekstraksi PDF: sambung kata putus di akhir baris, rapatkan spasi,
    buang nomor halaman yang menempel di ujung teks."""
    t = re.sub(r"-\s*\n\s*", "-", t)
    t = " ".join(t.split())
    return re.sub(r"\s+\d{1,3}$", "", t)


def judge_subbab(decomp: dict, subbab_title: str, student_text: str, guideline_text: str = "") -> dict:
    """decomp = hasil title_decomposer.decompose_title (judul, problem_domain, problem, uniqueness)."""
    spec = get_criteria(subbab_title)
    if spec is None:
        return {"status": "tidak_ada_kriteria", "subbab": subbab_title}

    ctx = {k: (decomp.get(k) or "(tidak ada di judul)") for k in ("problem_domain", "problem", "uniqueness")}
    student_text = clean_text(student_text)
    kriteria_txt = "\n".join(
        f"- {k['id']}: {k['deskripsi'].format(**ctx)}" + ("" if k["wajib"] else " (anjuran)")
        + ("" if k.get("butuh_bukti", True) else " [tanpa kutipan: jelaskan di alasan]")
        for k in spec["kriteria"]
    )
    user = USER_TEMPLATE.format(
        judul=decomp.get("judul", ""),
        **ctx,
        subbab_title=subbab_title,
        guideline=guideline_text or "(tidak disertakan)",
        kriteria=kriteria_txt,
        teks=student_text.strip(),
    )
    raw = call_llm_json(SYSTEM, user)

    # Validasi: bukti harus benar-benar ada di teks (cegah kutipan karangan)
    norm = match_key
    teks_norm = norm(student_text)
    by_id = {h.get("id"): h for h in raw.get("hasil", [])}
    hasil = []
    for k in spec["kriteria"]:
        h = by_id.get(k["id"], {"id": k["id"], "terpenuhi": False, "bukti": "", "alasan": "tidak dinilai model"})
        bukti_valid = (not h.get("bukti")) or norm(h["bukti"]) in teks_norm
        perlu = k.get("butuh_bukti", True)
        if h.get("terpenuhi") and perlu and (not h.get("bukti") or not bukti_valid):
            h["terpenuhi"] = False
            h["alasan"] = "bukti tidak ditemukan di teks; " + h.get("alasan", "")
        h["wajib"] = k["wajib"]
        hasil.append(h)

    # Verdict dihitung di kode, bukan oleh LLM
    wajib = [h for h in hasil if h["wajib"]]
    anjuran = [h for h in hasil if not h["wajib"]]
    if all(h["terpenuhi"] for h in wajib):
        status = "sesuai"
    elif any(h["terpenuhi"] for h in wajib):
        status = "sebagian"
    else:
        status = "tidak_sesuai"
    return {
        "status": status,
        "subbab": subbab_title,
        "wajib_terpenuhi": f"{sum(h['terpenuhi'] for h in wajib)}/{len(wajib)}",
        "anjuran_terpenuhi": f"{sum(h['terpenuhi'] for h in anjuran)}/{len(anjuran)}",
        "hasil": hasil,
        "saran": raw.get("saran", ""),
    }
