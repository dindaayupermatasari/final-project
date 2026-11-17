"""Langkah 0: dokumen mahasiswa (.pdf/.docx) -> daftar subbab, lewat student_sectionizer.py.
Juga: ambil judul dari halaman sampul (opsional, kalau --judul tidak diberikan)."""
import sys
from dataclasses import asdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent  # folder sistematika/
sys.path.insert(0, str(BASE / "sectionizer"))
sys.path.insert(0, str(BASE / "rag"))  # student_sectionizer memakai `from parser import ...`


# True : teks subbab tingkat >= 3 (x.y.z, x.y.z.w, ...) digabung ke induk x.y, jadi induk dinilai dengan teks utuh.
# False: teks tingkat >= 3 dibuang sama sekali (induk hanya dinilai dari teksnya sendiri).
# Subbab tingkat >= 3 TIDAK PERNAH dinilai sebagai subbab sendiri, karena panduan tidak punya aturan untuk tingkat itu.
GABUNG_KE_INDUK = True


def merge_subsub(rows: list[dict]) -> list[dict]:
    """Subbab tingkat >= 3 (mis. 1.2.1, 3.5.2.1) tidak dinilai sendiri.
    Kalau induk x.y ada: teksnya digabung ke induk (atau dibuang, lihat GABUNG_KE_INDUK).
    Kalau induk x.y tidak terdeteksi: baris tetap dikembalikan dengan tanda "tingkat_tiga" supaya
    muncul di laporan sebagai tidak_dinilai, bukan sebagai subbab biasa."""
    out: list[dict] = []
    for r in rows:
        parts = str(r["subbab_num"]).split(".")
        if len(parts) >= 3:
            parent = ".".join(parts[:2])
            host = next((o for o in out if o["bab_num"] == r["bab_num"] and o["subbab_num"] == parent), None)
            if host is not None:
                if GABUNG_KE_INDUK:
                    host["raw_text"] += f"\n{r['subbab_num']} {r['subbab_title']}\n{r['raw_text']}"
                continue
            out.append(dict(r) | {"tingkat_tiga": True})
            continue
        out.append(dict(r))
    return out


def sectionize_doc(path: str) -> tuple[list[dict], int]:
    """Kembalikan (daftar subbab sebagai dict, jumlah kandidat heading yang ditolak)."""
    try:
        from student_sectionizer import sectionize_student_doc
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(
            f"Gagal mengimpor student_sectionizer: {e}\n"
            f"Pastikan student_sectionizer.py ada di {BASE / 'sectionizer'} dan parser.py di {BASE / 'rag'}."
        ) from e
    sections, rejected = sectionize_student_doc(path)
    return merge_subsub([asdict(s) for s in sections]), len(rejected)


def cover_text(path: str) -> str:
    ext = path.lower().rsplit(".", 1)[-1]
    if ext == "pdf":
        import fitz

        doc = fitz.open(path)
        text = doc[0].get_text("text")
        doc.close()
        return text
    import docx

    paras = [p.text for p in docx.Document(path).paragraphs if p.text.strip()]
    return "\n".join(paras[:40])


TITLE_SYSTEM = """Kamu mengambil JUDUL proyek akhir dari teks halaman sampul sebuah buku proyek akhir.
Ambil hanya judulnya, persis seperti tertulis. Jangan sertakan kata "PROYEK AKHIR", nama mahasiswa,
NRP, dosen pembimbing, program studi, departemen, atau nama institusi. Jika judul terpecah di beberapa
baris, gabungkan. Jika tidak ada judul yang jelas, isi string kosong.
Balas hanya JSON: {"judul": "<judul>"}"""


def extract_title(path: str) -> tuple[str, bool]:
    """(judul, valid). valid = judul benar-benar potongan dari teks sampul."""
    from rag_judge import ask

    cover = cover_text(path)
    out, _ = ask(TITLE_SYSTEM, "TEKS SAMPUL:\n" + cover[:1500])
    judul = " ".join((out.get("judul") or "").split())
    ok = bool(judul) and "".join(judul.split()).lower() in "".join(cover.split()).lower()
    return judul, ok
