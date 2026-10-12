"""
rtf_converter.py
=================
Konversi file .rtf (atau .doc) ke .docx via LibreOffice headless,
supaya bisa lanjut lewat extractor.py (jalur docx) apa adanya --
BUKAN dikonversi ke PDF (lihat hasil pengujian: jalur PDF merusak
font, alignment justify, dan indentasi kutipan; jalur DOCX jauh lebih
presisi karena tidak lewat heuristik posisi/bbox sama sekali).

Kenapa subprocess ke `soffice`, bukan library Python murni:
- Tidak ada library Python yang bisa BACA RTF dan LANGSUNG hasilkan
  objek setara python-docx Document dengan fidelity tinggi.
  `striprtf` cuma buang semua formatting (jadi plain text doang).
  `pyth` sudah tidak di-maintain dan tidak lengkap dukung RTF modern.
- LibreOffice sendiri yang sudah dipakai project ini untuk render PDF
  (lihat komentar "-Bold" quirk di pdf_extractor.py) -- jadi dependency
  ini bukan penambahan baru, cuma pemakaian ulang binary yang sama.
"""

import subprocess
import tempfile
from pathlib import Path


class ConversionError(RuntimeError):
    pass


def convert_to_docx(src_path: str, timeout: int = 60) -> str:
    """
    Konversi file .rtf ATAU .doc ke .docx pakai LibreOffice headless.
    Return path absolut ke file .docx hasil konversi (di folder
    temporary sendiri -- caller bertanggung jawab hapus kalau perlu).

    Raise ConversionError kalau soffice gagal / timeout / hasil docx
    tidak ditemukan.
    """
    src = Path(src_path).resolve()
    if not src.exists():
        raise ConversionError(f"File sumber tidak ditemukan: {src}")
    if src.suffix.lower() not in (".rtf", ".doc"):
        raise ConversionError(
            f"Ekstensi tidak didukung untuk konversi ini: {src.suffix} "
            "(hanya .rtf / .doc)"
        )

    out_dir = tempfile.mkdtemp(prefix="rtf2docx_")

    try:
        proc = subprocess.run(
            [
                "soffice",
                "--headless",
                "--norestore",
                "--convert-to",
                "docx:Office Open XML Text",
                "--outdir",
                out_dir,
                str(src),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as e:
        raise ConversionError(
            f"soffice timeout ({timeout}s) saat konversi {src.name}"
        ) from e
    except FileNotFoundError as e:
        raise ConversionError(
            "Binary 'soffice' (LibreOffice) tidak ditemukan di server. "
            "Pastikan LibreOffice ter-install di environment produksi."
        ) from e

    if proc.returncode != 0:
        raise ConversionError(
            f"soffice gagal (exit {proc.returncode}) untuk {src.name}: "
            f"{proc.stderr.strip() or proc.stdout.strip()}"
        )

    expected = Path(out_dir) / (src.stem + ".docx")
    if not expected.exists():
        raise ConversionError(
            f"Konversi 'sukses' tapi file .docx tidak ditemukan di "
            f"{expected} -- stdout: {proc.stdout.strip()}"
        )

    return str(expected)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python rtf_converter.py <file.rtf|file.doc>")
        sys.exit(1)

    result_path = convert_to_docx(sys.argv[1])
    print(f"OK -> {result_path}")
