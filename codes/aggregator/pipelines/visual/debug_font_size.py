"""
debug_font_size.py -- cetak ukuran font ASLI tiap span di PDF untuk baris yang
mengandung teks tertentu. Dipakai untuk memastikan apakah temuan "Ukuran font"
benar-benar berasal dari isi PDF atau dari cara program menghitungnya.

Pemakaian:
    python debug_font_size.py dokumen.pdf "Bogor, Yogyakarta"
    python debug_font_size.py dokumen.pdf "Bogor, Yogyakarta" --page 82
(--page = nomor halaman FISIK di PDF, bukan nomor tercetak di footer)
"""

import sys
import fitz


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return
    pdf_path, needle = sys.argv[1], sys.argv[2].lower()
    only_page = (
        int(sys.argv[sys.argv.index("--page") + 1]) if "--page" in sys.argv else None
    )
    with fitz.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf, start=1):
            if only_page and pno != only_page:
                continue
            for b in page.get_text("dict")["blocks"]:
                if b.get("type") != 0:
                    continue
                block_text = "".join(
                    sp["text"] for ln in b["lines"] for sp in ln["spans"]
                ).lower()
                if needle not in block_text:
                    continue
                print(
                    f"=== halaman fisik {pno}, bbox blok {[round(v, 1) for v in b['bbox']]}"
                )
                for li, ln in enumerate(b["lines"]):
                    print(f"  baris {li} bbox {[round(v, 1) for v in ln['bbox']]}")
                    for sp in ln["spans"]:
                        print(
                            f"    size={sp['size']:.2f}  font={sp['font']!r}  "
                            f"text={sp['text']!r}"
                        )


if __name__ == "__main__":
    main()
