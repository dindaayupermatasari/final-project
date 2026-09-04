"""
debug_line_spacing.py
======================
Script SEKALI PAKAI buat diagnosa akar masalah line_spacing di PDF asli
(bukan file test buatan Claude) -- jalankan ini di laptop Dinda, lalu
paste HASIL OUTPUT-nya (bukan PDF-nya) balik ke chat kalau gak nyaman
share dokumen aslinya.

CARA PAKAI
----------
1. pip install pymupdf (kalau belum ada)
2. Ganti PDF_PATH di bawah ke path PDF buku TA asli.
3. Ganti PAGE_NUMBER ke halaman yang ada paragraf yang KAMU TAHU PASTI
   di-set 1.5 spasi di Word aslinya (body text biasa, bukan caption/
   abstrak/daftar pustaka -- biar gak ke-filter zona header/footer).
4. Jalankan: python debug_line_spacing.py
5. Paste seluruh output-nya ke chat.

APA YANG DICEK
--------------
Untuk tiap blok teks di halaman itu: berapa baris di dalamnya (kalau
cuma 1 -- itu tandanya PyMuPDF MEMECAH paragraf, bukan salah kalibrasi),
dan kalau >=2 baris, berapa gap baseline-ke-baseline sesungguhnya vs
ukuran font -- ini angka MENTAH yang dipakai buat evaluasi apakah
konstanta 1.17 di pdf_extractor.py sudah pas atau perlu diganti.
"""

import statistics
import fitz  # pip install pymupdf

PDF_PATH = "GANTI_DENGAN_PATH_PDF_ASLI.pdf"
PAGE_NUMBER = 1  # 1-indexed, halaman yang ada paragraf 1.5-spasi yang kamu tahu pasti

HEADER_FOOTER_ZONE_PT = 70


def main():
    doc = fitz.open(PDF_PATH)
    page = doc[PAGE_NUMBER - 1]
    page_height = page.rect.height
    d = page.get_text("dict")

    for bi, b in enumerate(d["blocks"]):
        if b.get("type") != 0:
            continue
        bbox = b["bbox"]
        y0, y1 = bbox[1], bbox[3]
        zona = ""
        if y1 <= HEADER_FOOTER_ZONE_PT or y0 >= page_height - HEADER_FOOTER_ZONE_PT:
            zona = "  [DIBUANG -- masuk zona header/footer]"

        lines = b["lines"]
        text_preview = "".join(s.get("text", "") for l in lines for s in l["spans"])[
            :70
        ]

        print(f"--- Blok #{bi}{zona} ---")
        print(f"  teks: {text_preview!r}")
        print(f"  jumlah baris (PyMuPDF): {len(lines)}")

        if len(lines) < 2:
            print("  >> CUMA 1 BARIS -- kalau ini seharusnya paragraf multi-baris,")
            print("     berarti PyMuPDF MEMECAH paragraf ini jadi blok terpisah.")
            print()
            continue

        baselines = []
        sizes = []
        for line in lines:
            spans = line.get("spans") or []
            if spans:
                baselines.append(spans[0]["origin"][1])
            for sp in spans:
                if sp.get("size"):
                    sizes.append(round(sp["size"], 1))

        if len(baselines) < 2 or not sizes:
            print("  >> data tidak cukup buat dihitung.")
            print()
            continue

        avg_size = statistics.median(sizes)
        gaps = [baselines[i + 1] - baselines[i] for i in range(len(baselines) - 1)]
        avg_gap = statistics.median(gaps)
        raw_ratio = avg_gap / avg_size if avg_size else None
        old_formula = round(avg_gap / (avg_size * 1.17), 2) if avg_size else None

        print(f"  ukuran font (median): {avg_size}pt")
        print(f"  gap baseline antar baris (median): {round(avg_gap,2)}pt")
        print(f"  semua gap: {[round(g,2) for g in gaps]}")
        print(
            f"  rasio mentah gap/fontsize: {round(raw_ratio,3) if raw_ratio else None}"
        )
        print(f"  line_spacing hasil formula lama (/1.17): {old_formula}")
        print()


if __name__ == "__main__":
    main()
