"""
pdf_extractor.py
=================
TAHAP 2 (Extract Parameter) versi PDF-native.

KENAPA MODUL INI ADA
---------------------
extractor.py (versi lama) membaca parameter LANGSUNG dari XML docx
(w:pgMar, w:jc, style, run.bold, dst). Itu paling akurat, TAPI cuma
bisa jalan kalau file sumbernya .docx. Mayoritas mahasiswa upload PDF
(bukan docx), dan pembimbing eksplisit minta PDF diutamakan sebagai
sumber pengecekan -- jadi modul ini dibuat sebagai PENGGANTI/PIPELINE
UTAMA, bukan cuma pelengkap check_page_layout() seperti render PDF yang
sudah ada sebelumnya.

PRINSIP DESAIN (PALING PENTING)
---------------------------------
extract_all() di sini mengembalikan bentuk yang SAMA PERSIS dengan
extract_all() versi docx: {"paragraphs", "sections", "tables",
"mirror_margins"}, dan tiap paragraf tetap punya "runs": list of
{"text","font_name","font_size_pt","bold","italic"}.

Konsekuensinya: comparator.py, bibliography_checker.py,
foreign_term_checker.py, quotation_checker.py TIDAK PERLU diubah SAMA
SEKALI -- semua modul itu sudah generic, cuma konsumsi dict/list biasa,
tidak pernah pegang objek docx langsung. Yang beda HANYA cara isi dict
itu didapat.

Fungsi classify_paragraph(), is_titlecase(), dan semua REGEX heading/
caption/dst DIIMPOR ULANG dari extractor.py (bukan diduplikasi) --
logic itu murni berbasis STRING, sama sekali tidak menyentuh objek
docx, jadi 100% valid dipakai ulang di sini.

KETERBATASAN DIBANDING VERSI DOCX (baca sebelum pakai sebagai sumber
utama produksi -- ini bukan pengganti 1:1, tapi pendekatan terbaik yang
bisa didapat dari PDF)
----------------------------------------------------------------------
1. Heading via w:numPr/ilvl (Jalur 2 classify_paragraph, "bab yang
   nomornya digenerate murni oleh multilevel-list Word tanpa teks
   'BAB N' literal") TIDAK bisa dideteksi dari PDF -- info itu sama
   sekali tidak ada lagi begitu dokumen dirender jadi PDF. Dampaknya
   kecil karena mayoritas dokumen menulis "BAB N" literal.
2. is_unnumbered_heading() versi PDF (_is_unnumbered_heading_runs di
   bawah) pakai `runs` hasil ekstraksi sendiri (bukan para.runs docx) --
   secara prinsip setara (bold penuh + pendek + tanpa tanda baca akhir),
   bukan cuma didekati.
3. Rumus matematika OMML (m:oMath) tidak terdeteksi presisi -- OMML
   sudah "diratakan" jadi glyph biasa saat render ke PDF. Proxy yang
   dipakai SAMA seperti _is_math_run_text() di extractor.py: karakter
   Unicode matematis & nama font matematis (Cambria Math) -- ini tetap
   valid karena berbasis teks/font run, bukan struktur XML.
4. alignment, indentasi kiri/kanan, line_spacing, margin section semua
   DIDEKATI dari posisi (bbox) tiap span/blok teks terhadap tepi
   halaman -- BUKAN dibaca langsung dari atribut seperti di docx.
   Toleransi tiap fungsi dijelaskan di docstring masing-masing.
5. Satu paragraf PDF didekati sebagai SATU "block" hasil deteksi
   PyMuPDF (page.get_text("dict")["blocks"]). Ini heuristik bawaan
   PyMuPDF (dikelompokkan dari jarak antar baris) -- pada dokumen rapi
   umumnya sudah cukup akurat, tapi BUKAN jaminan 1:1 sama seperti
   docx.paragraphs (yang eksplisit dari w:p). Kasus tepi yang mungkin
   meleset: dua paragraf pendek berurutan tanpa jarak antar-baris
   ekstra bisa ke-gabung jadi satu "paragraf".
6. Section docx (tiap section bisa beda margin/orientasi) disederhanakan
   jadi SATU section untuk seluruh isi dokumen (margin diambil median
   dari semua halaman body, BUKAN per-section) -- dokumen dengan margin
   berbeda antar section (mis. lampiran landscape) tidak dibedakan.
"""

import re
import statistics

import fitz  # PyMuPDF

from aggregator.pipelines.visual.extractor import (
    CHAPTER_TITLE_RE,
    CHAPTER_TITLE_ONLY_RE,
    SPECIAL_FRONT_PAGE_MARKERS,
    NORMAL_FRONT_MATTER_MARKERS,
    BLANK_PAGE_NOTICE_RE,
    ABSTRACT_KEYWORD_RE,
    SOURCE_CITATION_RE,
    MATH_FONT_NAMES,
    _is_math_run_text,
    classify_paragraph,
    is_titlecase,
)

PT_TO_CM = 2.54 / 72.0

# Blok di zona ini (dari tepi atas/bawah halaman) dianggap header/footer
# -- BUKAN badan dokumen. 70pt konsisten dengan FOOTER_ZONE_PT yang
# sudah diverifikasi empiris di page_layout_checker.py.
HEADER_FOOTER_ZONE_PT = 70

# Blok pendek (nomor halaman, dsb) di zona footer -- pola longgar,
# angka biasa atau angka romawi saja.
_PAGE_NUM_ONLY_RE = re.compile(r"^[ivxlcdm\d]+$", re.IGNORECASE)

# Ciri baris entri Daftar Isi/Daftar Gambar/Daftar Tabel: diakhiri
# dot-leader (>=3 titik berurutan) + opsional nomor halaman di ujung.
# Ini SINYAL YANG TIDAK ADA di docx (dot-leader di sana cuma properti
# tab-stop Word, bukan karakter sungguhan di w:t) -- tapi begitu
# dokumen dirender jadi PDF, dot-leader itu IKUT TERBACA sebagai teks
# biasa. Sinyal ini dipakai buat membedakan entri TOC (mis. "BAB 1
# PENDAHULUAN ...................... 5") dari heading ASLI (mis. "BAB
# 1 PENDAHULUAN" tanpa embel-embel dot-leader), yang teksnya bisa
# identik persis dan SAMA-SAMA lolos classify_paragraph() sebagai
# "chapter_title" -- lihat penjelasan lengkap di blok penanganan
# in_toc/in_figure_table_list pada extract_paragraph_params_pdf().
_TOC_LEADER_RE = re.compile(r"\.{3,}\s*[ivxlcdm\d]*\s*$", re.IGNORECASE)


# --------------------------------------------------------------------
# Span-level helpers
# --------------------------------------------------------------------


_SUBSET_PREFIX_RE = re.compile(r"^[A-Z]{6}\+")
_FONT_SUFFIX_RE = re.compile(
    r"-?(BoldItalic|BoldItal|Bold|Italic|Oblique|Regular)$", re.IGNORECASE
)


def _normalize_font_name(name: str | None) -> str | None:
    """
    PDF (baik hasil export Word langsung maupun render LibreOffice)
    menyimpan nama font sebagai PostScript name internal, BUKAN nama
    family yang manusia baca -- mis. "Times New Roman" tersimpan
    sebagai "TimesNewRomanPSMT" (regular), "TimesNewRomanPS-BoldMT"
    (bold), "TimesNewRomanPS-ItalicMT" (italic). docx XML sebaliknya
    langsung simpan "Times New Roman" apa adanya.

    Tanpa normalisasi ini, RULES["font_name"] ("Times New Roman") di
    config.py TIDAK PERNAH match nama PostScript-nya -- comparator.py
    jadi salah tandai HAMPIR SEMUA paragraf sebagai "font tidak
    sesuai", padahal font-nya sudah benar. Ini bukan kasus tepi, tapi
    penyebab UTAMA skor jalur PDF jatuh drastis dibanding jalur docx.

    Strategi: lucuti awalan subset ("ABCDEF+..."), lucuti akhiran
    "MT"/style ("PS", "-Bold", "-Italic", "-BoldItalic", dst) satu per
    satu dari belakang, lalu pisahkan sisa nama camelCase jadi kata per
    kata ("TimesNewRoman" -> "Times New Roman"). Nama yang BUKAN pola
    PostScript umum (mis. "Calibri", "Arial" polos) tetap apa adanya --
    itu memang bukan Times New Roman, dan harus tetap ke-flag.
    """
    if not name:
        return name
    n = _SUBSET_PREFIX_RE.sub("", name)
    # Lucuti "MT" di akhir (penanda umum PostScript TrueType), lalu
    # lucuti akhiran gaya (Bold/Italic/dst, dengan/tanpa "-"), lalu
    # lucuti sisa "PS" penanda PostScript -- ulangi karena bisa
    # bertumpuk (mis. "PS-BoldMT" = PS + "-Bold" + "MT").
    for _ in range(3):
        n = re.sub(r"MT$", "", n)
        n = _FONT_SUFFIX_RE.sub("", n)
        n = re.sub(r"-?PS$", "", n)
    n = n.strip("-").strip()
    if n and " " not in n and re.search(r"[a-z][A-Z]", n):
        n = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", n)
    return n or name


def _span_bold(span: dict) -> bool:
    """
    Flag bit 2**4 di span["flags"] PyMuPDF = bold. Ditambah fallback ke
    nama font ("...-Bold...") karena sebagian PDF hasil LibreOffice
    tidak selalu set flag itu dengan benar untuk font tertentu.
    """
    if span.get("flags", 0) & (1 << 4):
        return True
    return "bold" in (span.get("font") or "").lower()


def _span_italic(span: dict) -> bool:
    """Simetris dengan _span_bold(), flag bit 2**1 = italic."""
    if span.get("flags", 0) & (1 << 1):
        return True
    fname = (span.get("font") or "").lower()
    return "italic" in fname or "oblique" in fname


def _span_is_math(span: dict) -> bool:
    return (span.get("font") or "") in MATH_FONT_NAMES or _is_math_run_text(
        span.get("text", "")
    )


# --------------------------------------------------------------------
# Paragraph (block) reconstruction per halaman
# --------------------------------------------------------------------


def _looks_like_page_number(text: str) -> bool:
    t = text.strip()
    return bool(t) and len(t) <= 6 and bool(_PAGE_NUM_ONLY_RE.match(t))


def _overlaps_table(
    bbox: tuple, table_bboxes: list[tuple], threshold: float = 0.5
) -> bool:
    """
    True kalau proporsi area `bbox` yang beririsan dengan salah satu
    `table_bboxes` >= threshold. Dipakai buat MEMBUANG blok teks yang
    sebenarnya isi sel tabel (bukan paragraf body) dari daftar paragraf
    -- tanpa ini, teks tabel ikut nyasar jadi "paragraf" body biasa,
    mengacaukan klasifikasi & index paragraf (yang dipakai tables_data
    buat mencari prev/next paragraf tetangga tabel).
    """
    bx0, by0, bx1, by1 = bbox
    area = max(bx1 - bx0, 0) * max(by1 - by0, 0)
    if area <= 0:
        return False
    for tx0, ty0, tx1, ty1 in table_bboxes:
        ix0, iy0 = max(bx0, tx0), max(by0, ty0)
        ix1, iy1 = min(bx1, tx1), min(by1, ty1)
        inter = max(ix1 - ix0, 0) * max(iy1 - iy0, 0)
        if inter / area >= threshold:
            return True
    return False


def _extract_blocks(
    page: "fitz.Page", table_bboxes: list[tuple] | None = None
) -> list[dict]:
    """
    Ambil semua blok teks badan halaman (header/footer/nomor-halaman/
    isi-tabel sudah difilter). Tiap blok balik: {"full_text", "runs",
    "bbox", "line_bboxes"}.

    "runs" dibangun dengan MENYISIPKAN run spasi sintetis di antar
    baris yang di-wrap -- supaya "".join(r["text"] for r in runs) ==
    full_text, PERSIS asumsi yang dipakai foreign_term_checker.py &
    quotation_checker.py saat merekonstruksi paragraf dari runs.
    """
    page_height = page.rect.height
    table_bboxes = table_bboxes or []
    raw = page.get_text("dict")["blocks"]
    blocks = []

    for b in raw:
        if b.get("type") != 0:  # 0 = teks, 1 = gambar
            continue
        bbox = b["bbox"]
        y0, y1 = bbox[1], bbox[3]
        # Buang blok yang SELURUHNYA ada di zona header/footer.
        if y1 <= HEADER_FOOTER_ZONE_PT or y0 >= page_height - HEADER_FOOTER_ZONE_PT:
            continue
        # Buang blok yang isinya sel tabel (lihat docstring _overlaps_table).
        if _overlaps_table(bbox, table_bboxes):
            continue

        # Gabung dulu "line" PyMuPDF yang baseline-nya nyaris identik
        # (selisih <=1pt) SEBELUM diproses jadi runs/line_bboxes.
        #
        # BUG (terverifikasi empiris via debug_line_spacing.py, lihat
        # temuan "Paragraf #429... Spasi baris 0.0"): pada baris yang
        # rata kiri-kanan (justify), PyMuPDF kadang memecah SATU baris
        # visual jadi BEBERAPA "line" internal di titik antar-kata --
        # spasi antar-kata yang melebar akibat justify membuat heuristik
        # clustering-nya salah baca celah itu sebagai baris baru.
        # Baseline semua fragmen itu IDENTIK (karena memang satu baris
        # yang sama), jadi kalau dibiarkan apa adanya, gap baseline
        # antar "baris" ini kehitung 0 -> line_spacing_infer() balik 0.0
        # padahal aslinya paragraf itu bisa saja 1.5 spasi.
        raw_lines = b["lines"]
        merged_pymupdf_lines = []
        for line in raw_lines:
            line_spans = line.get("spans") or []
            baseline = line_spans[0]["origin"][1] if line_spans else line["bbox"][3]
            if merged_pymupdf_lines:
                prev_baseline, prev_line = merged_pymupdf_lines[-1]
                if abs(baseline - prev_baseline) <= 1.0:
                    # Gabung ke line sebelumnya: sambung spans (urutan
                    # asli PyMuPDF sudah left-to-right dalam blok, jadi
                    # concat langsung aman) + union bbox.
                    prev_line["spans"] = prev_line["spans"] + line_spans
                    pb, lb = prev_line["bbox"], line["bbox"]
                    prev_line["bbox"] = (
                        min(pb[0], lb[0]),
                        min(pb[1], lb[1]),
                        max(pb[2], lb[2]),
                        max(pb[3], lb[3]),
                    )
                    continue
            merged_pymupdf_lines.append((baseline, dict(line)))
        pdf_lines = [ml[1] for ml in merged_pymupdf_lines]

        # BUG (baru ditemukan 2026-08-24): PyMuPDF kadang menyisipkan
        # "baris hantu" -- line dengan span yang teksnya kosong/spasi
        # saja -- terutama di halaman dengan banyak baris kosong sisa
        # (mis. halaman "DAFTAR TABEL" yang sebagian besar blank, atau
        # baris kosong bekas paragraf kosong hasil export Word->PDF).
        # Baris begini TIDAK menyumbang karakter apa pun ke full_text
        # (spans-nya toh di-skip saat membangun runs, lihat filter
        # `text == ""` di bawah), tapi kalau bbox-nya tetap dibiarkan
        # ikut serta di line_bboxes/group_bbox, union bbox blok jadi
        # jauh melenceng dari bbox teks yang BENERAN tercetak.
        # Efeknya fatal untuk _infer_alignment(): judul yang aslinya
        # rata tengah (mis. "DAFTAR GAMBAR", "DAFTAR TABEL") keliru
        # terbaca "rata kiri" karena bbox gabungannya "ditarik" sampai
        # ke posisi baris hantu yang kebetulan nempel margin kiri.
        # Verifikasi empiris: baris "DAFTAR GAMBAR" (14pt bold, benar2
        # center secara visual) punya line_bboxes asli [(242.0, ..,
        # 371.1, ..), (85.1, .., 88.1, ..)] -- baris kedua itu 3pt
        # doang lebarnya dan spans-nya kosong, tapi cukup buat menarik
        # x0 gabungan dari 242.0 turun ke 85.1, mengubah titik tengah
        # blok dari ~306pt (SANGAT dekat titik tengah kolom, harusnya
        # lolos sebagai "center") jadi ~228pt (jauh dari titik tengah,
        # kebaca "left"). Fix: buang baris yang sama sekali tidak
        # punya teks tercetak SEBELUM dipakai di mana pun (bbox atau
        # full_text) -- baris begini tidak pernah dibutuhkan untuk
        # keduanya.
        pdf_lines = [
            line
            for line in pdf_lines
            if "".join(s.get("text", "") for s in (line.get("spans") or [])).strip()
        ]

        # BUG DIPERBAIKI (jalur PDF -- "Ukuran font 12.0pt tidak sesuai
        # untuk tipe 'caption', seharusnya 11pt" padahal captionnya
        # sendiri 100% 11pt): PyMuPDF sering menggabungkan baris caption
        # gambar ("Gambar N.N ...") dengan baris atribusi "Sumber: ..."
        # tepat di bawahnya jadi SATU block PyMuPDF, karena jaraknya
        # rapat (cuma selisih satu baris single-spacing). Padahal
        # keduanya jenis paragraf yang beda (caption vs source_citation)
        # dengan aturan font BERBEDA (11pt vs 12pt) -- kalau dibiarkan
        # jadi satu elemen gabungan, size_char_weight di comparator.py
        # milih ukuran MAYORITAS karakter, dan teks "Sumber: <url>"
        # (seringnya lebih panjang dari teks caption-nya sendiri karena
        # URL panjang) menang jadi "dominant_size" -- caption yang
        # aslinya 100% 11pt pun ke-flag "12pt tidak sesuai". Fix: pecah
        # blok PyMuPDF di titik mana pun baris berikutnya diawali
        # "Sumber:"/"Sumber ." (SOURCE_CITATION_RE) -- baris itu SELALU
        # memulai paragraf/elemen baru sendiri, tidak peduli seberapa
        # rapat jaraknya dengan baris di atasnya.
        split_pdf_line_groups = [[]]
        for line in pdf_lines:
            line_spans = line.get("spans") or []
            line_text = "".join(s.get("text", "") for s in line_spans).strip()
            if (
                line_text
                and SOURCE_CITATION_RE.match(line_text)
                and split_pdf_line_groups[-1]
            ):
                split_pdf_line_groups.append([])
            split_pdf_line_groups[-1].append(line)

        for pdf_lines_group in split_pdf_line_groups:
            if not pdf_lines_group:
                continue
            runs = []
            line_bboxes = []
            line_baselines = []
            for li, line in enumerate(pdf_lines_group):
                line_bboxes.append(line["bbox"])
                # Baseline ASLI baris (koordinat "origin" span pertama), BUKAN
                # top-of-bbox -- lihat docstring _infer_line_spacing() untuk
                # alasan kenapa ini penting.
                line_spans = line.get("spans") or []
                line_baselines.append(
                    line_spans[0]["origin"][1] if line_spans else line["bbox"][3]
                )
                for span in line["spans"]:
                    text = span.get("text", "")
                    if text == "":
                        continue
                    runs.append(
                        {
                            "text": text,
                            "font_name": _normalize_font_name(span.get("font")),
                            "font_size_pt": round(span.get("size", 0.0), 1),
                            "bold": _span_bold(span),
                            "italic": _span_italic(span),
                            "is_math": _span_is_math(span),
                        }
                    )
                if li < len(pdf_lines_group) - 1:
                    # Sambungan antar baris yang di-wrap -- spasi sintetis,
                    # format-nya ikut run TERAKHIR baris ini (paling wajar
                    # untuk keperluan italic_map di checker downstream).
                    inherit = runs[-1] if runs else None
                    runs.append(
                        {
                            "text": " ",
                            "font_name": inherit["font_name"] if inherit else None,
                            "font_size_pt": (
                                inherit["font_size_pt"] if inherit else None
                            ),
                            "bold": inherit["bold"] if inherit else False,
                            "italic": inherit["italic"] if inherit else False,
                            "is_math": False,
                        }
                    )

            full_text = "".join(r["text"] for r in runs)
            if not full_text.strip():
                continue
            # Nomor halaman badan yang nyasar ke luar zona footer (jarang,
            # tapi ada dokumen dengan margin bawah kecil) -- buang juga.
            if _looks_like_page_number(full_text) and len(pdf_lines_group) == 1:
                continue

            group_bbox = (
                min(l["bbox"][0] for l in pdf_lines_group),
                min(l["bbox"][1] for l in pdf_lines_group),
                max(l["bbox"][2] for l in pdf_lines_group),
                max(l["bbox"][3] for l in pdf_lines_group),
            )
            blocks.append(
                {
                    "full_text": full_text,
                    "runs": runs,
                    "bbox": group_bbox,
                    "line_bboxes": line_bboxes,
                    "line_baselines": line_baselines,
                }
            )

    return blocks


_SENTENCE_END_RE = re.compile(r"""["')\]]*[.!?:;]["')\]]*$""")


def _merge_wrapped_paragraph_blocks(
    blocks: list[dict], left_margin_pt: float, right_margin_pt: float
) -> list[dict]:
    """
    Gabung kembali blok-blok satu-baris (atau multi-baris pendek) yang
    SEBENARNYA satu paragraf yang sama, tapi kepecah jadi blok terpisah
    oleh heuristik clustering internal PyMuPDF -- kejadian PALING sering
    di paragraf ber-spasi lebar (>=1.5x), terverifikasi lewat
    debug_line_spacing.py di file hasil konversi LibreOffice.

    Tanpa fungsi ini, dampaknya BUKAN cuma line_spacing balik None
    (blok 1-baris tidak ada gap buat dihitung) -- font/alignment/indent
    per paragraf ikut salah karena satu paragraf asli dihitung sebagai
    banyak paragraf kecil terpisah, dan index paragraf (dipakai
    comparator.py buat referensi lokasi temuan) jadi tidak match jumlah
    paragraf yang sebenarnya.

    Syarat gabung blocks[i] -> blocks[i+1] (SEMUA harus terpenuhi):
      1. Baris terakhir blocks[i] TIDAK diakhiri tanda baca penutup
         kalimat (. ! ? : ; -- boleh diikuti kutip/kurung tutup).
         Paragraf yang beneran sudah selesai hampir selalu diakhiri
         tanda baca ini; kalau tidak, kalimatnya kemungkinan besar
         masih berlanjut ke blok berikutnya.
      2. Baris terakhir blocks[i] dan baris pertama blocks[i+1] sama-
         sama rata ke left_margin_pt (toleransi noise render). Format
         PENS tidak pakai indentasi first-line, jadi baris lanjutan
         paragraf yang di-wrap SELALU rata margin kiri kolom -- sama
         kayak baris pertama paragraf body biasa (makanya syarat ini
         SENDIRIAN tidak cukup, harus dibarengi #1 dan #3).
      3. Gap baseline (baris pertama blocks[i+1] - baris terakhir
         blocks[i]) masuk akal sebagai jarak antar-baris NORMAL
         (<= ~2.1x baseline single-line-height font tsb), bukan gap
         besar yang menandakan baris kosong/spasi-sebelum paragraf
         baru beneran.
      4. Ukuran font median kedua blok sama (toleransi 0.5pt) --
         paragraf lanjutan hampir pasti font sama; ganti ukuran font
         biasanya menandakan elemen lain (caption, heading, dst).

    CATATAN PENTING: ini heuristik yang BELUM diverifikasi terhadap PDF
    asli mahasiswa (baru diuji lewat simulasi/mock), beda dari fix
    baseline/1.17 yang sudah diverifikasi empiris. Kalau nanti dicoba
    di dokumen asli dan ada gejala:
      - Paragraf body ke-gabung sama heading/caption di sebelahnya
        (over-merge) -> longgarkan/perketat syarat #4, atau tambah cek
        bold sebagai syarat kelima.
      - Paragraf panjang yang seharusnya satu tetap kepecah
        (under-merge) -> gap aslinya di luar rentang 2.1x yang
        diasumsikan di syarat #3, perlu angka gap mentah dari
        debug_line_spacing.py di paragraf yang kepecah itu.
    """
    if len(blocks) < 2:
        return blocks

    tol = 0.25 / PT_TO_CM  # ~7pt, toleransi rata margin kiri
    merged: list[dict] = []
    for block in blocks:
        if merged:
            prev = merged[-1]
            prev_text = prev["full_text"].rstrip()
            prev_last_line = prev["line_bboxes"][-1]
            cur_first_line = block["line_bboxes"][0]
            prev_last_baseline = prev["line_baselines"][-1]
            cur_first_baseline = block["line_baselines"][0]

            prev_sizes = [r["font_size_pt"] for r in prev["runs"] if r["font_size_pt"]]
            cur_sizes = [r["font_size_pt"] for r in block["runs"] if r["font_size_pt"]]
            font_size = statistics.median(prev_sizes) if prev_sizes else None
            same_size = bool(
                prev_sizes
                and cur_sizes
                and font_size is not None
                and abs(font_size - statistics.median(cur_sizes)) <= 0.5
            )

            not_sentence_end = bool(prev_text) and not _SENTENCE_END_RE.search(
                prev_text
            )
            # FIX (baru ditemukan 2026-08-24, kasus nyata: paragraf
            # "...sebagaimana ditunjukkan pada" [baris terakhir] gagal
            # gabung dengan "Gambar 3.10." [baris berikutnya] walau
            # keduanya SAMA-SAMA rata pada x=70.9pt -- PERSIS sama satu
            # sama lain, cuma beda dari left_margin_pt GLOBAL (85.1pt,
            # median seluruh dokumen). Sebagian halaman body text
            # dokumen ini ternyata rata pada margin kiri yang sedikit
            # berbeda dari median global (indikasi indentasi/margin
            # tidak 100% seragam antar bagian dokumen) -- versi lama
            # yang mensyaratkan KEDUA baris cocok ke left_margin_pt
            # GLOBAL gagal total di kasus begini, blok "Gambar 3.10."
            # (fragmen akhir kalimat, BUKAN caption sungguhan) jadi
            # tidak pernah ke-merge dan lolos ke classify_paragraph()
            # sebagai ptype "caption" murni karena teksnya kebetulan
            # cocok pola CAPTION_RE. Fix: syarat rata-kiri cukup
            # dicek antara DUA BARIS ITU SENDIRI (harus sama-sama
            # nempel di posisi x yang SAMA), bukan terhadap referensi
            # eksternal yang bisa saja tidak match margin lokal
            # halaman itu.
            left_ok = abs(prev_last_line[0] - cur_first_line[0]) <= tol
            gap = cur_first_baseline - prev_last_baseline
            gap_ok = bool(font_size) and 0 < gap <= font_size * 1.17 * 2.1

            if not_sentence_end and left_ok and gap_ok and same_size:
                inherit = prev["runs"][-1] if prev["runs"] else None
                prev["runs"].append(
                    {
                        "text": " ",
                        "font_name": inherit["font_name"] if inherit else None,
                        "font_size_pt": inherit["font_size_pt"] if inherit else None,
                        "bold": inherit["bold"] if inherit else False,
                        "italic": inherit["italic"] if inherit else False,
                        "is_math": False,
                    }
                )
                prev["runs"].extend(block["runs"])
                prev["full_text"] = prev["full_text"] + " " + block["full_text"]
                pb, cb = prev["bbox"], block["bbox"]
                prev["bbox"] = (
                    min(pb[0], cb[0]),
                    min(pb[1], cb[1]),
                    max(pb[2], cb[2]),
                    max(pb[3], cb[3]),
                )
                prev["line_bboxes"].extend(block["line_bboxes"])
                prev["line_baselines"].extend(block["line_baselines"])
                continue

        merged.append(
            {
                "full_text": block["full_text"],
                "runs": list(block["runs"]),
                "bbox": block["bbox"],
                "line_bboxes": list(block["line_bboxes"]),
                "line_baselines": list(block["line_baselines"]),
            }
        )
    return merged


# --------------------------------------------------------------------
# Heuristik pengganti atribut docx (alignment, indent, dsb)
# --------------------------------------------------------------------


def _infer_alignment(
    block: dict, left_margin_pt: float, right_margin_pt: float
) -> str | None:
    """
    Dekati w:jc dari posisi baris:
      - center : titik tengah bbox blok ~ titik tengah kolom teks
      - right  : sisi kanan blok nempel batas kanan, sisi kiri menjorok
      - both (justify) : SEMUA baris (kecuali baris terakhir) nyaris
        menyentuh batas kiri MAUPUN kanan kolom
      - left   : default kalau tidak match pola di atas
    Toleransi 0.45 cm (dikonversi ke pt) -- dilonggarkan dari 0.3cm
    (2026-08-23) supaya cek "center" tetap tahan terhadap noise residual
    di left_margin_pt/right_margin_pt (nilai itu ESTIMASI per-dokumen
    dari extract_section_params_pdf(), bukan dibaca presis seperti w:jc
    di docx) -- tanpa longgar secukupnya, judul yang SUDAH center di
    dokumen asli masih bisa salah ditandai "belum rata tengah" kalau
    estimasi margin dokumen itu meleset beberapa milimeter dari nilai
    sebenarnya. 0.45cm masih cukup ketat untuk tetap menangkap judul
    yang BENERAN rata kiri/kanan (yang titik tengahnya biasanya bergeser
    >1cm dari titik tengah kolom).
    """
    tol = 0.45 / PT_TO_CM  # ~12.75pt
    col_left, col_right = left_margin_pt, right_margin_pt
    col_center = (col_left + col_right) / 2
    bbox = block["bbox"]
    x0, x1 = bbox[0], bbox[2]
    center = (x0 + x1) / 2

    # NOTE (fixed): versi sebelumnya juga mensyaratkan (col_right -
    # col_left) - (x1 - x0) > tol, dengan maksud "blok tidak boleh
    # selebar kolom" biar gak salah anggap paragraf body rata-kiri
    # sebagai center. Tapi ini gagal untuk baris CENTER yang kebetulan
    # hampir selebar kolom (kasus nyata: baris "Sumber: ..." di bawah
    # gambar/tabel) -- baris itu pas di tengah kolom TAPI lebarnya
    # nyaris = lebar kolom, jadi syarat kedua ini menolaknya walau
    # syarat pertama (titik tengah blok == titik tengah kolom) sudah
    # cukup buat bedain dari rata-kiri (yang titik tengahnya BAKAL
    # geser ke kiri kalau baris itu tidak selebar kolom). Titik tengah
    # sendirian sudah cukup diskriminatif; syarat kedua dibuang.
    if abs(center - col_center) <= tol:
        return "center"

    if abs(x1 - col_right) <= tol and (x0 - col_left) > tol:
        return "right"

    lines = block["line_bboxes"]
    if len(lines) >= 2:
        body_lines = lines[:-1]  # baris terakhir paragraf wajar tidak penuh
        left_ok = all(abs(l[0] - col_left) <= tol for l in body_lines)
        right_ok = all(abs(l[2] - col_right) <= tol for l in body_lines)
        if left_ok and right_ok:
            return "both"

    return "left"


def _infer_indent_cm(
    block: dict, left_margin_pt: float, right_margin_pt: float
) -> tuple[float | None, float | None]:
    """
    Indentasi kiri/kanan paragraf dalam cm, didekati dari selisih bbox
    blok terhadap batas kolom teks. None-kan (bukan 0.0) kalau selisih
    di bawah noise floor (~0.1cm) -- konsisten dengan semantik docx
    "tidak di-set" vs "di-set 0".
    """
    noise_floor_pt = 0.1 / PT_TO_CM
    bbox = block["bbox"]
    left_delta = bbox[0] - left_margin_pt
    right_delta = right_margin_pt - bbox[2]
    left_cm = left_delta * PT_TO_CM if left_delta > noise_floor_pt else None
    right_cm = right_delta * PT_TO_CM if right_delta > noise_floor_pt else None
    return left_cm, right_cm


def _infer_line_spacing(block: dict) -> float | None:
    """
    Rasio jarak-antar-baseline terhadap ukuran font. Cuma valid kalau
    blok >= 2 baris -- balik None kalau cuma 1 baris (tidak ada gap
    untuk dihitung, BUKAN berarti spacing-nya salah).

    PENTING (bug yang diperbaiki 2026-08-19): versi sebelumnya memakai
    top-of-bbox tiap baris (line["bbox"][1]) buat hitung gap. Itu BUKAN
    ukuran yang stabil -- tinggi bbox satu baris ikut naik-turun
    tergantung glyph yang ADA di baris itu (baris dengan huruf berekor
    seperti g/y/p naik sedikit lebih rendah, baris dengan aksen/huruf
    kapital semua naik lebih tinggi), jadi gap top-to-top antar baris
    bisa jauh meleset dari jarak baris sebenarnya (diuji empiris:
    paragraf 1.5 spasi yang SAMA bisa kebaca rasio 0.6x-1.45x tergantung
    isi barisnya). Baseline (line["spans"][0]["origin"][1] dari
    PyMuPDF) adalah titik jangkar vertikal yang presisi -- konsisten
    berapa pun bentuk glyph di baris tsb -- jadi dipakai di sini,
    bukan bbox.

    KALIBRASI KONSTANTA (diperbaiki lagi 2026-08-19, setelah fix merge
    paragraf di atas bikin bug ini KELIHATAN): konstanta 1.17 yang
    dipakai sebelumnya TERBUKTI sedikit kebesaran -- setelah paragraf
    yang tadinya kepecah jadi blok satu-baris berhasil digabung
    (_merge_wrapped_paragraph_blocks), puluhan paragraf yang jelas-
    jelas 1.5 spasi di dokumen asli malah konsisten kebaca 1.47-1.48,
    BUKAN acak/noise tapi BIAS SISTEMATIS ke satu arah. Dihitung balik
    dari situ (gap/size = hasil_lama * 1.17): rasio mentah gap/size
    ~1.72-1.73, dan 1.72/1.5 ~ 1.15 -- bukan 1.17. Konstanta diganti
    jadi 1.15.

    TOLERANSI (masalah kedua, terpisah dari kalibrasi): comparator.py
    membandingkan line_spacing ke RULES["line_spacing"] dengan toleransi
    CUMA 0.01 (abs(x - 1.5) > 0.01 -> gagal). Toleransi seketat itu
    masuk akal buat jalur DOCX (nilai persis dari XML, bukan estimasi),
    tapi buat jalur PDF nilai ini SELALU estimasi dari metrik render
    font -- noise pt-level dari hinting/DPI bisa geser hasil 0.01-0.03
    walau konstanta sudah benar. Daripada melonggarkan toleransi
    comparator.py (yang jadi kelonggaran untuk docx juga, padahal disitu
    memang harus presisi), noise itu diserap DI SINI: kalau rasio mentah
    dekat (toleransi 0.05) ke salah satu nilai lazim di Word (1.0, 1.15,
    2.0, 2.5, 3.0, dan 1.5 -- preset "Single/1.5 lines/Double" dst),
    di-snap PAS ke nilai itu. Paragraf yang beneran salah spasi (mis.
    hasil mentah 0.98x atau 1.3x, di luar jangkauan 0.05 dari semua
    preset) TETAP dilaporkan apa adanya, tidak ke-snap ke 1.5 -- jadi
    ini bukan melonggarkan deteksi, cuma membersihkan noise di sekitar
    nilai yang MEMANG dimaksud.

    CATATAN UNTUK VERIFIKASI LANJUTAN: konstanta 1.15 ini didapat dari
    reverse-engineering nilai 1.47/1.48 yang muncul di laporan real
    (bukan dari pengukuran font-metric langsung) -- cukup solid karena
    biasnya sistematis & sempit (1.47-1.48, bukan tersebar random), tapi
    kalau nanti dicoba ke PDF lain dan masih ada bias residual, cek dulu
    apa font_size yang kepakai (Arial vs TNR bisa beda dikit metriknya)
    sebelum utak-atik konstanta lagi.
    """
    baselines = block.get("line_baselines") or [
        bbox[1] for bbox in block["line_bboxes"]
    ]
    if len(baselines) < 2:
        return None
    sizes = [r["font_size_pt"] for r in block["runs"] if r["font_size_pt"]]
    if not sizes:
        return None
    avg_size = statistics.median(sizes)
    if avg_size <= 0:
        return None
    gaps = [baselines[i + 1] - baselines[i] for i in range(len(baselines) - 1)]
    avg_gap = statistics.median(gaps)
    # Baseline-to-baseline "single spacing" wajar ~1.15x font size
    # (line-height bawaan font, bukan tepat 1.0x) -- dibagi konstanta
    # ini supaya hasilnya kalibrasi ke skala "1.0 = single" yang sama
    # dengan RULES["line_spacing"] (1.5) di config.py.
    raw_ratio = avg_gap / (avg_size * 1.15)
    for preset in _LINE_SPACING_PRESETS:
        if abs(raw_ratio - preset) <= _LINE_SPACING_SNAP_TOL:
            return preset
    return round(raw_ratio, 2)


# Preset spasi baris lazim di Word (Single/1.5 lines/Double, dst) --
# lihat docstring _infer_line_spacing() bagian TOLERANSI.
_LINE_SPACING_PRESETS = (1.0, 1.15, 1.5, 2.0, 2.5, 3.0)
_LINE_SPACING_SNAP_TOL = 0.05


def _is_unnumbered_heading_runs(runs: list[dict], text: str) -> bool:
    """
    Padanan is_unnumbered_heading(para) di extractor.py, tapi berbasis
    `runs` (list dict generik) alih-alih para.runs docx -- prinsipnya
    identik: seluruh run berisi teks harus bold, teks pendek (<=8 kata),
    dan tidak diakhiri tanda baca kalimat.
    """
    text_runs = [r for r in runs if r["text"].strip()]
    if not text_runs:
        return False
    all_bold = all(r["bold"] for r in text_runs)
    is_short = len(text.split()) <= 8
    no_sentence_ending = not text.rstrip().endswith((".", ",", ";"))
    return all_bold and is_short and no_sentence_ending


# --------------------------------------------------------------------
# Section params (margin & ukuran halaman) dari geometri PDF
# --------------------------------------------------------------------


def extract_section_params_pdf(pdf_path: str) -> list[dict]:
    """
    Padanan extract_section_params(doc) versi docx, TAPI disederhanakan
    jadi SATU section untuk seluruh dokumen (lihat keterbatasan #6 di
    docstring modul). page_width/height diambil LANGSUNG dari
    page.rect -- ini EXACT (bukan inferensi) karena PDF menyimpan
    ukuran halaman literal, malah lebih presisi dari cm-rounding docx.

    BUG (diperbaiki 2026-08-19): versi sebelumnya menghitung margin
    kiri/kanan dari min(bbox[0])/max(bbox[2]) SEMUA blok di tiap
    halaman (lalu median antar-halaman) -- ini TERMASUK judul center
    (DAFTAR ISI, BAB N JUDUL), cover, caption, "Sumber:", dst yang
    MEMANG tidak menyentuh margin kiri/kanan asli (center/indent
    sengaja). Kalau di suatu halaman kebetulan tidak ada body text yang
    mepet margin asli (mis. halaman isinya cuma judul bab + gambar),
    margin halaman itu kebaca ngaco dan menarik median keseluruhan
    dokumen jadi salah. Dampaknya BUKAN cuma temuan margin sendiri
    (skor "Umum (Margin Halaman)" 0%) -- left_margin_pt/right_margin_pt
    yang salah ini dipakai juga oleh _infer_alignment() buat menghitung
    titik tengah kolom, jadi JUDUL YANG ASLINYA SUDAH CENTER (mis.
    "DAFTAR ISI", "BAB 1 PENDAHULUAN") ikut salah dilaporkan "seharusnya
    rata tengah" -- terverifikasi persis pola ini di laporan nyata,
    findings-nya SELALU 3 rangkap (font+bold+center) utk semua
    chapter_title, termasuk yang kemungkinan besar sudah benar di
    dokumen asli.

    Fix: filter kandidat margin ke paragraf BODY multi-baris SAJA
    (>=2 baris) -- heading/caption/cover/"Sumber:" hampir selalu cuma
    1 baris, jadi filter ini otomatis menyingkirkan mereka TANPA perlu
    klasifikasi penuh (classify_paragraph baru bisa jalan SETELAH
    margin section ada -- circular dependency kalau mau pakai
    ptype=="body" beneran, makanya dipakai proxy struktural "multi-
    baris" di sini). Konvensi PENS tidak pakai indentasi first-line,
    jadi SEMUA baris paragraf body (termasuk baris pertama) rata ke
    margin kiri -- dipakai buat estimasi kiri. Margin kanan dari SEMUA
    baris KECUALI baris terakhir tiap paragraf (baris terakhir wajar
    tidak penuh/tidak mepet margin kanan).

    BUG LANJUTAN (diperbaiki 2026-08-23, sebelumnya tercatat sebagai
    "margin kanan overcorrection" yang belum tuntas): fungsi ini
    memanggil `_extract_blocks(page)` MENTAH -- TANPA lewat
    `_merge_wrapped_paragraph_blocks()` -- padahal PyMuPDF sering
    memecah SATU paragraf body ber-spasi 1.5 jadi BEBERAPA blok
    terpisah (masing-masing seringnya cuma 1 baris) akibat heuristik
    clustering internalnya sendiri (root cause yang SAMA dengan yang
    dijelaskan panjang lebar di docstring `_merge_wrapped_paragraph_
    blocks`). Filter ">=2 baris" di atas otomatis MEMBUANG mayoritas
    blok pecahan ini (karena tiap pecahannya cuma 1 baris), sehingga
    kandidat margin kanan yang TERSISA didominasi sisa blok pendek/tak
    representatif (mis. potongan baris terakhir paragraf, sub-blok
    hasil pecahan yang kebetulan 2 baris tapi baris terakhirnya
    pendek) -- inilah kenapa median-nya bisa melenceng jauh (mis.
    terbaca ~14.9cm padahal margin asli ~3cm).

    Fix: jalankan `_merge_wrapped_paragraph_blocks()` DULU per halaman
    sebelum memfilter ">=2 baris", supaya paragraf body yang aslinya
    banyak baris tapi kepecah oleh PyMuPDF tergabung kembali dulu.
    Karena fungsi merge itu sendiri butuh estimasi margin kiri/kanan
    (buat cek "rata ke margin kiri" di syarat #2-nya) -- circular
    dependency yang sama seperti klasifikasi paragraf -- dipakai
    pendekatan DUA TAHAP:
      1. Kumpulkan estimasi KASAR (dari SEMUA blok mentah, tanpa
         filter jumlah baris) hanya untuk menyeed proses merge.
         Blok satu-baris (heading dst) tetap ikut di tahap ini, tapi
         karena cuma dipakai sebagai TITIK ACUAN merge (bukan hasil
         akhir), noise dari situ tidak fatal -- merge cuma butuh
         margin kiri yang "cukup dekat", bukan presisi.
      2. Jalankan merge per halaman pakai estimasi kasar itu, lalu
         hitung margin PRESISI dari blok hasil merge yang >=2 baris
         (logika sama seperti sebelumnya, cuma sumber bloknya beda).
    """
    with fitz.open(pdf_path) as pdf:
        if pdf.page_count == 0:
            raise ValueError("PDF tidak punya halaman sama sekali.")

        first_page = pdf[0]
        width_pt, height_pt = first_page.rect.width, first_page.rect.height

        # --- Tahap 1: estimasi KASAR margin kiri/kanan, cuma buat
        #     menyeed _merge_wrapped_paragraph_blocks() (lihat catatan
        #     di docstring). Dikumpulkan dari SEMUA blok mentah di
        #     SELURUH halaman (bukan per-halaman), lalu diambil MEDIAN --
        #     bukan min() seperti versi sebelumnya. min() per halaman
        #     rapuh terhadap SATU blok outlier (mis. tabel/gambar lebar
        #     penuh yang bbox-nya menclok sampai nyaris tepi fisik
        #     halaman) -- satu outlier begini cukup buat menarik seluruh
        #     estimasi kasar meleset jauh dari margin sebenarnya, padahal
        #     tujuannya cuma jadi "kira-kira" buat menyeed merge, jadi
        #     harus tahan-outlier. Median atas SEMUA blok (bukan cuma
        #     yang >=2 baris) tetap valid buat tujuan ini karena mayoritas
        #     blok di dokumen memang teks body yang rata margin kiri/
        #     kanan -- minoritas blok center/indent tidak cukup untuk
        #     menggeser median.
        rough_pages_blocks: list[list[dict]] = []
        rough_left: list[float] = []
        rough_right: list[float] = []
        for page in pdf:
            blocks = _extract_blocks(page)
            rough_pages_blocks.append(blocks)
            for b in blocks:
                rough_left.append(b["bbox"][0])
                rough_right.append(page.rect.width - b["bbox"][2])
        rough_left_pt = statistics.median(rough_left) if rough_left else 0.0
        rough_right_pt = statistics.median(rough_right) if rough_right else 0.0

        # --- Tahap 2: gabungkan blok yang terpecah pakai estimasi
        #     kasar di atas, baru filter ">=2 baris" & hitung margin
        #     presisi dari hasil gabungan itu.
        left_candidates: list[float] = []
        right_candidates: list[float] = []
        tops, bottoms = [], []
        for page, raw_blocks in zip(pdf, rough_pages_blocks):
            if not raw_blocks:
                continue
            blocks = _merge_wrapped_paragraph_blocks(
                raw_blocks, rough_left_pt, rough_right_pt
            )
            for b in blocks:
                lines = b["line_bboxes"]
                if len(lines) < 2:
                    # Blok satu-baris (heading/caption/cover/"Sumber:")
                    # sering SENGAJA tidak mepet margin -- disingkirkan
                    # dari kandidat, lihat docstring di atas.
                    continue
                for l in lines:
                    left_candidates.append(l[0])
                for l in lines[:-1]:
                    right_candidates.append(page.rect.width - l[2])
            tops.append(min(b["bbox"][1] for b in blocks))
            bottoms.append(page.rect.height - max(b["bbox"][3] for b in blocks))

        if not left_candidates:
            # Fallback kalau dokumen aneh (tidak ada body multi-baris
            # sama sekali, bahkan setelah merge) -- balik ke pendekatan
            # paling kasar biar tidak crash.
            left_candidates, right_candidates = [], []
            for blocks in rough_pages_blocks:
                if not blocks:
                    continue
                left_candidates.append(min(b["bbox"][0] for b in blocks))
                right_candidates.append(min(width_pt - b["bbox"][2] for b in blocks))
        if not left_candidates:
            raise ValueError("Tidak ada teks badan terdeteksi di PDF ini sama sekali.")

        margin_left_pt = statistics.median(left_candidates)
        margin_right_pt = (
            statistics.median(right_candidates) if right_candidates else margin_left_pt
        )
        margin_top_pt = statistics.median(tops)
        margin_bottom_pt = statistics.median(bottoms)

    def twips(pt: float) -> int:
        return round(pt * 20)  # 1pt = 20 twips

    return [
        {
            "index": 0,
            "left_cm": margin_left_pt * PT_TO_CM,
            "left_twips": twips(margin_left_pt),
            "top_cm": margin_top_pt * PT_TO_CM,
            "top_twips": twips(margin_top_pt),
            "bottom_cm": margin_bottom_pt * PT_TO_CM,
            "bottom_twips": twips(margin_bottom_pt),
            "right_cm": margin_right_pt * PT_TO_CM,
            "right_twips": twips(margin_right_pt),
            "width_cm": width_pt * PT_TO_CM,
            "width_twips": twips(width_pt),
            "height_cm": height_pt * PT_TO_CM,
            "height_twips": twips(height_pt),
            # pgnum_fmt/start & footer_* TIDAK diisi di sini -- itu
            # sudah dihitung lebih akurat oleh extract_footer_numbers()
            # (baca angka BENERAN tercetak) di page_layout_checker.py,
            # yang dipanggil terpisah di main.py. Diisi None supaya
            # skema dict tetap sama, bukan dihitung dobel/tidak akurat.
            "pgnum_fmt": None,
            "pgnum_start": None,
            "footer_has_page_field": None,
            "footer_alignment": None,
        }
    ]


def _infer_mirror_margins(pdf_path: str) -> bool:
    """
    Padanan _get_mirror_margins(doc). Didekati dari POLA: margin kiri
    halaman ganjil vs genap berbeda signifikan (>0.3cm) DAN konsisten
    berselang-seling -- itu ciri cetak bolak-balik/mirror margin.
    Kalau dokumen cuma sedikit halaman atau tidak konsisten, default
    False (asumsi paling umum/aman).
    """
    tol_pt = 0.3 / PT_TO_CM
    odd_lefts, even_lefts = [], []
    with fitz.open(pdf_path) as pdf:
        for i, page in enumerate(pdf):
            blocks = _extract_blocks(page)
            if not blocks:
                continue
            left = min(b["bbox"][0] for b in blocks)
            physical_page_no = i + 1
            (odd_lefts if physical_page_no % 2 == 1 else even_lefts).append(left)

    if len(odd_lefts) < 3 or len(even_lefts) < 3:
        return False
    return abs(statistics.median(odd_lefts) - statistics.median(even_lefts)) > tol_pt


# --------------------------------------------------------------------
# Table params (best-effort, pakai fitur deteksi tabel bawaan PyMuPDF)
# --------------------------------------------------------------------


def _row_has_shading(page: "fitz.Page", row_bbox: tuple) -> bool:
    """
    Deteksi latar-belakang gelap baris header, dari VECTOR DRAWING
    (rect terisi warna) yang bbox-nya beririsan signifikan dengan
    row_bbox -- PDF tidak simpan "shading cell" sebagai atribut
    langsung kayak w:shd di docx, jadi ini didekati dari objek gambar
    (kotak warna) yang dirender DI BELAKANG teks header. BEST EFFORT:
    kalau gagal terbaca / rendernya bukan vector rect (mis. shading di-
    rasterize jadi gambar), balik False (dianggap tidak shading),
    bukan error -- konsisten dengan filosofi "gagal ke arah aman/
    presisi" yang sudah dipakai checker lain di kodebase ini.
    """
    try:
        drawings = page.get_drawings()
    except Exception:
        return False
    rx0, ry0, rx1, ry1 = row_bbox
    row_area = max(rx1 - rx0, 0) * max(ry1 - ry0, 0)
    if row_area <= 0:
        return False
    for d in drawings:
        fill = d.get("fill")
        if not fill:
            continue
        # Putih/nyaris putih dianggap "tidak ada shading" (kertas
        # polos), bukan warna latar sungguhan.
        if all(c >= 0.95 for c in fill[:3]):
            continue
        dx0, dy0, dx1, dy1 = d.get("rect", (0, 0, 0, 0))
        ix0, iy0 = max(rx0, dx0), max(ry0, dy0)
        ix1, iy1 = min(rx1, dx1), min(ry1, dy1)
        inter = max(ix1 - ix0, 0) * max(iy1 - iy0, 0)
        if inter / row_area >= 0.5:
            return True
    return False


def _detect_tables_on_page(page: "fitz.Page") -> list[dict]:
    """
    Deteksi tabel di SATU halaman pakai page.find_tables() bawaan
    PyMuPDF (heuristik garis/kolom) -- BEST EFFORT, tidak seakurat baca
    w:tbl langsung dari XML docx, terutama untuk tabel tanpa garis
    (borderless) yang cukup umum dipakai di format buku TA. Balik list
    dict MENTAH (belum ada prev_para_idx/next_para_idx/seq -- itu diisi
    belakangan di _extract_paragraphs_and_tables_pdf() yang tahu posisi
    tabel ini relatif ke paragraf sekitarnya), diurutkan dari atas ke
    bawah halaman (dibutuhkan supaya penggabungan dengan paragraf per
    halaman urut benar).
    """
    try:
        found = page.find_tables()
    except Exception:
        return []

    results = []
    for t in found.tables:
        bbox = t.bbox
        rows = list(t.rows) if t.rows else []
        n_rows = len(rows)

        header_format = None
        if n_rows > 1:
            first_row_bbox = rows[0].bbox
            spans_in_row = []
            text_dict = page.get_text("dict")
            for b in text_dict["blocks"]:
                if b.get("type") != 0:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sx0, sy0, sx1, sy1 = span["bbox"]
                        if (
                            sy0 >= first_row_bbox[1] - 1
                            and sy1 <= first_row_bbox[3] + 1
                            and sx0 >= first_row_bbox[0] - 1
                            and sx1 <= first_row_bbox[2] + 1
                            and span.get("text", "").strip()
                        ):
                            spans_in_row.append(span)
            if spans_in_row:
                row_center = (first_row_bbox[0] + first_row_bbox[2]) / 2
                text_x0 = min(s["bbox"][0] for s in spans_in_row)
                text_x1 = max(s["bbox"][2] for s in spans_in_row)
                text_center = (text_x0 + text_x1) / 2
                header_format = {
                    "header_all_bold": all(_span_bold(s) for s in spans_in_row),
                    "header_all_center": abs(text_center - row_center) <= 10,
                    "header_has_shading": _row_has_shading(page, first_row_bbox),
                }

        alignment = None
        page_center = page.rect.width / 2
        tbl_center = (bbox[0] + bbox[2]) / 2
        if abs(tbl_center - page_center) <= 15:
            alignment = "center"

        results.append(
            {
                "bbox": bbox,
                "n_rows": n_rows,
                "alignment": alignment,
                "header_format": header_format,
            }
        )

    results.sort(key=lambda t: t["bbox"][1])
    return results


def _find_prev_para_idx(
    paragraphs: list[dict], page_index: int, table_bbox: tuple
) -> int | None:
    """Index paragraf TERAKHIR yang posisinya di atas/sebelum tabel ini."""
    best = None
    for p in paragraphs:
        before = p["page_index"] < page_index or (
            p["page_index"] == page_index and p["bbox"][3] <= table_bbox[1] + 1
        )
        if before and (best is None or p["index"] > best):
            best = p["index"]
    return best


def _find_next_para_idx(
    paragraphs: list[dict], page_index: int, table_bbox: tuple
) -> int | None:
    """Index paragraf PERTAMA yang posisinya di bawah/setelah tabel ini."""
    best = None
    for p in paragraphs:
        after = p["page_index"] > page_index or (
            p["page_index"] == page_index and p["bbox"][1] >= table_bbox[3] - 1
        )
        if after and (best is None or p["index"] < best):
            best = p["index"]
    return best


def extract_table_params_pdf(pdf_path: str, paragraphs: list[dict]) -> list[dict]:
    """
    Padanan extract_table_params(doc) versi docx -- skema output SAMA
    PERSIS (seq/alignment/prev_para_idx/next_para_idx/n_rows/
    header_format) supaya compare_visual_objects() di comparator.py
    bisa dipakai tanpa perubahan sama sekali.

    WAJIB dipanggil SETELAH extract_paragraph_params_pdf() -- prev/next
    para idx dicari dari posisi (page_index + bbox) paragraf yang sudah
    diekstrak lebih dulu, bukan dihitung independen.
    """
    tables_data = []
    seq = 0
    with fitz.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf):
            for t in _detect_tables_on_page(page):
                seq += 1
                tables_data.append(
                    {
                        "seq": seq,
                        "alignment": t["alignment"],
                        "prev_para_idx": _find_prev_para_idx(
                            paragraphs, page_index, t["bbox"]
                        ),
                        "next_para_idx": _find_next_para_idx(
                            paragraphs, page_index, t["bbox"]
                        ),
                        "n_rows": t["n_rows"],
                        "header_format": t["header_format"],
                    }
                )
    return tables_data


# --------------------------------------------------------------------
# Paragraph params -- state machine chapter tracking (DIPORTING dari
# extract_paragraph_params() di extractor.py, bagian yang murni
# berbasis TEKS dipertahankan APA ADANYA; bagian yang butuh objek docx
# diganti versi PDF-nya, lihat keterbatasan #1-#2 di docstring modul)
# --------------------------------------------------------------------


def extract_paragraph_params_pdf(pdf_path: str) -> list[dict]:
    sections = extract_section_params_pdf(pdf_path)
    left_margin_pt = sections[0]["left_twips"] / 20
    right_margin_pt = sections[0]["width_twips"] / 20 - sections[0]["right_twips"] / 20

    elements = []

    current_chapter = "Pra-Bab"
    expect_chapter_name = False
    in_figure_table_list = False
    in_toc = False
    seen_bab = False
    abstract_state = None
    in_special_page = True

    idx = 0
    with fitz.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf):
            table_bboxes = [t["bbox"] for t in _detect_tables_on_page(page)]
            blocks = _extract_blocks(page, table_bboxes)
            blocks = _merge_wrapped_paragraph_blocks(
                blocks, left_margin_pt, right_margin_pt
            )
            for block in blocks:
                stripped = block["full_text"].strip()
                if not stripped:
                    continue

                ptype = classify_paragraph(stripped, para=None)
                upper_now = stripped.upper()

                if upper_now in NORMAL_FRONT_MATTER_MARKERS:
                    in_special_page = False
                elif (
                    upper_now in SPECIAL_FRONT_PAGE_MARKERS
                    and not in_toc
                    and not in_figure_table_list
                ):
                    in_special_page = True

                if BLANK_PAGE_NOTICE_RE.match(stripped):
                    ptype = "notice"
                    expect_chapter_name = False
                elif in_special_page and not seen_bab:
                    ptype = "cover"
                    expect_chapter_name = False
                else:
                    upper_stripped = stripped.upper()
                    # BUG UTAMA (fixed): classify_paragraph() murni
                    # berbasis string, jadi TIDAK TAHU dia lagi ada di
                    # dalam region Daftar Isi/Daftar Gambar/Daftar
                    # Tabel atau bukan. Entri TOC yang teksnya "BAB 1
                    # PENDAHULUAN .......... 5" tetap match
                    # CHAPTER_TITLE_RE sama seperti heading ASLI "BAB 1
                    # PENDAHULUAN" -- versi lama percaya begitu saja ke
                    # ptype=="chapter_title" dan MENIMPA current_chapter/
                    # seen_bab/in_toc di baris TOC itu sendiri, padahal
                    # ini bukan heading sungguhan. Akibatnya domino:
                    # in_toc keburu di-reset False di entri TOC pertama
                    # yang match, jadi SEMUA entri TOC berikutnya
                    # (subbab, DAFTAR GAMBAR, dst) ikut lolos jadi
                    # heading beneran juga, minta font 14pt/bold/center
                    # yang seharusnya tidak berlaku buat baris daftar
                    # isi. Fix: kalau lagi di dalam TOC/daftar
                    # gambar/tabel DAN baris ini masih berciri entri
                    # (dot-leader di ujung -- lihat _TOC_LEADER_RE -- atau
                    # classify_paragraph() TIDAK bilang chapter_title
                    # sama sekali), demosi paksa ke "toc_entry" TANPA
                    # menyentuh current_chapter/seen_bab/abstract_state/
                    # expect_chapter_name apapun. Heading ASLI (tanpa
                    # dot-leader) tetap lolos ke cabang normal di bawah,
                    # termasuk yang menutup region TOC (bukan didemosi).
                    is_toc_leader_line = bool(_TOC_LEADER_RE.search(stripped))
                    if (in_toc or in_figure_table_list) and (
                        is_toc_leader_line or ptype != "chapter_title"
                    ):
                        ptype = "toc_entry"
                    elif is_toc_leader_line and not seen_bab:
                        # BUG DIPERBAIKI (jalur PDF -- entri "Daftar
                        # Gambar"/"Daftar Tabel" individual ke-flag salah
                        # ukuran font, mis. "Ukuran font 12.0pt tidak
                        # sesuai untuk tipe 'caption', seharusnya 11pt"
                        # padahal baris itu SAMA SEKALI BUKAN caption
                        # gambar -- itu entri daftar isi di halaman
                        # "Daftar Gambar" yang isinya memang body text
                        # 12pt, bukan caption 11pt): in_toc/
                        # in_figure_table_list di atas cuma di-set True
                        # lewat heading literal "DAFTAR ISI"/"DAFTAR
                        # GAMBAR"/"DAFTAR TABEL", dan gampang ke-reset
                        # False lagi kalau ADA baris di antaranya yang
                        # (keliru) lolos classify_paragraph() sebagai
                        # ptype=="chapter_title" tanpa dot-leader --
                        # domino yang sama seperti bug "BUG UTAMA" di atas,
                        # cuma triggernya baris lain di tengah daftar,
                        # bukan baris pembuka daftarnya sendiri. Daripada
                        # terus menerus tambal state tracking yang rapuh
                        # ini, dipasang jaring pengaman independen: baris
                        # apa pun yang berpola dot-leader + nomor halaman
                        # di ujung (is_toc_leader_line) DAN muncul SEBELUM
                        # heading "BAB 1 ..." pertama (not seen_bab) HAMPIR
                        # PASTI entri Daftar Isi/Daftar Gambar/Daftar Tabel
                        # -- pola ini tidak pernah muncul di body text
                        # sungguhan pada bagian depan buku (cover,
                        # abstrak, dst.), jadi aman didemosi ke "toc_entry"
                        # tanpa bergantung sama sekali ke status
                        # in_toc/in_figure_table_list.
                        ptype = "toc_entry"
                    elif ptype == "chapter_title":
                        if upper_stripped == "ABSTRAK":
                            abstract_state = "id"
                        elif upper_stripped == "ABSTRACT":
                            abstract_state = "en"
                        else:
                            abstract_state = None

                        if (
                            CHAPTER_TITLE_RE.match(upper_stripped)
                            and stripped == upper_stripped
                        ):
                            current_chapter = re.match(
                                r"^BAB\s+\d+", upper_stripped
                            ).group(0)
                            seen_bab = True
                            expect_chapter_name = bool(
                                CHAPTER_TITLE_ONLY_RE.match(upper_stripped)
                            )
                        elif seen_bab:
                            # NOTE (keterbatasan #1): cabang "heading
                            # Heading-1-style tanpa teks BAB N literal"
                            # (jalur bab_counter di extractor.py) TIDAK
                            # ada padanannya di sini -- info numbering
                            # Word sudah hilang begitu jadi PDF. Kalau
                            # dokumen konsisten menulis "BAB N" literal
                            # (mayoritas kasus), baris ini tidak pernah
                            # kepakai.
                            current_chapter = upper_stripped
                            expect_chapter_name = False
                        else:
                            expect_chapter_name = False

                        in_figure_table_list = upper_stripped in (
                            "DAFTAR GAMBAR",
                            "DAFTAR TABEL",
                        )
                        in_toc = upper_stripped == "DAFTAR ISI"
                    elif expect_chapter_name:
                        ptype = "chapter_title"
                        expect_chapter_name = False
                    elif in_toc or in_figure_table_list:
                        # Jaga-jaga: entri TOC yang classify_paragraph()
                        # sudah benar (subchapter/subsubchapter/body,
                        # ptype != "chapter_title") TAPI entah kenapa
                        # tidak match _TOC_LEADER_RE (mis. dot-leader
                        # tidak ke-render sebagai karakter di beberapa
                        # PDF) -- tetap didemosi di sini sebagai
                        # fallback kedua, konsisten dengan perilaku lama.
                        ptype = "toc_entry"
                    elif ptype == "body" and _is_unnumbered_heading_runs(
                        block["runs"], stripped
                    ):
                        # Padanan Jalur 3 classify_paragraph() (heading
                        # tanpa nomor) -- di extractor.py ini dicoba DI
                        # DALAM classify_paragraph() lewat objek para;
                        # di sini classify_paragraph() dipanggil dengan
                        # para=None (jalur itu otomatis skip), jadi
                        # ditambal manual di sini pakai runs PDF.
                        ptype = "subsubchapter_title"

                if ptype != "cover":
                    if any(r.get("is_math") for r in block["runs"]):
                        # Proxy has_equation -- BUKAN deteksi OMML asli
                        # (lihat keterbatasan #3), cuma menandai run
                        # matematis di level PARAGRAF kalau MAYORITAS
                        # blok ini memang simbol matematis (bukan cuma
                        # satu variabel di tengah kalimat body).
                        math_chars = sum(
                            len(r["text"]) for r in block["runs"] if r.get("is_math")
                        )
                        if math_chars / max(len(stripped), 1) > 0.5:
                            ptype = "equation"
                    if ptype != "equation" and SOURCE_CITATION_RE.match(stripped):
                        ptype = "source_citation"

                runs_data = [
                    {
                        "text": r["text"],
                        "font_name": r["font_name"],
                        "font_size_pt": r["font_size_pt"],
                        "bold": r["bold"],
                        "italic": r["italic"],
                        "is_math": r.get("is_math", False),
                    }
                    for r in block["runs"]
                ]

                abstract_part = None
                if abstract_state is not None and ptype == "body":
                    is_keyword_line = bool(ABSTRACT_KEYWORD_RE.match(stripped))
                    if abstract_state == "id":
                        abstract_part = (
                            "keyword_id" if is_keyword_line else "abstrak_body"
                        )
                    else:
                        abstract_part = (
                            "keyword_en" if is_keyword_line else "abstract_body"
                        )

                left_indent_cm, right_indent_cm = _infer_indent_cm(
                    block, left_margin_pt, right_margin_pt
                )

                elements.append(
                    {
                        "index": idx,
                        "page_index": page_index,
                        "bbox": block["bbox"],
                        "type": ptype,
                        "full_text": stripped,
                        "is_upper": stripped == stripped.upper(),
                        "is_titlecase": is_titlecase(stripped),
                        "line_spacing": _infer_line_spacing(block),
                        # PDF tidak punya konsep "rule" (EXACTLY/AT_LEAST
                        # vs MULTIPLE) seperti docx -- line_spacing di
                        # atas SELALU angka kelipatan hasil estimasi.
                        "line_spacing_is_multiple": True,
                        "line_spacing_rule": None,
                        "runs": runs_data,
                        "chapter": current_chapter,
                        "section": 0,
                        "abstract_part": abstract_part,
                        "alignment": _infer_alignment(
                            block, left_margin_pt, right_margin_pt
                        ),
                        "left_indent_cm": left_indent_cm,
                        "right_indent_cm": right_indent_cm,
                        # blank_before/blank_after: konsep "baris kosong"
                        # tidak berlaku sama untuk PDF (baris kosong
                        # tidak menghasilkan block sama sekali) --
                        # selalu False di sini. Cek "1 baris kosong
                        # sebelum/sesudah gambar" jadi TIDAK tervalidasi
                        # lewat jalur PDF; perlu ditinjau ulang di
                        # comparator.py kalau field ini krusial.
                        "blank_before": False,
                        "blank_after": False,
                    }
                )
                idx += 1

    return elements


def extract_all_pdf(pdf_path: str) -> dict:
    """
    Entry point tahap Extract Parameter versi PDF -- padanan extract_all(doc)
    di extractor.py. Bentuk return SAMA PERSIS, supaya main.py &
    comparator.py bisa pakai salah satu (docx atau PDF) tanpa tahu
    bedanya.
    """
    paragraphs = extract_paragraph_params_pdf(pdf_path)
    sections = extract_section_params_pdf(pdf_path)
    tables = extract_table_params_pdf(pdf_path, paragraphs)
    mirror_margins = _infer_mirror_margins(pdf_path)
    return {
        "paragraphs": paragraphs,
        "sections": sections,
        "tables": tables,
        "mirror_margins": mirror_margins,
    }
