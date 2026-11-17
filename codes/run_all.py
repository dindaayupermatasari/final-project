"""Satu pintu SPPA: dokumen (.pdf / .docx) -> pemeriksaan format (visual) + sistematika Bab 1 -> satu report.html.

Taruh file ini di folder proyek yang berisi `aggregator/` (visual ada di aggregator/pipelines/visual,
sistematika di aggregator/pipelines/sistematika atau sejajar dengan file ini).

Contoh:
  python run_all.py --doc "data/samples/Buku PA_Awal Raya.pdf" --judul "JUDUL YANG BENAR" --retrieval cached
  python run_all.py --doc "data/samples/Buku PA_Awal Raya.docx" --retrieval cached
  python run_all.py --doc "...pdf" --skip-sistematika      # hanya format
  python run_all.py --doc "...docx" --skip-visual          # hanya sistematika

Jalur format:
  .pdf  -> pdf_extractor (PyMuPDF), compare_all(source="pdf")
  .docx -> extractor (python-docx), dirender ke PDF lewat LibreOffice hanya untuk tata letak halaman
           (ganjil/genap, nomor footer), compare_all(source="docx")
Jalur sistematika membaca .pdf maupun .docx langsung lewat sectionizer.

Argumen yang tidak dikenal di sini (--judul, --mode, --retrieval, --yes, --pause, --only, ...)
diteruskan apa adanya ke sistematika/pipeline/run_bab1.py.

Nanti backend cukup memanggil run(doc, extra_args) lalu mengirim report.html / summary.json dari folder yang dikembalikan.
"""
import argparse
import json
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
SUPPORTED_EXT = (".pdf", ".docx")

# `aggregator` (modul visual meng-import aggregator.config dan aggregator.pipelines.visual.*) harus bisa di-import
for _p in (HERE, *HERE.parents):
    if (_p / "aggregator").is_dir():
        sys.path.insert(0, str(_p))
        break


def _sistematika_dir():
    kandidat = [HERE / "sistematika", HERE / "aggregator" / "pipelines" / "sistematika", HERE.parent / "sistematika"]
    kandidat += [p.parent.parent for p in sorted(HERE.glob("aggregator/**/sistematika/pipeline/run_bab1.py"))]
    for base in kandidat:
        if (base / "pipeline" / "run_bab1.py").exists():
            sys.path.insert(0, str(base / "pipeline"))
            return base
    sys.exit(f"!! sistematika/pipeline/run_bab1.py tidak ditemukan di bawah {HERE}.")


def _save(run_dir, name, obj):
    (run_dir / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def _extract_pdf(pdf_path, plc, comparator):
    """Jalur PDF. Mengembalikan (extracted, findings, checks, page_texts, footer_numbers, section_pages, font_status, render_error)."""
    from aggregator.pipelines.visual import pdf_extractor

    extracted = pdf_extractor.extract_all_pdf(pdf_path)
    extracted = plc.finalize_page_numbering(extracted, pdf_path)
    section_pages = {s["index"]: s["start_physical_page"] for s in extracted["sections"]}
    findings, checks = comparator.compare_all(extracted, section_pages, source="pdf")
    # PDF mahasiswa adalah halaman aslinya (bukan hasil render ulang), jadi font_status "verified"
    return (extracted, findings, checks, plc.extract_page_texts(pdf_path),
            plc.extract_footer_numbers(pdf_path), section_pages, "verified", None)


def _extract_docx(docx_path, plc, comparator, tmp_dir):
    """Jalur DOCX. Parameter format dibaca langsung dari XML Word (extractor.py); render ke PDF
    hanya dipakai untuk tata letak halaman, karena nomor halaman tidak tersimpan di XML.
    Kalau render gagal (LibreOffice crash, font tidak ada, dst), cek tata letak halaman dilewati
    (page_texts=None) dan alasannya dikembalikan di `render_error`; cek format lain tetap jalan."""
    import docx
    from aggregator.pipelines.visual import extractor

    extracted = extractor.extract_all(docx.Document(docx_path))
    page_texts = footer_numbers = None
    section_pages, font_status, render_error = {}, "unconfirmed", None
    try:
        render_pdf, font_status = plc.render_to_pdf(docx_path, tmp_dir)  # gagal keras kalau font tidak terpasang
        page_texts = plc.extract_page_texts(render_pdf)
        footer_numbers = plc.extract_footer_numbers(render_pdf)
        section_pages = plc.find_section_start_pages(page_texts, extracted["paragraphs"])
    except Exception as e:  # noqa: BLE001
        page_texts = footer_numbers = None
        section_pages = {}
        render_error = f"{type(e).__name__}: {e}"
    findings, checks = comparator.compare_all(extracted, section_pages, source="docx")
    return extracted, findings, checks, page_texts, footer_numbers, section_pages, font_status, render_error


def run_visual(doc_path, run_dir):
    """Pemeriksaan format untuk PDF atau DOCX. Urutan panggilannya mengikuti docstring di modul visual.
    Tahap utama gagal -> {"error": ...}; tahap tambahan (Daftar Pustaka, tata letak halaman) gagal -> dicatat di "errors"."""
    ext = Path(doc_path).suffix.lower()
    tmp_dir = tempfile.mkdtemp(prefix="sppa_render_")
    try:
        from aggregator.pipelines.visual import bibliography_checker, comparator
        from aggregator.pipelines.visual import page_layout_checker as plc
        from aggregator.pipelines.visual.scorer import calculate_score

        if ext == ".pdf":
            hasil = _extract_pdf(doc_path, plc, comparator)
        else:
            hasil = _extract_docx(doc_path, plc, comparator, tmp_dir)
        extracted, findings, checks, page_texts, footer_numbers, section_pages, font_status, render_error = hasil
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return {"error": f"{type(e).__name__}: {e}", "file": Path(doc_path).name}

    errors = []

    def tambah(tahap, fn):
        try:
            f, c = fn()
        except Exception as e:  # noqa: BLE001
            errors.append({"tahap": tahap, "error": f"{type(e).__name__}: {e}"})
            return
        findings.extend(f)
        for ch, v in c.items():
            checks.setdefault(ch, {"checked": 0, "failed": 0})["checked"] += v.get("checked", 0)

    tambah("Daftar Pustaka", lambda: bibliography_checker.check_bibliography(extracted["paragraphs"]))
    if page_texts is None:
        errors.append({"tahap": "Tata letak halaman", "error": render_error or "render PDF tidak tersedia"})
    else:
        tambah("Tata letak halaman", lambda: plc.check_page_layout(
            page_texts, extracted["sections"], section_pages, footer_numbers, font_status=font_status))
    shutil.rmtree(tmp_dir, ignore_errors=True)

    # jumlah temuan per bab dihitung ulang dari daftar temuan akhir (cara yang sama seperti compare_all)
    for c in checks.values():
        c["failed"] = 0
    for f in findings:
        checks.setdefault(f.get("chapter") or "Lainnya", {"checked": 0, "failed": 0})["failed"] += 1
    score = calculate_score(checks)

    _save(run_dir, "10_visual_findings.json", findings)
    _save(run_dir, "11_visual_score.json", score)
    print(f"    [{ext[1:]}] skor format {score['overall_score']} dengan {score['total_findings']} temuan"
          + (f"; tahap gagal: {[e['tahap'] for e in errors]}" if errors else ""))
    for e in errors:
        print(f"    !! {e['tahap']}: {e['error']}")
    return {"file": Path(doc_path).name, "findings": findings, "score": score, "errors": errors}


def run(doc, extra_args=(), skip_visual=False, skip_sistematika=False):
    """Jalankan kedua pemeriksaan ke satu folder hasil. Mengembalikan Path folder itu."""
    doc = str(Path(doc).resolve())
    if not Path(doc).is_file():
        sys.exit(f"!! File tidak ditemukan: {doc}")
    if Path(doc).suffix.lower() not in SUPPORTED_EXT:
        sys.exit("!! run_all.py hanya menerima .pdf atau .docx.")
    base = _sistematika_dir()
    from report import build_visual_only  # report.py ada di sistematika/pipeline (sudah di sys.path)

    run_dir = base / "debug_output" / "runs" / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True)
    print(f"Folder hasil: {run_dir}")

    visual = None
    if not skip_visual:
        print("\n=== Format (visual)")
        visual = run_visual(doc, run_dir)
        if visual.get("error"):
            print(f"    !! {visual['error']}")

    if not skip_sistematika:
        import run_bab1 as rb  # baru di-import di sini: menarik openai/chromadb, tidak perlu kalau hanya format

        try:
            rb.main(["--doc", doc, *extra_args], run_dir=run_dir, visual=visual)
            return run_dir
        except SystemExit as e:  # run_bab1 berhenti lewat sys.exit(pesan) kalau ada yang kurang
            if e.code in (0, None):
                return run_dir
            catatan = e.code if isinstance(e.code, str) else f"keluar dengan kode {e.code}"
            print(catatan, file=sys.stderr)
            if not visual:
                raise
            (run_dir / "report.html").write_text(build_visual_only(visual, Path(doc).name, catatan), encoding="utf-8")
            print(f"    Hasil format tetap disimpan: {run_dir / 'report.html'}")
            return run_dir

    (run_dir / "report.html").write_text(build_visual_only(visual, Path(doc).name), encoding="utf-8")
    print(f"    Buka: {run_dir / 'report.html'}")
    return run_dir


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--doc", required=True, help="dokumen mahasiswa (.pdf atau .docx)")
    ap.add_argument("--skip-visual", action="store_true")
    ap.add_argument("--skip-sistematika", action="store_true")
    a, rest = ap.parse_known_args()
    if a.skip_visual and a.skip_sistematika:
        ap.error("tidak ada yang dijalankan")
    run(a.doc, rest, a.skip_visual, a.skip_sistematika)


if __name__ == "__main__":
    main()
