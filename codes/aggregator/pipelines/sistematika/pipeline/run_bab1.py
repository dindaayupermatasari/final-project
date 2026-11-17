"""Pipeline sistematika: judul -> subbab (semua bab) -> RAG -> LLM memilih panduan -> penilaian -> laporan.

Default memeriksa SEMUA bab yang terbaca dari dokumen (--bab all). Bisa dibatasi: --bab 1  atau  --bab 1,5.
Nama file tetap run_bab1.py supaya run_all.py tidak perlu diubah.

Mode penilaian (--mode):
  kalimat (default)  kriteria dari chunk panduan (RAG) + aturan "dilarang" per kalimat, digabung jadi satu status
  rag                hanya kriteria dari chunk panduan (RAG)
  manual             kriteria tulisan tangan di criteria.py
  both               rag + manual

Contoh:
  python run_bab1.py --doc "C:\\...\\CheckRun\\data\\samples\\Buku PA_Awal Raya.pdf"      # dari dokumen asli (judul diambil dari sampul)
  python run_bab1.py --doc "...pdf" --judul "JUDUL YANG BENAR"                          # judul diketik manual
  python run_bab1.py --sections hasil_sectionizer.json --judul "..."                   # dari JSON sectionizer yang sudah ada
  python run_bab1.py --judul "..." --mode both --pause          # tekan Enter di tiap langkah
  python run_bab1.py --judul "..." --retrieval cached           # tanpa Ollama/ChromaDB
  python run_bab1.py --judul "..." --only "TUJUAN,PERMASALAHAN"
  python run_bab1.py --doc "...pdf" --bab 1,5                                           # hanya Bab 1 dan Bab 5

Hemat token: jawaban LLM disimpan di debug_output/llm_cache.jsonl. Kalau kena batas token harian, jalankan lagi
perintah yang sama setelah kuota reset: bagian yang sudah selesai diambil dari cache, hanya sisanya yang memanggil LLM.

Keluaran tiap langkah disimpan di debug_output/runs/<waktu>/ (01_title.json ... report.html).
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent  # folder sistematika/
sys.path.insert(0, str(BASE / "llm"))
sys.path.insert(0, str(HERE))

from criteria import get_criteria  # noqa: E402
from llm_client import DailyLimitError  # noqa: E402
from doc_input import extract_title, sectionize_doc  # noqa: E402
from judge import clean_text, judge_subbab  # noqa: E402
from rag_judge import judge_with_guideline, match_guideline, rule_pick  # noqa: E402
from sentence_judge import _canon, aturan_sebagai_kriteria, gabung_status, has_aturan, judge_subbab_kalimat, selaraskan_dengan_label  # noqa: E402
from report import build_report  # noqa: E402
from retriever import SnapshotIndex, make_retriever  # noqa: E402
from title_decomposer import decompose_title  # noqa: E402


def find_file(name):
    hits = [p for p in sorted(BASE.rglob(name)) if "runs" not in p.parts]
    if not hits:
        sys.exit(f"!! {name} tidak ditemukan di bawah {BASE}.")
    return hits[0]


def save(run_dir, name, obj):
    (run_dir / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def header(n, text, pause):
    print(f"\n=== Langkah {n}: {text}")
    if pause:
        input("    (Enter untuk menjalankan) ")


# Subbab yang menjawab Bab 1 diberi cuplikan Bab 1 mahasiswa sebagai konteks (kunci = id chunk panduan).
KONTEKS = {
    "bab4_sub4-7_instruction": ["TUJUAN"],                  # analisis harus menjawab apakah solusi pada Tujuan menyelesaikan masalah
    "bab5_sub5-1_instruction": ["PERMASALAHAN", "TUJUAN"],  # kesimpulan terhadap Problem, solusi, dan kinerja
}


def konteks_bab1(all_secs, chunk_id, maks=1500):
    teks = []
    for judul in KONTEKS.get(chunk_id, []):
        for sec in all_secs:
            if str(sec.get("bab_num")) == "1" and _canon(sec["subbab_title"]) == judul:
                teks.append(f"{judul}: {clean_text(sec.get('raw_text', ''))[:maks]}")
                break
    return "\n".join(teks)


def pilih_bab(nilai, tersedia):
    if nilai.strip().lower() in ("all", "semua"):
        return list(tersedia)
    want = {w.strip() for w in nilai.split(",") if w.strip()}
    return [b for b in tersedia if b in want]


def cek_struktur(snap, subs, bab_dinilai):
    """Kelengkapan: tiap petunjuk panduan di bab yang diperiksa harus punya subbab mahasiswa yang cocok."""
    dipilih = {}
    for x in subs:
        cid = (x.get("match") or {}).get("pilihan_llm")
        if cid:
            dipilih.setdefault(cid, []).append(f"{x['num']} {x['title']}")
    rows = []
    for cid, it in snap.items.items():
        m = it.get("metadata", {})
        if not cid.endswith("_instruction") or str(m.get("bab_num")) not in bab_dinilai:
            continue
        rows.append({"chunk_id": cid, "bab": str(m.get("bab_num")), "num": m.get("subbab_num"), "title": m.get("subbab_title"),
                     "ditemukan": cid in dipilih, "subbab_mahasiswa": dipilih.get(cid, [])})
    return rows


def confirm_decomp(decomp, auto_yes):
    """Gerbang konfirmasi: pipeline berhenti sampai user membenarkan/mengoreksi PD, Problem, Uniqueness."""
    names = {"problem_domain": "Problem Domain", "problem": "Problem", "uniqueness": "Uniqueness"}
    print("\n    Hasil pemecahan judul (usulan LLM):")
    for k, n in names.items():
        print(f"      {n:14}: {decomp[k] or '(KOSONG - tidak ketemu di judul)'}")
    if decomp.get("catatan"):
        print(f"      catatan LLM   : {decomp['catatan']}")
    if auto_yes and not any(not decomp[k] for k in names):
        decomp["dikonfirmasi"] = True
        return decomp
    if not sys.stdin.isatty():
        sys.exit("!! Butuh konfirmasi user (ada unsur kosong / tanpa --yes). Isi lewat --problem-domain/--problem/--uniqueness.")
    for k, n in names.items():
        v = input(f"    {n} [Enter/ok = pakai usulan di atas; atau ketik pengganti]: ").strip()
        if v and v.lower() not in {"ok", "ya", "y", "iya", "benar", "betul"}:
            decomp[k] = v
    if not all(decomp[k] for k in names):
        sys.exit("!! Problem Domain, Problem, dan Uniqueness harus terisi sebelum Bab 1 dicek.")
    decomp["perlu_dicek"] = []
    decomp["dikonfirmasi"] = True
    return decomp


def main(argv=None, run_dir=None, visual=None):
    """argv/run_dir/visual dipakai run_all.py: argv = daftar argumen CLI, run_dir = folder hasil yang dipakai bersama,
    visual = hasil pemeriksaan format untuk ikut ditampilkan di report.html. Dipanggil dari terminal: semuanya kosong.
    Mengembalikan dict `data` yang sama dengan yang dipakai build_report."""
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", help="dokumen mahasiswa (.pdf/.docx); sectionizer dijalankan otomatis")
    ap.add_argument("--judul", help="judul proyek akhir (kalau kosong dan memakai --doc: diambil dari sampul)")
    ap.add_argument("--sections", help="hasil_sectionizer.json yang sudah ada (alternatif dari --doc)")
    ap.add_argument("--bab", default="all", help="all (default) atau daftar nomor bab dipisah koma, mis. 1,5")
    ap.add_argument("--only", help="judul subbab dipisah koma, mis. TUJUAN,PERMASALAHAN")
    ap.add_argument("--mode", choices=["kalimat", "rag", "manual", "both"], default="kalimat")
    ap.add_argument("--retrieval", choices=["live", "cached"], default="live")
    ap.add_argument("--topk", type=int, default=3)
    ap.add_argument("--match", choices=["auto", "llm"], default="auto",
                    help="auto: judul subbab sama dengan judul panduan -> dipakai langsung, LLM hanya untuk judul yang tidak sama; "
                         "llm: LLM memilih untuk semua subbab (lebih boros token)")
    ap.add_argument("--max-chars", type=int, default=9000, help="batas karakter teks mahasiswa per panggilan LLM (subbab lebih panjang dibaca per jendela)")
    ap.add_argument("--max-windows", type=int, default=6, help="maksimal jendela yang dibaca per subbab")
    ap.add_argument("--chroma-dir")
    ap.add_argument("--pause", action="store_true", help="berhenti di tiap langkah (untuk demo)")
    ap.add_argument("--yes", action="store_true", help="lewati konfirmasi kalau ketiga unsur sudah terisi (hanya untuk tes)")
    ap.add_argument("--problem-domain")
    ap.add_argument("--problem")
    ap.add_argument("--uniqueness")
    a = ap.parse_args(argv)
    if not (a.doc or a.sections):
        ap.error("isi --doc (dokumen mahasiswa) atau --sections (JSON sectionizer)")
    if not a.judul and not a.doc:
        ap.error("--judul wajib kalau memakai --sections")

    run_dir = Path(run_dir) if run_dir else BASE / "debug_output" / "runs" / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    snapshot_path = find_file("production_snapshot.json")
    snap = SnapshotIndex(snapshot_path)
    sumber = Path(a.doc).name if a.doc else Path(a.sections).name
    save(run_dir, "00_input.json", {"doc": a.doc, "sections": a.sections, "judul": a.judul, "opsi": vars(a)})
    print(f"Folder hasil: {run_dir}")
    stop = {"msg": None}  # terisi kalau kuota token harian habis; sisa pekerjaan ditandai "ditunda"

    # ---- Langkah 0: baca dokumen lewat sectionizer
    header(0, f"baca dokumen ({sumber})", a.pause)
    if a.doc:
        try:
            all_secs, n_rej = sectionize_doc(a.doc)
        except Exception as e:  # noqa: BLE001
            sys.exit(f"!! Sectionizer gagal: {e}")
        save(run_dir, "00_sectionizer.json", [{k: v for k, v in x.items() if k != "raw_text"} | {"chars": len(x["raw_text"])} for x in all_secs])
        print(f"    {len(all_secs)} subbab terdeteksi, {n_rej} kandidat heading ditolak validasi")
    else:
        sections_path = Path(a.sections)
        all_secs = json.load(open(sections_path, encoding="utf-8"))["sections"]
        print(f"    {len(all_secs)} subbab dari {sections_path.name}")
    babs = sorted({str(s.get("bab_num")) for s in all_secs}, key=lambda b: (not b.isdigit(), int(b) if b.isdigit() else 0, b))
    print(f"    bab yang terbaca: {', '.join(babs)}")
    bab_dinilai = pilih_bab(a.bab, babs)
    if not bab_dinilai:
        sys.exit(f"!! Bab {a.bab} tidak ditemukan. Bab yang terbaca: {', '.join(babs)}")

    # ---- Langkah 1: pecah judul
    header(1, "pemecahan judul (LLM)", a.pause)
    try:
        if not a.judul:
            a.judul, ok = extract_title(a.doc)
            print(f"    Judul dari sampul: {a.judul or '(tidak ketemu)'}")
            if not ok:
                sys.exit("!! Judul dari sampul tidak bisa dipastikan. Jalankan lagi dengan --judul \"...\".")
        if a.problem_domain and a.problem and a.uniqueness:
            decomp = {"judul": a.judul, "problem_domain": a.problem_domain, "problem": a.problem,
                      "uniqueness": a.uniqueness, "catatan": "diisi manual", "perlu_dicek": []}
        else:
            decomp = decompose_title(a.judul)
            for f in ("problem_domain", "problem", "uniqueness"):
                v = getattr(a, f)
                if v:  # koreksi manual satu unsur saja
                    decomp[f] = v
                    decomp["perlu_dicek"] = [p for p in decomp["perlu_dicek"] if not p.startswith(f)]
                    decomp["catatan"] = (decomp.get("catatan", "") + f" [{f} diisi manual]").strip()
    except DailyLimitError as e:
        sys.exit(f"!! Kuota token harian LLM habis sebelum pemeriksaan dimulai: {str(e)[:250]}\n   Jalankan lagi nanti; panggilan yang sudah selesai tersimpan di cache.")
    decomp = confirm_decomp(decomp, a.yes)
    save(run_dir, "01_title.json", decomp)
    print(f"    Problem Domain: {decomp['problem_domain']}\n    Problem       : {decomp['problem']}\n    Uniqueness    : {decomp['uniqueness']}")
    if decomp["perlu_dicek"]:
        print(f"    !! perlu dicek manual: {decomp['perlu_dicek']}")

    # ---- Langkah 2: subbab dokumen mahasiswa
    header(2, f"subbab dokumen mahasiswa (bab {', '.join(bab_dinilai)})", a.pause)
    secs = [s for s in all_secs if str(s.get("bab_num")) in bab_dinilai]
    secs = [s for s in secs if not str(s["subbab_num"]).endswith(".0")]  # teks pembuka bab, bukan subbab
    if not secs:
        sys.exit(f"!! Tidak ada subbab di bab {', '.join(bab_dinilai)}.")
    if a.only:
        wanted = {w.strip().upper() for w in a.only.split(",")}
        secs = [s for s in secs if s["subbab_title"].strip().upper() in wanted]
    subs = []
    for s in secs:
        text = clean_text(s.get("raw_text", ""))
        skip = bool(s.get("tingkat_tiga"))
        subs.append({"bab": str(s.get("bab_num")), "num": s["subbab_num"], "title": s["subbab_title"].strip(), "chars": len(text), "text": text, "skip": skip})
        print(f"    {s['subbab_num']} {s['subbab_title']} ({len(text)} karakter)" + ("  [tingkat 3, tidak dinilai]" if skip else ""))
    save(run_dir, "02_sections.json", [{k: v for k, v in x.items() if k != "text"} | {"preview": x["text"][:200]} for x in subs])
    if not subs:
        sys.exit("!! Tidak ada subbab yang cocok.")
    total = sum(x["chars"] for x in subs)
    print(f"    total {total} karakter teks mahasiswa (perkiraan <= {total // 3} token masukan kalau semuanya dibaca penuh)")

    use_rag = a.mode in ("kalimat", "rag", "both")

    # ---- Langkah 3: retrieval
    header(3, f"retrieval panduan ({a.retrieval}, top-{a.topk})", a.pause)
    try:
        retr = make_retriever(a.retrieval, snap, str(find_file("retrieval_test_from_student_doc.json")) if a.retrieval == "cached" else None, a.chroma_dir)
    except Exception as e:  # noqa: BLE001
        sys.exit(f"!! Retriever gagal dibuat: {e}\n   Coba: --retrieval cached, atau cek Ollama/ChromaDB dan --chroma-dir.")
    for x in subs:
        if x.get("skip"):  # subbab tingkat 3 tanpa induk: tidak ada panduan, tidak diretrieval
            x["candidates"] = []
            continue
        try:
            x["candidates"] = retr.retrieve(x["title"].upper(), a.topk, label_hint=f"BAB {x['bab']} > {x['num']} {x['title']}")
            print(f"    {x['num']} {x['title']}: " + ", ".join(f"{c['rank']}.{c['subbab_title']}({c['distance']})" for c in x["candidates"]))
        except Exception as e:  # noqa: BLE001
            x["candidates"], x["retrieval_error"] = [], str(e)
            print(f"    {x['num']} {x['title']}: !! {e}")
    save(run_dir, "03_retrieval.json", [{"subbab": f"{x['num']} {x['title']}", "error": x.get("retrieval_error"),
        "kandidat": [{k: v for k, v in c.items() if k not in ("text", "example")} for c in x["candidates"]]} for x in subs])

    # ---- Langkah 4: pilih chunk panduan (judul sama -> langsung; selain itu LLM)
    header(4, "memilih chunk panduan dari kandidat", a.pause)
    for x in subs:
        if x.get("skip"):
            continue
        if not x["candidates"] or not x["text"]:
            x["match_error"] = "tidak ada kandidat atau teks kosong"
            continue
        if stop["msg"]:
            x["match_error"] = "ditunda: kuota token harian habis"
            continue
        try:
            sama = rule_pick(x["title"], x["candidates"])
            if a.match == "auto" and sama:
                x["match"] = {"pilihan_llm": sama, "alasan": "judul subbab sama dengan judul panduan (tanpa LLM)", "pilihan_valid": True,
                              "pilihan_aturan": sama, "dari": "judul"}
            else:
                x["match"] = match_guideline(x["title"], x["text"], x["candidates"]) | {"dari": "llm"}
            m = x["match"]
            print(f"    {x['num']} {x['title']} -> {m['pilihan_llm'] or 'tidak ada padanan'} ({m['dari']}; aturan: {m['pilihan_aturan'] or '-'})")
        except DailyLimitError as e:
            stop["msg"] = str(e)
            x["match_error"] = "ditunda: kuota token harian habis"
            print(f"    !! kuota token harian habis, sisa subbab ditunda")
        except Exception as e:  # noqa: BLE001
            x["match_error"] = str(e)
            print(f"    {x['num']} {x['title']}: !! {e}")
    save(run_dir, "04_match.json", [{"subbab": f"{x['num']} {x['title']}", "match": x.get("match"), "error": x.get("match_error")} for x in subs])

    # ---- Langkah 5: penilaian
    header(5, f"penilaian (mode: {a.mode})", a.pause)
    for x in subs:
        nama = f"{x['num']} {x['title']}"
        if x.get("skip"):
            x["rag"] = {"status": "tidak_dinilai", "catatan": "subbab tingkat 3 (x.y.z) tidak punya panduan; tidak dinilai"} if use_rag else None
            print(f"    [RAG]    {nama}: tidak_dinilai (tingkat 3)")
            continue
        if not x["text"]:
            x["rag"] = {"status": "kosong", "catatan": "teks subbab kosong (heading langsung diikuti subbab lain)"} if use_rag else None
            continue
        if use_rag:
            chosen = (x.get("match") or {}).get("pilihan_llm")
            if stop["msg"] and not x.get("match") and x.get("match_error", "").startswith("ditunda"):
                x["rag"] = {"status": "ditunda", "catatan": "kuota token harian LLM habis; jalankan ulang nanti (hasil sebelumnya ada di cache)"}
            elif not chosen:
                x["rag"] = {"status": "tidak_ada_padanan", "catatan": "tidak ada chunk panduan yang cocok untuk subbab ini (di luar panduan)"}
            elif stop["msg"]:
                x["rag"] = {"status": "ditunda", "catatan": "kuota token harian LLM habis; jalankan ulang nanti (hasil sebelumnya ada di cache)"}
            else:
                try:
                    chunk = next(c for c in x["candidates"] if c["chunk_id"] == chosen)
                    x["rag"] = judge_with_guideline(decomp, x["title"], x["text"], chunk, konteks=konteks_bab1(all_secs, chosen),
                                                    max_chars=a.max_chars, max_windows=a.max_windows)
                except DailyLimitError as e:
                    stop["msg"] = str(e)
                    x["rag"] = {"status": "ditunda", "catatan": "kuota token harian LLM habis; jalankan ulang nanti (hasil sebelumnya ada di cache)"}
                except Exception as e:  # noqa: BLE001
                    x["rag"] = {"status": "error", "error": str(e), "raw": getattr(e, "raw", "")}
            j = x["rag"].get("jendela")
            print(f"    [RAG]    {nama}: {x['rag']['status']} {x['rag'].get('wajib_terpenuhi', '')}" + (f" (jendela {j['dibaca']}/{j['total']})" if j and j["total"] > 1 else ""))
        if a.mode in ("manual", "both"):
            try:
                res = judge_subbab(decomp, x["title"], x["text"], "")
                spec = get_criteria(x["title"])
                if spec:
                    ctx = {k: decomp.get(k, "") for k in ("problem_domain", "problem", "uniqueness")}
                    desc = {k["id"]: k["deskripsi"].format(**ctx) for k in spec["kriteria"]}
                    for h in res.get("hasil", []):
                        h["deskripsi"] = desc.get(h["id"], "")
                x["manual"] = res
            except Exception as e:  # noqa: BLE001
                x["manual"] = {"status": "error", "error": str(e)}
            print(f"    [manual] {nama}: {x['manual']['status']} {x['manual'].get('wajib_terpenuhi', '')}")
    if use_rag:
        save(run_dir, "05_judge_rag.json", [{"subbab": f"{x['num']} {x['title']}", **(x.get("rag") or {})} for x in subs])
    if a.mode in ("manual", "both"):
        save(run_dir, "05_judge_manual.json", [{"subbab": f"{x['num']} {x['title']}", **(x.get("manual") or {})} for x in subs])

    # ---- Aturan penulisan per kalimat (hanya mode kalimat; hanya subbab Bab 1 yang punya aturan), digabung dengan hasil panduan
    if a.mode == "kalimat":
        header(6, "aturan penulisan per kalimat (LLM melabeli tiap kalimat) dan status akhir", a.pause)
        for x in subs:
            if x.get("skip") or not has_aturan(x["title"]) or x["bab"] != "1":
                x["kalimat_judge"] = {"status": "tidak_ada_aturan", "aturan": [], "kalimat": []}
            elif stop["msg"]:
                x["kalimat_judge"] = {"status": "ditunda", "aturan": [], "kalimat": []}
            else:
                try:
                    x["kalimat_judge"] = judge_subbab_kalimat(decomp, x["title"], x["text"])
                except DailyLimitError as e:
                    stop["msg"] = str(e)
                    x["kalimat_judge"] = {"status": "ditunda", "aturan": [], "kalimat": []}
                except Exception as e:  # noqa: BLE001
                    x["kalimat_judge"] = {"status": "error", "error": str(e), "aturan": [], "kalimat": []}
            r = x["kalimat_judge"]
            selaraskan_dengan_label(x["title"], x.get("rag"), r)  # kriteria PD/Problem/Solusi ikut hasil label kalimat
            x["aturan_rows"] = aturan_sebagai_kriteria(r)
            x["akhir"] = gabung_status(x.get("rag"), r)
            print(f"    {x['num']} {x['title']}: panduan={(x.get('rag') or {}).get('status', '-')} "
                  f"aturan={r['status']} -> akhir={x['akhir']['status']}" + (f"  [{x['akhir']['catatan']}]" if x["akhir"]["catatan"] else ""))
            for h in r.get("aturan", []):
                print(f"        {h['unsur']:8} ({h['aturan']}): {h['status']}" + (f" -> {len(h['kalimat'])} kalimat" if h["status"] == "SALAH" else "") + (f"  [{h['catatan']}]" if h.get("catatan") else ""))
        save(run_dir, "05_judge_kalimat.json", [{"subbab": f"{x['num']} {x['title']}", **x["kalimat_judge"]} for x in subs])
    else:
        for x in subs:
            x["akhir"] = gabung_status(x.get("rag"), None)

    # ---- Kelengkapan struktur: petunjuk panduan yang tidak punya subbab padanan di dokumen
    struktur = None
    if use_rag and not a.only:
        struktur = cek_struktur(snap, subs, bab_dinilai)
        if stop["msg"]:  # pemilihan chunk belum selesai untuk semua subbab, jadi "tidak ditemukan" belum bisa dipastikan
            for r in struktur:
                r["pasti"] = False
        save(run_dir, "06_struktur.json", struktur)
        hilang = [f"{r['num']} {r['title']}" for r in struktur if not r["ditemukan"]]
        print(f"    Kelengkapan struktur: {len(struktur) - len(hilang)}/{len(struktur)} petunjuk panduan punya subbab padanan"
              + (f"; tidak ditemukan: {', '.join(hilang)}" if hilang else ""))

    # ---- Laporan
    header(7 if a.mode == "kalimat" else 6, "laporan", a.pause)
    data = {"meta": {"bab": ", ".join(bab_dinilai), "judul": a.judul, "waktu": datetime.now().strftime("%Y-%m-%d %H:%M"), "mode": a.mode,
                     "retrieval": a.retrieval, "topk": a.topk, "sections_file": sumber, "match": a.match, "limit": stop["msg"]},
            "decomp": decomp, "subbab": subs, "struktur": struktur, "visual": visual}
    save(run_dir, "summary.json", [{"subbab": f"{x['num']} {x['title']}", "bab": x["bab"], "rag": (x.get("rag") or {}).get("status"),
                                    "manual": (x.get("manual") or {}).get("status"),
                                    "aturan": (x.get("kalimat_judge") or {}).get("status"),
                                    "akhir": (x.get("akhir") or {}).get("status")} for x in subs])
    (run_dir / "report.html").write_text(build_report(data), encoding="utf-8")
    print(f"    Buka: {run_dir / 'report.html'}")
    if stop["msg"]:
        n = sum(1 for x in subs if (x.get("akhir") or {}).get("status") == "ditunda")
        print(f"\n!! Kuota token harian LLM habis: {n} subbab ditunda. Jalankan lagi perintah yang sama setelah kuota reset;\n"
              f"   panggilan yang sudah selesai diambil dari cache (debug_output/llm_cache.jsonl), jadi hanya sisanya yang memakai token.")
    return data


if __name__ == "__main__":
    main()
