"""Buat report.html: tampilan per langkah dari satu kali jalan pipeline."""
import html

esc = lambda s: html.escape(str(s if s is not None else ""))

CSS = """
:root{--bg:#fff;--fg:#1f2328;--mut:#656d76;--line:#d0d7de;--card:#f6f8fa;--ok:#1a7f37;--mid:#9a6700;--bad:#cf222e;--na:#6e7781}
@media (prefers-color-scheme:dark){:root{--bg:#0d1117;--fg:#e6edf3;--mut:#8b949e;--line:#30363d;--card:#161b22;--ok:#3fb950;--mid:#d29922;--bad:#f85149;--na:#8b949e}}
body{font:14px/1.5 system-ui,Segoe UI,sans-serif;background:var(--bg);color:var(--fg);max-width:1100px;margin:0 auto;padding:24px}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:32px 0 8px;border-bottom:1px solid var(--line);padding-bottom:4px}
h3{font-size:15px;margin:18px 0 6px}.mut{color:var(--mut)}table{border-collapse:collapse;width:100%;margin:6px 0}
th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}th{background:var(--card)}
details{border:1px solid var(--line);border-radius:6px;margin:6px 0;padding:6px 10px;background:var(--card)}
summary{cursor:pointer;font-weight:600}pre{white-space:pre-wrap;word-break:break-word;margin:6px 0;font:12px/1.5 ui-monospace,Consolas,monospace}
.b{display:inline-block;border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600;border:1px solid currentColor}
.sesuai,.yes{color:var(--ok)}.sebagian{color:var(--mid)}.tidak_sesuai,.no{color:var(--bad)}
.tidak_ada_padanan,.ditunda,.tidak_ditemukan,.tidak_dinilai,.tidak_ada_kriteria,.tidak_ada_aturan,.kosong,.error{color:var(--na)}
.row{display:flex;gap:16px;flex-wrap:wrap}.row>div{flex:1;min-width:320px}
.big{font-size:30px;font-weight:700}
.babhead{font-size:16px;font-weight:700;margin:34px 0 4px;padding:4px 10px;background:var(--card);border-left:4px solid var(--mut)}
.warn{border:1px solid var(--bad);border-radius:6px;padding:8px 12px;margin:10px 0}
"""


def badge(status):
    return f'<span class="b {esc(status)}">{esc(status)}</span>'


def _crit_table(rows):
    out = ["<table><tr><th>ID</th><th>Kriteria</th><th>Wajib</th><th>Hasil</th><th>Bukti</th><th>Alasan</th></tr>"]
    for r in rows:
        dasar = ""
        if "sumber" in r:
            ok = r.get("sumber_valid", True)
            dasar = (
                f'<div class="mut">dasar panduan: “{esc(r["sumber"])}” '
                f'{"" if ok else "<b class=no>(tidak valid, tidak dihitung)</b>"}</div>'
            )
        res = '<b class="yes">✓</b>' if r.get("terpenuhi") else '<b class="no">✗</b>'
        out.append(
            f"<tr><td>{esc(r.get('id'))}</td><td>{esc(r.get('deskripsi'))}{dasar}</td>"
            f"<td>{'wajib' if r.get('wajib') else 'anjuran'}</td><td>{res}</td>"
            f"<td>{esc(r.get('bukti'))}</td><td>{esc(r.get('alasan'))}</td></tr>"
        )
    out.append("</table>")
    return "".join(out)


def _simple_table(rows):
    """Tabel hasil akhir: hanya ID, kriteria, dan benar/salah. Tanpa wajib, kutipan, atau alasan."""
    out = ["<table><tr><th>ID</th><th>Kriteria</th><th style='width:70px'>Hasil</th></tr>"]
    for r in rows:
        res = ('<b class="mut" title="tidak dijawab LLM">?</b>' if r.get("dinilai", True) is False else
               '<b class="yes">✓</b>' if r.get("terpenuhi") else '<b class="no">✗</b>')
        out.append(f"<tr><td>{esc(r.get('id'))}</td><td>{esc(r.get('sumber') or r.get('deskripsi'))}</td><td>{res}</td></tr>")
    out.append("</table>")
    return "".join(out)


def _result_block(title, res):
    if res is None:
        return ""
    if res.get("error"):
        return f"<h3>{esc(title)} {badge('error')}</h3><pre>{esc(res['error'])}</pre>"
    if "kriteria" not in res and "hasil" not in res:
        return f"<h3>{esc(title)} {badge(res.get('status', ''))}</h3><p class=mut>{esc(res.get('alasan', ''))}</p>"
    rows = res.get("kriteria") or res.get("hasil") or []
    head = (
        f"<h3>{esc(title)} {badge(res.get('status'))} "
        f"<span class=mut>wajib {esc(res.get('wajib_terpenuhi'))}"
        + (f" · anjuran {esc(res['anjuran_terpenuhi'])}" if res.get("anjuran_terpenuhi") else "") + "</span></h3>"
    )
    saran = f"<p><b>Saran:</b> {esc(res['saran'])}</p>" if res.get("saran") else ""
    trace = ""
    if res.get("trace"):
        t = res["trace"]
        trace = (
            "<details><summary>Jejak LLM (prompt dan jawaban mentah)</summary>"
            f"<pre><b>SYSTEM</b>\n{esc(t['system'])}\n\n<b>USER</b>\n{esc(t['user'])}\n\n<b>JAWABAN</b>\n{esc(t['raw'])}</pre></details>"
        )
    tabel = _simple_table(rows) if rows and "sumber" in rows[0] else _crit_table(rows)  # skema RAG baru vs mode manual
    return head + tabel + saran + trace


def _kalimat_table(rows):
    out = ["<table><tr><th>#</th><th>Kalimat</th><th>Label</th><th>Alasan</th></tr>"]
    for k in rows:
        lab = esc(k.get("label"))
        if k.get("label_awal"):
            lab += f" <span class=mut>(semula {esc(k['label_awal'])})</span>"
        alasan = esc(k.get("alasan"))
        if k.get("dibebaskan"):
            alasan += f" <span class=yes>[dibebaskan: {esc(k.get('alasan_verifikasi'))}]</span>"
        elif k.get("alasan_verifikasi"):
            alasan += f" <span class=no>[verifikasi: melanggar, {esc(k['alasan_verifikasi'])}]</span>"
        if k.get("verifikasi_gagal"):
            alasan += " <span class=mut>[belum terverifikasi]</span>"
        out.append(f"<tr><td>{esc(k.get('n'))}</td><td>{esc(k.get('kalimat'))}</td><td>{lab}</td><td>{alasan}</td></tr>")
    out.append("</table>")
    return "".join(out)


def _gabungan_block(s):
    """Mode kalimat: satu tabel (kriteria panduan wajib + aturan penulisan) dan status akhir.
    Label per kalimat dan jejak LLM tetap ada, tapi dilipat, hanya untuk debugging."""
    rag = s.get("rag") or {}
    kal = s.get("kalimat_judge") or {}
    akhir = s.get("akhir") or {}
    out = [f"<h3>Status akhir {badge(akhir.get('status', '-'))}</h3>"]
    if akhir.get("catatan"):
        out.append(f"<p class=mut>{esc(akhir['catatan'])}</p>")
    for nama, r in (("Penilaian panduan", rag), ("Aturan penulisan", kal)):
        if r.get("error"):
            out.append(f"<p class=no><b>{nama} error:</b></p><pre>{esc(r['error'])}</pre>")
            if r.get("raw"):
                out.append(f"<details><summary>Debug: jawaban mentah LLM yang gagal dibaca</summary><pre>{esc(r['raw'])}</pre></details>")
    if rag.get("status") in ("tidak_ada_padanan", "kosong", "ditunda", "tidak_dinilai") or rag.get("catatan"):
        out.append(f"<p class=mut>Panduan: {esc(rag.get('catatan', ''))}</p>")
    j = rag.get("jendela") or {}
    if j.get("total", 1) > 1:
        sisa = j["total"] - j["dibaca"]
        out.append(f"<p class=mut>Teks panjang: dibaca per bagian, {j['dibaca']} dari {j['total']} bagian"
                   + (" (berhenti karena semua kriteria sudah terpenuhi)" if sisa and rag.get("status") == "sesuai" else
                      f"; {sisa} bagian tidak dibaca (batas --max-windows)" if sisa else "") + ".</p>")
    rows = (rag.get("kriteria") or []) + (s.get("aturan_rows") or [])
    rows = [r for r in rows if r.get("sumber_valid", True)]  # kriteria yang sumbernya tidak valid tidak ditampilkan
    if rows:
        out.append(_simple_table(rows))
    if rag.get("diabaikan"):
        dl = "\n".join(f"[{d['alasan']}] {d['kalimat']}" for d in rag["diabaikan"])
        out.append(f"<details><summary>Kalimat petunjuk yang tidak dijadikan kriteria ({len(rag['diabaikan'])})</summary><pre>{esc(dl)}</pre></details>")
    if kal.get("kalimat"):
        h = kal.get("hitung") or {}
        out.append(f"<details><summary>Debug: label per kalimat <span class=mut>({', '.join(f'{k} {v}' for k, v in h.items())})</span></summary>"
                   f"{_kalimat_table(kal['kalimat'])}</details>")
    if rag.get("kriteria"):
        dbg = "\n".join(f"{k['id']} [{'✓' if k['terpenuhi'] else '✗'}] {k.get('dinilai_oleh', 'llm')}"
                        f"{' (kutipan tidak valid)' if k.get('kutipan_tidak_valid') else ''}: {k.get('kutipan', '')}"
                        for k in rag["kriteria"])
        out.append(f"<details><summary>Debug: kutipan kriteria panduan</summary><pre>{esc(dbg)}</pre></details>")
    if rag.get("trace"):
        t = rag["trace"]
        out.append("<details><summary>Debug: jejak LLM penilaian panduan</summary>"
                   f"<pre><b>SYSTEM</b>\n{esc(t['system'])}\n\n<b>USER</b>\n{esc(t['user'])}\n\n<b>JAWABAN</b>\n{esc(t['raw'])}</pre></details>")
    return "".join(out)


def _skor_class(skor):
    return "sesuai" if skor >= 90 else "sebagian" if skor >= 70 else "tidak_sesuai"


def visual_section(v):
    """Bagian 'Pemeriksaan format (visual)': skor keseluruhan, skor per bab, dan temuan per bab."""
    if not v:
        return ""
    out = ['<h2 id="format">Pemeriksaan format (visual)</h2>']
    if v.get("error"):
        return out[0] + f"<p class=no><b>Pemeriksaan format gagal dijalankan:</b></p><pre>{esc(v['error'])}</pre>"
    sc, findings = v["score"], v["findings"]
    out.append(
        f"<p><span class=\"big {_skor_class(sc['overall_score'])}\">{esc(sc['overall_score'])}</span> "
        f"<span class=mut>/ 100 · {esc(sc['total_findings'])} temuan · berkas: {esc(v.get('file'))}</span></p>"
    )
    for e in v.get("errors", []):
        out.append(f"<p class=no><b>Tahap {esc(e['tahap'])} gagal</b> (hasilnya tidak ikut dihitung): {esc(e['error'])}</p>")
    out.append("<table><tr><th>Bab</th><th>Diperiksa</th><th>Temuan</th><th>Skor</th></tr>")
    for ch, c in sc["per_chapter"].items():
        out.append(f"<tr><td>{esc(ch)}</td><td>{c['checked']}</td><td>{c['failed']}</td>"
                   f"<td><span class=\"b {_skor_class(c['score'])}\">{c['score']}</span></td></tr>")
    out.append("</table>")
    by_ch = {}
    for f in findings:
        by_ch.setdefault(f.get("chapter") or "Lainnya", []).append(f)
    if by_ch:
        out.append("<h3>Temuan per bab</h3>")
    for ch, fs in by_ch.items():
        rows = []
        for f in fs:
            low = ' <span class="b sebagian">keyakinan rendah</span>' if f.get("confidence") == "rendah" else ""
            rows.append(f"<tr><td>{esc(f.get('category'))}</td><td>{esc(f.get('location'))}</td><td>{esc(f.get('description'))}{low}</td></tr>")
        out.append(f"<details><summary>{esc(ch)} <span class=mut>({len(fs)} temuan)</span></summary>"
                   f"<table><tr><th style='width:80px'>Kategori</th><th style='width:30%'>Lokasi</th><th>Keterangan</th></tr>{''.join(rows)}</table></details>")
    return "".join(out)


def build_visual_only(v, nama_berkas, catatan=""):
    """Laporan kalau tahap sistematika berhenti di tengah jalan: hasil format tetap bisa dilihat."""
    p = [f"<!doctype html><html lang=id><meta charset=utf-8><title>Hasil pemeriksaan</title><style>{CSS}</style><body>",
         f"<h1>Pemeriksaan dokumen</h1><div class=mut>{esc(nama_berkas)}</div>"]
    if catatan:
        p.append(f"<p class=no><b>Tahap sistematika tidak selesai:</b> {esc(catatan)}</p>")
    p.append(visual_section(v))
    p.append("</body></html>")
    return "".join(p)


def _per_bab(subs):
    """Tabel ringkas: jumlah subbab per bab menurut status akhir."""
    order, by = [], {}
    for s in subs:
        if s["bab"] not in by:
            by[s["bab"]] = []
            order.append(s["bab"])
        by[s["bab"]].append((s.get("akhir") or {}).get("status", "-"))
    cols = ["sesuai", "tidak_sesuai", "tidak_ada_padanan", "tidak_dinilai", "kosong", "ditunda", "error"]
    out = ["<table><tr><th>Bab</th><th>Subbab</th>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"]
    for b in order:
        st = by[b]
        out.append(f"<tr><td>BAB {esc(b)}</td><td>{len(st)}</td>" + "".join(f"<td>{st.count(c) or '-'}</td>" for c in cols) + "</tr>")
    out.append("</table>")
    return "".join(out)


def _struktur(rows):
    """Petunjuk panduan yang tidak punya subbab padanan di dokumen mahasiswa."""
    if not rows:
        return ""
    hilang = [r for r in rows if not r["ditemukan"]]
    pasti = all(r.get("pasti", True) for r in rows)
    out = [f"<h2>Kelengkapan struktur</h2><p class=mut>{len(rows) - len(hilang)} dari {len(rows)} petunjuk panduan di bab yang diperiksa "
           f"punya subbab padanan di dokumen.{'' if pasti else ' <b class=no>Belum pasti: pemilihan chunk terhenti karena kuota token.</b>'}</p>"]
    if hilang:
        out.append("<table><tr><th>Petunjuk panduan</th><th>Status</th></tr>")
        for r in hilang:
            out.append(f"<tr><td>{esc(r['num'])} {esc(r['title'])}</td><td>{badge('tidak_ditemukan')}</td></tr>")
        out.append("</table><p class=mut>Daftar ini hanya memberi tahu; belum ada keputusan mana yang wajib ada.</p>")
    return "".join(out)


def build_report(data: dict) -> str:
    m, d, subs = data["meta"], data["decomp"], data["subbab"]
    p = [f"<!doctype html><html lang=id><meta charset=utf-8><title>Hasil pemeriksaan sistematika</title><style>{CSS}</style><body>"]
    p.append(f"<h1>Pemeriksaan sistematika (Bab {esc(m['bab'])})</h1>")
    if m.get("limit"):
        p.append(f"<div class=warn><b>Kuota token harian LLM habis di tengah pemeriksaan.</b> Subbab bertanda <i>ditunda</i> belum dinilai. "
                 f"Jalankan lagi perintah yang sama setelah kuota reset; yang sudah selesai diambil dari cache.<br><span class=mut>{esc(str(m['limit'])[:200])}</span></div>")
    p.append(
        f"<div class=mut>{esc(m['judul'])}<br>{esc(m['waktu'])} · mode penilaian: {esc(m['mode'])} · "
        f"retrieval: {esc(m['retrieval'])} (top-{esc(m['topk'])}) · pemilihan chunk: {esc(m.get('match', 'auto'))} · berkas: {esc(m['sections_file'])}</div>"
    )

    if data.get("visual"):
        p.append('<p><a href="#format">Format (visual)</a> · <a href="#sistematika">Sistematika</a></p>')
        p.append(visual_section(data["visual"]))
        p.append('<h2 id="sistematika" style="margin-top:44px">Pemeriksaan sistematika (Bab ' + esc(m['bab']) + ')</h2>')

    # Ringkasan
    kalimat_mode = m["mode"] == "kalimat"
    p.append("<h2>Ringkasan per bab</h2>" + _per_bab(subs))
    p.append(_struktur(data.get("struktur")))
    if kalimat_mode:
        p.append("<h2>Ringkasan per subbab</h2><table><tr><th>Subbab</th><th>Panduan terpilih</th><th>Status akhir</th></tr>")
        for s in subs:
            rag = s.get("rag") or {}
            kal = s.get("kalimat_judge") or {}
            p.append(
                f"<tr><td>{esc(s['num'])} {esc(s['title'])}</td><td>{esc((s.get('match') or {}).get('pilihan_llm') or '-')}</td>"
                f"<td>{badge((s.get('akhir') or {}).get('status', '-'))}</td></tr>"
            )
        p.append("</table>")
    else:
        p.append("<h2>Ringkasan per subbab</h2><table><tr><th>Subbab</th><th>Panduan terpilih</th><th>Mode RAG</th><th>Mode kriteria manual</th></tr>")
        for s in subs:
            rag = s.get("rag") or {}
            man = s.get("manual") or {}
            p.append(
                f"<tr><td>{esc(s['num'])} {esc(s['title'])}</td><td>{esc((s.get('match') or {}).get('pilihan_llm') or '-')}</td>"
                f"<td>{badge(rag.get('status', '-')) if rag else '-'}</td><td>{badge(man.get('status', '-')) if man else '-'}</td></tr>"
            )
        p.append("</table>")

    # Langkah 1
    p.append("<h2>Langkah 1. Pemecahan judul</h2>")
    p.append(
        f"<table><tr><th>Problem Domain</th><td>{esc(d.get('problem_domain'))}</td></tr>"
        f"<tr><th>Problem</th><td>{esc(d.get('problem'))}</td></tr>"
        f"<tr><th>Uniqueness</th><td>{esc(d.get('uniqueness'))}</td></tr></table>"
    )
    if d.get("perlu_dicek"):
        p.append(f"<p class=no><b>Perlu dicek manual:</b> {esc('; '.join(d['perlu_dicek']))}</p>")
    if d.get("catatan"):
        p.append(f"<p class=mut>Catatan: {esc(d['catatan'])}</p>")

    # Langkah 2
    p.append("<h2>Langkah 2. Subbab dari dokumen mahasiswa</h2>")
    for s in subs:
        p.append(
            f"<details><summary>{esc(s['num'])} {esc(s['title'])} <span class=mut>({s['chars']} karakter)</span></summary>"
            f"<pre>{esc(s['text'][:3000])}</pre>{'<p class=mut>(dipotong 3000 karakter pertama untuk tampilan; seluruh teks tetap dinilai)</p>' if s['chars'] > 3000 else ''}</details>"
        )

    # Langkah 3
    p.append("<h2>Langkah 3. Retrieval: kandidat chunk panduan</h2><p class=mut>Query = judul subbab. Jarak lebih kecil = lebih mirip.</p>")
    for s in subs:
        p.append(f"<h3>{esc(s['num'])} {esc(s['title'])}</h3>")
        if s.get("retrieval_error"):
            p.append(f"<pre>{esc(s['retrieval_error'])}</pre>")
            continue
        p.append("<table><tr><th>Rank</th><th>Chunk</th><th>Judul panduan</th><th>Jarak</th><th>Asal</th></tr>")
        for c in s.get("candidates", []):
            p.append(f"<tr><td>{c['rank']}</td><td>{esc(c['chunk_id'])}</td><td>{esc(c['subbab_title'])}</td><td>{esc(c['distance'])}</td><td class=mut>{esc(c.get('sumber', 'vektor'))}</td></tr>")
        p.append("</table>")
        for c in s.get("candidates", [])[:1]:
            p.append(f"<details><summary>Isi chunk peringkat 1</summary><pre>{esc(c['text'])}</pre></details>")

    # Langkah 4
    p.append("<h2>Langkah 4. LLM memilih chunk panduan</h2>")
    p.append("<table><tr><th>Subbab</th><th>Pilihan LLM</th><th>Pilihan aturan (judul sama)</th><th>Sama?</th><th>Cara / alasan</th></tr>")
    for s in subs:
        mt = s.get("match")
        if not mt:
            p.append(f"<tr><td>{esc(s['title'])}</td><td colspan=4 class=mut>{esc(s.get('match_error') or '-')}</td></tr>")
            continue
        same = mt["pilihan_llm"] == mt["pilihan_aturan"]
        p.append(
            f"<tr><td>{esc(s['title'])}</td><td>{esc(mt['pilihan_llm'] or 'tidak ada padanan')}</td>"
            f"<td>{esc(mt['pilihan_aturan'] or '-')}</td><td>{'<b class=yes>ya</b>' if same else '<b class=no>beda</b>'}</td>"
            f"<td>{esc(mt.get('dari', 'llm'))}: {esc(mt['alasan'])}</td></tr>"
        )
    p.append("</table>")

    # Langkah 5
    p.append("<h2>Langkah 5. Penilaian</h2>")
    bab_now = None
    for s in subs:
        if s["bab"] != bab_now:
            bab_now = s["bab"]
            p.append(f"<div class=babhead>BAB {esc(bab_now)}</div>")
        p.append(f"<h3 style='margin-top:26px'>{esc(s['num'])} {esc(s['title'])}</h3>")
        if kalimat_mode:
            p.append(_gabungan_block(s))
            continue
        blocks = []
        if s.get("rag") is not None:
            blocks.append(_result_block("Mode RAG (kriteria diturunkan dari chunk panduan)", s["rag"]))
        if s.get("manual") is not None:
            blocks.append(_result_block("Mode kriteria manual", s["manual"]))
        p.append("".join(f"<div>{b}</div>" for b in blocks) if len(blocks) < 2 else '<div class="row">' + "".join(f"<div>{b}</div>" for b in blocks) + "</div>")
    p.append("</body></html>")
    return "".join(p)
