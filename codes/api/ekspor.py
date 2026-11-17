"""Ekspor dokumen asli yang sudah diberi sorotan dan komentar dari hasil pemeriksaan.

- PDF  : PDF asli (atau hasil konversi LibreOffice untuk .docx) + anotasi sorotan kuning dan komentar
         pada posisi yang dihitung `api/preview.hitung_sorotan`.
- DOCX : berkas .docx asli + sorotan kuning pada teks bermasalah dan komentar Word (butuh python-docx >= 1.2).

Isi komentar tiap temuan: "<no>. <jenis temuan> (<bab>)", penjelasan, dan acuan pedoman bila ada.
"""
import copy
import re
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:  # PyMuPDF versi lama
    import fitz  # type: ignore

KUNING = (1.0, 0.85, 0.1)
PENULIS = "PENScek"


def _bagian(bab, k, f):
    """(judul, uraian, acuan) satu temuan. Nomor temuan sama dengan nomor di daftar temuan tiap bab pada halaman hasil."""
    judul = f"{bab} \u00b7 Temuan {k + 1}: {f.get('t') or 'Temuan'}"
    return judul, (f.get("d") or "").strip(), (f"Acuan: {f['r']}" if f.get("r") else "")


def _komentar(bab, k, f):
    return "\n".join(b for b in _bagian(bab, k, f) if b)


def _gabung(teks):
    """Beberapa temuan pada satu tempat: dipisah satu baris kosong."""
    return "\n\n".join(teks)


def _semua_temuan(hasil):
    """(indeks bab, indeks temuan, nama bab, temuan, teks komentar)"""
    for i, b in enumerate(hasil.get("bab", [])):
        for j, f in enumerate(b.get("f", [])):
            yield i, j, b.get("n", ""), f, _komentar(b.get("n", ""), j, f)


# ---------------------------------------------------------------- PDF

def _popup_lebar(a, page, teks, lebar=300):
    """Popup komentar di PDF viewer berukuran tetap kecil, sehingga komentar panjang (beberapa temuan sekaligus)
    terpotong. Ukuran popup disesuaikan dengan panjang teks dan dijaga tetap di dalam halaman."""
    baris = sum(max(1, -(-len(b) // 48)) for b in teks.split("\n"))
    tinggi = min(page.rect.height - 20, baris * 13 + 30)
    x0 = min(max(10, a.rect.x1 + 6), page.rect.width - lebar - 10)
    y0 = max(10, min(a.rect.y0, page.rect.height - tinggi - 10))
    try:
        a.set_popup(fitz.Rect(x0, y0, x0 + lebar, y0 + tinggi))
    except Exception:  # noqa: BLE001  (PyMuPDF lama tanpa set_popup: popup tetap ukuran bawaan)
        pass


def _catatan(page, titik, teks, judul=None):
    """Ikon catatan di halaman (untuk temuan yang tidak punya teks untuk disorot)."""
    a = page.add_text_annot(titik, teks, icon="Comment")
    a.set_colors(stroke=KUNING)
    a.set_info(title=judul or PENULIS)
    a.update()
    _popup_lebar(a, page, teks)
    return a


def pdf_beranotasi(pdf_path, hasil, sorotan, tujuan):
    """Salin PDF dengan anotasi sorotan + komentar. `sorotan` = keluaran hitung_sorotan."""
    doc = fitz.open(pdf_path)
    try:
        belum = []  # temuan yang posisinya tidak ketemu: satu catatan di halaman pertama
        per_halaman = {}  # temuan yang hanya diketahui halamannya: satu catatan per halaman
        # Beberapa temuan bisa menunjuk kotak yang sama persis (mis. posisi, jarak, dan penomoran satu persamaan).
        # Kalau dibuat satu anotasi per temuan, anotasi-anotasi itu bertumpuk dan PDF viewer hanya menampilkan
        # popup yang paling atas. Jadi temuan dengan kotak identik digabung menjadi satu sorotan dengan satu komentar.
        kelompok = {}
        for i, j, _bab, f, teks in _semua_temuan(hasil):
            s = sorotan.get("temuan", {}).get(f"{i}-{j}")
            rects = (s or {}).get("rects") or []
            if rects:
                kunci = tuple((r["p"], tuple(r["r"])) for r in rects)
                kelompok.setdefault(kunci, {"rects": rects, "teks": []})["teks"].append(teks)
            elif s and s.get("hal") and 1 <= s["hal"] <= len(doc):
                per_halaman.setdefault(s["hal"], []).append(teks)
            else:
                belum.append(teks)
        for g in kelompok.values():
            komentar = _gabung(g["teks"])
            for n, r in enumerate(g["rects"]):
                page = doc[r["p"] - 1]
                w, h = page.rect.width, page.rect.height
                x0, y0, x1, y1 = r["r"]
                a = page.add_highlight_annot(fitz.Rect(x0 * w, y0 * h, x1 * w, y1 * h))
                a.set_colors(stroke=KUNING)
                a.set_info(title=PENULIS, content=komentar if n == 0 else "")
                a.update()
                if n == 0:
                    _popup_lebar(a, page, komentar)
        for hal, daftar in sorted(per_halaman.items()):
            page = doc[hal - 1]
            _catatan(page, fitz.Point(page.rect.width - 40, 30), _gabung(daftar))
        if belum and len(doc):
            page = doc[0]
            _catatan(page, fitz.Point(page.rect.width - 40, 30 if 1 not in per_halaman else 56),
                     "Temuan yang posisinya di halaman tidak ditemukan\n\n" + _gabung(belum))
        Path(tujuan).parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(tujuan), garbage=3, deflate=True)
    finally:
        doc.close()
    return Path(tujuan)


# ---------------------------------------------------------------- DOCX

def _padat(s):
    """Huruf kecil tanpa spasi: -> (teks padat, indeks asli tiap karakter)."""
    teks, idx = [], []
    for n, c in enumerate(s):
        if not c.isspace() and c != "\u00ad":
            teks.append(c.lower())
            idx.append(n)
    return "".join(teks), idx


class _Par:
    """Satu paragraf: teks gabungan run langsung (tanpa hyperlink), supaya indeks karakter cocok dengan run."""

    def __init__(self, p):
        self.p = p
        self.padat, self.idx = _padat("".join(r.text for r in p.runs))


def _pisah(run, k):
    """Belah run di karakter ke-k; run kiri tetap, run kanan baru. Run bertab/baris-baru tidak dibelah."""
    from docx.text.run import Run

    teks = run.text
    if not (0 < k < len(teks)) or "\t" in teks or "\n" in teks or "\r" in teks:
        return None
    baru = copy.deepcopy(run._r)
    run._r.addnext(baru)
    kanan = Run(baru, run._parent)
    run.text = teks[:k]
    kanan.text = teks[k:]
    return kanan


def _runs_rentang(par, a, b):
    """Run yang menutup karakter [a, b) paragraf, dengan membelah run di tepi rentang."""
    for batas in (a, b):
        pos = 0
        for r in par.runs:
            n = len(r.text)
            if pos < batas < pos + n:
                _pisah(r, batas - pos)
                break
            pos += n
    hasil, pos = [], 0
    for r in par.runs:
        n = len(r.text)
        if n and pos < b and pos + n > a:
            hasil.append(r)
        pos += n
    return hasil


def _cari_kalimat(pars, q, dari):
    """-> (indeks paragraf, mulai, akhir) pada teks asli paragraf, atau None."""
    kunci, _ = _padat(q)
    if not kunci:
        return None
    urutan = list(range(dari, len(pars))) + list(range(0, dari))
    for panjang in (len(kunci), min(len(kunci), 40)):  # utuh dulu; kalau kalimat melintasi paragraf, cocokkan awalnya
        awal = kunci[:panjang]
        for n in urutan:
            p = pars[n]
            i = p.padat.find(awal)
            if i >= 0:
                j = min(i + len(kunci), len(p.padat)) - 1
                return n, p.idx[i], p.idx[j] + 1
    return None


def _cari_judul(pars, kandidat, dari):
    """Paragraf yang teksnya diawali salah satu `kandidat` (urut prioritas), dicari mulai dari `dari`."""
    urutan = list(range(dari, len(pars))) + list(range(0, dari))
    for k in kandidat:
        kp, _ = _padat(k)
        if not kp:
            continue
        for n in urutan:
            pt = pars[n].padat
            if pt == kp or (pt.startswith(kp) and len(pt) <= len(kp) + 6):
                return n
    return None


def _kandidat_judul(f, bab):
    lokasi = f.get("lokasi") or ""
    kand = []
    m = re.search(r"\('(.+?)'\)", lokasi)  # Paragraf #163 ('1.1 LATAR BELAKANG')
    if m:
        kand.append(m.group(1).rstrip("…. "))
    m = re.match(r"Subbab\s+(\d+(?:\.\d+)*)\s+(.+)$", lokasi)  # Subbab 1.3 TUJUAN
    if m:
        kand += [f"{m.group(1)} {m.group(2)}", m.group(2)]
    if "(" in lokasi and not lokasi.startswith(("Halaman", "Paragraf", "Section")):
        kand.append(lokasi.split("(")[0].strip())  # Abstrak (Paragraf #49-#55)
    return kand


_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _paragraf_persamaan(pars, nomor):
    """{nomor paragraf pemeriksa: indeks di pars} untuk paragraf persamaan (yang memuat objek Equation Word).
    Nomor paragraf pemeriksa dicocokkan langsung (indeks 0 atau 1 based) kalau semua nomor jatuh pada paragraf
    persamaan; kalau tidak, dicocokkan berurutan hanya bila jumlah persamaan dokumen sama dengan jumlah nomor."""
    math = [n for n, p in enumerate(pars) if p.p._p.find(f".//{{{_NS_M}}}oMath") is not None]
    if not nomor or not math:
        return {}
    ada = set(math)
    for geser in (0, -1):
        if all((no + geser) in ada for no in nomor):
            return {no: no + geser for no in nomor}
    if len(nomor) == len(math):
        return dict(zip(sorted(nomor), math))
    return {}


def _sorot_persamaan(par):
    """Sorot isi persamaan (run matematika) dan run biasa di paragrafnya (mis. nomor persamaan)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    for mr in par._p.iter(f"{{{_NS_M}}}r"):
        rpr = mr.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            t = mr.find(f"{{{_NS_M}}}t")
            if t is not None:
                t.addprevious(rpr)
            else:
                mr.append(rpr)
        if rpr.find(qn("w:highlight")) is None:
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)


_POLA_PARAGRAF = re.compile(r"^Paragraf\s+#(\d+)\s*\((.*)\)\s*$")
_POLA_TABEL = re.compile(r"^Persamaan dalam tabel \(tabel ke-(\d+)")


def _cocok_paragraf(teks, kutipan):
    """Paragraf docx cocok dengan kutipan dari pemeriksa ('LATAR BELAKANG', atau awal kalimat yang dipotong '...')."""
    a, _ = _padat(teks)
    k = kutipan.strip()
    if len(k) >= 2 and k[0] == k[-1] and k[0] in "'\"":
        k = k[1:-1]
    b, _ = _padat(k.rstrip(".\u2026 "))
    return bool(a and b) and (a.startswith(b[:40]) or b.startswith(a[:40]))


def docx_beranotasi(sumber, hasil, tujuan):
    """Salin .docx dengan sorotan kuning dan komentar Word pada tiap temuan yang posisinya ketemu."""
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX
    from docx.text.paragraph import Paragraph

    if not hasattr(Document(str(sumber)), "add_comment"):
        raise RuntimeError("python-docx terlalu lama untuk komentar Word. Jalankan: pip install -U \"python-docx>=1.2\"")

    doc = Document(str(sumber))
    pars = [_Par(Paragraph(p, doc._body)) for p in doc.element.body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")]
    kosong = [n for n, p in enumerate(pars) if p.padat]
    # Nomor 'Paragraf #N' dari pemeriksa = urutan paragraf tingkat atas dokumen (doc.paragraphs, mulai 0), tanpa isi tabel.
    # Pencarian teks sering salah sasaran untuk judul pendek ('TUJUAN', 'PENELITIAN TERKAIT' yang juga muncul di daftar isi
    # atau kalimat lain), jadi nomor paragraf dipakai dulu dan dicocokkan dengan kutipannya.
    indeks_par = {p.p._p: n for n, p in enumerate(pars)}
    tingkat_atas = doc.paragraphs
    judul_tebal = {b[0] for _i, _j, bab, f, _t in _semua_temuan(hasil) for b in [_bagian(bab, _j, f)]}
    judul_tebal.add("Temuan yang tidak terkait satu paragraf")

    def sorot(runs):
        for r in runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    def komentari(runs, teks):
        c = doc.add_comment(runs=[runs[0], runs[-1]] if len(runs) > 1 else runs[0], text=teks, author=PENULIS, initials="PS")
        for para in c.paragraphs:  # judul tiap temuan ditebalkan supaya komentar yang berisi beberapa temuan mudah dibaca
            if para.text.strip() in judul_tebal:
                for r in para.runs:
                    if r.text.strip():
                        r.bold = True

    umum = []  # temuan tanpa posisi di teks (nomor halaman, bab yang kurang subbab, dst)
    dari_bab = {}
    pola_par = re.compile(r"^Paragraf\s+#(\d+)\s*\(persamaan\)")
    nomor_pers = sorted({int(m.group(1)) for _i, _j, _b, f, _t in _semua_temuan(hasil)
                         for m in [pola_par.match(f.get("lokasi") or "")] if m})
    peta_pers = _paragraf_persamaan(pars, nomor_pers)
    komentar_pers = {}  # indeks paragraf persamaan -> teks komentar (semua temuan satu persamaan digabung jadi satu)
    for i, j, _bab, f, teks in _semua_temuan(hasil):
        lokasi = f.get("lokasi") or ""
        m = pola_par.match(lokasi)
        if m:
            n = peta_pers.get(int(m.group(1)))
            if n is None:
                umum.append(teks)
            else:
                komentar_pers.setdefault(n, []).append(teks)
            continue
        mt = _POLA_TABEL.match(lokasi)  # persamaan Word di dalam tabel: 'tabel ke-N' = urutan di doc.tables
        if mt:
            nt = int(mt.group(1))
            n = None
            if 1 <= nt <= len(doc.tables):
                for pp in doc.tables[nt - 1]._tbl.iter("{%s}p" % _NS_W):
                    if pp.find(f".//{{{_NS_M}}}oMath") is not None and pp in indeks_par:
                        n = indeks_par[pp]
                        break
            if n is None:
                umum.append(teks)
            else:
                komentar_pers.setdefault(n, []).append(teks)
            continue
        dari = dari_bab.get(i, 0)
        runs = []
        for q in f.get("q") or []:
            ketemu = _cari_kalimat(pars, q, dari)
            if not ketemu:
                continue
            n, a, b = ketemu
            rr = _runs_rentang(pars[n].p, a, b)
            pars[n] = _Par(pars[n].p)  # teks run sudah berubah karena pembelahan
            if rr:
                sorot(rr)
                if not runs:
                    runs = rr
                dari = n
        if not runs:
            mp = _POLA_PARAGRAF.match(lokasi)
            if mp and int(mp.group(1)) < len(tingkat_atas):
                par = tingkat_atas[int(mp.group(1))]
                if par._p in indeks_par and _cocok_paragraf(par.text, mp.group(2)):
                    runs = par.runs
                    sorot(runs)
                    dari = indeks_par[par._p]
        if not runs:
            n = _cari_judul(pars, _kandidat_judul(f, _bab), dari)
            if n is not None:
                runs = pars[n].p.runs
                sorot(runs)
                dari = n
        dari_bab[i] = dari
        if runs:
            komentari(runs, teks)
        else:
            umum.append(teks)

    for n, daftar in komentar_pers.items():
        par = pars[n].p
        _sorot_persamaan(par)
        sorot(par.runs)
        runs = par.runs or [par.add_run("")]  # persamaan tanpa teks lain: jangkar komentar berupa run kosong
        komentari(runs, _gabung(daftar))

    if umum and kosong:
        # Satu komentar untuk semua temuan yang tidak punya letak di teks, menempel di paragraf pertama yang berisi teks.
        runs = pars[kosong[0]].p.runs
        if runs:
            komentari(runs[:1], "Temuan yang tidak terkait satu paragraf\n\n" + _gabung(umum))

    Path(tujuan).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(tujuan))
    return Path(tujuan)
