"""Ubah folder hasil run_all (debug_output/runs/<waktu>) menjadi satu JSON untuk frontend.

Bentuk keluaran (kunci pendek mengikuti data contoh di frontend/src/data/data.js):
{
  "file", "judul",
  "skor":  {"total", "format", "sistematika"},        # 0-100, None kalau pemeriksaannya tidak berjalan
  "bab":   [{"n": "BAB 1", "s", "sf", "ss",           # skor gabungan, format, sistematika
             "f": [{"jenis", "t", "lokasi", "p", "d", "r", "q"}]}],
  "catatan": [str]
}
Temuan `f`: jenis = "format" | "sistematika"; t = judul; lokasi = teks lokasi; p = nomor halaman (kalau ada);
d = uraian; r = acuan/aturan; q = daftar kutipan kalimat mahasiswa (kalau ada).
"""
import json
import re
from pathlib import Path

MAKS_KUTIPAN = 5

# (pola pada uraian temuan format) -> judul pendek
_JUDUL_FORMAT = [
    (r"^Ukuran font", "Ukuran huruf"),
    (r"menyisakan .* baris kosong", "Jarak persamaan"),
    (r"rata kiri", "Posisi persamaan"),
    (r"[Nn]omor identitas persamaan|tidak memiliki nomor identitas", "Penomoran persamaan"),
    (r"^Abstrak", "Abstrak"),
    (r"^Teks Abstract", "Abstract"),
    (r"berakhir di halaman ganjil", "Halaman kosong akhir bab"),
    (r"^Ukuran kertas", "Ukuran kertas"),
    (r"^Penomoran halaman", "Penomoran halaman"),
    (r"[Mm]irror-margin", "Margin cetak bolak-balik"),
    (r"Daftar Pustaka", "Daftar pustaka"),
    (r"Sumber:", "Keterangan sumber"),
    (r"sampul", "Halaman sampul"),
    (r"[Mm]argin", "Margin"),
    (r"[Ss]pasi", "Spasi baris"),
]
_UNSUR = {"PD": "problem domain", "MASALAH": "masalah", "SOLUSI": "solusi"}


def _baca(run_dir, nama, default=None):
    p = Path(run_dir) / nama
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except ValueError:
        return default


def _kunci_bab(nama):
    if nama == "Pra-Bab":
        return (0, 0, nama)
    m = re.match(r"^BAB\s+(\d+)", nama)
    if m:
        return (1, int(m.group(1)), nama)
    if nama == "DAFTAR PUSTAKA":
        return (2, 0, nama)
    return (3, 0, nama)  # Umum, Lainnya


def _nama_bab(chapter):
    return "Umum" if (chapter or "").startswith("Umum") else (chapter or "Lainnya")


def _judul_format(uraian, kategori):
    for pola, judul in _JUDUL_FORMAT:
        if re.search(pola, uraian):
            return judul
    return kategori or "Format"


def _halaman(lokasi):
    m = re.match(r"^Halaman\s+(\d+)", lokasi or "")
    return int(m.group(1)) if m else None


def _rapikan(judul_subbab):
    """'1.1 LATAR BELAKANG' -> '1.1 Latar Belakang'"""
    num, _, nama = judul_subbab.partition(" ")
    return f"{num} {nama.title()}".strip()


def _skor(checked, failed):
    return round((checked - failed) / checked * 100, 1) if checked else None


def _rata(*nilai):
    v = [x for x in nilai if x is not None]
    return round(sum(v) / len(v), 1) if v else None


# ---------- bahasa awam ----------
# Uraian dari pemeriksa format masih memuat istilah teknis (nama tipe gaya, kata Inggris, istilah Word).
# Semua aturan di bawah mengubah uraian mentah jadi kalimat biasa. Hasilnya tidak lagi cocok dengan pola
# mentahnya, jadi aman dipanggil dua kali (dipakai juga untuk result.json lama di app.py).

_ELEMEN = {
    "chapter_title": "judul bab",
    "subchapter_title": "judul subbab",
    "subsubchapter_title": "judul anak subbab",
    "body": "isi teks",
    "caption": "keterangan gambar/tabel",
    "figure_caption": "keterangan gambar",
    "table_caption": "keterangan tabel",
    "reference": "daftar pustaka",
    "abstract": "abstrak",
}
_PERATAAN = {
    "tidak diset": "belum diatur perataannya",
    "center": "rata tengah",
    "right": "rata kanan",
    "both": "rata kiri-kanan",
    "justify": "rata kiri-kanan",
    "distribute": "rata kiri-kanan",
    "left": "rata kiri",
}
_NOMOR_HALAMAN = {
    "lowerRoman": "angka Romawi kecil (i, ii, iii)",
    "upperRoman": "angka Romawi besar (I, II, III)",
    "decimal": "angka Arab (1, 2, 3)",
    "lowerLetter": "huruf kecil (a, b, c)",
    "upperLetter": "huruf besar (A, B, C)",
}


def _angka(x):
    """'14.0' -> '14', '11.5' -> '11,5' (format Indonesia)."""
    x = float(x)
    return str(int(x)) if x == int(x) else str(x).replace(".", ",")


def _huruf_awal_besar(t):
    return t[:1].upper() + t[1:]


def _ukuran_kertas(m, lokasi):
    saat_ini = re.search(r"\(([\d.]+)\s*cm\s*x\s*([\d.]+)\s*cm\)", lokasi or "")
    awal = f"Ukuran kertas di bagian dokumen ini {_angka(saat_ini.group(1))} x {_angka(saat_ini.group(2))} cm" if saat_ini \
        else "Ukuran kertas di bagian dokumen ini tidak sesuai"
    return f"{awal}, seharusnya A4 (21 x 29,7 cm, tegak maupun mendatar)."


def _nomor_halaman(m, lokasi):
    bab, fmt, mulai = m.group(1), m.group(2), int(m.group(3))
    saat_ini = _NOMOR_HALAMAN.get(fmt, fmt)
    kalimat = f"Nomor halaman mulai {bab} harus memakai angka Arab (1, 2, 3) dan dimulai lagi dari 1. "
    if fmt != "decimal":
        kalimat += f"Saat ini memakai {saat_ini}"
        kalimat += f" dan dimulai dari {mulai}." if mulai != 1 else "."
    else:
        kalimat += f"Saat ini dimulai dari {mulai}."
    return kalimat.strip()


# (pola uraian mentah, fungsi(match, lokasi) -> uraian baru)
_ATURAN_URAIAN = [
    (re.compile(r"^Ukuran font ([\d.]+)\s?pt pada halaman sampul tidak sesuai, seharusnya ([\d.]+)\s?pt\.?"),
     lambda m, l: f"Huruf di halaman sampul berukuran {_angka(m.group(1))} pt, seharusnya {_angka(m.group(2))} pt."),
    (re.compile(r"^Ukuran font ([\d.]+)\s?pt tidak sesuai untuk tipe '(\w+)', seharusnya ([\d.]+)\s?pt\.?"),
     lambda m, l: f"{_huruf_awal_besar(_ELEMEN.get(m.group(2), m.group(2).replace('_', ' ')))} memakai huruf "
                  f"{_angka(m.group(1))} pt, seharusnya {_angka(m.group(3))} pt."),
    (re.compile(r"^Rumus \(persamaan\) seharusnya rata kiri \(mentok kiri\), terdeteksi: ([^.]+?)\..*", re.S),
     lambda m, l: f"Persamaan ini {_PERATAAN.get(m.group(1).strip(), m.group(1).strip())}, seharusnya rata kiri. "
                  "Atur perataan pada kotak persamaannya langsung, bukan hanya perataan paragraf."),
    (re.compile(r"^Persamaan seharusnya menyisakan (\d+) baris kosong di atas dan di bawahnya\.?"),
     lambda m, l: f"Beri {m.group(1)} baris kosong di atas dan di bawah persamaan ini."),
    (re.compile(r"^Nomor identitas persamaan seharusnya dicetak tebal \(bold\)\.?"),
     lambda m, l: "Nomor persamaan seharusnya dicetak tebal."),
    (re.compile(r"^Persamaan tidak memiliki nomor identitas rata kanan, mis\. '\(2\.1\)' atau '\(Persamaan 2\.1\)'\.?"),
     lambda m, l: "Persamaan ini belum diberi nomor di sisi kanan, misalnya (2.1) atau (Persamaan 2.1)."),
    (re.compile(r"^Baris 'Sumber:' seharusnya berukuran ([\d.]+)\s?pt\.?"),
     lambda m, l: f"Keterangan sumber seharusnya berukuran {_angka(m.group(1))} pt."),
    (re.compile(r"^Ukuran kertas di section tsb tidak sesuai A4.*"), _ukuran_kertas),
    (re.compile(r"^Penomoran halaman di section awal (BAB \d+) harus format angka Arab dan dimulai ulang dari 1, "
                r"terdeteksi fmt='(\w+)', start=(\d+)\.?"), _nomor_halaman),
]

_SUFIKS_RENDAH_LAMA = " (Kepastian rendah: font pada hasil render belum terverifikasi.)"
_SUFIKS_RENDAH = " (Perlu dicek sendiri: jenis huruf pada tampilan dokumen belum bisa dipastikan.)"

# sisa istilah teknis yang mungkin muncul di uraian lain
_GANTI_UMUM = [
    (r"\bsection\b", "bagian dokumen"),
    (r"\btsb\b", "tersebut"),
    (r"[Mm]irror-margin", "margin cetak bolak-balik"),
    (r"\bfont\b", "huruf"),
    (r"\bbold\b", "tebal"),
]


def rapikan_uraian(d, lokasi=""):
    """Uraian mentah pemeriksa format -> kalimat biasa. Uraian yang sudah rapi dikembalikan apa adanya."""
    d = d or ""
    rendah = d.endswith(_SUFIKS_RENDAH_LAMA)
    if rendah:
        d = d[: -len(_SUFIKS_RENDAH_LAMA)]
    for pola, fn in _ATURAN_URAIAN:
        m = pola.match(d)
        if m:
            d = fn(m, lokasi) + d[m.end():].strip()
            break
    for pola, ganti in _GANTI_UMUM:
        d = re.sub(pola, ganti, d)
    d = re.sub(r"(\d+)\.0(?=\s?pt\b)", r"\1", d)
    d = re.sub(r"(\d+)\.(\d+)\s?pt\b", r"\1,\2 pt", d)
    d = re.sub(r"(\d+)pt\b", r"\1 pt", d)
    return d.strip() + (_SUFIKS_RENDAH if rendah else "")


def rapikan_judul(t, d_mentah):
    """Judul temuan 'Ukuran huruf' dibuat spesifik: 'Ukuran huruf judul subbab'."""
    m = re.match(r"^Ukuran font [\d.]+\s?pt tidak sesuai untuk tipe '(\w+)'", d_mentah or "")
    if m and t == "Ukuran huruf":
        return f"Ukuran huruf {_ELEMEN.get(m.group(1), m.group(1).replace('_', ' '))}"
    if re.match(r"^Ukuran font [\d.]+\s?pt pada halaman sampul", d_mentah or "") and t == "Ukuran huruf":
        return "Ukuran huruf sampul"
    return t


def rapikan_hasil(hasil):
    """Rapikan result.json yang sudah tersimpan (aman dipanggil berulang)."""
    for b in hasil.get("bab", []):
        for f in b.get("f", []):
            if f.get("jenis") == "format":
                f["t"] = rapikan_judul(f.get("t"), f.get("d"))
                f["d"] = rapikan_uraian(f.get("d"), f.get("lokasi"))
    return hasil



def _temuan_format(run_dir):
    hasil = {}
    for f in _baca(run_dir, "10_visual_findings.json", []):
        d = f.get("description", "")
        lokasi = f.get("location") or ""
        item = {
            "jenis": "format",
            "t": rapikan_judul(_judul_format(d, f.get("category")), d),
            "lokasi": lokasi,  # tetap mentah: dipakai preview.py untuk mencari posisi sorotan
            "p": _halaman(lokasi),
            "d": rapikan_uraian(d, lokasi),
            "r": None,
            "q": [],
        }
        if f.get("confidence") == "rendah":
            item["d"] += _SUFIKS_RENDAH
        hasil.setdefault(_nama_bab(f.get("chapter")), []).append(item)
    return hasil


def _skor_format(run_dir):
    s = _baca(run_dir, "11_visual_score.json")
    if not s:
        return None, {}
    per = {}
    for chapter, c in s.get("per_chapter", {}).items():
        nama = _nama_bab(chapter)
        acc = per.setdefault(nama, {"checked": 0, "failed": 0})
        acc["checked"] += c.get("checked", 0)
        acc["failed"] += c.get("failed", 0)
    return s.get("overall_score"), {n: _skor(c["checked"], c["failed"]) for n, c in per.items()}


def _temuan_sistematika(run_dir):
    """Mengembalikan (temuan per bab, skor per bab, skor keseluruhan, jumlah subbab dinilai)."""
    ringkasan = _baca(run_dir, "summary.json")
    if not ringkasan:
        return {}, {}, None, 0
    rag = {x["subbab"]: x for x in _baca(run_dir, "05_judge_rag.json", [])}
    kal = {x["subbab"]: x for x in _baca(run_dir, "05_judge_kalimat.json", [])}
    temuan, hitung = {}, {}

    for s in ringkasan:
        nomor = str(s.get("bab") or s.get("subbab", "").split(".")[0])
        bab = f"BAB {nomor}"
        status = s.get("akhir")
        if status not in ("sesuai", "tidak_sesuai"):
            continue  # ditunda / error: tidak ikut dinilai
        h = hitung.setdefault(bab, {"ok": 0, "n": 0})
        h["n"] += 1
        if status == "sesuai":
            h["ok"] += 1
            continue

        nama = _rapikan(s["subbab"])
        daftar = temuan.setdefault(bab, [])
        mulai = len(daftar)

        for k in (rag.get(s["subbab"]) or {}).get("kriteria", []):
            if not k.get("terpenuhi"):
                daftar.append({
                    "jenis": "sistematika", "t": f"{nama}: petunjuk panduan belum terpenuhi",
                    "lokasi": f"Subbab {nama}", "p": None,
                    "d": f"Isi subbab {nama} belum memenuhi salah satu petunjuk panduan.",
                    "r": k.get("sumber"), "q": [],
                })
        for a in (kal.get(s["subbab"]) or {}).get("aturan", []):
            if a.get("status") != "SALAH":
                continue
            unsur = _UNSUR.get(a.get("unsur"), (a.get("unsur") or "").lower())
            larangan = a.get("aturan") == "dilarang"
            daftar.append({
                "jenis": "sistematika",
                "t": f"{nama}: {'memuat' if larangan else 'belum memuat'} {unsur}",
                "lokasi": f"Subbab {nama}", "p": None,
                "d": a.get("saran") or a.get("deskripsi") or "",
                "r": a.get("deskripsi"), "q": (a.get("kalimat") or [])[:MAKS_KUTIPAN],
            })
        if len(daftar) == mulai:  # subbab ditandai tidak sesuai, tetapi tidak ada rincian tersimpan
            daftar.append({"jenis": "sistematika", "t": f"{nama}: belum sesuai panduan", "lokasi": f"Subbab {nama}",
                           "p": None, "d": f"Subbab {nama} dinilai belum sesuai panduan.", "r": None, "q": []})

    for r in _baca(run_dir, "06_struktur.json", []):
        if r.get("ditemukan") or r.get("pasti") is False:
            continue
        nama = _rapikan(f"{r['num']} {r['title']}")
        temuan.setdefault(f"BAB {r['bab']}", []).append({
            "jenis": "sistematika", "t": f"Subbab {nama} tidak ditemukan", "lokasi": "Struktur bab", "p": None,
            "d": "Panduan memuat subbab ini, tetapi tidak ada subbab padanan di dokumen.",
            "r": f"Petunjuk panduan subbab {nama}", "q": [],
        })

    skor_bab = {b: _skor(h["n"], h["n"] - h["ok"]) for b, h in hitung.items()}
    n = sum(h["n"] for h in hitung.values())
    ok = sum(h["ok"] for h in hitung.values())
    return temuan, skor_bab, _skor(n, n - ok), n


def build_result(run_dir, nama_file, catatan=()):
    run_dir = Path(run_dir)
    tf = _temuan_format(run_dir)
    skor_f_total, skor_f = _skor_format(run_dir)
    ts, skor_s, skor_s_total, n_sistematika = _temuan_sistematika(run_dir)

    catatan = list(catatan)
    if not (run_dir / "summary.json").exists():
        catatan.append("Pemeriksaan sistematika tidak berjalan, jadi hanya format yang dinilai.")
    if not (run_dir / "11_visual_score.json").exists():
        catatan.append("Pemeriksaan format tidak berjalan, jadi hanya sistematika yang dinilai.")

    nama_bab = set(tf) | set(ts) | set(skor_f) | set(skor_s)
    bab = []
    for n in sorted(nama_bab, key=_kunci_bab):
        bab.append({
            "n": n,
            "s": _rata(skor_f.get(n), skor_s.get(n)),
            "sf": skor_f.get(n),
            "ss": skor_s.get(n),
            "f": tf.get(n, []) + ts.get(n, []),
        })

    judul = (_baca(run_dir, "01_title.json") or {}).get("judul")
    return {
        "file": nama_file,
        "judul": judul,
        "skor": {"total": _rata(skor_f_total, skor_s_total), "format": skor_f_total, "sistematika": skor_s_total},
        "bab": bab,
        "catatan": catatan,
    }
