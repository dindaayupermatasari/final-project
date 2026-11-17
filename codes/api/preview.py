"""Pratinjau PDF untuk halaman Hasil: konversi DOCX -> PDF dan penghitungan posisi sorotan temuan.

Sorotan dicari di PDF yang sama dengan yang ditampilkan ke pengguna (PDF asli, atau hasil konversi LibreOffice
untuk DOCX), jadi nomor halaman dan koordinatnya selalu cocok dengan pratinjau.

Format keluaran `hitung_sorotan`:
{
  "versi": 2,
  "halaman": 35,
  "ukuran": [[595.3, 841.9], ...],             # lebar, tinggi tiap halaman (poin PDF), untuk rasio halaman
  "temuan": {"<indeks bab>-<indeks temuan>": {  # sama dengan urutan result["bab"][i]["f"][j]
      "jenis": "teks" | "judul" | "bab" | "halaman",
      "hal": 3,                                 # halaman pertama (1-based), None kalau tidak ketemu
      "rects": [{"p": 3, "r": [x0, y0, x1, y1]}]  # pecahan 0-1 dari lebar/tinggi halaman
  }}
}
"""
import os
import re
import shutil
import subprocess
import tempfile
from collections import defaultdict
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:  # PyMuPDF versi lama
    import fitz  # type: ignore


# ---------- DOCX -> PDF ----------

def _soffice():
    for nama in ("soffice", "libreoffice"):
        p = shutil.which(nama)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def docx_ke_pdf(sumber, tujuan, batas_detik=180):
    """Ubah .docx ke PDF dengan LibreOffice headless. Profil sementara dipakai supaya tidak bentrok
    dengan LibreOffice yang sedang terbuka atau konversi lain."""
    exe = _soffice()
    if not exe:
        raise RuntimeError("LibreOffice (soffice) tidak ditemukan, pratinjau DOCX tidak bisa dibuat.")
    with tempfile.TemporaryDirectory() as tmp:
        keluar = Path(tmp) / "out"
        profil = (Path(tmp) / "profil").as_uri()
        cmd = [exe, f"-env:UserInstallation={profil}", "--headless", "--convert-to", "pdf",
               "--outdir", str(keluar), str(sumber)]
        try:
            subprocess.run(cmd, check=True, timeout=batas_detik, capture_output=True)
        except (subprocess.SubprocessError, OSError) as e:
            raise RuntimeError(f"Konversi DOCX ke PDF gagal: {e}") from e
        hasil = keluar / (Path(sumber).stem + ".pdf")
        if not hasil.exists():
            raise RuntimeError("Konversi DOCX ke PDF tidak menghasilkan berkas.")
        Path(tujuan).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(hasil), str(tujuan))
    return Path(tujuan)


# ---------- sorotan ----------

# Naikkan angka ini kalau cara menghitung sorotan berubah, supaya cache sorotan.json dan beranotasi.* yang lama dibuang.
VERSI = 4

_POLA_NOMOR_PERSAMAAN = re.compile(r"^\(\s*(?:Persamaan\s*)?\d+(?:\.\d+)+\s*\)$")
_AWAL_NOMOR = re.compile(r"^\s*\d+(?:\.\d+)*\s*")  # '2.3' pada '2.3PENELITIAN TERKAIT' (penomoran otomatis Word)
_POLA_TABEL_PERSAMAAN = re.compile(r"^Persamaan dalam tabel \(tabel ke-(\d+)")


def _norm(s):
    return re.sub(r"[^0-9a-zA-Z\u00C0-\u024F]+", "", s).lower()


class _Dokumen:
    """Aliran kata seluruh PDF (urutan baca) dengan posisinya, supaya kalimat yang terpotong baris
    atau pindah halaman tetap bisa ditemukan."""

    def __init__(self, path):
        self.doc = fitz.open(path)
        self.ukuran = [[p.rect.width, p.rect.height] for p in self.doc]
        self.tok, self.pos = [], []            # token ternormalisasi, (halaman, bbox, blok, baris)
        self.baris = []                        # (halaman, teks baris, bbox) untuk pencarian judul
        self.bbox_baris = {}                   # (halaman, blok, baris) -> bbox seluruh baris
        for n, page in enumerate(self.doc):
            per_baris = defaultdict(list)
            for x0, y0, x1, y1, w, blk, ln, _ in page.get_text("words", sort=True):
                t = _norm(w)
                if t:
                    self.tok.append(t)
                    self.pos.append((n, (x0, y0, x1, y1), blk, ln))
                per_baris[(blk, ln)].append((x0, y0, x1, y1, w))
            for (blk, ln), kata in per_baris.items():
                bb = (min(k[0] for k in kata), min(k[1] for k in kata), max(k[2] for k in kata), max(k[3] for k in kata))
                self.baris.append((n, " ".join(k[4] for k in kata), bb))
                self.bbox_baris[(n, blk, ln)] = bb
        self.indeks = defaultdict(list)
        for i, t in enumerate(self.tok):
            self.indeks[t].append(i)

    def tutup(self):
        self.doc.close()

    def _pecahan(self, n, bbox, pad=0.0):
        w, h = self.ukuran[n]
        x0, y0, x1, y1 = bbox
        x0, y0, x1, y1 = max(0, x0 - pad), max(0, y0 - pad), min(w, x1 + pad), min(h, y1 + pad)
        return {"p": n + 1, "r": [round(x0 / w, 5), round(y0 / h, 5), round(x1 / w, 5), round(y1 / h, 5)]}

    def token_dari(self, n, y=0.0):
        """Indeks token pertama pada halaman n (0-based) di bawah koordinat y; tempat mulai pencarian berikutnya."""
        for i, (hal, bb, _blk, _ln) in enumerate(self.pos):
            if hal > n or (hal == n and bb[1] >= y - 2):
                return i
        return len(self.tok)

    def cari_kalimat(self, kalimat, dari=0, awalan=False):
        """Rentang token (awal, akhir) kalimat di dokumen, atau None.
        awalan=True: kata terakhir boleh terpotong (kutipan yang dipendekkan dengan '...')."""
        kal = [t for t in (_norm(x) for x in kalimat.split()) if t]
        n = len(kal)
        if n == 0:
            return None
        awal = [i for i in self.indeks.get(kal[0], []) if i >= dari]
        for i in awal:  # cocok persis
            if self.tok[i:i + n] == kal:
                return i, i + n
        if awalan:  # kata terakhir hanya awalan dari kata di dokumen
            for i in awal:
                if self.tok[i:i + n - 1] == kal[:-1] and i + n <= len(self.tok) and self.tok[i + n - 1].startswith(kal[-1]):
                    return i, i + n
        k = min(4, n)  # nomor halaman/header yang menyela, atau tanda baca aneh: cocokkan awal dan akhirnya
        kepala, ekor = kal[:k], kal[-k:]
        for i in awal:
            if self.tok[i:i + k] != kepala:
                continue
            for j in range(i + int(n * 0.6), min(i + int(n * 1.6) + 10, len(self.tok) - k) + 1):
                if self.tok[j:j + k] == ekor:
                    return i, j + k
        return None

    def kotak_rentang(self, a, b, penuh=False):
        """Satu kotak per baris teks pada rentang token [a, b). penuh=True: kotak seluruh baris, bukan hanya kata-katanya."""
        per = {}
        for i in range(a, b):
            n, bb, blk, ln = self.pos[i]
            kunci = (n, blk, ln)
            if penuh:
                per[kunci] = self.bbox_baris[kunci]
            elif kunci in per:
                o = per[kunci]
                per[kunci] = (min(o[0], bb[0]), min(o[1], bb[1]), max(o[2], bb[2]), max(o[3], bb[3]))
            else:
                per[kunci] = bb
        return [self._pecahan(k[0], bb) for k, bb in sorted(per.items())]

    def _kandidat(self, sama_dengan, kapital=True, nomor=False):
        """Baris judul (tanpa titik pengisi daftar isi) yang isinya persis `sama_dengan`.
        kapital=True: hanya yang huruf kapital semua (judul bab/subbab biasa).
        nomor=True: nomor subbab di awal baris diabaikan. Hasil konversi DOCX mencetak penomoran otomatis Word
        menempel di judul ('2.3PENELITIAN TERKAIT'), padahal pemeriksa hanya menyebut 'PENELITIAN TERKAIT'."""
        target = _norm(sama_dengan)
        hasil = []
        for n, teks, bb in self.baris:
            if ".." in teks or "\u2026" in teks or (kapital and teks != teks.upper()):
                continue
            if _norm(teks) == target or (nomor and _norm(_AWAL_NOMOR.sub("", teks)) == target):
                hasil.append((n, bb))
        return hasil

    def nomor_halaman_cetak(self):
        """{nomor yang tercetak di margin halaman: halaman PDF 1-based}. Nomor cetak dimulai ulang setelah halaman
        romawi, jadi 'halaman 18' dari pemeriksa DOCX bukan urutan halaman di PDF."""
        peta = {}
        for n, page in enumerate(self.doc):
            h = page.rect.height
            angka = [w[4] for w in page.get_text("words") if (w[3] < 70 or w[1] > h - 80) and w[4].isdigit()]
            if len(angka) == 1:
                peta.setdefault(int(angka[0]), n + 1)
        return peta

    def cari_baris(self, sama_dengan, terakhir=False):
        ketemu = self._kandidat(sama_dengan)
        if not ketemu:
            return None
        n, bb = ketemu[-1] if terakhir else ketemu[0]
        return self._pecahan(n, bb)

    def cari_judul_subbab(self, nomor, judul):
        """Judul subbab seperti '1.3 TUJUAN'. Nomor kadang berada di baris terpisah dari judulnya (penomoran
        otomatis Word), jadi dicari juga baris nomor yang sejajar dengan baris judul. Daftar isi dilewati
        karena yang dipakai kemunculan terakhir."""
        for kapital in (True, False):
            for n, bb in reversed(self._kandidat(f"{nomor} {judul}", kapital)):
                return self._pecahan(n, bb)
            hanya_judul = self._kandidat(judul, kapital)
            for n, bb in reversed(hanya_judul):
                for m, teks, bn in self.baris:
                    if m == n and _norm(teks) == _norm(nomor) and abs(bn[1] - bb[1]) < 6:
                        return self._pecahan(n, (min(bb[0], bn[0]), min(bb[1], bn[1]), max(bb[2], bn[2]), max(bb[3], bn[3])))
            if len(hanya_judul) == 1:  # tidak ada baris nomor, tetapi judulnya tidak ambigu
                return self._pecahan(*hanya_judul[0])
        return None

    def rentang_bab(self, nama_bab):
        """(halaman awal, halaman akhir) 0-based sebuah bab. Pra-Bab = sebelum BAB 1; selain itu seluruh dokumen."""
        akhir = len(self.ukuran) - 1
        m = re.match(r"BAB\s+(\d+)", nama_bab or "")

        def hal(n):
            k = self.cari_baris(f"BAB {n}")
            return k["p"] - 1 if k else None

        if nama_bab == "Pra-Bab":
            b1 = hal(1)
            return (0, max(0, b1 - 1)) if b1 is not None else (0, akhir)
        if not m:
            return 0, akhir
        n = int(m.group(1))
        awal = hal(n)
        if awal is None:
            return 0, akhir
        berikut = hal(n + 1)
        return awal, (max(awal, berikut - 1) if berikut is not None else akhir)

    def _font_pengganti_matematika(self):
        """True kalau font berawalan 'Liberation' hanya muncul untuk persamaan. Tanpa Cambria Math, LibreOffice mencetak
        persamaan Word dengan Liberation Serif/OpenSymbol. Kalau seluruh dokumen ber-font Liberation (tidak ada
        Times New Roman di mesin konversi), font itu tidak bisa dipakai sebagai petunjuk."""
        if not hasattr(self, "_liberation_khusus"):
            hitung = defaultdict(int)
            for page in list(self.doc)[:60]:
                for b in page.get_text("dict")["blocks"]:
                    for ln in b.get("lines", []):
                        for sp in ln["spans"]:
                            hitung[sp["font"].lower()] += len(sp["text"])
            utama = max(hitung, key=hitung.get) if hitung else ""
            self._liberation_khusus = not utama.startswith("liberation")
        return self._liberation_khusus

    def klaster_persamaan(self, hal_a, hal_b, longgar=False):
        """Persamaan mandiri (bukan di dalam kalimat) pada halaman hal_a..hal_b, urut baca: [(halaman 0-based, bbox, punya_nomor)].
        Persamaan Word dicetak dengan font Cambria Math; baris-baris yang berdekatan (pembilang, penyebut, nomor
        persamaan di kanan) digabung jadi satu persamaan, lalu yang tak memuat '=' atau nomor dibuang (mis. satuan di tabel)."""
        liberation = longgar and self._font_pengganti_matematika()

        def adalah_mat(font):
            f = font.lower()
            if "math" in f:
                return True
            return longgar and ("opensymbol" in f or (f.startswith("liberation") and (liberation or "italic" in f)))

        hasil = []
        for n in range(hal_a, hal_b + 1):
            garis = []
            for b in self.doc[n].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    teks = "".join(s["text"] for s in ln["spans"]).strip()
                    if not teks:
                        continue
                    tot = sum(len(s["text"]) for s in ln["spans"]) or 1
                    mat = sum(len(s["text"]) for s in ln["spans"] if adalah_mat(s["font"]))
                    nomor = bool(_POLA_NOMOR_PERSAMAAN.match(teks))
                    if mat / tot > 0.5 or nomor:
                        garis.append((list(ln["bbox"]), teks, nomor))
            garis.sort(key=lambda g: (g[0][1], g[0][0]))
            grup = []
            for bb, teks, nomor in garis:
                if grup and bb[1] - grup[-1]["bb"][3] < 4:  # bertumpuk/menempel = satu persamaan; ada celah = persamaan lain
                    g = grup[-1]
                    g["bb"] = [min(g["bb"][0], bb[0]), min(g["bb"][1], bb[1]), max(g["bb"][2], bb[2]), max(g["bb"][3], bb[3])]
                    g["ada"] = g["ada"] or ("=" in teks)
                    g["nomor"] = g["nomor"] or nomor
                else:
                    grup.append({"bb": bb, "ada": ("=" in teks), "nomor": nomor})
            hasil += [(n, tuple(g["bb"]), g["nomor"]) for g in grup if g["ada"] or g["nomor"]]
        return hasil


_POLA_PARAGRAF = re.compile(r"^Paragraf\s+#(\d+)\s*\((.*)\)\s*$")


def _teks_potongan(s):
    """'Sumber : https://www.dat...' -> ('Sumber : https://www.dat', True)"""
    s = s.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "'\"":
        s = s[1:-1]
    terpotong = s.endswith(("...", "\u2026"))
    return s.rstrip(".\u2026 ").strip(), terpotong


def _sorotan_temuan(dok, bab_nama, f, ctx):
    """-> data sorotan, atau None kalau posisinya tidak bisa dipastikan.
    Sorotan selalu menunjuk ke kalimat/baris/persamaan yang bermasalah; judul bab hanya dipakai untuk
    temuan tentang struktur bab. ctx: keadaan pencarian satu bab (posisi terakhir yang sudah ketemu)."""
    # 1. kalimat mahasiswa yang dikutip pemeriksa sistematika
    # Tiap kalimat dicari dari awal bab, tidak maju terus antar temuan: urutan kutipan di tiap temuan tidak
    # selalu sama dengan urutan di dokumen, dan kalimat temuan berikutnya bisa terletak sebelum temuan sebelumnya.
    rects = []
    for q in f.get("q") or []:
        r = dok.cari_kalimat(q, ctx["q"])
        if r:
            rects += dok.kotak_rentang(*r)
    if rects:
        return {"jenis": "teks", "hal": rects[0]["p"], "rects": rects}

    lokasi = f.get("lokasi") or ""

    # 2. 'Paragraf #163 ('1.1 LATAR BELAKANG')' atau 'Paragraf #206 (persamaan)'
    m = _POLA_PARAGRAF.match(lokasi)
    if m:
        no_par, isi = int(m.group(1)), m.group(2).strip()
        if isi == "persamaan":
            k = ctx["persamaan"].get(no_par)
            if not k:
                return None
            kotak = dok._pecahan(k[0], k[1], pad=3)
            return {"jenis": "teks", "hal": kotak["p"], "rects": [kotak]}
        teks, terpotong = _teks_potongan(isi)
        mj = re.match(r"^(\d+(?:\.\d+)+)\s+(.+)$", teks)
        if mj:  # judul subbab bernomor
            kotak = dok.cari_judul_subbab(mj.group(1), mj.group(2))
            if kotak:
                ctx["t"] = dok.token_dari(kotak["p"] - 1, kotak["r"][1] * dok.ukuran[kotak["p"] - 1][1])
                return {"jenis": "judul", "hal": kotak["p"], "rects": [kotak]}
        if teks and teks == teks.upper() and not terpotong:  # judul kapital tanpa nomor, mis. 'PROYEK AKHIR' atau 'TUJUAN'
            ketemu = [(n, bb) for n, bb in dok._kandidat(teks, nomor=True) if ctx["hal"] <= n <= ctx["hal_akhir"]]
            if ketemu:
                # judul-judul subbab berurutan: pilih kemunculan setelah judul sebelumnya, bukan yang pertama di bab
                setelah = [k for k in ketemu if (k[0], k[1][1]) > ctx.get("judul_akhir", (-1, -1))]
                n, bb = (setelah or ketemu)[0]
                ctx["judul_akhir"] = (n, bb[1])
                kotak = dok._pecahan(n, bb)
                ctx["t"] = dok.token_dari(n, bb[1])
                return {"jenis": "judul", "hal": kotak["p"], "rects": [kotak]}
            if len(teks.split()) <= 8:
                # Judul tidak ketemu sebagai baris sendiri. Mencari kata-katanya di seluruh teks malah menyorot
                # kalimat atau caption yang kebetulan memuat kata yang sama (mis. 'Penelitian Terkait'), jadi tanpa sorotan.
                return None
        r = dok.cari_kalimat(teks, ctx["t"], awalan=True) if teks else None
        if r and dok.pos[r[0]][0] <= ctx["hal_akhir"]:
            rects = dok.kotak_rentang(*r, penuh=True)
            ctx["t"] = r[0]
            return {"jenis": "teks", "hal": rects[0]["p"], "rects": rects}
        return None

    # 2b. 'Persamaan dalam tabel (tabel ke-3 dalam dokumen)': hanya ada untuk DOCX (persamaan Word di dalam tabel dua kolom)
    m = _POLA_TABEL_PERSAMAAN.match(lokasi)
    if m:
        return _sorotan_tabel_persamaan(dok, int(m.group(1)), ctx)

    # 3. temuan sistematika: judul subbab
    m = re.match(r"Subbab\s+(\d+(?:\.\d+)*)\s+(.+)$", lokasi)
    if m:
        kotak = dok.cari_judul_subbab(m.group(1), m.group(2))
        return {"jenis": "judul", "hal": kotak["p"], "rects": [kotak]} if kotak else None

    # 4. 'Abstrak (Paragraf #37-#40)': judul bagian
    if "(" in lokasi and not lokasi.startswith(("Halaman", "Section")):
        kotak = dok.cari_baris(lokasi.split("(")[0].strip())
        return {"jenis": "judul", "hal": kotak["p"], "rects": [kotak]} if kotak else None

    # 5. subbab yang seharusnya ada tetapi tidak ada: tidak ada teks untuk disorot, tunjuk judul babnya
    if lokasi == "Struktur bab":
        mb = re.match(r"BAB\s+(\d+)", bab_nama)
        kotak = dok.cari_baris(f"BAB {mb.group(1)}") if mb else None
        return {"jenis": "bab", "hal": kotak["p"], "rects": [kotak]} if kotak else None

    # 6. hanya tahu halamannya. Pemeriksa DOCX menyebut nomor halaman cetak, bukan urutan halaman di PDF.
    if f.get("p"):
        hal = ctx.get("peta_cetak", {}).get(f["p"], f["p"])
        return {"jenis": "halaman", "hal": hal, "rects": []}
    return None


def _sorotan_tabel_persamaan(dok, no, ctx):
    """Persamaan di dalam tabel: tabel tidak punya teks yang bisa dicari, jadi dicari paragraf di depannya
    (dari DOCX), lalu persamaan pertama sesudah paragraf itu di PDF."""
    sebelum = ctx.get("tabel_sebelum", {}).get(no)
    if not sebelum:
        return None
    r = dok.cari_kalimat(sebelum, ctx["mulai"])
    if r:
        akhir = r[1] - 1
    else:
        # Akhir kalimat tak cocok (persamaan sebaris, nomor rujukan, dst): cukup awalnya, akhir paragraf = akhir blok teksnya
        kal = [t for t in (_norm(x) for x in sebelum.split()) if t][:5]
        awal = next((i for i in dok.indeks.get(kal[0], []) if i >= ctx["mulai"] and dok.tok[i:i + len(kal)] == kal), None) if kal else None
        if awal is None:
            return None
        akhir = awal
        while akhir + 1 < len(dok.pos) and dok.pos[akhir + 1][0] == dok.pos[awal][0] and dok.pos[akhir + 1][2] == dok.pos[awal][2]:
            akhir += 1
    n_a, bb_a, _blk, _ln = dok.pos[akhir]
    for n, bb, _nomor in dok.klaster_persamaan(n_a, min(n_a + 1, len(dok.ukuran) - 1), longgar=True):
        if n > n_a or bb[1] >= bb_a[3] - 2:
            kotak = dok._pecahan(n, bb, pad=3)
            return {"jenis": "teks", "hal": kotak["p"], "rects": [kotak]}
    return None


def _paragraf_sebelum_tabel(docx_path):
    """{nomor tabel (1-based, urutan di dokumen): teks paragraf tak kosong terakhir sebelum tabel itu}.
    Hanya tabel yang memuat persamaan Word."""
    from docx import Document

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    hasil = {}
    for k, t in enumerate(Document(str(docx_path)).tables, 1):
        if t._tbl.find(f".//{m}oMath") is None:
            continue
        el = t._tbl.getprevious()
        while el is not None:
            if el.tag == f"{w}p":
                teks = " ".join("".join(x.text or "" for x in el.iter(f"{w}t")).split())
                if len(teks.split()) >= 4:
                    hasil[k] = teks
                    break
            el = el.getprevious()
    return hasil


def _perlu_nomor(temuan):
    """{nomor paragraf persamaan yang bermasalah: True (harus punya nomor) / False (harus tanpa nomor) / None}"""
    perlu = {}
    for f in temuan:
        m = re.match(r"Paragraf\s+#(\d+)\s*\(persamaan\)", f.get("lokasi") or "")
        if not m:
            continue
        no, d = int(m.group(1)), f.get("d") or ""
        perlu.setdefault(no, None)
        if re.search(r"belum diberi nomor|tidak memiliki nomor", d):
            perlu[no] = False
        elif "dicetak tebal" in d:
            perlu[no] = True
    return perlu


def _satuan(paragraf, gabung):
    """Paragraf persamaan -> satuan persamaan. gabung=True: nomor paragraf yang berurutan (#791, #792) dianggap satu
    persamaan yang terpecah oleh ekstraksi PDF (pembilang, penyebut, dan nomornya sering jadi paragraf sendiri)."""
    hasil = []
    for no in paragraf:
        if gabung and hasil and no == hasil[-1][-1] + 1:
            hasil[-1].append(no)
        else:
            hasil.append([no])
    return hasil


def _cocok_nomor(satuan, klaster, perlu):
    return all(perlu[no] is None or perlu[no] == klaster[2] for no in satuan)


def _halaman_perkiraan(no, jangkar):
    """Perkiraan posisi (halaman 0-based + pecahan) paragraf nomor `no`, diinterpolasi dari paragraf lain yang posisinya
    sudah pasti (judul subbab, kalimat kutipan). None kalau tidak ada jangkar."""
    bawah = [j for j in jangkar if j[0] <= no]
    atas = [j for j in jangkar if j[0] >= no]
    if bawah and atas:
        (n1, h1), (n2, h2) = bawah[-1], atas[0]
        return float(h1) if n2 == n1 else h1 + (h2 - h1) * (no - n1) / (n2 - n1)
    if bawah:
        return float(bawah[-1][1])
    if atas:
        return float(atas[0][1])
    return None


def _petakan_interpolasi(dok, satuan, klaster, perlu, jangkar):
    """Jumlah persamaan di halaman lebih banyak daripada yang bermasalah (sebagian persamaan sudah benar). Pilih
    persamaan yang paling dekat dengan perkiraan posisi tiap paragraf (tetap urut baca), dan hanya dipakai kalau
    pilihan terbaik jelas lebih baik daripada pilihan terbaik berikutnya."""
    from itertools import combinations
    from math import comb

    m, n = len(satuan), len(klaster)
    if m == 0 or m > n or comb(n, m) > 5000:
        return {}
    harap = []
    for u in satuan:
        h = _halaman_perkiraan(u[0], jangkar)
        if h is None:
            return {}
        harap.append(h)
    posisi = [k[0] + k[1][1] / dok.ukuran[k[0]][1] for k in klaster]
    skor = []
    for pilih in combinations(range(n), m):
        biaya = 0.0
        for u, h, c in zip(satuan, harap, pilih):
            biaya += abs(posisi[c] - h)
            if not _cocok_nomor(u, klaster[c], perlu):
                biaya += 3  # petunjuk nomor tidak cocok: bukan tolak mutlak karena pemeriksa dan PDF bisa beda membaca nomor
        skor.append((biaya, pilih))
    skor.sort()
    if len(skor) > 1 and skor[1][0] - skor[0][0] < 1.0:
        return {}  # dua pilihan hampir sama bagusnya: lebih baik tanpa sorotan daripada menyorot persamaan lain
    return {no: klaster[c] for u, c in zip(satuan, skor[0][1]) for no in u}


def _petakan_persamaan(dok, temuan, hal_awal, hal_akhir, jangkar=(), longgar=False):
    """{nomor paragraf: (halaman, bbox, punya_nomor)} untuk persamaan yang bermasalah di satu bab, atau {} kalau tidak yakin.
    Urutan percobaan: (1) jumlah persamaan di halaman bab sama dengan jumlah paragraf bermasalah; (2) sama setelah
    paragraf bernomor berurutan digabung jadi satu persamaan; (3) persamaan lebih banyak daripada yang bermasalah:
    dicocokkan dengan perkiraan posisi dari jangkar `jangkar` = [(nomor paragraf, halaman 0-based)] yang sudah pasti."""
    perlu = _perlu_nomor(temuan)
    if not perlu:
        return {}
    paragraf = sorted(perlu)
    klaster = dok.klaster_persamaan(hal_awal, hal_akhir, longgar=longgar)
    if not klaster:
        return {}
    pilihan = [_satuan(paragraf, False)]
    gabung = _satuan(paragraf, True)
    if len(gabung) != len(pilihan[0]):
        pilihan.append(gabung)
    for satuan in pilihan:  # (1) dan (2)
        if len(satuan) == len(klaster) and all(_cocok_nomor(u, k, perlu) for u, k in zip(satuan, klaster)):
            return {no: k for u, k in zip(satuan, klaster) for no in u}
    for satuan in reversed(pilihan):  # (3), gabungan dulu karena lebih sedikit tebakan
        hasil = _petakan_interpolasi(dok, satuan, klaster, perlu, sorted(jangkar))
        if hasil:
            return hasil
    return {}


def _persamaan(f):
    return bool(re.match(r"Paragraf\s+#\d+\s*\(persamaan\)", f.get("lokasi") or ""))


def hitung_sorotan(pdf_path, hasil, docx_path=None):
    """docx_path: DOCX asli kalau PDF-nya hasil konversi. Dipakai untuk temuan yang acuannya urutan tabel di DOCX
    dan untuk menerjemahkan nomor halaman cetak ke halaman PDF."""
    dok = _Dokumen(pdf_path)
    try:
        temuan, konteks, jangkar = {}, {}, []
        ekstra = {}
        if docx_path:
            ekstra = {"tabel_sebelum": _paragraf_sebelum_tabel(docx_path), "peta_cetak": dok.nomor_halaman_cetak()}

        # Tahap 1: semua temuan selain persamaan. Judul subbab dan kalimat yang paragrafnya bernomor menjadi jangkar
        # (nomor paragraf -> halaman) untuk memperkirakan letak persamaan pada tahap 2.
        for i, b in enumerate(hasil.get("bab", [])):
            nama = b.get("n", "")
            hal_awal, hal_akhir = dok.rentang_bab(nama)
            mulai = dok.token_dari(hal_awal)
            ctx = {"q": mulai, "t": mulai, "mulai": mulai, "hal": hal_awal, "hal_akhir": hal_akhir, "persamaan": {}, **ekstra}
            konteks[i] = (hal_awal, hal_akhir, ctx)
            for j, f in enumerate(b.get("f", [])):
                if _persamaan(f):
                    continue
                data = _sorotan_temuan(dok, nama, f, ctx)
                if data:
                    temuan[f"{i}-{j}"] = data
                    m = _POLA_PARAGRAF.match(f.get("lokasi") or "")
                    if m and data.get("hal") and data["jenis"] in ("judul", "teks") and data["rects"]:
                        r = data["rects"][0]
                        jangkar.append((int(m.group(1)), r["p"] - 1 + r["r"][1]))

        # Jangkar harus naik bersama nomor paragraf; yang melanggar (salah cocok) dibuang.
        jangkar.sort()
        bersih = []
        for no, h in jangkar:
            if not bersih or h >= bersih[-1][1] - 1e-6:
                bersih.append((no, h))
        jangkar = bersih

        # Tahap 2: persamaan. Tidak ada teks untuk dicari, jadi dipetakan ke persamaan di halaman bab itu
        # (lihat _petakan_persamaan); kalau tidak yakin, tanpa sorotan daripada menyorot persamaan lain.
        for i, b in enumerate(hasil.get("bab", [])):
            hal_awal, hal_akhir, ctx = konteks[i]
            ctx["persamaan"] = _petakan_persamaan(dok, b.get("f", []), hal_awal, hal_akhir, jangkar, longgar=bool(docx_path))
            for j, f in enumerate(b.get("f", [])):
                if not _persamaan(f):
                    continue
                data = _sorotan_temuan(dok, b.get("n", ""), f, ctx)
                if data:
                    temuan[f"{i}-{j}"] = data
        temuan = {k: temuan[k] for k in sorted(temuan, key=lambda x: tuple(int(v) for v in x.split("-")))}
        return {"versi": VERSI, "halaman": len(dok.ukuran), "ukuran": dok.ukuran, "temuan": temuan}
    finally:
        dok.tutup()
