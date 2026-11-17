"""Aturan penulisan Bab 1 per kalimat: LLM melabeli SETIAP kalimat (PD / MASALAH / SOLUSI / LAIN),
verdict dihitung di kode.

Modul ini hanya memegang aturan "dilarang" (unsur yang tidak boleh muncul di suatu subbab).
Kriteria isi (apa yang HARUS dibahas) tidak ditulis di sini: itu diturunkan LLM dari chunk panduan
hasil RAG (lihat rag_judge.py). Hasil keduanya digabung oleh gabung_status().

Label:
  PD      membahas problem domain (bidang/objek, kondisi, tren, urgensi bidangnya)
  MASALAH membahas problem: kesulitan, kendala, kekurangan yang ingin diselesaikan
  SOLUSI  membahas solusi/metode/sistem yang ditawarkan atau dibangun (uniqueness)
  LAIN    di luar ketiganya (definisi umum, transisi, dsb)
"""
import re

from llm_client import call_llm_json

LABELS = ("PD", "MASALAH", "SOLUSI", "LAIN")

# Aturan "dilarang": unsur yang tidak boleh muncul di subbab (0 kalimat). Kriteria isi yang
# HARUS ada tidak ditulis di sini; itu datang dari chunk panduan (RAG).
ATURAN = {
    "LATAR BELAKANG": {"MASALAH": "dilarang", "SOLUSI": "dilarang"},
    "PERMASALAHAN": {"SOLUSI": "dilarang"},
    "TUJUAN": {"MASALAH": "dilarang"},
}
_NAMA_SUBBAB = {"LATAR BELAKANG": "latar belakang", "PERMASALAHAN": "permasalahan", "TUJUAN": "tujuan"}
_URAIAN = {
    "MASALAH": "uraian permasalahan (problem) yang ingin diselesaikan",
    "SOLUSI": "uraian solusi/metode yang ditawarkan",
}


def deskripsi_aturan(canon: str, lab: str) -> str:
    """Teks aturan yang tampil di laporan, mis. 'Pada permasalahan tidak boleh memuat uraian solusi/metode yang ditawarkan.'"""
    return f"Pada {_NAMA_SUBBAB.get(canon, canon.lower())} tidak boleh memuat {_URAIAN[lab]}."

ALIAS = {
    "LATAR BELAKANG MASALAH": "LATAR BELAKANG",
    "RUMUSAN MASALAH": "PERMASALAHAN",
    "TUJUAN PENELITIAN": "TUJUAN",
    "TUJUAN PROYEK AKHIR": "TUJUAN",
}
PINDAH = {  # saran templat untuk unsur yang salah tempat
    ("LATAR BELAKANG", "MASALAH"): "Latar belakang tidak boleh membahas problem; pindahkan ke Permasalahan.",
    ("LATAR BELAKANG", "SOLUSI"): "Latar belakang tidak boleh membahas solusi; pindahkan ke Tujuan.",
    ("PERMASALAHAN", "SOLUSI"): "Permasalahan tidak boleh memuat solusi; pindahkan ke Tujuan.",
    ("TUJUAN", "MASALAH"): "Tujuan tidak boleh mengulang permasalahan; pindahkan ke Permasalahan.",
}

SYSTEM = """Kamu membaca satu subbab Bab 1 proyek akhir KALIMAT PER KALIMAT dan memberi satu label
untuk setiap kalimat. Label:
- PD: membahas problem domain yang tertulis di bawah SECARA LANGSUNG (bidang/objek itu sendiri, kondisinya,
  tren, dampak bidangnya, urgensinya). Topik yang berkaitan tetapi bukan problem domain itu (mis. konteks umum,
  dokumen asal, institusi) diberi LAIN.
- MASALAH: membahas problem yang ingin diselesaikan penelitian: kesulitan, kendala, atau kekurangan pada
  problem domain atau cara yang ada saat ini (kata penanda umum: namun, tetapi, sulit, belum, tantangan).
  Kalimat yang menyebut AKIBAT atau konsekuensi dari kesulitan/kendala/kekurangan itu (mis. "Akibatnya,
  hasil penelitian hanya tersimpan sebagai laporan ...") juga MASALAH, bukan PD, karena melanjutkan uraian
  masalah. PD hanya untuk dampak/kondisi bidangnya sendiri yang tidak bergantung pada kesulitan tadi.
- SOLUSI: membahas solusi/metode/sistem yang ditawarkan atau dibangun penulis (uniqueness), termasuk
  fungsi, komponen, dan alasan adanya komponen itu. Keterbatasan atau kebutuhan dari solusi yang sedang
  dibangun sendiri (mis. "sistem sulit diakses tanpa antarmuka", lalu kalimat berikutnya menyebut
  antarmuka) adalah SOLUSI, bukan MASALAH.
  Kalimat yang menyatakan perlunya solusi, sistem, atau pendekatan (mis. "oleh karena itu diperlukan ...",
  "kondisi ini menunjukkan perlunya ...") juga SOLUSI, walaupun solusinya belum diuraikan.
- LAIN: di luar ketiganya.
Satu kalimat satu label; pilih yang paling dominan di kalimat itu. Baca kalimat sebelum dan sesudahnya
sebelum memutuskan. Nilai berdasarkan isi kalimat, bukan letaknya. Jangan menilai ejaan atau gaya bahasa. Balas hanya satu objek JSON."""

USER = """JUDUL: {judul}
Problem Domain: {problem_domain}
Problem: {problem}
Uniqueness/Solusi: {uniqueness}

SUBBAB: {subbab}

KALIMAT:
{kalimat}

Format JSON: {{"label": [{{"n": 1, "label": "PD|MASALAH|SOLUSI|LAIN", "alasan": "<satu frasa>"}}]}}
Harus ada tepat satu entri untuk setiap nomor kalimat."""

# Singkatan yang titiknya BUKAN akhir kalimat. "dll."/"dst."/"dsb." sengaja tidak dimasukkan
# karena sering menutup kalimat.
_ABBR = r"(?:Dr|Prof|Ir|Gambar|Tabel|Bab|No|Hal|Vol|al|Rp|mis)"
# Penanda kalimat "perlunya solusi": dipakai hanya untuk memilih kandidat yang diperiksa ulang oleh LLM.
NEED = re.compile(r"\b(diperlukan(nya)?|dibutuhkan(nya)?|perlunya|membutuhkan|memerlukan|perlu adanya)\b", re.I)
# Lebih ketat dari NEED: kalimat yang jelas menyatakan perlunya solusi/sistem. Dipakai sebagai pagar
# supaya pemeriksa ulang tidak memindahkan kalimat semacam ini ke MASALAH.
NEED_SOLUSI = re.compile(
    r"\b(diperlukan(nya)?|dibutuhkan(nya)?|perlunya|perlu adanya|membutuhkan|memerlukan)\b[^.]{0,40}\b(solusi|sistem|pendekatan|metode|aplikasi)\b",
    re.I,
)
_SPLIT = re.compile(r"(?<=[.!?])\s+(?=[A-Z\"\u201c(\[])")


def split_sentences(text: str) -> list[str]:
    """Pecah teks jadi kalimat; singkatan umum (dll., dst., et al., Gambar 1.) dan angka desimal tidak memecah."""
    t = " ".join((text or "").split())
    t = re.sub(r"\b(%s)\." % _ABBR, lambda m: m.group(1) + "\u2024", t, flags=re.I)  # titik singkatan -> penanda sementara
    t = re.sub(r"(\d)\.(\d)", "\\1\u2024\\2", t)
    parts = _SPLIT.split(t)
    return [p.replace("\u2024", ".").strip() for p in parts if len(p.strip()) > 3]


# Kalimat pembuka "akibat": kalau kalimat sebelumnya MASALAH, kalimat ini melanjutkan uraian masalah.
AKIBAT = re.compile(r"^(?:akibatnya|akibat (?:hal|kondisi|ini|tersebut)|sebagai akibat(?:nya)?|dampaknya|imbasnya|"
                    r"hal (?:ini|tersebut) (?:menyebabkan|berdampak|mengakibatkan)|kondisi (?:ini|tersebut) (?:menyebabkan|mengakibatkan))\b", re.I)


def lanjutkan_akibat(rows: list[dict]) -> None:
    """Pagar setelah pelabelan: kalimat berpenanda akibat tepat setelah kalimat MASALAH dan berlabel PD/LAIN
    dipindah ke MASALAH (label semula disimpan di label_awal)."""
    for i in range(1, len(rows)):
        r = rows[i]
        if r["label"] in ("PD", "LAIN") and rows[i - 1]["label"] == "MASALAH" and AKIBAT.search(r["kalimat"]):
            r["label_awal"] = r["label"]
            r["label"] = "MASALAH"
            r["alasan"] = (r.get("alasan") or "") + " (akibat dari masalah sebelumnya)"


def label_sentences(decomp: dict, subbab: str, sentences: list[str]) -> list[dict]:
    kal = "\n".join(f"{i}. {s}" for i, s in enumerate(sentences, 1))
    ctx = {k: (decomp.get(k) or "(tidak ada)") for k in ("problem_domain", "problem", "uniqueness")}
    out = call_llm_json(SYSTEM, USER.format(judul=decomp.get("judul", ""), subbab=subbab, kalimat=kal, **ctx), retries=2)
    by_n = {}
    for h in out.get("label", []):
        try:
            by_n[int(h.get("n"))] = h
        except (TypeError, ValueError):
            continue
    rows = []
    for i, s in enumerate(sentences, 1):
        h = by_n.get(i, {})
        lab = str(h.get("label", "")).upper()
        rows.append({"n": i, "kalimat": s, "label": lab if lab in LABELS else "TIDAK_DINILAI",
                     "alasan": h.get("alasan", "")})
    return rows


VERIF_SYSTEM = """Kamu memeriksa ulang kalimat-kalimat yang ditandai melanggar aturan penulisan Bab 1.
Untuk setiap kalimat bertanda, jawab "melanggar": true HANYA jika kalimat itu sendiri menguraikan atau
mengulang unsur terlarang yang disebut. Jawab false jika kalimat itu hanya alasan atau pengantar menuju
bagian solusi/fitur berikutnya, hanya menyebut hal lain, atau tidak terkait dengan unsur terlarang itu.
Isi juga "label_benar": label yang paling tepat untuk kalimat itu (PD, MASALAH, SOLUSI, atau LAIN); jika
kalimat itu sebenarnya berlabel lain, tulis label yang lain itu. Gunakan nomor kalimat PERSIS seperti pada daftar.
Jangan menjawab false hanya karena nama metode atau istilah tertentu tidak disebut persis di kalimat itu;
yang dinilai adalah apakah isinya membahas unsur terlarang tersebut.
Kalimat yang menyatakan perlunya solusi, sistem, atau pendekatan (mis. "oleh karena itu, diperlukan suatu
solusi yang ...") berlabel SOLUSI, bukan MASALAH, walaupun alasannya menyebut kesulitan.
Gunakan kalimat sebelum dan sesudahnya sebagai konteks. Balas hanya satu objek JSON."""

VERIF_USER = """SUBBAB: {subbab}

KALIMAT BERTANDA:
{items}

Format JSON: {{"hasil": [{{"n": <nomor kalimat bertanda>, "label_benar": "PD|MASALAH|SOLUSI|LAIN", "melanggar": true, "alasan": "<satu frasa>"}}]}}
Harus ada tepat satu entri untuk setiap nomor."""


def verify_flagged(decomp: dict, subbab: str, rows: list[dict]) -> None:
    """Lapis kedua: periksa ulang kalimat yang kena aturan 'dilarang'. Kalimat yang tidak benar-benar
    menguraikan unsur terlarang (mis. alasan menuju fitur solusi) diberi dibebaskan=True.
    Kalau pemeriksaan gagal atau tidak menjawab suatu kalimat, kalimat itu tetap dianggap melanggar."""
    key = subbab.strip().upper()
    spec = ATURAN.get(ALIAS.get(key, key)) or {}
    defn = {
        "MASALAH": ('masalah yang ingin diselesaikan penelitian: '
                    f'"{decomp.get("problem") or "(tidak ada)"}". Termasuk kesulitan atau kekurangan cara/metode yang ada '
                    'saat ini dalam menyelesaikan masalah itu. TIDAK termasuk keterbatasan atau kebutuhan dari '
                    'sistem/solusi yang sedang dibangun penelitian sendiri (mis. sistem sulit diakses tanpa antarmuka, '
                    'lalu kalimat berikutnya menyebut antarmuka)'),
        "SOLUSI": ('solusi yang ditawarkan penelitian, yaitu sistem/metode yang dibangun, dengan keunikan: '
                   f'"{decomp.get("uniqueness") or "(tidak ada)"}". Kalimat membahas solusi jika menyebut sistem, '
                   'metode, atau pendekatan (termasuk simulasi atau optimasi) yang akan dipakai atau dibangun '
                   'untuk menyelesaikan masalah, walaupun nama metode keunikannya tidak disebut persis. Kalimat yang hanya '
                   'menyatakan perlunya solusi/sistem/pendekatan (mis. "oleh karena itu diperlukan ...") juga dianggap '
                   'membahas solusi'),
        "PD": f'problem domain: "{decomp.get("problem_domain") or "(tidak ada)"}"',
    }
    targets = [r for lab, mode in spec.items() if mode == "dilarang" for r in rows if r["label"] == lab]
    kand = []
    if spec.get("SOLUSI") == "dilarang":  # kalimat "perlunya solusi" yang tidak dilabeli SOLUSI ikut diperiksa
        kand = [r for r in rows if r["label"] != "SOLUSI" and r not in targets and NEED.search(r["kalimat"])]
    kand_n = {r["n"] for r in kand}
    targets = targets + kand
    if not targets:
        return
    items_all = []
    for r in targets:
        i = r["n"]
        prev = rows[i - 2]["kalimat"] if i >= 2 else "(awal subbab)"
        nxt = rows[i]["kalimat"] if i < len(rows) else "(akhir subbab)"
        unsur = defn["SOLUSI" if i in kand_n else r["label"]]
        items_all.append(f"{i}. Unsur terlarang: {unsur}\n   sebelum: {prev}\n   KALIMAT: {r['kalimat']}\n   sesudah: {nxt}")
    belum = {r["n"]: r for r in targets}
    for _ in range(2):  # percobaan kedua hanya untuk kalimat yang belum dijawab
        if not belum:
            break
        items = [it for it in items_all if int(it.split(".", 1)[0]) in belum]
        try:
            out = call_llm_json(VERIF_SYSTEM, VERIF_USER.format(subbab=subbab, items="\n".join(items)), retries=2)
        except Exception as e:  # noqa: BLE001
            print(f"    !! verifikasi {subbab} gagal: {str(e)[:120]}")
            continue
        hs = _items(out)
        cocok = [h for h in hs if _num(h) in belum]
        if not cocok and hs and len(hs) == len(belum):  # model menomori ulang dari 1: cocokkan menurut urutan
            cocok = [dict(h, n=n) for h, n in zip(hs, sorted(belum))]
        if not cocok:
            print(f"    !! verifikasi {subbab}: balasan tidak cocok dengan nomor kalimat: {str(out)[:160]}")
        for h in cocok:
            r = belum.pop(_num(h))
            baru = str(h.get("label_benar", "")).strip().upper()
            if r["label"] == "SOLUSI" and baru == "MASALAH" and NEED_SOLUSI.search(r["kalimat"]):
                baru = "SOLUSI"  # "diperlukan solusi ..." tetap SOLUSI; pemeriksa tidak boleh memindahkannya ke MASALAH
            if r["n"] in kand_n and baru != "SOLUSI":
                continue  # kandidat hanya dipindah kalau memang SOLUSI; selain itu label semula dipertahankan
            r["alasan_verifikasi"] = h.get("alasan", "")  # disimpan untuk semua kalimat bertanda
            if baru in LABELS and baru != r["label"]:
                r["label_awal"] = r["label"]  # pelabel keliru: pindahkan ke label yang benar
                r["label"] = baru
            elif _bool(h.get("melanggar")) is False:
                r["dibebaskan"] = True
    belum = {n: r for n, r in belum.items() if n not in kand_n}  # kandidat yang tak terjawab dibiarkan apa adanya
    for r in belum.values():  # tetap dianggap melanggar, tapi ditandai belum terverifikasi
        r["verifikasi_gagal"] = True
    if belum:
        print(f"    !! {subbab}: {len(belum)} kalimat bertanda belum terverifikasi (kalimat {sorted(belum)})")


def _num(h):
    try:
        return int(h.get("n"))
    except (TypeError, ValueError, AttributeError):
        return None


def _items(out) -> list[dict]:
    """Ambil daftar jawaban dari balasan LLM: {"hasil": [...]}, daftar lain, atau satu objek {"n": ...}."""
    if isinstance(out, list):
        return [x for x in out if isinstance(x, dict)]
    if isinstance(out, dict):
        for v in out.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
        if "n" in out:
            return [out]
    return []


def _bool(v):
    if isinstance(v, str):
        return v.strip().lower() in ("true", "ya", "1", "melanggar")
    return None if v is None else bool(v)


def _canon(subbab: str) -> str:
    key = subbab.strip().upper()
    return ALIAS.get(key, key)


def has_aturan(subbab: str) -> bool:
    """True kalau subbab ini punya aturan 'dilarang' (Latar Belakang, Permasalahan, Tujuan)."""
    return _canon(subbab) in ATURAN


def verdict(subbab: str, rows: list[dict]) -> dict:
    """Terapkan aturan 'dilarang' pada label per kalimat (kalimat yang dibebaskan verifikasi tidak dihitung)."""
    canon = _canon(subbab)
    spec = ATURAN.get(canon)
    hitung = {l: len([r for r in rows if r["label"] == l]) for l in LABELS}
    if spec is None:
        return {"status": "tidak_ada_aturan", "aturan": [], "tidak_dinilai": [], "hitung": hitung}
    hasil = []
    for lab, mode in spec.items():
        catatan = ""
        kal = [r["kalimat"] for r in rows if r["label"] == lab and not r.get("dibebaskan")]
        bebas = [r["kalimat"] for r in rows if r["label"] == lab and r.get("dibebaskan")]
        st = "SALAH" if kal else "BENAR"
        if bebas:
            catatan = f"{len(bebas)} kalimat dibebaskan setelah verifikasi"
        gagal = [r["n"] for r in rows if r["label"] == lab and r.get("verifikasi_gagal")]
        if gagal:
            catatan = (catatan + "; " if catatan else "") + f"kalimat {gagal} belum terverifikasi"
        hasil.append({"unsur": lab, "aturan": mode, "deskripsi": deskripsi_aturan(canon, lab), "status": st, "kalimat": kal,
                      "dibebaskan": bebas, "saran": PINDAH.get((canon, lab), "") if st == "SALAH" else "",
                      "catatan": catatan})
    tidak_dinilai = [r["n"] for r in rows if r["label"] == "TIDAK_DINILAI"]
    status = "SALAH" if any(h["status"] == "SALAH" for h in hasil) else "BENAR"
    return {"status": status, "aturan": hasil, "tidak_dinilai": tidak_dinilai, "hitung": hitung}


def judge_subbab_kalimat(decomp: dict, subbab: str, text: str) -> dict:
    sents = split_sentences(text)
    if not sents:
        return {"status": "kosong", "kalimat": [], "aturan": []}
    rows = label_sentences(decomp, subbab, sents)
    lanjutkan_akibat(rows)
    verify_flagged(decomp, subbab, rows)
    return {**verdict(subbab, rows), "kalimat": rows}


def aturan_sebagai_kriteria(kal: dict | None) -> list[dict]:
    """Ubah hasil aturan 'dilarang' ke baris kriteria (ID, kriteria, terpenuhi) supaya tampil satu tabel
    dengan kriteria panduan. Kalimat pelanggar tetap ada di kal['kalimat'] untuk debugging, tidak di tabel."""
    return [{"id": f"A{i}", "sumber": h["deskripsi"], "terpenuhi": h["status"] != "SALAH"}
            for i, h in enumerate((kal or {}).get("aturan", []), 1)]


# Kriteria panduan yang sudah tercakup pemecahan judul (problem domain / problem / solusi) tidak dinilai
# ulang oleh LLM RAG. Hasilnya diambil dari label kalimat supaya dua jalur tidak saling bertentangan.
# Kriteria bernada kualitas (urgensi) tetap dinilai LLM RAG.
SAMA_DENGAN_JUDUL = {
    "LATAR BELAKANG": (re.compile(r"problem domain", re.I), "PD"),
    "PERMASALAHAN": (re.compile(r"\bproblem\b(?!\s+domain)", re.I), "MASALAH"),
    "TUJUAN": (re.compile(r"\bsolusi\b", re.I), "SOLUSI"),
}
_KUALITAS = re.compile(r"urgensi|meyakinkan|\bsingkat\b|\bjelas\b", re.I)


def selaraskan_dengan_label(subbab: str, rag: dict | None, kal: dict | None) -> None:
    """Timpa kriteria RAG yang membahas unsur judul dengan hasil label kalimat:
    terpenuhi kalau ada minimal satu kalimat berlabel unsur itu (yang tidak dibebaskan verifikasi)."""
    spec = SAMA_DENGAN_JUDUL.get(_canon(subbab))
    if not spec or not rag or not rag.get("kriteria") or not kal or not kal.get("kalimat"):
        return
    pola, label = spec
    ada = any(r["label"] == label and not r.get("dibebaskan") for r in kal["kalimat"])
    for k in rag["kriteria"]:
        teks = k.get("sumber", "")
        if pola.search(teks) and not _KUALITAS.search(teks):
            k["terpenuhi"] = ada
            k["dinilai_oleh"] = "label_kalimat"
    from rag_judge import compute_status  # noqa: E402  (hindari impor melingkar saat modul dimuat)
    rag.update(compute_status(rag["kriteria"]))


_NILAI = {"sesuai": 0, "sebagian": 1, "tidak_sesuai": 2}  # "sebagian" hanya dipakai mode manual


def gabung_status(rag: dict | None, kal: dict | None) -> dict:
    """Status akhir subbab = yang terburuk antara penilaian panduan (RAG) dan aturan 'dilarang'.
    Aturan SALAH dihitung tidak_sesuai. Lapisan yang tidak bisa dinilai (tidak ada padanan panduan,
    teks kosong, error) tidak ikut dihitung, tetapi dicatat; kalau hanya aturan 'dilarang' yang lolos
    sementara panduan tidak bisa dinilai, hasilnya tidak_dinilai (bukan sesuai). Kalau ada lapisan error
    dan hasil lainnya sesuai, hasilnya error (jalankan ulang), bukan sesuai."""
    lap = {}
    if rag and rag.get("status") in _NILAI:
        lap["panduan"] = rag["status"]
    if kal and kal.get("status") in ("BENAR", "SALAH"):
        lap["aturan"] = "sesuai" if kal["status"] == "BENAR" else "tidak_sesuai"
    gagal = [n for n, r in (("panduan", rag), ("aturan", kal)) if r and r.get("status") == "error"]
    catatan = ""
    if lap and not ("panduan" in lap) and rag is not None and lap.get("aturan") == "sesuai":
        status = "tidak_dinilai"
        catatan = f"panduan tidak dinilai ({(rag or {}).get('status', '-')}); aturan penulisan lolos"
    elif lap:
        status = max(lap.values(), key=_NILAI.get)
    elif gagal:
        status = "error"
    else:
        status = (rag or {}).get("status") or (kal or {}).get("status") or "tidak_dinilai"
    if gagal:
        catatan = (catatan + "; " if catatan else "") + "lapisan error: " + ", ".join(gagal)
        if status == "sesuai":  # satu lapisan tidak terperiksa, jadi belum boleh dinyatakan sesuai
            status = "error"
    return {"status": status, "lapisan": lap, "gagal": gagal, "catatan": catatan}
