"""Tahap LLM mode RAG: (1) pilih chunk panduan dari kandidat, (2) turunkan checklist dari
chunk itu dan nilai teks mahasiswa. Semua hasil divalidasi di kode (sumber kriteria dan
kutipan harus benar-benar ada di teks), dan status akhir dihitung di kode, bukan oleh LLM."""
import difflib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "llm"))
from judge import match_key as _norm_sp  # noqa: E402
from llm_client import call_llm, parse_json  # noqa: E402

_norm_title = lambda s: " ".join((s or "").split()).upper()


class LLMParseError(ValueError):
    """Jawaban LLM tidak bisa dibaca sebagai JSON; jawaban mentahnya disimpan di .raw untuk debugging."""

    def __init__(self, msg: str, raw: str = ""):
        super().__init__(msg)
        self.raw = raw


def salvage_kriteria(text: str):
    """JSON rusak/terpotong: ambil objek kriteria yang masih utuh ({"id": ...}). None kalau tidak ada yang selamat."""
    text = re.sub(r"<think>.*?</think>", "", text or "", flags=re.DOTALL)
    i = text.find('"kriteria"')
    if i < 0:
        return None
    dec, objs, pos = json.JSONDecoder(), [], max(text.find("[", i), i) + 1
    while True:
        j = text.find("{", pos)
        if j < 0:
            break
        try:
            o, end = dec.raw_decode(text, j)
        except ValueError:
            pos = j + 1
            continue
        if isinstance(o, dict) and "id" in o:
            objs.append(o)
        pos = end
    return {"kriteria": objs, "_dipulihkan": True} if objs else None


def ask(system: str, user: str):
    """Panggil LLM, parse JSON (ulang sekali kalau rusak). Mengembalikan (hasil, teks_mentah).
    Kalau dua-duanya rusak, coba selamatkan objek kriteria yang utuh; kalau tetap gagal, LLMParseError membawa jawaban mentahnya."""
    last, raw = None, ""
    for attempt in range(2):
        raw = call_llm(system, user, fresh=attempt > 0)  # ulangan melewati cache supaya tidak mengambil jawaban rusak yang sama
        try:
            return parse_json(raw), raw
        except ValueError as e:
            last = e
    rescued = salvage_kriteria(raw)
    if rescued:
        return rescued, raw
    raise LLMParseError(str(last), raw)


# ---------- Langkah A: LLM memilih chunk panduan dari kandidat RAG ----------

MATCH_SYSTEM = """Kamu mencocokkan satu subbab dokumen mahasiswa dengan petunjuk penulisan di panduan.
Pilih SATU petunjuk yang memang mengatur subbab tersebut, berdasarkan judul dan isi subbab.
Jika tidak ada kandidat yang mengatur subbab ini, pilih null; jangan memaksa.
Balas hanya JSON: {"pilihan": "<id kandidat>" atau null, "alasan": "<satu kalimat>"}"""


def rule_pick(title: str, candidates: list[dict], min_ratio: float = 0.92):
    """Pembanding tanpa LLM: kandidat yang judulnya sama dengan judul subbab. Selain persis sama, selisih ejaan
    kecil juga dianggap sama (mis. \"KARAKTERISTIK\" vs \"KARATERISTIK\" di panduan), supaya tidak jatuh ke LLM."""
    t = _norm_title(title)
    best, best_r = None, 0.0
    for c in candidates:
        r = 1.0 if _norm_title(c["subbab_title"]) == t else difflib.SequenceMatcher(None, _norm_title(c["subbab_title"]), t).ratio()
        if r > best_r:
            best, best_r = c["chunk_id"], r
    return best if best_r >= min_ratio else None


def match_guideline(title: str, student_text: str, candidates: list[dict]) -> dict:
    kand = "\n\n".join(
        f"ID: {c['chunk_id']}\nJudul panduan: {c['subbab_title']}\nIsi: {c['text'][:700]}"
        for c in candidates
    )
    user = (
        f"SUBBAB MAHASISWA\nJudul: {title}\nCuplikan isi: {student_text[:500]}\n\n"
        f"KANDIDAT PETUNJUK PANDUAN\n{kand}"
    )
    out, raw = ask(MATCH_SYSTEM, user)
    ids = {c["chunk_id"] for c in candidates}
    pilihan = out.get("pilihan")
    valid = pilihan in ids
    return {
        "pilihan_llm": pilihan if valid else None,
        "alasan": out.get("alasan", ""),
        "pilihan_valid": valid or pilihan is None,
        "pilihan_aturan": rule_pick(title, candidates),
        "trace": {"system": MATCH_SYSTEM, "user": user, "raw": raw},
    }


# ---------- Langkah B: kriteria dari chunk (deterministik) lalu LLM menilai ----------
# Kriteria TIDAK lagi dipecah oleh LLM: tiap kalimat petunjuk di chunk = satu kriteria (dipecah oleh kode, jadi daftar
# K1..Kn sama di setiap run). Kalimat yang berpola anjuran, opsional, contoh, atau fragmen tidak dijadikan kriteria,
# tetapi tetap dicatat beserta alasannya supaya terlihat di laporan. LLM hanya menjawab terpenuhi/tidak per kriteria
# (+ kutipan bukti, untuk validasi di kode). Ubah daftar pola di bawah kalau ada kalimat yang salah golongan.

ABAIKAN = [
    ("anjuran", re.compile(r"\b(?:sebaiknya|akan lebih (?:baik|mudah)|yang baik|jangan lupa|boleh|tidak terlalu)\b", re.I)),
    ("opsional (penulis dapat ...)", re.compile(r"\b(?:penulis|pembaca)\s+dapat\b", re.I)),
    ("gaya penulisan subjektif (singkat, padat, jelas)", re.compile(r"\bsingkat,?\s+padat\b|\bjelas,?\s+singkat\b", re.I)),
    ("klaim orisinalitas tidak dicek", re.compile(r"klaim\s+orisinalitas|\borisinalitas\b", re.I)),
    ("contoh atau template", re.compile(
        r"\(\d+\.\d+\)|seperti\s+(?:ini|kalimat berikut)|\b(?:misalkan|misalnya|sebagai contoh)\b"
        r"|\.{4,}|\(hubungan dengan|^(?:untuk|sedangkan untuk)\s+kutipan", re.I)),
    ("format penulisan (ranah checker visual)", re.compile(
        r"\b(?:tabel|gambar|grafik|persamaan|simbol|notasi|nomer|nomor)\b|sumber referensi|rujukan|daftar pustaka", re.I)),
]
MIN_PANJANG_KRITERIA = 20  # kalimat lebih pendek dari ini dianggap fragmen
MAX_KRITERIA = 12


def kriteria_dari_chunk(chunk_text: str) -> tuple[list[dict], list[dict]]:
    """Pecah teks chunk per kalimat. Mengembalikan (kriteria, diabaikan); diabaikan = [{"kalimat", "alasan"}]."""
    kal = [k.strip() for k in re.split(r"(?<=[.!?])\s+", " ".join((chunk_text or "").split())) if k.strip()]
    kriteria, diabaikan = [], []
    for k in kal:
        alasan = "fragmen" if len(k) < MIN_PANJANG_KRITERIA else next((a for a, p in ABAIKAN if p.search(k)), None)
        if alasan:
            diabaikan.append({"kalimat": k, "alasan": alasan})
        elif len(kriteria) < MAX_KRITERIA:
            kriteria.append({"id": f"K{len(kriteria) + 1}", "sumber": k})
        else:
            diabaikan.append({"kalimat": k, "alasan": f"melebihi batas {MAX_KRITERIA} kriteria"})
    return kriteria, diabaikan


JUDGE_SYSTEM = """Kamu memeriksa kesesuaian ISI satu subbab dokumen proyek akhir terhadap daftar KRITERIA yang sudah
ditetapkan. Tiap kriteria adalah satu kalimat petunjuk panduan. Jangan menambah, menghapus, atau memecah kriteria.
Nilai TEKS MAHASISWA per kriteria. Baca SELURUH teks subbab sekaligus; sebuah kriteria bisa dipenuhi oleh gabungan
beberapa kalimat. "terpenuhi": true hanya jika ada "kutipan": satu sampai tiga kalimat lengkap dari teks mahasiswa,
persis seperti tertulis, dipisah " | ". Jika tidak ada buktinya, "terpenuhi": false dan "kutipan": "". Untuk kriteria
bernada "tidak memuat ...", set "butuh_kutipan": false.
CONTOH di panduan hanya ilustrasi, bukan kriteria. Jangan menilai ejaan atau gaya bahasa.
Judul proyek sudah dipecah menjadi Problem Domain, Problem, dan Uniqueness; gunakan unsur itu untuk
menilai apakah isi subbab membahas hal yang tepat. Jangan menulis alasan atau saran.
Jika ada KONTEKS DARI BAB 1, itu hanya acuan (mis. tujuan atau permasalahan yang ditulis mahasiswa sendiri) untuk
menilai apakah subbab ini selaras dengannya; jangan menilai teks konteks itu dan jangan mengutipnya.
Jika TEKS MAHASISWA hanya sebagian dari subbab, nilai false untuk kriteria yang tidak terbukti di bagian ini
(bagian lain diperiksa terpisah).
Balas hanya satu objek JSON. Gunakan tanda petik ganda hanya untuk struktur JSON; tanda petik di dalam kutipan harus di-escape.
"""

JUDGE_USER = """JUDUL PROYEK AKHIR: {judul}
Problem Domain: {problem_domain} | Problem: {problem} | Uniqueness: {uniqueness}

SUBBAB: {title}
{konteks_blok}
KRITERIA (dari petunjuk panduan):
{kriteria}

CONTOH DI PANDUAN (ilustrasi):
{example}

TEKS MAHASISWA{bagian}:
\"\"\"
{teks}
\"\"\"

Format JSON (satu entri per kriteria, id sama dengan di atas):
{{"kriteria": [{{"id": "K1", "butuh_kutipan": true, "terpenuhi": true, "kutipan": "<kalimat dari teks mahasiswa atau kosong>"}}]}}"""


# Subbab panjang (mis. Hasil Eksperimen puluhan ribu karakter) dibaca per jendela. Jendela 1 menurunkan checklist dan
# menilai; jendela berikutnya HANYA menilai kriteria yang belum terpenuhi dan berhenti begitu semuanya terpenuhi,
# sehingga token tidak terbuang untuk bagian teks yang tidak perlu dibaca.
LANJUT_SYSTEM = """Kamu melanjutkan pemeriksaan isi satu subbab dokumen proyek akhir. Daftar kriteria sudah ditetapkan.
Nilai TEKS MAHASISWA (hanya satu bagian dari subbab) untuk kriteria yang tercantum saja. \"terpenuhi\": true hanya jika
ada \"kutipan\": satu sampai tiga kalimat lengkap dari teks ini, persis seperti tertulis, dipisah \" | \". Jika tidak ada
buktinya di bagian ini, nilai false dengan kutipan kosong. Jangan menulis alasan. Balas hanya satu objek JSON."""

LANJUT_USER = """SUBBAB: {title}

KRITERIA YANG BELUM TERPENUHI:
{daftar}

TEKS MAHASISWA (bagian {k} dari {n}):
\"\"\"
{teks}
\"\"\"

Format JSON: {{"kriteria": [{{"id": "K1", "terpenuhi": true, "kutipan": "<kalimat dari teks ini atau kosong>"}}]}}"""


def pecah_jendela(text: str, max_chars: int) -> list[str]:
    """Bagi teks jadi jendela <= max_chars, dipotong di batas kalimat (kalimat yang lebih panjang dari
    max_chars dipotong paksa)."""
    if len(text) <= max_chars:
        return [text]
    kal = re.split(r"(?<=[.!?])\s+", text)
    out, cur = [], ""
    for k in kal:
        while len(k) > max_chars:
            if cur:
                out.append(cur)
                cur = ""
            out.append(k[:max_chars])
            k = k[max_chars:]
        if cur and len(cur) + 1 + len(k) > max_chars:
            out.append(cur)
            cur = k
        else:
            cur = f"{cur} {k}".strip()
    if cur:
        out.append(cur)
    return out


def compute_status(rows: list[dict]) -> dict:
    """sesuai = semua kriteria wajib terpenuhi; selain itu tidak_sesuai. Tidak ada status 'sebagian'."""
    valid = [r for r in rows if r.get("sumber_valid", True) and r.get("dinilai", True)]  # kriteria yang tidak dijawab LLM tidak dihitung
    if not valid:
        return {"status": "tidak_dinilai", "wajib_terpenuhi": "0/0"}
    ok = sum(bool(r["terpenuhi"]) for r in valid)
    return {"status": "sesuai" if ok == len(valid) else "tidak_sesuai", "wajib_terpenuhi": f"{ok}/{len(valid)}"}


_CITE = re.compile(r"\[\s*\d+(?:\s*[,;\u2013-]\s*\d+)*\s*\]")


def _kata(s: str) -> list[str]:
    return re.findall(r"\w+", _CITE.sub("", s or "").lower())


def _ada_di(frag: str, acuan: str, acuan_n: str, min_cov: float = 0.9) -> bool:
    """Potongan teks dianggap ada di acuan kalau cocok persis (setelah dinormalkan), atau hampir persis:
    >= 90% katanya muncul di acuan. Model sering menyambung/memotong kalimat atau menambah tanda baca;
    kutipan karangan punya banyak kata yang tidak ada di acuan, jadi tetap ditolak."""
    n = _norm_sp(frag.strip().rstrip(".,;:"))
    if n and n in acuan_n:
        return True
    toks = [t for t in _kata(frag) if len(t) > 3]
    if len(toks) < 3:
        return False
    vocab = set(_kata(acuan))
    return sum(t in vocab for t in toks) / len(toks) >= min_cov


def _kutipan_ada(kutipan: str, text: str, text_n: str) -> bool:
    """Kutipan boleh beberapa kalimat (dipisah | atau ...); tiap potongan harus ada di teks mahasiswa."""
    frags = [f for f in re.split(r"\s\|\s|\.\.\.|\u2026", kutipan or "") if len(_norm_sp(f)) >= 12]
    return bool(frags) and all(_ada_di(f, text, text_n) for f in frags)


def _baris_valid(k: dict, i: int, guideline_raw: str, guideline_all: str, student_text: str, text_n: str) -> dict:
    """Satu baris kriteria dari keluaran LLM -> baris tervalidasi (sumber ada di panduan, kutipan ada di teks)."""
    row = {
        "id": k.get("id") or f"K{i}",
        "sumber": k.get("sumber", ""),  # kutipan dari panduan; sekaligus teks kriteria yang ditampilkan
        "butuh_kutipan": bool(k.get("butuh_kutipan", True)),
        "terpenuhi": bool(k.get("terpenuhi", False)),
        "kutipan": k.get("kutipan", "") or "",  # kutipan dari teks mahasiswa; hanya validasi/debug
    }
    # kriteria harus berdasar kalimat yang benar-benar ada di panduan
    # model sering menyingkat kutipan dengan "..."; tiap potongan harus ada di panduan
    frags = [f for f in re.split(r"\.\.\.|\u2026", row["sumber"]) if len(_norm_sp(f)) >= 12]
    row["sumber_valid"] = bool(frags) and all(_ada_di(f, guideline_raw, guideline_all) for f in frags)
    # kriteria terpenuhi harus punya kutipan yang benar-benar ada di teks mahasiswa
    if row["terpenuhi"] and row["butuh_kutipan"] and not _kutipan_ada(row["kutipan"], student_text, text_n):
        row["terpenuhi"] = False
        row["kutipan_tidak_valid"] = True
    return row


def judge_with_guideline(decomp: dict, title: str, student_text: str, chunk: dict, konteks: str = "",
                         max_chars: int = 9000, max_windows: int = 6) -> dict:
    """konteks: teks Bab 1 mahasiswa (mis. Tujuan) sebagai acuan untuk subbab yang menjawabnya (Analisis, Kesimpulan).
    max_chars/max_windows: batas ukuran teks per panggilan LLM dan jumlah jendela yang dibaca."""
    krit, diabaikan = kriteria_dari_chunk(chunk["text"])
    if not krit:
        return {"status": "tidak_dinilai", "wajib_terpenuhi": "0/0", "chunk_id": chunk["chunk_id"], "kriteria": [],
                "diabaikan": diabaikan,
                "catatan": "semua kalimat petunjuk subbab ini berupa anjuran, contoh, atau fragmen, jadi tidak ada kriteria yang dinilai"}

    windows = pecah_jendela(student_text, max_chars)
    n_all = len(windows)
    windows = windows[:max_windows]
    konteks_blok = f"\nKONTEKS DARI BAB 1 (acuan, bukan teks yang dinilai):\n{konteks}\n" if konteks else ""
    user = JUDGE_USER.format(
        judul=decomp.get("judul", ""),
        problem_domain=decomp.get("problem_domain") or "(tidak ada di judul)",
        problem=decomp.get("problem") or "(tidak ada di judul)",
        uniqueness=decomp.get("uniqueness") or "(tidak ada di judul)",
        title=title,
        konteks_blok=konteks_blok,
        kriteria="\n".join(f"{c['id']}: {c['sumber']}" for c in krit),
        example=chunk.get("example") or "(tidak ada)",
        bagian=f" (bagian 1 dari {n_all})" if n_all > 1 else "",
        teks=windows[0],
    )
    out, raw = ask(JUDGE_SYSTEM, user)

    guideline_all = _norm_sp(chunk["text"] + chunk.get("example", ""))
    guideline_raw = chunk["text"] + " " + chunk.get("example", "")
    wn0 = _norm_sp(windows[0])
    by_id = {str(h.get("id")): h for h in (out.get("kriteria") or []) if isinstance(h, dict)}
    rows = []
    for i, c in enumerate(krit, 1):
        h = by_id.get(c["id"])
        if h is None:  # LLM tidak menjawab kriteria ini (jawaban terpotong/rusak): tidak dihitung, bukan dianggap gagal
            rows.append({"id": c["id"], "sumber": c["sumber"], "butuh_kutipan": True, "terpenuhi": False,
                         "kutipan": "", "sumber_valid": True, "dinilai": False})
            continue
        rows.append(_baris_valid({**h, "sumber": c["sumber"]}, i, guideline_raw, guideline_all, windows[0], wn0))

    # jendela lanjutan: hanya kriteria "butuh kutipan" yang belum terpenuhi
    dibaca, traces = 1, []
    for w_i in range(1, len(windows)):
        belum = [r for r in rows if r["sumber_valid"] and r["butuh_kutipan"] and not r["terpenuhi"]]
        if not belum:
            break
        daftar = "\n".join(f"- {r['id']}: {r['sumber']}" for r in belum)
        u2 = LANJUT_USER.format(title=title, daftar=daftar, k=w_i + 1, n=n_all, teks=windows[w_i])
        o2, raw2 = ask(LANJUT_SYSTEM, u2)
        dibaca += 1
        traces.append({"user": u2, "raw": raw2})
        by2 = {str(h.get("id")): h for h in (o2.get("kriteria") or []) if isinstance(h, dict)}
        wn = _norm_sp(windows[w_i])
        for r in belum:
            h = by2.get(str(r["id"]))
            if h and bool(h.get("terpenuhi")) and _kutipan_ada(h.get("kutipan", "") or "", windows[w_i], wn):
                r["terpenuhi"], r["kutipan"], r["dinilai"] = True, h.get("kutipan", ""), True
                r.pop("kutipan_tidak_valid", None)

    res = compute_status(rows)
    res.update({"chunk_id": chunk["chunk_id"], "kriteria": rows, "diabaikan": diabaikan,
                "jendela": {"total": n_all, "dibaca": dibaca},
                "trace": {"system": JUDGE_SYSTEM, "user": user, "raw": raw}})
    tak = sum(1 for r in rows if not r.get("dinilai", True))
    if out.get("_dipulihkan") or tak:
        res["catatan"] = f"jawaban LLM rusak atau terpotong: {len(rows) - tak} dari {len(rows)} kriteria terbaca"
        if tak and res["status"] == "sesuai":  # sebagian kriteria belum terjawab, jadi belum boleh dinyatakan sesuai
            res["status"] = "tidak_dinilai"
            res["catatan"] += "; status sesuai tidak bisa dipastikan, jalankan ulang subbab ini"
    if traces:
        res["trace_lanjutan"] = traces
    return res
