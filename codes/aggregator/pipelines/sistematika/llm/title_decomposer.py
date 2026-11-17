import json
import re

from llm_client import call_llm_json

SYSTEM = """Kamu membantu memeriksa judul proyek akhir. Pecah judul menjadi tiga unsur:
- problem_domain: bidang/objek yang menjadi latar masalah (mis. "penyakit kanker", "proyek akhir").
- problem: masalah yang ingin diselesaikan, DITULIS DENGAN POLA "sulitnya <kata kerja> <objek>"
  (mis. "sulitnya mendeteksi penyakit kanker"). Kata kerja boleh diturunkan dari kata benda di judul
  (Pengecekan -> mengecek, Deteksi -> mendeteksi, Prediksi -> memprediksi). Objeknya harus kata-kata
  dari judul. Jangan menambah konteks, alasan, atau fakta yang tidak ada di judul.
- uniqueness: solusi, metode, atau pendekatan yang ditawarkan (mis. "sistem pakar berbasis fuzzy"),
  biasanya muncul setelah kata seperti "dengan", "menggunakan", "berbasis", "using", atau pada pola
  "Sistem/Aplikasi/Model X untuk Y" (X adalah uniqueness, Y bagian dari problem domain/problem).
Aturan: kata-kata problem_domain dan uniqueness harus berasal dari judul; jangan menerjemahkan atau
menambah kata. Kalau judul tidak memuat salah satu unsur, isi dengan string kosong dan jelaskan di
"catatan" (jangan menebak). Balas hanya dengan satu objek JSON."""

# (judul, problem_domain, problem, uniqueness, catatan)
CONTOH = [
    ("Sistem Pengecekan Kebenaran Buku Proyek Akhir Menggunakan LLM dan RAG",
     "proyek akhir", "sulitnya mengecek kebenaran buku proyek akhir",
     "sistem pengecekan kebenaran menggunakan LLM dan RAG", ""),
    ("Deteksi Penyakit Kanker Dengan Sistem Pakar Berbasis Fuzzy",
     "penyakit kanker", "sulitnya mendeteksi penyakit kanker", "sistem pakar berbasis fuzzy", ""),
    ("Cancer Detection using Neural Networks",
     "cancer", "sulitnya mendeteksi cancer", "neural networks", ""),
    ("Sistem Peringkas Otomatis untuk Berita Daring",
     "berita daring", "sulitnya meringkas berita daring", "sistem peringkas otomatis", ""),
    ("Prediksi Harga Beras Menggunakan LSTM",
     "harga beras", "sulitnya memprediksi harga beras", "LSTM", ""),
    ("Klasifikasi Sentimen Ulasan Aplikasi Mobile Berbasis IndoBERT",
     "ulasan aplikasi mobile", "sulitnya mengklasifikasikan sentimen ulasan aplikasi mobile", "IndoBERT", ""),
    ("Optimasi Rute Pengiriman Barang Dengan Algoritma Genetika",
     "pengiriman barang", "sulitnya mengoptimalkan rute pengiriman barang", "algoritma genetika", ""),
    ("Aplikasi Pemantauan Kualitas Udara Berbasis IoT",
     "kualitas udara", "sulitnya memantau kualitas udara", "aplikasi berbasis IoT", ""),
    ("Chatbot Layanan Akademik Menggunakan RAG",
     "layanan akademik", "sulitnya mendapatkan layanan akademik", "chatbot menggunakan RAG", ""),
    ("Analisis Sentimen Ulasan Produk",
     "ulasan produk", "sulitnya menganalisis sentimen ulasan produk", "",
     "Judul tidak memuat solusi/metode, uniqueness perlu diisi user."),
    ("Sistem Informasi Akademik",
     "akademik", "", "sistem informasi",
     "Judul tidak memuat tindakan/masalah, problem perlu diisi user."),
]


def _fewshot() -> str:
    parts = []
    for i, (j, pd, p, u, c) in enumerate(CONTOH, 1):
        obj = {"problem_domain": pd, "problem": p, "uniqueness": u, "catatan": c}
        parts.append(f"Contoh {i}\nJudul: {j}\n{json.dumps(obj, ensure_ascii=False)}")
    return "\n\n".join(parts) + "\n"


FEWSHOT = _fewshot()

_STOP = {"yang", "dan", "untuk", "dengan", "di", "pada", "dari", "ke", "dalam", "atau", "sulitnya", "sulit",
         "kesulitan", "the", "of", "in", "a", "using"}


def _kata(s: str) -> list[str]:
    return [w for w in re.findall(r"\w+", (s or "").lower()) if w not in _STOP]


def _overlap(unsur: str, judul: str) -> float:
    """Porsi kata bermakna di unsur yang juga ada di judul."""
    kt = _kata(unsur)
    jt = set(_kata(judul))
    return sum(w in jt for w in kt) / len(kt) if kt else 0.0


def decompose_title(judul: str) -> dict:
    user = f"{FEWSHOT}\nSekarang kerjakan.\nJudul: {judul}\nBalas dengan JSON berkunci problem_domain, problem, uniqueness, catatan."
    out = call_llm_json(SYSTEM, user)

    masalah = []
    # problem domain & uniqueness: kata-katanya harus dari judul; problem boleh memuat kata kerja turunan
    batas = {"problem_domain": 0.8, "problem": 0.5, "uniqueness": 0.8}
    for k, b in batas.items():
        v = (out.get(k) or "").strip()
        out[k] = v
        if not v:
            masalah.append(f"{k} kosong")
        elif _overlap(v, judul) < b:
            masalah.append(f"{k} banyak memuat kata di luar judul")
    out["perlu_dicek"] = masalah  # untuk ditampilkan di gerbang konfirmasi
    out["judul"] = judul
    return out


if __name__ == "__main__":
    import sys

    j = sys.argv[1] if len(sys.argv) > 1 else "Deteksi Penyakit Kanker Dengan Sistem Pakar Berbasis Fuzzy"
    print(json.dumps(decompose_title(j), ensure_ascii=False, indent=2))
