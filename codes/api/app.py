"""API PENScek: unggah dokumen -> jalankan run_all.run() di latar belakang -> hasil JSON untuk frontend.

Riwayat, hasil, dan dokumen disimpan di Supabase (lihat api/db.py dan supabase/schema.sql). Folder api/data
hanya tempat kerja sementara dan dibersihkan otomatis.

Jalankan dari folder proyek (yang berisi run_all.py dan aggregator/):
    pip install fastapi uvicorn python-multipart supabase python-dotenv
    uvicorn api.app:app --reload --port 8000

Semua endpoint (kecuali /api/health) butuh header  Authorization: Bearer <token login Supabase>.
Setiap orang hanya bisa melihat pemeriksaan yang ia unggah sendiri.

Endpoint:
    GET    /api/me                  profil yang sedang masuk: id, email, role (mhs|dsn), nama
    POST   /api/checks              form-data: file (.pdf/.docx), judul (opsional), mhs_nama (wajib untuk dosen)
                                    -> {"id": ...}
    GET    /api/checks/{id}         status: queued | running | menunggu | done | error
                                    menunggu = judul sudah dipecah, tunggu konfirmasi user (ada "decomp");
                                    done = ada "result"
    POST   /api/checks/{id}/konfirmasi   json: judul, problem_domain, problem, uniqueness -> mulai pemeriksaan penuh
    GET    /api/checks              daftar pemeriksaan milik pengguna (untuk Riwayat / Dashboard), terbaru dulu;
                                    "dokumen_tersedia" false = dokumen sudah dihapus karena melewati masa simpan
    DELETE /api/checks/{id}         hapus pemeriksaan beserta dokumennya
    GET    /api/checks/{id}/pdf       PDF untuk pratinjau (PDF asli; untuk .docx hasil konversi LibreOffice)
    GET    /api/checks/{id}/sorotan   posisi sorotan temuan di PDF itu (lihat api/preview.py)
    GET    /api/checks/{id}/unduh/pdf|docx   dokumen asli + sorotan + komentar temuan (lihat api/ekspor.py);
                                    docx hanya untuk dokumen yang diunggah sebagai .docx.
                                    pdf, sorotan, dan unduh butuh dokumennya; setelah masa simpan habis jawabannya 410.

Pengaturan lewat environment variable (boleh di file .env di folder proyek):
    SUPABASE_URL, SUPABASE_SERVICE_KEY   wajib (service key hanya untuk backend, jangan masuk frontend)
    PENSCEK_RETRIEVAL     live (default, butuh Ollama + ChromaDB) atau cached
    PENSCEK_DATA          folder kerja sementara (default api/data)
    PENSCEK_RETENSI_HARI  lama dokumen disimpan di Supabase (default 30; hasil pemeriksaan disimpan selamanya)
    PENSCEK_CACHE_HARI    lama salinan kerja di folder lokal sebelum dibersihkan (default 2)
    PENSCEK_HAPUS_RUN     1 (default) = folder run pipeline dihapus setelah hasilnya tersimpan; 0 = disimpan untuk debug
    PENSCEK_ORIGINS       alamat frontend tambahan untuk CORS, pisahkan dengan koma (untuk deploy)
"""
import contextlib
import copy
import io
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.background import BackgroundTask

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:  # .env di folder proyek (opsional)
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
except ImportError:
    pass

from api import db  # noqa: E402
from api.adapter import build_result, rapikan_hasil  # noqa: E402
from api.ekspor import docx_beranotasi, pdf_beranotasi  # noqa: E402
from api.preview import VERSI as VERSI_SOROTAN, docx_ke_pdf, hitung_sorotan  # noqa: E402

DATA = Path(os.getenv("PENSCEK_DATA") or ROOT / "api" / "data")
RETRIEVAL = os.getenv("PENSCEK_RETRIEVAL", "live")
CACHE_HARI = float(os.getenv("PENSCEK_CACHE_HARI", "2"))
HAPUS_RUN = os.getenv("PENSCEK_HAPUS_RUN", "1") != "0"
ALLOWED = {".pdf", ".docx"}
MAX_MB = 50
PENANDA = ".cache"  # folder kerja bikinan kode ini; hanya folder bertanda ini yang dibersihkan otomatis

JOBS: dict[str, dict] = {}  # hanya pemeriksaan yang sedang berjalan atau menunggu; yang selesai dibaca dari Supabase
# Satu dokumen sekali jalan: LibreOffice headless dan kuota token LLM tidak cocok dengan pemeriksaan paralel,
# dan run_all.run() membuat folder hasil bernama detik (dua pemeriksaan di detik yang sama akan bentrok).
_pool = ThreadPoolExecutor(max_workers=1)


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _folder_kerja(job_id: str) -> Path:
    d = DATA / job_id
    d.mkdir(parents=True, exist_ok=True)
    (d / PENANDA).touch()
    return d


def _bersihkan_lokal():
    """Buang folder kerja lama. Aman: dokumen dan hasilnya ada di Supabase, dan akan diunduh lagi kalau dibutuhkan."""
    if not DATA.exists():
        return
    batas = time.time() - CACHE_HARI * 86400
    for d in DATA.iterdir():
        if d.is_dir() and (d / PENANDA).exists() and d.name not in JOBS and d.stat().st_mtime < batas:
            shutil.rmtree(d, ignore_errors=True)


def _putaran_retensi(henti: threading.Event):
    while not henti.is_set():
        try:
            n = db.hapus_kadaluarsa()
            if n:
                print(f"Retensi: {n} dokumen melewati masa simpan dan dihapus.")
            _bersihkan_lokal()
        except Exception as e:  # noqa: BLE001
            print(f"!! Putaran retensi gagal: {e}")
        henti.wait(6 * 3600)


@contextlib.asynccontextmanager
async def _lifespan(_app):
    db.tandai_terhenti()  # pemeriksaan yang berjalan saat server mati tidak mungkin lanjut
    henti = threading.Event()
    threading.Thread(target=_putaran_retensi, args=(henti,), daemon=True).start()
    yield
    henti.set()


app = FastAPI(title="PENScek API", lifespan=_lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173",
                   *[o.strip() for o in os.getenv("PENSCEK_ORIGINS", "").split(",") if o.strip()]],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _pengguna(authorization: str = Header(default="")) -> dict:
    token = authorization[7:].strip() if authorization.lower().startswith("bearer ") else ""
    try:
        user = db.pengguna(token) if token else None
    except RuntimeError as e:  # konfigurasi Supabase belum lengkap
        raise HTTPException(503, str(e))
    if not user:
        raise HTTPException(401, "Masuk dulu untuk memakai PENScek.")
    return user


def _save(job):
    """Sinkronkan status pemeriksaan ke Supabase. Kegagalan dicatat, tidak menghentikan pemeriksaan."""
    kolom = {"status": job["status"], "error": job.get("error"), "decomp": job.get("decomp"),
             "judul": (job.get("decomp") or {}).get("judul") or job.get("judul") or ""}
    if job.get("skor") is not None:
        kolom["skor"] = job["skor"]
    if job["status"] in ("done", "error"):
        kolom["finished_at"] = _now()
    try:
        db.perbarui(job["id"], **kolom)
    except Exception as e:  # noqa: BLE001
        print(f"!! Supabase gagal menyimpan status {job['id']}: {e}")


def _nama_lokal(nama: str) -> str:
    return re.sub(r"[^\w.\- ()\[\]]", "_", Path(nama).name)


def _dari_baris(r: dict) -> dict:
    return {"id": r["id"], "status": r["status"], "file": r["file_name"], "created": r["created_at"], "error": r.get("error"),
            "skor": r.get("skor"), "judul": r.get("judul") or "", "decomp": r.get("decomp"), "uploader_id": r.get("uploader_id"),
            "file_path": r.get("file_path"), "file_expires_at": r.get("file_expires_at"),
            "mhs_nama": r.get("mhs_nama") or ""}


def _lokal(job) -> Path:
    """Salinan dokumen di folder kerja; diunduh dari Supabase kalau belum ada (mis. setelah dibersihkan atau server pindah)."""
    p = DATA / job["id"] / _nama_lokal(job["file"])
    if not p.exists():
        if not job.get("file_path"):
            raise HTTPException(410, "Dokumen sudah melewati masa simpan dan dihapus. Skor dan temuannya tetap tersimpan; "
                                     "unggah ulang dokumennya kalau butuh pratinjau atau unduhan.")
        _folder_kerja(job["id"])
        db.unduh_file(job["file_path"], p)
    return p


class _Tee(io.StringIO):
    """Tangkap keluaran print() pipeline sambil tetap menampilkannya di terminal server."""

    def __init__(self, asli):
        super().__init__()
        self.asli = asli

    def write(self, s):
        try:
            self.asli.write(s)
        except Exception:  # noqa: BLE001  (mis. karakter yang tidak bisa dicetak di konsol Windows)
            pass
        return super().write(s)

    def flush(self):
        try:
            self.asli.flush()
        except Exception:  # noqa: BLE001
            pass


def _catatan(log: str) -> list[str]:
    """Baris '!! ...' dari pipeline = alasan sebuah tahap tidak berjalan; ditampilkan ke pengguna."""
    baris = [b.strip()[2:].strip() for b in log.splitlines() if b.strip().startswith("!!")]
    return list(dict.fromkeys(baris))[:5]


@contextlib.contextmanager
def _tangkap():
    """Tangkap print() pipeline. input() di pipeline tidak boleh menunggu terminal server: stdin dikosongkan,
    sehingga pipeline berhenti dengan pesan biasa dan pemeriksaan format tetap selesai."""
    awal = sys.stdin, sys.stdout, sys.stderr
    log = _Tee(sys.stdout)
    sys.stdin, sys.stdout, sys.stderr = io.StringIO(""), log, log
    try:
        yield log
    finally:
        sys.stdin, sys.stdout, sys.stderr = awal


def _tahap_judul(job_id: str):
    """Tahap 1: ambil judul dari sampul (kalau belum diisi) dan pecah jadi Problem Domain / Problem / Uniqueness.
    Hasilnya hanya usulan; pemeriksaan penuh baru jalan setelah user mengonfirmasi (POST .../konfirmasi)."""
    job = JOBS[job_id]
    job.update(status="running", started=_now())
    _save(job)
    decomp = {"judul": job.get("judul") or "", "problem_domain": "", "problem": "", "uniqueness": "",
              "catatan": "", "perlu_dicek": [], "judul_pasti": bool(job.get("judul"))}
    with _tangkap():
        try:
            import run_all

            base = run_all._sistematika_dir()  # juga menambahkan folder pipeline ke sys.path
            sys.path.insert(0, str(base / "llm"))
            from doc_input import extract_title
            from title_decomposer import decompose_title

            if not decomp["judul"]:
                decomp["judul"], decomp["judul_pasti"] = extract_title(job["path"])
            if decomp["judul"]:
                decomp.update(decompose_title(decomp["judul"]))
            else:
                decomp["catatan"] = "Judul tidak terbaca dari sampul. Isi judul dan ketiga unsur di bawah."
        except BaseException as e:  # noqa: BLE001  (termasuk SystemExit dari pipeline)
            # Pemecahan otomatis gagal (LLM mati, kuota habis, dst): user tetap bisa mengisi sendiri.
            pesan = str(e.code) if isinstance(e, SystemExit) and isinstance(e.code, str) else f"{type(e).__name__}: {e}"
            decomp["catatan"] = "Pemecahan judul otomatis gagal, isi manual. " + pesan.lstrip("! ")[:200]
    job.update(status="menunggu", decomp=decomp)
    _save(job)


def _jalankan(job_id: str):
    """Tahap 2: pemeriksaan penuh (format + sistematika) dengan judul dan tiga unsur yang sudah dikonfirmasi."""
    job = JOBS[job_id]
    job.update(status="running", started=_now())
    _save(job)

    with _tangkap() as log:
        try:
            import run_all

            d = job["decomp"]
            extra = ["--yes", "--retrieval", RETRIEVAL, "--judul", d["judul"], "--problem-domain", d["problem_domain"],
                     "--problem", d["problem"], "--uniqueness", d["uniqueness"]]
            run_dir = Path(run_all.run(job["path"], extra))
            hasil = build_result(run_dir, job["file"], _catatan(log.getvalue()))
            hasil["id"], hasil["waktu"] = job_id, job["created"]
            skor = hasil["skor"]["total"]
            db.perbarui(job_id, status="done", skor=skor, result=hasil, error=None, finished_at=_now())
            job.update(status="done", skor=skor)
            if HAPUS_RUN and "runs" in run_dir.parts:  # hasilnya sudah di Supabase; folder run hanya menumpuk di disk
                shutil.rmtree(run_dir, ignore_errors=True)
            try:
                _pdf_pratinjau(job)
            except Exception as e:  # noqa: BLE001  (pratinjau opsional; hasil pemeriksaan tetap tersedia)
                print(f"!! Pratinjau PDF tidak disiapkan: {e}")
        except SystemExit as e:  # run_all/pipeline berhenti lewat sys.exit(pesan)
            job.update(status="error", error=str(e.code).lstrip("! ") if isinstance(e.code, str) else f"Pemeriksaan berhenti (kode {e.code}).")
        except Exception as e:  # noqa: BLE001
            job.update(status="error", error=f"{type(e).__name__}: {e}")
    _save(job)
    JOBS.pop(job_id, None)


_konversi_lock = threading.Lock()


def _pdf_pratinjau(job) -> Path:
    """PDF yang ditampilkan di halaman Hasil: berkas PDF unggahan, atau hasil konversi (di-cache) untuk .docx."""
    sumber = _lokal(job)
    if sumber.suffix.lower() == ".pdf":
        return sumber
    tujuan = sumber.parent / "pratinjau.pdf"
    with _konversi_lock:  # satu konversi LibreOffice sekali jalan
        if not tujuan.exists():
            docx_ke_pdf(sumber, tujuan)
    return tujuan


def _hasil_rapi(row: dict) -> dict:
    """Hasil dengan uraian temuan yang sudah dibahasakan (juga untuk hasil lama yang tersimpan mentah)."""
    return rapikan_hasil(copy.deepcopy(row["result"]))


def _publik(job, dengan_hasil=False, row=None):
    out = {k: job.get(k) for k in ("id", "status", "file", "created", "error", "skor", "mhs_nama")}
    out["dokumen_tersedia"] = bool(job.get("file_path"))
    out["dokumen_hingga"] = job.get("file_expires_at") if job.get("file_path") else None
    if job.get("status") == "menunggu":
        out["decomp"] = job.get("decomp")
    if dengan_hasil and job.get("status") == "done":
        out["result"] = _hasil_rapi(row) if row and row.get("result") else None
    return out


def _baris_selesai(job_id: str, user: dict) -> dict:
    row = db.ambil(job_id, user["id"])
    if not row:
        raise HTTPException(404, "Pemeriksaan tidak ditemukan.")
    if row.get("status") != "done":
        raise HTTPException(409, "Pemeriksaan belum selesai.")
    return row


@app.get("/api/health")
def health():
    return {"ok": True, "retrieval": RETRIEVAL}


@app.get("/api/me")
def saya(user: dict = Depends(_pengguna)):
    return {k: user[k] for k in ("id", "email", "role", "nama")}


@app.post("/api/checks", status_code=202)
async def buat_pemeriksaan(file: UploadFile = File(...), judul: str = Form(""),
                           mhs_nama: str = Form(""), user: dict = Depends(_pengguna)):
    nama = Path(file.filename or "").name
    if Path(nama).suffix.lower() not in ALLOWED:
        raise HTTPException(400, "Format tidak didukung. Unggah file .pdf atau .docx.")

    if user["role"] == "dsn":  # dosen mengunggah dokumen mahasiswa: nama mahasiswanya diisi di sini
        mhs_nama = " ".join(mhs_nama.split())
        if not mhs_nama:
            raise HTTPException(422, "Isi nama mahasiswa yang dokumennya diperiksa.")
    else:  # mahasiswa memeriksa dokumennya sendiri
        mhs_nama = user["nama"]

    job_id = uuid.uuid4().hex[:12]
    folder = _folder_kerja(job_id)
    tujuan = folder / _nama_lokal(nama)

    ukuran = 0
    try:
        with open(tujuan, "wb") as out:
            while chunk := await file.read(1024 * 1024):
                ukuran += len(chunk)
                if ukuran > MAX_MB * 1024 * 1024:
                    raise HTTPException(413, f"Ukuran file melebihi {MAX_MB} MB.")
                out.write(chunk)
    except HTTPException:
        shutil.rmtree(folder, ignore_errors=True)
        raise

    lokasi = db.lokasi_file(user["id"], job_id, tujuan.suffix)
    try:
        db.unggah_file(tujuan, lokasi)
        db.buat(job_id, user["id"], nama, judul.strip(), mhs_nama, lokasi)
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(folder, ignore_errors=True)
        with contextlib.suppress(Exception):
            db.hapus_file(lokasi)
        raise HTTPException(502, f"Dokumen gagal disimpan ke penyimpanan: {e}")

    job = {"id": job_id, "status": "queued", "file": nama, "path": str(tujuan), "judul": judul.strip(),
           "created": _now(), "error": None, "skor": None, "uploader_id": user["id"], "file_path": lokasi,
           "mhs_nama": mhs_nama}
    JOBS[job_id] = job
    _pool.submit(_tahap_judul, job_id)
    return {"id": job_id}


@app.get("/api/checks/{job_id}")
def lihat_pemeriksaan(job_id: str, user: dict = Depends(_pengguna)):
    job = JOBS.get(job_id)
    if job and job["uploader_id"] == user["id"] and job["status"] != "done":
        return _publik(job)  # sedang berjalan: cukup dari memori, polling tidak perlu ke Supabase
    row = db.ambil(job_id, user["id"])
    if not row:
        raise HTTPException(404, "Pemeriksaan tidak ditemukan.")
    return _publik(_dari_baris(row), dengan_hasil=True, row=row)


class Konfirmasi(BaseModel):
    judul: str
    problem_domain: str
    problem: str
    uniqueness: str


@app.post("/api/checks/{job_id}/konfirmasi", status_code=202)
def konfirmasi(job_id: str, body: Konfirmasi, user: dict = Depends(_pengguna)):
    job = JOBS.get(job_id)
    if not job or job["uploader_id"] != user["id"]:
        row = db.ambil(job_id, user["id"])  # server sempat dimulai ulang saat menunggu konfirmasi
        if not row:
            raise HTTPException(404, "Pemeriksaan tidak ditemukan.")
        job = _dari_baris(row)
        if job["status"] == "menunggu":
            job["path"] = str(_lokal(job))
            JOBS[job_id] = job
    if job["status"] != "menunggu":
        raise HTTPException(409, "Pemeriksaan ini tidak sedang menunggu konfirmasi.")
    isi = {k: " ".join(v.split()) for k, v in body.model_dump().items()}
    kosong = [k for k, v in isi.items() if not v]
    if kosong:
        raise HTTPException(422, "Masih ada isian yang kosong: " + ", ".join(kosong) + ".")
    job["decomp"] = {**job.get("decomp", {}), **isi, "dikonfirmasi": True}
    job["status"] = "queued"
    _save(job)
    _pool.submit(_jalankan, job_id)
    return {"id": job_id}


@app.get("/api/checks")
def daftar_pemeriksaan(user: dict = Depends(_pengguna)):
    return [_publik(_dari_baris(r)) for r in db.daftar(user["id"])]


@app.delete("/api/checks/{job_id}", status_code=204)
def hapus_pemeriksaan(job_id: str, user: dict = Depends(_pengguna)):
    job = JOBS.get(job_id)
    if job and job["uploader_id"] == user["id"] and job["status"] in ("queued", "running"):
        raise HTTPException(409, "Pemeriksaan sedang berjalan. Tunggu sampai selesai.")
    if not db.hapus(job_id, user["id"]):
        raise HTTPException(404, "Pemeriksaan tidak ditemukan.")
    JOBS.pop(job_id, None)
    shutil.rmtree(DATA / job_id, ignore_errors=True)
    return Response(status_code=204)


@app.get("/api/checks/{job_id}/pdf")
def pdf(job_id: str, user: dict = Depends(_pengguna)):
    row = _baris_selesai(job_id, user)
    try:
        path = _pdf_pratinjau(_dari_baris(row))
    except RuntimeError as e:
        raise HTTPException(503, str(e))
    return FileResponse(path, media_type="application/pdf", headers={"Content-Disposition": "inline"})


def _sorotan(row: dict) -> dict:
    tersimpan = row.get("sorotan")
    if tersimpan and tersimpan.get("versi") == VERSI_SOROTAN:  # sorotan dari versi lama (masih ke judul bab) dihitung ulang
        return tersimpan
    job = _dari_baris(row)
    try:
        path = _pdf_pratinjau(job)
        sumber = _lokal(job)
        # Untuk .docx, DOCX aslinya ikut dibaca: urutan tabel dan nomor halaman cetak hanya diketahui dari sana.
        data = hitung_sorotan(path, _hasil_rapi(row), docx_path=sumber if sumber.suffix.lower() == ".docx" else None)
    except RuntimeError as e:
        raise HTTPException(503, str(e))
    db.perbarui(row["id"], sorotan=data)
    return data


@app.get("/api/checks/{job_id}/sorotan")
def sorotan(job_id: str, user: dict = Depends(_pengguna)):
    return _sorotan(_baris_selesai(job_id, user))


@app.get("/api/checks/{job_id}/unduh/{jenis}")
def unduh(job_id: str, jenis: str, user: dict = Depends(_pengguna)):
    """Dokumen asli yang sudah disorot dan diberi komentar per temuan. Dibuat saat diminta dan tidak disimpan:
    hanya dokumen asli dan hasil pemeriksaannya yang disimpan."""
    if jenis not in ("pdf", "docx"):
        raise HTTPException(404, "Jenis unduhan harus pdf atau docx.")
    row = _baris_selesai(job_id, user)
    job = _dari_baris(row)
    sumber = _lokal(job)
    if jenis == "docx" and sumber.suffix.lower() != ".docx":
        raise HTTPException(422, "Unduhan DOCX hanya tersedia untuk dokumen yang diunggah sebagai .docx. Unduh versi PDF-nya.")

    sementara = Path(tempfile.mkdtemp(prefix="penscek-"))
    tujuan = sementara / f"beranotasi.{jenis}"
    try:
        hasil = _hasil_rapi(row)
        if jenis == "docx":
            docx_beranotasi(sumber, hasil, tujuan)
        else:
            pdf_beranotasi(_pdf_pratinjau(job), hasil, _sorotan(row), tujuan)
    except RuntimeError as e:
        shutil.rmtree(sementara, ignore_errors=True)
        raise HTTPException(503, str(e))
    except BaseException:
        shutil.rmtree(sementara, ignore_errors=True)
        raise
    dasar = re.sub(r"[^\w.\- ()]", "_", sumber.stem)
    nama = f"{dasar}-penscek.{jenis}"
    media = "application/pdf" if jenis == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(tujuan, media_type=media, filename=nama,
                        background=BackgroundTask(shutil.rmtree, sementara, ignore_errors=True))
