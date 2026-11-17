"""Penyimpanan PENScek di Supabase: tabel `checks`, bucket `dokumen`, dan pengecekan token login.

Backend memakai service key, yang MELEWATI Row Level Security. Karena itu setiap fungsi yang membaca atau
mengubah data pengguna di sini wajib menyaring `uploader_id`; jangan menambah fungsi yang tidak menyaringnya.

Environment variable:
    SUPABASE_URL           https://<project>.supabase.co
    SUPABASE_SERVICE_KEY   secret key / service_role key (hanya di backend, jangan masuk frontend atau git)
    PENSCEK_RETENSI_HARI   lama dokumen disimpan sebelum dihapus (default 30)
"""
import os
import threading
import time
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from pathlib import Path

BUCKET = "dokumen"
RETENSI_HARI = int(os.getenv("PENSCEK_RETENSI_HARI", "30"))

_KOLOM_DAFTAR = "id,status,file_name,mhs_nama,judul,skor,created_at,file_path,file_expires_at,error"


@lru_cache(maxsize=1)
def _sb():
    from supabase import ClientOptions, create_client

    url, kunci = os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_SERVICE_KEY")
    if not url or not kunci:
        raise RuntimeError("SUPABASE_URL dan SUPABASE_SERVICE_KEY belum diisi di environment.")
    # Unggahan dokumen bisa sampai 50 MB, batas waktu bawaan storage (20 detik) terlalu pendek.
    return create_client(url, kunci, options=ClientOptions(storage_client_timeout=180, auto_refresh_token=False, persist_session=False))


def _sekarang():
    return datetime.now(timezone.utc)


# ---------- pengguna ----------

_cache_pengguna: dict[str, tuple[float, dict]] = {}
_kunci = threading.Lock()


def pengguna(token: str) -> dict | None:
    """Token login dari frontend -> {id, email, role, nama}, atau None kalau tidak sah.
    Hasilnya disimpan 60 detik supaya polling status tidak menanyai Supabase tiap 2,5 detik."""
    sekarang = time.time()
    with _kunci:
        hit = _cache_pengguna.get(token)
        if hit and hit[0] > sekarang:
            return hit[1]
    try:
        u = _sb().auth.get_user(token)
        u = u.user if u else None
    except Exception:  # noqa: BLE001  (token kedaluwarsa, salah, atau Supabase tidak terjangkau)
        return None
    if not u:
        return None
    baris = _sb().table("profiles").select("role,nama").eq("id", u.id).limit(1).execute().data
    p = baris[0] if baris else {}
    info = {"id": u.id, "email": u.email, "role": p.get("role", "mhs"), "nama": p.get("nama") or ""}
    with _kunci:
        if len(_cache_pengguna) > 500:
            _cache_pengguna.clear()
        _cache_pengguna[token] = (sekarang + 60, info)
    return info


# ---------- pemeriksaan ----------

def buat(job_id: str, uploader_id: str, file_name: str, judul: str, mhs_nama: str, file_path: str):
    kedaluwarsa = _sekarang() + timedelta(days=RETENSI_HARI)
    _sb().table("checks").insert({
        "id": job_id, "uploader_id": uploader_id, "file_name": file_name, "judul": judul,
        "mhs_nama": mhs_nama, "status": "queued",
        "file_path": file_path, "file_expires_at": kedaluwarsa.isoformat(),
    }).execute()


def perbarui(job_id: str, **kolom):
    _sb().table("checks").update(kolom).eq("id", job_id).execute()


def ambil(job_id: str, uploader_id: str) -> dict | None:
    baris = _sb().table("checks").select("*").eq("id", job_id).eq("uploader_id", uploader_id).limit(1).execute().data
    return baris[0] if baris else None


def daftar(uploader_id: str) -> list[dict]:
    return (_sb().table("checks").select(_KOLOM_DAFTAR).eq("uploader_id", uploader_id)
            .order("created_at", desc=True).execute().data)


def hapus(job_id: str, uploader_id: str) -> bool:
    baris = ambil(job_id, uploader_id)
    if not baris:
        return False
    if baris.get("file_path"):
        hapus_file(baris["file_path"])
    _sb().table("checks").delete().eq("id", job_id).eq("uploader_id", uploader_id).execute()
    return True


def tandai_terhenti():
    """Dipanggil saat server menyala: pemeriksaan yang masih 'queued'/'running' tidak mungkin lagi berjalan."""
    _sb().table("checks").update({
        "status": "error", "finished_at": _sekarang().isoformat(),
        "error": "Server dimulai ulang saat pemeriksaan berjalan. Unggah ulang dokumennya.",
    }).in_("status", ["queued", "running"]).execute()


# ---------- berkas ----------

def lokasi_file(uploader_id: str, job_id: str, ekstensi: str) -> str:
    # Nama berkas asli tidak dipakai di storage (spasi, kurung, dan huruf non-ASCII sering ditolak); nama asli ada di tabel.
    return f"{uploader_id}/{job_id}/dokumen{ekstensi.lower()}"


def unggah_file(lokal: Path, tujuan: str):
    tipe = "application/pdf" if lokal.suffix.lower() == ".pdf" else \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    _sb().storage.from_(BUCKET).upload(tujuan, lokal.read_bytes(), {"content-type": tipe, "upsert": "true"})


def unduh_file(sumber: str, lokal: Path):
    data = _sb().storage.from_(BUCKET).download(sumber)
    lokal.parent.mkdir(parents=True, exist_ok=True)
    lokal.write_bytes(data)


def hapus_file(path: str):
    _sb().storage.from_(BUCKET).remove([path])


def hapus_kadaluarsa() -> int:
    """Hapus dokumen yang melewati masa simpan. Baris (skor, temuan, tanggal) tetap ada.
    Lewat API Storage, bukan DELETE SQL, karena SQL hanya membuang catatannya dan meninggalkan berkas di disk storage."""
    baris = (_sb().table("checks").select("id,file_path").not_.is_("file_path", "null")
             .lt("file_expires_at", _sekarang().isoformat()).limit(200).execute().data)
    n = 0
    for b in baris:
        try:
            hapus_file(b["file_path"])
            perbarui(b["id"], file_path=None)
            n += 1
        except Exception as e:  # noqa: BLE001  (dicoba lagi pada putaran berikutnya)
            print(f"!! Gagal menghapus {b['file_path']}: {e}")
    return n
