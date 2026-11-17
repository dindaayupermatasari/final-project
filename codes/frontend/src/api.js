import { supabase } from "./supabase.js";

// Klien untuk backend PENScek (api/app.py). Alamat bisa diganti lewat VITE_API_URL saat build.
const BASE = import.meta.env.VITE_API_URL || "/api";

const SAMBUNGAN_GAGAL = "Tidak bisa terhubung ke server pemeriksa. Pastikan backend sudah berjalan.";

// Semua permintaan membawa token login. getSession() memperbarui token yang sudah kedaluwarsa.
async function minta(url, opsi = {}) {
  const { data } = await supabase.auth.getSession();
  const token = data.session?.access_token;
  const headers = { ...(opsi.headers || {}), ...(token ? { Authorization: `Bearer ${token}` } : {}) };

  let res;
  try {
    res = await fetch(`${BASE}${url}`, { ...opsi, headers });
  } catch {
    throw new Error(SAMBUNGAN_GAGAL);
  }
  if (!res.ok) {
    let pesan = `Server membalas ${res.status}.`;
    try { pesan = (await res.json()).detail || pesan; } catch { /* badan bukan JSON */ }
    if (res.status === 401) await supabase.auth.signOut(); // token ditolak: kembali ke halaman masuk
    throw new Error(pesan);
  }
  return res;
}

const kirim = async (url, opsi) => (await minta(url, opsi)).json();

export const saya = () => kirim("/me"); // -> { id, email, role, nama }

// Dosen mengunggah dokumen mahasiswa dan mengirim `mhsNama`. Mahasiswa tidak mengirimnya.
export function unggah(file, judul = "", mhsNama) {
  const f = new FormData();
  f.append("file", file);
  if (judul) f.append("judul", judul);
  if (mhsNama) f.append("mhs_nama", mhsNama);
  return kirim("/checks", { method: "POST", body: f }); // -> { id }
}

export const ambil = (id) => kirim(`/checks/${id}`); // -> { status, decomp?, result?, mhs_nama, dokumen_tersedia, ... }

// Setelah judul dipecah, pemeriksaan menunggu user: kirim judul dan tiga unsur yang sudah dikoreksi.
export const konfirmasi = (id, isi) =>
  kirim(`/checks/${id}/konfirmasi`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(isi),
  });
export const riwayat = () => kirim("/checks");

// Posisi sorotan temuan di PDF.
export const sorotan = (id) => kirim(`/checks/${id}/sorotan`); // -> { halaman, ukuran, temuan: {"bab-temuan": {...}} }

// PDF pratinjau (untuk .docx: hasil konversi di backend). Diambil utuh dengan token yang masih segar,
// karena pdf.js tidak bisa menyertakan header Authorization yang diperbarui di tengah jalan.
export async function ambilPdf(id) {
  return (await minta(`/checks/${id}/pdf`)).arrayBuffer();
}

// Tanya status tiap `jeda` ms sampai selesai atau sampai menunggu konfirmasi user ("menunggu").
// `berhenti()` dipanggil sebelum tiap tanya (mis. komponen sudah ditutup).
export async function tunggu(id, { berhenti, jeda = 2500 } = {}) {
  for (;;) {
    if (berhenti?.()) return null;
    const s = await ambil(id);
    if (s.status === "done" || s.status === "menunggu") return s;
    if (s.status === "error") throw new Error(s.error || "Pemeriksaan gagal.");
    await new Promise((r) => setTimeout(r, jeda));
  }
}

// Dokumen asli + sorotan + komentar temuan (dibuat di backend). jenis: "pdf" | "docx". -> Blob
export async function unduhDokumen(id, jenis) {
  return (await minta(`/checks/${id}/unduh/${jenis}`)).blob();
}
