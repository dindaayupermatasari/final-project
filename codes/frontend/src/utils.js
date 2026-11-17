export const label = (s) => (s >= 90 ? "Sangat baik" : s >= 80 ? "Baik" : s >= 70 ? "Cukup" : "Kurang");

// Warna indikator skor: hijau (>= 80), kuning (60-79), merah (< 60). Dipakai cincin skor dan bilah per bab.
export const levelSkor = (s) => (s >= 80 ? "ok" : s >= 60 ? "mid" : "lo");
export const warnaSkor = (s) => `var(--${levelSkor(s)})`;

// Aturan domain yang sama dengan trigger handle_new_user di database:
// dosen @pens.ac.id, mahasiswa @<prodi>.student.pens.ac.id (mis. ds.student.pens.ac.id).
export const emailPens = (e) => {
  const v = String(e || "").trim().toLowerCase();
  return /^[^@\s]+@pens\.ac\.id$/.test(v) || /^[^@\s]+@[a-z0-9.-]+\.student\.pens\.ac\.id$/.test(v);
};

export const inisial = (nama) =>
  String(nama || "")
    .split(/\s+/).filter(Boolean).slice(0, 2)
    .map((k) => k[0].toUpperCase()).join("") || "?";

export const tanggalPendek = (iso) =>
  iso ? new Date(iso).toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "numeric" }) : "-";
