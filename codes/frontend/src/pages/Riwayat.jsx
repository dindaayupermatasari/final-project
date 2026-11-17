import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import Mascot from "../components/Mascot.jsx";
import Ring from "../components/Ring.jsx";
import { riwayat } from "../api.js";
import { label, tanggalPendek } from "../utils.js";

const STATUS = { queued: "Antre", running: "Sedang diperiksa", menunggu: "Menunggu konfirmasi judul", error: "Gagal" };

export default function Riwayat() {
  const [daftar, setDaftar] = useState(null);
  const [error, setError] = useState("");

  useEffect(() => {
    let batal = false;
    riwayat()
      .then((d) => { if (!batal) setDaftar([...d].sort((a, b) => String(b.created).localeCompare(String(a.created)))); })
      .catch((e) => { if (!batal) setError(e.message); });
    return () => { batal = true; };
  }, []);

  return (
    <>
      <div className="pagehead">
        <Mascot pose="search" size={150} style={{ position: "absolute", right: 0, top: 0 }} />
        <h1>Riwayat pengecekan</h1>
        <p>Setiap unggahan tersimpan, jadi kamu bisa melihat skor naik dari draf ke draf. Dokumen aslinya disimpan 30 hari, hasil pemeriksaannya tetap ada.</p>
      </div>

      {error ? (
        <div className="card" style={{ padding: 28, textAlign: "center", color: "var(--mute)" }}>{error}</div>
      ) : daftar === null ? (
        <p style={{ color: "var(--mute)" }}>Memuat riwayat…</p>
      ) : daftar.length === 0 ? (
        <div className="card" style={{ padding: 36, textAlign: "center", display: "grid", gap: 12, justifyItems: "center" }}>
          <Mascot pose="idle" size={110} />
          <p style={{ color: "var(--mute)" }}>Belum ada pemeriksaan. Unggah draf pertamamu dari Beranda.</p>
          <Link className="btn" to="/beranda">Unggah dokumen</Link>
        </div>
      ) : (
        <div className="hist">
          {daftar.map((h) => {
            const selesai = h.status === "done" && h.skor != null;
            const skor = Math.round(h.skor);
            return (
              <div className="card hrow" key={h.id}>
                {selesai ? <Ring score={skor} /> : <Mascot pose={h.status === "error" ? "sleepy" : "idle"} size={56} />}
                <div>
                  <b>{h.file}</b>
                  <p style={{ color: "var(--mute)", fontSize: 14 }}>
                    Diperiksa {tanggalPendek(h.created)}{selesai ? ` · ${label(skor)}` : ` · ${STATUS[h.status] || h.status}`}
                    {selesai && !h.dokumen_tersedia ? " · dokumen sudah dihapus" : ""}
                  </p>
                </div>
                {selesai && <span className={`pill ${skor < 80 ? "rev" : ""}`.trim()}>{skor >= 80 ? "Layak" : "Perlu revisi"}</span>}
                {selesai && <Link className="btn ghost" to={`/hasil?id=${h.id}`}>Lihat hasil</Link>}
              </div>
            );
          })}
        </div>
      )}
    </>
  );
}
