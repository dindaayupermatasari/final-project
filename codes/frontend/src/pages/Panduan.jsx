import Mascot from "../components/Mascot.jsx";

// File asli ada di public/panduan/. Ganti nama/edisi di sini bila pedoman direvisi.
const PDF = "/panduan/Panduan_Buku_PA_SarjanaTerapan_A4_Rev05.pdf";
const PAGES = 35;

export default function Panduan() {
  return (
    <>
      <div className="pagehead">
        <Mascot pose="wink" size={150} style={{ position: "absolute", right: 0, top: 0 }} />
        <h1>Template panduan</h1>
        <p>Pedoman penulisan yang dipakai sebagai acuan, serta cara menggunakan sistem ini.</p>
      </div>
      <div className="pan">
        <div className="card in">
          <p style={{ fontWeight: 600, color: "var(--b)" }}>Referensi dokumen</p>
          <h2 style={{ fontSize: 30, margin: "6px 0 20px" }}>Pedoman penulisan yang digunakan</h2>
          <span className="tag">Rev05 · Sarjana Terapan</span>
          <h3 style={{ fontSize: 24, margin: "12px 0 8px" }}>Panduan Buku Proyek Akhir PENS</h3>
          <p style={{ color: "var(--mute)" }}>Dokumen ini menjadi acuan utama sistem dalam memeriksa kesesuaian format visual dan sistematika penulisan buku proyek akhir.</p>
          <div className="facts">
            <div><small>Edisi tahun</small><b>2025</b></div>
            <div><small>Institusi</small><b>Politeknik Elektronika Negeri Surabaya</b></div>
            <div><small>Jumlah halaman</small><b>{PAGES} halaman</b></div>
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <a className="btn" href={PDF} download>Unduh file</a>
            <a className="btn ghost" href={PDF} target="_blank" rel="noreferrer">Lihat dokumen</a>
          </div>
        </div>
        <div className="card in">
          <div className="hd" style={{ padding: "0 0 12px", border: 0 }}>Preview dokumen <span>{PAGES} halaman</span></div>
          <iframe
            title="Panduan Buku Proyek Akhir PENS"
            src={`${PDF}#page=1&view=FitH&toolbar=0`}
            style={{ width: "100%", aspectRatio: "3/4", border: "1.5px solid var(--line)", borderRadius: 8, background: "#fff" }}
          />
          <p style={{ marginTop: 10, fontSize: 13, color: "var(--mute)" }}>
            Preview tidak tampil? <a className="link" href={PDF} target="_blank" rel="noreferrer">Buka PDF di tab baru</a>.
          </p>
        </div>
      </div>
      <div className="card note"><b aria-hidden="true">ⓘ</b> Sistem akan diperbarui apabila ada revisi pedoman dari institusi.</div>
    </>
  );
}
