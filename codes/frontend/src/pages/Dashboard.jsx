import { useEffect, useMemo, useRef, useState } from "react";
import { Link } from "react-router-dom";
import Mascot from "../components/Mascot.jsx";
import Ring from "../components/Ring.jsx";
import { riwayat } from "../api.js";
import { inisial, label, levelSkor, tanggalPendek } from "../utils.js";

const PER = 4;
const LAYAK = 80; // skor minimal agar dokumen dinyatakan layak

const FILTER = [["", "Semua status"], ["ok", "Layak"], ["rev", "Perlu revisi"]];

const Ikon = ({ children }) => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{children}</svg>
);

// Satu baris per mahasiswa (nama, tanpa membedakan huruf besar kecil): pemeriksaan selesai yang terbaru.
// Pemeriksaan yang belum selesai belum punya skor.
const kunciNama = (n) => String(n || "").trim().replace(/\s+/g, " ").toLowerCase();
function kelompokkan(daftar) {
  const per = new Map();
  for (const c of daftar) {
    const k = kunciNama(c.mhs_nama);
    if (c.status !== "done" || c.skor == null || !k) continue;
    const g = per.get(k) || { key: k, nama: c.mhs_nama, id: c.id, skor: Math.round(c.skor), created: c.created, jumlah: 0 };
    g.jumlah += 1;
    if (String(c.created) > String(g.created) || g.jumlah === 1) {
      Object.assign(g, { nama: c.mhs_nama, id: c.id, skor: Math.round(c.skor), created: c.created });
    }
    per.set(k, g);
  }
  return [...per.values()].sort((a, b) => String(b.created).localeCompare(String(a.created)));
}

export default function Dashboard() {
  const [data, setData] = useState(null); // null = memuat
  const [error, setError] = useState("");
  const [q, setQ] = useState("");
  const [status, setStatus] = useState("");
  const [page, setPage] = useState(1);
  const [menu, setMenu] = useState(false);
  const menuRef = useRef(null);

  // Menu filter menutup saat klik di luar atau tekan Escape.
  useEffect(() => {
    if (!menu) return;
    const tutup = (e) => { if (e.type === "keydown" ? e.key === "Escape" : !menuRef.current?.contains(e.target)) setMenu(false); };
    document.addEventListener("mousedown", tutup);
    document.addEventListener("keydown", tutup);
    return () => { document.removeEventListener("mousedown", tutup); document.removeEventListener("keydown", tutup); };
  }, [menu]);

  useEffect(() => {
    let batal = false;
    riwayat()
      .then((d) => { if (!batal) setData(kelompokkan(d)); })
      .catch((e) => { if (!batal) { setError(e.message); setData([]); } });
    return () => { batal = true; };
  }, []);

  const mhs = data || [];

  // Angka ringkasan dihitung dari daftar yang sama dengan tabel, jadi tidak pernah selisih.
  const stats = useMemo(() => {
    const total = mhs.length;
    const layak = mhs.filter((m) => m.skor >= LAYAK).length;
    const rata = mhs.reduce((a, m) => a + m.skor, 0) / (total || 1);
    return { total, layak, revisi: total - layak, rata, persen: Math.round((layak / (total || 1)) * 100) };
  }, [mhs]);

  const rows = useMemo(() => {
    const k = q.trim().toLowerCase();
    return mhs.filter((m) =>
      (!k || m.key.includes(k.replace(/\s+/g, " "))) &&
      (!status || (status === "ok") === (m.skor >= LAYAK))
    );
  }, [mhs, q, status]);

  const pages = Math.max(1, Math.ceil(rows.length / PER));
  const current = Math.min(page, pages);
  const dari = (current - 1) * PER;
  const slice = rows.slice(dari, dari + PER);

  const kartu = [
    { t: "Total mahasiswa", v: stats.total, h: "Sudah diperiksa", i: <><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" /></> },
    { t: "Dokumen layak", v: stats.layak, h: `${stats.persen}% dari total`, i: <><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" /><polyline points="22 4 12 14.01 9 11.01" /></> },
    { t: "Perlu revisi", v: stats.revisi, h: `Skor di bawah ${LAYAK}%`, i: <><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" /><line x1="12" y1="9" x2="12" y2="13" /><line x1="12" y1="17" x2="12.01" y2="17" /></> },
    { t: "Rata-rata skor", v: `${stats.rata.toFixed(1).replace(".", ",")}%`, h: `Batas layak ${LAYAK}%`, i: <><line x1="18" y1="20" x2="18" y2="10" /><line x1="12" y1="20" x2="12" y2="4" /><line x1="6" y1="20" x2="6" y2="14" /></> },
  ];

  return (
    <>
      <div className="stats">
        {kartu.map((k) => (
          <div className="card stat" key={k.t}>
            <div className="stat-h"><span className="stat-i"><Ikon>{k.i}</Ikon></span><span className="stat-n">{k.h}</span></div>
            <small>{k.t}</small>
            <b>{k.v}</b>
          </div>
        ))}
      </div>

      <div className="tools">
        <div className="tools-t">
          <h2>Mahasiswa</h2>
          <p>Skor terbaru tiap mahasiswa dari dokumen yang kamu unggah. Unggah dokumen baru dari Beranda.</p>
        </div>
        <div className="tools-f">
          <div className="srch">
            <label className="sr" htmlFor="q">Cari mahasiswa</label>
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="7" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>
            <input id="q" type="search" placeholder="Cari nama mahasiswa" value={q} onChange={(e) => { setQ(e.target.value); setPage(1); }} />
          </div>
          <div className="dd" ref={menuRef}>
            <button className={`btn ghost flt${status ? " aktif" : ""}`} type="button" aria-haspopup="menu" aria-expanded={menu} onClick={() => setMenu((m) => !m)}>
              <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><line x1="4" y1="6" x2="20" y2="6" /><line x1="7" y1="12" x2="17" y2="12" /><line x1="10" y1="18" x2="14" y2="18" /></svg>
              {status ? FILTER.find(([v]) => v === status)[1] : "Filter"}
            </button>
            {menu && (
              <div className="dd-menu fm" role="menu">
                {FILTER.map(([v, t]) => (
                  <button key={v} type="button" role="menuitemradio" aria-checked={v === status} onClick={() => { setStatus(v); setPage(1); setMenu(false); }}>
                    <span className="ck" aria-hidden="true">{v === status ? "✓" : ""}</span>{t}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="card tw">
        <table>
          <thead><tr><th>Nama mahasiswa</th><th>Tanggal pengajuan</th><th>Skor kesesuaian</th><th>Status</th><th><span className="sr">Aksi</span></th></tr></thead>
          <tbody>
            {slice.length ? slice.map((m) => (
              <tr key={m.key}>
                <td><div className="who"><span className="av">{inisial(m.nama)}</span><div><b>{m.nama}</b><small>{m.jumlah} unggahan</small></div></div></td>
                <td>{tanggalPendek(m.created)}</td>
                <td><div className={`sc ${levelSkor(m.skor)}`}><Ring score={m.skor} />{label(m.skor)}</div></td>
                <td><span className={`pill ${m.skor < LAYAK ? "rev" : ""}`.trim()}>{m.skor >= LAYAK ? "Layak" : "Perlu revisi"}</span></td>
                <td>
                  <Link className="link" to={`/hasil?id=${m.id}`} aria-label={`Lihat detail ${m.nama}`}>
                    Lihat detail
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><line x1="5" y1="12" x2="19" y2="12" /><polyline points="12 5 19 12 12 19" /></svg>
                  </Link>
                </td>
              </tr>
            )) : (
              <tr><td colSpan={5} className="kosong">
                <Mascot pose="confused" size={96} badge="?" /><br />{data === null ? "Memuat data…" : error ? error : mhs.length ? "Tidak ada mahasiswa yang cocok. Ubah kata kunci atau filter status." : "Belum ada dokumen yang diperiksa. Unggah dokumen mahasiswa dari Beranda."}
              </td></tr>
            )}
          </tbody>
        </table>
        <div className="pg-f">
          <span>{rows.length ? `Menampilkan ${dari + 1}–${dari + slice.length} dari ${rows.length} mahasiswa` : "Tidak ada data"}</span>
          <div role="navigation" aria-label="Halaman tabel">
            <button type="button" aria-label="Halaman sebelumnya" disabled={current === 1} onClick={() => setPage(current - 1)}>‹</button>
            {Array.from({ length: pages }, (_, n) => (
              <button key={n} type="button" aria-current={n + 1 === current ? "page" : undefined} onClick={() => setPage(n + 1)}>{n + 1}</button>
            ))}
            <button type="button" aria-label="Halaman berikutnya" disabled={current === pages} onClick={() => setPage(current + 1)}>›</button>
          </div>
        </div>
      </div>
    </>
  );
}
