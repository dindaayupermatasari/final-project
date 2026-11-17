import { useEffect, useRef, useState } from "react";
import { Link, Navigate, useSearchParams } from "react-router-dom";
import Mascot from "../components/Mascot.jsx";
import Ring from "../components/Ring.jsx";
import PdfViewer from "../components/PdfViewer.jsx";
import { useAuth } from "../context/AuthContext.jsx";
import { ambil, sorotan as ambilSorotan, ambilPdf, unduhDokumen } from "../api.js";
import { inisial, levelSkor } from "../utils.js";

const bulat = (n) => (n == null ? null : Math.round(n));

const tanggal = (iso) =>
  iso ? new Date(iso).toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "numeric" }) : "-";

// Garis-garis abu sebagai sisa isi halaman, supaya pratinjau terlihat seperti lembar dokumen.
const Skeleton = ({ widths }) => (
  <div className="sk">{widths.map((w, i) => <i key={i} style={{ width: `${w}%` }} />)}</div>
);

export default function Hasil() {
  const [params] = useSearchParams();
  const id = params.get("id"); // hasil pemeriksaan dari backend
  // Dosen hanya melihat bagian yang salah: tanpa penjelasan perbaikan, acuan pedoman, dan unduhan.
  const { role } = useAuth();
  const tinjau = role === "dsn";

  const [data, setData] = useState(null);
  const [info, setInfo] = useState(null); // status pemeriksaan: nama mahasiswa, tanggal, dokumen_tersedia
  const dokumenAda = info?.dokumen_tersedia !== false; // false setelah dokumen dihapus karena melewati masa simpan
  const [error, setError] = useState("");
  const [unduh, setUnduh] = useState(""); // "docx" | "pdf" selama berkas dibuat
  const [menu, setMenu] = useState(false); // dropdown unduh terbuka
  const [errUnduh, setErrUnduh] = useState("");
  const menuRef = useRef(null);
  useEffect(() => {
    if (!menu) return;
    const tutup = (e) => { if (e.type === "keydown" ? e.key === "Escape" : !menuRef.current?.contains(e.target)) setMenu(false); };
    document.addEventListener("mousedown", tutup);
    document.addEventListener("keydown", tutup);
    return () => { document.removeEventListener("mousedown", tutup); document.removeEventListener("keydown", tutup); };
  }, [menu]);
  useEffect(() => {
    if (!id) return;
    let batal = false;
    setData(null); setInfo(null); setError("");
    ambil(id)
      .then((s) => {
        if (batal) return;
        if (s.result) { setData(s.result); setInfo(s); } else setError("Hasil pemeriksaan belum tersedia.");
      })
      .catch((e) => { if (!batal) setError(e.message); });
    return () => { batal = true; };
  }, [id]);

  // Posisi sorotan di PDF asli. Kalau PDF/sorotan tidak bisa disiapkan (mis. LibreOffice belum terpasang untuk .docx),
  // pratinjau kembali ke tampilan teks per temuan.
  const [sor, setSor] = useState(null);
  const [pdfGagal, setPdfGagal] = useState(false);
  useEffect(() => {
    if (!id || !data || !dokumenAda) return;
    let batal = false;
    setSor(null); setPdfGagal(false);
    ambilSorotan(id).then((s) => { if (!batal) setSor(s); }).catch(() => { if (!batal) setPdfGagal(true); });
    return () => { batal = true; };
  }, [id, data, dokumenAda]);

  const view = data;

  const [cur, setCur] = useState(0); // indeks BAB
  const [act, setAct] = useState(0); // indeks temuan aktif
  const markRef = useRef(null);

  // Gulung ke kalimat yang disorot saat temuan atau BAB berganti (bukan saat halaman pertama dibuka).
  const first = useRef(true);
  useEffect(() => {
    if (first.current) { first.current = false; return; }
    markRef.current?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [act, cur]);

  const back = tinjau
    ? <Link className="btn ghost" to="/dashboard">← Kembali ke dashboard</Link>
    : <Link className="btn ghost" to="/beranda">← Unggah baru</Link>;

  if (!id) return <Navigate to={tinjau ? "/dashboard" : "/beranda"} replace />;

  if (!view) {
    return (
      <>
        <p style={{ marginBottom: 16 }}>{back}</p>
        <div className="card" style={{ padding: 36, textAlign: "center", display: "grid", gap: 12, justifyItems: "center" }}>
          <Mascot pose={error ? "sleepy" : "idle"} size={110} />
          <p style={{ color: "var(--mute)" }}>{error || "Maskot lagi merapikan hasilnya, sebentar ya…"}</p>
        </div>
      </>
    );
  }

  const bab = view.bab;
  const score = bulat(view.skor.total);
  const b = bab[Math.min(cur, bab.length - 1)];
  const a = b?.f[act];

  // Temuan bab ini yang punya posisi di PDF; nomor urutnya sama dengan daftar temuan di kanan.
  const temuanPdf = sor && b
    ? b.f.map((_, k) => ({ k, ...sor.temuan[`${Math.min(cur, bab.length - 1)}-${k}`] })).filter((t) => t.rects)
    : [];
  const modePdf = !!id && dokumenAda && !pdfGagal;

  // Pratinjau halaman. Backend boleh mengirim `konteks` (daftar paragraf halaman temuan) dan `jumlah_halaman`;
  // kalau belum ada, yang tampil kalimat bermasalahnya saja di antara garis-garis abu.
  const hal = a?.p ?? (Number((a?.lokasi || "").match(/\d+/)?.[0]) || null);
  const total = view.halamanTotal ?? view.jumlah_halaman;
  const ketemu = !!a?.konteks?.length && !!a.q?.some((q) => a.konteks.some((t) => t.includes(q)));
  const paras = ketemu ? a.konteks : [a?.q?.length ? a.q.join(" ") : a?.lokasi || ""];
  // Temuan lain di halaman yang sama ikut ditandai, yang aktif paling tebal.
  const sama = a ? b.f.map((f, k) => ({ f, k })).filter(({ f, k }) => k === act || (f.p != null && f.p === a.p)) : [];
  let refTaken = false;
  const tandai = (text) => {
    const hits = [];
    sama.forEach(({ f, k }) => (f.q || []).forEach((q) => { const i = q ? text.indexOf(q) : -1; if (i >= 0) hits.push({ i, e: i + q.length, k }); }));
    hits.sort((x, y) => x.i - y.i);
    const out = [];
    let pos = 0;
    for (const h of hits) {
      if (h.i < pos) continue;
      if (h.i > pos) out.push(text.slice(pos, h.i));
      const on = h.k === act;
      const pakaiRef = on && !refTaken;
      if (pakaiRef) refTaken = true;
      out.push(
        <mark key={h.i} className={on ? "on" : ""} ref={pakaiRef ? markRef : null} onClick={() => setAct(h.k)}>
          {text.slice(h.i, h.e)}<sup>{h.k + 1}</sup>
        </mark>
      );
      pos = h.e;
    }
    out.push(text.slice(pos));
    return out;
  };

  // Dokumen yang diunggah + sorotan + komentar, dibuat backend.
  const laporan = async (jenis) => {
    setMenu(false); setErrUnduh(""); setUnduh(jenis);
    try {
      const blob = await unduhDokumen(id, jenis);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${String(view.file).replace(/\.[^.]+$/, "")}-penscek.${jenis}`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    } catch (e) {
      setErrUnduh(e.message);
    } finally { setUnduh(""); }
  };
  // DOCX asli hanya bisa dibuat dari dokumen yang diunggah sebagai .docx.
  const docxBisa = dokumenAda && /\.docx$/i.test(view.file || "");

  return (
    <>
      <p style={{ marginBottom: 16 }}>{back}</p>

      <div className="top">
        {tinjau ? (
          <div className="card doc">
            <span className="av lg" aria-hidden="true">{inisial(info?.mhs_nama)}</span>
            <div>
              <h2>{info?.mhs_nama || "Mahasiswa"}</h2>
              <p className="meta">Diperiksa {tanggal(info?.created ?? view.waktu)}</p>
              <p className="meta">{view.file}</p>
            </div>
          </div>
        ) : (
        <div className="card doc">
          <div>
            <h2>{view.file}</h2>
            <p className="meta">Diperiksa {tanggal(view.waktu ?? info?.created)}{view.judul ? ` · ${view.judul}` : ""}</p>
          </div>
          <div className="acts">
            <Link className="btn ghost" to="/panduan">Panduan buku PA</Link>
            <div className="dd" ref={menuRef}>
              <button className="btn" type="button" aria-haspopup="menu" aria-expanded={menu} disabled={!!unduh} onClick={() => setMenu((m) => !m)}>
                <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>
                {unduh ? "Menyiapkan…" : "Unduh"}
                <svg className="cv" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="6 9 12 15 18 9" /></svg>
              </button>
              {menu && (
                <div className="dd-menu" role="menu">
                  <button type="button" role="menuitem" disabled={!dokumenAda} onClick={() => laporan("pdf")}>
                    <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /></svg>
                    PDF<small>{dokumenAda ? "Dokumen asli dengan sorotan dan komentar" : "Dokumen asli sudah dihapus"}</small>
                  </button>
                  <button type="button" role="menuitem" disabled={!docxBisa} onClick={() => laporan("docx")}>
                    <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="8" y1="13" x2="16" y2="13" /><line x1="8" y1="17" x2="16" y2="17" /></svg>
                    DOCX<small>{!dokumenAda ? "Dokumen asli sudah dihapus" : !docxBisa ? "Hanya untuk dokumen yang diunggah sebagai .docx" : "Dokumen asli dengan sorotan dan komentar Word"}</small>
                  </button>
                </div>
              )}
            </div>
          </div>
          {errUnduh && <p className="dd-err" role="alert">{errUnduh}</p>}
        </div>
        )}
        <div className="score">
          <small>Skor Kesesuaian Keseluruhan</small>
          {score == null ? <p style={{ color: "var(--mute)" }}>Belum tersedia</p> : <Ring score={score} color="#fff" />}
        </div>
      </div>

      {tinjau && (
        <div className="card tips tinjau" style={{ marginBottom: 20 }}>
          <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" /><circle cx="12" cy="12" r="3" /></svg>
          <span><b>Mode tinjau.</b> Hanya bagian yang bermasalah yang ditandai. Penjelasan dan cara memperbaikinya hanya tampil untuk mahasiswa.</span>
        </div>
      )}

      {!dokumenAda && (
        <div className="card tips" style={{ marginBottom: 20 }}>
          <b>Dokumen dihapus</b><span>Dokumen asli sudah dihapus dari server karena melewati masa simpan, jadi pratinjau dan unduhan tidak tersedia. Hasil pemeriksaan tetap tersimpan.</span>
        </div>
      )}

      {view.catatan?.length > 0 && (
        <div className="card tips" style={{ marginBottom: 20 }}>
          <b>Catatan</b><span>{view.catatan.join(" ")}</span>
        </div>
      )}

      {!b ? (
        <div className="card" style={{ padding: 36, textAlign: "center", color: "var(--mute)" }}>Tidak ada bab yang bisa ditampilkan.</div>
      ) : (
        <>
          <div className="card tabs" role="tablist" aria-label="Pilih bab, skor kesesuaian per bab">
            {bab.map((x, j) => (
              <button key={x.n} role="tab" aria-selected={j === cur} onClick={() => { setCur(j); setAct(0); }}>
                <span className="tr"><span>{x.n}</span><b>{bulat(x.s) ?? "–"}%</b></span>
                <span className="bar-t"><i className={x.s != null ? levelSkor(x.s) : ""} style={{ width: `${x.s ?? 0}%` }} /></span>
              </button>
            ))}
          </div>

          <div className="cols">
            <div className="card pv">
              <div className="hd">Pratinjau Dokumen{modePdf ? "" : ` ${b.n}`}<span>{modePdf ? (sor ? `${sor.halaman} halaman` : "") : (hal ? `Halaman ${hal}${total ? ` dari ${total}` : ""}` : "")}</span></div>
              {modePdf ? (
                sor ? (
                  <PdfViewer
                    kunci={id} muat={() => ambilPdf(id)} ukuran={sor.ukuran} temuan={temuanPdf} aktif={act}
                    tandaiGulir={`${cur}-${act}`} onPilih={setAct} onGagal={() => setPdfGagal(true)}
                  />
                ) : (
                  <p style={{ padding: 24, color: "var(--mute)" }}>Menyiapkan dokumen…</p>
                )
              ) : (
              <div className="page" tabIndex={0}>
                {a ? (
                  <>
                    {!ketemu && <Skeleton widths={[97, 92, 99, 88]} />}
                    {paras.map((t, i) => <p key={i}>{tandai(t)}</p>)}
                    {!ketemu && <Skeleton widths={[94, 98, 90, 96, 64]} />}
                  </>
                ) : (
                  <p>Tidak ada temuan pada bab ini.</p>
                )}
              </div>
              )}
            </div>

            <div className="card">
              <div className="hd">Temuan ketidaksesuaian <span>{b.f.length} temuan</span></div>
              <ul className="find">
                {b.f.map((f, k) => (
                  <li key={`${f.t}-${k}`} className={k === act ? "on" : ""}>
                    <button className="h" aria-expanded={k === act} onClick={() => setAct(k)}>
                      <span className="n">{k + 1}</span><span className="t">{f.t}</span>
                      {(() => { // nomor halaman di pratinjau (urutan halaman PDF); nomor dari pemeriksa dipakai hanya bila sorotan belum ada
                        const hlm = sor?.temuan?.[`${Math.min(cur, bab.length - 1)}-${k}`]?.hal ?? f.p; return hlm ? <span className="pg">hlm. {hlm}</span> : null; })()}
                    </button>
                    {k === act && (tinjau
                      ? (<div className="flag"><small>Bagian yang ditandai</small>{f.q?.length ? <q>{f.q.join(" … ")}</q> : <span>{f.lokasi}</span>}</div>)
                      : (<><p>{f.d}</p>{f.r && <div className="ref">{f.r}</div>}</>))}
                  </li>
                ))}
              </ul>
              {!b.f.length && (
                <div className="nofind"><Mascot pose="happy" size={84} /><p>Tidak ada temuan pada bab ini.</p></div>
              )}
            </div>
          </div>
        </>
      )}
    </>
  );
}
