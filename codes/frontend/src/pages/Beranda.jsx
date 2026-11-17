import { useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import Mascot from "../components/Mascot.jsx";
import { unggah, tunggu, konfirmasi, riwayat } from "../api.js";
import { useAuth } from "../context/AuthContext.jsx";

// Tiga tahap, tiap tahap punya beberapa versi kalimat; satu versi dipilih acak tiap pemeriksaan dimulai.
const STEPS = [
  [
    "Lagi baca bukumu dari awal, nyari di mana tiap bab mulai",
    "Buka daftar isi dulu biar tahu letak tiap bab",
    "Scroll dari halaman pertama, biar nggak ada yang kelewat",
    "Lagi misahin mana sampul, mana bab, mana daftar pustaka",
  ],
  [
    "Lagi ngecek margin, spasi, sama jenis huruf, satu per satu",
    "Ngukur margin tiap halaman, abis itu lanjut ke spasi dan jenis huruf",
    "Lagi lihat judul bab, nomor halaman, sama rata teksnya",
    "Cek font dan spasi, yang beda dikit pun ketahuan",
  ],
  [
    "Masih nyocokin sama pedoman PENS, bentar lagi kelar",
    "Tinggal bandingin temuannya sama buku pedoman",
    "Lagi ngerapihin hasilnya biar gampang dibaca",
  ],
];
// Selingan tiap beberapa detik supaya menunggu tidak terasa kosong; urutannya diacak tiap pemeriksaan.
const SELINGAN = [
  "Tenang, tidak ada halaman yang dilewati",
  "Boleh ambil minum dulu, belum kelar kok",
  "Jangan ditutup dulu tab-nya ya",
  "Daftar pustaka juga ikut diintip, lho",
  "Nomor halaman dan margin tidak akan lolos",
  "Dokumen yang tebal emang lebih lama, sabar ya",
  "Sambil nunggu, boleh regangin badan dulu",
  "Sambil menunggu, coba cek lagi apakah semua gambar punya keterangan",
  "Lagi serius nih, sabar dulu",
  "Spasi ganda dan huruf yang beda sedikit pun kelihatan di sini",
  "Tips dari aku, simpan dokumenmu dalam satu font yang sama dari awal sampai akhir",
  "Bab yang skornya paling rendah nanti kelihatan jelas, kok",
  "Ssst… ada yang lagi hitung kata abstrakmu",
];
const acak = (arr) => arr.map((x) => [Math.random(), x]).sort((p, q) => p[0] - q[0]).map(([, x]) => x);
const POSES = ["search", "surprised", "happy"];
const UNSUR = [
  ["problem_domain", "Problem Domain", "Bidang atau objek yang menjadi latar masalah, mis. “opinion mining pada sosial media”."],
  ["problem", "Problem", "Masalah yang diselesaikan, dengan pola “sulitnya …”."],
  ["uniqueness", "Uniqueness", "Solusi, metode, atau pendekatan yang ditawarkan."],
];
const kotak = { width: "100%", padding: "10px 12px", border: "1px solid var(--bt)", borderRadius: 10, font: "inherit", background: "#fff" };
const SEKALI = 12000; // ms per langkah teks; lamanya pemeriksaan asli tidak diketahui, jadi ini hanya penanda

export default function Beranda() {
  const navigate = useNavigate();
  const { role } = useAuth();
  const [mhsNama, setMhsNama] = useState(""); // khusus dosen: nama mahasiswa pemilik dokumen
  const [errForm, setErrForm] = useState("");
  const [namaLama, setNamaLama] = useState([]); // nama yang pernah diisi dosen, untuk saran ketikan
  const [file, setFile] = useState(null); // nama file yang sedang diperiksa
  const [step, setStep] = useState(0);
  const [over, setOver] = useState(false);
  const [full, setFull] = useState(false); // memicu transisi progress bar
  const [error, setError] = useState("");
  const [cek, setCek] = useState(null); // { id, isi, perlu } saat menunggu konfirmasi judul
  const [kirimKonfirmasi, setKirim] = useState(false);
  const timers = useRef([]);
  const alive = useRef(true);
  const [teksTahap, setTeksTahap] = useState(() => STEPS.map((v) => v[0]));
  const [urutan, setUrutan] = useState(() => acak(SELINGAN));

  useEffect(() => {
    alive.current = true;
    return () => { alive.current = false; timers.current.forEach(clearTimeout); };
  }, []);

  // Saran nama dari unggahan sebelumnya, supaya mahasiswa yang sama tidak terpecah karena salah ketik.
  useEffect(() => {
    if (role !== "dsn") return;
    let batal = false;
    riwayat()
      .then((d) => { if (!batal) setNamaLama([...new Set(d.map((c) => c.mhs_nama).filter(Boolean))].sort()); })
      .catch(() => { /* saran hanya pelengkap */ });
    return () => { batal = true; };
  }, [role]);

  const jalankanTimer = () => {
    setTeksTahap(STEPS.map((v) => v[Math.floor(Math.random() * v.length)]));
    setUrutan(acak(SELINGAN));
    setStep(0);
    setFull(false);
    requestAnimationFrame(() => requestAnimationFrame(() => setFull(true)));
    timers.current.forEach(clearTimeout);
    timers.current = [1, 2].map((i) => setTimeout(() => setStep(i), i * SEKALI));
  };

  // Menunggu server sampai selesai, atau sampai ia meminta konfirmasi judul.
  const tungguHasil = async (id, nama) => {
    const s = await tunggu(id, { berhenti: () => !alive.current });
    if (!s || !alive.current) return;
    timers.current.forEach(clearTimeout);
    if (s.status === "menunggu") {
      const d = s.decomp || {};
      setCek({ id, nama, isi: { judul: d.judul || "", problem_domain: d.problem_domain || "", problem: d.problem || "", uniqueness: d.uniqueness || "" },
               perlu: d.perlu_dicek || [], catatan: d.catatan || "", judulPasti: d.judul_pasti !== false,
               edit: !(d.judul && d.problem_domain && d.problem && d.uniqueness) }); // ada yang kosong: langsung isi manual
    } else {
      navigate(`/hasil?id=${id}`, { state: { file: nama } });
    }
  };

  const start = async (f) => {
    let mhs;
    if (role === "dsn") { // dosen wajib mengisi nama pemilik dokumen, supaya skornya masuk ke tabel dashboard
      mhs = mhsNama.trim().replace(/\s+/g, " ");
      if (!mhs) { setErrForm("Isi nama mahasiswa dulu, baru unggah dokumennya."); return; }
    }
    setErrForm("");
    setError("");
    setFile(f.name);
    jalankanTimer();
    try {
      const { id } = await unggah(f, "", mhs);
      await tungguHasil(id, f.name);
    } catch (e) {
      if (!alive.current) return;
      timers.current.forEach(clearTimeout);
      setError(e.message);
    }
  };

  const lanjut = async () => {
    const { id, nama, isi } = cek;
    setKirim(true);
    try {
      await konfirmasi(id, isi);
      setCek(null);
      jalankanTimer();
      await tungguHasil(id, nama);
    } catch (e) {
      if (alive.current) { setCek(null); setError(e.message); }
    } finally {
      if (alive.current) setKirim(false);
    }
  };

  const ubah = (k, v) => setCek((c) => ({ ...c, isi: { ...c.isi, [k]: v } }));
  const lengkap = cek && Object.values(cek.isi).every((v) => v.trim());

  const reset = () => { setFile(null); setError(""); setCek(null); setFull(false); setStep(0); };

  const busy = file !== null;
  const [selingan, setSelingan] = useState(0);
  useEffect(() => {
    if (!busy || error || cek) return;
    setSelingan(0);
    const t = setInterval(() => setSelingan((i) => (i + 1) % SELINGAN.length), 4500);
    return () => clearInterval(t);
  }, [busy, error, cek]);
  const stop = (e) => e.preventDefault();

  return (
    <>
      <div className="hero">
        {role === "dsn" ? (
          <>
            <h1>Tinjau buku proyek akhir mahasiswa, <b>sebelum</b> <u>disetujui</u></h1>
            <p>Unggah dokumen mahasiswa. Sistem menandai bagian yang belum sesuai pedoman penulisan, lengkap dengan halamannya.</p>
          </>
        ) : (
          <>
            <h1>Cek buku proyek akhir, <b>sebelum</b> dosen yang <u>mengecek</u></h1>
            <p>Unggah dokumen kamu. Sistem menandai bagian yang belum sesuai pedoman penulisan, lengkap dengan halaman dan acuannya.</p>
          </>
        )}
      </div>
      {role === "dsn" && !busy && (
        <div className="card" style={{ padding: 20, marginBottom: 20, display: "grid", gap: 12 }}>
          <label style={{ display: "grid", gap: 4 }}>
            <b>Nama mahasiswa pemilik dokumen</b>
            <input style={kotak} value={mhsNama} list="nama-mhs" onChange={(e) => setMhsNama(e.target.value)} autoComplete="off" />
            <datalist id="nama-mhs">{namaLama.map((n) => <option key={n} value={n} />)}</datalist>
            <small style={{ color: "var(--mute)" }}>Tulis nama yang sama tiap kali, supaya skor draf berikutnya masuk ke baris mahasiswa yang sama.</small>
          </label>
          {errForm && <small role="alert" style={{ color: "var(--a)" }}>{errForm}</small>}
        </div>
      )}
      <div style={{ position: "relative" }}>
        <div className="stamp" aria-hidden="true">PDF<small>atau DOCX</small></div>
        <div
          className={`drop ${over ? "over" : ""}`.trim()}
          onDragOver={(e) => { stop(e); setOver(true); }}
          onDragEnter={(e) => { stop(e); setOver(true); }}
          onDragLeave={(e) => { stop(e); setOver(false); }}
          onDrop={(e) => { stop(e); setOver(false); const f = e.dataTransfer.files[0]; if (f && !busy) start(f); }}
        >
          <i /><i /><i /><i />
          <Mascot pose={error ? "sleepy" : cek ? "confused" : busy ? POSES[step] : "idle"} size={cek ? 96 : 130} badge={error ? "!" : cek ? "?" : "↑"} tone="blue" />
          {!busy ? (
            <div style={{ display: "grid", gap: 10, justifyItems: "center" }}>
              <h3 style={{ fontSize: 22 }}>Unggah dokumen untuk diperiksa</h3>
              <p style={{ color: "var(--mute)" }}>Format yang didukung: .pdf dan .docx</p>
              <label className="btn" style={{ margin: "6px 0 0", cursor: "pointer" }} htmlFor="fi">Pilih file</label>
              <input className="sr" id="fi" type="file" accept=".pdf,.docx" onChange={(e) => { const f = e.target.files[0]; e.target.value = ""; if (f) start(f); }} />
              <p style={{ color: "var(--mute)", fontSize: 14 }}>Atau seret berkas ke sini</p>
            </div>
          ) : cek ? (
            <div style={{ display: "grid", gap: 12, width: "min(580px,100%)", textAlign: "left" }}>
              <h3 style={{ fontSize: 22, textAlign: "center" }}>Cek judul kamu dulu</h3>
              <p style={{ color: "var(--mute)", textAlign: "center" }}>
                Sistem menebak tiga unsur dari judulmu. Ini dipakai untuk menilai Bab 1: kalau sudah benar, langsung mulai periksa; kalau salah, ubah dulu.
              </p>
              {cek.catatan && <p style={{ color: "var(--a)", fontSize: 14 }}>{cek.catatan}</p>}
              {!cek.edit ? (
                <div style={{ display: "grid", gap: 10 }}>
                  {[["judul", "Judul"], ...UNSUR.map(([k, n]) => [k, n])].map(([k, n]) => (
                    <div key={k} style={{ ...kotak, background: "var(--bt)", border: 0 }}>
                      <b style={{ fontSize: 13, display: "block", color: "var(--mute)" }}>{n}</b>
                      {cek.isi[k]}
                      {cek.perlu.find((p) => p.startsWith(k)) && (
                        <small style={{ display: "block", color: "var(--a)" }}>Perlu dicek: {cek.perlu.find((p) => p.startsWith(k)).slice(k.length + 1)}.</small>
                      )}
                    </div>
                  ))}
                </div>
              ) : (
                <>
              <label style={{ display: "grid", gap: 4 }}>
                <b style={{ fontSize: 14 }}>Judul</b>
                <input style={kotak} value={cek.isi.judul} onChange={(e) => ubah("judul", e.target.value)} />
                {!cek.judulPasti && <small style={{ color: "var(--a)" }}>Judul ini mungkin belum tepat, cocokkan dengan sampul bukumu.</small>}
              </label>
              {UNSUR.map(([k, nama, petunjuk]) => (
                <label key={k} style={{ display: "grid", gap: 4 }}>
                  <b style={{ fontSize: 14 }}>{nama}</b>
                  <textarea style={{ ...kotak, resize: "vertical" }} rows={2} value={cek.isi[k]} onChange={(e) => ubah(k, e.target.value)} />
                  <small style={{ color: cek.perlu.some((p) => p.startsWith(k)) ? "var(--a)" : "var(--mute)" }}>
                    {cek.perlu.find((p) => p.startsWith(k)) ? `Perlu dicek: ${cek.perlu.find((p) => p.startsWith(k)).slice(k.length + 1)}. ` : ""}{petunjuk}
                  </small>
                </label>
              ))}
                </>
              )}
              <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 6 }}>
                <button className="btn ghost" onClick={reset} disabled={kirimKonfirmasi}>Batal</button>
                {!cek.edit && <button className="btn ghost" onClick={() => setCek((c) => ({ ...c, edit: true }))} disabled={kirimKonfirmasi}>Ada yang salah, ubah</button>}
                <button className="btn" onClick={lanjut} disabled={!lengkap || kirimKonfirmasi}>Sudah benar, mulai periksa</button>
              </div>
            </div>
          ) : error ? (
            <div style={{ display: "grid", gap: 12, justifyItems: "center", width: "100%" }}>
              <h3 style={{ fontSize: 22 }}>Pemeriksaan belum berhasil</h3>
              <p style={{ color: "var(--mute)", maxWidth: 480, textAlign: "center" }}>{error}</p>
              <button className="btn" onClick={reset}>Coba lagi</button>
            </div>
          ) : (
            <div style={{ display: "grid", gap: 12, justifyItems: "center", width: "100%" }}>
              <h3 style={{ fontSize: 22 }}>{file}</h3>
              <div className="prog">
                <div style={{ width: full ? "92%" : 0, transitionDuration: "90s" }} />
              </div>
              <p style={{ color: "var(--mute)", maxWidth: 480, textAlign: "center" }}>{teksTahap[step]}</p>
              <small key={selingan} style={{ color: "var(--b)", fontWeight: 600 }}>{urutan[selingan]}</small>
            </div>
          )}
        </div>
      </div>
      <div className="card tips"><b>Tips</b><span>Pastikan dokumen sudah memakai format dan pedoman yang berlaku agar hasil pengecekan optimal.</span></div>
    </>
  );
}
