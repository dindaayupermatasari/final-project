import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "../context/AuthContext.jsx";
import useParallax from "../hooks/useParallax.js";
import Mascot from "../components/Mascot.jsx";
import Ring from "../components/Ring.jsx";
import { emailPens } from "../utils.js";

const CHECKS = ["Ukuran huruf", "Margin halaman", "Spasi baris", "Penomoran subbab", "Format tabel", "Keterangan gambar", "Judul bab", "Penomoran halaman"];

const go = (id) =>
  document.getElementById(id)?.scrollIntoView({
    behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth",
  });

export default function Landing() {
  const { user, galat, masuk, daftar } = useAuth();
  const navigate = useNavigate();
  useParallax();

  const [mode, setMode] = useState("masuk"); // "masuk" | "daftar"
  const [nama, setNama] = useState("");
  const [email, setEmail] = useState("");
  const [sandi, setSandi] = useState("");
  const [lihat, setLihat] = useState(false); // tampilkan kata sandi
  const [sibuk, setSibuk] = useState(false);
  const [pesan, setPesan] = useState("");

  // Sudah punya sesi: langsung ke beranda.
  useEffect(() => { if (user) navigate("/beranda", { replace: true }); }, [user, navigate]);

  const ganti = (e, m) => { e.preventDefault(); setMode(m); setPesan(""); setSandi(""); };

  const submit = async (e) => {
    e.preventDefault();
    setPesan("");
    if (!emailPens(email)) {
      setPesan("Pakai email PENS: nama@pens.ac.id untuk dosen atau nama@ds.student.pens.ac.id untuk mahasiswa.");
      return;
    }
    if (mode === "daftar" && !nama.trim()) { setPesan("Isi nama lengkap."); return; }
    if (mode === "daftar" && sandi.length < 8) { setPesan("Kata sandi minimal 8 karakter."); return; }
    if (!sandi) { setPesan("Isi kata sandi."); return; }

    setSibuk(true);
    const { data, error } = mode === "daftar" ? await daftar(email, sandi, nama) : await masuk(email, sandi);
    setSibuk(false);

    if (error) {
      const m = error.message || "";
      setPesan(
        /invalid login credentials/i.test(m) ? "Email atau kata sandi salah."
        : /already registered|already been registered/i.test(m) ? "Email ini sudah terdaftar. Silakan masuk."
        : /database error/i.test(m) ? "Hanya akun email PENS yang bisa mendaftar."
        : error.status === 429 ? "Terlalu banyak percobaan. Tunggu sebentar lalu coba lagi."
        : /password/i.test(m) ? "Kata sandi terlalu lemah. Pakai minimal 8 karakter."
        : mode === "daftar" ? "Pendaftaran belum berhasil. Coba lagi sebentar lagi."
        : "Belum bisa masuk. Coba lagi sebentar lagi."
      );
      return;
    }
    // Jika "Confirm email" masih aktif di Supabase, signUp tidak mengembalikan sesi.
    if (mode === "daftar" && !data?.session) {
      setPesan("Akun dibuat. Cek email untuk konfirmasi, lalu masuk.");
      return;
    }
    // Berhasil: sesi terbaca AuthContext, lalu efek di atas membawa ke /beranda.
  };

  return (
    <main>
      <section id="login">
        <header className="ln">
          <div className="bar">
            <a className="brand" href="/" aria-label="PENScek"><span className="logo" role="img" aria-label="PENScek" /></a>
            <nav aria-label="Bagian halaman">
              <button className="nl" onClick={() => go("fitur")}>Fitur</button>
              <button className="nl" onClick={() => go("cara")}>Cara kerja</button>
              <button className="nl" onClick={() => go("dosen")}>Untuk dosen</button>
            </nav>
            <button className="btn" style={{ padding: "9px 22px" }} onClick={() => go("masuk")}>Masuk</button>
          </div>
        </header>

        {/* HERO: poster kiri, login kanan */}
        <div className="w hero2">
          <div className="hl">
            <h1>Buku PA rapi, tanpa begadang</h1>
            <p>Unggah draf buku proyek akhir kamu, dan PENScek akan menandai bagian yang belum sesuai dengan pedoman penulisan PENS</p>
            <div className="poster" aria-hidden="true">
              <span className="cap">Cek dulu, revisi belakangan</span>
              <i className="dt" style={{ left: 12, top: 12 }} /><i className="dt" style={{ right: 12, top: 12 }} />
              <i className="dt" style={{ left: 12, bottom: 12 }} /><i className="dt" style={{ right: 12, bottom: 12 }} />
              <svg className="ph" viewBox="0 0 560 300" preserveAspectRatio="xMidYMid slice">
                <defs>
                  <pattern id="dots2" width="9" height="9" patternUnits="userSpaceOnUse"><circle cx="4.5" cy="4.5" r="2.7" fill="#fff" /></pattern>
                  <linearGradient id="fd2" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stopColor="#fff" /><stop offset="1" stopColor="#000" /></linearGradient>
                  <mask id="mk2"><rect width="560" height="300" fill="url(#fd2)" /></mask>
                </defs>
                <rect className="fb" width="560" height="300" />
                <g className="fn"><circle cx="440" cy="62" r="30" /><circle cx="474" cy="76" r="24" /><circle cx="408" cy="78" r="22" /><rect x="390" y="76" width="106" height="26" rx="13" /></g>
                <path className="fn" d="M0 300V196h110v32h110v32h110v28h230v12z" />
                <path d="M20 170L250 96M52 198L262 126" stroke="#FBFBF4" strokeWidth="6" fill="none" />
                <rect width="560" height="300" fill="url(#dots2)" mask="url(#mk2)" opacity=".4" />
              </svg>
              <div className="big" data-s=".07" data-rot="-8"><span>C</span><span>E</span><span>K</span></div>
              <div className="sheet" data-s="-.1" data-rot="5" style={{ right: "-3%", top: "14%", width: "34%", transform: "rotate(5deg)" }}>
                <b>BAB 1 · hlm. 3</b>
                <div className="l" /><div className="l hot" /><div className="l" style={{ width: "70%" }} /><div className="l" />
              </div>
              <div className="badge" data-s=".16" style={{ left: "34%", top: "20%" }}>
                <Ring score={85} color="var(--b)" />Kesesuaian<br />keseluruhan
              </div>
              <Mascot pose="search" size={118} data-s="-.09" style={{ position: "absolute", right: "11%", bottom: "10%" }} />
            </div>
          </div>

          <aside className="lgp" id="masuk" aria-label="Masuk">
            <Mascot pose="happy" size={176} data-s=".08" aria-hidden="true" />
            <div className="card lgc">
              <h2 style={{ fontSize: 32 }}>{mode === "masuk" ? "Masuk ke PENScek" : "Buat akun PENScek"}</h2>
              <p style={{ color: "var(--mute)", fontSize: 15, marginTop: 6 }}>
                {mode === "masuk"
                  ? "Pertama kali ke sini? Yuk, pilih Daftar dulu untuk membuat akun dengan email PENS!"
                  : "Daftar dengan email PENS dan kata sandi buatanmu. Tidak perlu kode verifikasi."}
              </p>
              <div className="seg" role="group" aria-label="Pilih masuk atau daftar">
                <button type="button" aria-pressed={mode === "masuk"} onClick={(e) => ganti(e, "masuk")}>Masuk</button>
                <button type="button" aria-pressed={mode === "daftar"} onClick={(e) => ganti(e, "daftar")}>Daftar</button>
              </div>
              <form onSubmit={submit} noValidate>
                {mode === "daftar" && (
                  <>
                    <label htmlFor="nama" style={{ marginTop: 0 }}>Nama lengkap</label>
                    <input id="nama" type="text" autoComplete="name" value={nama} onChange={(e) => setNama(e.target.value)} />
                  </>
                )}
                <label htmlFor="nid" style={mode === "daftar" ? undefined : { marginTop: 0 }}>Email PENS</label>
                <input id="nid" type="email" autoComplete="username" placeholder="nama@pens.ac.id" value={email} onChange={(e) => setEmail(e.target.value)} />
                <label htmlFor="sandi">Kata sandi</label>
                <div className="pw">
                  <input id="sandi" type={lihat ? "text" : "password"} autoComplete={mode === "daftar" ? "new-password" : "current-password"} value={sandi} onChange={(e) => setSandi(e.target.value)} />
                  <button type="button" className="pwb" onClick={() => setLihat((v) => !v)} aria-label={lihat ? "Sembunyikan kata sandi" : "Tampilkan kata sandi"} aria-pressed={lihat}>
                    {lihat ? (
                      <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <path d="M17.94 17.94A10.9 10.9 0 0 1 12 20C5 20 1 12 1 12a18.5 18.5 0 0 1 5.06-5.94" />
                        <path d="M9.9 4.24A10.9 10.9 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19" />
                        <path d="M14.12 14.12a3 3 0 1 1-4.24-4.24" />
                        <path d="M1 1l22 22" />
                      </svg>
                    ) : (
                      <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                        <circle cx="12" cy="12" r="3" />
                      </svg>
                    )}
                  </button>
                </div>
                {(pesan || galat) && <p role="alert" style={{ color: "var(--a)", fontSize: 14, marginTop: 10 }}>{pesan || galat}</p>}
                <button className="btn hot" disabled={sibuk} style={{ width: "100%", justifyContent: "center", marginTop: 12 }}>
                  {sibuk ? "Memeriksa…" : mode === "masuk" ? "Masuk" : "Daftar"}
                </button>
              </form>
              <p style={{ textAlign: "center", fontSize: 14, marginTop: 14, color: "var(--mute)" }}>
                Butuh bantuan akses? <a href="#" onClick={(e) => e.preventDefault()}>Hubungi kami</a>
              </p>
            </div>
          </aside>
        </div>

        {/* Yang dicek */}
        <div className="chk"><div className="w">
          <h2>Yang dicek PENScek</h2>
          <div className="chips">{CHECKS.map((c) => <span key={c}>{c}</span>)}</div>
        </div></div>

        {/* Fitur */}
        <div className="sec" id="fitur"><div className="w">
          <h2 style={{ maxWidth: "24ch" }}>Temuan yang biasanya baru kamu tahu saat revisi</h2>
          <p className="lead">Semua dicek terhadap panduan buku proyek akhir PENS, bab demi bab.</p>
          <div className="bento">
            <div className="bg1" data-s="-.03">
              <h3>Setiap temuan ada halaman dan acuannya</h3>
              <p>Klik satu temuan, kalimat yang bermasalah langsung disorot di pratinjau dokumen.</p>
              <div className="fl">
                <div><span className="n">1</span>Ukuran huruf<small>hlm. 3</small></div>
                <div><span className="n">2</span>Margin kiri<small>hlm. 4</small></div>
                <div><span className="n">3</span>Penomoran subbab<small>hlm. 4</small></div>
              </div>
              <div className="dv" aria-hidden="true">
                <i style={{ width: "94%" }} /><i style={{ width: "88%" }} />
                <mark>Nomor persamaan seharusnya dicetak tebal.</mark>
                <i style={{ width: "72%" }} />
                <span className="cm"><b>Penomoran persamaan</b>Beri 1 baris kosong di atas dan di bawah persamaan.</span>
                <Mascot pose="search" size={84} style={{ position: "absolute", right: -6, bottom: -14 }} />
              </div>
            </div>
            <div className="bx">
              <h3>Skor per bab</h3>
              <p>Lihat bab mana yang paling perlu dibenahi lebih dulu.</p>
              <div className="mini">
                <i style={{ width: "85%" }} /><i style={{ width: "90%" }} />
                <i style={{ width: "70%", background: "var(--a)" }} /><i style={{ width: "80%" }} />
              </div>
            </div>
            <div className="bx p">
              <h3>Riwayat draf</h3>
              <p>Bandingkan skor dari draf ke draf dan lihat progresmu naik.</p>
              <svg viewBox="0 0 240 70" width="100%" style={{ marginTop: 16 }} aria-hidden="true">
                <path d="M8 58 L120 36 L232 10" stroke="#2F6BE0" strokeWidth="4" fill="none" strokeLinecap="round" />
                <circle cx="8" cy="58" r="7" fill="#2F6BE0" /><circle cx="120" cy="36" r="7" fill="#2F6BE0" /><circle cx="232" cy="10" r="9" fill="#E3257F" />
              </svg>
            </div>
          </div>
        </div></div>

        {/* Cara kerja */}
        <div className="sec" id="cara" style={{ background: "var(--n2)", borderBlock: "1.5px solid var(--ink)" }}><div className="w">
          <h2>Tiga langkah sampai draf siap</h2>
          <div className="steps">
            <div className="st" data-s="-.04"><Mascot pose="surprised" size={92} badge="↑" tone="blue" style={{ position: "absolute", right: 24, top: 22 }} /><div className="num">1</div><h3>Unggah</h3><p>Pilih draf buku PA dalam format PDF atau DOCX, atau seret langsung ke halaman.</p></div>
            <div className="st" data-s="-.07"><Mascot pose="search" size={92} badge="!" style={{ position: "absolute", right: 24, top: 22 }} /><div className="num">2</div><h3>Periksa</h3><p>Sistem membaca tiap bab dan mencocokkannya dengan pedoman penulisan.</p></div>
            <div className="st" data-s="-.1"><Mascot pose="happy" size={92} badge="✓" tone="blue" style={{ position: "absolute", right: 24, top: 22 }} /><div className="num">3</div><h3>Perbaiki</h3><p>Buka daftar temuan, benahi di dokumen, lalu unggah lagi untuk melihat skor baru.</p></div>
          </div>
        </div></div>

        {/* Untuk dosen */}
        <div className="sec" id="dosen"><div className="w dos">
          <div>
            <h2>Pantau mahasiswa bimbingan di satu tabel</h2>
            <p className="lead">Lihat skor kesesuaian dan status layak atau perlu revisi, lalu buka detail temuan tiap mahasiswa.</p>
          </div>
          <div className="dpanel">
            <Mascot pose="wink" size={120} style={{ position: "absolute", right: 30, top: -6, zIndex: 1 }} />
            <div className="mc" data-s="-.05">
              {[["AN", "Aditya Nugroho", 100], ["BP", "Bima Pratama", 75], ["DA", "Dinda Ayu", 85]].map(([av, name, s]) => (
                <div className="rw" key={av}>
                  <span className="av">{av}</span><b>{name}</b>
                  <span className={`pill ${s < 80 ? "rev" : ""}`.trim()}>{s >= 80 ? "Layak" : "Perlu revisi"}</span>
                  <b style={{ flex: "none" }}>{s}%</b>
                </div>
              ))}
            </div>
          </div>
        </div></div>

        {/* Penutup */}
        <div className="mas"><div className="w" style={{ gridTemplateColumns: "1fr", textAlign: "center", justifyItems: "center" }}>
          <h2 style={{ maxWidth: "14ch" }}>Siap cek draf kamu?</h2>
          <p>Masuk dengan email PENS dan unggah draf pertamamu.</p>
          <button className="btn hot" style={{ marginTop: 28 }} onClick={() => go("masuk")}>Masuk sekarang</button>
        </div></div>
        <footer className="ft">
          <Mascot pose="sleepy" size={70} style={{ display: "block", margin: "0 auto 8px" }} />
          PENScek · Pemeriksa kesesuaian buku proyek akhir dengan pedoman penulisan PENS
        </footer>
      </section>
    </main>
  );
}
