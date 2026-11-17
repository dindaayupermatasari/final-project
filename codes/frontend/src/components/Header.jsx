import { Link, useLocation, useNavigate } from "react-router-dom";
import { useAuth } from "../context/AuthContext.jsx";

const NAV = {
  mhs: [["/beranda", "Beranda"], ["/riwayat", "Riwayat"], ["/panduan", "Panduan"]],
  dsn: [["/beranda", "Beranda"], ["/dashboard", "Dashboard"], ["/panduan", "Panduan"]],
};

// Salam menurut jam setempat.
const salam = () => {
  const j = new Date().getHours();
  return j >= 4 && j < 11 ? "Selamat pagi" : j < 15 ? "Selamat siang" : j < 18 ? "Selamat sore" : "Selamat malam";
};

export default function Header() {
  const { user, role, logout } = useAuth();
  // Nama dari profil; jika kosong, pakai bagian depan email.
  const nama = user?.nama?.trim() || user?.email?.split("@")[0] || "";
  const { pathname } = useLocation();
  const navigate = useNavigate();

  // Halaman hasil menyorot menu asalnya: Dashboard untuk dosen, Beranda untuk mahasiswa.
  const current = pathname === "/hasil" ? (role === "dsn" ? "/dashboard" : "/beranda") : pathname;

  return (
    <header id="hd">
      <div className="bar">
        <Link className="brand" to="/beranda" aria-label="PENScek">
          <span className="logo" role="img" aria-label="PENScek" />
        </Link>
        <nav id="nav" aria-label="Menu utama">
          {(NAV[role] || []).map(([to, text]) => (
            <Link key={to} to={to} aria-current={to === current ? "page" : undefined}>{text}</Link>
          ))}
        </nav>
        {nama && (
          <div className="hi">
            <div>
              <small>{salam()}</small>
              <b title={nama}>{nama}</b>
            </div>
          </div>
        )}
        <button className="out" type="button" aria-label="Keluar" title="Keluar" onClick={async () => { await logout(); navigate("/"); }}>
          <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
            <polyline points="16 17 21 12 16 7" />
            <line x1="21" y1="12" x2="9" y2="12" />
          </svg>
        </button>
      </div>
    </header>
  );
}
