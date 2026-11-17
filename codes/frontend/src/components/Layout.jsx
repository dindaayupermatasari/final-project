import { useEffect } from "react";
import { Navigate, Outlet, useLocation } from "react-router-dom";
import { useAuth } from "../context/AuthContext.jsx";
import Header from "./Header.jsx";

// Pembungkus halaman yang butuh login: header + konten. Belum masuk -> kembali ke landing.
export default function Layout() {
  const { role, siap } = useAuth();
  const { pathname } = useLocation();
  useEffect(() => { window.scrollTo(0, 0); }, [pathname]);

  if (!siap) return null; // sesi masih dibaca; jangan melempar pengguna ke landing dulu
  if (!role) return <Navigate to="/" replace />;
  return (
    <>
      <Header />
      <main><Outlet /></main>
    </>
  );
}
