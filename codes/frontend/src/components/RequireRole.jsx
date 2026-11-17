import { Navigate } from "react-router-dom";
import { useAuth } from "../context/AuthContext.jsx";

// Membatasi halaman untuk satu peran ("mhs" atau "dsn").
export default function RequireRole({ role, children }) {
  const { role: current } = useAuth();
  return current === role ? children : <Navigate to="/beranda" replace />;
}
