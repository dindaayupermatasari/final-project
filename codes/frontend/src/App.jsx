import { Routes, Route, Navigate } from "react-router-dom";
import Layout from "./components/Layout.jsx";
import RequireRole from "./components/RequireRole.jsx";
import Landing from "./pages/Landing.jsx";
import Beranda from "./pages/Beranda.jsx";
import Hasil from "./pages/Hasil.jsx";
import Riwayat from "./pages/Riwayat.jsx";
import Panduan from "./pages/Panduan.jsx";
import Dashboard from "./pages/Dashboard.jsx";

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Landing />} />
      <Route element={<Layout />}>
        <Route path="/beranda" element={<Beranda />} />
        <Route path="/hasil" element={<Hasil />} />
        <Route path="/panduan" element={<Panduan />} />
        <Route path="/riwayat" element={<RequireRole role="mhs"><Riwayat /></RequireRole>} />
        <Route path="/dashboard" element={<RequireRole role="dsn"><Dashboard /></RequireRole>} />
      </Route>
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}
