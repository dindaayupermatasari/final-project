import { createContext, useContext, useEffect, useState } from "react";
import { supabase } from "../supabase.js";
import { saya } from "../api.js";

const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  // undefined = sesi belum dibaca; null = tidak ada sesi.
  const [session, setSession] = useState(undefined);
  const [user, setUser] = useState(null); // { id, email, role, nama } dari /api/me
  const [siap, setSiap] = useState(false); // false selama sesi dan profil belum selesai dibaca
  const [galat, setGalat] = useState("");

  // Callback ini sengaja hanya menyimpan sesi. Memanggil fungsi auth Supabase di dalamnya bisa membuatnya macet.
  useEffect(() => {
    const { data } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => data.subscription.unsubscribe();
  }, []);

  const uid = session === undefined ? undefined : session?.user?.id ?? null;
  useEffect(() => {
    if (uid === undefined) return;
    if (uid === null) { setUser(null); setGalat(""); setSiap(true); return; }
    let batal = false;
    setGalat("");
    // Peran dan nama diambil dari backend (tabel profiles), bukan dari pilihan di layar.
    saya()
      .then((u) => { if (!batal) setUser(u); })
      .catch((e) => { if (!batal) { setUser(null); setGalat(e.message); } })
      .finally(() => { if (!batal) setSiap(true); });
    return () => { batal = true; };
  }, [uid]);

  const masuk = (email, password) =>
    supabase.auth.signInWithPassword({ email: email.trim().toLowerCase(), password });
  const daftar = (email, password, nama) =>
    supabase.auth.signUp({
      email: email.trim().toLowerCase(),
      password,
      options: { data: { full_name: nama.trim() } },
    });
  const logout = () => supabase.auth.signOut();

  return (
    <AuthContext.Provider value={{ user, role: user?.role ?? null, siap, galat, masuk, daftar, logout }}>
      {children}
    </AuthContext.Provider>
  );
}

export const useAuth = () => useContext(AuthContext);
