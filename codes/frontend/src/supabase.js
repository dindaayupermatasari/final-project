import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !key) {
  throw new Error("VITE_SUPABASE_URL dan VITE_SUPABASE_ANON_KEY belum diisi. Salin .env.example menjadi .env lalu isi nilainya.");
}

// Sesi disimpan di localStorage dan token diperbarui otomatis, jadi pengguna tidak perlu masuk ulang setiap membuka situs.
export const supabase = createClient(url, key, {
  auth: { persistSession: true, autoRefreshToken: true },
});
