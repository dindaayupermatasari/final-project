# PENScek (React)

Antarmuka untuk memeriksa kesesuaian buku proyek akhir dengan pedoman penulisan PENS.
React 18 + React Router 6 + Vite. Gaya memakai satu file CSS biasa.

## Menjalankan
Butuh Node.js 18 atau lebih baru.

    npm install
    cp .env.example .env   # isi VITE_SUPABASE_URL dan VITE_SUPABASE_ANON_KEY
    npm run dev        # http://localhost:5173
    npm run build      # hasil di dist/
    npm run preview    # coba hasil build

Font (Archivo, Public Sans) dimuat dari Google Fonts, jadi butuh internet.

## Struktur
    index.html
    src/
      main.jsx                 titik masuk: Router + AuthProvider
      App.jsx                  daftar rute
      pages/                   satu file per halaman
        Landing.jsx            "/"           landing + login
        Beranda.jsx            "/beranda"    unggah dokumen
        Hasil.jsx              "/hasil?id=<id>"  hasil per BAB (dosen: mode tinjau)
        Riwayat.jsx            "/riwayat"    khusus mahasiswa, dari GET /api/checks
        Dashboard.jsx          "/dashboard"  khusus dosen, skor per mahasiswa (dikelompokkan per nama) dari unggahannya sendiri
        Panduan.jsx            "/panduan"
      components/
        Layout.jsx             header + konten, memeriksa login
        Header.jsx             menu sesuai peran
        RequireRole.jsx        membatasi halaman per peran
        Mascot.jsx             <Mascot pose="happy" size={96} badge="✓" tone="blue" />
        Ring.jsx               cincin skor
      context/AuthContext.jsx  sesi Supabase + profil dari /api/me (peran dari backend)
      supabase.js              klien Supabase (kunci anon)
      api.js                   klien backend, semua permintaan membawa token login
      hooks/useParallax.js     efek parallax landing
      utils.js
      styles/style.css         semua gaya + token palet 60-30-10
      assets/img/              logo dan 5 pose maskot

## Palet 60-30-10 (src/styles/style.css, bagian :root)
| Porsi | Token | Warna |
|---|---|---|
| 60% netral | --n, --n2 | #F1F2E6, #FBFBF4 |
| 30% pendukung | --b, --b2, --bt | #2F6BE0, #2557C2, #DDE8FC |
| 10% aksen | --a, --at | #E3257F, #FAD7E8 |

## Login
Masuk dengan kode 6 digit yang dikirim ke email PENS (`@pens.ac.id` untuk dosen, `@<prodi>.student.pens.ac.id`
untuk mahasiswa). Kode hanya diminta saat pertama masuk di sebuah perangkat; sesi disimpan dan diperbarui otomatis.
Peran (`mhs` atau `dsn`) ditentukan database dari domain email, lalu dibaca frontend lewat `GET /api/me`.
Template email "Magic Link" di Supabase harus memakai `{{ .Token }}` supaya yang terkirim berupa kode.

## Backend
`api.js` memanggil FastAPI (`api/app.py`) lewat proxy `/api` saat development. Di produksi isi `VITE_API_URL`
dan `PENSCEK_ORIGINS` di backend. Dokumen disimpan 30 hari; sesudahnya hasil tetap ada, pratinjau dan unduhan tidak.
`data/data.js` dan `laporan.js` tidak dipakai lagi dan boleh dihapus.

## Deploy
Rute memakai BrowserRouter, jadi server harus mengarahkan semua path ke `index.html`
(di Netlify tambahkan `public/_redirects` berisi `/* /index.html 200`; di Vercel gunakan rewrite ke `/index.html`).
