-- Opsional. Jalankan di Supabase > SQL Editor SETELAH backend baru (api/app.py dan api/db.py) terpasang,
-- karena backend lama masih membaca kolom nim_nip.
-- NIM tidak dipakai lagi: dosen hanya mengisi nama mahasiswa, jadi indeks per mahasiswa pindah ke nama.

drop index if exists public.checks_mhs_idx;a
alter table public.checks   drop column if exists mhs_nim;
alter table public.profiles drop column if exists nim_nip;   -- hak update kolom ini ikut hilang

create index if not exists checks_mhs_idx on public.checks (uploader_id, mhs_nama);
