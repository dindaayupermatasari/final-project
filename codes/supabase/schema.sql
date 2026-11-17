-- PENScek: jalankan sekali di Supabase > SQL Editor.

-- 1. Profil pengguna. Dibuat otomatis saat akun baru terdaftar.
create table if not exists public.profiles (
  id         uuid primary key references auth.users (id) on delete cascade,
  role       text not null default 'mhs' check (role in ('mhs', 'dsn')),
  nama       text not null default '',
  nim_nip    text,
  created_at timestamptz not null default now()
);

-- Role TIDAK diambil dari data pendaftaran, supaya orang tidak bisa mendaftar sebagai dosen.
-- Semua akun baru jadi 'mhs'. Jadikan dosen secara manual:
--   update public.profiles set role = 'dsn' where id = '<uuid dari Authentication > Users>';
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  insert into public.profiles (id, nama, nim_nip)
  values (new.id,
          coalesce(new.raw_user_meta_data ->> 'nama', ''),
          new.raw_user_meta_data ->> 'nim_nip');
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- 2. Pemeriksaan. Satu baris per unggahan.
create table if not exists public.checks (
  id               text primary key,                -- 12 karakter hex dari backend
  uploader_id      uuid not null references auth.users (id) on delete cascade,
  mhs_nama         text not null default '',
  mhs_nim          text not null default '',
  file_name        text not null,
  judul            text not null default '',
  status           text not null default 'queued'
                   check (status in ('queued', 'running', 'menunggu', 'done', 'error')),
  error            text,
  skor             numeric(5, 1),
  decomp           jsonb,
  result           jsonb,                           -- hasil lengkap (kecil, sekitar 5-20 KB)
  sorotan          jsonb,                           -- posisi sorotan di PDF
  file_path        text,                            -- null = dokumen sudah dihapus
  file_expires_at  timestamptz,
  created_at       timestamptz not null default now(),
  finished_at      timestamptz
);

create index if not exists checks_uploader_idx on public.checks (uploader_id, created_at desc);
create index if not exists checks_mhs_idx      on public.checks (uploader_id, mhs_nim);
create index if not exists checks_expiry_idx   on public.checks (file_expires_at) where file_path is not null;

-- 3. Row Level Security: tiap orang hanya melihat miliknya.
alter table public.profiles enable row level security;
alter table public.checks   enable row level security;

drop policy if exists "profil sendiri (baca)"  on public.profiles;
drop policy if exists "profil sendiri (ubah)"  on public.profiles;
drop policy if exists "checks sendiri (baca)"  on public.checks;

create policy "profil sendiri (baca)" on public.profiles
  for select to authenticated using (id = (select auth.uid()));
create policy "profil sendiri (ubah)" on public.profiles
  for update to authenticated using (id = (select auth.uid())) with check (id = (select auth.uid()));
create policy "checks sendiri (baca)" on public.checks
  for select to authenticated using (uploader_id = (select auth.uid()));

-- Pengguna hanya boleh mengubah nama dan NIM/NIP, bukan role.
revoke update on public.profiles from anon, authenticated;
grant  update (nama, nim_nip) on public.profiles to authenticated;

-- Tulis ke checks hanya lewat backend (service key). Klien tidak boleh insert/update/delete.
revoke insert, update, delete on public.checks from anon, authenticated;

-- 4. Bucket dokumen: privat, tanpa policy, jadi hanya backend yang bisa menyentuhnya.
insert into storage.buckets (id, name, public)
values ('dokumen', 'dokumen', false)
on conflict (id) do nothing;
