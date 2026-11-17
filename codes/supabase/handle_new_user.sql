-- Jalankan di Supabase > SQL Editor. Menggantikan handle_new_user di schema.sql.
-- Hanya email PENS yang bisa mendaftar; peran ditentukan dari domain email.
-- Aman karena login memakai kode OTP: akun baru hanya terbentuk setelah pemilik email terbukti menerima kodenya.
-- Jangan aktifkan pendaftaran email + password tanpa konfirmasi email, atau peran dosen bisa dipalsukan.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_email text := lower(coalesce(new.email, ''));
  v_role  text;
begin
  if v_email like '%@pens.ac.id' then
    v_role := 'dsn';
  elsif v_email like '%.student.pens.ac.id' then
    v_role := 'mhs';
  else
    raise exception 'Hanya akun PENS yang dapat masuk';
  end if;

  insert into public.profiles (id, role, nama)
  values (new.id, v_role,
          coalesce(new.raw_user_meta_data ->> 'full_name',
                   new.raw_user_meta_data ->> 'name', ''));
  return new;
end;
$$;
