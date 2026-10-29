-- Jalankan di Supabase: SQL Editor -> New query -> Run.
create table if not exists public.masa_manfaat (
  kode7 text primary key, uraian7 text, kode6 text, uraian6 text,
  masa_manfaat_bulan int not null check (masa_manfaat_bulan >= 0)
);
alter table public.masa_manfaat enable row level security;
create policy "master dapat dibaca semua" on public.masa_manfaat for select using (true);
-- Isi data: Table Editor -> masa_manfaat -> Import data from CSV -> supabase/masa_manfaat_seed.csv

create table if not exists public.penyusutan_runs (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  nama_file text, pengguna_barang text, tahun int not null, params jsonb not null default '{}',
  ringkasan jsonb not null default '{}', created_at timestamptz not null default now()
);
create table if not exists public.penyusutan_items (
  id bigserial primary key,
  run_id uuid not null references public.penyusutan_runs(id) on delete cascade,
  urut int not null, data jsonb not null
);
create index if not exists penyusutan_items_run on public.penyusutan_items(run_id, urut);
alter table public.penyusutan_runs enable row level security;
alter table public.penyusutan_items enable row level security;
create policy "run milik sendiri" on public.penyusutan_runs for all using (user_id = auth.uid()) with check (user_id = auth.uid());
create policy "item milik run sendiri" on public.penyusutan_items for all
  using (exists (select 1 from public.penyusutan_runs r where r.id = run_id and r.user_id = auth.uid()))
  with check (exists (select 1 from public.penyusutan_runs r where r.id = run_id and r.user_id = auth.uid()));
