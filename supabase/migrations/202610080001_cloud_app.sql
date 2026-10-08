-- Shared packing lists, short-lived uploads, and reviewed example queue.
create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  role text not null default 'sales' check (role in ('sales', 'admin')),
  key_mode text not null default 'shared' check (key_mode in ('shared', 'personal')),
  created_at timestamptz not null default now()
);

create or replace function public.create_profile_for_user()
returns trigger language plpgsql security definer set search_path = '' as $$
begin
  insert into public.profiles (id, email)
  values (new.id, coalesce(new.email, ''))
  on conflict (id) do update set email = excluded.email;
  return new;
end;
$$;
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created after insert on auth.users
for each row execute function public.create_profile_for_user();

create table if not exists public.packing_lists (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references public.profiles(id),
  data jsonb not null,
  source_hashes text[] not null default '{}',
  status text not null default 'final' check (status in ('final', 'training_candidate', 'training_approved', 'training_rejected')),
  source_paths text[] not null default '{}',
  origin text not null default 'recognition' check (origin in ('recognition', 'manual_import')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  approved_by uuid references public.profiles(id),
  approved_at timestamptz
);
create index if not exists packing_lists_owner_created_idx on public.packing_lists(owner_id, created_at desc);
create index if not exists packing_lists_status_idx on public.packing_lists(status, created_at desc);
create index if not exists packing_lists_source_hashes_idx on public.packing_lists using gin(source_hashes);

create table if not exists public.temp_uploads (
  path text primary key,
  owner_id uuid not null references public.profiles(id),
  created_at timestamptz not null default now()
);
create index if not exists temp_uploads_created_idx on public.temp_uploads(created_at);

-- Credentials are only read by the Edge Function with the server-side secret key.
create table if not exists public.gemini_credentials (
  owner_key text primary key,
  encrypted_key text not null,
  updated_by uuid not null references public.profiles(id),
  updated_at timestamptz not null default now()
);

alter table public.profiles enable row level security;
alter table public.packing_lists enable row level security;
alter table public.temp_uploads enable row level security;
alter table public.gemini_credentials enable row level security;

revoke all on public.gemini_credentials from anon, authenticated;
revoke all on public.temp_uploads from anon;
grant select, insert, update, delete on public.profiles, public.packing_lists, public.temp_uploads, public.gemini_credentials to service_role;
grant select, update on public.profiles to authenticated;
grant select on public.packing_lists to authenticated;
grant select, insert, delete on public.temp_uploads to authenticated;

create or replace function public.is_admin()
returns boolean language sql stable security definer set search_path = '' as $$
  select exists(select 1 from public.profiles where id = (select auth.uid()) and role = 'admin')
$$;

create policy "read own profile or admin" on public.profiles for select to authenticated
using (id = (select auth.uid()) or public.is_admin());
create policy "update own key mode" on public.profiles for update to authenticated
using (id = (select auth.uid())) with check (id = (select auth.uid()));
-- Column privilege keeps role/email immutable from browser clients.
revoke update on public.profiles from authenticated;
grant update (key_mode) on public.profiles to authenticated;

create policy "read own lists or admin" on public.packing_lists for select to authenticated
using (owner_id = (select auth.uid()) or public.is_admin());
create policy "read own uploads" on public.temp_uploads for select to authenticated
using (owner_id = (select auth.uid()) or public.is_admin());
create policy "insert own uploads" on public.temp_uploads for insert to authenticated
with check (owner_id = (select auth.uid()) and split_part(path, '/', 1) = (select auth.uid())::text);
create policy "delete own uploads" on public.temp_uploads for delete to authenticated
using (owner_id = (select auth.uid()) or public.is_admin());

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('packing-sources', 'packing-sources', false, 18874368, array['image/jpeg', 'image/png', 'image/webp'])
on conflict (id) do nothing;

create policy "upload own source pages" on storage.objects for insert to authenticated
with check (bucket_id = 'packing-sources' and (storage.foldername(name))[1] = (select auth.uid())::text);
create policy "read own source pages" on storage.objects for select to authenticated
using (bucket_id = 'packing-sources' and ((storage.foldername(name))[1] = (select auth.uid())::text or public.is_admin()));
create policy "delete own source pages" on storage.objects for delete to authenticated
using (bucket_id = 'packing-sources' and ((storage.foldername(name))[1] = (select auth.uid())::text or public.is_admin()));
