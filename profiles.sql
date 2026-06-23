-- =============================================================================
--  GESTÃO DE SALAS DE REUNIÃO — Tabela de perfis de usuários
--  -----------------------------------------------------------------------------
--  Como aplicar:
--    1. Supabase → SQL Editor → New query
--    2. Cole TODO este arquivo e clique em RUN.
-- =============================================================================

-- =============================================================================
--  TABELA: profiles
--  Criada automaticamente quando um usuário se cadastra via Supabase Auth.
-- =============================================================================
create table if not exists public.profiles (
  id          uuid primary key references auth.users(id) on delete cascade,
  name        text,
  email       text not null,
  created_at  timestamptz not null default now(),
  last_login  timestamptz
);

-- =============================================================================
--  ROW LEVEL SECURITY
-- =============================================================================
alter table public.profiles enable row level security;

-- Usuário lê apenas o próprio perfil
drop policy if exists "profiles_self_read" on public.profiles;
create policy "profiles_self_read"
  on public.profiles for select
  using (auth.uid() = id);

-- Usuário atualiza apenas o próprio perfil
drop policy if exists "profiles_self_update" on public.profiles;
create policy "profiles_self_update"
  on public.profiles for update
  using (auth.uid() = id);

-- O sistema insere automaticamente via trigger (sem restrição de uid)
drop policy if exists "profiles_insert_trigger" on public.profiles;
create policy "profiles_insert_trigger"
  on public.profiles for insert
  with check (true);

-- =============================================================================
--  TRIGGER — cria o perfil automaticamente no cadastro
-- =============================================================================
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  insert into public.profiles (id, name, email)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'name', split_part(new.email, '@', 1)),
    new.email
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- =============================================================================
--  TRIGGER — atualiza last_login a cada acesso
-- =============================================================================
create or replace function public.handle_user_login()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  update public.profiles
  set last_login = now()
  where id = new.id;
  return new;
end;
$$;

drop trigger if exists on_auth_user_login on auth.users;
create trigger on_auth_user_login
  after update of last_sign_in_at on auth.users
  for each row execute function public.handle_user_login();

-- =============================================================================
--  FIM
-- =============================================================================
