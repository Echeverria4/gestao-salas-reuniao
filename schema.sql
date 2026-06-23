-- =============================================================================
--  GESTÃO DE SALAS DE REUNIÃO — Esquema do banco (Supabase / PostgreSQL)
--  -----------------------------------------------------------------------------
--  Como aplicar:
--    1. Supabase  ->  SQL Editor  ->  New query
--    2. Cole TODO este arquivo e clique em RUN.
--  Pode rodar novamente sem medo: usa IF NOT EXISTS / DROP onde necessário.
-- =============================================================================

-- Extensões necessárias -------------------------------------------------------
create extension if not exists "pgcrypto";   -- gen_random_uuid()
create extension if not exists "btree_gist";  -- exclusão de horários sobrepostos

-- =============================================================================
--  TABELA: floors (andares)
-- =============================================================================
create table if not exists public.floors (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,                    -- ex.: "Térreo", "3º Andar"
  number      integer,                          -- ordem de exibição
  description text,
  created_at  timestamptz not null default now()
);

-- =============================================================================
--  TABELA: rooms (salas de reunião)
-- =============================================================================
create table if not exists public.rooms (
  id          uuid primary key default gen_random_uuid(),
  floor_id    uuid not null references public.floors(id) on delete cascade,
  name        text not null,                    -- ex.: "Sala Atlântico"
  capacity    integer not null default 4,       -- nº de pessoas
  location    text,                             -- referência física
  equipment   text[] default '{}',              -- {'TV','Projetor','Webcam'}
  color       text default '#2563eb',           -- cor para o calendário
  active      boolean not null default true,
  created_at  timestamptz not null default now()
);

create index if not exists rooms_floor_id_idx on public.rooms(floor_id);

-- =============================================================================
--  TABELA: bookings (agendamentos)
-- =============================================================================
create table if not exists public.bookings (
  id              uuid primary key default gen_random_uuid(),
  room_id         uuid not null references public.rooms(id) on delete cascade,
  title           text not null,                -- assunto da reunião
  organizer_name  text not null,
  organizer_email text,
  department      text,
  attendees       integer default 1,
  start_time      timestamptz not null,
  end_time        timestamptz not null,
  status          text not null default 'confirmed',  -- confirmed | cancelled
  notes           text,
  created_at      timestamptz not null default now(),

  constraint bookings_time_check check (end_time > start_time),
  constraint bookings_status_check check (status in ('confirmed','cancelled'))
);

create index if not exists bookings_room_id_idx    on public.bookings(room_id);
create index if not exists bookings_start_time_idx on public.bookings(start_time);

-- Impede DOIS agendamentos confirmados na MESMA sala com horários sobrepostos.
-- (a mágica anti-conflito acontece aqui no banco, não só no front-end)
alter table public.bookings drop constraint if exists bookings_no_overlap;
alter table public.bookings
  add constraint bookings_no_overlap
  exclude using gist (
    room_id with =,
    tstzrange(start_time, end_time) with &&
  ) where (status = 'confirmed');

-- =============================================================================
--  ROW LEVEL SECURITY
--  -----------------------------------------------------------------------------
--  Para começar rápido, liberamos leitura/escrita pública (chave anon).
--  >>> Em produção, troque por políticas baseadas em auth.uid() / papéis. <<<
-- =============================================================================
alter table public.floors   enable row level security;
alter table public.rooms    enable row level security;
alter table public.bookings enable row level security;

-- floors
drop policy if exists "floors_public_read"  on public.floors;
drop policy if exists "floors_public_write" on public.floors;
create policy "floors_public_read"  on public.floors for select using (true);
create policy "floors_public_write" on public.floors for all using (true) with check (true);

-- rooms
drop policy if exists "rooms_public_read"  on public.rooms;
drop policy if exists "rooms_public_write" on public.rooms;
create policy "rooms_public_read"  on public.rooms for select using (true);
create policy "rooms_public_write" on public.rooms for all using (true) with check (true);

-- bookings
drop policy if exists "bookings_public_read"  on public.bookings;
drop policy if exists "bookings_public_write" on public.bookings;
create policy "bookings_public_read"  on public.bookings for select using (true);
create policy "bookings_public_write" on public.bookings for all using (true) with check (true);

-- =============================================================================
--  DADOS DE EXEMPLO (opcional — comente se não quiser)
-- =============================================================================
insert into public.floors (name, number, description)
select * from (values
  ('Térreo',   0, 'Recepção e salas de apoio'),
  ('1º Andar', 1, 'Área comercial'),
  ('3º Andar', 3, 'Diretoria e reuniões executivas')
) as t(name, number, description)
where not exists (select 1 from public.floors);

-- Salas de exemplo vinculadas aos andares acima
insert into public.rooms (floor_id, name, capacity, location, equipment, color)
select f.id, r.name, r.capacity, r.location, r.equipment, r.color
from public.floors f
join (values
  ('Térreo',   'Sala Recepção',  4,  'Próx. à entrada', array['TV'],                 '#0ea5e9'),
  ('Térreo',   'Sala Apoio',     6,  'Corredor B',      array['TV','Quadro'],        '#14b8a6'),
  ('1º Andar', 'Sala Atlântico', 8,  'Ala leste',       array['TV','Webcam'],        '#2563eb'),
  ('1º Andar', 'Sala Pacífico',  12, 'Ala oeste',       array['Projetor','Webcam'],  '#7c3aed'),
  ('3º Andar', 'Sala Diretoria', 16, 'Sala principal',  array['Projetor','TV','Webcam','Telefone'], '#dc2626')
) as r(floor, name, capacity, location, equipment, color)
  on r.floor = f.name
where not exists (select 1 from public.rooms);

-- =============================================================================
--  FIM
-- =============================================================================
