-- =========================================================================
-- SCHEMA 'Manutenzione' (Mezzi + Carrozzine + Struttura) -- versione sicura
-- da rilanciare piu' volte senza errori (idempotente).
-- =========================================================================

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text,
  role text not null default 'lettore' check (role in ('admin','operatore','lettore')),
  created_at timestamptz not null default now()
);
alter table public.profiles enable row level security;

create or replace function public.my_role()
returns text
language sql
security definer
stable
set search_path = public
as $$
  select role from public.profiles where id = auth.uid();
$$;

drop policy if exists "profiles_select_authenticated" on public.profiles;
create policy "profiles_select_authenticated"
  on public.profiles for select
  to authenticated
  using (true);

drop policy if exists "profiles_update_admin_only" on public.profiles;
create policy "profiles_update_admin_only"
  on public.profiles for update
  to authenticated
  using (public.my_role() = 'admin')
  with check (public.my_role() = 'admin');

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email) values (new.id, new.email)
  on conflict (id) do nothing;
  return new;
end;
$$;

create or replace trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- Tabelle dati + permessi

create table if not exists public.vehicles (
  "id" text primary key,
  "targa" text,
  "marca" text,
  "modello" text,
  "tipo" text,
  "anno" integer,
  "km" integer,
  "carburante" text,
  "colore" text,
  "assicurazione" date,
  "revisione" date,
  "bollo" date,
  "note" text
);
alter table public.vehicles enable row level security;
drop policy if exists "vehicles_select_authenticated" on public.vehicles;
create policy "vehicles_select_authenticated"
  on public.vehicles for select
  to authenticated
  using (true);

drop policy if exists "vehicles_insert_admin_operatore" on public.vehicles;
create policy "vehicles_insert_admin_operatore"
  on public.vehicles for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "vehicles_update_admin_operatore" on public.vehicles;
create policy "vehicles_update_admin_operatore"
  on public.vehicles for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "vehicles_delete_admin_only" on public.vehicles;
create policy "vehicles_delete_admin_only"
  on public.vehicles for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.maints (
  "id" text primary key,
  "targa" text references public.vehicles(id) on delete set null,
  "data" date,
  "km" integer,
  "tipo" text,
  "descrizione" text,
  "officina" text,
  "costo" numeric,
  "stato" text
);
alter table public.maints enable row level security;
drop policy if exists "maints_select_authenticated" on public.maints;
create policy "maints_select_authenticated"
  on public.maints for select
  to authenticated
  using (true);

drop policy if exists "maints_insert_admin_operatore" on public.maints;
create policy "maints_insert_admin_operatore"
  on public.maints for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "maints_update_admin_operatore" on public.maints;
create policy "maints_update_admin_operatore"
  on public.maints for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "maints_delete_admin_only" on public.maints;
create policy "maints_delete_admin_only"
  on public.maints for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.carrozzine (
  "id" text primary key,
  "data" date,
  "marca" text,
  "modello" text,
  "seriale" text,
  "tipologia" text,
  "fornitore" text,
  "nucleo" text,
  "stanza" text,
  "ospite" text,
  "stato" text,
  "c" jsonb,
  "note" text
);
alter table public.carrozzine enable row level security;
drop policy if exists "carrozzine_select_authenticated" on public.carrozzine;
create policy "carrozzine_select_authenticated"
  on public.carrozzine for select
  to authenticated
  using (true);

drop policy if exists "carrozzine_insert_admin_operatore" on public.carrozzine;
create policy "carrozzine_insert_admin_operatore"
  on public.carrozzine for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "carrozzine_update_admin_operatore" on public.carrozzine;
create policy "carrozzine_update_admin_operatore"
  on public.carrozzine for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "carrozzine_delete_admin_only" on public.carrozzine;
create policy "carrozzine_delete_admin_only"
  on public.carrozzine for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.camere (
  "codice" text primary key,
  "piano" text,
  "nucleo" text,
  "tipo" text,
  "stato" text,
  "note" text
);
alter table public.camere enable row level security;
drop policy if exists "camere_select_authenticated" on public.camere;
create policy "camere_select_authenticated"
  on public.camere for select
  to authenticated
  using (true);

drop policy if exists "camere_insert_admin_operatore" on public.camere;
create policy "camere_insert_admin_operatore"
  on public.camere for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "camere_update_admin_operatore" on public.camere;
create policy "camere_update_admin_operatore"
  on public.camere for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "camere_delete_admin_only" on public.camere;
create policy "camere_delete_admin_only"
  on public.camere for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.reparti (
  "codice" text primary key,
  "nome" text,
  "categoria" text,
  "responsabile" text,
  "note" text
);
alter table public.reparti enable row level security;
drop policy if exists "reparti_select_authenticated" on public.reparti;
create policy "reparti_select_authenticated"
  on public.reparti for select
  to authenticated
  using (true);

drop policy if exists "reparti_insert_admin_operatore" on public.reparti;
create policy "reparti_insert_admin_operatore"
  on public.reparti for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "reparti_update_admin_operatore" on public.reparti;
create policy "reparti_update_admin_operatore"
  on public.reparti for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "reparti_delete_admin_only" on public.reparti;
create policy "reparti_delete_admin_only"
  on public.reparti for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.tecnici (
  "id" text primary key,
  "nome" text,
  "tipo" text,
  "specializzazione" text,
  "telefono" text,
  "email" text,
  "note" text
);
alter table public.tecnici enable row level security;
drop policy if exists "tecnici_select_authenticated" on public.tecnici;
create policy "tecnici_select_authenticated"
  on public.tecnici for select
  to authenticated
  using (true);

drop policy if exists "tecnici_insert_admin_operatore" on public.tecnici;
create policy "tecnici_insert_admin_operatore"
  on public.tecnici for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "tecnici_update_admin_operatore" on public.tecnici;
create policy "tecnici_update_admin_operatore"
  on public.tecnici for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "tecnici_delete_admin_only" on public.tecnici;
create policy "tecnici_delete_admin_only"
  on public.tecnici for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.interventi (
  "id" text primary key,
  "dataSegnalazione" date,
  "cameraZona" text,
  "descrizione" text,
  "priorita" text,
  "tecnico" text,
  "stato" text,
  "dataChiusura" date,
  "costo" numeric,
  "note" text
);
alter table public.interventi enable row level security;
drop policy if exists "interventi_select_authenticated" on public.interventi;
create policy "interventi_select_authenticated"
  on public.interventi for select
  to authenticated
  using (true);

drop policy if exists "interventi_insert_admin_operatore" on public.interventi;
create policy "interventi_insert_admin_operatore"
  on public.interventi for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "interventi_update_admin_operatore" on public.interventi;
create policy "interventi_update_admin_operatore"
  on public.interventi for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "interventi_delete_admin_only" on public.interventi;
create policy "interventi_delete_admin_only"
  on public.interventi for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.manutenzioni (
  "id" text primary key,
  "cameraZona" text,
  "tipoManutenzione" text,
  "frequenza" text,
  "ultimaEsecuzione" date,
  "prossimaScadenza" date,
  "tecnico" text,
  "note" text
);
alter table public.manutenzioni enable row level security;
drop policy if exists "manutenzioni_select_authenticated" on public.manutenzioni;
create policy "manutenzioni_select_authenticated"
  on public.manutenzioni for select
  to authenticated
  using (true);

drop policy if exists "manutenzioni_insert_admin_operatore" on public.manutenzioni;
create policy "manutenzioni_insert_admin_operatore"
  on public.manutenzioni for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "manutenzioni_update_admin_operatore" on public.manutenzioni;
create policy "manutenzioni_update_admin_operatore"
  on public.manutenzioni for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "manutenzioni_delete_admin_only" on public.manutenzioni;
create policy "manutenzioni_delete_admin_only"
  on public.manutenzioni for delete
  to authenticated
  using (public.my_role() = 'admin');


create table if not exists public.costi (
  "id" text primary key,
  "idIntervento" text,
  "tipo" text,
  "descrizione" text,
  "fornitore" text,
  "numeroDocumento" text,
  "data" date,
  "importo" numeric,
  "statoPagamento" text,
  "note" text
);
alter table public.costi enable row level security;
drop policy if exists "costi_select_authenticated" on public.costi;
create policy "costi_select_authenticated"
  on public.costi for select
  to authenticated
  using (true);

drop policy if exists "costi_insert_admin_operatore" on public.costi;
create policy "costi_insert_admin_operatore"
  on public.costi for insert
  to authenticated
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "costi_update_admin_operatore" on public.costi;
create policy "costi_update_admin_operatore"
  on public.costi for update
  to authenticated
  using (public.my_role() in ('admin','operatore'))
  with check (public.my_role() in ('admin','operatore'));

drop policy if exists "costi_delete_admin_only" on public.costi;
create policy "costi_delete_admin_only"
  on public.costi for delete
  to authenticated
  using (public.my_role() = 'admin');

-- Sincronizzazione in tempo reale (sicura da rilanciare piu' volte)
do $$
declare
  t text;
begin
  foreach t in array array['vehicles','maints','carrozzine','camere','reparti','tecnici','interventi','manutenzioni','costi']
  loop
    if not exists (
      select 1 from pg_publication_tables
      where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = t
    ) then
      execute format('alter publication supabase_realtime add table public.%I', t);
    end if;
  end loop;
end $$;
