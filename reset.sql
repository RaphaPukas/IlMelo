-- =========================================================================
-- RESET COMPLETO — da usare SOLO se vuoi ripartire da zero (cancella tutte
-- le tabelle e i relativi dati creati dallo schema "Manutenzione").
-- Dopo averlo eseguito, rilancia schema_idempotente.sql per ricreare tutto.
-- =========================================================================

drop table if exists public.costi cascade;
drop table if exists public.manutenzioni cascade;
drop table if exists public.interventi cascade;
drop table if exists public.tecnici cascade;
drop table if exists public.reparti cascade;
drop table if exists public.camere cascade;
drop table if exists public.carrozzine cascade;
drop table if exists public.maints cascade;
drop table if exists public.vehicles cascade;

drop trigger if exists on_auth_user_created on auth.users;
drop function if exists public.handle_new_user();
drop function if exists public.my_role();
drop table if exists public.profiles cascade;
