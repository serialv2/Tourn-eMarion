-- À exécuter une fois dans le SQL Editor du projet Tournée Marion.
-- Accès partagé sans connexion, demandé explicitement par le propriétaire.
-- Les visiteurs peuvent lire, ajouter, modifier et supprimer les dépenses.
-- Ce script ne modifie ni les patients, ni les tournées, ni les journées.
begin;

create table if not exists public.marion_expenses (
  id uuid primary key default gen_random_uuid(),
  expense_date date not null,
  category text not null check (category in (
    'Prélèvement salaire', 'Essence', 'Réparations', 'Leasing',
    'Assurance', 'Cotisations', 'URSSAF', 'CARPIMKO', 'Matériel', 'Autre'
  )),
  label text not null check (length(btrim(label)) between 1 and 200),
  amount numeric(11,2) not null check (amount > 0 and amount <= 999999999.99),
  created_at timestamptz not null default now()
);

create index if not exists marion_expenses_date_id_idx on public.marion_expenses (expense_date, id);
alter table public.marion_expenses enable row level security;
revoke all on public.marion_expenses from anon, authenticated;
grant select, insert, update, delete on public.marion_expenses to anon, authenticated;

drop policy if exists marion_expenses_shared_read on public.marion_expenses;
create policy marion_expenses_shared_read on public.marion_expenses
  for select to anon, authenticated using (true);
drop policy if exists marion_expenses_shared_insert on public.marion_expenses;
create policy marion_expenses_shared_insert on public.marion_expenses
  for insert to anon, authenticated with check (true);
drop policy if exists marion_expenses_shared_update on public.marion_expenses;
create policy marion_expenses_shared_update on public.marion_expenses
  for update to anon, authenticated using (true) with check (true);
drop policy if exists marion_expenses_shared_delete on public.marion_expenses;
create policy marion_expenses_shared_delete on public.marion_expenses
  for delete to anon, authenticated using (true);

notify pgrst, 'reload schema';
commit;

-- Vérification sans lecture des données : table, RLS et quatre règles d'accès.
select c.relname, c.relrowsecurity from pg_class c
join pg_namespace n on n.oid = c.relnamespace
where n.nspname = 'public' and c.relname = 'marion_expenses';
select policyname, cmd, roles from pg_policies
where schemaname = 'public' and tablename = 'marion_expenses';
