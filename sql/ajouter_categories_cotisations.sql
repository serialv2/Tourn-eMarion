-- Pour une installation existante : autoriser la distinction URSSAF / CARPIMKO.
-- Aucune donnée existante n'est modifiée. L'accès reste partagé sans connexion.
begin;
alter table public.marion_expenses drop constraint if exists marion_expenses_category_check;
alter table public.marion_expenses add constraint marion_expenses_category_check
  check (category in ('Prélèvement salaire','Essence','Réparations','Leasing',
    'Assurance','Cotisations','URSSAF','CARPIMKO','Matériel','Autre'));
notify pgrst, 'reload schema';
commit;
