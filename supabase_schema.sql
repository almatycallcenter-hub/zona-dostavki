-- ═══════════════════════════════════════════════════════════
-- Тәп-Тәтті · общее хранилище для callcenter.html
-- Вставьте целиком в Supabase → SQL Editor → Run
-- ═══════════════════════════════════════════════════════════

-- Настройки: адреса филиалов, ставки KPI, ручные связки Wolt,
-- ручные нарушения. Одна строка на ключ, значение — JSON.
create table if not exists app_state (
  key        text primary key,
  value      jsonb       not null,
  updated_at timestamptz not null default now(),
  updated_by text
);

-- Загруженные отчёты: чтобы с другого ноутбука не загружать файлы заново
create table if not exists datasets (
  id         bigint generated always as identity primary key,
  kind       text not null,          -- orders | wolt | invoices
  label      text,
  date_from  date,
  date_to    date,
  rows       int,
  payload    jsonb not null,
  created_at timestamptz not null default now(),
  created_by text
);
create index if not exists datasets_kind_created on datasets (kind, created_at desc);

-- ── Доступ ────────────────────────────────────────────────
-- Читать и писать может только тот, кто вошёл по логину и паролю.
-- Анонимный ключ сам по себе ничего не открывает.
alter table app_state enable row level security;
alter table datasets  enable row level security;

drop policy if exists app_state_read   on app_state;
drop policy if exists app_state_write  on app_state;
drop policy if exists app_state_update on app_state;
drop policy if exists datasets_read    on datasets;
drop policy if exists datasets_write   on datasets;
drop policy if exists datasets_delete  on datasets;

create policy app_state_read   on app_state for select to authenticated using (true);
create policy app_state_write  on app_state for insert to authenticated with check (true);
create policy app_state_update on app_state for update to authenticated using (true) with check (true);

create policy datasets_read    on datasets for select to authenticated using (true);
create policy datasets_write   on datasets for insert to authenticated with check (true);
create policy datasets_delete  on datasets for delete to authenticated using (true);
