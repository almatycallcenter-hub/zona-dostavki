-- ═══════════════════════════════════════════════════════════
-- Чеки за Яндекс и Индрайвер, которые присылает телеграм-бот
-- Supabase → SQL Editor → Run (после supabase_schema.sql)
-- ═══════════════════════════════════════════════════════════

create table if not exists receipts (
  id          bigint generated always as identity primary key,
  service     text,                    -- Яндекс | Индрайвер | ''
  paid_date   date,                    -- дата с чека
  paid_time   text,                    -- ЧЧ:ММ с чека
  amount      numeric,                 -- сколько отдали за курьера
  order_num   text,                    -- № заказа из iiko, если указали
  status      text not null default 'pending',   -- pending | ready
  note        text,
  file_path   text,                    -- сам чек в Storage
  raw_text    text,                    -- что вытащили из PDF, на случай разбора
  tg_chat_id  bigint,
  tg_user     text,
  created_at  timestamptz not null default now()
);
create index if not exists receipts_date on receipts (paid_date desc);
create index if not exists receipts_order on receipts (order_num);

alter table receipts enable row level security;

drop policy if exists receipts_read   on receipts;
drop policy if exists receipts_write  on receipts;
drop policy if exists receipts_update on receipts;
drop policy if exists receipts_delete on receipts;

create policy receipts_read   on receipts for select to authenticated using (true);
create policy receipts_write  on receipts for insert to authenticated with check (true);
create policy receipts_update on receipts for update to authenticated using (true) with check (true);
create policy receipts_delete on receipts for delete to authenticated using (true);

-- Хранилище для самих чеков. Приватное: файлы отдаются только по токену.
insert into storage.buckets (id, name, public)
values ('receipts', 'receipts', false)
on conflict (id) do nothing;

drop policy if exists receipts_files_read on storage.objects;
create policy receipts_files_read on storage.objects for select
  to authenticated using (bucket_id = 'receipts');
