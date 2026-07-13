-- ABLAB Sales Report V3 - Phase 1
-- Safe, transactional sales storage. Existing public.app_state is preserved.

create extension if not exists pgcrypto;

create table if not exists public.sales_upload_batches (
  id uuid primary key default gen_random_uuid(),
  period text not null check (period in ('current','prevMonth','prevYear')),
  ref_month text not null,
  file_name text,
  row_count integer not null default 0,
  uploaded_dates date[] not null default '{}',
  status text not null default 'success' check (status in ('success','rolled_back')),
  replaced_batch_ids uuid[] not null default '{}',
  created_at timestamptz not null default now(),
  rolled_back_at timestamptz
);

create table if not exists public.sales_records (
  id uuid primary key default gen_random_uuid(),
  row_key text not null,
  period text not null check (period in ('current','prevMonth','prevYear')),
  ref_month text not null,
  sale_date date not null,
  store_code text not null,
  store_name text not null default '',
  channel text not null default '',
  manager text not null default '',
  store_type text not null default '',
  brand text not null default '',
  item_code text not null default '',
  item_name text not null default '',
  quantity numeric not null default 0,
  sales_amount numeric not null default 0,
  cost_amount numeric not null default 0,
  profit_amount numeric not null default 0,
  profit_rate numeric not null default 0,
  batch_id uuid not null references public.sales_upload_batches(id),
  active boolean not null default true,
  created_at timestamptz not null default now()
);

create index if not exists sales_records_ref_month_active_idx
  on public.sales_records(ref_month, period, active);
create index if not exists sales_records_sale_date_active_idx
  on public.sales_records(sale_date, active);
create index if not exists sales_records_store_code_idx
  on public.sales_records(store_code);
create index if not exists sales_records_item_code_idx
  on public.sales_records(item_code);
create index if not exists sales_records_batch_id_idx
  on public.sales_records(batch_id);
create unique index if not exists sales_records_active_row_key_uq
  on public.sales_records(row_key)
  where active = true;

alter table public.sales_upload_batches enable row level security;
alter table public.sales_records enable row level security;

drop policy if exists "sales batches read" on public.sales_upload_batches;
create policy "sales batches read" on public.sales_upload_batches
  for select to anon, authenticated using (true);

drop policy if exists "sales records read" on public.sales_records;
create policy "sales records read" on public.sales_records
  for select to anon, authenticated using (active = true);

create or replace function public.replace_sales_batch(
  p_period text,
  p_ref_month text,
  p_file_name text,
  p_uploaded_dates date[],
  p_rows jsonb
) returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_batch_id uuid := gen_random_uuid();
  v_replaced uuid[] := '{}';
  v_count integer := 0;
begin
  if p_period not in ('current','prevMonth','prevYear') then
    raise exception 'invalid period: %', p_period;
  end if;
  if p_ref_month !~ '^\\d{4}-\\d{2}$' then
    raise exception 'invalid ref_month: %', p_ref_month;
  end if;
  if jsonb_typeof(p_rows) <> 'array' then
    raise exception 'p_rows must be a json array';
  end if;

  select coalesce(array_agg(distinct batch_id), '{}')
    into v_replaced
  from public.sales_records
  where active = true
    and period = p_period
    and ref_month = p_ref_month
    and (
      p_period <> 'current'
      or sale_date = any(coalesce(p_uploaded_dates, '{}'))
    );

  update public.sales_records
     set active = false
   where active = true
     and period = p_period
     and ref_month = p_ref_month
     and (
       p_period <> 'current'
       or sale_date = any(coalesce(p_uploaded_dates, '{}'))
     );

  insert into public.sales_upload_batches(
    id, period, ref_month, file_name, row_count, uploaded_dates,
    status, replaced_batch_ids
  ) values (
    v_batch_id, p_period, p_ref_month, p_file_name,
    jsonb_array_length(p_rows), coalesce(p_uploaded_dates, '{}'),
    'success', v_replaced
  );

  insert into public.sales_records(
    row_key, period, ref_month, sale_date,
    store_code, store_name, channel, manager, store_type, brand,
    item_code, item_name, quantity, sales_amount, cost_amount,
    profit_amount, profit_rate, batch_id, active
  )
  select
    x.id,
    x.period,
    x.ref_month,
    x.sale_date::date,
    x.store_code,
    coalesce(x.store_name, ''),
    coalesce(x.channel, ''),
    coalesce(x.manager, ''),
    coalesce(x.store_type, ''),
    coalesce(x.brand, ''),
    coalesce(x.item_code, ''),
    coalesce(x.item_name, ''),
    coalesce(x.quantity, 0),
    coalesce(x.sales_amount, 0),
    coalesce(x.cost_amount, 0),
    coalesce(x.profit_amount, 0),
    coalesce(x.profit_rate, 0),
    v_batch_id,
    true
  from jsonb_to_recordset(p_rows) as x(
    id text,
    period text,
    ref_month text,
    sale_date text,
    store_code text,
    store_name text,
    channel text,
    manager text,
    store_type text,
    brand text,
    item_code text,
    item_name text,
    quantity numeric,
    sales_amount numeric,
    cost_amount numeric,
    profit_amount numeric,
    profit_rate numeric
  );

  get diagnostics v_count = row_count;
  return jsonb_build_object(
    'ok', true,
    'batch_id', v_batch_id,
    'row_count', v_count,
    'replaced_batch_ids', v_replaced
  );
end;
$$;

create or replace function public.delete_current_sales_date(
  p_ref_month text,
  p_sale_date date
) returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_count integer := 0;
begin
  update public.sales_records
     set active = false
   where active = true
     and period = 'current'
     and ref_month = p_ref_month
     and sale_date = p_sale_date;
  get diagnostics v_count = row_count;
  return jsonb_build_object('ok', true, 'deleted_count', v_count);
end;
$$;

grant execute on function public.replace_sales_batch(text,text,text,date[],jsonb) to anon, authenticated;
grant execute on function public.delete_current_sales_date(text,date) to anon, authenticated;
grant select on public.sales_upload_batches to anon, authenticated;
grant select on public.sales_records to anon, authenticated;

comment on table public.sales_records is 'ABLAB Sales V3 row-level storage. Old rows are retained with active=false.';
comment on table public.sales_upload_batches is 'Atomic upload history for ABLAB Sales V3.';
