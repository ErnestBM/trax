-- Trax: personal finance tracker. Lives in its own `trax` schema inside the leanislaw
-- Supabase project so leanislaw's Prisma (which owns `public`) never sees these tables
-- as drift. After running this, add `trax` to Project Settings → Data API → Exposed schemas.

create schema if not exists trax;

-- Month boundaries follow Jakarta time (WIB), not the server's UTC clock.
create or replace function trax.current_month() returns date
language sql stable as $$
  select date_trunc('month', now() at time zone 'Asia/Jakarta')::date
$$;

-- Pockets -------------------------------------------------------------------------------
-- A pocket with a parent is a sub-pocket of the same bank account (e.g. Kos inside BCA):
-- the account balance is parent + children. Only top-level pockets count in the total balance.
create table trax.pockets (
  id               uuid primary key default gen_random_uuid(),
  user_id          uuid not null default auth.uid() references auth.users on delete cascade,
  name             text not null,
  parent_id        uuid references trax.pockets on delete set null,
  sort             int not null default 0,
  archived         boolean not null default false,
  created_at       timestamptz not null default now()
);

-- Categories ----------------------------------------------------------------------------
-- Never hard-deleted: archiving keeps old transactions and locked budgets intact.
create table trax.categories (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users on delete cascade,
  name       text not null,
  kind       text not null default 'expense' check (kind in ('expense', 'income')),
  sort       int not null default 0,
  archived   boolean not null default false,
  created_at timestamptz not null default now()
);

-- Transactions --------------------------------------------------------------------------
-- One transaction holds one or more signed lines, one per pocket it touches — the sheet's
-- Kredit/Debit columns on a single row. The sum of the lines (net) is what the
-- transaction did to total wealth: an expense is negative, income positive, and a
-- transfer nets to minus its admin fee. Budget usage is measured on that net.
create table trax.transactions (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null default auth.uid() references auth.users on delete cascade,
  date        date not null default (now() at time zone 'Asia/Jakarta')::date,
  name        text not null,
  -- Null = not counted against any budget (transfers, reimbursements, lending).
  category_id uuid references trax.categories,
  created_at  timestamptz not null default now()
);
create index on trax.transactions (user_id, date);

create table trax.transaction_lines (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null default auth.uid() references auth.users on delete cascade,
  transaction_id uuid not null references trax.transactions on delete cascade,
  pocket_id      uuid not null references trax.pockets,
  amount         numeric(16, 2) not null check (amount <> 0)
);
create index on trax.transaction_lines (transaction_id);
create index on trax.transaction_lines (pocket_id);

create view trax.pocket_balances with (security_invoker = true) as
  select p.id as pocket_id, coalesce(sum(l.amount), 0)::numeric(16, 2) as balance
  from trax.pockets p
  left join trax.transaction_lines l on l.pocket_id = p.id
  group by p.id;

-- AR / AP -------------------------------------------------------------------------------
-- AR: money owed to me (fronted payments, loans). AP: money I hold for or owe others.
create table trax.debts (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users on delete cascade,
  kind       text not null check (kind in ('AR', 'AP')),
  date       date not null default (now() at time zone 'Asia/Jakarta')::date,
  name       text not null,
  amount     numeric(16, 2) not null check (amount > 0),
  settled_at date,
  created_at timestamptz not null default now()
);

-- Budgets -------------------------------------------------------------------------------
-- One period per month. A period is locked once its month has started: changes made
-- today go into next month's draft. The trigger below enforces that; ensure_budget is
-- the only path that may write a current-month period (by copying the previous one).
create table trax.budget_periods (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null default auth.uid() references auth.users on delete cascade,
  month       date not null check (extract(day from month) = 1),
  income      numeric(16, 2) not null default 0,
  target_eomb numeric(16, 2) not null default 0,
  created_at  timestamptz not null default now(),
  unique (user_id, month)
);

create table trax.budget_items (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null default auth.uid() references auth.users on delete cascade,
  period_id   uuid not null references trax.budget_periods on delete cascade,
  category_id uuid not null references trax.categories,
  price       numeric(16, 2) not null default 0,
  unit        text not null default 'Bulan',
  qty         numeric(10, 2) not null default 1,
  sort        int not null default 0,
  unique (period_id, category_id)
);

create or replace function trax.guard_budget_lock() returns trigger
language plpgsql as $$
declare
  m date;
begin
  if current_setting('trax.bypass_lock', true) = 'on' then
    return case when tg_op = 'DELETE' then old else new end;
  end if;
  if tg_table_name = 'budget_periods' then
    m := case when tg_op = 'INSERT' then new.month else old.month end;
  else
    select month into m from trax.budget_periods
    where id = case when tg_op = 'DELETE' then old.period_id else new.period_id end;
  end if;
  if m <= trax.current_month() then
    raise exception 'The % budget is locked. Edit next month''s draft instead.', to_char(m, 'Mon YYYY');
  end if;
  return case when tg_op = 'DELETE' then old else new end;
end
$$;

create trigger budget_periods_lock before insert or update or delete on trax.budget_periods
  for each row execute function trax.guard_budget_lock();
create trigger budget_items_lock before insert or update or delete on trax.budget_items
  for each row execute function trax.guard_budget_lock();

-- Returns the period for p_month, creating it as a copy of the latest earlier period if
-- missing. This is how a month "locks on the 1st": whatever draft exists becomes the
-- month's budget, and if none was made, last month's budget carries over unchanged.
create or replace function trax.ensure_budget(p_month date) returns uuid
language plpgsql as $$
declare
  pid uuid;
  src trax.budget_periods;
begin
  select id into pid from trax.budget_periods where user_id = auth.uid() and month = p_month;
  if pid is not null then
    return pid;
  end if;
  if p_month > (trax.current_month() + interval '1 month')::date then
    raise exception 'Budgets can only be created up to next month';
  end if;

  perform set_config('trax.bypass_lock', 'on', true);
  select * into src from trax.budget_periods
  where user_id = auth.uid() and month < p_month order by month desc limit 1;

  insert into trax.budget_periods (month, income, target_eomb)
  values (p_month, coalesce(src.income, 0), coalesce(src.target_eomb, 0))
  returning id into pid;

  if src.id is not null then
    insert into trax.budget_items (period_id, category_id, price, unit, qty, sort)
    select pid, bi.category_id, bi.price, bi.unit, bi.qty, bi.sort
    from trax.budget_items bi
    join trax.categories c on c.id = bi.category_id
    where bi.period_id = src.id and not c.archived;
  end if;
  perform set_config('trax.bypass_lock', 'off', true);
  return pid;
end
$$;

-- Creates (p_id null) or replaces a transaction and all its lines atomically.
-- p_lines: [{"pocket_id": uuid, "amount": signed number}, ...]
create or replace function trax.save_transaction(
  p_id uuid, p_date date, p_name text, p_category uuid, p_lines jsonb
) returns uuid
language plpgsql as $$
declare
  tid uuid := p_id;
begin
  if jsonb_array_length(coalesce(p_lines, '[]')) = 0 then
    raise exception 'A transaction needs at least one pocket';
  end if;
  if tid is null then
    insert into trax.transactions (date, name, category_id)
    values (p_date, p_name, p_category) returning id into tid;
  else
    update trax.transactions set date = p_date, name = p_name, category_id = p_category
    where id = tid;
    if not found then
      raise exception 'Transaction not found';
    end if;
    delete from trax.transaction_lines where transaction_id = tid;
  end if;
  insert into trax.transaction_lines (transaction_id, pocket_id, amount)
  select tid, (x->>'pocket_id')::uuid, (x->>'amount')::numeric
  from jsonb_array_elements(p_lines) x;
  return tid;
end
$$;

-- Access: every row belongs to its auth user. Functions run as the caller, so RLS applies.
do $$
declare
  t text;
begin
  foreach t in array array['pockets', 'categories', 'transactions', 'transaction_lines',
                           'debts', 'budget_periods', 'budget_items'] loop
    execute format('alter table trax.%I enable row level security', t);
    execute format('create policy own_rows on trax.%I for all to authenticated
                    using (user_id = auth.uid()) with check (user_id = auth.uid())', t);
  end loop;
end
$$;

grant usage on schema trax to authenticated, service_role;
grant select, insert, update, delete on all tables in schema trax to authenticated;
grant all on all tables in schema trax to service_role;
grant execute on all functions in schema trax to authenticated, service_role;
