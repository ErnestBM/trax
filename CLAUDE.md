# Trax

Personal finance tracker replacing the "Financial Statement" Google Sheet (monthly `Oct 26`-style tracking tabs + `RAB Bulanan`). iPhone 17 PWA (402×874), same deploy/PWA pattern as `../leanislaw` (pinned shell, `status-bar-style: black`, safe areas, 16px inputs). UI copy and data model are English; the user writes transaction names in Indonesian.

Same staging as `../knhi` (`knhi-mvp/` prototype → `web/` Next.js):

- **`mvp/index.html` — the prototype (current).** One self-contained file: Preact + htm from jsdelivr, data in localStorage (`trax.mvp.v4`). The `store` object mirrors the Supabase schema and rules (budget lock, ensure-period copy, signed lines). Seeded from the embedded `BACKFILL` (real October data); with `BACKFILL = null` it falls back to dummy data. The "Prototype" button on Balance resets data and sets a simulated today (`trax.mvp.today`). `window.__trax` exposes the store for the console. Run: `python3 -m http.server -d mvp 8000`. **Contains real financial data — never deploy it publicly.**
- **`app/` — the real Next.js app (not started; only when the user says so).** Will use `supabase/001_init.sql`.
- **`backfill/build_backfill.py`** — reads a tracking tab + `RAB Bulanan` from the xlsx, applies the categorisation agreed with the user (rules and per-row corrections are in the script), asserts every pocket closes on the sheet's balance and open AR equals RAB's total, writes `backfill/<tab>.json`; `--embed` also writes it into `mvp/index.html`. Run: `python3 -I backfill/build_backfill.py "Untitled spreadsheet.xlsx" "Oct 26" --embed`.

## Setup (for `app/`)
- DB: the leanislaw Supabase project (`pruvwucsxwkqjopbmvjv`), but tables live in schema **`trax`** so leanislaw's Prisma (owns `public`) never reports them as drift. Run `supabase/001_init.sql` in the SQL editor, then add `trax` to Project Settings → Data API → Exposed schemas.
- Auth: Supabase Auth email/password. Create the user in the dashboard (Authentication → Users); no sign-up screen.
- Env: `.env` has `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, `SUPABASE_SECRET_KEY`. Only the first two may reach the browser; the secret key is for server code and backfill scripts.

## Model
- **pockets** — `parent_id` groups sub-pockets into one bank account (Kos inside BCA → "BCA Total" = BCA + Kos). The total balance counts top-level pockets only, like the sheet's Saldo Total (BCA → Ovo): a sub-pocket is set-aside money, shown under its account's total but not counted. No per-pocket toggle.
- **transactions + transaction_lines** — one transaction, N signed lines (one per pocket), mirroring the sheet's Kredit/Debit columns on a row. `net = Σ lines` is the effect on wealth; budget usage is measured on net. A transfer is exactly one negative + one positive line where |neg| ≥ pos, the difference being the admin fee; transfers are named "Top up from {from} to {to}" and have no category. Input types are Expense / Income / Transfer (Custom only appears when editing an older multi-pocket row). Saved atomically via `trax.save_transaction`.
- **pocket_balances** view — Σ lines per pocket. Opening balances are "Opening balance"/"Balance adjustment" transactions, not a column.
- **categories** — `expense|income|debt`, archived not deleted. Every expense/income has one; **Others** is the default. Only transfers and balance adjustments have none. Chip order is the user's: Office lunch, Office transport, Gas/toll, Date, Hobby, Spotify, Gym, Claude, Groceries, Lifestyle, Health, Admin, Kos, Others.
- **budget_periods / budget_items** — one period per month, items = category × price × qty × unit. A period whose month has started is **locked** (trigger `guard_budget_lock`); edits go to next month's draft. `trax.ensure_budget(month)` creates a missing period by copying the latest earlier one — that copy is how "locks on the 1st" works without a cron.
- **debts** — AR (receivables) / AP (payables), `settled_at` when done. Each debt links up to two transactions (`open_tx_id`, `settle_tx_id`) in the system categories **AR** / **AP** (kind `debt`, excluded from spending, income and budgets): AR opens as a debit from the chosen pocket and settles as a credit, AP the reverse; "Not sure" records no transaction. The RAB's AR amount is the source of truth even when the money that left differs (Oct 21 dinner: AR 32.000, ShopeePay −27.000), so editing a debt keeps its lines unless amount, kind or pocket change. Balance screen shows Balance, −AP, +AR, −AP+AR. *Not yet in `supabase/001_init.sql`* (needs the two FK columns and `kind` allowing `debt`).
- Expected EOMB (Budget → current month) = total balance + remaining income + AR − Σ max(0, budget − spent) − AP, vs `target_eomb`. This intentionally differs from the sheet's Expected EOMB: the sheet's "Expenses" there was a hand-entered stand-in for budget minus actual spending, which the budget bars now compute (user, 2026-10-28).

## Open
- Backfill of earlier months from the full xlsx: the user will send it; extend `build_backfill.py` (row overrides are per tab).
- `supabase/001_init.sql` lags the MVP: debt categories/links, Others default.
