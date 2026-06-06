[README.md](https://github.com/user-attachments/files/28668107/README.md)
# Eudai ControlCenter

The founders' operating system for Eudai — tasks, product, financials, strategy, meetings, and team discussion in one app. Built React + Supabase + Vercel.

## Cost: $0/month

| Service | Plan | Cost |
|---|---|---|
| Vercel (hosting + the AI proxy) | Hobby | **$0** |
| Supabase (database + logins) | Free | **$0** |
| Anthropic API (copilot + AI meeting notes) | pay-as-you-go | **optional** — skip the key entirely and the app runs in offline-AI mode (built-in fallback answers). With a key, expect ~$5–15/mo at founder-level usage; set a spend limit in the Anthropic console. |
| Custom domain | optional | ~$10–12/yr, or use the free `*.vercel.app` URL |

Hard ceiling if you turn everything on: well under $50 for months.

## Deploy (about 45 minutes, one founder does this once)

### 1. Supabase — database + logins (15 min)

1. [supabase.com](https://supabase.com) → New project → name `eudai-controlcenter`. Note the **Project URL** and **anon public key** (Settings → API).
2. SQL Editor → New query → paste & run:

```sql
create table kv (
  key text primary key,
  value text not null,
  updated_at timestamptz default now()
);

alter table kv enable row level security;

create policy "founders read"   on kv for select to authenticated using (true);
create policy "founders insert" on kv for insert to authenticated with check (true);
create policy "founders update" on kv for update to authenticated using (true);
create policy "founders delete" on kv for delete to authenticated using (true);
```

3. **Security — do not skip:** Authentication → Sign In / Up → turn **OFF** "Allow new users to sign up". Then Authentication → Users → **Add user → Send invitation** to each founder's email (Grayson, Phoenix, Dylan, Marcellus). Only invited accounts can ever log in.

### 2. Run locally (10 min)

```bash
npm install
cp .env.example .env.local     # paste your Supabase URL + anon key into .env.local
npm run dev
```

Open the printed localhost URL → magic-link login screen → sign in with an invited email. (Local AI calls 404 without Vercel; the offline fallbacks answer instead. AI goes live after deploy.)

### 3. Deploy to Vercel (15 min)

```bash
git init && git add -A && git commit -m "Eudai ControlCenter v1"
```

Create an empty GitHub repo named `eudai-controlcenter`, then:

```bash
git remote add origin https://github.com/YOURUSERNAME/eudai-controlcenter.git
git push -u origin main
```

1. [vercel.com](https://vercel.com) → Add New → Project → import the repo (Vite auto-detected).
2. Environment Variables (Settings → Environment Variables):
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `ANTHROPIC_API_KEY` ← optional; add later to turn AI on
3. Deploy → you get `https://eudai-controlcenter.vercel.app`.
4. Back in Supabase: Authentication → URL Configuration → set **Site URL** to that Vercel URL (magic links redirect there).

### 4. Move your data over (5 min)

In the Claude artifact: Settings → Backup & Migration → **Export workspace** (downloads a JSON).
In the deployed app: Settings → Backup & Migration → **Import backup** → pick the file. Everything — tasks, financials, vendors, the lot — lands in production.

### 5. Founder test checklist

- [ ] All four founders get magic links and can log in
- [ ] A task added on one laptop appears on another after refresh
- [ ] Discussion messages flow between two founders
- [ ] Copilot answers with live numbers (if API key set) or labeled offline answers (if not)
- [ ] Statements → Export to Excel downloads the .xlsx
- [ ] Settings → Export workspace downloads a backup (do this weekly — it's your undo button)

## Architecture (30-second version)

- `src/Eudai_ControlCenter.jsx` — the entire app, single file by design
- `src/storage.js` — recreates the artifact storage API on a Supabase `kv` table; the app's persistence code is unchanged
- `src/AuthGate.jsx` — magic-link login; the logged-in name feeds the app's existing personalization
- `api/claude.js` — server-side AI proxy; holds the Anthropic key, verifies the caller's session before forwarding, returns 503 (→ graceful offline mode) when no key is set

## Later, when it matters

- **Realtime sync** — swap Discussion's 7s polling for Supabase Realtime
- **Per-record tables** — split the kv blob into `tasks`/`vendors`/`investors` tables (see StartupOS-Architecture.md)
- **Per-founder permissions** — RLS keyed to user IDs; the Leads name-arrays map 1:1 to accounts
- **Mobile pass** — the app is desktop-first; tables and the 5-column stat strip get cramped on phones
