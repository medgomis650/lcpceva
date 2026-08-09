-- Run this once in Supabase: Project > SQL Editor > New query > paste & Run.

create table if not exists kv_store (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz default now()
);

alter table kv_store enable row level security;

-- NOTE: these policies allow anyone with your anon public key (i.e. anyone
-- who has your deployed app's URL, since the key is embedded client-side)
-- to read and write this table. There is no login/authentication in this
-- version of the app. This is fine for an internal tool on a private link,
-- but do not treat the URL as public. Ask if you'd like real user accounts
-- added later.
create policy "kv_store select" on kv_store for select using (true);
create policy "kv_store insert" on kv_store for insert with check (true);
create policy "kv_store update" on kv_store for update using (true);
