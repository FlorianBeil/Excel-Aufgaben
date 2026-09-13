-- Excel.Flo – anonyme Nutzungs-Ereignisse der Übungsportale (Funktionen + Pivot-Tabellen)
--
-- Einmalig im Supabase-Dashboard ausführen:
--   SQL Editor → New query → dieses Skript einfügen → Run
-- Das Skript kann gefahrlos mehrfach ausgeführt werden.
--
-- Datenschutz: keine E-Mail, keine Nutzer-ID – nur eine zufällige ID pro Browser-Sitzung.
-- Die Portale dürfen Ereignisse nur SCHREIBEN. Lesen geht ausschließlich im Dashboard.

create table if not exists public.events (
  id          bigint generated always as identity primary key,
  created_at  timestamptz not null default now(),
  portal      text not null check (portal in ('funktionen', 'pivot', 'powerquery')),
  exercise_id text check (char_length(exercise_id) <= 100),
  event       text not null check (char_length(event) <= 40),
  detail      jsonb not null default '{}'::jsonb check (pg_column_size(detail) <= 2000),
  session_id  text check (char_length(session_id) <= 64)
);

create index if not exists events_created_at_idx on public.events (created_at);
create index if not exists events_exercise_idx on public.events (portal, exercise_id, event);

alter table public.events enable row level security;

-- Nur Einfügen erlaubt – kein Lesen, Ändern oder Löschen über die öffentliche API.
drop policy if exists "events_insert_only" on public.events;
create policy "events_insert_only" on public.events
  for insert to anon, authenticated
  with check (true);

revoke all on public.events from anon, authenticated;
grant insert on public.events to anon, authenticated;

-- Auswertung pro Übung. Im Dashboard unter Table Editor → events_uebersicht ansehen
-- (security_invoker + revoke: über die öffentliche API nicht lesbar).
create or replace view public.events_uebersicht
with (security_invoker = true) as
select
  portal,
  exercise_id,
  count(distinct session_id) filter (where event = 'exercise_open')                              as sitzungen_geoeffnet,
  count(*)                   filter (where event = 'check')                                      as pruefungen,
  count(distinct session_id) filter (where event = 'check' and detail->>'correct' = 'true')      as sitzungen_geloest,
  round(100.0 * count(distinct session_id) filter (where event = 'check' and detail->>'correct' = 'true')
        / nullif(count(distinct session_id) filter (where event = 'exercise_open'), 0), 1)       as loesungsquote_prozent,
  round(avg(case when detail->>'attempt' ~ '^\d{1,6}$' then (detail->>'attempt')::int end)
        filter (where event = 'check' and detail->>'correct' = 'true'), 1)                       as versuche_bis_richtig,
  round(avg(case when detail->>'seconds' ~ '^\d{1,7}$' then (detail->>'seconds')::int end)
        filter (where event = 'check' and detail->>'correct' = 'true'))                          as sekunden_bis_richtig,
  count(distinct session_id) filter (where event = 'hints_open')                                 as sitzungen_tipps,
  count(distinct session_id) filter (where event = 'solution_show')                              as sitzungen_loesung_angezeigt,
  max(created_at)                                                                                as zuletzt
from public.events
group by portal, exercise_id
order by portal, exercise_id;

revoke all on public.events_uebersicht from anon, authenticated;
