-- Excel.Flo – Einstufungstest: Tracking freischalten und Auswertungen anlegen
--
-- Voraussetzung: supabase/events.sql wurde schon einmal ausgeführt (Tabelle public.events).
-- Einmalig im Supabase-Dashboard ausführen:
--   SQL Editor → New query → dieses Skript einfügen → Run
-- Das Skript kann gefahrlos mehrfach ausgeführt werden.
--
-- Ereignisse des Einstufungstests (portal = 'einstufungstest'):
--   test_start       „Los geht's“ geklickt
--   test_resume      abgebrochenen Test fortgesetzt           detail: nummer, beantwortet
--   question_answer  Frage beantwortet (beim Weiterklicken)    exercise_id = Frage-ID
--                    detail: nummer, typ, kapitel, antwort, richtig, weiss_nicht (Abschlussfrage: abschlussfrage=true)
--   test_abort       Seite verlassen, während eine Frage offen war   exercise_id = offene Frage
--   test_complete    Ergebnis angezeigt                        detail: punkte, startkapitel, antwortmuster, schwerpunkt …
--   chapter_click    Kapitel-Button / „Alle Kapitel“           detail: kapitel, punkte
--   test_restart     „Test neu starten“ bestätigt              detail: vorher

-- 1. Neuen Portalwert erlauben (die Tabelle ließ bisher nur funktionen/pivot/powerquery zu)
alter table public.events drop constraint if exists events_portal_check;
alter table public.events add constraint events_portal_check
  check (portal in ('funktionen', 'pivot', 'powerquery', 'einstufungstest'));

-- 2. Überblick: eine Zeile mit den wichtigsten Zahlen
-- (drop vorher, weil sich Spaltennamen geändert haben können – z. B. Startkapitel 2 → 1)
drop view if exists public."Auswertung Einstufungstest";
create view public."Auswertung Einstufungstest"
with (security_invoker = true) as
with sitzungen as (
  select
    session_id,
    bool_or(event = 'test_start')                                          as gestartet,
    bool_or(event = 'test_complete')                                       as abgeschlossen,
    bool_or(event = 'test_abort')                                          as abbruch,
    bool_or(event = 'chapter_click')                                       as kapitel_klick,
    (array_agg(detail order by created_at desc) filter (where event = 'test_complete'))[1] as ergebnis
  from public.events
  where portal = 'einstufungstest'
  group by session_id
)
select
  count(*) filter (where gestartet)                                                        as "Tests gestartet",
  count(*) filter (where abgeschlossen)                                                    as "Tests abgeschlossen",
  round(100.0 * count(*) filter (where abgeschlossen)
        / nullif(count(*) filter (where gestartet), 0), 1)                                 as "Abschlussquote in %",
  count(*) filter (where abbruch and not abgeschlossen)                                    as "Abgebrochen (nicht fertig)",
  round(avg(case when ergebnis->>'punkte' ~ '^\d{1,2}$' then (ergebnis->>'punkte')::int end), 1) as "Ø Punkte",
  count(*) filter (where ergebnis->>'startkapitel' = '1')                                  as "Startpunkt Kapitel 1",
  count(*) filter (where ergebnis->>'startkapitel' = '4')                                  as "Startpunkt Kapitel 4",
  count(*) filter (where ergebnis->>'startkapitel' = '7')                                  as "Startpunkt Kapitel 7",
  count(*) filter (where ergebnis->>'antwortmuster' = 'true')                              as "Mit Hinweis Antwortmuster",
  count(*) filter (where ergebnis->>'schwerpunkt' = 'a')                                   as "Schwerpunkt Berichte",
  count(*) filter (where ergebnis->>'schwerpunkt' = 'b')                                   as "Schwerpunkt Finanzen",
  count(*) filter (where ergebnis->>'schwerpunkt' = 'c')                                   as "Schwerpunkt Automatisieren",
  count(*) filter (where ergebnis->>'schwerpunkt' = 'd')                                   as "Schwerpunkt Gemischt",
  count(*) filter (where kapitel_klick)                                                    as "Klick auf Kapitel"
from sitzungen;

revoke all on public."Auswertung Einstufungstest" from anon, authenticated;

-- 3. Pro Frage: zählt je Sitzung nur die zuletzt gegebene Antwort (Zurück + Ändern ist erlaubt)
create or replace view public."Auswertung Einstufungstest Fragen"
with (security_invoker = true) as
with letzte as (
  select distinct on (session_id, exercise_id)
    session_id, exercise_id, detail
  from public.events
  where portal = 'einstufungstest' and event = 'question_answer'
    and coalesce(detail->>'abschlussfrage', 'false') <> 'true'
  order by session_id, exercise_id, created_at desc
)
select
  min(case when detail->>'nummer' ~ '^\d{1,2}$' then (detail->>'nummer')::int end)          as "Nr.",
  exercise_id                                                                                as "Frage",
  min(detail->>'kapitel')                                                                    as "Kapitel",
  count(*)                                                                                   as "Antworten",
  round(100.0 * count(*) filter (where detail->>'richtig' = 'true') / nullif(count(*), 0), 1) as "Richtig in %",
  count(*) filter (where detail->>'richtig' = 'false' and detail->>'weiss_nicht' = 'false')  as "Falsch",
  count(*) filter (where detail->>'weiss_nicht' = 'true')                                    as "Weiß ich nicht",
  mode() within group (order by detail->>'antwort')
    filter (where detail->>'richtig' = 'false' and detail->>'weiss_nicht' = 'false')         as "Häufigste falsche Antwort"
from letzte
group by exercise_id
order by 1;

revoke all on public."Auswertung Einstufungstest Fragen" from anon, authenticated;

-- 4. Wo wird abgebrochen? Nur Sitzungen, die den Test nie abgeschlossen haben, letzte offene Frage
create or replace view public."Auswertung Einstufungstest Abbrüche"
with (security_invoker = true) as
with abbrueche as (
  select distinct on (session_id)
    session_id, exercise_id, detail
  from public.events e
  where portal = 'einstufungstest' and event = 'test_abort'
    and not exists (
      select 1 from public.events c
      where c.portal = 'einstufungstest' and c.event = 'test_complete' and c.session_id = e.session_id
    )
  order by session_id, created_at desc
)
select
  min(case when detail->>'nummer' ~ '^\d{1,2}$' then (detail->>'nummer')::int end) as "Bei Frage Nr.",
  exercise_id                                                                       as "Frage",
  count(*)                                                                          as "Abbrüche"
from abbrueche
group by exercise_id
order by 1;

revoke all on public."Auswertung Einstufungstest Abbrüche" from anon, authenticated;
