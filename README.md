# Excel.Flo – Übungsportal

Statische Sammlung interaktiver Excel-Übungen. Kein Framework, kein Build-Schritt,
kein Backend, kein API-Key — reines HTML/CSS/JS. Jede Übung hat eine feste Aufgabe
mit fest bekannten richtigen Antworten, die per Mustervergleich geprüft werden.

## Lokal testen

Da die Seiten `fetch()` für die JSON-Dateien nutzen, funktioniert das direkte Öffnen
per `file://` in manchen Browsern nicht (CORS). Stattdessen einen einfachen lokalen
Server starten, z. B.:

```bash
python -m http.server 8080
```

und dann `http://localhost:8080` öffnen.

## Hosting auf GitHub Pages

Repo-Root wie hier vorliegend pushen, dann in den Repo-Einstellungen unter
**Pages** den Branch (z. B. `main`) und Ordner `/ (root)` als Quelle wählen.
`index.html` im Root wird automatisch als Startseite ausgeliefert.

## Neue Übung hinzufügen

Jede Übung besteht aus zwei Teilen: einem Eintrag im Manifest und einer Datendatei,
die die Tabelle, Aufgabe und Lösung beschreibt. Eine dünne HTML-Shell lädt beides
zur Laufzeit über `assets/engine.js`.

### 1. Übungsdaten anlegen

Neue Datei unter `assets/exercises/<slug>.json`. Aufbau:

```json
{
  "slug": "eindeutiger-slug",
  "title": "Titel der Übung",
  "category": "SVERWEIS",
  "difficulty": "Leicht | Mittel | Schwer",
  "description": "Kurzbeschreibung für die Übersichtskarte.",
  "task": {
    "intro": "Einleitender Text zur Aufgabe.",
    "steps": ["Schritt 1 ...", "Schritt 2 ..."]
  },
  "grid": {
    "cols": ["A", "B", "C"],
    "rowCount": 5,
    "cells": {
      "A1": { "value": "Kopfzeile", "type": "header" },
      "A2": { "value": "Textwert" },
      "B2": { "value": 12.5, "format": "currency" },
      "C2": {
        "type": "input",
        "answer": {
          "value": 12.5,
          "tolerance": 0.005,
          "patterns": ["^=?SUMME\\(A2:B2\\)$"]
        }
      }
    }
  },
  "hints": ["Tipp 1", "Tipp 2"],
  "solution": "=SUMME(A2:B2)"
}
```

**Zell-Typen** (`grid.cells["<Zellbezug>"]`):

- Kein Eintrag → leere Gridzelle.
- `{ "value": ... }` → normale Datenzelle. Optional `"type": "header"` für fette
  Kopfzeilen-Optik, `"format": "currency"` für Euro-Formatierung von Zahlen.
- `{ "type": "input", "answer": {...} }` → Eingabefeld, das geprüft wird.

**Antworten prüfen** (`answer`):

- `patterns`: Array von Regex-Strings. Die Nutzereingabe wird vor dem Test
  Leerzeichen-bereinigt und in Großbuchstaben umgewandelt — Patterns entsprechend
  ohne Leerzeichen schreiben. `[;,]` verwenden, um sowohl Semikolon (deutsches
  Excel) als auch Komma zu akzeptieren.
- `value` + `tolerance`: Erlaubt zusätzlich die direkte Eingabe des berechneten
  Zahlenergebnisses (mit Komma oder Punkt als Dezimaltrennzeichen) statt der Formel.
- Eine Zelle ist korrekt, wenn **entweder** ein Pattern **oder** der Zahlenwert
  passt.

### 2. Im Manifest eintragen

In `assets/exercises/manifest.json` einen Eintrag ergänzen (gleiche Felder wie
oben: `slug`, `title`, `category`, `difficulty`, `description`). Reihenfolge im
Array bestimmt die Reihenfolge auf der Übersichtsseite.

### 3. HTML-Shell erstellen

Neue Datei `uebungen/<slug>.html` — einfach `uebungen/sverweis-preisliste.html`
kopieren und in `data-exercise-path` sowie im `<title>` den neuen Slug/Titel
eintragen. Mehr ist nicht nötig, `engine.js` übernimmt Rendering und Prüfung.

## Design-System

Farben, Schriften und Komponenten (Karten, Tabellen-Grid, Badges, Buttons) liegen
zentral in `assets/engine.css`. Neue Übungen sollten ausschließlich vorhandene
Klassen nutzen, damit alle Seiten optisch konsistent bleiben.
