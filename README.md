# Excel.Flo – Übungsportal

Interaktive Excel-Übungen für den Excel Master Kurs. Kein Framework, kein
Build-Schritt — reines HTML/CSS/JS, gehostet auf GitHub Pages. Jede Übung ist
ein Datensatz (JSON), keine eigene Seite/Datei — die generische Portal-Engine
(`assets/engine.js`) rendert daraus Tabelle, Eingabe und Prüfung.

## Struktur

```
/index.html                → leitet auf /portal/ weiter
/portal/index.html         → Übersicht mit Level-Tabs, generiert aus manifest.json
/portal/uebung.html        → generische Übungsseite, lädt Übung anhand ?id=
/assets/engine.js          → Rendering, Navigation, Prüf-Logik
/assets/engine.css         → Design-System
/assets/formula-engine.js  → eigener, eng begrenzter Formel-Auswerter
/assets/progress.js        → Fortschritts-Speicherung (aktuell localStorage)
/assets/iframe-resize.js   → meldet Seitenhöhe per postMessage (Ablefy-Einbettung)
/assets/exercises/manifest.json     → Liste aller Übungen für die Übersicht
/assets/exercises/<id>.json         → ein Datensatz pro Übung
```

Jede Übung hat dadurch automatisch eine stabile, verlinkbare URL:
`/portal/uebung.html?id=<id>` — für Deep-Links am Ende einzelner Kapitel im
Kurs, während der zentrale Menüpunkt auf `/portal/` (Übersicht) zeigt.

## Lokal testen

Da die Seiten `fetch()` für die JSON-Dateien nutzen, funktioniert das direkte
Öffnen per `file://` in manchen Browsern nicht (CORS). Stattdessen einen
einfachen lokalen Server starten, z. B.:

```bash
python -m http.server 8080
```

und dann `http://localhost:8080` öffnen.

## Hosting auf GitHub Pages

Repo-Root wie hier vorliegend pushen, dann in den Repo-Einstellungen unter
**Pages** den Branch (z. B. `main`) und Ordner `/ (root)` als Quelle wählen.
`index.html` im Root leitet automatisch auf `/portal/` weiter.

## Neue Übung hinzufügen

Neue Übung = neuer Datensatz, kein neuer Seiten-Code.

### 1. Übungsdaten anlegen

Neue Datei unter `assets/exercises/<id>.json`. Aufbau:

```json
{
  "id": "eindeutige-id",
  "title": "Titel der Übung",
  "level": "anfaenger",
  "category": "verweisfunktionen",
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
          "acceptedFormulas": ["=SUMME(A2:B2)"]
        }
      }
    }
  },
  "hints": ["Tipp 1", "Tipp 2"],
  "explanation": "Wird nach erfolgreichem Lösen angezeigt sowie im Tipps-Panel.",
  "solution": "=SUMME(A2:B2)"
}
```

- `level`: `"anfaenger"` | `"fortgeschritten"` | `"profi"` — bestimmt den Tab
  auf der Übersicht. Profi = fortgeschrittene Formeln (verschachtelte
  Funktionen, INDEX/VERGLEICH, dynamische Arrays) — **nicht** Power Query/
  Power Pivot/DAX, das gehört zum separaten Power Kurs.
- `category`: freie Taxonomie zur Gruppierung (`verweisfunktionen`,
  `textfunktionen`, …), wird auch als Badge angezeigt.

**Zell-Typen** (`grid.cells["<Zellbezug>"]`):

- Kein Eintrag → leere Gridzelle.
- `{ "value": ... }` → normale Datenzelle. Optional `"type": "header"` für
  fette Kopfzeilen-Optik, `"format": "currency"` für Euro-Formatierung.
- `{ "type": "input", "answer": {...} }` → Eingabefeld, das geprüft wird. Eine
  Übung kann mehrere Eingabezellen haben; „abgeschlossen" heißt: alle davon
  korrekt.

**Antworten prüfen** (`answer`) — mehrstufig, in dieser Reihenfolge:

1. `value` + `tolerance`: direkte Zahlenwert-Eingabe statt Formel (Komma oder
   Punkt als Dezimaltrennzeichen).
2. `acceptedFormulas`: Array **im Klartext** geschriebener Formeln (kein
   Regex nötig). Vergleich erfolgt normalisiert — Groß-/Kleinschreibung,
   Leerzeichen, `;` vs. `,`, `$`-Fixierung und `WAHR`/`FALSCH` vs. `1`/`0`
   sind dabei egal. Eine Formel wie `=SVERWEIS(D2;A2:B7;2;FALSCH)` reicht
   als einziger Eintrag, um auch `=sverweis(d2,$a$2:$b$7,2,0)` zu akzeptieren.
3. `patterns`: Regex-Array, nur für Sonderfälle, die sich mit (2) nicht
   abbilden lassen.
4. **Formel-Auswertung** (`assets/formula-engine.js`): wenn keine der obigen
   Stufen greift, wird die eingetippte Formel tatsächlich gegen die
   Zelldaten der Übung ausgewertet und das Ergebnis mit `value` verglichen.
   Fängt Formulierungen ab, an die beim Schreiben der Übung niemand gedacht
   hat. Unterstützte Funktionen: siehe `FUNCTION_SIGNATURES` in
   `assets/engine.js` (SVERWEIS, WVERWEIS, WENN, SUMME, SUMMEWENN,
   ZÄHLENWENN, RUNDEN, VERGLEICH, INDEX, MIN, MAX, MITTELWERT). Bewusst kein
   HyperFormula o. ä. — Lizenz (AGPL/kommerziell) passt nicht zu einem
   bezahlten Kursprodukt; stattdessen ein kleiner eigener, isoliert
   getesteter Parser/Evaluator (Test: `node` gegen `formula-engine.js` mit
   einer lokalen Testdatei, siehe Git-Historie für ein Beispiel).

Eine Zelle ist korrekt, sobald **eine** der vier Stufen zutrifft.

### 2. Im Manifest eintragen

In `assets/exercises/manifest.json` einen Eintrag ergänzen (gleiche Felder
wie oben: `id`, `title`, `level`, `category`, `difficulty`, `description`).
Reihenfolge im Array bestimmt die Reihenfolge auf der Übersichtsseite
innerhalb einer Stufe. Das war's — `portal/uebung.html` lädt die neue Übung
automatisch über `?id=<id>`, keine weitere Datei nötig.

## Fortschritt

`assets/progress.js` erzeugt beim ersten Besuch automatisch eine anonyme ID
(kein Login, keine Pflicht-Anmeldung) und speichert abgeschlossene Übungen —
aktuell in `localStorage` des Browsers. Die API (`ExcelFloProgress.*`) ist
bewusst stabil gehalten, damit die Anbindung an ein serverseitiges Backend
(geplant: Supabase, für geräteübergreifenden Fortschritt + optionale
E-Mail-Verknüpfung) später nur diese eine Datei ersetzt, ohne den Rest des
Codes anzufassen.

Nach erfolgreichem Abschluss einer Übung: Erklärung wird angezeigt, die
nächste unerledigte Übung derselben Stufe wird vorgeschlagen (kein Zwang),
und bei vollständig abgeschlossener Stufe bzw. allen Stufen erscheint eine
entsprechende Meldung (auf der Übungs- und der Übersichtsseite).

## Einbettung per iframe (Ablefy)

GitHub Pages setzt keine `X-Frame-Options`/CSP-Header, die Embedding
blockieren würden — hier ist nichts zu konfigurieren. `assets/iframe-resize.js`
meldet die Seitenhöhe per `postMessage` an die Elternseite, damit im iframe
kein unnötiger Scrollbalken entsteht (die Elternseite muss die Nachricht
`{ type: "excelflo:resize", height }` selbst auswerten und die iframe-Höhe
setzen). Der Zugriffsschutz läuft vollständig über Ablefy — das Portal prüft
selbst keine Kaufberechtigung.

## Zell-Engine (Navigation, Formeln)

`assets/engine.js` baut pro Übung ein echtes Tabellen-Feeling nach: Zellen sind per
Klick/Pfeiltasten/Tab/Enter navigierbar, Formeln werden farbig nach Bezug
hervorgehoben, es gibt Strg+C/Strg+V (mit automatischer Anpassung relativer
Bezüge), ein Ausfüllkästchen zum Herunterkopieren, F4 zum Zyklisieren von
`$`-Fixierungen und einen Argument-Tooltip beim Tippen von Funktionsnamen.

Für den Tooltip müssen Funktionen mit ihren Argumentnamen in `FUNCTION_SIGNATURES`
(oben in `engine.js`) eingetragen sein — für eine neue Übung mit bisher unbekannter
Funktion dort einfach einen Eintrag ergänzen, z. B.:

```js
INDEX: ["Matrix", "Zeile", "[Spalte]"],
```

Beim Schreiben einer Formel (Zelle beginnt mit `=`) kann man außerdem andere
Zellen anklicken (oder bei gedrückter Maustaste über einen Bereich ziehen), um
deren Bezug an der aktuellen Cursor-Position einzufügen — wie in echtem Excel.
Steht der Cursor am Formel-Ende, geht das auch per Pfeiltasten: eine Pfeiltaste
peilt die Nachbarzelle an, Umschalt+Pfeiltaste erweitert das zu einem
Zellbereich (z. B. `B2:B7`).

## Design-System

Farben, Schriften und Komponenten (Karten, Tabellen-Grid, Badges, Buttons,
Level-Tabs) liegen zentral in `assets/engine.css`. Neue Übungen sollten
ausschließlich vorhandene Klassen nutzen, damit alle Seiten optisch
konsistent bleiben.
