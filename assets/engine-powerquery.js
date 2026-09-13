/* Excel.Flo – Power-Query-Engine (Kern): Aktionen, Schrittausführung, M-Code, Vergleich
 *
 * Reine Datenlogik ohne DOM. Die Oberfläche (Excel-Fenster mit Registerkarte
 * „Daten“ + Power Query-Editor) liegt in engine-powerquery-ui.js. Headless per
 * node testbar über globalThis.ExcelFloPowerQuery.
 *
 * Tabellen-Modell: { columns: [Name, …], types: { Name: Typ }, rows: [{ Name: Wert }] }
 * Zellwerte: Text | Zahl | null | Datum ({ kind: "date", y, m, d }) | Fehler ({ kind: "error", message })
 *
 * Jede Aktion ist eine reine Funktion (Tabelle, Parameter) → neue Tabelle. Ein
 * ungültiger Schritt (z. B. Spalte existiert nicht mehr) wirft einen StepError,
 * der – wie in Power Query – direkt am Schritt angezeigt wird.
 *
 * Geprüft wird ausschließlich das Endergebnis gegen expectedOutput (Spaltennamen,
 * Spaltenreihenfolge, Werte) – NICHT die Reihenfolge oder Anzahl der Schritte.
 */

(function () {
  "use strict";

  const TYPES = ["Text", "Zahl", "Dezimalzahl", "Datum"];
  const TYPE_M = { Text: "type text", Zahl: "Int64.Type", Dezimalzahl: "type number", Datum: "type date" };

  /* ---------------- Werte ---------------- */

  function isDate(v) {
    return v !== null && typeof v === "object" && v.kind === "date";
  }

  function isError(v) {
    return v !== null && typeof v === "object" && v.kind === "error";
  }

  function errorValue(message) {
    return { kind: "error", message };
  }

  function isEmptyValue(v) {
    return v === null || v === undefined || v === "";
  }

  function makeDate(y, m, d) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return { kind: "date", y, m, d };
  }

  function pad2(n) {
    return (n < 10 ? "0" : "") + n;
  }

  function formatDate(v) {
    return pad2(v.d) + "." + pad2(v.m) + "." + v.y;
  }

  function dateOrdinal(v) {
    return v.y * 10000 + v.m * 100 + v.d;
  }

  function formatNumber(n) {
    return n.toLocaleString("de-DE", { maximumFractionDigits: 10, useGrouping: false });
  }

  function displayText(v) {
    if (v === null || v === undefined) return "null";
    if (typeof v === "number") return formatNumber(v);
    if (isDate(v)) return formatDate(v);
    if (isError(v)) return "Error";
    return String(v);
  }

  function valueKind(v) {
    if (v === null || v === undefined) return "null";
    if (typeof v === "number") return "Zahl";
    if (isDate(v)) return "Datum";
    if (isError(v)) return "Fehler";
    return "Text";
  }

  // Texte in Anführungszeichen, damit Leerzeichen am Rand sichtbar werden.
  function describeValue(v) {
    if (typeof v === "string") return "„" + v + "“";
    return displayText(v);
  }

  function quoteList(names) {
    return names.map((n) => "„" + n + "“").join(", ");
  }

  // Typgenauer Schlüssel für Duplikate/Gruppierung/Werteliste (Text „1“ ≠ Zahl 1).
  function valueKey(v) {
    if (v === null || v === undefined) return "null";
    if (typeof v === "number") return "n:" + v;
    if (isDate(v)) return "d:" + dateOrdinal(v);
    if (isError(v)) return "e:" + v.message;
    return "s:" + v;
  }

  // Strenges Zahlenformat wie Power Query mit englischem Gebietsschema:
  // Punkt = Dezimaltrennzeichen, „129,90“ ergibt einen Fehler. Genau deshalb
  // muss in den Übungen erst das Komma ersetzt werden.
  function parseStrictNumber(text) {
    const s = String(text).trim();
    if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) return NaN;
    return Number(s);
  }

  // Tolerant für Vergleichswerte, die in Dialogen eingetippt werden.
  function parseLooseNumber(text) {
    return parseStrictNumber(String(text).replace(",", "."));
  }

  function parseDate(text) {
    const s = String(text).trim();
    let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return makeDate(+m[3], +m[2], +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return makeDate(+m[1], +m[2], +m[3]);
    return null;
  }

  // Power Query rundet bei „Ganze Zahl“ kaufmännisch zur geraden Zahl (2,5 → 2; 3,5 → 4).
  function roundHalfEven(x) {
    const r = Math.round(x);
    return Math.abs(x % 1) === 0.5 && r % 2 !== 0 ? r - 1 : r;
  }

  function convertValue(v, type) {
    if (v === undefined) return null;
    if (v === null || isError(v)) return v;
    if (type === "Text") return typeof v === "string" ? v : displayText(v);
    if (typeof v === "string" && v.trim() === "") return null;

    const shown = describeValue(v);
    if (type === "Dezimalzahl" || type === "Zahl") {
      const n = typeof v === "number" ? v : typeof v === "string" ? parseStrictNumber(v) : NaN;
      if (Number.isNaN(n)) return errorValue(shown + " kann nicht in den Typ „" + type + "“ konvertiert werden.");
      return type === "Zahl" ? roundHalfEven(n) : n;
    }
    if (type === "Datum") {
      if (isDate(v)) return v;
      const d = typeof v === "string" ? parseDate(v) : null;
      return d || errorValue(shown + " kann nicht in den Typ „Datum“ konvertiert werden.");
    }
    return v;
  }

  function compareValues(a, b) {
    const rank = (v) => (v === null || v === undefined ? 0 : typeof v === "number" ? 1 : isDate(v) ? 2 : isError(v) ? 4 : 3);
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 1) return a - b;
    if (ra === 2) return dateOrdinal(a) - dateOrdinal(b);
    if (ra === 3) return String(a).localeCompare(String(b), "de");
    return 0;
  }

  /* ---------------- Tabellen ---------------- */

  function stepError(message) {
    const err = new Error(message);
    err.isStepError = true;
    return err;
  }

  function requireColumn(table, name) {
    if (!name || table.columns.indexOf(name) === -1) {
      throw stepError("Die Spalte „" + (name || "") + "“ wurde in der Tabelle nicht gefunden.");
    }
  }

  function tableFromRows(rows) {
    const columns = [];
    rows.forEach((row) =>
      Object.keys(row).forEach((key) => {
        if (columns.indexOf(key) === -1) columns.push(key);
      })
    );
    const types = {};
    columns.forEach((c) => {
      const values = rows.map((r) => r[c]).filter((v) => v !== null && v !== undefined);
      types[c] = values.length && values.every((v) => typeof v === "number") ? "Dezimalzahl" : "Text";
    });
    return {
      columns,
      types,
      rows: rows.map((r) => {
        const out = {};
        columns.forEach((c) => (out[c] = r[c] === undefined ? null : r[c]));
        return out;
      }),
    };
  }

  function mapColumn(table, column, fn) {
    requireColumn(table, column);
    return {
      columns: table.columns.slice(),
      types: Object.assign({}, table.types),
      rows: table.rows.map((row) => Object.assign({}, row, { [column]: fn(row[column]) })),
    };
  }

  function mapColumns(table, columns, fn) {
    return columns.reduce((t, c) => mapColumn(t, c, fn), table);
  }

  function projectRows(rows, sourceColumns, rename) {
    return rows.map((row) => {
      const out = {};
      sourceColumns.forEach((c) => (out[rename ? rename(c) : c] = row[c]));
      return out;
    });
  }

  function withRows(table, rows) {
    return { columns: table.columns, types: table.types, rows };
  }

  function columnsOf(p) {
    return p.columns || (p.column ? [p.column] : []);
  }

  function requireCount(p) {
    const n = Number(p.count);
    if (!Number.isInteger(n) || n < 1) throw stepError("Die Anzahl der Zeilen muss eine ganze Zahl größer als 0 sein.");
    return n;
  }

  /* ---------------- M-Code ---------------- */

  function mString(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
  }

  function mRef(stepName) {
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(stepName) ? stepName : "#" + mString(stepName);
  }

  function mList(items) {
    return "{" + items.map(mString).join(", ") + "}";
  }

  function mLiteral(value, literalType) {
    if (literalType === "number") {
      const n = parseLooseNumber(value);
      return Number.isNaN(n) ? mString(value) : String(n);
    }
    if (literalType === "date") {
      const d = parseDate(value);
      return d ? "#date(" + d.y + ", " + d.m + ", " + d.d + ")" : mString(value);
    }
    return mString(value);
  }

  /* ---------------- Aktionen ---------------- */

  const DELIMITERS = [
    { value: ":", label: "Doppelpunkt" },
    { value: ",", label: "Komma" },
    { value: "=", label: "Gleichheitszeichen" },
    { value: ";", label: "Semikolon" },
    { value: " ", label: "Leerzeichen" },
    { value: "\t", label: "Tabstopp" },
  ];

  function delimiterLabel(d) {
    const hit = DELIMITERS.find((x) => x.value === d);
    return hit ? hit.label : "„" + d + "“";
  }

  const TEXT_OPERATORS = ["=", "≠", "beginnt mit", "beginnt nicht mit", "endet mit", "endet nicht mit", "enthält", "enthält nicht"];
  const COMPARE_OPERATORS = ["=", "≠", ">", "≥", "<", "≤"];

  // Liefert <0 / 0 / >0 oder null, wenn Zelle und Vergleichswert nicht vergleichbar sind.
  function compareForFilter(cell, raw) {
    if (typeof cell === "number") {
      const n = parseLooseNumber(raw);
      return Number.isNaN(n) ? null : cell - n;
    }
    if (isDate(cell)) {
      const d = parseDate(raw);
      return d ? dateOrdinal(cell) - dateOrdinal(d) : null;
    }
    if (typeof cell === "string") return cell === raw ? 0 : cell.localeCompare(raw, "de");
    return null;
  }

  // Semantik wie der von Power Query erzeugte M-Code: Textvergleiche beachten
  // Groß-/Kleinschreibung, null fällt bei Text-/Größenvergleichen heraus,
  // bleibt aber bei „ist nicht gleich“ erhalten (null <> "Ja" ist wahr).
  function rowMatches(cell, p) {
    const op = p.operator;
    const raw = p.value === undefined || p.value === null ? "" : String(p.value);
    if (isError(cell)) return false;

    if (op === "in" || op === "nicht in") {
      const keys = (p.values || []).map(valueKey);
      const hit = keys.indexOf(valueKey(cell)) !== -1;
      return op === "in" ? hit : !hit;
    }

    const textOps = {
      "enthält": (t) => t.indexOf(raw) !== -1,
      "beginnt mit": (t) => t.startsWith(raw),
      "endet mit": (t) => t.endsWith(raw),
    };
    const negated = { "enthält nicht": "enthält", "beginnt nicht mit": "beginnt mit", "endet nicht mit": "endet mit" };
    if (textOps[op] || negated[op]) {
      if (cell === null || cell === undefined) return false;
      const result = (textOps[op] || textOps[negated[op]])(displayText(cell));
      return negated[op] ? !result : result;
    }

    if (isEmptyValue(cell)) {
      if (op === "=") return raw === "";
      if (op === "≠") return raw !== "";
      return false;
    }
    if (raw === "") return op === "≠";
    const cmp = compareForFilter(cell, raw);
    if (cmp === null) return op === "≠";
    if (op === "=") return typeof cell === "string" ? cell === raw : cmp === 0;
    if (op === "≠") return typeof cell === "string" ? cell !== raw : cmp !== 0;
    if (op === ">") return cmp > 0;
    if (op === "≥") return cmp >= 0;
    if (op === "<") return cmp < 0;
    if (op === "≤") return cmp <= 0;
    return false;
  }

  function filterCondition(p) {
    const col = "[" + p.column + "]";
    const lit = (v) => (v === null ? "null" : mLiteral(v, p.literalType));
    switch (p.operator) {
      case "=": return "(" + col + " = " + lit(p.value) + ")";
      case "≠": return "(" + col + " <> " + lit(p.value) + ")";
      case ">": return "(" + col + " > " + lit(p.value) + ")";
      case "≥": return "(" + col + " >= " + lit(p.value) + ")";
      case "<": return "(" + col + " < " + lit(p.value) + ")";
      case "≤": return "(" + col + " <= " + lit(p.value) + ")";
      case "enthält": return "Text.Contains(" + col + ", " + mString(p.value) + ")";
      case "enthält nicht": return "not Text.Contains(" + col + ", " + mString(p.value) + ")";
      case "beginnt mit": return "Text.StartsWith(" + col + ", " + mString(p.value) + ")";
      case "beginnt nicht mit": return "not Text.StartsWith(" + col + ", " + mString(p.value) + ")";
      case "endet mit": return "Text.EndsWith(" + col + ", " + mString(p.value) + ")";
      case "endet nicht mit": return "not Text.EndsWith(" + col + ", " + mString(p.value) + ")";
      case "in":
        return "(" + (p.values || []).map((v) => col + " = " + (v === null ? "null" : mString(displayText(v)))).join(" or ") + ")";
      case "nicht in":
        return "(" + (p.values || []).map((v) => col + " <> " + (v === null ? "null" : mString(displayText(v)))).join(" and ") + ")";
      default: return "true";
    }
  }

  const AGGREGATIONS = {
    Anzahl: { label: "Zeilen zählen", m: () => "each Table.RowCount(_), Int64.Type" },
    Summe: { label: "Summe", m: (c) => "each List.Sum([" + c + "]), type nullable number" },
    Durchschnitt: { label: "Durchschnitt", m: (c) => "each List.Average([" + c + "]), type nullable number" },
    Median: { label: "Median", m: (c) => "each List.Median([" + c + "]), type nullable number" },
    Min: { label: "Min", m: (c) => "each List.Min([" + c + "]), type nullable number" },
    Max: { label: "Max", m: (c) => "each List.Max([" + c + "]), type nullable number" },
  };

  function aggregate(values, fn) {
    if (fn === "Anzahl") return values.length;
    const present = values.filter((v) => v !== null && v !== undefined);
    if (present.some((v) => typeof v !== "number")) {
      return errorValue("Nur Zahlen können zusammengefasst werden – ändere zuerst den Datentyp der Spalte.");
    }
    if (!present.length) return null;
    if (fn === "Min") return Math.min.apply(null, present);
    if (fn === "Max") return Math.max.apply(null, present);
    if (fn === "Median") {
      const s = present.slice().sort((a, b) => a - b);
      const mid = Math.floor(s.length / 2);
      return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
    }
    const sum = present.reduce((a, b) => a + b, 0);
    return fn === "Summe" ? sum : sum / present.length;
  }

  const TEXT_MODES = {
    lower: { stepName: "Kleingeschriebener Text", label: "Kleinbuchstaben", m: "Text.Lower", fn: (t) => t.toLowerCase() },
    upper: { stepName: "Großgeschriebener Text", label: "Großbuchstaben", m: "Text.Upper", fn: (t) => t.toUpperCase() },
    proper: {
      stepName: "Großbuchstaben am Wortanfang",
      label: "Jedes Wort großschreiben",
      m: "Text.Proper",
      fn: (t) => t.replace(/\p{L}+/gu, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()),
    },
    clean: { stepName: "Gesäuberter Text", label: "Säubern", m: "Text.Clean", fn: (t) => t.replace(/[ -]/g, "") },
  };

  function typeTransforms(p) {
    return p.transforms || [{ column: p.column, type: p.type }];
  }

  function renamesOf(p) {
    return p.renames || [{ from: p.column, to: p.newName }];
  }

  function sortsOf(p) {
    return p.sorts || [{ column: p.column, direction: p.direction || "asc" }];
  }

  // stepName: Name in „Angewandte Schritte“ (wie im deutschen Power Query).
  // path: Wo der Befehl im Editor zu finden ist (für die Musterlösung).
  // describe: Kurzbeschreibung der Parameter. toM: M-Ausdruck des Schritts.
  const ACTIONS = {
    removeEmptyRows: {
      stepName: "Entfernte leere Zeilen",
      path: "Start › Zeilen entfernen › Leere Zeilen entfernen",
      describe: () => "",
      apply: (table) => withRows(table, table.rows.filter((row) => table.columns.some((c) => !isEmptyValue(row[c])))),
      toM: (p, prev) =>
        "Table.SelectRows(" + prev + ', each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))',
    },

    trimColumn: {
      stepName: "Gekürzter Text",
      path: "Transformieren › Format › Kürzen",
      describe: (p) => columnsOf(p).join(", "),
      apply: (table, p) => mapColumns(table, columnsOf(p), (v) => (typeof v === "string" ? v.trim() : v)),
      toM: (p, prev) =>
        "Table.TransformColumns(" + prev + ",{" + columnsOf(p).map((c) => "{" + mString(c) + ", Text.Trim, type text}").join(", ") + "})",
    },

    transformText: {
      stepName: (p) => (TEXT_MODES[p.mode] || TEXT_MODES.lower).stepName,
      path: "Transformieren › Format",
      describe: (p) => columnsOf(p).join(", ") + " → " + (TEXT_MODES[p.mode] || TEXT_MODES.lower).label,
      apply: (table, p) => {
        const mode = TEXT_MODES[p.mode] || TEXT_MODES.lower;
        return mapColumns(table, columnsOf(p), (v) => (typeof v === "string" ? mode.fn(v) : v));
      },
      toM: (p, prev) =>
        "Table.TransformColumns(" +
        prev +
        ",{" +
        columnsOf(p).map((c) => "{" + mString(c) + ", " + (TEXT_MODES[p.mode] || TEXT_MODES.lower).m + ", type text}").join(", ") +
        "})",
    },

    splitColumn: {
      stepName: "Spalte nach Trennzeichen teilen",
      path: "Start › Spalte teilen › Nach Trennzeichen",
      describe: (p) => p.column + " an " + delimiterLabel(p.delimiter) + " → " + p.newColumnNames.join(", "),
      validate(p) {
        if (!p.delimiter) return "Bitte geben Sie ein Trennzeichen an.";
        if (p.newColumnNames.some((n) => !n)) return "Jede neue Spalte braucht einen Namen.";
        if (new Set(p.newColumnNames).size !== p.newColumnNames.length) return "Die neuen Spaltennamen müssen unterschiedlich sein.";
        return null;
      },
      apply(table, p) {
        requireColumn(table, p.column);
        const names = p.newColumnNames;
        names.forEach((n) => {
          if (n !== p.column && table.columns.indexOf(n) !== -1) throw stepError("Die Spalte „" + n + "“ existiert bereits.");
        });

        const idx = table.columns.indexOf(p.column);
        const columns = table.columns.slice(0, idx).concat(names, table.columns.slice(idx + 1));
        const types = Object.assign({}, table.types);
        delete types[p.column];
        names.forEach((n) => (types[n] = "Text"));

        const rows = table.rows.map((row) => {
          const v = row[p.column];
          let parts;
          if (v === null || v === undefined || isError(v)) {
            parts = names.map(() => (v === undefined ? null : v));
          } else {
            const text = displayText(v);
            let pieces;
            if (p.mode === "left" || p.mode === "right") {
              const at = p.mode === "left" ? text.indexOf(p.delimiter) : text.lastIndexOf(p.delimiter);
              pieces = at === -1 ? [text] : [text.slice(0, at), text.slice(at + p.delimiter.length)];
            } else {
              pieces = text.split(p.delimiter);
              // Mehr Teile als Spalten: Rest landet (mit Trennzeichen) in der letzten Spalte.
              if (pieces.length > names.length) {
                const tail = pieces.slice(names.length - 1).join(p.delimiter);
                pieces.length = names.length - 1;
                pieces.push(tail);
              }
            }
            parts = names.map((_, i) => (i < pieces.length ? pieces[i] : null));
          }
          const out = {};
          columns.forEach((c) => {
            const ni = names.indexOf(c);
            out[c] = ni !== -1 ? parts[ni] : row[c];
          });
          return out;
        });

        return { columns, types, rows };
      },
      toM(p, prev) {
        const splitter =
          p.mode === "left" || p.mode === "right"
            ? "Splitter.SplitTextByEachDelimiter({" + mString(p.delimiter) + "}, QuoteStyle.Csv, " + (p.mode === "right") + ")"
            : "Splitter.SplitTextByDelimiter(" + mString(p.delimiter) + ", QuoteStyle.Csv)";
        return "Table.SplitColumn(" + prev + ", " + mString(p.column) + ", " + splitter + ", " + mList(p.newColumnNames) + ")";
      },
    },

    changeType: {
      stepName: "Geänderter Typ",
      path: "Start › Datentyp",
      describe: (p) => typeTransforms(p).map((t) => t.column + " → " + t.type).join(", "),
      apply(table, p) {
        return typeTransforms(p).reduce((t, tr) => {
          const result = mapColumn(t, tr.column, (v) => convertValue(v, tr.type));
          result.types[tr.column] = tr.type;
          return result;
        }, table);
      },
      toM: (p, prev) =>
        "Table.TransformColumnTypes(" +
        prev +
        ",{" +
        typeTransforms(p).map((t) => "{" + mString(t.column) + ", " + TYPE_M[t.type] + "}").join(", ") +
        "})",
    },

    replaceValue: {
      stepName: "Ersetzter Wert",
      path: "Start › Werte ersetzen",
      describe: (p) => columnsOf(p).join(", ") + ": „" + p.find + "“ → „" + p.replace + "“",
      validate: (p) => (p.find === "" ? "Bitte geben Sie einen zu suchenden Wert ein." : null),
      // Textzellen: Teiltext ersetzen (oder ganze Zelle mit matchEntire).
      // Zahlen/Datumswerte: nur der ganze Wert wird ersetzt – wie Replacer.ReplaceValue.
      apply: (table, p) =>
        mapColumns(table, columnsOf(p), (v) => {
          if (p.find === "" || isError(v) || v === null || v === undefined) return v;
          if (typeof v === "string") {
            if (p.matchEntire) return v === p.find ? p.replace : v;
            return v.split(p.find).join(p.replace);
          }
          return displayText(v) === p.find ? p.replace : v;
        }),
      toM: (p, prev) =>
        "Table.ReplaceValue(" +
        prev +
        "," +
        mString(p.find) +
        "," +
        mString(p.replace) +
        "," +
        (p.matchEntire ? "Replacer.ReplaceValue" : "Replacer.ReplaceText") +
        "," +
        mList(columnsOf(p)) +
        ")",
    },

    filterRows: {
      stepName: "Gefilterte Zeilen",
      path: "Filterpfeil in der Spaltenüberschrift",
      describe(p) {
        if (p.operator === "in" || p.operator === "nicht in") {
          return p.column + (p.operator === "in" ? " = " : " ≠ ") + (p.values || []).map(describeValue).join(" / ");
        }
        return p.column + " " + p.operator + " „" + p.value + "“";
      },
      validate: (p) =>
        p.operator !== "in" && p.operator !== "nicht in" && p.value === "" && p.operator !== "=" && p.operator !== "≠"
          ? "Bitte geben Sie einen Wert ein."
          : null,
      apply(table, p) {
        requireColumn(table, p.column);
        return withRows(table, table.rows.filter((row) => rowMatches(row[p.column], p)));
      },
      toM: (p, prev) => "Table.SelectRows(" + prev + ", each " + filterCondition(p) + ")",
    },

    removeColumn: {
      stepName: "Entfernte Spalten",
      path: "Start › Spalten entfernen",
      describe: (p) => p.columns.join(", "),
      validate: (p) => (!p.columns.length ? "Bitte wählen Sie mindestens eine Spalte aus." : null),
      apply(table, p) {
        p.columns.forEach((c) => requireColumn(table, c));
        const columns = table.columns.filter((c) => p.columns.indexOf(c) === -1);
        const types = {};
        columns.forEach((c) => (types[c] = table.types[c]));
        return { columns, types, rows: projectRows(table.rows, columns) };
      },
      toM: (p, prev) => "Table.RemoveColumns(" + prev + "," + mList(p.columns) + ")",
    },

    selectColumns: {
      stepName: "Andere entfernte Spalten",
      path: "Start › Spalten entfernen › Andere Spalten entfernen",
      describe: (p) => "behalten: " + p.columns.join(", "),
      validate: (p) => (!p.columns.length ? "Bitte wählen Sie mindestens eine Spalte aus." : null),
      apply(table, p) {
        p.columns.forEach((c) => requireColumn(table, c));
        const columns = table.columns.filter((c) => p.columns.indexOf(c) !== -1);
        const types = {};
        columns.forEach((c) => (types[c] = table.types[c]));
        return { columns, types, rows: projectRows(table.rows, columns) };
      },
      toM: (p, prev) => "Table.SelectColumns(" + prev + "," + mList(p.columns) + ")",
    },

    duplicateColumn: {
      stepName: "Duplizierte Spalte",
      path: "Spalte hinzufügen › Spalte duplizieren",
      describe: (p) => p.column + " → " + p.newName,
      apply(table, p) {
        requireColumn(table, p.column);
        if (table.columns.indexOf(p.newName) !== -1) throw stepError("Die Spalte „" + p.newName + "“ existiert bereits.");
        const types = Object.assign({}, table.types, { [p.newName]: table.types[p.column] });
        return {
          columns: table.columns.concat([p.newName]),
          types,
          rows: table.rows.map((r) => Object.assign({}, r, { [p.newName]: r[p.column] })),
        };
      },
      toM: (p, prev) => "Table.DuplicateColumn(" + prev + ", " + mString(p.column) + ", " + mString(p.newName) + ")",
    },

    removeDuplicates: {
      stepName: "Entfernte Duplikate",
      path: "Start › Zeilen entfernen › Duplikate entfernen",
      describe: (p) => (p.columns.length ? "anhand von " + p.columns.join(", ") : "alle Spalten"),
      apply(table, p) {
        p.columns.forEach((c) => requireColumn(table, c));
        const keyColumns = p.columns.length ? p.columns : table.columns;
        const seen = new Set();
        return withRows(
          table,
          table.rows.filter((row) => {
            const key = keyColumns.map((c) => valueKey(row[c])).join("");
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          })
        );
      },
      toM: (p, prev) => (p.columns.length ? "Table.Distinct(" + prev + ", " + mList(p.columns) + ")" : "Table.Distinct(" + prev + ")"),
    },

    removeErrors: {
      stepName: "Entfernte Fehler",
      path: "Start › Zeilen entfernen › Fehler entfernen",
      describe: (p) => (p.columns.length ? p.columns.join(", ") : "alle Spalten"),
      apply(table, p) {
        p.columns.forEach((c) => requireColumn(table, c));
        const cols = p.columns.length ? p.columns : table.columns;
        return withRows(table, table.rows.filter((row) => !cols.some((c) => isError(row[c]))));
      },
      toM: (p, prev) => "Table.RemoveRowsWithErrors(" + prev + (p.columns.length ? ", " + mList(p.columns) : "") + ")",
    },

    removeTopRows: {
      stepName: "Entfernte oberste Zeilen",
      path: "Start › Zeilen entfernen › Oberste Zeilen entfernen",
      describe: (p) => String(p.count),
      apply: (table, p) => withRows(table, table.rows.slice(requireCount(p))),
      toM: (p, prev) => "Table.Skip(" + prev + "," + p.count + ")",
    },

    removeBottomRows: {
      stepName: "Entfernte unterste Zeilen",
      path: "Start › Zeilen entfernen › Untere Zeilen entfernen",
      describe: (p) => String(p.count),
      apply: (table, p) => withRows(table, table.rows.slice(0, Math.max(0, table.rows.length - requireCount(p)))),
      toM: (p, prev) => "Table.RemoveLastN(" + prev + "," + p.count + ")",
    },

    keepTopRows: {
      stepName: "Beibehaltene erste Zeilen",
      path: "Start › Zeilen beibehalten › Erste Zeilen beibehalten",
      describe: (p) => String(p.count),
      apply: (table, p) => withRows(table, table.rows.slice(0, requireCount(p))),
      toM: (p, prev) => "Table.FirstN(" + prev + "," + p.count + ")",
    },

    sortRows: {
      stepName: "Sortierte Zeilen",
      path: "Start › Aufsteigend/Absteigend sortieren",
      describe: (p) => sortsOf(p).map((s) => s.column + (s.direction === "desc" ? " ↓" : " ↑")).join(", "),
      apply(table, p) {
        const sorts = sortsOf(p);
        sorts.forEach((s) => requireColumn(table, s.column));
        const indexed = table.rows.map((row, i) => ({ row, i }));
        indexed.sort((a, b) => {
          for (const s of sorts) {
            const cmp = compareValues(a.row[s.column], b.row[s.column]);
            if (cmp !== 0) return s.direction === "desc" ? -cmp : cmp;
          }
          return a.i - b.i;
        });
        return withRows(table, indexed.map((x) => x.row));
      },
      toM: (p, prev) =>
        "Table.Sort(" +
        prev +
        ",{" +
        sortsOf(p).map((s) => "{" + mString(s.column) + ", " + (s.direction === "desc" ? "Order.Descending" : "Order.Ascending") + "}").join(", ") +
        "})",
    },

    renameColumn: {
      stepName: "Umbenannte Spalten",
      path: "Doppelklick auf die Spaltenüberschrift",
      describe: (p) => renamesOf(p).map((r) => r.from + " → " + r.to).join(", "),
      validate: (p) => (renamesOf(p).some((r) => !r.to) ? "Bitte geben Sie einen Spaltennamen ein." : null),
      apply(table, p) {
        return renamesOf(p).reduce((t, r) => {
          requireColumn(t, r.from);
          if (r.to !== r.from && t.columns.indexOf(r.to) !== -1) throw stepError("Die Spalte „" + r.to + "“ existiert bereits.");
          const rename = (c) => (c === r.from ? r.to : c);
          const types = {};
          t.columns.forEach((c) => (types[rename(c)] = t.types[c]));
          return { columns: t.columns.map(rename), types, rows: projectRows(t.rows, t.columns, rename) };
        }, table);
      },
      toM: (p, prev) =>
        "Table.RenameColumns(" + prev + ",{" + renamesOf(p).map((r) => "{" + mString(r.from) + ", " + mString(r.to) + "}").join(", ") + "})",
    },

    groupBy: {
      stepName: "Gruppierte Zeilen",
      path: "Start › Gruppieren nach",
      describe: (p) =>
        "nach " +
        p.groupColumns.join(", ") +
        ": " +
        (p.aggFunction === "Anzahl" ? "Zeilen zählen" : p.aggFunction + " von " + p.aggColumn) +
        " → " +
        p.newColumnName,
      validate(p) {
        if (!p.groupColumns.length) return "Bitte wählen Sie mindestens eine Spalte zum Gruppieren aus.";
        if (!p.newColumnName) return "Bitte geben Sie einen neuen Spaltennamen ein.";
        if (p.groupColumns.indexOf(p.newColumnName) !== -1) return "Der neue Spaltenname ist bereits vergeben.";
        return null;
      },
      apply(table, p) {
        p.groupColumns.forEach((c) => requireColumn(table, c));
        if (p.aggFunction !== "Anzahl") requireColumn(table, p.aggColumn);
        if (p.groupColumns.indexOf(p.newColumnName) !== -1) {
          throw stepError("Die Spalte „" + p.newColumnName + "“ existiert bereits.");
        }

        const groups = new Map();
        table.rows.forEach((row) => {
          const key = p.groupColumns.map((c) => valueKey(row[c])).join("");
          if (!groups.has(key)) groups.set(key, { first: row, items: [] });
          groups.get(key).items.push(row);
        });

        const columns = p.groupColumns.concat([p.newColumnName]);
        const types = {};
        p.groupColumns.forEach((c) => (types[c] = table.types[c]));
        types[p.newColumnName] = p.aggFunction === "Anzahl" ? "Zahl" : "Dezimalzahl";

        const rows = [];
        groups.forEach((g) => {
          const out = {};
          p.groupColumns.forEach((c) => (out[c] = g.first[c]));
          out[p.newColumnName] = aggregate(g.items.map((r) => r[p.aggColumn]), p.aggFunction);
          rows.push(out);
        });
        return { columns, types, rows };
      },
      toM: (p, prev) =>
        "Table.Group(" +
        prev +
        ", " +
        mList(p.groupColumns) +
        ", {{" +
        mString(p.newColumnName) +
        ", " +
        (AGGREGATIONS[p.aggFunction] || AGGREGATIONS.Anzahl).m(p.aggColumn) +
        "}})",
    },
  };

  function stepNameOf(step) {
    const a = ACTIONS[step.action];
    if (!a) return step.action;
    return typeof a.stepName === "function" ? a.stepName(step) : a.stepName;
  }

  function stepDetail(step) {
    return ACTIONS[step.action] ? ACTIONS[step.action].describe(step) : "";
  }

  // states[k] = Tabelle nach k Schritten (states[0] = Quelle). Bricht beim ersten
  // fehlerhaften Schritt ab – alle folgenden Schritte werden nicht ausgeführt.
  function runSteps(source, steps) {
    const states = [source];
    for (let i = 0; i < steps.length; i++) {
      try {
        const action = ACTIONS[steps[i].action];
        if (!action) throw stepError("Unbekannte Aktion „" + steps[i].action + "“.");
        states.push(action.apply(states[i], steps[i]));
      } catch (err) {
        if (!err.isStepError) console.error("Excel.Flo Power Query: interner Fehler in Schritt", steps[i], err);
        return {
          states,
          errorIndex: i,
          errorMessage: err.isStepError ? err.message : "Dieser Schritt konnte nicht ausgeführt werden.",
        };
      }
    }
    return { states, errorIndex: -1, errorMessage: null };
  }

  function sourceM(tableName) {
    return "Excel.CurrentWorkbook(){[Name=" + mString(tableName) + "]}[Content]";
  }

  // steps: [{ name, action, … }] – name ist der eindeutige Schrittname im Editor.
  function stepToM(steps, index, tableName) {
    if (index < 0) return sourceM(tableName);
    const prev = index === 0 ? "Quelle" : mRef(steps[index - 1].name);
    const action = ACTIONS[steps[index].action];
    return action && action.toM ? action.toM(steps[index], prev) : prev;
  }

  function queryToM(steps, tableName) {
    const lines = ["    Quelle = " + sourceM(tableName)];
    steps.forEach((s, i) => lines.push("    " + mRef(s.name) + " = " + stepToM(steps, i, tableName)));
    return "let\n" + lines.join(",\n") + "\nin\n    " + (steps.length ? mRef(steps[steps.length - 1].name) : "Quelle");
  }

  /* ---------------- Vergleich mit der Zieltabelle ---------------- */

  // expectedOutput ist reines JSON: Datumswerte stehen dort als Text „TT.MM.JJJJ“.
  // Eine in den Typ Datum umgewandelte Spalte gilt daher ebenfalls als korrekt.
  function valuesEqual(actual, expected) {
    if (isError(actual)) return false;
    if (expected === null || expected === undefined) return actual === null || actual === undefined;
    if (typeof expected === "number") return typeof actual === "number" && Math.abs(actual - expected) < 1e-9;
    if (typeof expected === "string") {
      return typeof actual === "string" ? actual === expected : isDate(actual) && formatDate(actual) === expected;
    }
    return false;
  }

  // Gleiche Gleichheit wie valuesEqual, aber als Schlüssel für den Zeilen-Abgleich.
  let errorKeySeq = 0;
  function matchKey(v) {
    if (v === null || v === undefined) return "null";
    if (typeof v === "number") return "n:" + Math.round(v * 1e6) / 1e6;
    if (isDate(v)) return "s:" + formatDate(v);
    if (isError(v)) return "e:" + ++errorKeySeq; // Fehler passen nie
    return "s:" + v;
  }

  function rowKey(row, columns) {
    return columns.map((c) => matchKey(row[c])).join("");
  }

  function rowPreview(row, columns) {
    return columns
      .slice(0, 4)
      .map((c) => c + ": " + (row[c] === "" ? "leer" : displayText(row[c])))
      .join(", ");
  }

  function describePair(actual, expected) {
    const kindA = valueKind(actual);
    const kindE = valueKind(expected);
    const showKinds = kindA !== kindE && kindA !== "null" && kindE !== "null";
    const fmt = (v, kind) => describeValue(v) + (showKinds ? " (" + kind + ")" : "");
    const actualText = isError(actual) ? "ein Fehler: " + actual.message : fmt(actual, kindA);
    return "erwartet " + fmt(expected, kindE) + ", bei dir " + actualText;
  }

  function compareTables(actual, expectedRows, expectedColumns) {
    const expCols = expectedColumns || (expectedRows[0] ? Object.keys(expectedRows[0]) : []);
    const issues = [];
    const wrongCells = new Set(); // "Zeilenindex|Spalte"
    const wrongColumns = new Set();
    const wrongRows = new Set();

    const missing = expCols.filter((c) => actual.columns.indexOf(c) === -1);
    const extra = actual.columns.filter((c) => expCols.indexOf(c) === -1);
    const common = expCols.filter((c) => actual.columns.indexOf(c) !== -1);

    if (missing.length) {
      issues.push((missing.length === 1 ? "Es fehlt die Spalte " : "Es fehlen die Spalten ") + quoteList(missing) + ".");
    }
    if (extra.length) {
      issues.push(
        (extra.length === 1 ? "Die Spalte " + quoteList(extra) + " gehört" : "Die Spalten " + quoteList(extra) + " gehören") +
          " nicht in die Zieltabelle."
      );
      extra.forEach((c) => wrongColumns.add(c));
    }
    if (!missing.length && !extra.length && actual.columns.join("") !== expCols.join("")) {
      issues.push(
        "Die Spalten stimmen, aber ihre Reihenfolge nicht. Erwartet: " +
          expCols.join(", ") +
          " – bei dir: " +
          actual.columns.join(", ") +
          "."
      );
      actual.columns.forEach((c, i) => {
        if (expCols[i] !== c) wrongColumns.add(c);
      });
    }

    const emptyRowCount = actual.rows.filter((r) => actual.columns.every((c) => isEmptyValue(r[c]))).length;

    // Welche Zeilen haben (über die gemeinsamen Spalten) kein Gegenstück in der Zieltabelle?
    const remaining = new Map();
    expectedRows.forEach((r) => {
      const k = rowKey(r, common);
      remaining.set(k, (remaining.get(k) || 0) + 1);
    });
    const unmatched = [];
    actual.rows.forEach((r, i) => {
      const k = rowKey(r, common);
      if (remaining.get(k)) remaining.set(k, remaining.get(k) - 1);
      else unmatched.push(i);
    });

    if (actual.rows.length !== expectedRows.length) {
      issues.push("Deine Tabelle hat " + actual.rows.length + " Zeilen, erwartet sind " + expectedRows.length + ".");
      if (emptyRowCount) {
        issues.push(emptyRowCount === 1 ? "Eine Zeile ist noch komplett leer." : emptyRowCount + " Zeilen sind noch komplett leer.");
      }
      if (actual.rows.length > expectedRows.length && unmatched.length && unmatched.length <= 3) {
        unmatched.forEach((i) => {
          wrongRows.add(i);
          if (actual.columns.every((c) => isEmptyValue(actual.rows[i][c]))) return; // oben schon zusammengefasst
          issues.push("Zeile " + (i + 1) + " gehört nicht in die Zieltabelle (" + rowPreview(actual.rows[i], actual.columns) + ").");
        });
      }
    } else if (common.length) {
      const perColumn = new Map();
      actual.rows.forEach((row, i) =>
        common.forEach((c) => {
          if (valuesEqual(row[c], expectedRows[i][c])) return;
          wrongCells.add(i + "|" + c);
          if (!perColumn.has(c)) perColumn.set(c, []);
          perColumn.get(c).push(i);
        })
      );

      if (perColumn.size && !unmatched.length) {
        issues.push("Alle Zeilen sind vorhanden, stehen aber in einer anderen Reihenfolge als in der Zieltabelle.");
      } else {
        perColumn.forEach((rowIdxs, c) => {
          const i = rowIdxs[0];
          const count = rowIdxs.length === 1 ? "1 Zelle weicht ab" : rowIdxs.length + " Zellen weichen ab";
          const example = rowIdxs.length === 1 ? "Zeile " : "z. B. Zeile ";
          const pair = describePair(actual.rows[i][c], expectedRows[i][c]);
          issues.push("Spalte „" + c + "“: " + count + " – " + example + (i + 1) + ": " + pair + (/[.!?]$/.test(pair) ? "" : "."));
        });
      }
    }

    return { ok: issues.length === 0, issues, wrongCells, wrongColumns, wrongRows };
  }

  (typeof window !== "undefined" ? window : globalThis).ExcelFloPowerQuery = {
    ACTIONS,
    TYPES,
    TEXT_MODES,
    AGGREGATIONS,
    DELIMITERS,
    TEXT_OPERATORS,
    COMPARE_OPERATORS,
    tableFromRows,
    runSteps,
    compareTables,
    displayText,
    describeValue,
    formatDate,
    isDate,
    isError,
    isEmptyValue,
    valueKey,
    compareValues,
    parseLooseNumber,
    parseDate,
    stepNameOf,
    stepDetail,
    stepToM,
    queryToM,
    typeTransforms,
    renamesOf,
    sortsOf,
    columnsOf,
  };
})();
