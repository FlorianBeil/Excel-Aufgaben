/* Excel.Flo – Dialoge, Filter- und Kontextmenüs des Power Query-Editors
 *
 * Nachbauten der Original-Dialoge (Spalte nach Trennzeichen teilen, Werte
 * ersetzen, Zeilen filtern, Gruppieren nach, Spalten auswählen, …) sowie der
 * Menüs am Filterpfeil, an Spaltenköpfen, Zellen und angewandten Schritten.
 * Jeder Dialog kann einen neuen Schritt anlegen oder – über das Zahnrad –
 * einen bestehenden Schritt bearbeiten (editIndex).
 */

(function () {
  "use strict";

  const UI = window.ExcelFloUI;
  const PQ = window.ExcelFloPowerQuery;
  const { h } = UI;
  const ED = () => window.ExcelFloPQEditor;

  function guard(app, action) {
    if (ED().isAllowed(app, action)) return true;
    ED().hint(app, "„" + (ED().ACTION_LABELS[action] || action) + "“ wird in dieser Übung nicht benötigt.");
    return false;
  }

  function baseTable(app, editIndex) {
    return editIndex !== undefined && editIndex !== null ? ED().tableAt(app, editIndex - 1) : ED().viewTable(app);
  }

  function literalType(type) {
    return type === "Zahl" || type === "Dezimalzahl" ? "number" : type === "Datum" ? "date" : "text";
  }

  function distinctValues(table, col) {
    const seen = new Map();
    table.rows.forEach((r) => {
      const v = r[col];
      if (PQ.isError(v)) return;
      const k = PQ.valueKey(v);
      if (!seen.has(k)) seen.set(k, v);
    });
    return Array.from(seen.values()).sort(PQ.compareValues);
  }

  function valueLabel(v) {
    if (v === null || v === undefined) return "(null)";
    if (v === "") return "(leer)";
    return PQ.displayText(v);
  }

  function advancedToggle(label, content) {
    content.hidden = true;
    const btn = h("button", { type: "button", class: "xl-advtoggle", "aria-expanded": "false" }, [h("span", { class: "xl-advtoggle__arrow", text: "▸" }), label]);
    btn.addEventListener("click", () => {
      content.hidden = !content.hidden;
      btn.setAttribute("aria-expanded", String(!content.hidden));
      btn.firstChild.textContent = content.hidden ? "▸" : "▾";
    });
    return h("div", { class: "xl-adv" }, [btn, content]);
  }

  /* ---------------- Allgemein ---------------- */

  function confirm(app, o, onOk) {
    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: o.title,
      subtitle: o.text,
      width: 460,
      buttons: [
        { label: o.okLabel || "OK", primary: true, onClick: () => onOk() },
        { label: "Abbrechen" },
      ],
    });
  }

  function keepChanges(app) {
    const ed = app.editor;
    if (!app.query || (!ed.dirty && app.query.loaded)) {
      ED().discardAndClose(app);
      return;
    }
    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Möchten Sie Ihre Änderungen beibehalten?",
      width: 430,
      buttons: [
        { label: "Beibehalten", primary: true, onClick: () => ED().closeAndLoad(app, "table") },
        { label: "Verwerfen", onClick: () => ED().discardAndClose(app) },
        { label: "Abbrechen" },
      ],
    });
  }

  function loadTo(app) {
    const state = { kind: "table", where: "new" };
    const radios = (name, opts) =>
      h("div", { class: "xl-radiogroup" }, opts.map((o) => UI.radio(name, o.label, state[name] === o.value, () => (state[name] = o.value))));
    UI.openDialog(app.host, {
      titlebar: "Daten importieren",
      title: "",
      width: 380,
      body: h("div", { class: "xl-import" }, [
        h("p", { text: "Wählen Sie aus, wie diese Daten in Ihrer Arbeitsmappe angezeigt werden sollen." }),
        radios("kind", [
          { value: "table", label: "Tabelle" },
          { value: "pivot", label: "PivotTable-Bericht" },
          { value: "pivotchart", label: "PivotChart" },
          { value: "connection", label: "Nur Verbindung erstellen" },
        ]),
        h("p", { text: "Wo sollen die Daten eingefügt werden?" }),
        radios("where", [
          { value: "existing", label: "Vorhandenes Arbeitsblatt:" },
          { value: "new", label: "Neues Arbeitsblatt" },
        ]),
        UI.checkbox("Diese Daten dem Datenmodell hinzufügen", false, () => {}, true),
      ]),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            if (state.kind === "pivot" || state.kind === "pivotchart") {
              setError("PivotTables werden in dieser Übung nicht benötigt – wähle „Tabelle“.");
              return false;
            }
            if (state.where === "existing" && state.kind === "table") {
              setError("Lade das Ergebnis für diese Übung auf ein neues Arbeitsblatt.");
              return false;
            }
            ED().closeAndLoad(app, state.kind === "connection" ? "connection" : "table");
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  function advancedEditor(app) {
    const code = PQ.queryToM(app.query.steps, app.tableName);
    UI.openDialog(app.host, {
      titlebar: "Erweiterter Editor",
      title: app.query.name,
      width: 720,
      body: h("div", { class: "xl-advanced-editor" }, [
        h("pre", { class: "xl-code", tabindex: "0", text: code }),
        h("p", { class: "xl-advanced-editor__ok" }, [h("span", { class: "xl-okmark", text: "✓" }), "Es wurden keine Syntaxfehler erkannt."]),
      ]),
      buttons: [{ label: "Fertig", primary: true }, { label: "Abbrechen" }],
    });
  }

  /* ---------------- Spalten & Zeilen ---------------- */

  function chooseColumns(app, editIndex) {
    if (!guard(app, "selectColumns")) return;
    const table = baseTable(app, editIndex);
    if (!table) return;
    const existing = editIndex != null ? app.query.steps[editIndex] : null;
    const checked = new Set(existing ? existing.columns : table.columns);
    const list = h("div", { class: "xl-checklist" });
    let filterText = "";

    function draw() {
      list.replaceChildren();
      const visible = table.columns.filter((c) => c.toLowerCase().indexOf(filterText.toLowerCase()) !== -1);
      const all = UI.checkbox("(Alle Spalten auswählen)", visible.every((c) => checked.has(c)), (v) => {
        visible.forEach((c) => (v ? checked.add(c) : checked.delete(c)));
        draw();
      });
      list.appendChild(all);
      visible.forEach((c) => list.appendChild(UI.checkbox(c, checked.has(c), (v) => (v ? checked.add(c) : checked.delete(c)))));
    }
    draw();

    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Spalten auswählen",
      subtitle: "Wählen Sie die beizubehaltenden Spalten aus.",
      width: 420,
      body: h("div", {}, [UI.input("", (v) => ((filterText = v), draw()), { placeholder: "Spalten durchsuchen", "aria-label": "Spalten durchsuchen" }), list]),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const cols = table.columns.filter((c) => checked.has(c));
            if (!cols.length) {
              setError("Wählen Sie mindestens eine Spalte aus.");
              return false;
            }
            ED().commitStep(app, { action: "selectColumns", columns: cols }, editIndex);
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  const ROW_DIALOGS = {
    removeTopRows: { title: "Oberste Zeilen entfernen", text: "Geben Sie an, wie viele Zeilen von oben entfernt werden sollen." },
    removeBottomRows: { title: "Untere Zeilen entfernen", text: "Geben Sie an, wie viele Zeilen von unten entfernt werden sollen." },
    keepTopRows: { title: "Erste Zeilen beibehalten", text: "Geben Sie an, wie viele Zeilen beibehalten werden sollen." },
  };

  function rowCount(app, action, editIndex) {
    if (!guard(app, action)) return;
    if (!baseTable(app, editIndex)) return;
    const existing = editIndex != null ? app.query.steps[editIndex] : null;
    let value = existing ? String(existing.count) : "";
    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: ROW_DIALOGS[action].title,
      subtitle: ROW_DIALOGS[action].text,
      width: 440,
      body: UI.field("Anzahl von Zeilen", UI.input(value, (v) => (value = v), { placeholder: "Beispiel: 3" })),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const n = Number(value.trim());
            if (!Number.isInteger(n) || n < 1) {
              setError("Geben Sie eine ganze Zahl größer als 0 ein.");
              return false;
            }
            ED().commitStep(app, { action, count: n }, editIndex);
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  /* ---------------- Spalte teilen ---------------- */

  function detectDelimiter(table, col) {
    const texts = table.rows.map((r) => r[col]).filter((v) => typeof v === "string" && v !== "");
    let best = null;
    let bestCount = 0;
    [",", ";", "\t", ":", "=", " "].forEach((d) => {
      const count = texts.filter((t) => t.trim().indexOf(d) !== -1).length;
      if (count > bestCount) {
        best = d;
        bestCount = count;
      }
    });
    return best;
  }

  function maxParts(table, col, delimiter) {
    if (!delimiter) return 2;
    return table.rows.reduce((m, r) => {
      const v = r[col];
      if (v === null || v === undefined || PQ.isError(v)) return m;
      return Math.max(m, PQ.displayText(v).split(delimiter).length);
    }, 1);
  }

  function splitByDelimiter(app, column, editIndex) {
    if (!guard(app, "splitColumn")) return;
    const table = baseTable(app, editIndex);
    if (!table) return;
    const existing = editIndex != null ? app.query.steps[editIndex] : null;
    const known = PQ.DELIMITERS.map((d) => d.value);
    const initial = existing ? existing.delimiter : detectDelimiter(table, column);

    const state = {
      choice: initial && known.indexOf(initial) !== -1 ? initial : "custom",
      custom: initial && known.indexOf(initial) === -1 ? initial : "",
      mode: existing ? existing.mode || "each" : "each",
      countTouched: !!existing,
      count: existing ? existing.newColumnNames.length : maxParts(table, column, initial),
    };
    const delimiter = () => (state.choice === "custom" ? state.custom : state.choice);

    const customInput = UI.input(state.custom, (v) => {
      state.custom = v;
      syncCount();
    }, { "aria-label": "Benutzerdefiniertes Trennzeichen" });
    customInput.hidden = state.choice !== "custom";

    const countInput = UI.input(String(state.count), (v) => {
      state.count = v;
      state.countTouched = true;
    }, { "aria-label": "Anzahl zu teilender Spalten" });

    function syncCount() {
      if (state.countTouched) return;
      state.count = Math.max(2, maxParts(table, column, delimiter()));
      countInput.value = String(state.count);
    }

    const delimiterSelect = UI.select(
      [{ value: "custom", label: "--Benutzerdefiniert--" }].concat(PQ.DELIMITERS),
      state.choice,
      (v) => {
        state.choice = v;
        customInput.hidden = v !== "custom";
        if (v === "custom") customInput.focus();
        syncCount();
      }
    );

    const modeRadios = h("div", { class: "xl-radiogroup" }, [
      UI.radio("pq-split-mode", "Trennzeichen ganz links", state.mode === "left", () => (state.mode = "left")),
      UI.radio("pq-split-mode", "Trennzeichen ganz rechts", state.mode === "right", () => (state.mode = "right")),
      UI.radio("pq-split-mode", "Jedes Vorkommen des Trennzeichens", state.mode === "each", () => (state.mode = "each")),
    ]);

    const advanced = h("div", { class: "xl-adv__content" }, [
      h("p", { class: "xl-field__label", text: "Teilen in" }),
      h("div", { class: "xl-radiogroup xl-radiogroup--row" }, [UI.radio("pq-split-into", "Spalten", true, () => {}), UI.radio("pq-split-into", "Zeilen", false, () => {})]),
      UI.field("Anzahl zu teilender Spalten", countInput),
      UI.field("Anführungszeichen", UI.select([{ value: '"', label: '"' }, { value: "", label: "Kein" }], '"')),
      UI.checkbox("Mit Sonderzeichen teilen", false, () => {}, true),
    ]);

    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Spalte nach Trennzeichen teilen",
      subtitle: "Geben Sie das zum Teilen der Textspalte verwendete Trennzeichen an.",
      width: 520,
      body: h("div", { class: "xl-form" }, [
        UI.field("Wählen Sie das Trennzeichen aus, oder geben Sie es ein", h("div", { class: "xl-inline" }, [delimiterSelect, customInput])),
        h("p", { class: "xl-field__label", text: "Teilen bei" }),
        modeRadios,
        advancedToggle("Erweiterte Optionen", advanced),
      ]),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const d = delimiter();
            if (!d) {
              setError("Geben Sie ein Trennzeichen ein.");
              return false;
            }
            const into = advanced.querySelector('input[name="pq-split-into"]:checked');
            if (into && into.parentNode.textContent === "Zeilen") {
              setError("Das Teilen in Zeilen wird in dieser Übung nicht benötigt.");
              return false;
            }
            let count = state.mode === "each" ? Number(String(state.count).trim()) : 2;
            if (!Number.isInteger(count) || count < 1) {
              setError("Geben Sie für die Anzahl der Spalten eine ganze Zahl größer als 0 ein.");
              return false;
            }
            const names =
              existing && existing.newColumnNames.length === count
                ? existing.newColumnNames.slice()
                : Array.from({ length: count }, (_, i) => column + "." + (i + 1));
            ED().commitStep(app, { action: "splitColumn", column, delimiter: d, mode: state.mode, newColumnNames: names }, editIndex);
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  /* ---------------- Werte ersetzen ---------------- */

  function replaceValues(app, columns, editIndex, prefill) {
    if (!guard(app, "replaceValue")) return;
    const table = baseTable(app, editIndex);
    if (!table) return;
    const existing = editIndex != null ? app.query.steps[editIndex] : null;
    const state = {
      find: existing ? existing.find : prefill || "",
      replace: existing ? existing.replace : "",
      matchEntire: existing ? !!existing.matchEntire : false,
    };
    const isText = columns.every((c) => (table.types[c] || "Text") === "Text");

    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Werte ersetzen",
      subtitle: "Ersetzen Sie in den ausgewählten Spalten einen Wert durch einen anderen.",
      width: 480,
      body: h("div", { class: "xl-form" }, [
        UI.field("Zu suchender Wert", UI.input(state.find, (v) => (state.find = v))),
        UI.field("Ersetzen durch", UI.input(state.replace, (v) => (state.replace = v))),
        isText
          ? advancedToggle(
              "Erweiterte Optionen",
              h("div", { class: "xl-adv__content" }, [
                UI.checkbox("Gesamten Zelleninhalt vergleichen", state.matchEntire, (v) => (state.matchEntire = v)),
                UI.checkbox("Mithilfe von Sonderzeichen ersetzen", false, () => {}, true),
              ])
            )
          : null,
      ]),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            if (state.find === "") {
              setError("Geben Sie einen Wert ein, der ersetzt werden soll.");
              return false;
            }
            ED().commitStep(
              app,
              { action: "replaceValue", columns: columns.slice(), find: state.find, replace: state.replace, matchEntire: isText ? state.matchEntire : true },
              editIndex
            );
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  /* ---------------- Zeilen filtern ---------------- */

  function operatorOptions(type) {
    if (type === "Zahl" || type === "Dezimalzahl") {
      return [
        { value: "=", label: "ist gleich" },
        { value: "≠", label: "ist nicht gleich" },
        { value: ">", label: "ist größer als" },
        { value: "≥", label: "ist größer als oder gleich" },
        { value: "<", label: "ist kleiner als" },
        { value: "≤", label: "ist kleiner als oder gleich" },
      ];
    }
    if (type === "Datum") {
      return [
        { value: "=", label: "ist gleich" },
        { value: "≠", label: "ist nicht gleich" },
        { value: "<", label: "liegt vor" },
        { value: "≤", label: "liegt vor oder am" },
        { value: ">", label: "liegt nach" },
        { value: "≥", label: "liegt nach oder am" },
      ];
    }
    return [
      { value: "=", label: "ist gleich" },
      { value: "≠", label: "ist nicht gleich" },
      { value: "beginnt mit", label: "beginnt mit" },
      { value: "beginnt nicht mit", label: "beginnt nicht mit" },
      { value: "endet mit", label: "endet mit" },
      { value: "endet nicht mit", label: "endet nicht mit" },
      { value: "enthält", label: "enthält" },
      { value: "enthält nicht", label: "enthält nicht" },
    ];
  }

  function filterRowsDialog(app, column, operator, value, editIndex) {
    if (!guard(app, "filterRows")) return;
    const table = baseTable(app, editIndex);
    if (!table) return;
    const type = table.types[column] || "Text";
    const ops = operatorOptions(type);
    const state = { operator: ops.some((o) => o.value === operator) ? operator : ops[0].value, value: value == null ? "" : String(value) };

    const listId = "pq-filter-values-" + Date.now();
    const valueInput = UI.input(state.value, (v) => (state.value = v), { list: listId, placeholder: "Wert eingeben oder auswählen", "aria-label": "Vergleichswert" });
    const datalist = h(
      "datalist",
      { id: listId },
      distinctValues(table, column)
        .filter((v) => !PQ.isEmptyValue(v))
        .map((v) => h("option", { value: PQ.displayText(v) }))
    );

    const disabledRow = h("div", { class: "xl-filter-row is-disabled" }, [
      UI.select(ops, ops[0].value),
      UI.input("", null, { disabled: true, placeholder: "Wert eingeben oder auswählen" }),
    ]);
    disabledRow.querySelectorAll("select, input").forEach((n) => (n.disabled = true));

    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Zeilen filtern",
      subtitle: "Wenden Sie eine oder mehrere Filterbedingungen auf die Zeilen in dieser Tabelle an.",
      width: 560,
      body: h("div", { class: "xl-form" }, [
        h("div", { class: "xl-radiogroup xl-radiogroup--row" }, [
          UI.radio("pq-filter-kind", "Einfach", true, () => {}),
          h("label", { class: "xl-radio is-disabled" }, [h("input", { type: "radio", name: "pq-filter-kind", disabled: true }), h("span", { text: "Erweitert" })]),
        ]),
        h("p", { class: "xl-filter-keep" }, ["Zeilen beibehalten, in denen „", h("strong", { text: column }), "“"]),
        h("div", { class: "xl-filter-row" }, [UI.select(ops, state.operator, (v) => (state.operator = v)), valueInput, datalist]),
        h("div", { class: "xl-radiogroup xl-radiogroup--row is-disabled" }, [
          h("label", { class: "xl-radio" }, [h("input", { type: "radio", disabled: true, checked: true }), h("span", { text: "Und" })]),
          h("label", { class: "xl-radio" }, [h("input", { type: "radio", disabled: true }), h("span", { text: "Oder" })]),
        ]),
        disabledRow,
      ]),
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const step = { action: "filterRows", column, operator: state.operator, value: state.value, literalType: literalType(type) };
            const msg = PQ.ACTIONS.filterRows.validate(step);
            if (msg) {
              setError(msg);
              return false;
            }
            ED().commitStep(app, step, editIndex);
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  /* ---------------- Gruppieren nach ---------------- */

  function groupBy(app, editIndex) {
    if (!guard(app, "groupBy")) return;
    const table = baseTable(app, editIndex);
    if (!table) return;
    const existing = editIndex != null ? app.query.steps[editIndex] : null;
    const firstSel = app.editor.selCols.find((c) => table.columns.indexOf(c) !== -1);
    const numericCols = table.columns.filter((c) => table.types[c] === "Zahl" || table.types[c] === "Dezimalzahl");
    const state = {
      advanced: existing ? existing.groupColumns.length > 1 : false,
      groupColumns: existing ? existing.groupColumns.slice() : [firstSel || table.columns[0]],
      aggFunction: existing ? existing.aggFunction : "Anzahl",
      aggColumn: existing ? existing.aggColumn || numericCols[0] || table.columns[0] : numericCols[0] || table.columns[0],
      newColumnName: existing ? existing.newColumnName : "Anzahl",
      nameTouched: !!existing,
    };
    const colOptions = table.columns.map((c) => ({ value: c, label: c }));
    const aggOptions = Object.keys(PQ.AGGREGATIONS).map((k) => ({ value: k, label: PQ.AGGREGATIONS[k].label })).concat([{ value: "all", label: "Alle Zeilen", disabled: true }]);
    const body = h("div", { class: "xl-form" });

    function draw() {
      const groupRows = (state.advanced ? state.groupColumns : state.groupColumns.slice(0, 1)).map((col, i) =>
        h("div", { class: "xl-inline" }, [
          UI.select(colOptions, col, (v) => (state.groupColumns[i] = v)),
          state.advanced && state.groupColumns.length > 1
            ? h("button", {
                type: "button",
                class: "xl-btn xl-btn--small",
                text: "Entfernen",
                on: {
                  click: () => {
                    state.groupColumns.splice(i, 1);
                    draw();
                  },
                },
              })
            : null,
        ])
      );

      const nameInput = UI.input(state.newColumnName, (v) => {
        state.newColumnName = v;
        state.nameTouched = true;
      });
      const aggSelect = UI.select(aggOptions, state.aggFunction, (v) => {
        if (!state.nameTouched) state.newColumnName = v;
        state.aggFunction = v;
        draw();
      });
      const colSelect = UI.select(colOptions, state.aggColumn, (v) => (state.aggColumn = v));
      colSelect.disabled = state.aggFunction === "Anzahl";

      body.replaceChildren(
        h("div", { class: "xl-radiogroup xl-radiogroup--row" }, [
          UI.radio("pq-group-kind", "Einfach", !state.advanced, () => {
            state.advanced = false;
            draw();
          }),
          UI.radio("pq-group-kind", "Erweitert", state.advanced, () => {
            state.advanced = true;
            draw();
          }),
        ]),
        h("div", { class: "xl-field" }, [h("span", { class: "xl-field__label", text: "Gruppieren nach" })].concat(groupRows)),
        state.advanced
          ? h("button", {
              type: "button",
              class: "xl-btn xl-btn--small",
              text: "Gruppierung hinzufügen",
              on: {
                click: () => {
                  state.groupColumns.push(table.columns.find((c) => state.groupColumns.indexOf(c) === -1) || table.columns[0]);
                  draw();
                },
              },
            })
          : null,
        h("div", { class: "xl-group-agg" }, [UI.field("Neuer Spaltenname", nameInput), UI.field("Vorgang", aggSelect), UI.field("Spalte", colSelect)])
      );
    }
    draw();

    UI.openDialog(app.host, {
      titlebar: "Power Query-Editor",
      title: "Gruppieren nach",
      subtitle: "Geben Sie die Spalte an, nach der gruppiert werden soll, und die gewünschte Ausgabe.",
      width: 600,
      body,
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const groupColumns = (state.advanced ? state.groupColumns : state.groupColumns.slice(0, 1)).filter((c, i, arr) => arr.indexOf(c) === i);
            const step = { action: "groupBy", groupColumns, aggFunction: state.aggFunction, newColumnName: state.newColumnName.trim() };
            if (state.aggFunction !== "Anzahl") step.aggColumn = state.aggColumn;
            const msg = PQ.ACTIONS.groupBy.validate(step);
            if (msg) {
              setError(msg);
              return false;
            }
            ED().commitStep(app, step, editIndex);
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  /* ---------------- Datentyp ---------------- */

  function applyType(app, columns, type) {
    if (!guard(app, "changeType")) return;
    const q = app.query;
    const ed = app.editor;
    const current = q.steps[ed.selectedStep];
    const isLastTypeStep = current && current.action === "changeType" && ed.selectedStep === q.steps.length - 1;

    const addNew = () => ED().insertStep(app, { action: "changeType", transforms: columns.map((c) => ({ column: c, type })) });
    const mergeIntoCurrent = () => {
      const transforms = PQ.typeTransforms(current).map((t) => Object.assign({}, t));
      columns.forEach((c) => {
        const hit = transforms.find((t) => t.column === c);
        if (hit) hit.type = type;
        else transforms.push({ column: c, type });
      });
      current.transforms = transforms;
      delete current.column;
      delete current.type;
      ed.dirty = true;
      ED().render(app);
    };

    if (!isLastTypeStep) {
      addNew();
      return;
    }
    const already = PQ.typeTransforms(current).some((t) => columns.indexOf(t.column) !== -1);
    if (!already) {
      mergeIntoCurrent();
      return;
    }
    UI.openDialog(app.host, {
      titlebar: "Spaltentyp ändern",
      title: "Spaltentyp ändern",
      subtitle:
        "Für die ausgewählte Spalte ist bereits eine Konvertierung vorhanden. Möchten Sie die vorhandene Konvertierung ersetzen oder die neue Konvertierung als separaten Schritt hinzufügen?",
      width: 500,
      buttons: [
        { label: "Aktuelle ersetzen", primary: true, onClick: () => mergeIntoCurrent() },
        { label: "Neuen Schritt hinzufügen", onClick: () => addNew() },
        { label: "Abbrechen" },
      ],
    });
  }

  function typeMenu(app, column, anchor) {
    if (!anchor) return;
    UI.openMenu({ anchor, placement: "below" }, ED().typeMenuItems(app, (type) => applyType(app, [column], type)));
  }

  /* ---------------- Filterpfeil ---------------- */

  function filterMenu(app, column, anchor) {
    const table = ED().viewTable(app);
    if (!table) return;
    const ed = app.editor;
    const type = table.types[column] || "Text";
    const values = distinctValues(table, column);
    const checked = new Set(values.map(PQ.valueKey));
    let search = "";

    const lastFilterIndex = (() => {
      for (let i = ed.selectedStep; i >= 0; i--) {
        const s = app.query.steps[i];
        if (s.action === "filterRows" && s.column === column) return i;
      }
      return -1;
    })();

    const content = h("div", { class: "pq-filtermenu" });
    const row = (label, opts) => {
      const el = h("div", { class: "xl-menu__item" + (opts.disabled ? " is-disabled" : "") + (opts.submenu ? " has-sub" : "") }, [
        h("span", { class: "xl-menu__icon" }, [opts.icon ? UI.icon(opts.icon, 16) : null]),
        h("span", { class: "xl-menu__label", text: label }),
        opts.submenu ? h("span", { class: "xl-menu__arrow", text: "›" }) : null,
      ]);
      if (opts.submenu && !opts.disabled) {
        const openSub = () => UI.openMenu({ anchor: el, placement: "right", level: 1 }, opts.submenu());
        el.addEventListener("mouseenter", openSub);
        el.addEventListener("click", openSub);
      } else {
        el.addEventListener("mouseenter", () => UI.closeMenus(1));
        if (!opts.disabled) {
          el.addEventListener("click", () => {
            UI.closeMenus(0);
            opts.onClick();
          });
        }
      }
      return el;
    };

    const c = ED().cmd(app);
    const filterKind = type === "Zahl" || type === "Dezimalzahl" ? "Zahlenfilter" : type === "Datum" ? "Datumsfilter" : "Textfilter";
    const subItems = () =>
      operatorOptions(type).map((o) => ({
        label: o.label.charAt(0).toUpperCase() + o.label.slice(1) + " …",
        onClick: () => filterRowsDialog(app, column, o.value, ""),
      }));

    const list = h("div", { class: "pq-filtermenu__list" });
    const okBtn = h("button", { type: "button", class: "xl-btn xl-btn--primary", text: "OK" });

    function drawList() {
      list.replaceChildren();
      const visible = values.filter((v) => valueLabel(v).toLowerCase().indexOf(search.toLowerCase()) !== -1);
      list.appendChild(
        UI.checkbox("(Alle auswählen)", visible.every((v) => checked.has(PQ.valueKey(v))), (on) => {
          visible.forEach((v) => (on ? checked.add(PQ.valueKey(v)) : checked.delete(PQ.valueKey(v))));
          drawList();
        })
      );
      visible.forEach((v) => {
        const k = PQ.valueKey(v);
        list.appendChild(
          UI.checkbox(valueLabel(v), checked.has(k), (on) => {
            if (on) checked.add(k);
            else checked.delete(k);
            okBtn.disabled = checked.size === 0;
            drawList();
          })
        );
      });
      okBtn.disabled = checked.size === 0;
    }
    drawList();

    okBtn.addEventListener("click", () => {
      UI.closeMenus(0);
      const chosen = values.filter((v) => checked.has(PQ.valueKey(v)));
      const rest = values.filter((v) => !checked.has(PQ.valueKey(v)));
      if (!rest.length) return;
      const plain = (v) => typeof v === "number" || (typeof v === "string" && v !== "") || PQ.isDate(v);
      let step;
      if (rest.length === 1 && plain(rest[0])) {
        step = { operator: "≠", value: PQ.displayText(rest[0]) };
      } else if (chosen.length === 1 && plain(chosen[0])) {
        step = { operator: "=", value: PQ.displayText(chosen[0]) };
      } else if (chosen.length <= rest.length) {
        step = { operator: "in", values: chosen };
      } else {
        step = { operator: "nicht in", values: rest };
      }
      ED().insertStep(app, Object.assign({ action: "filterRows", column, literalType: literalType(type) }, step));
    });

    const cancelBtn = h("button", { type: "button", class: "xl-btn", text: "Abbrechen", on: { click: () => UI.closeMenus(0) } });

    content.append(
      row("Aufsteigend sortieren", { icon: "sortAsc", onClick: () => ((ed.selCols = [column]), c.sort("asc")) }),
      row("Absteigend sortieren", { icon: "sortDesc", onClick: () => ((ed.selCols = [column]), c.sort("desc")) }),
      row("Sortierung löschen", { disabled: true }),
      h("div", { class: "xl-menu__sep" }),
      row("Filter löschen", {
        icon: "clearFilter",
        disabled: lastFilterIndex === -1,
        onClick: () => ED().deleteStep(app, lastFilterIndex),
      }),
      row("Leere entfernen", {
        onClick: () => ED().insertStep(app, { action: "filterRows", column, operator: "≠", value: "", literalType: literalType(type) }),
      }),
      row(filterKind, { submenu: subItems }),
      h("div", { class: "pq-filtermenu__search" }, [
        UI.input("", (v) => {
          search = v;
          drawList();
        }, { placeholder: "Suchen", "aria-label": "Werte durchsuchen" }),
      ]),
      list,
      h("div", { class: "pq-filtermenu__buttons" }, [okBtn, cancelBtn])
    );

    UI.openMenu({ anchor, placement: "below", content, className: "xl-menu--filter" });
  }

  /* ---------------- Kontextmenüs ---------------- */

  function point(x, y) {
    return { left: x, top: y, right: x, bottom: y };
  }

  function headerMenu(app, column, x, y) {
    const c = ED().cmd(app);
    const nn = (label) => ED().notNeeded(app, label);
    const multi = app.editor.selCols.length > 1;
    UI.openMenu({ rect: point(x, y), placement: "point" }, [
      { label: "Entfernen", icon: "removeColumns", onClick: c.removeColumns },
      { label: "Andere Spalten entfernen", onClick: c.removeOtherColumns },
      { label: "Spalte duplizieren", disabled: multi, onClick: c.duplicate },
      { label: "Spalte aus Beispielen hinzufügen …", onClick: nn("Spalte aus Beispielen hinzufügen") },
      { separator: true },
      { label: "Duplikate entfernen", onClick: c.removeDuplicates },
      { label: "Fehler entfernen", onClick: c.removeErrors },
      { separator: true },
      { label: "Typ ändern", submenu: ED().typeMenuItems(app, c.setType) },
      {
        label: "Transformieren",
        submenu: [
          { label: "Kleinbuchstaben", onClick: () => c.format("lower") },
          { label: "Großbuchstaben", onClick: () => c.format("upper") },
          { label: "Jedes Wort großschreiben", onClick: () => c.format("proper") },
          { label: "Kürzen", onClick: () => c.format("trim") },
          { label: "Säubern", onClick: () => c.format("clean") },
          { label: "Länge", onClick: nn("Länge") },
          { separator: true },
          { label: "JSON", onClick: nn("JSON") },
          { label: "XML", onClick: nn("XML") },
        ],
      },
      { separator: true },
      { label: "Werte ersetzen …", icon: "replaceValues", onClick: c.replaceValues },
      { label: "Fehler ersetzen …", onClick: nn("Fehler ersetzen") },
      { separator: true },
      { label: "Spalte teilen", disabled: multi, submenu: ED().splitMenu(app) },
      { label: "Gruppieren nach …", icon: "groupBy", onClick: c.groupBy },
      { label: "Ausfüllen", submenu: [{ label: "Nach unten", onClick: nn("Ausfüllen") }, { label: "Nach oben", onClick: nn("Ausfüllen") }] },
      { label: "Spalten entpivotieren", onClick: nn("Spalten entpivotieren") },
      { label: "Andere Spalten entpivotieren", onClick: nn("Andere Spalten entpivotieren") },
      { label: "Nur ausgewählte Spalten entpivotieren", onClick: nn("Nur ausgewählte Spalten entpivotieren") },
      { separator: true },
      { label: "Umbenennen …", icon: "rename", disabled: multi, onClick: c.rename },
      { label: "Verschieben", submenu: [{ label: "Nach links", onClick: nn("Verschieben") }, { label: "Nach rechts", onClick: nn("Verschieben") }, { label: "An den Anfang", onClick: nn("Verschieben") }, { label: "An das Ende", onClick: nn("Verschieben") }] },
      { separator: true },
      { label: "Drilldown ausführen", disabled: multi, onClick: nn("Drilldown ausführen") },
      { label: "Als neue Abfrage hinzufügen", disabled: multi, onClick: nn("Als neue Abfrage hinzufügen") },
    ]);
  }

  function cellMenu(app, column, value, x, y) {
    const table = ED().viewTable(app);
    const type = table ? table.types[column] || "Text" : "Text";
    const nn = (label) => ED().notNeeded(app, label);
    const quick = (operator) => () => {
      const step = { action: "filterRows", column, literalType: literalType(type) };
      if (PQ.isEmptyValue(value)) Object.assign(step, { operator: operator === "=" ? "in" : "nicht in", values: [value === undefined ? null : value] });
      else Object.assign(step, { operator, value: PQ.displayText(value) });
      ED().insertStep(app, step);
    };
    const isNumeric = type === "Zahl" || type === "Dezimalzahl" || type === "Datum";
    const empty = PQ.isEmptyValue(value) || PQ.isError(value);
    const filterItems = isNumeric
      ? [
          { label: "Ist gleich", onClick: quick("=") },
          { label: "Ist nicht gleich", onClick: quick("≠") },
          { label: type === "Datum" ? "Liegt vor" : "Kleiner als", disabled: empty, onClick: quick("<") },
          { label: type === "Datum" ? "Liegt vor oder am" : "Kleiner als oder gleich", disabled: empty, onClick: quick("≤") },
          { label: type === "Datum" ? "Liegt nach" : "Größer als", disabled: empty, onClick: quick(">") },
          { label: type === "Datum" ? "Liegt nach oder am" : "Größer als oder gleich", disabled: empty, onClick: quick("≥") },
        ]
      : [
          { label: "Ist gleich", onClick: quick("=") },
          { label: "Ist nicht gleich", onClick: quick("≠") },
          { label: "Beginnt mit", disabled: empty, onClick: quick("beginnt mit") },
          { label: "Endet mit", disabled: empty, onClick: quick("endet mit") },
          { label: "Enthält", disabled: empty, onClick: quick("enthält") },
          { label: "Enthält nicht", disabled: empty, onClick: quick("enthält nicht") },
        ];

    UI.openMenu({ rect: point(x, y), placement: "point" }, [
      { label: "Kopieren", onClick: nn("Kopieren") },
      { separator: true },
      { label: isNumeric ? (type === "Datum" ? "Datumsfilter" : "Zahlenfilter") : "Textfilter", submenu: filterItems },
      { separator: true },
      { label: "Werte ersetzen …", icon: "replaceValues", onClick: () => replaceValues(app, [column], null, typeof value === "string" ? value : PQ.isEmptyValue(value) ? "" : PQ.displayText(value)) },
      { separator: true },
      { label: "Drilldown ausführen", onClick: nn("Drilldown ausführen") },
      { label: "Als neue Abfrage hinzufügen", onClick: nn("Als neue Abfrage hinzufügen") },
    ]);
  }

  function tableMenu(app, anchor) {
    const c = ED().cmd(app);
    const nn = (label) => ED().notNeeded(app, label);
    UI.openMenu({ anchor, placement: "below" }, [
      { label: "Erste Zeile als Überschriften verwenden", icon: "useFirstRow", onClick: nn("Erste Zeile als Überschriften verwenden") },
      { separator: true },
      { label: "Benutzerdefinierte Spalte hinzufügen …", onClick: nn("Benutzerdefinierte Spalte hinzufügen") },
      { label: "Spalte aus Beispielen hinzufügen …", onClick: nn("Spalte aus Beispielen hinzufügen") },
      { label: "Bedingte Spalte hinzufügen …", onClick: nn("Bedingte Spalte hinzufügen") },
      { label: "Indexspalte hinzufügen", submenu: [{ label: "Ab 0", onClick: nn("Indexspalte") }, { label: "Ab 1", onClick: nn("Indexspalte") }] },
      { separator: true },
      { label: "Gesamte Tabelle kopieren", onClick: nn("Gesamte Tabelle kopieren") },
      { label: "Spalten auswählen …", icon: "chooseColumns", onClick: c.chooseColumns },
      { separator: true },
      { label: "Duplikate beibehalten", onClick: nn("Duplikate beibehalten") },
      { label: "Fehler beibehalten", onClick: nn("Fehler beibehalten") },
      { label: "Duplikate entfernen", onClick: () => ED().insertStep(app, { action: "removeDuplicates", columns: [] }) },
      { label: "Fehler entfernen", onClick: () => ED().insertStep(app, { action: "removeErrors", columns: [] }) },
      { separator: true },
      { label: "Abfragen zusammenführen …", onClick: nn("Abfragen zusammenführen") },
      { label: "Abfragen anfügen …", onClick: nn("Abfragen anfügen") },
    ]);
  }

  function canEditStep(step) {
    if (["splitColumn", "replaceValue", "groupBy", "removeTopRows", "removeBottomRows", "keepTopRows", "selectColumns"].indexOf(step.action) !== -1) return true;
    return step.action === "filterRows" && step.operator !== "in" && step.operator !== "nicht in";
  }

  function stepMenu(app, index, x, y) {
    const step = app.query.steps[index];
    const nn = (label) => ED().notNeeded(app, label);
    UI.openMenu({ rect: point(x, y), placement: "point" }, [
      { label: "Einstellungen bearbeiten", icon: "gearSmall", disabled: !canEditStep(step), onClick: () => editStep(app, index) },
      {
        label: "Umbenennen",
        icon: "rename",
        onClick: () => {
          app.editor.renamingStep = index;
          ED().render(app);
        },
      },
      { label: "Löschen", onClick: () => ED().deleteStep(app, index) },
      {
        label: "Bis zum Ende löschen",
        onClick: () =>
          confirm(
            app,
            { title: "Schritte löschen", text: "Möchten Sie diesen und alle nachfolgenden Schritte wirklich löschen?", okLabel: "Löschen" },
            () => ED().deleteStep(app, index, true)
          ),
      },
      { label: "Schritt danach einfügen", onClick: nn("Schritt danach einfügen") },
      { label: "Nach oben", disabled: true },
      { label: "Nach unten", disabled: true },
      { label: "Vorherige extrahieren", onClick: nn("Vorherige extrahieren") },
      { separator: true },
      { label: "Native Abfrage anzeigen", disabled: true },
      { label: "Diagnose für Schritt", disabled: true },
      { label: "Eigenschaften …", icon: "properties2", onClick: nn("Eigenschaften") },
    ]);
  }

  function editStep(app, index) {
    const step = app.query.steps[index];
    if (!ED().tableAt(app, index - 1)) {
      ED().hint(app, "Dieser Schritt kann nicht bearbeitet werden, weil ein vorheriger Schritt einen Fehler enthält.", "warning");
      return;
    }
    switch (step.action) {
      case "splitColumn":
        return splitByDelimiter(app, step.column, index);
      case "replaceValue":
        return replaceValues(app, PQ.columnsOf(step), index);
      case "filterRows":
        return filterRowsDialog(app, step.column, step.operator, step.value, index);
      case "groupBy":
        return groupBy(app, index);
      case "removeTopRows":
      case "removeBottomRows":
      case "keepTopRows":
        return rowCount(app, step.action, index);
      case "selectColumns":
        return chooseColumns(app, index);
      default:
        ED().hint(app, "Für diesen Schritt gibt es keine Einstellungen.");
    }
  }

  window.ExcelFloPQDialogs = {
    confirm,
    keepChanges,
    loadTo,
    advancedEditor,
    chooseColumns,
    rowCount,
    splitByDelimiter,
    replaceValues,
    filterRowsDialog,
    groupBy,
    applyType,
    typeMenu,
    filterMenu,
    headerMenu,
    cellMenu,
    tableMenu,
    stepMenu,
    editStep,
  };
})();
