/* Excel.Flo – Nachbau des Power Query-Editors
 *
 * Fenster mit Registerkarten Datei/Start/Transformieren/Spalte hinzufügen/Ansicht,
 * Abfragen-Bereich, Bearbeitungsleiste (M-Code des gewählten Schritts),
 * Datenvorschau mit Typ-Symbolen und Filterpfeilen, Abfrageeinstellungen mit
 * „Angewandte Schritte“ und grüner Statusleiste. Dialoge, Filter- und
 * Kontextmenüs liegen in engine-powerquery-dialogs.js.
 */

(function () {
  "use strict";

  const UI = window.ExcelFloUI;
  const PQ = window.ExcelFloPowerQuery;
  const { h, icon } = UI;
  const D = () => window.ExcelFloPQDialogs;

  const EDITOR_TABS = ["Datei", "Start", "Transformieren", "Spalte hinzufügen", "Ansicht"];
  const GEAR_ACTIONS = ["splitColumn", "replaceValue", "filterRows", "groupBy", "removeTopRows", "removeBottomRows", "keepTopRows", "selectColumns"];

  const ACTION_LABELS = {
    removeEmptyRows: "Leere Zeilen entfernen",
    trimColumn: "Kürzen",
    transformText: "Format",
    splitColumn: "Spalte teilen",
    changeType: "Datentyp",
    replaceValue: "Werte ersetzen",
    filterRows: "Zeilen filtern",
    removeColumn: "Spalten entfernen",
    selectColumns: "Andere Spalten entfernen",
    duplicateColumn: "Spalte duplizieren",
    removeDuplicates: "Duplikate entfernen",
    removeErrors: "Fehler entfernen",
    removeTopRows: "Oberste Zeilen entfernen",
    removeBottomRows: "Untere Zeilen entfernen",
    keepTopRows: "Erste Zeilen beibehalten",
    sortRows: "Sortieren",
    renameColumn: "Umbenennen",
    groupBy: "Gruppieren nach",
  };

  let stepSeq = 0;

  function clone(x) {
    return JSON.parse(JSON.stringify(x));
  }

  function nowTime() {
    const d = new Date();
    return (d.getHours() < 10 ? "0" : "") + d.getHours() + ":" + (d.getMinutes() < 10 ? "0" : "") + d.getMinutes();
  }

  /* ---------------- Abfrage & Schritte ---------------- */

  function ensureQuery(app) {
    if (app.query) return app.query;
    app.query = {
      name: app.tableName,
      loaded: null,
      steps: [
        {
          id: ++stepSeq,
          name: "Geänderter Typ",
          action: "changeType",
          auto: true,
          transforms: app.source.columns.map((c) => ({ column: c, type: app.source.types[c] })),
        },
      ],
    };
    return app.query;
  }

  function run(app) {
    return PQ.runSteps(app.source, app.query ? app.query.steps : []);
  }

  // Tabelle nach Schritt index (-1 = Quelle) oder null, wenn dort ein Fehler vorliegt.
  function tableAt(app, index) {
    const r = run(app);
    return index + 1 < r.states.length ? r.states[index + 1] : null;
  }

  function viewTable(app) {
    if (!app.query) return null;
    return tableAt(app, app.editor.selectedStep);
  }

  function isAllowed(app, action) {
    const list = app.data.availableActions;
    return !list || !list.length || list.indexOf(action) !== -1;
  }

  function hint(app, text, kind) {
    app.editor.message = { text, kind };
    render(app);
  }

  function uniqueName(app, base) {
    const names = new Set(["Quelle"].concat(app.query.steps.map((s) => s.name)));
    if (!names.has(base)) return base;
    let i = 1;
    while (names.has(base + i)) i++;
    return base + i;
  }

  function selectedColumns(app, single) {
    const table = viewTable(app);
    if (!table) return null;
    const cols = app.editor.selCols.filter((c) => table.columns.indexOf(c) !== -1);
    if (!cols.length) {
      hint(app, "Wählen Sie zuerst eine Spalte aus, indem Sie auf die Spaltenüberschrift klicken.");
      return null;
    }
    return single ? [cols[0]] : cols;
  }

  // Fügt einen Schritt hinter dem ausgewählten ein (wie Power Query). Bei einem
  // Zwischenschritt wird vorher – wie im Original – nachgefragt.
  function insertStep(app, step) {
    if (!isAllowed(app, step.action)) {
      hint(app, "„" + (ACTION_LABELS[step.action] || step.action) + "“ wird in dieser Übung nicht benötigt.");
      return;
    }
    const q = app.query;
    const ed = app.editor;
    const last = q.steps.length - 1;

    const doInsert = () => {
      const current = q.steps[ed.selectedStep];
      // Aufeinanderfolgende Umbenennungen fasst Power Query in einem Schritt zusammen.
      if (step.action === "renameColumn" && current && current.action === "renameColumn" && ed.selectedStep === last) {
        current.renames = PQ.renamesOf(current).concat(PQ.renamesOf(step));
        delete current.column;
        delete current.newName;
      } else {
        step.id = ++stepSeq;
        step.name = uniqueName(app, PQ.stepNameOf(step));
        const at = ed.selectedStep + 1;
        q.steps.splice(at, 0, step);
        ed.selectedStep = at;
        // Nach dem Teilen erkennt Power Query die Typen der neuen Spalten automatisch.
        if (step.action === "splitColumn") {
          const typeStep = {
            id: ++stepSeq,
            action: "changeType",
            auto: true,
            transforms: step.newColumnNames.map((c) => ({ column: c, type: "Text" })),
          };
          typeStep.name = uniqueName(app, "Geänderter Typ");
          q.steps.splice(at + 1, 0, typeStep);
          ed.selectedStep = at + 1;
        }
      }
      ed.dirty = true;
      ed.message = null;
      const table = viewTable(app);
      ed.selCols = table ? ed.selCols.filter((c) => table.columns.indexOf(c) !== -1) : [];
      render(app);
    };

    if (ed.selectedStep < last) {
      D().confirm(
        app,
        {
          title: "Schritt einfügen",
          text: "Möchten Sie diesen Schritt wirklich einfügen? Das Einfügen eines Zwischenschritts kann sich auf nachfolgende Schritte auswirken und dazu führen, dass die Abfrage nicht mehr funktioniert.",
          okLabel: "Einfügen",
        },
        doInsert
      );
    } else {
      doInsert();
    }
  }

  function replaceStep(app, index, step) {
    const old = app.query.steps[index];
    step.id = old.id;
    step.name = old.name;
    if (old.auto) step.auto = true;
    app.query.steps[index] = step;
    app.editor.dirty = true;
    render(app);
  }

  function commitStep(app, step, editIndex) {
    if (editIndex !== undefined && editIndex !== null) replaceStep(app, editIndex, step);
    else insertStep(app, step);
  }

  function deleteStep(app, index, toEnd) {
    const q = app.query;
    const doDelete = () => {
      q.steps.splice(index, toEnd ? q.steps.length - index : 1);
      app.editor.selectedStep = Math.min(index - 1, q.steps.length - 1);
      app.editor.dirty = true;
      render(app);
    };
    if (!toEnd && index < q.steps.length - 1) {
      D().confirm(
        app,
        {
          title: "Schritt löschen",
          text: "Möchten Sie diesen Schritt wirklich löschen? Nachfolgende Schritte, die auf diesen Schritt verweisen, funktionieren möglicherweise nicht mehr.",
          okLabel: "Löschen",
        },
        doDelete
      );
    } else {
      doDelete();
    }
  }

  /* ---------------- Befehle ---------------- */

  function cmd(app) {
    const withCols = (single, fn) => () => {
      const cols = selectedColumns(app, single);
      if (cols) fn(cols);
    };
    return {
      closeLoad: () => closeAndLoad(app, "table"),
      removeColumns: withCols(false, (cols) => insertStep(app, { action: "removeColumn", columns: cols })),
      removeOtherColumns: withCols(false, (cols) => insertStep(app, { action: "selectColumns", columns: cols })),
      chooseColumns: () => D().chooseColumns(app),
      keepTopRows: () => D().rowCount(app, "keepTopRows"),
      removeTopRows: () => D().rowCount(app, "removeTopRows"),
      removeBottomRows: () => D().rowCount(app, "removeBottomRows"),
      removeDuplicates: withCols(false, (cols) => insertStep(app, { action: "removeDuplicates", columns: cols })),
      removeEmptyRows: () => viewTable(app) && insertStep(app, { action: "removeEmptyRows" }),
      removeErrors: withCols(false, (cols) => insertStep(app, { action: "removeErrors", columns: cols })),
      sort: (direction) => withCols(true, (cols) => insertStep(app, { action: "sortRows", sorts: [{ column: cols[0], direction }] }))(),
      split: withCols(true, (cols) => D().splitByDelimiter(app, cols[0])),
      groupBy: () => viewTable(app) && D().groupBy(app),
      replaceValues: withCols(false, (cols) => D().replaceValues(app, cols)),
      format: (mode) =>
        withCols(false, (cols) =>
          insertStep(app, mode === "trim" ? { action: "trimColumn", columns: cols } : { action: "transformText", columns: cols, mode })
        )(),
      rename: withCols(true, (cols) => startRename(app, cols[0])),
      duplicate: withCols(true, (cols) => insertStep(app, { action: "duplicateColumn", column: cols[0], newName: cols[0] + " - Kopie" })),
      setType: (type) => withCols(false, (cols) => D().applyType(app, cols, type))(),
      advancedEditor: () => D().advancedEditor(app),
      refreshPreview: () => {
        app.editor.lastRefresh = nowTime();
        render(app);
      },
    };
  }

  function startRename(app, column) {
    if (!isAllowed(app, "renameColumn")) {
      hint(app, "„Umbenennen“ wird in dieser Übung nicht benötigt.");
      return;
    }
    app.editor.renaming = column;
    render(app);
  }

  function notNeeded(app, label) {
    return () => hint(app, "„" + label.replace("\n", " ") + "“ wird in dieser Übung nicht benötigt.");
  }

  function typeMenuItems(app, onPick) {
    const current = currentType(app);
    const t = (label, type) => ({
      label,
      checked: type && current === type,
      onClick: type ? () => onPick(type) : notNeeded(app, "Datentyp „" + label + "“"),
    });
    return [
      t("Dezimalzahl", "Dezimalzahl"),
      t("Feste Dezimalzahl"),
      t("Ganze Zahl", "Zahl"),
      t("Prozentsatz"),
      t("Datum/Uhrzeit"),
      t("Datum", "Datum"),
      t("Uhrzeit"),
      t("Datum/Uhrzeit/Zeitzone"),
      t("Dauer"),
      t("Text", "Text"),
      t("Wahr/Falsch"),
      t("Binär"),
      { separator: true },
      { label: "Mit Gebietsschema …", onClick: notNeeded(app, "Mit Gebietsschema") },
    ];
  }

  function currentType(app) {
    const table = viewTable(app);
    const col = app.editor.selCols[0];
    if (!table || !col || table.columns.indexOf(col) === -1) return "";
    return table.types[col] || "Text";
  }

  const TYPE_LABELS = { Text: "Text", Zahl: "Ganze Zahl", Dezimalzahl: "Dezimalzahl", Datum: "Datum" };

  /* ---------------- Menüband-Definitionen ---------------- */

  function ribbonGroups(app, tab) {
    const c = cmd(app);
    const nn = (label) => notNeeded(app, label);
    const hasQuery = !!app.query;
    const hasTable = !!viewTable(app);
    const hasCol = hasTable && app.editor.selCols.length > 0;
    const noTable = !hasTable;
    const noCol = !hasCol;

    if (tab === "Start") {
      return [
        {
          label: "Schließen",
          icon: "closeLoad",
          priority: 100,
          controls: [
            {
              type: "big",
              label: "Schließen &\nladen",
              icon: "closeLoad",
              disabled: !hasQuery,
              onClick: c.closeLoad,
              menu: () => [
                { label: "Schließen & laden", onClick: c.closeLoad },
                { label: "Schließen & laden in …", onClick: () => D().loadTo(app) },
              ],
              tip: { title: "Schließen & laden", text: "Schließen Sie den Power Query-Editor und laden Sie die Abfrage in die Arbeitsmappe." },
            },
          ],
        },
        {
          label: "Abfrage",
          icon: "refreshPreview",
          priority: 30,
          controls: [
            {
              type: "big",
              label: "Vorschau\naktualisieren",
              icon: "refreshPreview",
              disabled: !hasQuery,
              onClick: c.refreshPreview,
              menu: () => [
                { label: "Vorschau aktualisieren", onClick: c.refreshPreview },
                { label: "Alle aktualisieren", onClick: c.refreshPreview },
                { label: "Aktualisierung abbrechen", disabled: true },
              ],
            },
            {
              type: "stack",
              items: [
                { label: "Eigenschaften", icon: "properties2", disabled: !hasQuery, onClick: nn("Eigenschaften") },
                { label: "Erweiterter Editor", icon: "advancedEditor", disabled: !hasQuery, onClick: c.advancedEditor },
                { label: "Verwalten", icon: "manage", disabled: !hasQuery, menu: [{ label: "Löschen", onClick: nn("Löschen") }, { label: "Duplizieren", onClick: nn("Duplizieren") }, { label: "Verweisen", onClick: nn("Verweisen") }] },
              ],
            },
          ],
        },
        {
          label: "Spalten verwalten",
          icon: "chooseColumns",
          priority: 70,
          controls: [
            {
              type: "big",
              label: "Spalten\nauswählen",
              icon: "chooseColumns",
              disabled: noTable,
              menu: () => [
                { label: "Spalten auswählen", icon: "chooseColumns", onClick: c.chooseColumns },
                { label: "Gehe zu Spalte", onClick: nn("Gehe zu Spalte") },
              ],
            },
            {
              type: "big",
              label: "Spalten\nentfernen",
              icon: "removeColumns",
              disabled: noCol,
              menu: () => [
                { label: "Spalten entfernen", icon: "removeColumns", onClick: c.removeColumns },
                { label: "Andere Spalten entfernen", onClick: c.removeOtherColumns },
              ],
              tip: { title: "Spalten entfernen", text: "Entfernen Sie die ausgewählten Spalten aus der Tabelle." },
            },
          ],
        },
        {
          label: "Zeilen verringern",
          icon: "removeRows",
          priority: 60,
          controls: [
            {
              type: "big",
              label: "Zeilen\nbeibehalten",
              icon: "keepRows",
              disabled: noTable,
              menu: () => [
                { label: "Erste Zeilen beibehalten", onClick: c.keepTopRows },
                { label: "Letzte Zeilen beibehalten", onClick: nn("Letzte Zeilen beibehalten") },
                { label: "Zeilenbereich beibehalten", onClick: nn("Zeilenbereich beibehalten") },
                { separator: true },
                { label: "Duplikate beibehalten", disabled: noCol, onClick: nn("Duplikate beibehalten") },
                { label: "Fehler beibehalten", disabled: noCol, onClick: nn("Fehler beibehalten") },
              ],
            },
            {
              type: "big",
              label: "Zeilen\nentfernen",
              icon: "removeRows",
              disabled: noTable,
              menu: () => [
                { label: "Oberste Zeilen entfernen", onClick: c.removeTopRows },
                { label: "Untere Zeilen entfernen", onClick: c.removeBottomRows },
                { label: "Alternierende Zeilen entfernen", onClick: nn("Alternierende Zeilen entfernen") },
                { separator: true },
                { label: "Duplikate entfernen", disabled: noCol, onClick: c.removeDuplicates },
                { label: "Leere Zeilen entfernen", onClick: c.removeEmptyRows },
                { label: "Fehler entfernen", disabled: noCol, onClick: c.removeErrors },
              ],
              tip: { title: "Zeilen entfernen", text: "Entfernen Sie Zeilen aus der Tabelle, z. B. leere Zeilen, Duplikate oder Fehler." },
            },
          ],
        },
        {
          label: "Sortieren",
          icon: "sort",
          priority: 40,
          controls: [
            {
              type: "stack",
              iconOnly: true,
              items: [
                { label: "Aufsteigend sortieren", icon: "sortAsc", disabled: noCol, onClick: () => c.sort("asc") },
                { label: "Absteigend sortieren", icon: "sortDesc", disabled: noCol, onClick: () => c.sort("desc") },
              ],
            },
          ],
        },
        {
          label: "Transformieren",
          icon: "splitColumn",
          priority: 80,
          controls: [
            {
              type: "big",
              label: "Spalte\nteilen",
              icon: "splitColumn",
              disabled: noCol,
              menu: () => splitMenu(app),
            },
            { type: "big", label: "Gruppieren\nnach", icon: "groupBy", disabled: noTable, onClick: c.groupBy },
            {
              type: "stack",
              keepLabels: true,
              items: [
                {
                  label: "Datentyp: " + (TYPE_LABELS[currentType(app)] || ""),
                  disabled: noCol,
                  menu: () => typeMenuItems(app, c.setType),
                },
                {
                  label: "Erste Zeile als Überschriften verwenden",
                  icon: "useFirstRow",
                  disabled: noTable,
                  menu: [
                    { label: "Erste Zeile als Überschriften verwenden", onClick: nn("Erste Zeile als Überschriften verwenden") },
                    { label: "Überschriften als erste Zeile verwenden", onClick: nn("Überschriften als erste Zeile verwenden") },
                  ],
                },
                { label: "Werte ersetzen", icon: "replaceValues", disabled: noCol, onClick: c.replaceValues },
              ],
            },
          ],
        },
        {
          label: "Kombinieren",
          icon: "mergeQueries",
          priority: 10,
          controls: [
            {
              type: "stack",
              items: [
                { label: "Abfragen zusammenführen", icon: "mergeQueries", disabled: noTable, menu: [{ label: "Abfragen zusammenführen", onClick: nn("Abfragen zusammenführen") }, { label: "Abfragen als neue Abfrage zusammenführen", onClick: nn("Abfragen zusammenführen") }] },
                { label: "Abfragen anfügen", icon: "appendQueries", disabled: noTable, menu: [{ label: "Abfragen anfügen", onClick: nn("Abfragen anfügen") }, { label: "Abfragen als neue Abfrage anfügen", onClick: nn("Abfragen anfügen") }] },
                { label: "Dateien kombinieren", icon: "combineFiles", disabled: true },
              ],
            },
          ],
        },
        {
          label: "Parameter",
          icon: "manageParameters",
          priority: 15,
          controls: [
            {
              type: "big",
              label: "Parameter\nverwalten",
              icon: "manageParameters",
              menu: [{ label: "Parameter verwalten", onClick: nn("Parameter verwalten") }, { label: "Parameter bearbeiten", onClick: nn("Parameter bearbeiten") }, { label: "Neuer Parameter", onClick: nn("Neuer Parameter") }],
            },
          ],
        },
        {
          label: "Datenquellen",
          icon: "dataSourceSettings",
          priority: 12,
          controls: [{ type: "big", label: "Datenquellen-\neinstellungen", icon: "dataSourceSettings", onClick: nn("Datenquelleneinstellungen") }],
        },
        {
          label: "Neue Abfrage",
          icon: "newSource",
          priority: 20,
          controls: [
            {
              type: "stack",
              items: [
                { label: "Neue Quelle", icon: "newSource", menu: [{ label: "Datei", submenu: [{ label: "Excel-Arbeitsmappe", onClick: nn("Neue Quelle") }, { label: "Text/CSV", onClick: nn("Neue Quelle") }] }, { label: "Datenbank", onClick: nn("Neue Quelle") }, { label: "Andere Quellen", onClick: nn("Neue Quelle") }] },
                { label: "Zuletzt verwendete Quellen", icon: "recentSources", menu: [{ label: "Keine zuletzt verwendeten Quellen", disabled: true }] },
                { label: "Daten eingeben", icon: "enterData", onClick: nn("Daten eingeben") },
              ],
            },
          ],
        },
      ];
    }

    if (tab === "Transformieren") {
      return [
        {
          label: "Tabelle",
          icon: "groupBy",
          priority: 60,
          controls: [
            { type: "big", label: "Gruppieren\nnach", icon: "groupBy", disabled: noTable, onClick: c.groupBy },
            {
              type: "big",
              label: "Erste Zeile als\nÜberschriften",
              icon: "useFirstRowBig",
              disabled: noTable,
              menu: [
                { label: "Erste Zeile als Überschriften verwenden", onClick: nn("Erste Zeile als Überschriften verwenden") },
                { label: "Überschriften als erste Zeile verwenden", onClick: nn("Überschriften als erste Zeile verwenden") },
              ],
            },
            {
              type: "stack",
              items: [
                { label: "Transponieren", icon: "move", disabled: noTable, onClick: nn("Transponieren") },
                { label: "Zeilen umkehren", icon: "fillDown", disabled: noTable, onClick: nn("Zeilen umkehren") },
                { label: "Zeilen zählen", icon: "dataType", disabled: noTable, onClick: nn("Zeilen zählen") },
              ],
            },
          ],
        },
        {
          label: "Beliebige Spalte",
          icon: "detectType",
          priority: 70,
          controls: [
            {
              type: "stack",
              items: [
                { label: "Datentyp: " + (TYPE_LABELS[currentType(app)] || ""), disabled: noCol, menu: () => typeMenuItems(app, c.setType) },
                { label: "Datentyp erkennen", icon: "dataType", disabled: noCol, onClick: nn("Datentyp erkennen") },
                { label: "Umbenennen", icon: "rename", disabled: noCol, onClick: c.rename },
              ],
            },
            {
              type: "stack",
              items: [
                {
                  label: "Werte ersetzen",
                  icon: "replaceValues",
                  disabled: noCol,
                  menu: () => [
                    { label: "Werte ersetzen", onClick: c.replaceValues },
                    { label: "Fehler ersetzen", onClick: nn("Fehler ersetzen") },
                  ],
                },
                { label: "Ausfüllen", icon: "fillDown", disabled: noCol, menu: [{ label: "Nach unten", onClick: nn("Ausfüllen") }, { label: "Nach oben", onClick: nn("Ausfüllen") }] },
                { label: "Spalte pivotieren", icon: "move", disabled: noCol, onClick: nn("Spalte pivotieren") },
              ],
            },
            {
              type: "stack",
              items: [
                { label: "Spalten entpivotieren", icon: "toList", disabled: noCol, menu: [{ label: "Spalten entpivotieren", onClick: nn("Spalten entpivotieren") }, { label: "Andere Spalten entpivotieren", onClick: nn("Andere Spalten entpivotieren") }] },
                { label: "Verschieben", icon: "move", disabled: noCol, menu: [{ label: "Nach links", onClick: nn("Verschieben") }, { label: "Nach rechts", onClick: nn("Verschieben") }, { label: "An den Anfang", onClick: nn("Verschieben") }, { label: "An das Ende", onClick: nn("Verschieben") }] },
                { label: "In Liste konvertieren", icon: "toList", disabled: noCol, onClick: nn("In Liste konvertieren") },
              ],
            },
          ],
        },
        {
          label: "Textspalte",
          icon: "format",
          priority: 80,
          controls: [
            { type: "big", label: "Spalte\nteilen", icon: "splitColumn", disabled: noCol, menu: () => splitMenu(app) },
            {
              type: "big",
              label: "Format",
              icon: "format",
              disabled: noCol,
              menu: () => [
                { label: "Kleinbuchstaben", onClick: () => c.format("lower") },
                { label: "Großbuchstaben", onClick: () => c.format("upper") },
                { label: "Jedes Wort großschreiben", onClick: () => c.format("proper") },
                { label: "Kürzen", onClick: () => c.format("trim") },
                { label: "Säubern", onClick: () => c.format("clean") },
                { label: "Präfix hinzufügen", onClick: nn("Präfix hinzufügen") },
                { label: "Suffix hinzufügen", onClick: nn("Suffix hinzufügen") },
              ],
              tip: { title: "Format", text: "Ändern Sie die Groß-/Kleinschreibung oder entfernen Sie Leerzeichen am Anfang und Ende (Kürzen)." },
            },
            {
              type: "stack",
              items: [
                { label: "Spalten zusammenführen", icon: "mergeQueries", disabled: noCol, onClick: nn("Spalten zusammenführen") },
                { label: "Extrahieren", icon: "extract", disabled: noCol, menu: [{ label: "Länge", onClick: nn("Extrahieren") }, { label: "Erste Zeichen", onClick: nn("Extrahieren") }, { label: "Letzte Zeichen", onClick: nn("Extrahieren") }, { label: "Bereich", onClick: nn("Extrahieren") }, { label: "Text vor Trennzeichen", onClick: nn("Extrahieren") }, { label: "Text nach Trennzeichen", onClick: nn("Extrahieren") }] },
                { label: "Analysieren", icon: "parse", disabled: noCol, menu: [{ label: "XML", onClick: nn("Analysieren") }, { label: "JSON", onClick: nn("Analysieren") }] },
              ],
            },
          ],
        },
        {
          label: "Zahlenspalte",
          icon: "statistics",
          priority: 30,
          controls: [
            { type: "big", label: "Statistiken", icon: "statistics", disabled: noCol, menu: [{ label: "Summe", onClick: nn("Statistiken") }, { label: "Minimum", onClick: nn("Statistiken") }, { label: "Maximum", onClick: nn("Statistiken") }, { label: "Mittelwert", onClick: nn("Statistiken") }] },
            { type: "big", label: "Standard", icon: "standard", disabled: noCol, menu: [{ label: "Addieren", onClick: nn("Standard") }, { label: "Multiplizieren", onClick: nn("Standard") }, { label: "Subtrahieren", onClick: nn("Standard") }, { label: "Dividieren", onClick: nn("Standard") }] },
            { type: "big", label: "Wissen-\nschaftlich", icon: "scientific", disabled: noCol, menu: [{ label: "Absolutwert", onClick: nn("Wissenschaftlich") }, { label: "Potenz", onClick: nn("Wissenschaftlich") }] },
            {
              type: "stack",
              items: [
                { label: "Trigonometrie", icon: "trig", disabled: noCol, menu: [{ label: "Sinus", onClick: nn("Trigonometrie") }] },
                { label: "Runden", icon: "rounding", disabled: noCol, menu: [{ label: "Aufrunden", onClick: nn("Runden") }, { label: "Abrunden", onClick: nn("Runden") }, { label: "Runden …", onClick: nn("Runden") }] },
                { label: "Informationen", icon: "information", disabled: noCol, menu: [{ label: "Ist gerade", onClick: nn("Informationen") }, { label: "Ist ungerade", onClick: nn("Informationen") }] },
              ],
            },
          ],
        },
        {
          label: "Datums- & Uhrzeitspalte",
          icon: "date",
          priority: 25,
          controls: [
            { type: "big", label: "Datum", icon: "date", disabled: noCol, menu: [{ label: "Alter", onClick: nn("Datum") }, { label: "Nur Datum", onClick: nn("Datum") }, { label: "Jahr", onClick: nn("Datum") }, { label: "Monat", onClick: nn("Datum") }] },
            { type: "big", label: "Uhrzeit", icon: "time", disabled: noCol, menu: [{ label: "Nur Uhrzeit", onClick: nn("Uhrzeit") }] },
            { type: "big", label: "Dauer", icon: "duration", disabled: noCol, menu: [{ label: "Tage", onClick: nn("Dauer") }] },
          ],
        },
        {
          label: "Strukturierte Spalte",
          icon: "expand",
          priority: 20,
          controls: [
            { type: "big", label: "Erweitern", icon: "expand", disabled: true },
            { type: "big", label: "Aggregieren", icon: "aggregate", disabled: true },
            { type: "stack", items: [{ label: "Werte extrahieren", icon: "extractValues", disabled: true }] },
          ],
        },
      ];
    }

    if (tab === "Spalte hinzufügen") {
      return [
        {
          label: "Allgemein",
          icon: "customColumn",
          priority: 70,
          controls: [
            { type: "big", label: "Spalte aus\nBeispielen", icon: "columnFromExamples", disabled: noTable, menu: [{ label: "Aus allen Spalten", onClick: nn("Spalte aus Beispielen") }, { label: "Aus Auswahl", onClick: nn("Spalte aus Beispielen") }] },
            { type: "big", label: "Benutzerdefinierte\nSpalte", icon: "customColumn", disabled: noTable, onClick: nn("Benutzerdefinierte Spalte") },
            { type: "big", label: "Benutzerdefinierte\nFunktion aufrufen", icon: "invokeFunction", disabled: noTable, onClick: nn("Benutzerdefinierte Funktion aufrufen") },
            {
              type: "stack",
              items: [
                { label: "Bedingte Spalte", icon: "dataValidation", disabled: noTable, onClick: nn("Bedingte Spalte") },
                { label: "Indexspalte", icon: "dataType", disabled: noTable, menu: [{ label: "Ab 0", onClick: nn("Indexspalte") }, { label: "Ab 1", onClick: nn("Indexspalte") }, { label: "Benutzerdefiniert …", onClick: nn("Indexspalte") }] },
                { label: "Spalte duplizieren", icon: "table", disabled: noCol, onClick: c.duplicate },
              ],
            },
          ],
        },
        {
          label: "Aus Text",
          icon: "format",
          priority: 50,
          controls: [
            { type: "big", label: "Format", icon: "format", disabled: noCol, menu: [{ label: "Kleinbuchstaben", onClick: nn("Format (Spalte hinzufügen)") }, { label: "Großbuchstaben", onClick: nn("Format (Spalte hinzufügen)") }, { label: "Kürzen", onClick: nn("Format (Spalte hinzufügen)") }] },
            {
              type: "stack",
              items: [
                { label: "Spalten zusammenführen", icon: "mergeQueries", disabled: noCol, onClick: nn("Spalten zusammenführen") },
                { label: "Extrahieren", icon: "extract", disabled: noCol, menu: [{ label: "Länge", onClick: nn("Extrahieren") }] },
                { label: "Analysieren", icon: "parse", disabled: noCol, menu: [{ label: "JSON", onClick: nn("Analysieren") }] },
              ],
            },
          ],
        },
        {
          label: "Aus Zahl",
          icon: "statistics",
          priority: 30,
          controls: [
            { type: "big", label: "Statistiken", icon: "statistics", disabled: noCol, menu: [{ label: "Summe", onClick: nn("Statistiken") }] },
            { type: "big", label: "Standard", icon: "standard", disabled: noCol, menu: [{ label: "Addieren", onClick: nn("Standard") }] },
            { type: "big", label: "Wissen-\nschaftlich", icon: "scientific", disabled: noCol, menu: [{ label: "Potenz", onClick: nn("Wissenschaftlich") }] },
            {
              type: "stack",
              items: [
                { label: "Trigonometrie", icon: "trig", disabled: noCol, menu: [{ label: "Sinus", onClick: nn("Trigonometrie") }] },
                { label: "Runden", icon: "rounding", disabled: noCol, menu: [{ label: "Runden …", onClick: nn("Runden") }] },
                { label: "Informationen", icon: "information", disabled: noCol, menu: [{ label: "Ist gerade", onClick: nn("Informationen") }] },
              ],
            },
          ],
        },
        {
          label: "Aus Datum und Uhrzeit",
          icon: "date",
          priority: 20,
          controls: [
            { type: "big", label: "Datum", icon: "date", disabled: noCol, menu: [{ label: "Jahr", onClick: nn("Datum") }] },
            { type: "big", label: "Uhrzeit", icon: "time", disabled: noCol, menu: [{ label: "Stunde", onClick: nn("Uhrzeit") }] },
            { type: "big", label: "Dauer", icon: "duration", disabled: noCol, menu: [{ label: "Tage", onClick: nn("Dauer") }] },
          ],
        },
      ];
    }

    // Ansicht
    const ed = app.editor;
    const toggle = (key) => (checked) => {
      ed[key] = checked;
      render(app);
    };
    return [
      {
        label: "Layout",
        icon: "queriesConnectionsBig",
        priority: 50,
        controls: [
          { type: "big", label: "Abfrage-\neinstellungen", icon: "properties", pressed: ed.settingsPane, onClick: () => toggle("settingsPane")(!ed.settingsPane) },
          { type: "stack", items: [{ type: "check", label: "Bearbeitungsleiste", checked: ed.formulaBar, onChange: toggle("formulaBar") }] },
        ],
      },
      {
        label: "Datenvorschau",
        icon: "table",
        priority: 60,
        controls: [
          {
            type: "stack",
            items: [
              { type: "check", label: "Nichtproportionale Schriftart", checked: ed.mono, onChange: toggle("mono") },
              { type: "check", label: "Leerzeichen anzeigen", checked: ed.whitespace, onChange: toggle("whitespace") },
            ],
          },
          {
            type: "stack",
            items: [
              { type: "check", label: "Spaltenqualität", checked: ed.quality, onChange: toggle("quality") },
              { type: "check", label: "Spaltenverteilung", checked: ed.distribution, onChange: toggle("distribution") },
              { type: "check", label: "Spaltenprofil", checked: false, onChange: () => hint(app, "„Spaltenprofil“ wird in dieser Übung nicht benötigt.") },
            ],
          },
        ],
      },
      { label: "Spalten", icon: "goToColumn", priority: 30, controls: [{ type: "big", label: "Gehe zu\nSpalte", icon: "goToColumn", disabled: noTable, onClick: nn("Gehe zu Spalte") }] },
      { label: "Parameter", icon: "manageParameters", priority: 20, controls: [{ type: "stack", items: [{ type: "check", label: "Parametrisierung immer zulassen", checked: false, onChange: () => hint(app, "Parameter werden in dieser Übung nicht benötigt.") }] }] },
      { label: "Erweitert", icon: "advancedEditorBig", priority: 40, controls: [{ type: "big", label: "Erweiterter\nEditor", icon: "advancedEditorBig", disabled: !hasQuery, onClick: c.advancedEditor }] },
      { label: "Abhängigkeiten", icon: "queryDependencies", priority: 10, controls: [{ type: "big", label: "Abfrage-\nabhängigkeiten", icon: "queryDependencies", disabled: !hasQuery, onClick: nn("Abfrageabhängigkeiten") }] },
    ];
  }

  function splitMenu(app) {
    const nn = (label) => notNeeded(app, label);
    return [
      { label: "Nach Trennzeichen", onClick: () => cmd(app).split() },
      { label: "Nach Anzahl von Zeichen", onClick: nn("Nach Anzahl von Zeichen") },
      { label: "Nach Positionen", onClick: nn("Nach Positionen") },
      { label: "Nach Kleinbuchstabe zu Großbuchstabe", onClick: nn("Nach Kleinbuchstabe zu Großbuchstabe") },
      { label: "Nach Großbuchstabe zu Kleinbuchstabe", onClick: nn("Nach Großbuchstabe zu Kleinbuchstabe") },
      { label: "Nach Ziffer zu Nicht-Ziffer", onClick: nn("Nach Ziffer zu Nicht-Ziffer") },
      { label: "Nach Nicht-Ziffer zu Ziffer", onClick: nn("Nach Nicht-Ziffer zu Ziffer") },
    ];
  }

  /* ---------------- Datenvorschau ---------------- */

  function filteredColumns(app) {
    const set = new Set();
    app.query.steps.slice(0, app.editor.selectedStep + 1).forEach((s) => {
      if (s.action === "filterRows") set.add(s.column);
    });
    return set;
  }

  function columnQuality(table, col) {
    let valid = 0;
    let errors = 0;
    let empty = 0;
    table.rows.forEach((r) => {
      const v = r[col];
      if (PQ.isError(v)) errors++;
      else if (PQ.isEmptyValue(v)) empty++;
      else valid++;
    });
    const n = table.rows.length || 1;
    const pct = (x) => Math.round((x / n) * 100) + " %";
    return h("div", { class: "pq-quality" }, [
      h("div", { class: "pq-quality__bar" }, [
        h("span", { class: "is-valid", style: { width: (valid / n) * 100 + "%" } }),
        h("span", { class: "is-error", style: { width: (errors / n) * 100 + "%" } }),
        h("span", { class: "is-empty", style: { width: (empty / n) * 100 + "%" } }),
      ]),
      h("div", { class: "pq-quality__row" }, [h("i", { class: "is-valid" }), "Gültig", h("b", { text: pct(valid) })]),
      h("div", { class: "pq-quality__row" }, [h("i", { class: "is-error" }), "Fehler", h("b", { text: pct(errors) })]),
      h("div", { class: "pq-quality__row" }, [h("i", { class: "is-empty" }), "Leer", h("b", { text: pct(empty) })]),
    ]);
  }

  function columnDistribution(table, col) {
    const counts = new Map();
    table.rows.forEach((r) => {
      const k = PQ.valueKey(r[col]);
      counts.set(k, (counts.get(k) || 0) + 1);
    });
    const values = Array.from(counts.values());
    const max = Math.max(1, ...values);
    const unique = values.filter((v) => v === 1).length;
    return h("div", { class: "pq-distribution" }, [
      h("div", { class: "pq-distribution__bars" }, values.slice(0, 14).map((v) => h("span", { style: { height: Math.max(8, (v / max) * 100) + "%" } }))),
      h("div", { class: "pq-distribution__text", text: counts.size + " unterschiedliche, " + unique + " eindeutige" }),
    ]);
  }

  function renderGrid(app) {
    const ed = app.editor;
    const r = run(app);
    const q = app.query;

    if (ed.selectedStep + 1 >= r.states.length) {
      const failed = q.steps[r.errorIndex];
      const isFailed = ed.selectedStep === r.errorIndex;
      return h("div", { class: "pq-errorpane" }, [
        h("div", { class: "pq-errorpane__head" }, [icon("warning", 16), h("strong", { text: "Expression.Error: " + r.errorMessage })]),
        h("p", { class: "pq-errorpane__details", text: "Details: Schritt „" + failed.name + "“" }),
        !isFailed
          ? h("button", {
              type: "button",
              class: "xl-btn",
              text: "Zu Fehler wechseln",
              on: {
                click: () => {
                  ed.selectedStep = r.errorIndex;
                  render(app);
                },
              },
            })
          : null,
      ]);
    }

    const table = r.states[ed.selectedStep + 1];
    const filtered = filteredColumns(app);
    const D_ = D();

    const headRow = h("tr", {}, [
      h("th", { class: "pq-grid__corner" }, [
        h("button", {
          type: "button",
          class: "pq-grid__cornerbtn",
          "aria-label": "Tabellenmenü",
          html: window.ExcelFloIcons.get("table", 16) + '<span class="xl-caret"></span>',
          on: {
            mousedown: (e) => e.stopPropagation(),
            click: (e) => D_.tableMenu(app, e.currentTarget),
          },
        }),
      ]),
    ]);

    table.columns.forEach((col) => {
      const selected = ed.selCols.indexOf(col) !== -1;
      const type = table.types[col] || "Text";
      const typeBtn = h("button", {
        type: "button",
        class: "pq-colhead__type",
        "aria-label": "Datentyp ändern (" + (TYPE_LABELS[type] || type) + ")",
        html: window.ExcelFloIcons.typeGlyph(type),
        on: {
          mousedown: (e) => e.stopPropagation(),
          click: (e) => {
            e.stopPropagation();
            if (ed.selCols.indexOf(col) === -1) ed.selCols = [col];
            render(app);
            D_.typeMenu(app, col, app.windowEl.querySelector('th[data-col="' + CSS.escape(col) + '"] .pq-colhead__type'));
          },
        },
      });
      const nameEl =
        ed.renaming === col
          ? renameInput(app, col)
          : h("span", { class: "pq-colhead__name", text: col, on: { dblclick: () => startRename(app, col) } });
      const filterBtn = h("button", {
        type: "button",
        class: "pq-colhead__filter" + (filtered.has(col) ? " is-filtered" : ""),
        "aria-label": "Filter für " + col,
        html: window.ExcelFloIcons.get(filtered.has(col) ? "funnelActive" : "filterArrow", 16),
        on: {
          mousedown: (e) => e.stopPropagation(),
          click: (e) => {
            e.stopPropagation();
            D_.filterMenu(app, col, e.currentTarget);
          },
        },
      });

      const th = h("th", { class: "pq-colhead" + (selected ? " is-selected" : ""), dataset: { col } }, [
        h("div", { class: "pq-colhead__inner" }, [typeBtn, nameEl, filterBtn]),
        ed.quality ? columnQuality(table, col) : null,
        ed.distribution ? columnDistribution(table, col) : null,
      ]);
      th.addEventListener("mousedown", (e) => {
        if (e.button !== 0 || e.target.closest("input")) return;
        selectColumn(app, table, col, e);
      });
      th.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        if (ed.selCols.indexOf(col) === -1) {
          ed.selCols = [col];
          render(app);
        }
        D_.headerMenu(app, col, e.clientX, e.clientY);
      });
      headRow.appendChild(th);
    });

    const tbody = h("tbody");
    table.rows.forEach((row, i) => {
      const tr = h("tr", {}, [h("th", { class: "pq-grid__rownum", text: String(i + 1) })]);
      table.columns.forEach((col) => {
        const v = row[col];
        let cls = "pq-cell";
        if (typeof v === "number" || PQ.isDate(v)) cls += " pq-cell--num";
        if (v === null || v === undefined) cls += " pq-cell--null";
        if (PQ.isError(v)) cls += " pq-cell--error";
        if (ed.selCols.indexOf(col) !== -1) cls += " is-colsel";
        const td = h("td", { class: cls, text: PQ.displayText(v), title: PQ.isError(v) ? v.message : null });
        td.addEventListener("mousedown", (e) => {
          if (e.button !== 0) return;
          selectColumn(app, table, col, e);
        });
        td.addEventListener("contextmenu", (e) => {
          e.preventDefault();
          ed.selCols = [col];
          render(app);
          D_.cellMenu(app, col, v, e.clientX, e.clientY);
        });
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    const grid = h("table", { class: "pq-grid" + (ed.mono ? " is-mono" : "") + (ed.whitespace ? " is-whitespace" : "") }, [h("thead", {}, [headRow]), tbody]);
    const scroller = h("div", { class: "pq-grid-scroll", tabindex: "0", "aria-label": "Datenvorschau" }, [grid]);
    scroller.addEventListener("keydown", (e) => {
      if (e.key === "Delete" && ed.selCols.length && !ed.renaming) {
        e.preventDefault();
        cmd(app).removeColumns();
      }
    });
    if (!table.rows.length) scroller.appendChild(h("p", { class: "pq-grid-empty", text: "Diese Tabelle ist leer." }));
    return scroller;
  }

  function selectColumn(app, table, col, e) {
    const ed = app.editor;
    if (e.ctrlKey || e.metaKey) {
      ed.selCols = ed.selCols.indexOf(col) !== -1 ? ed.selCols.filter((c) => c !== col) : ed.selCols.concat([col]);
    } else if (e.shiftKey && ed.selCols.length) {
      const a = table.columns.indexOf(ed.selCols[ed.selCols.length - 1]);
      const b = table.columns.indexOf(col);
      ed.selCols = table.columns.slice(Math.min(a, b), Math.max(a, b) + 1);
    } else {
      if (ed.selCols.length === 1 && ed.selCols[0] === col) return;
      ed.selCols = [col];
    }
    render(app);
  }

  function renameInput(app, col) {
    const ed = app.editor;
    const input = h("input", { class: "pq-colhead__rename", type: "text", value: col, "aria-label": "Neuer Spaltenname" });
    let done = false;
    const finish = (commit) => {
      if (done) return;
      done = true;
      ed.renaming = null;
      const to = input.value.trim();
      if (commit && to && to !== col) {
        const table = viewTable(app);
        if (table && table.columns.indexOf(to) !== -1) {
          hint(app, "Der Spaltenname „" + to + "“ ist bereits vorhanden.", "warning");
          return;
        }
        ed.selCols = [to];
        insertStep(app, { action: "renameColumn", renames: [{ from: col, to }] });
      } else {
        render(app);
      }
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") finish(true);
      if (e.key === "Escape") finish(false);
      e.stopPropagation();
    });
    input.addEventListener("blur", () => finish(true));
    input.addEventListener("mousedown", (e) => e.stopPropagation());
    setTimeout(() => {
      input.focus();
      input.select();
    }, 0);
    return input;
  }

  /* ---------------- Seitenbereiche ---------------- */

  function queriesPane(app) {
    const ed = app.editor;
    const count = app.query ? 1 : 0;
    if (!ed.queriesPane) {
      return h("div", { class: "pq-queries is-collapsed" }, [
        h("button", {
          type: "button",
          class: "pq-queries__toggle",
          "aria-label": "Abfragenbereich erweitern",
          html: window.ExcelFloIcons.get("chevronRight", 16),
          on: {
            click: () => {
              ed.queriesPane = true;
              render(app);
            },
          },
        }),
        h("span", { class: "pq-queries__vertical", text: "Abfragen" }),
      ]);
    }
    return h("div", { class: "pq-queries" }, [
      h("div", { class: "pq-queries__head" }, [
        h("span", { class: "pq-queries__title", text: "Abfragen [" + count + "]" }),
        h("button", {
          type: "button",
          class: "pq-queries__toggle",
          "aria-label": "Abfragenbereich reduzieren",
          html: window.ExcelFloIcons.get("chevronLeft", 16),
          on: {
            click: () => {
              ed.queriesPane = false;
              render(app);
            },
          },
        }),
      ]),
      app.query ? h("div", { class: "pq-queries__item is-selected" }, [icon("query", 16), h("span", { text: app.query.name })]) : null,
    ]);
  }

  function stepsPane(app) {
    const ed = app.editor;
    const q = app.query;
    const r = run(app);

    const items = [h("li", { class: "pq-step" + (ed.selectedStep === -1 ? " is-selected" : "") }, [
      h("span", { class: "pq-step__del" }),
      h("span", { class: "pq-step__name", text: "Quelle" }),
      h("button", { type: "button", class: "pq-step__gear", "aria-label": "Einstellungen für Quelle", html: window.ExcelFloIcons.get("gearSmall", 16), on: { click: (e) => { e.stopPropagation(); hint(app, "Die Quelle ist die Tabelle „" + app.tableName + "“ in dieser Arbeitsmappe."); } } }),
    ])];
    items[0].addEventListener("click", () => {
      ed.selectedStep = -1;
      render(app);
    });

    q.steps.forEach((step, i) => {
      const hasError = r.errorIndex === i;
      const li = h("li", { class: "pq-step" + (ed.selectedStep === i ? " is-selected" : "") + (hasError ? " has-error" : "") + (r.errorIndex !== -1 && i > r.errorIndex ? " is-after-error" : "") }, [
        h("button", {
          type: "button",
          class: "pq-step__del",
          "aria-label": "Schritt „" + step.name + "“ löschen",
          html: window.ExcelFloIcons.get("close", 16),
          on: {
            click: (e) => {
              e.stopPropagation();
              deleteStep(app, i);
            },
          },
        }),
        ed.renamingStep === i ? stepRenameInput(app, i) : h("span", { class: "pq-step__name", text: step.name, title: PQ.stepDetail(step) }),
        hasError ? icon("warning", 16) : null,
        GEAR_ACTIONS.indexOf(step.action) !== -1 && !(step.action === "filterRows" && (step.operator === "in" || step.operator === "nicht in"))
          ? h("button", {
              type: "button",
              class: "pq-step__gear",
              "aria-label": "Einstellungen für „" + step.name + "“ bearbeiten",
              html: window.ExcelFloIcons.get("gearSmall", 16),
              on: {
                click: (e) => {
                  e.stopPropagation();
                  ed.selectedStep = i;
                  D().editStep(app, i);
                },
              },
            })
          : null,
      ]);
      li.addEventListener("click", () => {
        if (ed.selectedStep === i) return;
        ed.selectedStep = i;
        render(app);
      });
      li.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        ed.selectedStep = i;
        render(app);
        D().stepMenu(app, i, e.clientX, e.clientY);
      });
      items.push(li);
    });

    return h("aside", { class: "pq-settings" }, [
      h("div", { class: "pq-settings__head" }, [
        h("span", { text: "Abfrageeinstellungen" }),
        h("button", {
          type: "button",
          class: "pq-settings__close",
          "aria-label": "Abfrageeinstellungen schließen",
          html: window.ExcelFloIcons.get("close", 16),
          on: {
            click: () => {
              ed.settingsPane = false;
              render(app);
            },
          },
        }),
      ]),
      h("div", { class: "pq-settings__section", text: "EIGENSCHAFTEN" }),
      h("label", { class: "pq-settings__field" }, [h("span", { text: "Name" }), h("input", { class: "xl-input", type: "text", value: q.name, readonly: true })]),
      h("button", { type: "button", class: "pq-settings__link", text: "Alle Eigenschaften", on: { click: notNeeded(app, "Alle Eigenschaften") } }),
      h("div", { class: "pq-settings__section", text: "ANGEWANDTE SCHRITTE" }),
      h("ol", { class: "pq-steps" }, items),
    ]);
  }

  function stepRenameInput(app, i) {
    const step = app.query.steps[i];
    const input = h("input", { class: "pq-step__rename", type: "text", value: step.name });
    let done = false;
    const finish = (commit) => {
      if (done) return;
      done = true;
      const val = input.value.trim();
      app.editor.renamingStep = null;
      if (commit && val && val !== step.name) {
        if (val === "Quelle" || app.query.steps.some((s) => s !== step && s.name === val)) {
          hint(app, "Ein Schritt mit dem Namen „" + val + "“ ist bereits vorhanden.", "warning");
          return;
        }
        step.name = val;
        app.editor.dirty = true;
      }
      render(app);
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") finish(true);
      if (e.key === "Escape") finish(false);
      e.stopPropagation();
    });
    input.addEventListener("blur", () => finish(true));
    input.addEventListener("click", (e) => e.stopPropagation());
    setTimeout(() => {
      input.focus();
      input.select();
    }, 0);
    return input;
  }

  function formulaBar(app) {
    const ed = app.editor;
    const formula = "= " + PQ.stepToM(app.query.steps, ed.selectedStep, app.tableName);
    return h("div", { class: "pq-formulabar" + (ed.formulaExpanded ? " is-expanded" : "") }, [
      h("span", { class: "pq-formulabar__btn", text: "✕", "aria-hidden": "true" }),
      h("span", { class: "pq-formulabar__btn", text: "✓", "aria-hidden": "true" }),
      h("span", { class: "pq-formulabar__fx", text: "fx" }),
      h("div", { class: "pq-formulabar__text", title: "Die Formel wird durch die angewandten Schritte erzeugt.", text: formula }),
      h("button", {
        type: "button",
        class: "pq-formulabar__expand",
        "aria-label": ed.formulaExpanded ? "Bearbeitungsleiste reduzieren" : "Bearbeitungsleiste erweitern",
        html: '<span class="xl-caret' + (ed.formulaExpanded ? " is-up" : "") + '"></span>',
        on: {
          click: () => {
            ed.formulaExpanded = !ed.formulaExpanded;
            render(app);
          },
        },
      }),
    ]);
  }

  /* ---------------- Fenster ---------------- */

  function render(app) {
    const ed = app.editor;
    const q = app.query;

    const titlebar = h("div", { class: "pq-titlebar" }, [
      h("span", { class: "pq-titlebar__left" }, [icon("excelApp", 16), h("span", { class: "pq-titlebar__sep" }), icon("smiley", 16), UI.caret(), UI.caret(), h("span", { class: "pq-titlebar__sep" })]),
      h("span", { class: "pq-titlebar__title", text: "Power Query-Editor" }),
      h("span", { class: "pq-titlebar__controls" }, [
        h("span", { class: "pq-winbtn", text: "—", "aria-hidden": "true" }),
        h("span", { class: "pq-winbtn", text: "❐", "aria-hidden": "true" }),
        h("button", { type: "button", class: "pq-winbtn pq-winbtn--close", "aria-label": "Power Query-Editor schließen", text: "✕", on: { click: () => D().keepChanges(app) } }),
      ]),
    ]);

    const tabsEl = UI.tabs(
      EDITOR_TABS,
      ed.tab,
      (name, el) => {
        if (name === "Datei") {
          UI.openMenu({ anchor: el, placement: "below" }, [
            { label: "Schließen & laden", icon: "closeLoad", disabled: !q, onClick: () => closeAndLoad(app, "table") },
            { label: "Schließen & laden in …", disabled: !q, onClick: () => D().loadTo(app) },
            { label: "Verwerfen & schließen", onClick: () => discardAndClose(app) },
            { separator: true },
            { label: "Optionen und Einstellungen", submenu: [{ label: "Abfrageoptionen", onClick: notNeeded(app, "Abfrageoptionen") }, { label: "Datenquelleneinstellungen", onClick: notNeeded(app, "Datenquelleneinstellungen") }] },
            { label: "Hilfe", submenu: [{ label: "Dokumentation", onClick: notNeeded(app, "Dokumentation") }, { label: "Info", onClick: notNeeded(app, "Info") }] },
          ]);
          return;
        }
        ed.tab = name;
        render(app);
      },
      (name) => (name === "Datei" ? "pq-tab--file" : "")
    );
    tabsEl.querySelector(".pq-tab--file").addEventListener("mousedown", (e) => e.stopPropagation());

    const ribbonRow = UI.renderRibbon(ribbonGroups(app, ed.tab));
    const table = q ? viewTable(app) : null;
    const r = q ? run(app) : null;

    let center;
    if (!q) {
      center = h("div", { class: "pq-center pq-center--empty" });
    } else {
      center = h("div", { class: "pq-center" }, [ed.formulaBar ? formulaBar(app) : null, renderGrid(app)]);
    }

    const status = q
      ? h("div", { class: "pq-statusbar" }, [
          h("span", { text: table ? table.columns.length + " SPALTEN, " + table.rows.length + " ZEILEN" : "" }),
          h("span", { class: "pq-statusbar__profile", text: table ? "Spaltenprofil basierend auf den obersten 1000 Zeilen" : "" }),
          h("span", { class: "pq-statusbar__right", text: r && r.errorIndex !== -1 ? "FEHLER IN SCHRITT „" + q.steps[r.errorIndex].name.toUpperCase() + "“" : "VORSCHAU UM " + ed.lastRefresh + " HERUNTERGELADEN" }),
        ])
      : h("div", { class: "pq-statusbar" });

    const win = h("div", { class: "xl-window pq-window" }, [
      titlebar,
      h("div", { class: "pq-tabsrow" }, [
        tabsEl,
        h("span", { class: "pq-tabsrow__right" }, [h("span", { class: "xl-caret is-up", "aria-hidden": "true" }), h("span", { html: window.ExcelFloIcons.get("help", 16), "aria-hidden": "true" })]),
      ]),
      h("div", { class: "xl-ribbon pq-ribbon" }, [ribbonRow]),
      UI.messageBar(ed.message, () => {
        ed.message = null;
        render(app);
      }),
      h("div", { class: "pq-body" }, [queriesPane(app), center, q && ed.settingsPane ? stepsPane(app) : null]),
      status,
    ]);

    app.windowEl.replaceChildren(win);
    UI.observeRibbon(ribbonRow);
  }

  /* ---------------- Öffnen / Schließen ---------------- */

  function open(app) {
    const prev = app.editor || {};
    app.view = "editor";
    app.editor = {
      tab: "Start",
      selectedStep: app.query ? app.query.steps.length - 1 : -1,
      selCols: [],
      queriesPane: !app.query,
      settingsPane: prev.settingsPane !== false,
      formulaBar: prev.formulaBar !== false,
      mono: !!prev.mono,
      whitespace: prev.whitespace !== false,
      quality: !!prev.quality,
      distribution: !!prev.distribution,
      message: null,
      renaming: null,
      renamingStep: null,
      dirty: false,
      snapshot: app.query ? clone(app.query.steps) : null,
      lastRefresh: nowTime(),
    };
    app.render();
  }

  function closeAndLoad(app, mode) {
    if (!app.query) {
      backToExcel(app);
      return;
    }
    const r = run(app);
    if (r.errorIndex !== -1) {
      app.editor.selectedStep = r.errorIndex;
      hint(app, "Die Abfrage kann nicht geladen werden, weil der Schritt „" + app.query.steps[r.errorIndex].name + "“ einen Fehler enthält.", "warning");
      return;
    }
    const prevLoaded = app.query.loaded;
    app.query.loaded = {
      steps: clone(app.query.steps),
      result: r.states[r.states.length - 1],
      mode,
      sheetName: prevLoaded && prevLoaded.sheetName ? prevLoaded.sheetName : app.outputSheetName,
    };
    app.check = null;
    app.excel.queriesPane = true;
    if (mode === "table") {
      app.excel.sheet = app.query.loaded.sheetName;
      app.excel.sel = { r: 0, c: 0 };
    }
    backToExcel(app);
    if (app.onLoaded) app.onLoaded();
  }

  function discardAndClose(app) {
    const ed = app.editor;
    if (app.query) {
      if (app.query.loaded) app.query.steps = clone(ed.snapshot || app.query.loaded.steps);
      else app.query = null; // neue, nie geladene Abfrage wird verworfen
    }
    backToExcel(app);
  }

  function backToExcel(app) {
    UI.closeMenus(0);
    app.view = "excel";
    app.render();
  }

  function refreshAll(app) {
    if (!app.query || !app.query.loaded) {
      app.hint("Es gibt noch keine geladene Abfrage, die aktualisiert werden kann.");
      return;
    }
    const r = PQ.runSteps(app.source, app.query.loaded.steps);
    if (r.errorIndex === -1) app.query.loaded.result = r.states[r.states.length - 1];
    app.excel.status = "Bereit";
    app.render();
  }

  window.ExcelFloPQEditor = {
    open,
    render,
    ensureQuery,
    run,
    tableAt,
    viewTable,
    insertStep,
    replaceStep,
    commitStep,
    deleteStep,
    selectedColumns,
    isAllowed,
    hint,
    closeAndLoad,
    discardAndClose,
    refreshAll,
    startRename,
    cmd,
    typeMenuItems,
    splitMenu,
    notNeeded,
    TYPE_LABELS,
    ACTION_LABELS,
  };
})();
