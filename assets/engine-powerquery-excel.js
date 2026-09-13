/* Excel.Flo – Nachbau des Excel-Fensters (Registerkarte „Daten“)
 *
 * Zeigt die Rohdaten auf „Tabelle1“, das vollständige Menüband „Daten“ und –
 * nach „Schließen & laden“ – das Abfrageergebnis auf einem neuen Blatt samt
 * Aufgabenbereich „Abfragen & Verbindungen“. Funktional sind alle Befehle, die
 * für Power Query gebraucht werden; alle anderen öffnen ihre Menüs, zeigen aber
 * nur eine Hinweisleiste.
 */

(function () {
  "use strict";

  const UI = window.ExcelFloUI;
  const PQ = window.ExcelFloPowerQuery;
  const { h, icon } = UI;

  const TABS = ["Datei", "Start", "Einfügen", "Zeichnen", "Seitenlayout", "Formeln", "Daten", "Überprüfen", "Ansicht", "Automatisieren", "Entwicklertools", "Hilfe", "Power Pivot"];
  const MIN_COLS = 14;
  const MIN_ROWS = 22;

  function colLetter(i) {
    let s = "";
    let n = i + 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function cellText(v) {
    if (v === null || v === undefined || PQ.isError(v)) return "";
    if (typeof v === "string") return v;
    return PQ.displayText(v);
  }

  /* ---------------- Blätter ---------------- */

  function sheetsOf(app) {
    const list = [{ name: app.sourceSheetName, kind: "source" }];
    if (app.query && app.query.loaded && app.query.loaded.mode === "table") {
      list.push({ name: app.query.loaded.sheetName, kind: "output" });
    }
    return list;
  }

  function sheetTable(app, sheet) {
    if (sheet.kind === "output") return app.query.loaded.result;
    return app.source;
  }

  /* ---------------- Befehle ---------------- */

  function sourceRangeRef(app) {
    return "=$A$1:$" + colLetter(app.source.columns.length - 1) + "$" + (app.source.rows.length + 1);
  }

  function selectionInSource(app) {
    const sel = app.excel.sel;
    return app.excel.sheet === app.sourceSheetName && sel.r <= app.source.rows.length && sel.c < app.source.columns.length;
  }

  function fromTableRange(app) {
    if (app.excel.sheet !== app.sourceSheetName) {
      app.hint("Wähle auf dem Blatt „" + app.sourceSheetName + "“ eine Zelle in den Rohdaten aus, bevor du „Aus Tabelle/Bereich“ nutzt.");
      return;
    }
    if (app.tableCreated) {
      if (!selectionInSource(app)) {
        app.hint("Markiere eine Zelle in der Tabelle „" + app.tableName + "“, um sie in Power Query zu öffnen.");
        return;
      }
      app.ensureQuery();
      app.openEditor();
      return;
    }

    const inside = selectionInSource(app);
    const state = { ref: inside ? sourceRangeRef(app) : "", headers: true };
    const body = h("div", { class: "xl-create-table" }, [
      h("p", { class: "xl-create-table__label", text: "Wo sind die Daten für die Tabelle?" }),
      UI.input(state.ref, (v) => (state.ref = v), { "aria-label": "Datenbereich" }),
      UI.checkbox("Tabelle hat Überschriften", true, (v) => (state.headers = v)),
    ]);

    UI.openDialog(app.host, {
      titlebar: "Tabelle erstellen",
      width: 300,
      body,
      buttons: [
        {
          label: "OK",
          primary: true,
          onClick: ({ setError }) => {
            const norm = (s) => s.replace(/[\s$=]/g, "").toUpperCase();
            if (norm(state.ref) !== norm(sourceRangeRef(app))) {
              setError(
                state.ref.trim()
                  ? "Der Bereich passt nicht zu den Rohdaten (" + sourceRangeRef(app).replace("=", "") + ")."
                  : "Der Bezug ist ungültig. Markiere zuerst eine Zelle innerhalb der Rohdaten."
              );
              return false;
            }
            if (!state.headers) {
              setError("Die erste Zeile enthält die Spaltennamen – aktiviere „Tabelle hat Überschriften“.");
              return false;
            }
            app.tableCreated = true;
            app.ensureQuery();
            app.openEditor();
            return true;
          },
        },
        { label: "Abbrechen" },
      ],
    });
  }

  function launchEditor(app) {
    app.openEditor();
  }

  function dataGetMenu(app) {
    const n = (label, iconName) => ({ label, icon: iconName, onClick: () => app.hint("„" + label + "“ wird in dieser Übung nicht benötigt.") });
    return [
      { label: "Aus Datei", icon: "fromTextCsv", submenu: [n("Aus Arbeitsmappe", "table"), n("Aus Text/CSV", "fromTextCsv"), n("Aus XML", "fromTextCsv"), n("Aus JSON", "parse"), n("Aus PDF", "fromTextCsv"), n("Aus Ordner", "existingConnections"), n("Aus SharePoint-Ordner", "existingConnections")] },
      { label: "Aus Datenbank", icon: "manageDataModel", submenu: [n("Aus SQL Server-Datenbank", "manageDataModel"), n("Aus Microsoft Access-Datenbank", "manageDataModel"), n("Aus SQL Server Analysis Services-Datenbank (Import)", "manageDataModel")] },
      { label: "Aus Azure", icon: "fromWeb", submenu: [n("Aus Azure SQL-Datenbank", "manageDataModel"), n("Aus Azure Synapse Analytics", "manageDataModel"), n("Aus Azure Blob Storage", "fromWeb"), n("Aus Azure Data Lake Storage", "fromWeb")] },
      { label: "Aus Power Platform", icon: "dataAnalysis", submenu: [n("Aus Power BI-Semantikmodellen", "dataAnalysis"), n("Aus Dataflows", "dataAnalysis"), n("Aus Dataverse", "dataAnalysis")] },
      { label: "Aus Onlinedienst", icon: "fromWeb", submenu: [n("Aus SharePoint Online-Liste", "toList"), n("Aus Microsoft Exchange Online", "fromWeb"), n("Aus Salesforce-Objekten", "fromWeb")] },
      {
        label: "Aus anderen Quellen",
        icon: "existingConnections",
        submenu: [
          { label: "Aus Tabelle/Bereich", icon: "fromTable", onClick: () => fromTableRange(app) },
          n("Aus dem Web", "fromWeb"),
          n("Von Microsoft Query", "manageDataModel"),
          n("Aus SharePoint-Liste", "toList"),
          n("Aus OData-Feed", "fromWeb"),
          n("Aus Hadoop-Datei (HDFS)", "manageDataModel"),
          n("Aus Active Directory", "relationships"),
          n("Aus Microsoft Exchange", "fromWeb"),
          n("Aus ODBC", "manageDataModel"),
          n("Aus OLEDB", "manageDataModel"),
          { separator: true },
          n("Leere Abfrage", "advancedEditor"),
        ],
      },
      { separator: true },
      { label: "Abfragen kombinieren", icon: "mergeQueries", submenu: [n("Zusammenführen", "mergeQueries"), n("Anfügen", "appendQueries")] },
      { separator: true },
      { label: "Power Query-Editor starten …", icon: "advancedEditor", onClick: () => launchEditor(app) },
      n("Datenquelleneinstellungen …", "properties"),
      n("Abfrageoptionen", "gearSmall"),
    ];
  }

  function notNeeded(app, label) {
    return () => app.hint("„" + label.replace("\n", " ") + "“ wird in dieser Übung nicht benötigt.");
  }

  function dataRibbonGroups(app) {
    const nn = (label) => notNeeded(app, label);
    const paneOpen = app.excel.queriesPane;
    return [
      {
        label: "Daten abrufen und transformieren",
        icon: "dataGet",
        priority: 90,
        controls: [
          {
            type: "big",
            label: "Daten\nabrufen",
            icon: "dataGet",
            menu: () => dataGetMenu(app),
            tip: { title: "Daten abrufen", text: "Importieren Sie Daten aus vielen verschiedenen Quellen und bearbeiten Sie sie im Power Query-Editor." },
          },
          {
            type: "stack",
            iconOnly: true,
            items: [
              { label: "Aus Text/CSV", icon: "fromTextCsv", onClick: nn("Aus Text/CSV"), tip: { title: "Aus Text/CSV", text: "Importieren Sie Daten aus einer Text-, CSV- oder PRN-Datei." } },
              { label: "Aus dem Web", icon: "fromWeb", onClick: nn("Aus dem Web"), tip: { title: "Aus dem Web", text: "Importieren Sie Daten von einer Webseite." } },
              {
                label: "Aus Tabelle/Bereich",
                icon: "fromTable",
                onClick: () => fromTableRange(app),
                tip: { title: "Aus Tabelle/Bereich", text: "Erstellen Sie eine neue Abfrage, die mit der ausgewählten Tabelle oder dem benannten Bereich verknüpft ist." },
              },
            ],
          },
          {
            type: "stack",
            iconOnly: true,
            items: [
              {
                label: "Aus Bild",
                icon: "fromPicture",
                menu: [
                  { label: "Bild aus Datei", onClick: nn("Bild aus Datei") },
                  { label: "Bild aus Zwischenablage", onClick: nn("Bild aus Zwischenablage") },
                ],
                tip: { title: "Aus Bild", text: "Wandeln Sie ein Bild einer Tabelle in bearbeitbare Daten um." },
              },
              { label: "Zuletzt verwendete Quellen", icon: "recentSources", onClick: nn("Zuletzt verwendete Quellen") },
              { label: "Vorhandene Verbindungen", icon: "existingConnections", onClick: nn("Vorhandene Verbindungen") },
            ],
          },
        ],
      },
      {
        label: "Abfragen & Verbindungen",
        icon: "refreshAll",
        priority: 60,
        controls: [
          {
            type: "big",
            label: "Alle\naktualisieren",
            icon: "refreshAll",
            onClick: () => app.refreshAll(),
            menu: () => [
              { label: "Alle aktualisieren", icon: "refreshPreview", shortcut: "Strg+Alt+F5", onClick: () => app.refreshAll() },
              { label: "Aktualisieren", onClick: () => app.refreshAll(), shortcut: "Alt+F5" },
              { label: "Aktualisierungsstatus", disabled: true },
              { label: "Aktualisierung abbrechen", disabled: true },
              { label: "Verbindungseigenschaften …", onClick: nn("Verbindungseigenschaften") },
            ],
            tip: { title: "Alle aktualisieren (Strg+Alt+F5)", text: "Rufen Sie die neuesten Daten ab, indem Sie alle Quellen in der Arbeitsmappe aktualisieren." },
          },
          {
            type: "stack",
            iconOnly: true,
            items: [
              {
                label: "Abfragen und Verbindungen",
                icon: "queriesConnections",
                pressed: paneOpen,
                onClick: () => {
                  app.excel.queriesPane = !app.excel.queriesPane;
                  app.render();
                },
                tip: { title: "Abfragen und Verbindungen", text: "Zeigen Sie den Aufgabenbereich mit allen Abfragen und Verbindungen der Arbeitsmappe an." },
              },
              { label: "Eigenschaften", icon: "properties", disabled: true },
              { label: "Verknüpfungen bearbeiten", icon: "editLinks", disabled: true },
            ],
          },
        ],
      },
      {
        label: "Datentypen",
        icon: "stocks",
        priority: 20,
        controls: [
          {
            type: "gallery",
            items: [
              { label: "Aktien", icon: "stocks", onClick: nn("Aktien") },
              { label: "Währungen", icon: "currencies", onClick: nn("Währungen") },
            ],
            onMore: nn("Weitere Datentypen"),
          },
        ],
      },
      {
        label: "Sortieren und Filtern",
        icon: "filter",
        priority: 40,
        controls: [
          {
            type: "stack",
            iconOnly: true,
            items: [
              { label: "Von A bis Z sortieren", icon: "sortAZ", onClick: nn("Von A bis Z sortieren") },
              { label: "Von Z bis A sortieren", icon: "sortZA", onClick: nn("Von Z bis A sortieren") },
            ],
          },
          { type: "big", label: "Sortieren", icon: "sort", onClick: nn("Sortieren") },
          { type: "big", label: "Filtern", icon: "filter", onClick: nn("Filtern"), tip: { title: "Filtern (Strg+Umschalt+L)", text: "Aktivieren Sie die Filterung für die ausgewählten Zellen." } },
          {
            type: "stack",
            items: [
              { label: "Löschen", icon: "clearFilter", disabled: true },
              { label: "Erneut anwenden", icon: "reapply", disabled: true },
              { label: "Erweitert", icon: "advancedFilter", onClick: nn("Erweitert") },
            ],
          },
        ],
      },
      {
        label: "Datentools",
        icon: "textToColumns",
        priority: 30,
        controls: [
          { type: "big", label: "Text in\nSpalten", icon: "textToColumns", onClick: nn("Text in Spalten"), tip: { title: "Text in Spalten", text: "Teilen Sie den Inhalt einer Zelle in mehrere Spalten auf." } },
          { type: "big", label: "Daten\nbereinigen", icon: "cleanData", onClick: nn("Daten bereinigen") },
          {
            type: "stack",
            iconOnly: true,
            items: [
              { label: "Blitzvorschau", icon: "flashFill", onClick: nn("Blitzvorschau") },
              { label: "Duplikate entfernen", icon: "removeDuplicates", onClick: nn("Duplikate entfernen") },
              {
                label: "Datenüberprüfung",
                icon: "dataValidation",
                menu: [
                  { label: "Datenüberprüfung …", onClick: nn("Datenüberprüfung") },
                  { label: "Ungültige Daten einkreisen", onClick: nn("Ungültige Daten einkreisen") },
                  { label: "Gültigkeitskreise löschen", onClick: nn("Gültigkeitskreise löschen") },
                ],
              },
            ],
          },
          {
            type: "stack",
            iconOnly: true,
            items: [
              { label: "Konsolidieren", icon: "consolidate", onClick: nn("Konsolidieren") },
              { label: "Beziehungen", icon: "relationships", onClick: nn("Beziehungen") },
              { label: "Datenmodell verwalten", icon: "manageDataModel", onClick: nn("Datenmodell verwalten") },
            ],
          },
        ],
      },
      {
        label: "Prognose",
        icon: "whatIf",
        priority: 35,
        controls: [
          {
            type: "big",
            label: "Was-wäre-wenn-\nAnalyse",
            icon: "whatIf",
            menu: [
              { label: "Szenario-Manager …", onClick: nn("Szenario-Manager") },
              { label: "Zielwertsuche …", onClick: nn("Zielwertsuche") },
              { label: "Datentabelle …", onClick: nn("Datentabelle") },
            ],
          },
          { type: "big", label: "Prognose-\nblatt", icon: "forecastSheet", onClick: nn("Prognoseblatt") },
        ],
      },
      {
        label: "Gliederung",
        icon: "outline",
        priority: 25,
        controls: [
          {
            type: "big",
            label: "Gliederung",
            icon: "outline",
            menu: [
              { label: "Gruppieren", onClick: nn("Gruppieren") },
              { label: "Gruppierung aufheben", onClick: nn("Gruppierung aufheben") },
              { label: "Teilergebnis", onClick: nn("Teilergebnis") },
            ],
          },
        ],
      },
      {
        label: "Analyse",
        icon: "dataAnalysis",
        priority: 10,
        controls: [{ type: "stack", items: [{ label: "Datenanalyse", icon: "dataAnalysis", onClick: nn("Datenanalyse") }] }],
      },
    ];
  }

  /* ---------------- Rendering ---------------- */

  function titlebar(app) {
    return h("div", { class: "xl-titlebar" }, [
      h("div", { class: "xl-titlebar__left" }, [
        icon("excelApp", 16),
        h("span", { class: "xl-autosave" }, [h("span", { text: "Automatisches Speichern" }), h("span", { class: "xl-toggle", "aria-hidden": "true" }, [h("span", { text: "Aus" })])]),
        h("span", { class: "xl-qat" }, [icon("save", 16), icon("undo", 16), UI.caret(), icon("redo", 16), UI.caret()]),
      ]),
      h("div", { class: "xl-titlebar__center" }, [
        h("span", { class: "xl-filename" }, [h("strong", { text: app.workbookName }), h("span", { text: " • Gespeichert" }), UI.caret()]),
        h("span", { class: "xl-searchbox" }, [icon("search", 16), h("span", { text: "Suchen" })]),
      ]),
      h("div", { class: "xl-titlebar__right" }, [
        h("span", { class: "xl-avatar", text: "EF", "aria-hidden": "true" }),
        h("span", { class: "xl-winctl", "aria-hidden": "true" }, [h("span", { text: "—" }), h("span", { text: "▢" }), h("span", { text: "✕" })]),
      ]),
    ]);
  }

  function formulaBar(app, refs) {
    refs.nameBox = h("span", { class: "xl-namebox__value" });
    refs.fx = h("span", { class: "xl-formulabar__content" });
    return h("div", { class: "xl-formulabar" }, [
      h("div", { class: "xl-namebox" }, [refs.nameBox, UI.caret()]),
      h("span", { class: "xl-formulabar__dots", text: "⋮" }),
      h("span", { class: "xl-formulabar__btns" }, [h("span", { text: "✕" }), h("span", { text: "✓" }), h("span", { class: "xl-formulabar__fx", text: "fx" }), UI.caret()]),
      refs.fx,
    ]);
  }

  function grid(app, sheet, refs) {
    const table = sheetTable(app, sheet);
    const nCols = Math.max(MIN_COLS, table.columns.length + 4);
    const nRows = Math.max(MIN_ROWS, table.rows.length + 8);
    const isTable = sheet.kind === "output" || app.tableCreated;
    const check = sheet.kind === "output" ? app.check : null;

    const widths = [];
    for (let c = 0; c < nCols; c++) {
      if (c < table.columns.length) {
        const texts = [table.columns[c]].concat(table.rows.map((r) => cellText(r[table.columns[c]])));
        const max = texts.reduce((m, t) => Math.max(m, t.length), 0);
        widths.push(Math.min(220, Math.max(72, Math.round(max * 7.4 + (isTable ? 34 : 18)))));
      } else {
        widths.push(72);
      }
    }

    const colgroup = h("colgroup", {}, [h("col", { style: { width: "36px" } })].concat(widths.map((w) => h("col", { style: { width: w + "px" } }))));
    const headRow = h("tr", {}, [h("th", { class: "xl-grid__corner" })].concat(widths.map((_, c) => h("th", { class: "xl-grid__colhead", dataset: { c }, text: colLetter(c) }))));
    const tbody = h("tbody");

    for (let r = 0; r < nRows; r++) {
      const tr = h("tr", {}, [h("th", { class: "xl-grid__rowhead", dataset: { r }, text: String(r + 1) })]);
      for (let c = 0; c < nCols; c++) {
        let text = "";
        let cls = "xl-cell";
        const inData = c < table.columns.length && r <= table.rows.length;
        if (inData) {
          const col = table.columns[c];
          if (r === 0) {
            text = col;
            if (isTable) cls += " xl-cell--thead";
          } else {
            const v = table.rows[r - 1][col];
            text = cellText(v);
            if (typeof v === "number" || PQ.isDate(v)) cls += " xl-cell--num";
            if (isTable) cls += r % 2 === 1 ? " xl-cell--band" : " xl-cell--tbody";
            if (check && (check.wrongCells.has(r - 1 + "|" + col) || check.wrongRows.has(r - 1))) cls += " is-wrong";
          }
          if (r === 0 && check && check.wrongColumns.has(col)) cls += " is-wrong";
        }
        const td = h("td", { class: cls, dataset: { r, c } }, [
          h("span", { class: "xl-cell__text", text }),
          inData && r === 0 && isTable ? h("span", { class: "xl-cell__filterbtn", "aria-hidden": "true" }, [icon("filterArrow", 16)]) : null,
        ]);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    const tableEl = h("table", { class: "xl-grid" }, [colgroup, h("thead", {}, [headRow]), tbody]);
    const scroller = h("div", { class: "xl-grid-scroll", tabindex: "0", "aria-label": "Tabellenblatt " + sheet.name });

    function textAt(r, c) {
      if (c >= table.columns.length || r > table.rows.length) return "";
      return r === 0 ? table.columns[c] : cellText(table.rows[r - 1][table.columns[c]]);
    }

    function select(r, c) {
      r = Math.max(0, Math.min(nRows - 1, r));
      c = Math.max(0, Math.min(nCols - 1, c));
      app.excel.sel = { r, c };
      tableEl.querySelectorAll(".is-selected, .is-active-head").forEach((n) => n.classList.remove("is-selected", "is-active-head"));
      const td = tableEl.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
      if (td) td.classList.add("is-selected");
      const ch = tableEl.querySelector('th.xl-grid__colhead[data-c="' + c + '"]');
      const rh = tableEl.querySelector('th.xl-grid__rowhead[data-r="' + r + '"]');
      if (ch) ch.classList.add("is-active-head");
      if (rh) rh.classList.add("is-active-head");
      refs.nameBox.textContent = colLetter(c) + (r + 1);
      refs.fx.textContent = textAt(r, c);
    }

    tableEl.addEventListener("mousedown", (e) => {
      const td = e.target.closest("td[data-r]");
      if (!td) return;
      select(Number(td.dataset.r), Number(td.dataset.c));
      scroller.focus({ preventScroll: true });
    });
    tableEl.addEventListener("dblclick", (e) => {
      if (e.target.closest("td[data-r]")) app.hint("In dieser Übung bereinigst du die Daten nicht direkt in den Zellen, sondern mit Power Query.");
    });
    scroller.addEventListener("keydown", (e) => {
      const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, 1] };
      const mv = moves[e.key];
      if (!mv) return;
      e.preventDefault();
      select(app.excel.sel.r + mv[0], app.excel.sel.c + mv[1]);
      const td = tableEl.querySelector("td.is-selected");
      if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    });

    scroller.appendChild(tableEl);
    refs.select = select;
    return scroller;
  }

  function queriesPane(app) {
    const q = app.query;
    let items = [];
    if (q) {
      let status = "Nicht geladen.";
      if (q.loaded) status = q.loaded.mode === "table" ? q.loaded.result.rows.length + " Zeilen geladen." : "Nur Verbindung.";
      const item = h(
        "div",
        {
          class: "xl-querylist__item",
          tabindex: "0",
          "data-tip-title": q.name,
          "data-tip-text": "Doppelklicken, um die Abfrage im Power Query-Editor zu bearbeiten.",
        },
        [icon("query", 16), h("span", { class: "xl-querylist__text" }, [h("strong", { text: q.name }), h("span", { text: status })])]
      );
      item.addEventListener("dblclick", () => app.openEditor());
      item.addEventListener("keydown", (e) => e.key === "Enter" && app.openEditor());
      item.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        UI.openMenu({ rect: { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY }, placement: "point" }, [
          { label: "Bearbeiten", icon: "advancedEditor", onClick: () => app.openEditor() },
          { label: "Aktualisieren", icon: "refreshPreview", onClick: () => app.refreshAll() },
          { label: "Laden in …", onClick: notNeeded(app, "Laden in") },
          { separator: true },
          { label: "Löschen", onClick: notNeeded(app, "Löschen") },
          { label: "Umbenennen", onClick: notNeeded(app, "Umbenennen") },
          { separator: true },
          { label: "Eigenschaften …", icon: "properties", onClick: notNeeded(app, "Eigenschaften") },
        ]);
      });
      items = [item];
    }
    return h("aside", { class: "xl-taskpane" }, [
      h("div", { class: "xl-taskpane__head" }, [
        h("span", { class: "xl-taskpane__title", text: "Abfragen & Verbindungen" }),
        h("button", {
          type: "button",
          class: "xl-taskpane__close",
          "aria-label": "Aufgabenbereich schließen",
          html: window.ExcelFloIcons.get("close", 16),
          on: {
            click: () => {
              app.excel.queriesPane = false;
              app.render();
            },
          },
        }),
      ]),
      h("div", { class: "xl-taskpane__tabs" }, [h("span", { class: "is-active", text: "Abfragen" }), h("span", { text: "Verbindungen" })]),
      h("div", { class: "xl-taskpane__count", text: q ? "1 Abfrage" : "0 Abfragen" }),
      h("div", { class: "xl-querylist" }, items),
    ]);
  }

  function sheetTabs(app) {
    return h("div", { class: "xl-sheettabs" }, [
      h("span", { class: "xl-sheettabs__nav", "aria-hidden": "true" }, [icon("chevronLeft", 16), icon("chevronRight", 16)]),
      h(
        "div",
        { class: "xl-sheettabs__list", role: "tablist" },
        sheetsOf(app).map((s) =>
          h("button", {
            type: "button",
            role: "tab",
            class: "xl-sheettab" + (s.name === app.excel.sheet ? " is-active" : ""),
            text: s.name,
            on: {
              click: () => {
                if (s.name === app.excel.sheet) return;
                app.excel.sheet = s.name;
                app.excel.sel = { r: 0, c: 0 };
                app.render();
              },
            },
          })
        )
      ),
      h("button", { type: "button", class: "xl-sheettabs__add", "aria-label": "Neues Blatt", text: "+", on: { click: notNeeded(app, "Neues Blatt") } }),
    ]);
  }

  function statusBar(app) {
    return h("div", { class: "xl-statusbar" }, [
      h("span", { text: app.excel.status || "Bereit" }),
      h("span", { class: "xl-statusbar__right" }, [
        h("span", { class: "xl-statusbar__views", "aria-hidden": "true" }, [h("span", { text: "▦" }), h("span", { text: "▤" }), h("span", { text: "▥" })]),
        h("span", { class: "xl-zoom", "aria-hidden": "true" }, [h("span", { text: "−" }), h("span", { class: "xl-zoom__track" }, [h("span", { class: "xl-zoom__thumb" })]), h("span", { text: "+" }), h("span", { text: "100 %" })]),
      ]),
    ]);
  }

  function render(app) {
    const st = app.excel;
    const sheets = sheetsOf(app);
    if (!sheets.some((s) => s.name === st.sheet)) st.sheet = sheets[0].name;
    const sheet = sheets.find((s) => s.name === st.sheet);
    const refs = {};

    const tabsEl = UI.tabs(
      TABS,
      "Daten",
      (name) => {
        if (name === "Daten") return;
        st.message = { text: "Alles, was du für diese Übung brauchst, findest du auf der Registerkarte „Daten“." };
        app.render();
      },
      (name) => (name === "Datei" ? "xl-tab--file" : "")
    );

    const ribbonRow = UI.renderRibbon(dataRibbonGroups(app));
    const win = h("div", { class: "xl-window xl-window--excel" }, [
      titlebar(app),
      h("div", { class: "xl-tabsrow" }, [
        tabsEl,
        h("div", { class: "xl-tabsrow__right" }, [
          h("button", { type: "button", class: "xl-pillbtn", on: { click: notNeeded(app, "Kommentare") } }, [icon("comment", 16), "Kommentare"]),
          h("button", { type: "button", class: "xl-pillbtn xl-pillbtn--share", on: { click: notNeeded(app, "Freigeben") } }, [icon("share", 16), "Freigeben", UI.caret()]),
        ]),
      ]),
      h("div", { class: "xl-ribbon" }, [ribbonRow]),
      UI.messageBar(st.message, () => {
        st.message = null;
        app.render();
      }),
      formulaBar(app, refs),
      h("div", { class: "xl-body" }, [grid(app, sheet, refs), st.queriesPane ? queriesPane(app) : null]),
      sheetTabs(app),
      statusBar(app),
    ]);

    app.windowEl.replaceChildren(win);
    UI.observeRibbon(ribbonRow);
    refs.select(st.sel.r, st.sel.c);
  }

  window.ExcelFloPQExcel = { render, colLetter, fromTableRange };
})();
