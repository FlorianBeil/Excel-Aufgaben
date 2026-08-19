/* Excel.Flo – gemeinsame Render- und Prüf-Logik
 * Keine Abhängigkeiten, kein Build-Schritt. Läuft auf Übersichts- und Übungsseiten.
 */

(function () {
  "use strict";

  const LEVELS = [
    { id: "anfaenger", label: "Anfänger" },
    { id: "fortgeschritten", label: "Fortgeschritten" },
    { id: "profi", label: "Profi" },
  ];
  const LEVEL_LABELS = LEVELS.reduce((m, l) => {
    m[l.id] = l.label;
    return m;
  }, {});

  const REF_COLORS = ["#1a73e8", "#e8710a", "#a142f4", "#188038", "#d01884", "#0b8a8a"];

  const FUNCTION_SIGNATURES = {
    SVERWEIS: ["Suchkriterium", "Matrix", "Spaltenindex", "[Bereich_Verweis]"],
    WVERWEIS: ["Suchkriterium", "Matrix", "Zeilenindex", "[Bereich_Verweis]"],
    WENN: ["Prüfung", "[Dann_Wert]", "[Sonst_Wert]"],
    SUMME: ["Zahl1", "[Zahl2]"],
    SUMMEWENN: ["Bereich", "Kriterium", "[Summe_Bereich]"],
    ZÄHLENWENN: ["Bereich", "Kriterium"],
    RUNDEN: ["Zahl", "Anzahl_Stellen"],
    VERGLEICH: ["Suchkriterium", "Suchmatrix", "[Vergleichstyp]"],
    INDEX: ["Matrix", "Zeile", "[Spalte]"],
    MIN: ["Zahl1", "[Zahl2]"],
    MAX: ["Zahl1", "[Zahl2]"],
    MITTELWERT: ["Zahl1", "[Zahl2]"],
    ANZAHL: ["Wert1", "[Wert2]"],
    GROSS: ["Text"],
    KLEIN: ["Text"],
    GROSS2: ["Text"],
  };

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    if (attrs) {
      for (const key in attrs) {
        if (key === "class") node.className = attrs[key];
        else if (key === "text") node.textContent = attrs[key];
        else if (key === "html") node.innerHTML = attrs[key];
        else node.setAttribute(key, attrs[key]);
      }
    }
    (children || []).forEach((c) => c && node.appendChild(c));
    return node;
  }

  function colLetter(index) {
    // 0 -> A, 1 -> B, ...
    let n = index + 1;
    let s = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function colIndexFromLetters(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n - 1;
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function formatCategoryLabel(category) {
    if (!category) return "";
    return category.charAt(0).toUpperCase() + category.slice(1);
  }

  /* ---------------- Übersichtsseite (Level-Tabs, generiert aus dem Manifest) ---------------- */

  function initOverview() {
    const root = document.getElementById("exercise-list");
    if (!root) return;

    const manifestPath = root.dataset.manifestPath || "assets/exercises/manifest.json";
    const exercisePagePath = root.dataset.exercisePage || "uebung.html";
    const tabsRoot = document.getElementById("level-tabs");
    const bannerRoot = document.getElementById("stage-banner");

    fetch(manifestPath, { cache: "no-cache" })
      .then((res) => {
        if (!res.ok) throw new Error("Manifest konnte nicht geladen werden (" + res.status + ")");
        return res.json();
      })
      .then((exercises) => setupOverview(exercises, root, tabsRoot, bannerRoot, exercisePagePath))
      .catch((err) => {
        root.innerHTML = "";
        root.appendChild(
          el("p", { class: "exercise-grid__empty", text: "Übungen konnten nicht geladen werden: " + err.message })
        );
      });
  }

  function groupByLevel(exercises) {
    const groups = {};
    LEVELS.forEach((l) => (groups[l.id] = []));
    exercises.forEach((ex) => {
      const levelId = groups[ex.level] ? ex.level : LEVELS[0].id;
      groups[levelId].push(ex);
    });
    return groups;
  }

  function setupOverview(exercises, root, tabsRoot, bannerRoot, exercisePagePath) {
    const groups = groupByLevel(exercises);
    const nonEmptyLevels = LEVELS.filter((l) => groups[l.id].length > 0);
    const levels = nonEmptyLevels.length ? nonEmptyLevels : LEVELS;

    let activeLevel = (location.hash || "").replace("#", "");
    if (!levels.some((l) => l.id === activeLevel)) activeLevel = levels[0].id;

    function allDone(list) {
      return list.length > 0 && window.ExcelFloProgress && list.every((ex) => window.ExcelFloProgress.isCompleted(ex.id));
    }

    function renderTabs() {
      if (!tabsRoot) return;
      tabsRoot.innerHTML = "";
      levels.forEach((l) => {
        const count = groups[l.id].length;
        const doneCount = window.ExcelFloProgress
          ? groups[l.id].filter((ex) => window.ExcelFloProgress.isCompleted(ex.id)).length
          : 0;
        const btn = el(
          "button",
          { type: "button", class: "level-tab" + (l.id === activeLevel ? " is-active" : "") },
          [
            el("span", { text: l.label }),
            count ? el("span", { class: "level-tab__count", text: doneCount + "/" + count }) : null,
          ]
        );
        btn.addEventListener("click", () => {
          activeLevel = l.id;
          location.hash = l.id;
          renderTabs();
          renderList();
        });
        tabsRoot.appendChild(btn);
      });
    }

    function renderBanner() {
      if (!bannerRoot) return;
      bannerRoot.innerHTML = "";
      bannerRoot.classList.remove("is-visible");
      if (!window.ExcelFloProgress) return;

      if (levels.every((l) => allDone(groups[l.id]))) {
        bannerRoot.appendChild(el("p", { text: "🏆 Geschafft! Du hast alle Übungen in allen Stufen abgeschlossen." }));
        bannerRoot.classList.add("is-visible");
      } else if (allDone(groups[activeLevel])) {
        const label = LEVEL_LABELS[activeLevel] || activeLevel;
        bannerRoot.appendChild(el("p", { text: "✅ Stufe „" + label + "“ abgeschlossen! Wechsle oben zu einer weiteren Stufe." }));
        bannerRoot.classList.add("is-visible");
      }
    }

    function renderList() {
      renderOverview(root, groups[activeLevel] || [], exercisePagePath);
      renderBanner();
    }

    window.addEventListener("hashchange", () => {
      const h = (location.hash || "").replace("#", "");
      if (levels.some((l) => l.id === h) && h !== activeLevel) {
        activeLevel = h;
        renderTabs();
        renderList();
      }
    });

    // Fortschritt kann im Hintergrund von einem anderen Gerät nachgeladen werden (progress.js) –
    // Übersicht dann live aktualisieren.
    document.addEventListener("excelflo:progress-synced", () => {
      renderTabs();
      renderList();
    });

    renderTabs();
    renderList();
  }

  function renderOverview(root, exercises, exercisePagePath) {
    root.innerHTML = "";

    if (!exercises || exercises.length === 0) {
      root.appendChild(el("p", { class: "exercise-grid__empty", text: "In dieser Stufe sind noch keine Übungen verfügbar." }));
      return;
    }

    exercises.forEach((ex) => {
      const done = window.ExcelFloProgress && window.ExcelFloProgress.isCompleted(ex.id);

      const card = el(
        "a",
        { class: "exercise-card" + (done ? " is-done" : ""), href: exercisePagePath + "?id=" + encodeURIComponent(ex.id) },
        [
          el("div", { class: "exercise-card__badges" }, [
            el("span", { class: "badge badge--category", text: formatCategoryLabel(ex.category) }),
            done ? el("span", { class: "badge badge--done", text: "✓ erledigt" }) : null,
          ]),
          el("h3", { text: ex.title }),
          el("p", { text: ex.description || "" }),
          el("span", { class: "exercise-card__cta", text: done ? "Nochmal üben →" : "Übung starten →" }),
        ]
      );

      root.appendChild(card);
    });
  }

  /* ---------------- Übungsseite (generisch, lädt Übung anhand ?id=) ---------------- */

  function initExercise() {
    const root = document.getElementById("exercise-root");
    if (!root) return;

    const id = new URLSearchParams(location.search).get("id");
    if (!id) {
      root.textContent = "Keine Übungs-ID angegeben (erwartet: ?id=... in der URL).";
      return;
    }

    const exercisesDir = root.dataset.exercisesDir || "../assets/exercises/";
    const exercisePagePath = root.dataset.exercisePage || "uebung.html";

    Promise.all([
      fetch(exercisesDir + id + ".json", { cache: "no-cache" }).then((res) => {
        if (!res.ok) throw new Error("Übung „" + id + "“ konnte nicht geladen werden (" + res.status + ")");
        return res.json();
      }),
      fetch(exercisesDir + "manifest.json", { cache: "no-cache" })
        .then((res) => (res.ok ? res.json() : []))
        .catch(() => []),
    ])
      .then(([data, manifest]) => renderExercise(root, data, manifest, exercisePagePath))
      .catch((err) => {
        root.textContent = "Fehler beim Laden der Übung: " + err.message;
      });
  }

  function computeNextExercise(manifest, currentId) {
    if (!manifest || !manifest.length) return { next: null, levelDone: false, allDone: false };
    const current = manifest.find((ex) => ex.id === currentId);
    const level = current ? current.level : null;
    const sameLevel = manifest.filter((ex) => ex.level === level);
    const idx = sameLevel.findIndex((ex) => ex.id === currentId);

    const doneIds = window.ExcelFloProgress ? window.ExcelFloProgress.getCompletedIds() : [];
    const doneSet = new Set(doneIds.concat([currentId]));

    let next = null;
    for (let i = idx + 1; i < sameLevel.length; i++) {
      if (!doneSet.has(sameLevel[i].id)) {
        next = sameLevel[i];
        break;
      }
    }
    if (!next) next = sameLevel.find((ex) => !doneSet.has(ex.id)) || null;

    const levelDone = sameLevel.length > 0 && sameLevel.every((ex) => doneSet.has(ex.id));
    const allDone = manifest.every((ex) => doneSet.has(ex.id));
    return { next, levelDone, allDone, levelLabel: LEVEL_LABELS[level] || level };
  }

  function renderExercise(root, data, manifest, exercisePagePath) {
    document.title = "Excel.Flo – " + data.title;

    root.innerHTML = "";

    root.appendChild(
      el("div", { class: "exercise-header" }, [
        el("div", { class: "exercise-header__badges" }, [
          data.level ? el("span", { class: "badge badge--level", text: LEVEL_LABELS[data.level] || data.level }) : null,
          el("span", { class: "badge badge--category", text: formatCategoryLabel(data.category) }),
        ]),
        el("h1", { text: data.title }),
      ])
    );

    if (data.task) {
      const stepsList = (data.task.steps || []).map((step) => el("li", { text: step }));
      root.appendChild(
        el("div", { class: "exercise-task" }, [
          data.task.intro ? el("p", { class: "exercise-task__intro", text: data.task.intro }) : null,
          stepsList.length ? el("ol", { class: "exercise-task__steps" }, stepsList) : null,
        ])
      );
    }

    const sheet = createSheet(data.grid);
    root.appendChild(sheet.node);

    root.appendChild(
      el("div", { class: "exercise-actions" }, [
        el("button", { class: "btn btn--primary", type: "button", id: "btn-check", text: "Prüfen" }),
        el("button", { class: "btn btn--secondary", type: "button", id: "btn-reset", text: "Zurücksetzen" }),
      ])
    );

    const feedback = el("div", { class: "exercise-feedback", id: "exercise-feedback" });
    root.appendChild(feedback);

    if ((data.hints && data.hints.length) || data.explanation) {
      const hintItems = (data.hints || []).map((hint) => el("li", { text: hint }));
      const solutionBox = data.solution
        ? el("div", { class: "exercise-solution", id: "exercise-solution", text: data.solution })
        : null;

      const details = el("details", { class: "exercise-hints" }, [
        el("summary", { text: "Tipps anzeigen" }),
        hintItems.length ? el("ol", {}, hintItems) : null,
        data.solution
          ? el("button", {
              class: "btn btn--secondary",
              type: "button",
              id: "btn-solution",
              text: "Lösung anzeigen",
            })
          : null,
        solutionBox,
        data.explanation
          ? el("p", { class: "exercise-hints__explanation", text: data.explanation })
          : null,
      ]);

      root.appendChild(details);

      if (data.solution) {
        details.querySelector("#btn-solution").addEventListener("click", () => {
          solutionBox.classList.toggle("is-visible");
        });
      }
    }

    const context = { exerciseData: data, manifest, exercisePagePath };
    document.getElementById("btn-check").addEventListener("click", () => checkExercise(sheet, feedback, context));
    document.getElementById("btn-reset").addEventListener("click", () => resetExercise(sheet, feedback));
  }

  function formatValue(value, format) {
    if (value === undefined || value === null) return "";
    if (format === "currency" && typeof value === "number") {
      return value.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
    }
    if (format === "currency0" && typeof value === "number") {
      return value.toLocaleString("de-DE", { style: "currency", currency: "EUR", minimumFractionDigits: 0, maximumFractionDigits: 0 });
    }
    if (format === "percent" && typeof value === "number") {
      return value.toLocaleString("de-DE", { style: "percent", minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }
    return String(value);
  }

  /* ---------------- Tabellen-/Zell-Engine (Navigation, Formeln, Ausfüllen) ---------------- */

  const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
  const REF_TOKEN_RE = /\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?/;

  function tokenizeFormula(text) {
    // Zerlegt eine Formel in Tokens: Zellbezüge/Bereiche, Funktionsnamen, Zahlen, Rest.
    const tokens = [];
    const re = new RegExp(
      "(" + REF_TOKEN_RE.source + ")|([A-Za-zÄÖÜäöü_][A-Za-z0-9ÄÖÜäöü_.]*)(?=\\()|(-?\\d+(?:[.,]\\d+)?)|([^A-Za-z0-9]|[A-Za-z0-9])",
      "g"
    );
    let m;
    while ((m = re.exec(text)) !== null) {
      if (m[1]) tokens.push({ text: m[1], type: "ref" });
      else if (m[2]) tokens.push({ text: m[2], type: "func" });
      else if (m[3]) tokens.push({ text: m[3], type: "number" });
      else tokens.push({ text: m[4], type: "other" });
    }
    return tokens;
  }

  function refColorMap(tokens) {
    const map = {};
    let next = 0;
    tokens.forEach((t) => {
      if (t.type !== "ref") return;
      const key = t.text.toUpperCase();
      if (!(key in map)) {
        map[key] = REF_COLORS[next % REF_COLORS.length];
        next++;
      }
    });
    return map;
  }

  function renderFormulaMarkup(text) {
    if (!text.startsWith("=")) {
      return { html: escapeHtml(text), colorMap: {} };
    }
    const tokens = tokenizeFormula(text);
    const colorMap = refColorMap(tokens);
    const html = tokens
      .map((t) => {
        const safe = escapeHtml(t.text);
        if (t.type === "ref") return '<span style="color:' + colorMap[t.text.toUpperCase()] + '">' + safe + "</span>";
        if (t.type === "func") return '<span class="tok-func">' + safe + "</span>";
        if (t.type === "number") return '<span class="tok-number">' + safe + "</span>";
        return safe;
      })
      .join("");
    return { html, colorMap };
  }

  // Erweitert einen Zellbezug/Bereich um rowDelta Zeilen / colDelta Spalten; $-fixierte Teile bleiben unverändert.
  function shiftRefToken(token, rowDelta, colDelta) {
    return token.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g, (m, colDollar, colLetters, rowDollar, rowNum) => {
      let newCol = colLetters;
      if (!colDollar && colDelta) {
        const idx = colIndexFromLetters(colLetters) + colDelta;
        if (idx < 0) return m;
        newCol = colLetter(idx);
      }
      let newRow = rowNum;
      if (!rowDollar && rowDelta) {
        const parsed = parseInt(rowNum, 10) + rowDelta;
        if (parsed < 1) return m;
        newRow = String(parsed);
      }
      return colDollar + newCol + rowDollar + newRow;
    });
  }

  function shiftFormula(text, rowDelta, colDelta) {
    if (!text.startsWith("=")) return text;
    return tokenizeFormula(text)
      .map((t) => (t.type === "ref" ? shiftRefToken(t.text, rowDelta, colDelta || 0) : t.text))
      .join("");
  }

  // F4: zyklisiert $-Fixierung eines Bezugs/Bereichs: keine -> beide -> nur Zeile -> nur Spalte -> keine.
  function cycleRefDollars(token) {
    const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g;
    const first = re.exec(token);
    if (!first) return token;
    const curCol = !!first[1];
    const curRow = !!first[3];
    const states = [
      [false, false],
      [true, true],
      [false, true],
      [true, false],
    ];
    const idx = states.findIndex(([c, r]) => c === curCol && r === curRow);
    const [newCol, newRow] = states[(idx + 1) % states.length];
    return token.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g, (m, cd, letters, rd, num) => {
      return (newCol ? "$" : "") + letters + (newRow ? "$" : "") + num;
    });
  }

  // Ermittelt, in welcher Funktion und welchem Argument sich der Cursor gerade befindet (für den Argument-Tooltip).
  function findFunctionContext(text, caret) {
    const stack = [];
    let name = "";
    for (let i = 0; i < caret; i++) {
      const ch = text[i];
      if (/[A-Za-zÄÖÜäöü]/.test(ch)) {
        name += ch;
        continue;
      }
      if (ch === "(") {
        stack.push({ name: name.toUpperCase(), argIndex: 0 });
      } else if (ch === ")") {
        stack.pop();
      } else if ((ch === ";" || ch === ",") && stack.length) {
        stack[stack.length - 1].argIndex++;
      }
      name = "";
    }
    if (!stack.length) return null;
    const top = stack[stack.length - 1];
    if (!FUNCTION_SIGNATURES[top.name]) return null;
    return top;
  }

  function getCaretOffset(container) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return 0;
    const range = sel.getRangeAt(0);
    if (!container.contains(range.startContainer)) return 0;
    const pre = range.cloneRange();
    pre.selectNodeContents(container);
    pre.setEnd(range.endContainer, range.endOffset);
    return pre.toString().length;
  }

  function setCaretOffset(container, offset) {
    const range = document.createRange();
    const sel = window.getSelection();
    let remaining = offset;
    let found = false;

    (function walk(node) {
      if (found) return;
      if (node.nodeType === Node.TEXT_NODE) {
        const len = node.textContent.length;
        if (remaining <= len) {
          range.setStart(node, remaining);
          range.collapse(true);
          found = true;
        } else {
          remaining -= len;
        }
      } else {
        for (let i = 0; i < node.childNodes.length; i++) {
          walk(node.childNodes[i]);
          if (found) break;
        }
      }
    })(container);

    if (!found) {
      range.selectNodeContents(container);
      range.collapse(false);
    }
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function createSheet(grid) {
    const cols = grid.cols;
    const rowCount = grid.rowCount;
    const defs = grid.cells || {};

    const cellEls = {}; // ref -> td
    const inputEntries = {}; // ref -> { el, td, answer }
    const colHeadEls = {};
    const rowHeadEls = {};

    let selectedRef = null;
    let selectionAnchor = null; // Startzelle einer Mehrfachauswahl (Shift+Pfeiltaste/Klick); === selectedRef bei Einzelauswahl
    let rangeSelectedRefs = []; // aktuell hervorgehobene Zellen einer Mehrfachauswahl (ohne die aktive Zelle)
    let editingRef = null;
    let highlightedRefs = [];
    let keyPoint = null; // { before, after, anchor, current } – Pfeiltasten-Referenzierung beim Formel-Schreiben

    const nameBox = el("span", { class: "sheet-toolbar__namebox", text: "" });
    const contentPreview = el("span", { class: "sheet-toolbar__content" });
    const toolbar = el("div", { class: "sheet-toolbar" }, [
      nameBox,
      el("span", { class: "sheet-toolbar__divider" }),
      el("span", { class: "sheet-toolbar__fx", text: "fx" }),
      contentPreview,
    ]);

    const headRow = el("tr", {}, [el("th", { class: "row-head", text: "" })]);
    cols.forEach((c, i) => {
      const th = el("th", { text: c });
      colHeadEls[i] = th;
      headRow.appendChild(th);
    });
    const thead = el("thead", {}, [headRow]);
    const tbody = el("tbody");

    for (let r = 1; r <= rowCount; r++) {
      const rowHeadTd = el("td", { class: "row-head", text: String(r) });
      rowHeadEls[r] = rowHeadTd;
      const rowCells = [rowHeadTd];

      cols.forEach((col) => {
        const ref = col + r;
        const cellDef = defs[ref];
        const td = buildCell(ref, cellDef);
        cellEls[ref] = td;
        rowCells.push(td);
      });

      tbody.appendChild(el("tr", {}, rowCells));
    }

    const table = el("table", { class: "sheet" }, [thead, tbody]);
    const scrollArea = el("div", { class: "sheet-scroll" }, [table]);
    const argHint = el("div", { class: "formula-hint" });
    const wrap = el("div", { class: "sheet-wrap" }, [toolbar, scrollArea, argHint]);
    wrap.tabIndex = 0;

    function buildCell(ref, cellDef) {
      if (!cellDef) {
        return el("td", { class: "cell--empty", "data-ref": ref });
      }

      if (cellDef.type === "input") {
        const content = el("div", {
          class: "cell-editable",
          "data-ref": ref,
          spellcheck: "false",
          autocomplete: "off",
        });
        content.contentEditable = "false";

        const handle = el("span", { class: "fill-handle" });
        const td = el("td", { class: "cell--input", "data-ref": ref }, [content, handle]);

        inputEntries[ref] = { el: content, td, answer: cellDef.answer || {}, beforeEdit: "" };

        content.addEventListener("input", () => {
          keyPoint = null;
          handleContentChanged(ref);
        });
        content.addEventListener("keydown", (e) => handleEditKeydown(e, ref));
        content.addEventListener("blur", () => {
          if (editingRef === ref) commitEdit();
        });
        content.addEventListener("paste", (e) => {
          e.preventDefault();
          const text = (e.clipboardData || window.clipboardData).getData("text/plain").split("\n")[0];
          document.execCommand && document.execCommand("insertText", false, text);
        });

        handle.addEventListener("mousedown", (e) => startFillDrag(e, ref));

        return td;
      }

      const displayValue = formatValue(cellDef.value, cellDef.format);
      const cls = cellDef.type === "header" ? "cell--header" : "cell--data";
      return el("td", { class: cls, "data-ref": ref, text: displayValue });
    }

    /* ---- Auswahl & Navigation ---- */

    function refRowCol(ref) {
      const col = ref.match(/[A-Za-z]+/)[0];
      const row = parseInt(ref.match(/\d+/)[0], 10);
      return { col, row };
    }

    function cellText(ref) {
      if (inputEntries[ref]) return inputEntries[ref].el.textContent;
      const def = defs[ref];
      return def ? formatValue(def.value, def.format) : "";
    }

    // Für den Formel-Auswerter: liefert den rohen Zellwert (Zahl/Text), nicht die formatierte Anzeige.
    function getCellValue(ref) {
      const def = defs[ref];
      if (def && def.value !== undefined) return def.value;
      const entry = inputEntries[ref];
      if (entry) {
        const raw = entry.el.textContent.trim();
        if (raw === "") return undefined;
        const num = parseGermanNumber(raw);
        return num !== null ? num : raw;
      }
      return undefined;
    }

    function clearRangeHighlight() {
      rangeSelectedRefs.forEach((r) => {
        if (cellEls[r]) cellEls[r].classList.remove("is-range-selected");
      });
      rangeSelectedRefs = [];
    }

    function cellsInRange(r1, r2) {
      const a = refRowCol(r1);
      const b = refRowCol(r2);
      const c1 = cols.indexOf(a.col);
      const c2 = cols.indexOf(b.col);
      const colLo = Math.min(c1, c2);
      const colHi = Math.max(c1, c2);
      const rowLo = Math.min(a.row, b.row);
      const rowHi = Math.max(a.row, b.row);
      const refs = [];
      for (let ci = colLo; ci <= colHi; ci++) {
        for (let ri = rowLo; ri <= rowHi; ri++) {
          if (cols[ci]) refs.push(cols[ci] + ri);
        }
      }
      return refs;
    }

    function select(ref) {
      if (!cellEls[ref]) return;
      if (editingRef && editingRef !== ref) commitEdit();
      clearRangeHighlight();

      if (selectedRef && cellEls[selectedRef]) cellEls[selectedRef].classList.remove("is-selected");
      const prev = selectedRef ? refRowCol(selectedRef) : null;
      if (prev) {
        const prevColIdx = cols.indexOf(prev.col);
        if (colHeadEls[prevColIdx]) colHeadEls[prevColIdx].classList.remove("is-active");
        if (rowHeadEls[prev.row]) rowHeadEls[prev.row].classList.remove("is-active");
      }

      selectedRef = ref;
      selectionAnchor = ref;
      cellEls[ref].classList.add("is-selected");
      const { col, row } = refRowCol(ref);
      const colIdx = cols.indexOf(col);
      if (colHeadEls[colIdx]) colHeadEls[colIdx].classList.add("is-active");
      if (rowHeadEls[row]) rowHeadEls[row].classList.add("is-active");

      nameBox.textContent = ref;
      contentPreview.textContent = cellText(ref);
    }

    // Erweitert die Auswahl von selectionAnchor bis newRef (Shift+Pfeiltaste/Klick) – wie in Excel.
    function extendSelectionTo(newRef) {
      if (!cellEls[newRef] || !selectionAnchor) return;
      if (cellEls[selectedRef]) cellEls[selectedRef].classList.remove("is-selected");
      clearRangeHighlight();

      selectedRef = newRef;
      const rangeRefs = cellsInRange(selectionAnchor, newRef);
      rangeRefs.forEach((r) => {
        if (r === newRef || !cellEls[r]) return;
        cellEls[r].classList.add("is-range-selected");
        rangeSelectedRefs.push(r);
      });
      cellEls[newRef].classList.add("is-selected");

      nameBox.textContent = rangeRefs.length > 1 ? selectionAnchor + ":" + newRef : newRef;
      contentPreview.textContent = cellText(newRef);
    }

    function moveSelection(dRow, dCol) {
      const base = selectedRef || cols[0] + "1";
      const { col, row } = refRowCol(base);
      const colIdx = Math.min(Math.max(cols.indexOf(col) + dCol, 0), cols.length - 1);
      const newRow = Math.min(Math.max(row + dRow, 1), rowCount);
      select(cols[colIdx] + newRow);
      wrap.focus({ preventScroll: true });
    }

    // Shift+Pfeiltaste: Auswahl wie in Excel zu einem Zellbereich erweitern statt zu verschieben.
    function extendSelection(dRow, dCol) {
      const base = selectedRef || cols[0] + "1";
      const { col, row } = refRowCol(base);
      const colIdx = Math.min(Math.max(cols.indexOf(col) + dCol, 0), cols.length - 1);
      const newRow = Math.min(Math.max(row + dRow, 1), rowCount);
      if (!selectionAnchor) selectionAnchor = base;
      extendSelectionTo(cols[colIdx] + newRow);
      wrap.focus({ preventScroll: true });
    }

    /* ---- Bearbeiten ---- */

    function handleContentChanged(ref) {
      const entry = inputEntries[ref];
      if (!entry) return;
      entry.td.classList.remove("is-correct", "is-wrong");

      const offset = getCaretOffset(entry.el);
      const text = entry.el.textContent;
      const { html, colorMap } = renderFormulaMarkup(text);
      entry.el.innerHTML = html;
      if (document.activeElement === entry.el) setCaretOffset(entry.el, offset);

      clearRefHighlights();
      Object.keys(colorMap).forEach((refKey) => applyRefHighlight(refKey, colorMap[refKey]));

      if (selectedRef === ref) contentPreview.textContent = text;

      updateArgHint(ref, text, offset);
    }

    function updateArgHint(ref, text, caret) {
      const ctx = text.startsWith("=") ? findFunctionContext(text, caret) : null;
      if (!ctx) {
        argHint.classList.remove("is-visible");
        return;
      }
      const args = FUNCTION_SIGNATURES[ctx.name];
      const activeIdx = Math.min(ctx.argIndex, args.length - 1);
      argHint.innerHTML =
        '<span class="tok-func">' +
        escapeHtml(ctx.name) +
        "</span>(" +
        args
          .map((a, i) => '<span class="' + (i === activeIdx ? "is-active" : "") + '">' + escapeHtml(a) + "</span>")
          .join("; ") +
        ")";
      argHint.classList.add("is-visible");

      const td = cellEls[ref];
      if (td) {
        const wrapRect = wrap.getBoundingClientRect();
        const tdRect = td.getBoundingClientRect();
        let left = tdRect.left - wrapRect.left;
        argHint.style.left = left + "px";
        argHint.style.top = tdRect.bottom - wrapRect.top + 4 + "px";

        const hintWidth = argHint.getBoundingClientRect().width;
        const maxLeft = wrapRect.width - hintWidth - 4;
        if (left > maxLeft) argHint.style.left = Math.max(4, maxLeft) + "px";
      }
    }

    function clearRefHighlights() {
      highlightedRefs.forEach((r) => {
        if (cellEls[r]) cellEls[r].style.boxShadow = "";
      });
      highlightedRefs = [];
    }

    function applyRefHighlight(refToken, color) {
      const m = refToken.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?$/);
      if (!m) return;
      const c1 = colIndexFromLetters(m[2]);
      const r1 = parseInt(m[4], 10);
      const c2 = m[6] ? colIndexFromLetters(m[6]) : c1;
      const r2 = m[8] ? parseInt(m[8], 10) : r1;
      const colLo = Math.min(c1, c2),
        colHi = Math.max(c1, c2);
      const rowLo = Math.min(r1, r2),
        rowHi = Math.max(r1, r2);

      for (let ci = colLo; ci <= colHi; ci++) {
        if (!cols[ci]) continue;
        for (let ri = rowLo; ri <= rowHi; ri++) {
          const ref = cols[ci] + ri;
          if (cellEls[ref]) {
            cellEls[ref].style.boxShadow = "inset 0 0 0 2px " + color;
            highlightedRefs.push(ref);
          }
        }
      }
    }

    function enterEditMode(ref, replacementChar) {
      const entry = inputEntries[ref];
      if (!entry) return;
      if (editingRef && editingRef !== ref) commitEdit();
      if (selectedRef !== ref) select(ref);

      entry.beforeEdit = entry.el.textContent;
      if (typeof replacementChar === "string") entry.el.textContent = replacementChar;

      entry.el.contentEditable = "true";
      editingRef = ref;
      keyPoint = null;
      entry.el.focus();
      setCaretOffset(entry.el, entry.el.textContent.length);
      handleContentChanged(ref);
    }

    function commitEdit() {
      if (!editingRef) return;
      const entry = inputEntries[editingRef];
      if (entry) entry.el.contentEditable = "false";
      editingRef = null;
      keyPoint = null;
      clearRefHighlights();
      argHint.classList.remove("is-visible");
    }

    function cancelEdit() {
      if (!editingRef) return;
      const entry = inputEntries[editingRef];
      if (entry) {
        entry.el.textContent = entry.beforeEdit;
        handleContentChanged(editingRef);
      }
      commitEdit();
    }

    function handleEditKeydown(e, ref) {
      const entry = inputEntries[ref];
      // Verhindert, dass Tab/Enter/Escape/F4 etc. zum Container hochblubbern und dort
      // (mit inzwischen geändertem editingRef/selectedRef) ein zweites Mal ausgelöst werden.
      e.stopPropagation();

      if (e.key === "Tab") {
        e.preventDefault();
        commitEdit();
        moveSelection(0, e.shiftKey ? -1 : 1);
      } else if (e.key === "Enter") {
        e.preventDefault();
        commitEdit();
        moveSelection(1, 0);
      } else if (e.key === "Escape") {
        e.preventDefault();
        cancelEdit();
        wrap.focus({ preventScroll: true });
      } else if (
        ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key) &&
        entry.el.textContent.startsWith("=") &&
        getCaretOffset(entry.el) === entry.el.textContent.length
      ) {
        // Pfeiltaste am Formel-Ende: peilt statt Cursor-Bewegung eine Nachbarzelle (oder mit
        // Shift einen Zellbereich) als Bezug an.
        e.preventDefault();
        if (!keyPoint) {
          keyPoint = { before: entry.el.textContent, after: "", anchor: ref, current: ref };
        }
        const base = refRowCol(keyPoint.current);
        const baseColIdx = cols.indexOf(base.col);
        let newColIdx = baseColIdx;
        let newRow = base.row;
        if (e.key === "ArrowUp") newRow = Math.max(1, base.row - 1);
        else if (e.key === "ArrowDown") newRow = Math.min(rowCount, base.row + 1);
        else if (e.key === "ArrowLeft") newColIdx = Math.max(0, baseColIdx - 1);
        else if (e.key === "ArrowRight") newColIdx = Math.min(cols.length - 1, baseColIdx + 1);
        const newCurrent = cols[newColIdx] + newRow;

        keyPoint.current = newCurrent;
        if (!e.shiftKey) keyPoint.anchor = newCurrent;

        const insertText = rangeRefText(keyPoint.anchor, keyPoint.current);
        entry.el.textContent = keyPoint.before + insertText + keyPoint.after;
        handleContentChanged(ref);
        setCaretOffset(entry.el, keyPoint.before.length + insertText.length);
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        commitEdit();
        moveSelection(-1, 0);
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        commitEdit();
        moveSelection(1, 0);
      } else if (e.key === "ArrowLeft") {
        if (getCaretOffset(entry.el) === 0) {
          e.preventDefault();
          commitEdit();
          moveSelection(0, -1);
        }
      } else if (e.key === "ArrowRight") {
        if (getCaretOffset(entry.el) === entry.el.textContent.length) {
          e.preventDefault();
          commitEdit();
          moveSelection(0, 1);
        }
      } else if (e.key === "F4") {
        e.preventDefault();
        const text = entry.el.textContent;
        if (!text.startsWith("=")) return;
        const caret = getCaretOffset(entry.el);
        const tokens = tokenizeFormula(text);
        let pos = 0;
        let target = null;
        for (const t of tokens) {
          const start = pos;
          const end = pos + t.text.length;
          if (t.type === "ref" && caret > start && caret <= end) {
            target = { start, end, token: t };
            break;
          }
          pos = end;
        }
        if (!target) return;
        const newToken = cycleRefDollars(target.token.text);
        entry.el.textContent = text.slice(0, target.start) + newToken + text.slice(target.end);
        handleContentChanged(ref);
        setCaretOffset(entry.el, target.start + newToken.length);
      }
    }

    /* ---- Kopieren & Einfügen (Strg+C / Strg+V) ---- */

    let clipboard = null; // { ref, text }

    function clearCopiedVisual() {
      Object.values(cellEls).forEach((td) => td.classList.remove("is-copied"));
    }

    function copySelected() {
      if (!selectedRef) return;
      const text = cellText(selectedRef);
      clipboard = { ref: selectedRef, text };
      clearCopiedVisual();
      cellEls[selectedRef].classList.add("is-copied");
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).catch(() => {});
      }
    }

    function applyPaste(text, fromRef) {
      // Bei Mehrfachauswahl (Shift+Pfeiltaste/Klick) wird in jede Zelle des markierten
      // Bereichs eingefügt, jeweils mit angepassten relativen Bezügen – wie in Excel.
      const destRefs =
        selectionAnchor && selectionAnchor !== selectedRef ? cellsInRange(selectionAnchor, selectedRef) : [selectedRef];

      destRefs.forEach((destRef) => {
        const entry = inputEntries[destRef];
        if (!entry) return;
        let finalText = text;
        if (fromRef) {
          const src = refRowCol(fromRef);
          const dst = refRowCol(destRef);
          const rowDelta = dst.row - src.row;
          const colDelta = cols.indexOf(dst.col) - cols.indexOf(src.col);
          finalText = shiftFormula(text, rowDelta, colDelta);
        }
        entry.el.textContent = finalText;
        handleContentChanged(destRef);
      });
      clearCopiedVisual();
    }

    function pasteIntoSelected() {
      if (!selectedRef) return;

      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard
          .readText()
          .then((text) => {
            if (clipboard && text === clipboard.text) applyPaste(text, clipboard.ref);
            else if (text) applyPaste(text, null);
            else if (clipboard) applyPaste(clipboard.text, clipboard.ref);
          })
          .catch(() => {
            if (clipboard) applyPaste(clipboard.text, clipboard.ref);
          });
      } else if (clipboard) {
        applyPaste(clipboard.text, clipboard.ref);
      }
    }

    function handleContainerKeydown(e) {
      if (editingRef) return; // wird von handleEditKeydown behandelt
      if (!selectedRef) return;

      const ctrlOrCmd = e.ctrlKey || e.metaKey;

      if (e.key === "ArrowUp") {
        e.preventDefault();
        e.shiftKey ? extendSelection(-1, 0) : moveSelection(-1, 0);
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        e.shiftKey ? extendSelection(1, 0) : moveSelection(1, 0);
      } else if (e.key === "ArrowLeft") {
        e.preventDefault();
        e.shiftKey ? extendSelection(0, -1) : moveSelection(0, -1);
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        e.shiftKey ? extendSelection(0, 1) : moveSelection(0, 1);
      } else if (e.key === "Tab") {
        e.preventDefault();
        moveSelection(0, e.shiftKey ? -1 : 1);
      } else if (ctrlOrCmd && (e.key === "c" || e.key === "C")) {
        e.preventDefault();
        copySelected();
      } else if (ctrlOrCmd && (e.key === "v" || e.key === "V")) {
        e.preventDefault();
        pasteIntoSelected();
      } else if (e.key === "Enter" || e.key === "F2") {
        if (inputEntries[selectedRef]) {
          e.preventDefault();
          enterEditMode(selectedRef);
        }
      } else if (e.key === "Delete" || e.key === "Backspace") {
        if (inputEntries[selectedRef]) {
          e.preventDefault();
          inputEntries[selectedRef].el.textContent = "";
          handleContentChanged(selectedRef);
          contentPreview.textContent = "";
        }
      } else if (e.key === "Escape") {
        clearCopiedVisual();
      } else if (e.key.length === 1 && !ctrlOrCmd && !e.altKey) {
        if (inputEntries[selectedRef]) {
          e.preventDefault();
          enterEditMode(selectedRef, e.key);
        }
      }
    }

    wrap.addEventListener("keydown", handleContainerKeydown);

    /* ---- Zellen anklicken, um sie als Bezug in eine Formel einzufügen ("Point-Modus") ---- */

    let pointDrag = null; // { editRef, before, after, anchorRef, currentHover }
    let suppressNextClick = false;

    function rangeRefText(r1, r2) {
      if (r1 === r2) return r1;
      const a = refRowCol(r1);
      const b = refRowCol(r2);
      const c1 = cols.indexOf(a.col);
      const c2 = cols.indexOf(b.col);
      const colLo = cols[Math.min(c1, c2)];
      const colHi = cols[Math.max(c1, c2)];
      const rowLo = Math.min(a.row, b.row);
      const rowHi = Math.max(a.row, b.row);
      return colLo + rowLo + ":" + colHi + rowHi;
    }

    function updatePointDrag(hoverRef) {
      const refText = rangeRefText(pointDrag.anchorRef, hoverRef);
      const entry = inputEntries[pointDrag.editRef];
      entry.el.textContent = pointDrag.before + refText + pointDrag.after;
      pointDrag.currentHover = hoverRef;
      handleContentChanged(pointDrag.editRef);
    }

    function onPointMove(e) {
      if (!pointDrag) return;
      const targetEl = document.elementFromPoint(e.clientX, e.clientY);
      const td = targetEl && targetEl.closest ? targetEl.closest("td[data-ref]") : null;
      if (!td) return;
      updatePointDrag(td.dataset.ref);
    }

    function onPointDrop() {
      if (!pointDrag) return;
      const { editRef, before, anchorRef, currentHover } = pointDrag;
      const refText = rangeRefText(anchorRef, currentHover || anchorRef);
      const entry = inputEntries[editRef];
      const newCaret = before.length + refText.length;

      pointDrag = null;
      suppressNextClick = true;
      document.removeEventListener("mousemove", onPointMove);
      document.removeEventListener("mouseup", onPointDrop);

      entry.el.contentEditable = "true";
      editingRef = editRef;
      entry.el.focus();
      setCaretOffset(entry.el, newCaret);
      handleContentChanged(editRef);
    }

    function startPointDrag(clickedRef) {
      if (!editingRef) return false;
      const entry = inputEntries[editingRef];
      if (!entry || !entry.el.textContent.startsWith("=")) return false;
      if (clickedRef === editingRef) return false;

      const caret = getCaretOffset(entry.el);
      const fullText = entry.el.textContent;
      pointDrag = {
        editRef: editingRef,
        before: fullText.slice(0, caret),
        after: fullText.slice(caret),
        anchorRef: clickedRef,
        currentHover: clickedRef,
      };
      updatePointDrag(clickedRef);
      document.addEventListener("mousemove", onPointMove);
      document.addEventListener("mouseup", onPointDrop);
      return true;
    }

    table.addEventListener("mousedown", (e) => {
      const td = e.target.closest("td[data-ref]");
      if (!td || e.target.closest(".fill-handle")) return;
      if (startPointDrag(td.dataset.ref)) {
        e.preventDefault();
      }
    });

    table.addEventListener("click", (e) => {
      if (suppressNextClick) {
        suppressNextClick = false;
        return;
      }
      const td = e.target.closest("td[data-ref]");
      if (!td) return;
      if (e.shiftKey && selectionAnchor && !editingRef) {
        extendSelectionTo(td.dataset.ref);
      } else {
        select(td.dataset.ref);
      }
      if (document.activeElement !== (inputEntries[td.dataset.ref] || {}).el) wrap.focus({ preventScroll: true });
    });

    table.addEventListener("dblclick", (e) => {
      const td = e.target.closest("td[data-ref]");
      if (!td) return;
      const ref = td.dataset.ref;
      if (inputEntries[ref]) enterEditMode(ref);
    });

    /* ---- Ausfüllen per Fill-Handle (vertikal, wie Excel) ---- */

    let fillDrag = null;

    function startFillDrag(e, sourceRef) {
      e.preventDefault();
      e.stopPropagation();
      if (editingRef) commitEdit();
      const { col } = refRowCol(sourceRef);
      fillDrag = { sourceRef, col, previewRefs: [] };

      document.addEventListener("mousemove", onFillMove);
      document.addEventListener("mouseup", onFillDrop);
    }

    function clearFillPreview() {
      if (!fillDrag) return;
      fillDrag.previewRefs.forEach((r) => {
        if (cellEls[r]) cellEls[r].classList.remove("is-fill-preview");
      });
      fillDrag.previewRefs = [];
    }

    function onFillMove(e) {
      if (!fillDrag) return;
      const targetTd = document.elementFromPoint(e.clientX, e.clientY);
      const td = targetTd && targetTd.closest ? targetTd.closest("td[data-ref]") : null;
      if (!td) return;
      const { col, row } = refRowCol(td.dataset.ref);
      if (col !== fillDrag.col) return;

      const { row: sourceRow } = refRowCol(fillDrag.sourceRef);
      clearFillPreview();

      const lo = Math.min(sourceRow, row);
      const hi = Math.max(sourceRow, row);
      for (let r = lo; r <= hi; r++) {
        if (r === sourceRow) continue;
        const ref = col + r;
        if (cellEls[ref]) {
          cellEls[ref].classList.add("is-fill-preview");
          fillDrag.previewRefs.push(ref);
        }
      }
      fillDrag.targetRow = row;
    }

    function onFillDrop() {
      if (!fillDrag) return;
      const { sourceRef, targetRow } = fillDrag;
      const sourceEntry = inputEntries[sourceRef];
      const { col, row: sourceRow } = refRowCol(sourceRef);

      if (sourceEntry && targetRow !== undefined && targetRow !== sourceRow) {
        const sourceText = sourceEntry.el.textContent;
        const lo = Math.min(sourceRow, targetRow);
        const hi = Math.max(sourceRow, targetRow);
        for (let r = lo; r <= hi; r++) {
          if (r === sourceRow) continue;
          const targetRef = col + r;
          const targetEntry = inputEntries[targetRef];
          if (!targetEntry) continue;
          targetEntry.el.textContent = shiftFormula(sourceText, r - sourceRow, 0);
          handleContentChanged(targetRef);
        }
        clearRefHighlights();
      }

      clearFillPreview();
      fillDrag = null;
      document.removeEventListener("mousemove", onFillMove);
      document.removeEventListener("mouseup", onFillDrop);
    }

    function resetSheet() {
      if (editingRef) cancelEdit();
      clearRefHighlights();
      clearCopiedVisual();
      argHint.classList.remove("is-visible");
      Object.values(inputEntries).forEach((entry) => {
        entry.el.textContent = "";
        entry.td.classList.remove("is-correct", "is-wrong");
      });
      select(cols[0] + "1");
      wrap.focus({ preventScroll: true });
    }

    select(cols[0] + "1");

    return {
      node: wrap,
      inputEntries,
      select,
      cellText,
      getCellValue,
      reset: resetSheet,
    };
  }

  /* ---------------- Prüf-Logik (Mustervergleich) ---------------- */

  function parseGermanNumber(raw) {
    const cleaned = raw.trim().replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
    const num = parseFloat(cleaned);
    return Number.isNaN(num) ? null : num;
  }

  function normalizeFormula(raw) {
    return raw.trim().toUpperCase().replace(/\s+/g, "");
  }

  // Prüf-Reihenfolge: 1) direkter Zahlenwert  2) Klartext-Formelliste (acceptedFormulas,
  // normalisiert – $, ;/,, WAHR/FALSCH vs. 1/0 sind dabei egal)  3) Regex-Patterns (Altlast/
  // Spezialfälle)  4) echte Auswertung der eingegebenen Formel gegen die Zelldaten – fängt
  // Formulierungen ab, an die vorher niemand gedacht hat.
  function textMatches(a, b, caseSensitive) {
    if (caseSensitive) return String(a).trim() === String(b).trim();
    return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
  }

  function checkCell(rawInput, answer, getCellValue) {
    const raw = (rawInput || "").trim();
    if (raw === "") return null; // nicht beantwortet

    if (typeof answer.value === "number") {
      const num = parseGermanNumber(raw);
      const tolerance = answer.tolerance !== undefined ? answer.tolerance : 0.01;
      if (num !== null && Math.abs(num - answer.value) < tolerance) return true;
    } else if (typeof answer.value === "string" && !raw.startsWith("=")) {
      if (textMatches(raw, answer.value, answer.caseSensitive)) return true;
    }

    if (window.ExcelFloFormula && answer.acceptedFormulas && answer.acceptedFormulas.length) {
      if (window.ExcelFloFormula.acceptedFormulaMatch(raw, answer.acceptedFormulas)) return true;
    }

    if (answer.patterns && answer.patterns.length) {
      const normalized = normalizeFormula(raw);
      for (const pattern of answer.patterns) {
        try {
          const re = new RegExp(pattern, "i");
          if (re.test(normalized)) return true;
        } catch (e) {
          // ungültiges Pattern in den Übungsdaten – überspringen
        }
      }
    }

    if (window.ExcelFloFormula && getCellValue && raw.startsWith("=") && answer.value !== undefined) {
      const result = window.ExcelFloFormula.evaluate(raw, getCellValue);
      if (!window.ExcelFloFormula.isFormulaError(result)) {
        if (typeof answer.value === "number" && typeof result === "number") {
          const tolerance = answer.tolerance !== undefined ? answer.tolerance : 0.01;
          if (Math.abs(result - answer.value) < tolerance) return true;
        } else if (typeof answer.value === "string" && (typeof result === "string" || typeof result === "boolean")) {
          if (textMatches(result, answer.value, answer.caseSensitive)) return true;
        }
      }
    }

    return false;
  }

  function checkExercise(sheet, feedback, context) {
    const refs = Object.keys(sheet.inputEntries);
    let answered = 0;
    let correct = 0;

    refs.forEach((ref) => {
      const entry = sheet.inputEntries[ref];
      const result = checkCell(entry.el.textContent, entry.answer, sheet.getCellValue);
      entry.td.classList.remove("is-correct", "is-wrong");
      if (result === true) {
        entry.td.classList.add("is-correct");
        answered++;
        correct++;
      } else if (result === false) {
        entry.td.classList.add("is-wrong");
        answered++;
      }
    });

    feedback.classList.remove("is-success", "is-error");
    feedback.innerHTML = "";

    if (answered === 0) {
      feedback.classList.add("is-error");
      feedback.appendChild(el("p", { text: "Bitte trage zuerst eine Antwort ein." }));
      return;
    }

    if (correct === refs.length) {
      feedback.classList.add("is-success");
      const successText = refs.length === 1 ? "Richtig! 🎉" : "Richtig! Alle " + refs.length + " Felder stimmen. 🎉";
      feedback.appendChild(el("p", { text: successText }));

      const data = context && context.exerciseData;
      if (data && data.explanation) {
        feedback.appendChild(el("p", { class: "exercise-feedback__explanation", text: data.explanation }));
      }

      if (data && window.ExcelFloProgress) {
        window.ExcelFloProgress.markCompleted(data.id);
        const nextInfo = computeNextExercise(context.manifest, data.id);

        if (nextInfo.allDone) {
          feedback.appendChild(
            el("p", { class: "exercise-feedback__next", text: "🏆 Du hast alle Übungen in allen Stufen abgeschlossen!" })
          );
        } else if (nextInfo.levelDone) {
          feedback.appendChild(
            el("p", {
              class: "exercise-feedback__next",
              text: "✅ Stufe „" + nextInfo.levelLabel + "“ abgeschlossen!",
            })
          );
        } else if (nextInfo.next) {
          const link = el(
            "a",
            {
              class: "btn btn--primary",
              href: (context.exercisePagePath || "uebung.html") + "?id=" + encodeURIComponent(nextInfo.next.id),
            },
            [document.createTextNode("Nächste Übung: " + nextInfo.next.title + " →")]
          );
          feedback.appendChild(el("p", { class: "exercise-feedback__next" }, [link]));
        }
      }
    } else {
      feedback.classList.add("is-error");
      feedback.appendChild(el("p", { text: correct + " von " + refs.length + " Feldern korrekt. Versuch es weiter!" }));
    }
  }

  function resetExercise(sheet, feedback) {
    sheet.reset();
    feedback.classList.remove("is-success", "is-error");
    feedback.textContent = "";

    const solutionBox = document.getElementById("exercise-solution");
    if (solutionBox) solutionBox.classList.remove("is-visible");
  }

  document.addEventListener("DOMContentLoaded", () => {
    initOverview();
    initExercise();
  });

  window.ExcelFlo = { colLetter, checkCell };
})();
