/* Excel.Flo – gemeinsame Render- und Prüf-Logik
 * Keine Abhängigkeiten, kein Build-Schritt. Läuft auf Übersichts- und Übungsseiten.
 */

(function () {
  "use strict";

  const DIFFICULTY_CLASS = {
    "Leicht": "badge--difficulty-leicht",
    "Mittel": "badge--difficulty-mittel",
    "Schwer": "badge--difficulty-schwer",
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

  /* ---------------- Übersichtsseite ---------------- */

  function initOverview() {
    const root = document.getElementById("exercise-list");
    if (!root) return;

    const manifestPath = root.dataset.manifestPath || "assets/exercises/manifest.json";

    fetch(manifestPath)
      .then((res) => {
        if (!res.ok) throw new Error("Manifest konnte nicht geladen werden (" + res.status + ")");
        return res.json();
      })
      .then((exercises) => renderOverview(root, exercises))
      .catch((err) => {
        root.innerHTML = "";
        root.appendChild(
          el("p", { class: "exercise-grid__empty", text: "Übungen konnten nicht geladen werden: " + err.message })
        );
      });
  }

  function renderOverview(root, exercises) {
    root.innerHTML = "";

    if (!exercises || exercises.length === 0) {
      root.appendChild(el("p", { class: "exercise-grid__empty", text: "Noch keine Übungen verfügbar." }));
      return;
    }

    exercises.forEach((ex) => {
      const diffClass = DIFFICULTY_CLASS[ex.difficulty] || "badge--difficulty-mittel";

      const card = el("a", { class: "exercise-card", href: "uebungen/" + ex.slug + ".html" }, [
        el("div", { class: "exercise-card__badges" }, [
          el("span", { class: "badge badge--category", text: ex.category }),
          el("span", { class: "badge " + diffClass, text: ex.difficulty }),
        ]),
        el("h3", { text: ex.title }),
        el("p", { text: ex.description || "" }),
        el("span", { class: "exercise-card__cta", text: "Übung starten →" }),
      ]);

      root.appendChild(card);
    });
  }

  /* ---------------- Übungsseite ---------------- */

  function initExercise() {
    const root = document.getElementById("exercise-root");
    if (!root) return;

    const exercisePath = root.dataset.exercisePath;
    if (!exercisePath) {
      root.textContent = "Kein data-exercise-path auf #exercise-root gesetzt.";
      return;
    }

    fetch(exercisePath)
      .then((res) => {
        if (!res.ok) throw new Error("Übung konnte nicht geladen werden (" + res.status + ")");
        return res.json();
      })
      .then((data) => renderExercise(root, data))
      .catch((err) => {
        root.textContent = "Fehler beim Laden der Übung: " + err.message;
      });
  }

  function renderExercise(root, data) {
    document.title = "Excel.Flo – " + data.title;

    root.innerHTML = "";

    const diffClass = DIFFICULTY_CLASS[data.difficulty] || "badge--difficulty-mittel";

    root.appendChild(
      el("div", { class: "exercise-header" }, [
        el("div", { class: "exercise-header__badges" }, [
          el("span", { class: "badge badge--category", text: data.category }),
          el("span", { class: "badge " + diffClass, text: data.difficulty }),
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

    const inputRefs = {};
    root.appendChild(buildSheet(data.grid, inputRefs));

    root.appendChild(
      el("div", { class: "exercise-actions" }, [
        el("button", { class: "btn btn--primary", type: "button", id: "btn-check", text: "Prüfen" }),
        el("button", { class: "btn btn--secondary", type: "button", id: "btn-reset", text: "Zurücksetzen" }),
      ])
    );

    const feedback = el("div", { class: "exercise-feedback", id: "exercise-feedback" });
    root.appendChild(feedback);

    if (data.hints && data.hints.length) {
      const hintItems = data.hints.map((hint) => el("li", { text: hint }));
      const solutionBox = data.solution
        ? el("div", { class: "exercise-solution", id: "exercise-solution", text: data.solution })
        : null;

      const details = el("details", { class: "exercise-hints" }, [
        el("summary", { text: "Tipps anzeigen" }),
        el("ol", {}, hintItems),
        data.solution
          ? el("button", {
              class: "btn btn--secondary",
              type: "button",
              id: "btn-solution",
              text: "Lösung anzeigen",
              style: "margin: 0 0 0 0;",
            })
          : null,
        solutionBox,
      ]);

      root.appendChild(details);

      if (data.solution) {
        details.querySelector("#btn-solution").addEventListener("click", () => {
          solutionBox.classList.toggle("is-visible");
        });
      }
    }

    document.getElementById("btn-check").addEventListener("click", () => checkExercise(data, inputRefs, feedback));
    document.getElementById("btn-reset").addEventListener("click", () => resetExercise(inputRefs, feedback));
  }

  function buildSheet(grid, inputRefs) {
    const cols = grid.cols;
    const rowCount = grid.rowCount;
    const cells = grid.cells || {};

    const headRow = el("tr", {}, [
      el("th", { class: "row-head", text: "" }),
      ...cols.map((c) => el("th", { text: c })),
    ]);
    const thead = el("thead", {}, [headRow]);

    const tbody = el("tbody");

    for (let r = 1; r <= rowCount; r++) {
      const rowCells = [el("td", { class: "row-head", text: String(r) })];

      cols.forEach((col) => {
        const ref = col + r;
        const cellDef = cells[ref];
        rowCells.push(buildCell(ref, cellDef, inputRefs));
      });

      tbody.appendChild(el("tr", {}, rowCells));
    }

    return el("div", { class: "sheet-wrap" }, [el("table", { class: "sheet" }, [thead, tbody])]);
  }

  function buildCell(ref, cellDef, inputRefs) {
    if (!cellDef) {
      return el("td", { class: "cell--empty" });
    }

    if (cellDef.type === "input") {
      const input = el("input", { type: "text", "data-ref": ref, autocomplete: "off", spellcheck: "false" });
      const td = el("td", { class: "cell--input" }, [input]);
      inputRefs[ref] = { input, td, answer: cellDef.answer || {} };
      return td;
    }

    const displayValue = formatValue(cellDef.value, cellDef.format);
    const cls = cellDef.type === "header" ? "cell--header" : "cell--data";
    return el("td", { class: cls, text: displayValue });
  }

  function formatValue(value, format) {
    if (value === undefined || value === null) return "";
    if (format === "currency" && typeof value === "number") {
      return value.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
    }
    return String(value);
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

  function checkCell(rawInput, answer) {
    const raw = (rawInput || "").trim();
    if (raw === "") return null; // nicht beantwortet

    if (typeof answer.value === "number") {
      const num = parseGermanNumber(raw);
      const tolerance = answer.tolerance !== undefined ? answer.tolerance : 0.01;
      if (num !== null && Math.abs(num - answer.value) < tolerance) return true;
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
      return false;
    }

    return false;
  }

  function checkExercise(data, inputRefs, feedback) {
    const refs = Object.keys(inputRefs);
    let answered = 0;
    let correct = 0;

    refs.forEach((ref) => {
      const entry = inputRefs[ref];
      const result = checkCell(entry.input.value, entry.answer);
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

    if (answered === 0) {
      feedback.classList.add("is-error");
      feedback.textContent = "Bitte trage zuerst eine Antwort ein.";
      return;
    }

    if (correct === refs.length) {
      feedback.classList.add("is-success");
      feedback.textContent = "Richtig! Alle " + refs.length + " Felder stimmen. 🎉";
    } else {
      feedback.classList.add("is-error");
      feedback.textContent = correct + " von " + refs.length + " Feldern korrekt. Versuch es weiter!";
    }
  }

  function resetExercise(inputRefs, feedback) {
    Object.values(inputRefs).forEach((entry) => {
      entry.input.value = "";
      entry.td.classList.remove("is-correct", "is-wrong");
    });
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
