/* Excel.Flo – Power-Query-Übungsseite
 *
 * Verbindet Aufgabe, Excel-/Power-Query-Nachbau und Prüfung. Geprüft wird das
 * Ergebnis, das per „Schließen & laden“ in Excel geladen wurde – verglichen
 * mit expectedOutput (siehe engine-powerquery.js). Lädt nach engine.js,
 * engine-powerquery.js, -icons, -ui, -excel, -editor und -dialogs.
 */

(function () {
  "use strict";

  function initPowerQueryExercise() {
    const root = document.getElementById("pq-exercise-root");
    if (!root) return;

    if (!window.ExcelFlo || !window.ExcelFlo.el || !window.ExcelFloPQEditor) {
      root.textContent = "Fehler: Die Skripte der Power-Query-Übung wurden nicht vollständig geladen.";
      return;
    }

    const id = new URLSearchParams(location.search).get("id");
    if (!id) {
      root.textContent = "Keine Übungs-ID angegeben (erwartet: ?id=... in der URL).";
      return;
    }

    const exercisesDir = root.dataset.exercisesDir || "../assets/exercises/powerquery/";
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

  function createApp(data, host, windowEl) {
    const PQ = window.ExcelFloPowerQuery;
    const XL = window.ExcelFloPQExcel;
    const ED = window.ExcelFloPQEditor;
    const UI = window.ExcelFloUI;

    const app = {
      data,
      host,
      windowEl,
      source: PQ.tableFromRows(data.inputData || []),
      tableName: data.tableName || "Tabelle1",
      sourceSheetName: data.sheetName || "Tabelle1",
      outputSheetName: data.outputSheetName || "Tabelle2",
      workbookName: data.workbookName || "Mappe1",
      tableCreated: false,
      query: null,
      view: "excel",
      excel: { sheet: data.sheetName || "Tabelle1", sel: { r: 0, c: 0 }, queriesPane: false, message: null },
      editor: null,
      check: null,
    };

    app.render = () => {
      UI.hideTip();
      if (app.view === "editor") ED.render(app);
      else XL.render(app);
    };
    app.hint = (text) => {
      if (app.view === "editor") {
        ED.hint(app, text);
      } else {
        app.excel.message = { text };
        app.render();
      }
    };
    app.openEditor = () => ED.open(app);
    app.ensureQuery = () => ED.ensureQuery(app);
    app.refreshAll = () => ED.refreshAll(app);
    return app;
  }

  function renderExercise(root, data, manifest, exercisePagePath) {
    const shared = window.ExcelFlo;
    const PQ = window.ExcelFloPowerQuery;
    const UI = window.ExcelFloUI;
    const el = shared.el;

    document.title = "Excel.Flo – " + data.title;
    root.innerHTML = "";

    root.appendChild(
      el("div", { class: "exercise-header" }, [
        el("div", { class: "exercise-header__badges" }, [
          data.level ? el("span", { class: "badge badge--level", text: shared.LEVEL_LABELS[data.level] || data.level }) : null,
          el("span", { class: "badge badge--category", text: shared.formatCategoryLabel(data.category) }),
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

    const windowEl = el("div", { class: "xl-stage__window" });
    const stage = el("div", { class: "xl-stage" }, [windowEl]);
    root.appendChild(el("div", { class: "xl-stage-scroll" }, [stage]));

    const app = createApp(data, stage, windowEl);
    UI.bindTooltips(stage);

    root.appendChild(
      el("div", { class: "exercise-actions" }, [
        el("button", { class: "btn btn--primary", type: "button", id: "btn-check", text: "Prüfen" }),
        el("button", { class: "btn btn--secondary", type: "button", id: "btn-reset", text: "Zurücksetzen" }),
      ])
    );

    const feedback = el("div", { class: "exercise-feedback", id: "exercise-feedback" });
    root.appendChild(feedback);

    let solutionBox = null;
    const solutionSteps = Array.isArray(data.solution) ? data.solution : [];
    if ((data.hints && data.hints.length) || data.explanation || solutionSteps.length) {
      const hintItems = (data.hints || []).map((hint) => el("li", { text: hint }));
      if (solutionSteps.length) {
        const item = (path, detail) =>
          el("li", {}, [el("strong", { text: path }), detail ? document.createTextNode(" – " + detail) : null]);
        solutionBox = el("div", { class: "exercise-solution pq-solution", id: "exercise-solution" }, [
          el(
            "ol",
            {},
            [item("Daten › Aus Tabelle/Bereich", "„Tabelle hat Überschriften“ aktiviert lassen")]
              .concat(
                solutionSteps.map((s) => {
                  const action = PQ.ACTIONS[s.action];
                  return item(action ? action.path : s.action, PQ.stepDetail(s));
                })
              )
              .concat([item("Start › Schließen & laden", "Ergebnis erscheint auf einem neuen Blatt")])
          ),
          el("p", {
            class: "pq-solution__note",
            text: "Eine andere Reihenfolge kann ebenfalls richtig sein – geprüft wird nur, ob am Ende die richtige Tabelle in Excel geladen ist.",
          }),
        ]);
      }

      const details = el("details", { class: "exercise-hints" }, [
        el("summary", { text: "Tipps anzeigen" }),
        hintItems.length ? el("ol", {}, hintItems) : null,
        solutionBox ? el("button", { class: "btn btn--secondary", type: "button", id: "btn-solution", text: "Musterlösung anzeigen" }) : null,
        solutionBox,
        data.explanation ? el("p", { class: "exercise-hints__explanation", text: data.explanation }) : null,
      ]);
      root.appendChild(details);
      if (solutionBox) details.querySelector("#btn-solution").addEventListener("click", () => solutionBox.classList.toggle("is-visible"));
    }

    function clearFeedback() {
      feedback.classList.remove("is-success", "is-error");
      feedback.innerHTML = "";
    }

    function fail(message, details) {
      feedback.classList.add("is-error");
      feedback.appendChild(el("p", { text: message }));
      if (details) feedback.appendChild(details);
      shared.showErrorPopup(stage, message);
    }

    app.onLoaded = () => {
      clearFeedback();
    };

    function check() {
      clearFeedback();
      if (app.view === "editor") {
        fail("Der Power Query-Editor ist noch geöffnet. Lade dein Ergebnis zuerst mit „Schließen & laden“ in Excel.");
        return;
      }
      if (!app.query || !app.query.loaded) {
        fail("Es ist noch kein Ergebnis in Excel geladen. Bearbeite die Daten in Power Query und klicke dort auf „Schließen & laden“.");
        return;
      }

      const loaded = app.query.loaded;
      const result = PQ.compareTables(loaded.result, data.expectedOutput || [], data.expectedColumns);
      app.check = result;
      if (loaded.mode === "table") {
        app.excel.sheet = loaded.sheetName;
      }
      app.render();

      if (result.ok) {
        feedback.classList.add("is-success");
        feedback.appendChild(el("p", { text: "Richtig! 🎉 Das geladene Ergebnis entspricht genau der Zieltabelle." }));
        shared.showSuccessPopup(stage);
        shared.appendCompletionFeedback(feedback, data, { exerciseData: data, manifest, exercisePagePath });
        return;
      }

      const shown = result.issues.slice(0, 8).map((text) => el("li", { text }));
      if (result.issues.length > 8) shown.push(el("li", { text: "… und " + (result.issues.length - 8) + " weitere Abweichungen." }));
      const hasMarks = loaded.mode === "table" && (result.wrongCells.size || result.wrongColumns.size || result.wrongRows.size);

      feedback.classList.add("is-error");
      feedback.appendChild(el("p", { text: "Das geladene Ergebnis entspricht noch nicht der Zieltabelle:" }));
      feedback.appendChild(el("ul", { class: "pq-feedback-list" }, shown));
      if (hasMarks) {
        feedback.appendChild(el("p", { text: "Abweichende Zellen sind auf dem Blatt „" + loaded.sheetName + "“ rot markiert. Öffne die Abfrage per Doppelklick im Bereich „Abfragen & Verbindungen“, um sie weiter zu bearbeiten." }));
      }
      shared.showErrorPopup(
        stage,
        result.issues.length === 1 ? result.issues[0] : result.issues.length + " Abweichungen zur Zieltabelle – die Details stehen unter dem Prüfen-Button."
      );
    }

    function reset() {
      UI.closeMenus(0);
      stage.querySelectorAll(".xl-dialog-backdrop").forEach((n) => n.remove());
      Object.assign(app, {
        tableCreated: false,
        query: null,
        view: "excel",
        editor: null,
        check: null,
        excel: { sheet: app.sourceSheetName, sel: { r: 0, c: 0 }, queriesPane: false, message: null },
      });
      clearFeedback();
      if (solutionBox) solutionBox.classList.remove("is-visible");
      app.render();
    }

    document.getElementById("btn-check").addEventListener("click", check);
    document.getElementById("btn-reset").addEventListener("click", reset);

    app.render();
    window.__excelFloPQApp = app; // für Debugging/Tests im Browser
  }

  document.addEventListener("DOMContentLoaded", initPowerQueryExercise);
})();
