/* Excel.Flo – Einstufungstest: Bildschirme und Bedienung
 *
 * Inhalte kommen aus fragen.json, Auswertung und Speicherstand aus logik.js
 * (window.ExcelFloEinstufung). Bausteine und Stile aus engine.js/engine.css
 * werden wiederverwendet; einstufungstest.css ergänzt nur, was es dort nicht gibt.
 *
 * Bildschirme: Start (auch „Fortsetzen“) → Frage (inkl. Abschlussfrage) → Ergebnis.
 */

(function () {
  "use strict";

  const L = window.ExcelFloEinstufung;
  const el = window.ExcelFlo.el;

  let root = null;
  let daten = null;
  let stand = null;
  let storage = null;
  let ansicht = "start"; // "start" | "frage" | "ergebnis"

  /* ---------------- Tracking ---------------- */

  // Gleiches Schema wie das Übungsportal (assets/tracking.js → Supabase-Tabelle public.events).
  // exercise_id = Frage-ID, wo es um eine Frage geht. Ohne geladenes Modul passiert nichts.
  function melde(event, frageId, detail) {
    if (window.ExcelFloTracking) window.ExcelFloTracking.track("einstufungstest", frageId || null, event, detail);
  }

  const gemeldeteAntworten = {}; // pro Seitenaufruf: dieselbe Antwort nicht mehrfach melden
  let abbruchGemeldet = false;

  function meldeAntwort() {
    const schritt = stand.schritt;
    const frage = L.frageFuerSchritt(daten, schritt);
    const antwortId = L.antwortFuerSchritt(daten, stand, schritt);
    if (!frage || !antwortId || gemeldeteAntworten[frage.id] === antwortId) return;
    gemeldeteAntworten[frage.id] = antwortId;

    if (L.istAbschlussfrage(daten, schritt)) {
      melde("question_answer", frage.id, { nummer: schritt + 1, antwort: antwortId, abschlussfrage: true });
      return;
    }
    const option = frage.options.find((o) => o.id === antwortId);
    melde("question_answer", frage.id, {
      nummer: schritt + 1,
      typ: frage.typ,
      kapitel: frage.kapitelBezug,
      antwort: antwortId,
      richtig: antwortId === frage.correctOptionId,
      weiss_nicht: !!(option && option.weissNicht),
    });
  }

  function meldeAbschluss() {
    const e = L.auswerten(daten, stand);
    melde("test_complete", null, {
      punkte: e.punkte,
      gesamt: e.gesamt,
      startkapitel: e.startKapitel ? e.startKapitel.nummer : null,
      fehlantworten: e.fehlantworten,
      weiss_nicht: e.weissNicht,
      antwortmuster: !!e.antwortmusterText,
      schwerpunkt: e.schwerpunkt,
      schwerpunkt_empfehlung: !!e.schwerpunktEmpfehlung,
    });
  }

  // Abbruch = Seite wird verlassen, während eine Frage offen ist. Nicht exakt (z. B. zählt
  // auch ein Neuladen), deshalb gibt es zusätzlich test_resume beim Fortsetzen.
  function meldeAbbruch() {
    if (abbruchGemeldet || ansicht !== "frage" || !stand || stand.status !== "laeuft") return;
    abbruchGemeldet = true;
    const frage = L.frageFuerSchritt(daten, stand.schritt);
    melde("test_abort", frage ? frage.id : null, {
      nummer: stand.schritt + 1,
      beantwortet: Object.keys(stand.antworten).length,
    });
  }

  function text(key, werte) {
    return L.platzhalter(daten.texte[key], werte);
  }

  function holeStorage() {
    try {
      return window.localStorage; // Zugriff kann im iframe eine Ausnahme werfen
    } catch (e) {
      return null;
    }
  }

  function setzeStand(neu) {
    stand = neu;
    L.speichereStand(stand, storage);
  }

  // Text mit `Formeln` → Knoten, Formeln als <code>
  function mitFormeln(tag, attrs, inhalt) {
    const node = el(tag, attrs);
    L.formelAbschnitte(inhalt).forEach((a) => {
      node.appendChild(a.formel ? el("code", { class: "test-formula", text: a.text }) : document.createTextNode(a.text));
    });
    return node;
  }

  /* ---------------- Rendern ---------------- */

  // nachAktion: nach einem Klick den Fokus auf die neue Überschrift setzen. Gescrollt wird nur,
  // wenn der Anfang des neuen Bildschirms nicht zu sehen ist – sonst bleibt die Seite, wo sie ist.
  // Beim ersten Laden nie – sonst springt die einbettende Ablefy-Seite zum iframe.
  function render(nachAktion) {
    root.innerHTML = "";
    if (ansicht === "ergebnis") renderErgebnis();
    else if (ansicht === "frage") renderFrage();
    else renderStart();

    if (nachAktion) {
      const ziel = root.querySelector("[data-fokus]");
      if (ziel) ziel.focus({ preventScroll: true });
      anfangSichtbarMachen(root.firstElementChild);
    }
  }

  // IntersectionObserver prüft auch im iframe gegen das tatsächlich sichtbare Browserfenster
  // (also inklusive Scrollposition der Ablefy-Seite).
  function anfangSichtbarMachen(anfang) {
    if (!anfang) return;
    if (!("IntersectionObserver" in window)) {
      if (anfang.getBoundingClientRect().top < 0) anfang.scrollIntoView({ block: "start" });
      return;
    }
    const beobachter = new IntersectionObserver((eintraege) => {
      beobachter.disconnect();
      if (eintraege[0].intersectionRatio < 0.99) anfang.scrollIntoView({ block: "start" });
    });
    beobachter.observe(anfang);
  }

  /* ---------------- Startbildschirm ---------------- */

  function renderStart() {
    const fortsetzen = stand.status === "laeuft";

    const startButton = el("button", {
      type: "button",
      class: "btn btn--primary test-btn",
      text: text(fortsetzen ? "buttonFortsetzen" : "buttonStarten"),
    });
    startButton.addEventListener("click", () => {
      if (fortsetzen) {
        melde("test_resume", null, { nummer: stand.schritt + 1, beantwortet: Object.keys(stand.antworten).length });
      } else {
        setzeStand(L.starten(daten, stand));
        melde("test_start");
      }
      ansicht = "frage";
      render(true);
    });

    // Vier Elemente: Überschrift, Fließtext, Fakten, Button
    root.appendChild(
      el("section", { class: "test-panel test-start" }, [
        el("h1", { class: "test-start__title", text: text("startUeberschrift"), tabindex: "-1", "data-fokus": "" }),
        ...daten.texte.startText.map((absatz) => el("p", { class: "test-start__text", text: absatz })),
        el("ul", { class: "test-facts" }, daten.texte.startEckdaten.map((t) => el("li", { text: t }))),
        fortsetzen ? el("p", { class: "test-resume", text: text("fortsetzenHinweis") }) : null,
        el("div", { class: "test-actions" }, [startButton]),
      ])
    );

    if (fortsetzen) root.appendChild(neuStartenBereich());
  }

  /* ---------------- Fragebildschirm ---------------- */

  function fortschrittTextFuer(schritt) {
    return L.istAbschlussfrage(daten, schritt)
      ? text("fortschrittAbschluss")
      : text("fortschrittFrage", { nummer: schritt + 1, gesamt: L.anzahlFragen(daten) });
  }

  function renderFrage() {
    const schritt = stand.schritt;
    const abschluss = L.istAbschlussfrage(daten, schritt);
    const fortschrittText = fortschrittTextFuer(schritt);

    // Fortschritt: sichtbar oben; Screenreader hören ihn stattdessen direkt vor der Frage (siehe frageFieldset)
    const anteil = Math.round(((schritt + 1) / L.anzahlSchritte(daten)) * 100);
    root.appendChild(
      el("div", { class: "test-progress", "aria-hidden": "true" }, [
        el("p", { class: "test-progress__label", text: fortschrittText }),
        el("div", { class: "test-progress__bar", "aria-hidden": "true" }, [
          el("span", { class: "test-progress__fill", style: "width:" + anteil + "%" }),
        ]),
      ])
    );

    const hinweisSlot = el("div", { class: "test-hint-slot", id: "test-hinweis" });

    const nav = el("div", { class: "test-nav" });
    if (schritt > 0) {
      const zurueckButton = el("button", { type: "button", class: "btn btn--secondary test-btn", text: text("buttonZurueck") });
      zurueckButton.addEventListener("click", () => {
        setzeStand(L.zurueck(daten, stand));
        render(true);
      });
      nav.appendChild(zurueckButton);
    }
    nav.appendChild(
      el("button", {
        type: "submit",
        class: "btn btn--primary test-btn test-nav__next",
        text: text(abschluss ? "buttonAuswerten" : "buttonWeiter"),
      })
    );

    // Hinweis unter den Buttons: Beim Einblenden rutschen sie so nicht nach unten
    const form = el("form", { class: "test-panel test-question", novalidate: "" }, [
      frageFieldset(schritt, true),
      nav,
      hinweisSlot,
    ]);
    form.addEventListener("submit", (e) => {
      e.preventDefault();
      weiterKlick();
    });

    root.appendChild(form);
    fragenHoeheAngleichen();
  }

  // Frage mit Antwortkarten. mitAktionen = false: nur für die unsichtbare Höhenmessung.
  function frageFieldset(schritt, mitAktionen) {
    const frage = L.frageFuerSchritt(daten, schritt);
    const abschluss = L.istAbschlussfrage(daten, schritt);
    const gewaehlt = mitAktionen ? L.antwortFuerSchritt(daten, stand, schritt) : null;

    const optionen = frage.options.map((option, i) => {
      const input = el("input", {
        type: "radio",
        name: mitAktionen ? "antwort" : "messung",
        value: option.id,
        class: "answer-option__input",
        "aria-describedby": "test-hinweis",
      });
      if (mitAktionen) {
        if (option.id === gewaehlt) input.checked = true;
        input.addEventListener("change", () => waehleAntwort(option.id));
      }

      return el("label", { class: "answer-option" + (option.id === gewaehlt ? " is-selected" : "") }, [
        input,
        el("span", { class: "answer-option__key", "aria-hidden": "true", text: String(i + 1) }),
        mitFormeln("span", { class: "answer-option__text" }, option.text),
      ]);
    });

    // Beim Fokuswechsel auf die Frage liest der Screenreader „Frage 4 von 10: …“ vor
    const prompt = mitFormeln("h1", { class: "test-question__prompt", tabindex: "-1", "data-fokus": "" }, frage.prompt);
    prompt.insertBefore(el("span", { class: "test-sr-only", text: fortschrittTextFuer(schritt) + ": " }), prompt.firstChild);

    return el("fieldset", {}, [
      el("legend", {}, [prompt]),
      abschluss && frage.hinweis ? el("p", { class: "test-muted test-question__note", text: frage.hinweis }) : null,
      el("div", { class: "answer-list" }, optionen),
    ]);
  }

  // Alle Fragen gleich hoch, damit Zurück/Weiter beim Durchklicken an derselben Stelle bleiben:
  // Die längste Frage wird unsichtbar in der aktuellen Breite gemessen (neu bei anderer Breite).
  let fragenHoehe = { breite: 0, hoehe: 0 };

  function fragenHoeheAngleichen() {
    const form = root.querySelector(".test-question");
    if (!form) return;
    const breite = form.offsetWidth;

    if (fragenHoehe.breite !== breite) {
      const messung = el("div", { class: "test-panel test-question", "aria-hidden": "true" });
      messung.style.cssText = "position:absolute;left:-10000px;top:0;visibility:hidden;width:" + breite + "px";
      document.body.appendChild(messung);
      let hoehe = 0;
      for (let s = 0; s < L.anzahlSchritte(daten); s++) {
        const probe = frageFieldset(s, false);
        messung.appendChild(probe);
        hoehe = Math.max(hoehe, probe.offsetHeight);
        probe.remove();
      }
      messung.remove();
      fragenHoehe = { breite: breite, hoehe: hoehe };
    }

    form.querySelector("fieldset").style.minHeight = fragenHoehe.hoehe + "px";
  }

  function waehleAntwort(optionId) {
    setzeStand(L.antwortSetzen(daten, stand, optionId));
    root.querySelectorAll(".answer-option").forEach((label) => {
      const input = label.querySelector("input");
      const aktiv = input.value === optionId;
      input.checked = aktiv;
      label.classList.toggle("is-selected", aktiv);
    });
    zeigeHinweis(false);
  }

  function zeigeHinweis(sichtbar) {
    const slot = document.getElementById("test-hinweis");
    if (!slot) return;
    slot.innerHTML = "";
    if (sichtbar) {
      slot.appendChild(
        el("div", { class: "exercise-feedback is-error", role: "alert" }, [el("p", { text: text("hinweisKeineAuswahl") })])
      );
    }
  }

  function weiterKlick() {
    if (!L.kannWeiter(daten, stand)) {
      zeigeHinweis(true);
      return;
    }
    meldeAntwort();
    setzeStand(L.weiter(daten, stand));
    if (stand.status === "abgeschlossen") {
      ansicht = "ergebnis";
      meldeAbschluss();
    }
    render(true);
  }

  // Zahlentasten 1–4 wählen eine Antwort, Enter geht weiter
  function handleKeydown(e) {
    if (ansicht !== "frage" || e.ctrlKey || e.altKey || e.metaKey) return;
    const tag = e.target && e.target.tagName;

    if (e.key === "Enter") {
      if (tag === "BUTTON" || tag === "A" || tag === "SUMMARY") return; // eigene Aktion des Elements
      e.preventDefault();
      if (!e.repeat) weiterKlick();
      return;
    }

    if (/^[1-9]$/.test(e.key)) {
      const inputs = root.querySelectorAll(".answer-option__input");
      const input = inputs[Number(e.key) - 1];
      if (!input) return;
      e.preventDefault();
      input.focus();
      waehleAntwort(input.value);
    }
  }

  /* ---------------- Ergebnisbildschirm ---------------- */

  function renderErgebnis() {
    const e = L.auswerten(daten, stand);

    // 1. Punktzahl
    root.appendChild(
      el("div", { class: "test-result__head" }, [
        el("h1", { class: "test-result__title", text: text("ergebnisUeberschrift"), tabindex: "-1", "data-fokus": "" }),
        el("p", { class: "test-score" }, [
          el("span", { class: "test-score__value", text: text("ergebnisPunkte", { punkte: e.punkte, gesamt: e.gesamt }) }),
          el("span", { class: "test-score__label", text: text("ergebnisPunkteLabel") }),
        ]),
      ])
    );

    // 2.–4. Stufentext, Antwortmuster, Schwerpunkt
    const panel = el("section", { class: "test-panel test-result" });
    if (e.stufe) {
      panel.appendChild(el("p", { class: "test-result__text", text: e.stufe.beschreibung }));
      panel.appendChild(el("p", { class: "test-startpoint", text: e.stufe.startpunkt }));
    }
    if (e.antwortmusterText) panel.appendChild(el("p", { class: "test-result__note", text: e.antwortmusterText }));
    if (e.schwerpunktEmpfehlung) panel.appendChild(el("p", { class: "test-result__note", text: e.schwerpunktEmpfehlung }));

    // 5.–6. Kapitel-Button (bzw. Hinweis ohne Link) und „Alle Kapitel“
    const aktionen = el("div", { class: "test-actions test-result__actions" });
    const k = e.startKapitel;
    // target="_top": Kursseite im ganzen Fenster öffnen, nicht im iframe. Das Ereignis geht per keepalive trotzdem raus.
    if (k && k.url) {
      const kapitelButton = el("a", { class: "btn btn--primary test-btn", href: k.url, target: "_top", text: text("buttonKapitel", k) });
      kapitelButton.addEventListener("click", () => melde("chapter_click", null, { kapitel: k.nummer, punkte: e.punkte }));
      aktionen.appendChild(kapitelButton);
    } else if (k) {
      aktionen.appendChild(el("p", { class: "test-next-step", text: text("hinweisOhneLink", k) }));
    }
    if (daten.einstellungen && daten.einstellungen.urlAlleKapitel) {
      const alleLink = el("a", { class: "test-textlink", href: daten.einstellungen.urlAlleKapitel, target: "_top", text: text("linkAlleKapitel") });
      alleLink.addEventListener("click", () => melde("chapter_click", null, { kapitel: "alle", punkte: e.punkte }));
      aktionen.appendChild(alleLink);
    }
    panel.appendChild(aktionen);
    root.appendChild(panel);

    // 7. Falsch beantwortete Fragen, zugeklappt
    if (e.fehlerliste.length) root.appendChild(fehlerliste(e.fehlerliste));

    root.appendChild(neuStartenBereich());
  }

  // Zweistufig: aufgeklappt zeigt die Liste nur die Fragen, jede Erklärung klappt einzeln auf.
  // Hält die Seite kurz genug für die feste iframe-Höhe in Ablefy.
  function fehlerliste(liste) {
    const eintraege = liste.map((x) =>
      el("li", { class: "test-review__item" }, [
        el("details", { class: "test-review__entry" }, [
          mitFormeln("summary", { class: "test-review__prompt" }, x.frage.prompt),
          el("div", { class: "test-review__body" }, [
            el("dl", { class: "test-review__answers" }, [
              el("dt", { text: text("fehlerlisteDeineAntwort") }),
              x.weissNicht
                ? el("dd", { text: text("fehlerlisteNichtBeantwortet") })
                : mitFormeln("dd", {}, x.antwort.text),
              el("dt", { text: text("fehlerlisteRichtigeAntwort") }),
              mitFormeln("dd", { class: "test-review__correct" }, x.richtigeOption ? x.richtigeOption.text : ""),
            ]),
            mitFormeln("p", { class: "test-review__explanation" }, x.frage.erklaerung || ""),
          ]),
        ]),
      ])
    );

    const titel =
      liste.length === 1 ? text("fehlerlisteTitelEinzahl") : text("fehlerlisteTitelMehrzahl", { anzahl: liste.length });

    return el("details", { class: "exercise-hints test-review" }, [
      el("summary", { text: titel }),
      el("ul", { class: "test-review__list" }, eintraege),
    ]);
  }

  /* ---------------- Neu starten (mit Nachfrage direkt auf der Seite, kein Popup) ---------------- */

  function neuStartenBereich() {
    const bereich = el("div", { class: "test-restart" });

    function zeigeLink() {
      bereich.innerHTML = "";
      const link = el("button", { type: "button", class: "test-textlink", text: text("linkNeuStarten") });
      link.addEventListener("click", zeigeNachfrage);
      bereich.appendChild(link);
      return link;
    }

    function zeigeNachfrage() {
      bereich.innerHTML = "";
      const ja = el("button", { type: "button", class: "btn btn--secondary test-btn", text: text("buttonNeuStartenJa") });
      const abbrechen = el("button", { type: "button", class: "test-textlink", text: text("buttonAbbrechen") });
      ja.addEventListener("click", () => {
        melde("test_restart", null, { vorher: stand.status });
        setzeStand(L.neuStarten());
        ansicht = "frage";
        render(true);
      });
      abbrechen.addEventListener("click", () => zeigeLink().focus());
      bereich.appendChild(
        el("div", { class: "test-restart__confirm", role: "group", "aria-label": text("linkNeuStarten") }, [
          el("p", { text: text("neuStartenBestaetigung") }),
          el("div", { class: "test-actions" }, [ja, abbrechen]),
        ])
      );
      ja.focus();
    }

    zeigeLink();
    return bereich;
  }

  /* ---------------- Start ---------------- */

  function init() {
    root = document.getElementById("einstufungstest-root");
    if (!root || !L || !el) return;

    fetch(root.dataset.daten, { cache: "no-cache" })
      .then((res) => {
        if (!res.ok) throw new Error("fragen.json konnte nicht geladen werden (" + res.status + ")");
        return res.json();
      })
      .then((json) => {
        daten = json;
        const probleme = L.pruefeDaten(daten);
        if (probleme.length) console.warn("Excel.Flo Einstufungstest – Probleme in fragen.json:\n- " + probleme.join("\n- "));

        storage = holeStorage();
        stand = L.ladeStand(daten, storage);
        ansicht = stand.status === "abgeschlossen" ? "ergebnis" : "start";
        document.addEventListener("keydown", handleKeydown);
        window.addEventListener("pagehide", meldeAbbruch);
        window.addEventListener("pageshow", (ev) => {
          if (ev.persisted) abbruchGemeldet = false; // aus dem Zurück-Cache wiederhergestellt
        });
        window.addEventListener("resize", () => {
          if (ansicht === "frage") fragenHoeheAngleichen();
        });
        if (document.fonts && document.fonts.addEventListener) {
          document.fonts.addEventListener("loadingdone", () => {
            fragenHoehe.breite = 0; // Schrift nachgeladen → Höhen neu messen
            if (ansicht === "frage") fragenHoeheAngleichen();
          });
        }

        // Schriften abwarten (höchstens 1,5 s), damit die Fragenhöhe richtig gemessen wird
        const schriften = document.fonts
          ? Promise.race([document.fonts.ready, new Promise((r) => setTimeout(r, 1500))])
          : Promise.resolve();
        return schriften.then(() => render(false));
      })
      .catch((err) => {
        console.error(err);
        root.innerHTML = "";
        root.appendChild(
          el("p", {
            class: "exercise-feedback is-error test-load-error",
            text: daten ? daten.texte.ladeFehler : "Der Einstufungstest konnte nicht geladen werden. Bitte lade die Seite neu.",
          })
        );
      });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
