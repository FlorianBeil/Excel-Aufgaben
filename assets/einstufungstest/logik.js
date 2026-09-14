/* Excel.Flo – Einstufungstest: Logik ohne Bildschirmaufbau
 *
 * Alles hier arbeitet nur mit den Inhalten aus fragen.json und einem einfachen
 * Speicherstand-Objekt – kein DOM, kein fetch. Dadurch lässt sich die Auswertung
 * auch ohne Browser testen (node). Der Speicherstand wird nie verändert, sondern
 * jede Aktion gibt einen neuen Stand zurück.
 *
 * Ablauf: status "neu" (Startbildschirm) → "laeuft" (schritt 0 … anzahlFragen,
 * der letzte Schritt ist die Abschlussfrage) → "abgeschlossen" (Ergebnis).
 */

(function (root) {
  "use strict";

  const STORAGE_KEY = "excelflo_einstufungstest_v1";
  const STAND_VERSION = 1;

  /* ---------------- Schritte ---------------- */

  function anzahlFragen(daten) {
    return daten.fragen.length;
  }

  // Fragen plus Abschlussfrage
  function anzahlSchritte(daten) {
    return daten.fragen.length + 1;
  }

  function istAbschlussfrage(daten, schritt) {
    return schritt === daten.fragen.length;
  }

  function frageFuerSchritt(daten, schritt) {
    return istAbschlussfrage(daten, schritt) ? daten.abschlussfrage : daten.fragen[schritt];
  }

  /* ---------------- Speicherstand ---------------- */

  function neuerStand() {
    return { version: STAND_VERSION, status: "neu", schritt: 0, antworten: {}, schwerpunkt: null };
  }

  function kopie(stand) {
    return Object.assign({}, stand, { antworten: Object.assign({}, stand.antworten) });
  }

  function optionGibtEs(frage, optionId) {
    return frage.options.some((o) => o.id === optionId);
  }

  function antwortFuerSchritt(daten, stand, schritt) {
    if (istAbschlussfrage(daten, schritt)) return stand.schwerpunkt;
    const frage = daten.fragen[schritt];
    return frage ? stand.antworten[frage.id] || null : null;
  }

  function ersterUnbeantworteterSchritt(daten, stand) {
    for (let i = 0; i < anzahlSchritte(daten); i++) {
      if (!antwortFuerSchritt(daten, stand, i)) return i;
    }
    return -1;
  }

  // Gespeicherten Stand an die aktuellen Inhalte anpassen: Antworten zu Fragen oder
  // Optionen, die es nicht mehr gibt, fallen weg. Fehlt nach einer Inhaltsänderung
  // eine Antwort, läuft ein abgeschlossener Test an dieser Stelle weiter.
  function bereinigeStand(daten, roh) {
    if (!roh || typeof roh !== "object" || roh.version !== STAND_VERSION) return neuerStand();

    const stand = neuerStand();
    const antworten = roh.antworten && typeof roh.antworten === "object" ? roh.antworten : {};
    daten.fragen.forEach((frage) => {
      if (optionGibtEs(frage, antworten[frage.id])) stand.antworten[frage.id] = antworten[frage.id];
    });
    if (optionGibtEs(daten.abschlussfrage, roh.schwerpunkt)) stand.schwerpunkt = roh.schwerpunkt;

    if (roh.status === "laeuft" || roh.status === "abgeschlossen") stand.status = roh.status;

    const letzterSchritt = anzahlSchritte(daten) - 1;
    const schritt = Number.isInteger(roh.schritt) ? roh.schritt : 0;
    stand.schritt = Math.min(Math.max(schritt, 0), letzterSchritt);

    if (stand.status === "abgeschlossen") {
      const offen = ersterUnbeantworteterSchritt(daten, stand);
      if (offen !== -1) {
        stand.status = "laeuft";
        stand.schritt = offen;
      }
    } else if (stand.status === "laeuft") {
      // Nicht hinter die erste offene Frage springen lassen
      const offen = ersterUnbeantworteterSchritt(daten, stand);
      if (offen !== -1 && stand.schritt > offen) stand.schritt = offen;
    }
    return stand;
  }

  // storage: localStorage oder ein Ersatz mit getItem/setItem/removeItem.
  // Im iframe kann localStorage gesperrt sein – dann läuft der Test ohne Fortsetzen.
  function ladeStand(daten, storage) {
    try {
      const raw = storage && storage.getItem(STORAGE_KEY);
      return raw ? bereinigeStand(daten, JSON.parse(raw)) : neuerStand();
    } catch (e) {
      return neuerStand();
    }
  }

  function speichereStand(stand, storage) {
    try {
      if (storage) storage.setItem(STORAGE_KEY, JSON.stringify(stand));
    } catch (e) {
      // Speichern nicht möglich – der Test läuft trotzdem weiter
    }
  }

  function loescheStand(storage) {
    try {
      if (storage) storage.removeItem(STORAGE_KEY);
    } catch (e) {}
  }

  /* ---------------- Aktionen (geben jeweils einen neuen Stand zurück) ---------------- */

  function starten(daten, stand) {
    const neu = kopie(stand);
    if (neu.status === "neu") {
      neu.status = "laeuft";
      neu.schritt = 0;
    }
    return neu;
  }

  function neuStarten() {
    return Object.assign(neuerStand(), { status: "laeuft" });
  }

  function antwortSetzen(daten, stand, optionId) {
    const frage = frageFuerSchritt(daten, stand.schritt);
    if (stand.status !== "laeuft" || !frage || !optionGibtEs(frage, optionId)) return stand;
    const neu = kopie(stand);
    if (istAbschlussfrage(daten, stand.schritt)) neu.schwerpunkt = optionId;
    else neu.antworten[frage.id] = optionId;
    return neu;
  }

  function kannWeiter(daten, stand) {
    return stand.status === "laeuft" && !!antwortFuerSchritt(daten, stand, stand.schritt);
  }

  // Ohne Antwort bleibt der Stand unverändert (die Oberfläche zeigt dann den Hinweis).
  function weiter(daten, stand) {
    if (!kannWeiter(daten, stand)) return stand;
    const neu = kopie(stand);
    if (istAbschlussfrage(daten, stand.schritt)) neu.status = "abgeschlossen";
    else neu.schritt = stand.schritt + 1;
    return neu;
  }

  function zurueck(daten, stand) {
    if (stand.status !== "laeuft" || stand.schritt === 0) return stand;
    const neu = kopie(stand);
    neu.schritt = stand.schritt - 1;
    return neu;
  }

  /* ---------------- Auswertung ---------------- */

  function kapitelInfo(daten, nummer) {
    const k = daten.kapitel[nummer] || {};
    return { nummer: nummer, titel: k.titel || "", url: k.url || "" };
  }

  function auswerten(daten, stand) {
    let punkte = 0;
    let fehlantworten = 0;
    let weissNicht = 0;
    const fehlerliste = [];

    daten.fragen.forEach((frage) => {
      const gewaehlt = frage.options.find((o) => o.id === stand.antworten[frage.id]) || null;
      const richtigeOption = frage.options.find((o) => o.id === frage.correctOptionId);
      const richtig = !!gewaehlt && gewaehlt.id === frage.correctOptionId;

      if (richtig) {
        punkte++;
        return;
      }
      // „Weiß ich nicht“ (und eine fehlende Antwort) zählt wie falsch, aber nicht als echte Fehlantwort
      const unbekannt = !gewaehlt || !!gewaehlt.weissNicht;
      if (unbekannt) weissNicht++;
      else fehlantworten++;

      fehlerliste.push({ frage: frage, antwort: unbekannt ? null : gewaehlt, weissNicht: unbekannt, richtigeOption: richtigeOption });
    });

    const stufe = daten.stufen.find((s) => punkte >= s.minPunkte && punkte <= s.maxPunkte) || null;

    const muster = daten.antwortmuster;
    const antwortmusterText =
      muster && punkte >= muster.minPunkte && fehlantworten >= muster.minFehlantworten ? muster.text : null;

    // Schwerpunkt-Empfehlung entfällt bei „Gemischt“ und wenn das Kapitel schon im Stufentext steht
    const schwerpunktOption = daten.abschlussfrage.options.find((o) => o.id === stand.schwerpunkt) || null;
    let schwerpunktEmpfehlung = null;
    if (schwerpunktOption && schwerpunktOption.kapitel && schwerpunktOption.empfehlung) {
      const genannt = stufe && (stufe.genannteKapitel || []).indexOf(schwerpunktOption.kapitel) !== -1;
      if (!genannt) schwerpunktEmpfehlung = schwerpunktOption.empfehlung;
    }

    return {
      punkte: punkte,
      gesamt: daten.fragen.length,
      fehlantworten: fehlantworten,
      weissNicht: weissNicht,
      stufe: stufe,
      startKapitel: stufe ? kapitelInfo(daten, stufe.startKapitel) : null,
      antwortmusterText: antwortmusterText,
      schwerpunkt: schwerpunktOption ? schwerpunktOption.id : null,
      schwerpunktEmpfehlung: schwerpunktEmpfehlung,
      fehlerliste: fehlerliste,
    };
  }

  /* ---------------- Texthelfer ---------------- */

  // "Frage {nummer} von {gesamt}" → "Frage 4 von 10"; unbekannte Platzhalter bleiben stehen
  function platzhalter(text, werte) {
    return String(text || "").replace(/\{(\w+)\}/g, (m, name) =>
      werte && werte[name] !== undefined && werte[name] !== null ? String(werte[name]) : m
    );
  }

  // "Die Formel `=B2*$F$1` …" → [{text:"Die Formel ", formel:false}, {text:"=B2*$F$1", formel:true}, …]
  function formelAbschnitte(text) {
    return String(text || "")
      .split("`")
      .map((teil, i) => ({ text: teil, formel: i % 2 === 1 }))
      .filter((a) => a.text !== "");
  }

  /* ---------------- Inhaltsprüfung (Hilfe beim Bearbeiten von fragen.json) ---------------- */

  function pruefeDaten(daten) {
    const probleme = [];
    if (!daten || !Array.isArray(daten.fragen) || !daten.fragen.length) return ["Keine Fragen gefunden."];
    if (!daten.abschlussfrage || !Array.isArray(daten.abschlussfrage.options)) probleme.push("Abschlussfrage fehlt.");
    if (!Array.isArray(daten.stufen) || !daten.stufen.length) probleme.push("Keine Stufen gefunden.");

    const ids = new Set();
    daten.fragen.forEach((f, i) => {
      const name = "Frage " + (i + 1) + " (" + f.id + ")";
      if (!f.id) probleme.push("Frage " + (i + 1) + ": id fehlt.");
      else if (ids.has(f.id)) probleme.push(name + ": id kommt doppelt vor.");
      ids.add(f.id);
      if (!Array.isArray(f.options) || f.options.length < 2) {
        probleme.push(name + ": zu wenige Antwortoptionen.");
        return;
      }
      const richtige = f.options.find((o) => o.id === f.correctOptionId);
      if (!richtige) probleme.push(name + ": correctOptionId passt zu keiner Option.");
      else if (richtige.weissNicht) probleme.push(name + ": „Weiß ich nicht“ ist als richtige Antwort markiert.");
      if (!f.options.some((o) => o.weissNicht)) probleme.push(name + ": keine Option „Weiß ich nicht“.");
    });

    if (Array.isArray(daten.stufen)) {
      for (let p = 0; p <= daten.fragen.length; p++) {
        const treffer = daten.stufen.filter((s) => p >= s.minPunkte && p <= s.maxPunkte).length;
        if (treffer !== 1) probleme.push(p + " Punkte passen zu " + treffer + " Stufen (erwartet: genau 1).");
      }
      daten.stufen.forEach((s) => {
        if (!daten.kapitel || !daten.kapitel[s.startKapitel]) probleme.push("Kapitel " + s.startKapitel + " fehlt unter „kapitel“.");
      });
    }
    return probleme;
  }

  const api = {
    STORAGE_KEY,
    anzahlFragen,
    anzahlSchritte,
    istAbschlussfrage,
    frageFuerSchritt,
    antwortFuerSchritt,
    neuerStand,
    ladeStand,
    speichereStand,
    loescheStand,
    starten,
    neuStarten,
    antwortSetzen,
    kannWeiter,
    weiter,
    zurueck,
    auswerten,
    kapitelInfo,
    platzhalter,
    formelAbschnitte,
    pruefeDaten,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ExcelFloEinstufung = api;
})(typeof window !== "undefined" ? window : this);
