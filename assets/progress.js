/* Excel.Flo – Fortschritts-Speicherung
 *
 * Stabile API (ExcelFloProgress.*), die vom Rest der App genutzt wird. Die
 * Implementierung liegt aktuell lokal (localStorage) – anonyme ID wird beim
 * ersten Besuch automatisch erzeugt, kein Login nötig. Sobald ein Supabase-
 * Projekt eingerichtet ist, wird NUR diese Datei ausgetauscht (gegen eine
 * Version, die zusätzlich mit Supabase synchronisiert); die Aufrufer ändern
 * sich nicht.
 */

(function () {
  "use strict";

  const STORAGE_KEY = "excelflo_progress_v1";
  const ID_KEY = "excelflo_anon_id";

  function uuid() {
    if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
    return "anon-" + Date.now() + "-" + Math.random().toString(16).slice(2);
  }

  function getAnonId() {
    let id;
    try {
      id = localStorage.getItem(ID_KEY);
    } catch (e) {
      return null; // z. B. localStorage im iframe blockiert
    }
    if (!id) {
      id = uuid();
      try {
        localStorage.setItem(ID_KEY, id);
      } catch (e) {
        // Speichern nicht möglich – Fortschritt bleibt dann nur für diese Sitzung im Speicher
      }
    }
    return id;
  }

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return { completedExerciseIds: [], email: null, updatedAt: null };
      const parsed = JSON.parse(raw);
      return {
        completedExerciseIds: Array.isArray(parsed.completedExerciseIds) ? parsed.completedExerciseIds : [],
        email: parsed.email || null,
        updatedAt: parsed.updatedAt || null,
      };
    } catch (e) {
      return { completedExerciseIds: [], email: null, updatedAt: null };
    }
  }

  function save(state) {
    state.updatedAt = new Date().toISOString();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      // z. B. Storage-Kontingent voll oder blockiert – Fortschritt bleibt für diese Sitzung erhalten
    }
    return state;
  }

  function isCompleted(exerciseId) {
    return load().completedExerciseIds.indexOf(exerciseId) !== -1;
  }

  function markCompleted(exerciseId) {
    const state = load();
    if (state.completedExerciseIds.indexOf(exerciseId) === -1) {
      state.completedExerciseIds.push(exerciseId);
      save(state);
    }
    return state;
  }

  function getCompletedIds() {
    return load().completedExerciseIds.slice();
  }

  function setEmail(email) {
    const state = load();
    state.email = email;
    save(state);
    return state;
  }

  function getEmail() {
    return load().email;
  }

  window.ExcelFloProgress = {
    getAnonId,
    isCompleted,
    markCompleted,
    getCompletedIds,
    setEmail,
    getEmail,
  };
})();
