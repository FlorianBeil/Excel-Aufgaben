/* Excel.Flo – Fortschritts-Speicherung
 *
 * Stabile API (ExcelFloProgress.*), die vom Rest der App genutzt wird und
 * überall SYNCHRON bleibt (kein await nötig), damit engine.js unverändert
 * bleibt: localStorage ist die sofort verfügbare Quelle der Wahrheit für
 * das Rendering. Im Hintergrund wird zusätzlich – best effort, blockiert
 * nichts – mit Supabase synchronisiert (anonyme Auth-Session, später per
 * setEmail() mit einer E-Mail verknüpfbar), damit Fortschritt
 * geräteübergreifend erhalten bleibt.
 */

(function () {
  "use strict";

  const SUPABASE_URL = "https://hbhagmmbowplzjzfvuao.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_Ee47HbgO5Ne8PP9Jh5ugag_iE_lZcp6";

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
    syncPush();
    return state;
  }

  function getCompletedIds() {
    return load().completedExerciseIds.slice();
  }

  function setEmail(email) {
    const state = load();
    state.email = email;
    save(state);
    linkEmail(email);
    return state;
  }

  function getEmail() {
    return load().email;
  }

  /* ---------------- Supabase-Sync (best effort, läuft im Hintergrund) ---------------- */

  let supabaseClient = null;
  let sessionPromise = null;

  function getClient() {
    if (!window.supabase || !window.supabase.createClient) return null; // SDK nicht geladen (z. B. offline/geblockt)
    if (!supabaseClient) supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    return supabaseClient;
  }

  function ensureSession() {
    const client = getClient();
    if (!client) return Promise.resolve(null);
    if (sessionPromise) return sessionPromise;

    sessionPromise = client.auth
      .getSession()
      .then(({ data }) => {
        if (data && data.session) return data.session;
        return client.auth.signInAnonymously().then(({ data: signInData, error }) => {
          if (error) throw error;
          return signInData.session;
        });
      })
      .catch((err) => {
        console.warn("Excel.Flo: Supabase-Session nicht verfügbar, Fortschritt bleibt lokal.", err);
        return null;
      });

    return sessionPromise;
  }

  function syncPull() {
    return ensureSession().then((session) => {
      const client = getClient();
      if (!client || !session) return;
      return client
        .from("progress")
        .select("completed_exercise_ids, email")
        .eq("user_id", session.user.id)
        .maybeSingle()
        .then(({ data, error }) => {
          if (error || !data) return;
          const state = load();
          const merged = Array.from(new Set(state.completedExerciseIds.concat(data.completed_exercise_ids || [])));
          const changed =
            merged.length !== state.completedExerciseIds.length || (data.email && !state.email);
          state.completedExerciseIds = merged;
          if (data.email && !state.email) state.email = data.email;
          save(state);
          if (changed) document.dispatchEvent(new CustomEvent("excelflo:progress-synced"));
        });
    });
  }

  function syncPush() {
    return ensureSession()
      .then((session) => {
        const client = getClient();
        if (!client || !session) return;
        const state = load();
        return client.from("progress").upsert({
          user_id: session.user.id,
          email: state.email,
          completed_exercise_ids: state.completedExerciseIds,
        });
      })
      .catch((err) => {
        console.warn("Excel.Flo: Fortschritt konnte nicht synchronisiert werden.", err);
      });
  }

  function linkEmail(email) {
    return ensureSession()
      .then((session) => {
        const client = getClient();
        if (!client || !session) return;
        return client.auth.updateUser({ email }).then(({ error }) => {
          if (error) throw error;
          return syncPush();
        });
      })
      .catch((err) => {
        console.warn("Excel.Flo: E-Mail konnte nicht verknüpft werden.", err);
      });
  }

  // Beim Laden bereits vorhandenen Fortschritt (z. B. von einem anderen Gerät) abholen.
  syncPull();

  window.ExcelFloProgress = {
    getAnonId,
    isCompleted,
    markCompleted,
    getCompletedIds,
    setEmail,
    getEmail,
  };
})();
