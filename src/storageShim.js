// TEMPORARY DEV-ONLY SHIM
// -------------------------------------------------------------------------
// This app was originally built inside Claude, where `window.storage` is a
// built-in API that saves data on Anthropic's servers and shares it between
// everyone who opens the same published link.
//
// That API does not exist outside Claude. This file creates a stand-in
// version of it, backed by your browser's localStorage, so the app runs
// and you can click around during development.
//
// IMPORTANT: localStorage only lives in ONE browser on ONE device. It is
// NOT shared between teammates and is NOT a real multi-user backend.
// Before you deploy this for your team to use together, this file needs to
// be replaced with real calls to a database (e.g. via a small API route
// backed by Postgres, Supabase, etc.). Until then, treat this build as
// "preview only" — good for checking the UI works, not for real team data.
// -------------------------------------------------------------------------

if (typeof window !== "undefined" && !window.storage) {
  const keyFor = (key, shared) => `scc-shim:${shared ? "shared" : "personal"}:${key}`;

  window.storage = {
    async get(key, shared = false) {
      const raw = localStorage.getItem(keyFor(key, shared));
      if (raw === null) throw new Error(`Key not found: ${key}`);
      return { key, value: raw, shared };
    },
    async set(key, value, shared = false) {
      localStorage.setItem(keyFor(key, shared), value);
      return { key, value, shared };
    },
    async delete(key, shared = false) {
      localStorage.removeItem(keyFor(key, shared));
      return { key, deleted: true, shared };
    },
    async list(prefix = "", shared = false) {
      const scope = `scc-shim:${shared ? "shared" : "personal"}:`;
      const keys = Object.keys(localStorage)
        .filter(k => k.startsWith(scope + prefix))
        .map(k => k.slice(scope.length));
      return { keys, prefix, shared };
    },
  };
}
