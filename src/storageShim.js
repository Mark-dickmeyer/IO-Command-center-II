// This file replaces Claude's built-in `window.storage` with a real,
// shared, permanent backend using Supabase (a hosted Postgres database).
//
// It keeps the EXACT same get/set/delete/list shape that the rest of the
// app already expects, so nothing in App.jsx needs to change.
//
// HOW IT DECIDES WHAT TO USE:
//   - If VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY are set (see README),
//     it talks to your real Supabase database. This is shared between
//     everyone who opens the app — the real, production behavior.
//   - If those aren't set yet, it falls back to your browser's own local
//     storage, purely so the app still loads and is clickable while you're
//     getting Supabase set up. That fallback is NOT shared between people.

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const TABLE = "app_storage";

async function setupSupabaseStorage() {
  const { createClient } = await import("@supabase/supabase-js");
  const supabase = createClient(supabaseUrl, supabaseKey);

  window.storage = {
    async get(key, shared = false) {
      const { data, error } = await supabase
        .from(TABLE)
        .select("value")
        .eq("key", key)
        .eq("shared", shared)
        .maybeSingle();
      if (error) throw error;
      if (!data) throw new Error(`Key not found: ${key}`);
      return { key, value: data.value, shared };
    },
    async set(key, value, shared = false) {
      const { error } = await supabase
        .from(TABLE)
        .upsert({ key, shared, value, updated_at: new Date().toISOString() }, { onConflict: "key,shared" });
      if (error) throw error;
      return { key, value, shared };
    },
    async delete(key, shared = false) {
      const { error } = await supabase.from(TABLE).delete().eq("key", key).eq("shared", shared);
      if (error) throw error;
      return { key, deleted: true, shared };
    },
    async list(prefix = "", shared = false) {
      let query = supabase.from(TABLE).select("key").eq("shared", shared);
      if (prefix) query = query.like("key", `${prefix}%`);
      const { data, error } = await query;
      if (error) throw error;
      return { keys: (data || []).map(r => r.key), prefix, shared };
    },
  };
}

function setupLocalFallbackStorage() {
  console.warn(
    "[storage] VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY are not set. " +
    "Falling back to browser-only local storage — data will NOT be shared " +
    "between teammates. See README.md to connect Supabase."
  );
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

export async function initStorage() {
  if (typeof window === "undefined" || window.storage) return;
  if (supabaseUrl && supabaseKey) {
    await setupSupabaseStorage();
  } else {
    setupLocalFallbackStorage();
  }
}
