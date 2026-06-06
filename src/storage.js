/* Recreates the artifact storage API on Supabase so Eudai_ControlCenter.jsx
   needs zero storage-code changes. One kv table, shared by all founders. */
import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

export function installStorage() {
  window.storage = {
    async get(key) {
      const { data, error } = await supabase.from("kv").select("value").eq("key", key).single();
      if (error || !data) throw new Error("not found");
      return { key, value: data.value };
    },
    async set(key, value) {
      const { error } = await supabase.from("kv").upsert({ key, value, updated_at: new Date().toISOString() });
      if (error) { console.error("storage.set failed", error); return null; }
      return { key, value };
    },
    async delete(key) {
      await supabase.from("kv").delete().eq("key", key);
      return { key, deleted: true };
    },
  };

  // Bearer token for the AI proxy (api/claude.js verifies it).
  window.getAuthToken = async () => {
    const { data } = await supabase.auth.getSession();
    return data.session?.access_token || "";
  };
}
