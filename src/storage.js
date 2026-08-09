import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabase = supabaseUrl && supabaseKey ? createClient(supabaseUrl, supabaseKey) : null;
export const supabaseConfigured = !!supabase;

// Drop-in replacements for the previous window.storage-based helpers.
// Data lives in a single "kv_store" table (key text primary key, value jsonb),
// see supabase.sql for the schema. This keeps the rest of the app unchanged.
export async function loadKey(key, fallback) {
  if (!supabase) return fallback;
  try {
    const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value ?? fallback;
  } catch (e) {
    console.error("Erreur de chargement", key, e);
    return fallback;
  }
}

export async function saveKey(key, value) {
  if (!supabase) return false;
  try {
    const { error } = await supabase.from("kv_store").upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) console.error("Erreur de sauvegarde", key, error);
    return !error;
  } catch (e) {
    console.error("Erreur de sauvegarde", key, e);
    return false;
  }
}
