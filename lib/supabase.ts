import { createClient, type SupabaseClient } from "@supabase/supabase-js";

let client: SupabaseClient | null = null;

function getSupabaseClient(): SupabaseClient {
  if (client) return client;

  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const anonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!url || !anonKey) {
    throw new Error(
      "Supabase non configurato: impostare NEXT_PUBLIC_SUPABASE_URL e NEXT_PUBLIC_SUPABASE_ANON_KEY."
    );
  }

  client = createClient(url, anonKey);
  return client;
}

// Init lazy: il client viene creato al PRIMO uso, non al caricamento del modulo.
// - Con le env configurate il comportamento e' identico a prima.
// - Senza env (es. build/prerender) il modulo non lancia piu' al caricamento;
//   l'errore chiaro compare solo se il client viene realmente usato.
export const supabase = new Proxy({} as SupabaseClient, {
  get(_target, prop, receiver) {
    const c = getSupabaseClient();
    const value = Reflect.get(c as object, prop, receiver);
    return typeof value === "function" ? (value as (...a: unknown[]) => unknown).bind(c) : value;
  },
});
