import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabaseConfigured = !!(supabaseUrl && supabaseAnonKey);

if (!supabaseConfigured) {
  // eslint-disable-next-line no-console
  console.error(
    'Configurazione Supabase mancante: imposta VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY ' +
    '(in un file .env in locale, o come secret di GitHub Actions per la build online).'
  );
}

// Con credenziali mancanti, createClient('','') lancia un errore SUBITO al
// caricamento del modulo e manda in crash l'intera pagina (schermo bianco).
// Uso un URL segnaposto valido solo per evitare il crash: l'app mostrera'
// comunque un messaggio chiaro invece di provare a usarlo.
export const supabase = createClient(
  supabaseConfigured ? supabaseUrl : 'https://placeholder.supabase.co',
  supabaseConfigured ? supabaseAnonKey : 'placeholder-non-configurato'
);
