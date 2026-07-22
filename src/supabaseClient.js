import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.error(
    'Configurazione Supabase mancante: imposta VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY ' +
    '(in un file .env in locale, o come secret di GitHub Actions per la build online).'
  );
}

export const supabase = createClient(supabaseUrl || '', supabaseAnonKey || '');
