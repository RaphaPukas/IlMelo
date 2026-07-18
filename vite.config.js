import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// base: './' funziona sia su GitHub Pages (in un sottopercorso tipo
// tuo-utente.github.io/nome-repo/) sia in anteprima locale.
export default defineConfig({
  plugins: [react()],
  base: './',
});
