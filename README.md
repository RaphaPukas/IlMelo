# Manutenzione

App di gestione manutenzione (moduli Mezzi e Carrozzine).

## Uso in locale

```bash
npm install
npm run dev
```

## Pubblicazione su GitHub Pages

1. Crea un repository su GitHub e carica questi file (vedi istruzioni ricevute in chat).
2. Su GitHub: **Settings → Pages → Source → GitHub Actions**.
3. Ad ogni push sul branch `main`, il workflow in `.github/workflows/deploy.yml`
   compila l'app e la pubblica automaticamente.
4. L'app sarà visibile su `https://<tuo-utente>.github.io/<nome-repo>/`.

## Dati

I dati (veicoli, carrozzine, interventi) restano salvati nel browser di chi apre
l'app (localStorage), non su un server: ogni dispositivo/browser ha la propria copia.
