# Fix Edge Function Push Notification

## Problema risolto

L'Edge Function `push-notification` rispondeva 404 alle richieste OPTIONS (CORS preflight) e usava `SUPABASE_ANON_KEY` non definito invece di `SUPABASE_SERVICE_ROLE_KEY`. Inoltre non usava la libreria `web-push` per la firma VAPID.

## File corretto

Sostituisci `supabase/functions/push-notification/index.ts` con il file fornito.

## Variabili d'ambiente richieste su Supabase

Vai su **Supabase Dashboard → Project Settings → Edge Functions** e aggiungi:

| Variabile | Valore | Dove trovarlo |
|-----------|--------|---------------|
| `VAPID_PUBLIC_KEY` | Chiave pubblica VAPID | Generata con `npx web-push generate-vapid-keys` |
| `VAPID_PRIVATE_KEY` | Chiave privata VAPID | Stesso comando sopra |
| `SUPABASE_URL` | `https://<project>.supabase.co` | Project Settings → API |
| `SUPABASE_SERVICE_ROLE_KEY` | `service_role` key | Project Settings → API → service_role key |

⚠️ **NON usare** `SUPABASE_ANON_KEY` per leggere le subscriptions: serve il `service_role` per bypassare RLS.

## Deploy

```bash
# 1. Installa Supabase CLI (se non ce l'hai)
npm install -g supabase

# 2. Login
supabase login

# 3. Link al progetto
supabase link --project-ref <TUO-PROJECT-REF>

# 4. Deploy la funzione
supabase functions deploy push-notification

# 5. Verifica i log in tempo reale
supabase functions logs push-notification --tail
```

## Test rapido

Dall'app, invia una segnalazione. Poi controlla i log:

```bash
supabase functions logs push-notification --tail
```

Dovresti vedere:
- `OPTIONS /functions/v1/push-notification → 200` (CORS ok)
- `POST /functions/v1/push-notification → 200` con `{ ok: true, sent: N }`

## Generare chiavi VAPID (se non le hai)

```bash
npx web-push generate-vapid-keys
```

Copia la **public key** in:
- Variabile d'ambiente `VAPID_PUBLIC_KEY` su Supabase Edge Function
- `VITE_VAPID_PUBLIC_KEY` nel frontend (GitHub Secrets per il deploy)

Copia la **private key** in:
- Variabile d'ambiente `VAPID_PRIVATE_KEY` su Supabase Edge Function

## Troubleshooting

| Errore | Causa | Soluzione |
|--------|-------|-----------|
| `OPTIONS 404` | CORS non gestito | Usa il nuovo `index.ts` con handler OPTIONS |
| `SUPABASE_ANON_KEY is not defined` | Variabile mancante | Usa `SUPABASE_SERVICE_ROLE_KEY` |
| `sent: 0` | Nessuna subscription | L'utente deve accettare le notifiche push nel browser |
| `410 Gone` | Subscription scaduta | La funzione le pulisce automaticamente |
| `VAPID keys not set` | Manca env var | Controlla VAPID_PUBLIC_KEY e VAPID_PRIVATE_KEY su Supabase |
