import { serve } from 'https://deno.land/std@0.168.0/http/server.ts';
import webPush from 'https://esm.sh/web-push@3.6.6?target=deno';

const VAPID_PUBLIC_KEY = Deno.env.get('VAPID_PUBLIC_KEY')!;
const VAPID_PRIVATE_KEY = Deno.env.get('VAPID_PRIVATE_KEY')!;
const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const SUPABASE_ANON_KEY = Deno.env.get('SUPABASE_ANON_KEY')!;

webPush.setVapidDetails(
  'mailto:fantagalla@libero.it',
  VAPID_PUBLIC_KEY,
  VAPID_PRIVATE_KEY
);

serve(async (req) => {
  const { record } = await req.json();

  // 1. Recupera tutte le subscription push dal database
  const subsRes = await fetch(`${SUPABASE_URL}/rest/v1/push_subscriptions?select=*`, {
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`
    }
  });

  if (!subsRes.ok) {
    console.error('Errore fetch subscriptions:', await subsRes.text());
    return new Response('Error', { status: 500 });
  }

  const subs = await subsRes.json();
  if (!Array.isArray(subs) || subs.length === 0) {
    return new Response('No subscriptions');
  }

  // 2. Prepara il payload della notifica
  const payload = JSON.stringify({
    titolo: record.titolo,
    corpo: record.corpo,
    id: record.id,
    link: `/?notifica=${record.id}`
  });

  // 3. Invia la push a tutti i dispositivi registrati
  const results = await Promise.allSettled(
    subs.map((sub) =>
      webPush.sendNotification(
        {
          endpoint: sub.endpoint,
          keys: {
            p256dh: sub.p256dh,
            auth: sub.auth
          }
        },
        payload
      )
    )
  );

  // 4. Rimuovi subscription scadute (410 Gone = dispositivo non più raggiungibile)
  const dead = subs.filter((_, i) => {
    const r = results[i];
    return r.status === 'rejected' && (r.reason?.statusCode === 410 || r.reason?.statusCode === 404);
  });

  if (dead.length > 0) {
    const endpointList = dead.map(s => `"${s.endpoint}"`).join(',');
    await fetch(`${SUPABASE_URL}/rest/v1/push_subscriptions?endpoint=in.(${endpointList})`, {
      method: 'DELETE',
      headers: {
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`
      }
    });
  }

  return new Response('OK');
});
