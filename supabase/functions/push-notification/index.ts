import { serve } from 'https://deno.land/std@0.168.0/http/server.ts';

const VAPID_PUBLIC_KEY = Deno.env.get('VAPID_PUBLIC_KEY')!;
const VAPID_PRIVATE_KEY = Deno.env.get('VAPID_PRIVATE_KEY')!;
const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;

serve(async (req) => {
  const { record } = await req.json();
  
  // Recupera tutte le sottoscrizioni push
  const subsRes = await fetch(`${SUPABASE_URL}/rest/v1/push_subscriptions?select=*`, {
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`
    }
  });
  const subs = await subsRes.json();
  if (!Array.isArray(subs) || !subs.length) return new Response('OK');

  const payload = JSON.stringify({
    titolo: record.titolo,
    corpo: record.corpo,
    id: record.id,
    link: `/?notifica=${record.id}`
  });

  const results = await Promise.allSettled(
    subs.map((sub) =>
      fetch(sub.endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + VAPID_PUBLIC_KEY, // Simplified; for production use web-push lib
        },
        body: payload
      })
    )
  );

  return new Response('OK');
});
