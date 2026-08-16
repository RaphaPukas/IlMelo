import { serve } from 'https://deno.land/std@0.168.0/http/server.ts';
import webpush from 'https://esm.sh/web-push@3.6.6';

const VAPID_PUBLIC_KEY = Deno.env.get('VAPID_PUBLIC_KEY')!;
const VAPID_PRIVATE_KEY = Deno.env.get('VAPID_PRIVATE_KEY')!;
const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;

// CORS headers for browser preflight
const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

serve(async (req) => {
  // 1. Handle CORS preflight (OPTIONS)
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders });
  }

  try {
    const { record } = await req.json();

    // 2. Setup web-push with VAPID keys
    webpush.setVapidDetails(
      'mailto:admin@il-melo.it',
      VAPID_PUBLIC_KEY,
      VAPID_PRIVATE_KEY
    );

    // 3. Fetch all push subscriptions from DB
    const subsRes = await fetch(`${SUPABASE_URL}/rest/v1/push_subscriptions?select=*`, {
      headers: {
        'apikey': SUPABASE_SERVICE_ROLE_KEY,
        'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      }
    });

    if (!subsRes.ok) {
      throw new Error(`Failed to fetch subscriptions: ${subsRes.status}`);
    }

    const subs = await subsRes.json();
    if (!Array.isArray(subs) || !subs.length) {
      return new Response(JSON.stringify({ ok: true, sent: 0, message: 'No subscriptions' }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }

    // 4. Build notification payload
    const payload = JSON.stringify({
      title: record.titolo,
      body: record.corpo,
      icon: '/apple-touch-icon.png',
      badge: '/apple-touch-icon.png',
      tag: record.id,
      requireInteraction: true,
      data: {
        url: `/?notifica=${record.id}`,
        notificaId: record.id,
        tipo: record.tipo,
      }
    });

    // 5. Send push to all subscriptions
    const results = await Promise.allSettled(
      subs.map((sub) => {
        const pushSub = {
          endpoint: sub.endpoint,
          keys: {
            p256dh: sub.p256dh,
            auth: sub.auth,
          }
        };
        return webpush.sendNotification(pushSub, payload);
      })
    );

    // 6. Clean up expired subscriptions (HTTP 410 Gone)
    const expiredEndpoints = results
      .map((r, i) => ({ result: r, index: i }))
      .filter(({ result }) => result.status === 'rejected' && result.reason?.statusCode === 410)
      .map(({ index }) => subs[index].endpoint);

    for (const endpoint of expiredEndpoints) {
      await fetch(
        `${SUPABASE_URL}/rest/v1/push_subscriptions?endpoint=eq.${encodeURIComponent(endpoint)}`,
        {
          method: 'DELETE',
          headers: {
            'apikey': SUPABASE_SERVICE_ROLE_KEY,
            'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
          }
        }
      );
    }

    const sent = results.filter(r => r.status === 'fulfilled').length;
    const failed = results.filter(r => r.status === 'rejected').length;

    return new Response(JSON.stringify({ ok: true, sent, failed, cleaned: expiredEndpoints.length }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    });

  } catch (err) {
    console.error('Push notification error:', err);
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    });
  }
});
