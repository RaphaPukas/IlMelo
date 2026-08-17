import { serve } from 'https://deno.land/std@0.168.0/http/server.ts';
import webpush from 'https://esm.sh/web-push@3.6.6';

const VAPID_PUBLIC_KEY    = Deno.env.get('VAPID_PUBLIC_KEY')!;
const VAPID_PRIVATE_KEY   = Deno.env.get('VAPID_PRIVATE_KEY')!;
const SUPABASE_URL        = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders });
  }

  try {
    const { record } = await req.json();

    webpush.setVapidDetails(
      'mailto:admin@il-melo.it',
      VAPID_PUBLIC_KEY,
      VAPID_PRIVATE_KEY
    );

    // Fetch subscriptions, excluding the sender (inviata_da) so the operator
    // who created the report doesn't receive their own push.
    const senderFilter = record.inviata_da
      ? `&user_id=neq.${encodeURIComponent(record.inviata_da)}`
      : '';

    const subsRes = await fetch(
      `${SUPABASE_URL}/rest/v1/push_subscriptions?select=*${senderFilter}`,
      {
        headers: {
          'apikey': SUPABASE_SERVICE_ROLE_KEY,
          'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
        }
      }
    );

    if (!subsRes.ok) throw new Error(`Fetch subs failed: ${subsRes.status}`);
    const subs = await subsRes.json();
    if (!Array.isArray(subs) || !subs.length) {
      return new Response(JSON.stringify({ ok: true, sent: 0 }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }

    const payload = JSON.stringify({
      title: record.titolo,
      body:  record.corpo,
      icon:  'https://raphapukas.github.io/IlMelo/apple-touch-icon.png',
      badge: 'https://raphapukas.github.io/IlMelo/apple-touch-icon.png',
      tag:   record.id,
      data: {
        url: `https://raphapukas.github.io/IlMelo/?notifica=${encodeURIComponent(record.id)}`,
      },
    });

    const results = await Promise.allSettled(
      subs.map((sub) => {
        const pushSub = {
          endpoint: sub.endpoint,
          keys: { p256dh: sub.p256dh, auth: sub.auth },
        };
        return webpush.sendNotification(pushSub, payload);
      })
    );

    // Clean expired subscriptions (HTTP 410)
    const expired = results
      .map((r, i) => ({ r, i }))
      .filter(({ r }) => r.status === 'rejected' && r.reason?.statusCode === 410)
      .map(({ i }) => subs[i].endpoint);

    for (const ep of expired) {
      await fetch(
        `${SUPABASE_URL}/rest/v1/push_subscriptions?endpoint=eq.${encodeURIComponent(ep)}`,
        { method: 'DELETE', headers: { 'apikey': SUPABASE_SERVICE_ROLE_KEY, 'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
      );
    }

    const sent = results.filter(r => r.status === 'fulfilled').length;
    return new Response(JSON.stringify({ ok: true, sent, cleaned: expired.length }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    });

  } catch (err) {
    console.error('Push error:', err);
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    });
  }
});
