// public/sw.js — Service Worker per notifiche push

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('push', (event) => {
  if (!event.data) return;
  const data = event.data.json();
  const options = {
    body: data.corpo || 'Nuova notifica',
    icon: '/apple-touch-icon.png',
    badge: '/favicon-32x32.png',
    tag: data.id || 'default',
    requireInteraction: true,
    data: { url: data.link || '/' },
    // NOTA: il suono è gestito dal sistema operativo, non personalizzabile via web push
  };
  event.waitUntil(
    self.registration.showNotification(data.titolo || 'IlMelo', options)
  );
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.openWindow(event.notification.data?.url || '/')
  );
});