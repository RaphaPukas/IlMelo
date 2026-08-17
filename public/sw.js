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
    body: data.body || 'Nuova notifica',
    icon: data.icon || 'https://raphapukas.github.io/IlMelo/apple-touch-icon.png',
    badge: data.badge || 'https://raphapukas.github.io/IlMelo/apple-touch-icon.png',
    tag: data.tag || 'default',
    requireInteraction: true,
    data: data.data || {
      url: 'https://raphapukas.github.io/IlMelo/'
    }
  };

  event.waitUntil(
    self.registration.showNotification(
      data.title || 'IlMelo',
      options
    )
  );
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  const url =
    event.notification.data?.url ||
    'https://raphapukas.github.io/IlMelo/';

  event.waitUntil(
    self.clients.matchAll({
      type: 'window',
      includeUncontrolled: true
    }).then((clients) => {
      for (const client of clients) {
        if (client.url.startsWith('https://raphapukas.github.io/IlMelo/')) {
          return client.focus();
        }
      }

      return self.clients.openWindow(url);
    })
  );
});
