// public/sw.js — Service Worker notifiche push IlMelo

const APP_URL = 'https://raphapukas.github.io/IlMelo/';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('push', (event) => {
  if (!event.data) return;

  let data;

  try {
    data = event.data.json();
  } catch {
    data = {
      title: 'IlMelo',
      body: event.data.text()
    };
  }

  const options = {
    body: data.body || 'Nuova notifica',
    icon: data.icon || `${APP_URL}apple-touch-icon.png`,
    badge: data.badge || `${APP_URL}apple-touch-icon.png`,
    tag: data.tag || 'default',
    requireInteraction: true,
    data: data.data || {
      url: APP_URL
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

  const targetUrl =
    event.notification.data?.url || APP_URL;

  event.waitUntil(
    self.clients.matchAll({
      type: 'window',
      includeUncontrolled: true
    }).then(async (clients) => {

      // Se l'app è già aperta, aggiorna la pagina con l'URL
      // della notifica invece di limitarsi a fare focus.
      for (const client of clients) {
        if (client.url.startsWith(APP_URL)) {
          await client.navigate(targetUrl);
          return client.focus();
        }
      }

      // App non aperta: aprila direttamente sull'URL della notifica.
      return self.clients.openWindow(targetUrl);
    })
  );
});
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
