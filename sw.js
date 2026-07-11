const CACHE_NAME = 'sorteurpro-v9';

// Fichiers mis en cache pour fonctionner hors ligne
const ASSETS = [
  '/',
  '/index.html',
];

// Installation : mise en cache des fichiers essentiels
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activation : supprime les anciens caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Interception des requêtes
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Les appels API (/api/*) : toujours réseau d'abord
  // Si pas de réseau, on répond avec une file d'attente locale
  if (url.pathname.startsWith('/api/')) {
    event.respondWith(
      fetch(event.request).catch(() => {
        // Si c'est une validation POST sans réseau, on répond OK quand même
        // (le résultat est déjà sauvegardé en localStorage côté app)
        if (event.request.method === 'POST') {
          return new Response(JSON.stringify({ ok: true, offline: true }), {
            headers: { 'Content-Type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ error: 'Hors ligne' }), {
          status: 503,
          headers: { 'Content-Type': 'application/json' }
        });
      })
    );
    return;
  }

  // Pour tout le reste : cache d'abord, puis réseau
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // Met en cache les nouvelles ressources statiques
        if (response.ok && event.request.method === 'GET') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        // Retourne la page principale si hors ligne
        return caches.match('/index.html');
      });
    })
  );
});
