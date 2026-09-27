// Service Worker for Delivery Navi PWA (Network-First strategy to ensure latest privacy-cleared code)
const CACHE_NAME = 'delivery-navi-v8-fix';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './xlsx.full.min.js',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Network-First for HTML/Manifest to immediately update UI, Cache-Fallback for offline
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  
  // HTML or navigation requests: Network first, fall back to cache
  if (event.request.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/')) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const resClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
          return response;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Other assets: Stale-while-revalidate or Cache-first
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const fetchPromise = fetch(event.request).then((networkResponse) => {
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, networkResponse.clone()));
        return networkResponse;
      }).catch(() => null);

      return cachedResponse || fetchPromise;
    })
  );
});
