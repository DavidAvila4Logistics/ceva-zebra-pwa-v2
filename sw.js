/**
 * CEVA Zebra PWA v2 — Service Worker
 * Estrategia: Cache-first para assets estáticos, network-only para APIs externas.
 * Compatible con GitHub Pages (sin backend).
 */

const CACHE_NAME = "ceva-zebra-v2-cache-v5";

// Assets to pre-cache on install
const PRECACHE_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./assets/css/styles.css",
  "./assets/js/app.js",
  "./assets/img/ceva-logo.jpg",
  "./assets/img/icon-192.png",
  "./assets/img/icon-512.png",
  // CDN libraries are NOT pre-cached (they come from external hosts)
  // The app still works offline for UI but needs network for CDN libs first load
];

// Install: pre-cache static assets
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => {
        return cache.addAll(PRECACHE_ASSETS).catch((err) => {
          console.warn("[SW] Pre-cache parcial (algunos assets pueden no estar disponibles):", err);
        });
      })
      .then(() => {
        return self.skipWaiting();
      })
  );
});

// Activate: remove old caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => {
        return Promise.all(
          keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
        );
      })
      .then(() => self.clients.claim())
  );
});

// Fetch: cache-first for same-origin, network-first for CDN
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // For same-origin requests: cache-first strategy
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        if (cached) return cached;
        return fetch(event.request)
          .then((response) => {
            if (response && response.status === 200) {
              const clone = response.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
            }
            return response;
          })
          .catch(() => {
            // Fallback: return index.html for navigation requests
            if (event.request.mode === "navigate") {
              return caches.match("./index.html");
            }
          });
      })
    );
    return;
  }

  // For CDN/external requests: network-first, cache on success
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (response && response.status === 200 && event.request.method === "GET") {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        return caches.match(event.request);
      })
  );
});
