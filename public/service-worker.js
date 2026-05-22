const CACHE_NAME = "openframe-internal-v1";
const STATIC_ASSETS = [
  "/apple-touch-icon.png",
  "/favicon-32.png",
  "/icons/icon-192.png",
  "/icons/icon-512.png",
  "/install.js",
  "/openframe-logo.png",
  "/site.webmanifest",
  "/styles.css"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(STATIC_ASSETS))
      .catch(() => null)
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  if (event.request.method !== "GET" || url.origin !== self.location.origin || url.pathname.startsWith("/api/")) {
    return;
  }

  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request).then((cached) => cached || Response.error()))
  );
});
