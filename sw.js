/**
 * FNRG Portal — Service Worker
 * App shell: cache-first (instant loads, works offline).
 * API calls (script.google.com): network-only, never cached.
 */

const CACHE_NAME = "fnrg-portal-v38";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./manifest.json",
  "./icon.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Never intercept API traffic — data must always be live
  if (url.hostname.includes("script.google.com") ||
      url.hostname.includes("googleusercontent.com")) {
    return;
  }

  if (event.request.method !== "GET") return;

  // Network-First with Cache Fallback for instant updates when online
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (response && response.ok && url.origin === self.location.origin) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return response;
      })
      .catch(() => {
        return caches.match(event.request);
      })
  );
});
