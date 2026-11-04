const CACHE_NAME = "tiktok-reverse-dashboard-v5";
const FILES = ["./", "./index.html", "./tiktok_reverse_dashboard.html", "./manifest.webmanifest", "./icon.svg", "./icon-192.png", "./icon-512.png", "./xlsx.full.min.js"];
self.addEventListener("install", event => event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener("activate", event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith("tiktok-reverse-dashboard-") && key !== CACHE_NAME).map(key => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET" || new URL(event.request.url).origin !== self.location.origin) return;
  if (event.request.mode === "navigate") {
    event.respondWith(fetch(event.request).then(response => { const copy=response.clone(); caches.open(CACHE_NAME).then(cache=>cache.put("./tiktok_reverse_dashboard.html",copy)); return response; }).catch(() => caches.match("./index.html")));
  } else {
    event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
  }
});
