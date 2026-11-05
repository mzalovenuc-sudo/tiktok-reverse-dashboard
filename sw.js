const CACHE_NAME = "daily-rmm-dashboard-v4";
const FILES = ["./index.html", "./tiktok_reverse_dashboard.html", "./manifest.webmanifest", "./icon.svg", "./icon-192.png", "./icon-512.png", "./xlsx.full.min.js"];
self.addEventListener("install", event => event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener("activate", event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith("daily-rmm-dashboard-") && key !== CACHE_NAME).map(key => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET" || new URL(event.request.url).origin !== self.location.origin) return;
  if (event.request.mode === "navigate") {
    event.respondWith(fetch(event.request).then(response => { caches.open(CACHE_NAME).then(cache=>cache.put(event.request,response.clone())); return response; }).catch(() => caches.match(event.request).then(cached=>cached||caches.match("./index.html"))));
  } else {
    event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
  }
});


