const CACHE = 'workout-v4';
const IMMUTABLE_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'lh3.googleusercontent.com'];

self.addEventListener('install', e => {
  // Do not pre-cache workout.html — always load it fresh from the network
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = e.request.url;

  // Fonts and GIFs: cache-first — these never change
  if (IMMUTABLE_HOSTS.some(h => url.includes(h))) {
    e.respondWith(
      caches.open(CACHE).then(async cache => {
        const cached = await cache.match(e.request);
        if (cached) return cached;
        const res = await fetch(e.request);
        if (res.ok) cache.put(e.request, res.clone());
        return res;
      })
    );
  }
  // workout.html: always network, no caching — ensures every open gets the latest
});
