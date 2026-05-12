const CACHE = 'workout-v2';

// Cache the app shell on install
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(['./workout.html']))
  );
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

// Runtime caching: serve from cache if available, fetch and cache if not
self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Only cache GET requests for Google Fonts and Google Drive (GIFs)
  if (
    e.request.method !== 'GET' ||
    (!url.includes('fonts.googleapis.com') &&
     !url.includes('fonts.gstatic.com') &&
     !url.includes('lh3.googleusercontent.com') &&
     !url.includes('workout.html'))
  ) return;

  e.respondWith(
    caches.open(CACHE).then(async cache => {
      const cached = await cache.match(e.request);
      if (cached) return cached;

      try {
        const response = await fetch(e.request);
        if (response.ok) cache.put(e.request, response.clone());
        return response;
      } catch {
        return cached || new Response('Offline', { status: 503 });
      }
    })
  );
});
