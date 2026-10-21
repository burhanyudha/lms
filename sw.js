const CACHE_NAME = 'al-ilmi-lms-v4';
const ASSETS_TO_CACHE = [
  'index.html',
  'dashboard.html',
  'lms-api.js',
  'pwa-handler.js',
  'manifest.json',
  'icon.svg',
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // Use a more resilient approach: try to cache everything, 
      // but don't let one missing file break the whole installation
      return Promise.allSettled(
        ASSETS_TO_CACHE.map(url => 
          cache.add(url).catch(err => console.warn(`Failed to cache: ${url}`, err))
        )
      );
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Clone the response and cache it for offline use
        const responseToCache = response.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(event.request, responseToCache);
        });
        return response;
      })
      .catch(() => {
        // Fallback to cache if network fails (offline)
        return caches.match(event.request);
      })
  );
});
