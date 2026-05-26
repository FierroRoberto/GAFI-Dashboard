// service-worker.js — GAFI Ferrelectrico PWA v2
const CACHE_NAME = 'gafi-dashboard-v2.78';

const LOCAL_ASSETS = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './logo-gafi.png',
  './icon-192x192.png',
  './icon-512x512.png',
  './icon-180x180.png',
  './icon-32x32.png',
  './icon-16x16.png'
];

const EXTERNAL_ASSETS = [
  'https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;400;500;600;700&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css',
  'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// ── INSTALACIÓN ──────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      // Primero cachear archivos locales (críticos, deben existir)
      await cache.addAll(LOCAL_ASSETS);
      // Cachear externos uno por uno sin fallar si alguno no está disponible
      await Promise.allSettled(
        EXTERNAL_ASSETS.map(url =>
          fetch(url).then(res => { if (res.ok) return cache.put(url, res); })
                    .catch(() => console.warn('[SW] No cacheado:', url))
        )
      );
      console.log('[SW] Instalación completa — caché lista.');
    }).then(() => self.skipWaiting())
  );
});

// ── ACTIVACIÓN: limpia cachés viejos ────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => {
      console.log('[SW] Activado. Tomando control.');
      return self.clients.claim();
    })
  );
});

// ── FETCH: Cache First → Network Fallback ───────────────────
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
