// service-worker.js - GAFI Ferrelectrico PWA
const CACHE_NAME = 'gafi-dashboard-v1';

// Archivos que se guardan en caché para funcionar offline
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './logo-gafi.png',
  './icon-192x192.png',
  './icon-512x512.png',
  // Librerías externas (se cachean la primera vez que se usan con internet)
  'https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;400;500;600;700&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css',
  'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// ── INSTALACIÓN: guarda todos los archivos en caché ──
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[SW] Instalando y cacheando archivos...');
      // Cachear archivos locales (críticos)
      const localAssets = ASSETS_TO_CACHE.filter(url => !url.startsWith('http'));
      return cache.addAll(localAssets).then(() => {
        // Cachear recursos externos uno por uno (no fallar si alguno falla)
        const externalAssets = ASSETS_TO_CACHE.filter(url => url.startsWith('http'));
        return Promise.allSettled(
          externalAssets.map(url =>
            fetch(url).then(res => {
              if (res.ok) return cache.put(url, res);
            }).catch(() => console.warn('[SW] No se pudo cachear:', url))
          )
        );
      });
    }).then(() => {
      console.log('[SW] Instalación completa.');
      return self.skipWaiting();
    })
  );
});

// ── ACTIVACIÓN: elimina cachés viejos ──
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter(name => name !== CACHE_NAME)
          .map(name => {
            console.log('[SW] Eliminando caché viejo:', name);
            return caches.delete(name);
          })
      );
    }).then(() => {
      console.log('[SW] Activado. Tomando control de la app.');
      return self.clients.claim();
    })
  );
});

// ── FETCH: estrategia Cache First → Network Fallback ──
self.addEventListener('fetch', (event) => {
  // Solo interceptar peticiones GET
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      // Si está en caché, devuelve el caché inmediatamente
      if (cachedResponse) {
        return cachedResponse;
      }
      // Si no está en caché, busca en red y lo guarda para la próxima vez
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.ok) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }
        return networkResponse;
      }).catch(() => {
        // Sin internet y sin caché: mostrar página de fallback si es HTML
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
