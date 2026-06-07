'use strict';

const CACHE_NAME   = 'gafi-dashboard-v3';
const SHELL_ASSETS = [
    './index.html',
    './style.css',
    './script.js',
    './manifest.json',
    './logo-gafi.png',
    './icons/icon-72.png',
    './icons/icon-96.png',
    './icons/icon-128.png',
    './icons/icon-144.png',
    './icons/icon-152.png',
    './icons/icon-192.png',
    './icons/icon-384.png',
    './icons/icon-512.png',
];

const CDN_ORIGINS = [
    'fonts.googleapis.com',
    'fonts.gstatic.com',
    'cdnjs.cloudflare.com',
    'cdn.sheetjs.com',
    'cdn.jsdelivr.net',
];

/* ── INSTALL ── */
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                // Cachear cada asset individualmente para que un error
                // en uno no bloquee toda la instalación
                return Promise.allSettled(
                    SHELL_ASSETS.map(url =>
                        cache.add(url).catch(err =>
                            console.warn('[SW] No se pudo cachear:', url, err)
                        )
                    )
                );
            })
            .then(() => self.skipWaiting())
    );
});

/* ── ACTIVATE ── */
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys()
            .then(keys => Promise.all(
                keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
            ))
            .then(() => self.clients.claim())
    );
});

/* ── FETCH ── */
self.addEventListener('fetch', event => {
    const { request } = event;
    const url = new URL(request.url);

    if (request.method !== 'GET') return;

    // Archivos Excel — nunca cachear
    if (/\.(xlsx|xls|xlsm|xlsb)$/i.test(url.pathname)) return;

    // CDN — Cache First
    if (CDN_ORIGINS.some(o => url.hostname.includes(o))) {
        event.respondWith(cacheFirst(request));
        return;
    }

    // App shell — Stale While Revalidate
    event.respondWith(staleWhileRevalidate(request));
});

async function cacheFirst(request) {
    const cached = await caches.match(request);
    if (cached) return cached;
    try {
        const response = await fetch(request);
        if (response.ok) {
            const cache = await caches.open(CACHE_NAME);
            cache.put(request, response.clone());
        }
        return response;
    } catch {
        return new Response('Sin conexión', { status: 503 });
    }
}

async function staleWhileRevalidate(request) {
    const cache  = await caches.open(CACHE_NAME);
    const cached = await cache.match(request);
    const fetchPromise = fetch(request)
        .then(response => {
            if (response.ok) cache.put(request, response.clone());
            return response;
        })
        .catch(() => cached);
    return cached || fetchPromise;
}
