/**
 * sw.js — Service Worker para GAFI Dashboard PWA
 * Estrategia:
 *   - Shell (HTML, CSS, JS, fuentes): Cache-First con actualización en background.
 *   - Archivos Excel: Network-Only (datos siempre frescos, no se cachean).
 *   - Recursos CDN (Chart.js, SheetJS, FontAwesome): Cache-First.
 */

'use strict';

const CACHE_NAME    = 'gafi-dashboard-v2.0';
const SHELL_ASSETS  = [
    './',
    './index.html',
    './style.css',
    './script.js',
    './manifest.json',
    './logo-gafi.png',
];

const CDN_ORIGINS = [
    'fonts.googleapis.com',
    'fonts.gstatic.com',
    'cdnjs.cloudflare.com',
    'cdn.sheetjs.com',
    'cdn.jsdelivr.net',
];

/* ── INSTALL: pre-cachear el app shell ── */
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(SHELL_ASSETS))
            .then(() => self.skipWaiting())   // activar de inmediato
    );
});

/* ── ACTIVATE: limpiar caches viejos ── */
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys()
            .then(keys => Promise.all(
                keys
                    .filter(key => key !== CACHE_NAME)
                    .map(key => caches.delete(key))
            ))
            .then(() => self.clients.claim())
    );
});

/* ── FETCH: estrategia mixta ── */
self.addEventListener('fetch', event => {
    const { request } = event;
    const url         = new URL(request.url);

    // Ignorar solicitudes no-GET
    if (request.method !== 'GET') return;

    // Archivos Excel → siempre red (nunca cachear datos de usuario)
    if (/\.(xlsx|xls|xlsm|xlsb)$/i.test(url.pathname)) return;

    // Recursos CDN → Cache-First
    if (CDN_ORIGINS.some(o => url.hostname.includes(o))) {
        event.respondWith(cacheFirst(request));
        return;
    }

    // App shell → Stale-While-Revalidate
    event.respondWith(staleWhileRevalidate(request));
});

/* ── ESTRATEGIAS ── */

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
        return new Response('Sin conexión', { status: 503, statusText: 'Offline' });
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
        .catch(() => cached); // fallback a cache si no hay red

    return cached || fetchPromise;
}
