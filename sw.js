/* ============================================================
   sw.js — GAFI Dashboard PWA  |  Service Worker único v4.1
   Estrategia: Cache-First shell + Network-First datos
   ============================================================ */
'use strict';

const CACHE = 'gafi-v4.1';

/* Assets del shell — se cachean en install.
   Promise.allSettled garantiza que un asset faltante
   (ej. logo-gafi.png en primer deploy) no rompe toda la instalación. */
const SHELL = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './icons/icon-72.png',
  './icons/icon-96.png',
  './icons/icon-128.png',
  './icons/icon-144.png',
  './icons/icon-152.png',
  './icons/icon-192.png',
  './icons/icon-384.png',
  './icons/icon-512.png',
];

const CDN = [
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'cdnjs.cloudflare.com',
  'cdn.sheetjs.com',
  'cdn.jsdelivr.net',
];

/* ── INSTALL ─────────────────────────────────────────────── */
self.addEventListener('install', ev => {
  ev.waitUntil(
    caches.open(CACHE).then(cache =>
      Promise.allSettled(SHELL.map(url => cache.add(url)))
    ).then(() => self.skipWaiting())
  );
});

/* ── ACTIVATE: borrar caches viejos ─────────────────────── */
self.addEventListener('activate', ev => {
  ev.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

/* ── FETCH ───────────────────────────────────────────────── */
self.addEventListener('fetch', ev => {
  const req = ev.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  /* Archivos Excel → nunca cachear, siempre red */
  if (/\.(xlsx|xls|xlsm|xlsb)$/i.test(url.pathname)) return;

  /* CDN externas → Cache-First */
  if (CDN.some(d => url.hostname.includes(d))) {
    ev.respondWith(cacheFirst(req));
    return;
  }

  /* Shell local → Stale-While-Revalidate (offline-first) */
  ev.respondWith(swr(req));
});

async function cacheFirst(req) {
  const hit = await caches.match(req);
  if (hit) return hit;
  try {
    const res = await fetch(req);
    if (res.ok) (await caches.open(CACHE)).put(req, res.clone());
    return res;
  } catch {
    return new Response('Sin conexión', { status: 503 });
  }
}

async function swr(req) {
  const cache = await caches.open(CACHE);
  const hit   = await cache.match(req);
  const net   = fetch(req)
    .then(res => { if (res.ok) cache.put(req, res.clone()); return res; })
    .catch(() => hit);
  return hit || net;
}
