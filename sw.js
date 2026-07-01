// ============================================================
//  Fee Manager — Service Worker v1.0.48
//  STRATEGY:
//    HTML pages → NETWORK-FIRST
//    Static assets → CACHE-FIRST with background refresh
//    External CDNs → pass-through
//    version.json → NETWORK-ONLY
// ============================================================

const CACHE_VERSION = 'feemanager-v1.0.50';
const ASSETS = [
  './',
  './app.html',
  './css/app.css',
  './appstart/config.js',
  './appstart/license.js',
  './appstart/keystore.js',
  './appstart/schema.js',
  './appstart/translator.js',
  './appstart/appstart.js',
  './appstart/appstart.css',
  './js/api.js',
  './js/app.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/feeflow-logo.png',
  './icons/ai-logo.png',
  // CDN dependencies pre-caching:
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap',
  'https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap',
  'https://unpkg.com/@phosphor-icons/web',
  'https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.cache === 'only-if-cached' && e.request.mode !== 'same-origin') return;

  const url = new URL(e.request.url);

  // Allowed external CDN origins
  const ALLOWED_CDN_ORIGINS = [
    'https://fonts.googleapis.com',
    'https://fonts.gstatic.com',
    'https://unpkg.com',
    'https://cdn.jsdelivr.net',
    'https://cdnjs.cloudflare.com'
  ];

  const isAllowedCDN = ALLOWED_CDN_ORIGINS.some(origin => e.request.url.startsWith(origin));

  // External requests not in the whitelist → pass through
  if (url.origin !== self.location.origin && !isAllowedCDN) return;

  // version.json → always network
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(fetch(e.request));
    return;
  }

  // HTML (navigation) → NETWORK-FIRST
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then(response => {
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(c => c.put(e.request, clone));
          return response;
        })
        .catch(() => {
          return caches.match(e.request, { ignoreSearch: true })
            .then(cached => cached || caches.match('./app.html'));
        })
    );
    return;
  }

  // Static assets → CACHE-FIRST with background refresh
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(cached => {
      const networkFetch = fetch(e.request).then(response => {
        if (response && (response.ok || response.type === 'opaque')) {
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(c => c.put(e.request, clone));
        }
        return response;
      }).catch(() => null);
      return cached || networkFetch;
    })
  );
});

self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
