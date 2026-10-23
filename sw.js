// Service worker: guarda la app en el celular para que abra rápido y sin señal.
// Al publicar cambios, subir VERSION para que los celulares descarguen la nueva versión.
const VERSION = 'hc-v13';
const MSAL = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.12.0/lib/msal-browser.min.js';
const DOCX = 'https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js';
const FIREBASE = ['app', 'auth', 'firestore'].map(m => `https://www.gstatic.com/firebasejs/10.14.1/firebase-${m}-compat.js`);
const SHELL = [
  './', 'index.html', 'styles.css', 'app.js', 'informe.js', 'config.js', 'catalogos.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', MSAL, DOCX, ...FIREBASE
];
const EXTERNOS = [MSAL, DOCX, ...FIREBASE];

self.addEventListener('install', e => {
  // Se guarda cada archivo por separado: si alguno falta, el resto igual queda disponible sin señal.
  e.waitUntil(caches.open(VERSION)
    .then(c => Promise.allSettled(SHELL.map(u => c.add(u))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const propio = url.origin === location.origin;
  const fuentes = url.host === 'fonts.googleapis.com' || url.host === 'fonts.gstatic.com';
  if (!propio && !EXTERNOS.includes(req.url) && !fuentes) return; // Graph y login van directo a la red

  // Abrir la app: siempre index.html (también al volver del login con ?code=...)
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).catch(() => caches.match('index.html')));
    return;
  }
  // Resto: responde desde el celular y actualiza en segundo plano
  e.respondWith(caches.open(VERSION).then(async cache => {
    const guardado = await cache.match(req);
    const red = fetch(req).then(r => { if (r.ok) cache.put(req, r.clone()); return r; }).catch(() => guardado);
    return guardado || red;
  }));
});
