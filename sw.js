const CACHE = 'reembolso-igd-v2';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './index.html', './manifest.json'])));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if(e.request.method !== 'GET') return;
  // Só arquivos do próprio app. Chamadas ao Google (Drive, login) passam direto:
  // interceptá-las não ajuda (não dá pra usar cache delas) e, se o fetch do SW
  // falhasse, caches.match devolvia vazio e a chamada virava erro de rede.
  if(new URL(e.request.url).origin !== self.location.origin) return;
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
