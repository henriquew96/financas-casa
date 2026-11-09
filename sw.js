// Finanças da Casa — service worker: deixa o app abrir sem internet.
// Arquivos do app: rede primeiro (pega sempre a versão nova), cópia guardada se estiver offline.
// Bibliotecas dos CDNs (versão fixa): cópia guardada primeiro.
// Supabase (dados e login): sempre direto pela rede — nunca guardado aqui.
const VERSAO = 'financas-casa-v2';
const BASE = ['./', './index.html', './config.js', './manifest.webmanifest', './icones/icone-180.png', './icones/icone-192.png', './icones/icone-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(BASE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k.startsWith('financas-casa-') && k !== VERSAO).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === self.location.origin) {
    e.respondWith(
      fetch(req).then(r => {
        if (r.ok) { const copia = r.clone(); caches.open(VERSAO).then(c => c.put(req, copia)); }
        return r;
      }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('./index.html')))
    );
  } else if (url.hostname === 'cdn.jsdelivr.net' || url.hostname === 'cdnjs.cloudflare.com') {
    e.respondWith(
      caches.match(req).then(r => r || fetch(req).then(resp => {
        if (resp.ok) { const copia = resp.clone(); caches.open(VERSAO).then(c => c.put(req, copia)); }
        return resp;
      }))
    );
  }
});
