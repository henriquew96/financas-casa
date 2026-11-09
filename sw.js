// Finanças da Casa — service worker: deixa o app abrir sem internet.
// Só guarda arquivos do próprio site (rede primeiro, para pegar sempre a versão nova;
// a cópia guardada só é usada sem internet). Nada de outros sites é carregado.
// Dados e login (Supabase) passam direto pela rede — nunca ficam guardados aqui.
const VERSAO = 'financas-casa-v3';
const BASE = ['./', './index.html', './app.js', './config.js', './manifest.webmanifest', './libs/supabase.js',
  './icones/icone-180.png', './icones/icone-192.png', './icones/icone-512.png'];

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
  if (url.origin !== self.location.origin) return;   // Supabase e qualquer outro endereço: direto pela rede
  e.respondWith(
    fetch(req).then(r => {
      if (r.ok && r.type === 'basic') { const copia = r.clone(); caches.open(VERSAO).then(c => c.put(req, copia)); }
      return r;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('./index.html')))
  );
});
