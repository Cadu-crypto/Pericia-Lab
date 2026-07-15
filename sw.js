// PeríciaLab — Service Worker
const CACHE_VERSION = 'pericia-v4';

const ARQUIVOS_CACHE = [
  '/Pericia-Lab/app-ipad-pericia.html',
  '/Pericia-Lab/supabase_client.js',
  '/Pericia-Lab/manifest.json',
  '/Pericia-Lab/icon-192.png',
  '/Pericia-Lab/icon-512.png',
  '/Pericia-Lab/apple-touch-icon.png',
  '/Pericia-Lab/index.html',
];

// ── INSTALL: baixa e cacheia todos os arquivos ────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(cache => {
      console.log('[SW] Instalando cache offline...');
      return cache.addAll(ARQUIVOS_CACHE);
    }).then(() => {
      console.log('[SW] Cache instalado com sucesso.');
      return self.skipWaiting(); // Ativa imediatamente
    })
  );
});

// ── ACTIVATE: limpa caches antigos ───────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== CACHE_VERSION)
          .map(k => { console.log('[SW] Removendo cache antigo:', k); return caches.delete(k); })
      )
    ).then(() => self.clients.claim())
  );
});

// ── FETCH: rede primeiro, cache só como fallback (offline de verdade) ────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Requisições ao Supabase: sempre vai para a rede (sync em background)
  if (url.hostname.includes('supabase.co')) {
    event.respondWith(
      fetch(event.request).catch(() => {
        // Sem internet — ignora silenciosamente (sync acontece quando voltar)
        return new Response(JSON.stringify({ error: 'offline' }), {
          headers: { 'Content-Type': 'application/json' },
        });
      })
    );
    return;
  }

  // Arquivos do app: tenta a rede primeiro para sempre pegar a versão mais
  // nova quando há internet. Só cai pro cache se a rede falhar de verdade
  // (offline real) — antes era "cache primeiro", que fazia o app nunca
  // buscar atualizações depois da primeira instalação.
  event.respondWith(
    fetch(event.request).then(response => {
      if (response.ok) {
        const clone = response.clone();
        caches.open(CACHE_VERSION).then(cache => cache.put(event.request, clone));
      }
      return response;
    }).catch(() => {
      return caches.match(event.request).then(cached => {
        return cached || new Response('App offline — arquivo não encontrado no cache.', {
          status: 503,
          headers: { 'Content-Type': 'text/plain' },
        });
      });
    })
  );
});

// ── SYNC EM BACKGROUND (quando reconecta) ───────────────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-diligencias') {
    console.log('[SW] Background sync: enviando diligências pendentes...');
    // A sincronização real acontece no app via syncToCloud()
  }
});
