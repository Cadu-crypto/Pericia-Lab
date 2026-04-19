// PeríciaLab — Service Worker
// Versão do cache — incremente para forçar atualização
const CACHE_VERSION = 'pericia-v1';

// Arquivos que serão cacheados para uso offline
const ARQUIVOS_CACHE = [
  '/app-ipad-pericia.html',
  '/supabase_client.js',
  '/manifest.json',
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

// ── FETCH: serve do cache quando offline ─────────────────────
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

  // Arquivos do app: cache primeiro, rede como fallback
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      // Não está no cache — tenta a rede e cacheia para próxima vez
      return fetch(event.request).then(response => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        // Completamente offline e não estava no cache
        return new Response('App offline — arquivo não encontrado no cache.', {
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
