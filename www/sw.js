// Service Worker do painel Extreme Wind — Logística & Frotas
// Estrategia: "network-first" para as paginas HTML (sempre tenta buscar a versao
// mais recente, ja que os dashboards sao atualizados diariamente), com fallback
// pro cache quando estiver offline. Para imagens/assets estaticos, usa cache-first
// (nao mudam com frequencia).

const CACHE_NAME = "extreme-wind-v1";
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./dashboard-resumo-equipes.html",
  "./dashboard-frota-manutencao.html",
  "./manifest.json",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/logo-extreme-wind-white.png",
  "./assets/logo-extreme-wind-color.png",
  "./assets/icone-helice-white.png",
  "./assets/icone-helice-color.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE_URLS)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const isHTML = req.mode === "navigate" || (req.headers.get("accept") || "").includes("text/html");

  if (isHTML) {
    // network-first: dados de hoje em primeiro lugar, cache so como reserva offline
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req).then((cached) => cached || caches.match("./index.html")))
    );
  } else {
    // cache-first pra imagens/estaticos
    event.respondWith(
      caches.match(req).then((cached) => cached || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        return res;
      }).catch(() => cached))
    );
  }
});
