// PlanMetric Pro — Service Worker
// Stratégie : l'interface (coquille de l'app) est mise en cache pour un
// lancement instantané et un fonctionnement hors-ligne de l'écran d'accueil ;
// les appels /api/ passent toujours par le réseau (données métier à jour),
// avec repli sur le cache uniquement pour les fichiers statiques.
const CACHE = "planmetric-v1";
const COQUILLE = ["/", "/index.html", "/manifest.webmanifest",
  "/icons/icon-192.png", "/icons/icon-512.png", "/icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(COQUILLE)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(cles => Promise.all(
    cles.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (url.pathname.startsWith("/api/")) { e.respondWith(fetch(e.request)); return; } // réseau seul pour le métier
  e.respondWith(
    fetch(e.request).then(rep => {
      const copie = rep.clone();
      caches.open(CACHE).then(c => c.put(e.request, copie));
      return rep;
    }).catch(() => caches.match(e.request).then(r => r || caches.match("/index.html")))
  );
});
