// Service worker H.S.E Motors — PWA
// Stratégie : réseau en priorité partout (l'app doit toujours refléter les
// dernières données), avec un filet de sécurité en cache pour l'écran de
// connexion et les fichiers statiques si le réseau est momentanément coupé.
// Les appels /api/* ne sont JAMAIS mis en cache (données toujours fraîches).

const CACHE_VERSION = "hse-motors-v1";
const APP_SHELL = [
  "/",
  "/manifest.webmanifest",
  "/logo.png",
  "/logo-icon.png",
  "/icon-192.png",
  "/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((noms) =>
      Promise.all(noms.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Jamais de cache pour l'API : toujours des données à jour.
  if (url.pathname.startsWith("/api/")) {
    return;
  }

  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((reponse) => {
        const copie = reponse.clone();
        caches.open(CACHE_VERSION).then((cache) => cache.put(event.request, copie));
        return reponse;
      })
      .catch(() => caches.match(event.request).then((r) => r || caches.match("/")))
  );
});
