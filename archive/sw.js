const CACHE_NAME = "irrigation-v3";
const ASSETS = ["./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
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
  if (event.request.method !== "GET") return;

  // Ne jamais intercepter les requêtes vers un autre domaine (ex. script.google.com,
  // utilisé pour la consultation en JSONP via une balise <script>). Le service worker
  // renvoyait ces requêtes vides (réponse opaque), ce qui cassait la consultation.
  // On laisse le navigateur les gérer nativement, sans passer par ici.
  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  // Pour la page HTML elle-même : toujours essayer le réseau en premier, pour être
  // sûr d'avoir la dernière version publiée. On ne se rabat sur le cache que si le
  // téléphone est hors-ligne. Ça évite de rester bloqué sur une vieille version.
  if (event.request.mode === "navigate" || event.request.url.endsWith("index.html")) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // Pour le reste (manifest, icônes) : cache d'abord, réseau en secours.
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});