/* Compta Mobile — service worker (offline-first)
 * Network-first pour le shell (version fraîche à CHAQUE ouverture, repli
 * cache hors ligne) ; GET same-origin uniquement ;
 * JAMAIS de cache pour les appels API externes (Graph / OneDrive) — toujours réseau.
 * Version : v1 (bump CACHE pour invalider).
 */
"use strict";
const CACHE = "compta-shell-v21";
const SHELL = [
  "./", "./index.html", "./manifest.webmanifest", "./icon.svg",
  "./css/styles.css", "./js/config.js", "./js/db.js", "./js/xlsx-surgery.js",
  "./js/graph.js", "./js/link.js", "./js/sync.js", "./js/app.js"
];

self.addEventListener("install", function (ev) {
  ev.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (ev) {
  ev.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.filter(function (k) { return k !== CACHE; })
                               .map(function (k) { return caches.delete(k); }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (ev) {
  const req = ev.request;
  if (req.method !== "GET") return; /* écritures : toujours réseau */
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; /* API externes : jamais cachées */
  ev.respondWith(
    fetch(req).then(function (resp) {
      if (resp && resp.ok) {
        const copy = resp.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
      }
      return resp;
    }).catch(function () {
      return caches.match(req);
    })
  );
});