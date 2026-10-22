/* Mapa de Fincas · Entrega 98 · trabajador sin conexión (service worker)
   - La app (la página) se guarda en el aparato: con conexión siempre se pide la última versión y se guarda;
     sin conexión (o si la red no contesta en 6 s) se abre la copia guardada.
   - Los trozos de la foto aérea PNOA y de los nombres (IGN) que la app ha descargado para usar sin conexión
     se sirven desde el aparato; lo que no está guardado se pide a internet como siempre.
   - Todo lo demás (sincronización, SIGPAC, Catastro, radar, NDVI…) va directo a internet, sin tocarlo. */
'use strict';
var APP = 'mf-app-1', MAPAS = 'mf-mapas', ESPERA = 6000;

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return /^mf-app-/.test(k) && k !== APP; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

function esMapa(u) { return u.hostname === 'www.ign.es' && /^\/wmts\/(pnoa-ma|ign-base)\b/.test(u.pathname); }
function clave(u) { return u.origin + u.pathname; }

function pagina(req) {
  var u = new URL(req.url), k = clave(u);
  return caches.open(APP).then(function (c) {
    var red = fetch(req).then(function (r) {
      if (r && r.ok && r.type === 'basic') c.put(k, r.clone()).catch(function () {});
      return r;
    });
    var plazo = new Promise(function (ok) { setTimeout(function () { ok(null); }, ESPERA); });
    return Promise.race([red.catch(function () { return null; }), plazo]).then(function (r) {
      if (r) return r;
      return c.match(k).then(function (m) {
        if (m) return m;
        return c.match(clave(new URL(self.registration.scope))).then(function (m2) { return m2 || red; });
      });
    });
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var u; try { u = new URL(req.url); } catch (x) { return; }
  if (req.mode === 'navigate' && u.origin === location.origin) { e.respondWith(pagina(req)); return; }
  if (esMapa(u)) {
    /* «Actualizar» en la app pide el trozo de nuevo a internet (cache: reload): entonces no se usa lo guardado */
    if (req.cache === 'reload' || req.cache === 'no-store') return;
    e.respondWith(caches.open(MAPAS).then(function (c) {
      return c.match(req.url).then(function (m) { return m || fetch(req); });
    }));
  }
});

/* la app pide guardar su propia página nada más abrirse (la primera vez aún no pasó por aquí) */
self.addEventListener('message', function (e) {
  var d = e.data || {};
  if (d.t !== 'guarda' || typeof d.u !== 'string') return;
  var u; try { u = new URL(d.u, self.registration.scope); } catch (x) { return; }
  if (u.origin !== location.origin) return;
  e.waitUntil(fetch(u.href, { cache: 'no-cache', credentials: 'same-origin' }).then(function (r) {
    if (r && r.ok) return caches.open(APP).then(function (c) { return c.put(clave(u), r); });
  }).catch(function () {}));
});
