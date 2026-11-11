/* Mapa de Fincas · Entrega 102 · trabajador sin conexión (service worker)
   - La app (la página) se guarda en el aparato. Entrega 101: se abre al momento desde la copia guardada y, a la vez,
     se pide la última versión a internet y se guarda para la próxima vez (antes se esperaba a internet hasta 6 s).
     Si aún no hay copia guardada, se pide a internet como antes (sin conexión o si no contesta en 6 s, la copia de la portada).
   - Las librerías de PDF (pdf.json) se guardan también: se sirven desde el aparato y solo se piden cuando cambia la versión.
   - Los trozos de la foto aérea PNOA y de los nombres (IGN) que la app ha descargado para usar sin conexión
     se sirven desde el aparato; lo que no está guardado se pide a internet como siempre.
     Entrega 102: esos trozos están en una base de datos propia (IndexedDB «mf-mapas»), no en el almacén de la app:
     en el iPhone ese almacén se revisa entero al abrir y con miles de trozos la app tardaba más de 10 s en arrancar.
     Mientras la app los pasa del almacén viejo al nuevo, lo que aún no se ha pasado se sirve del viejo.
   - Todo lo demás (sincronización, SIGPAC, Catastro, radar, NDVI…) va directo a internet, sin tocarlo.
   - Entrega 140: recibe los avisos de los recordatorios (los manda la función «avisos» de Supabase), los enseña,
     pone el número del icono y, al tocar el aviso, abre la app en esa nota. */
'use strict';
var APP = 'mf-app-1', MAPAS = 'mf-mapas', PIEZAS = 'mf-piezas', ESPERA = 6000;

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return /^mf-app-/.test(k) && k !== APP; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

var TROZOS = 'mf-mapas', tdbP = null;
function tdb() {
  if (tdbP) return tdbP;
  tdbP = new Promise(function (ok) {
    try {
      var r = indexedDB.open(TROZOS);
      r.onupgradeneeded = function () { var d = r.result; if (!d.objectStoreNames.contains('t')) d.createObjectStore('t'); };
      r.onsuccess = function () { var d = r.result; d.onversionchange = function () { try { d.close(); } catch (x) {} tdbP = null; }; ok(d); };
      r.onerror = function () { tdbP = null; ok(null); };
      r.onblocked = function () { tdbP = null; ok(null); };
    } catch (x) { tdbP = null; ok(null); }
  });
  return tdbP;
}
function trozo(u) {
  return tdb().then(function (d) {
    if (!d) return null;
    return new Promise(function (ok) {
      try {
        var q = d.transaction('t', 'readonly').objectStore('t').get(u);
        q.onsuccess = function () { ok(q.result || null); }; q.onerror = function () { ok(null); };
      } catch (x) { tdbP = null; ok(null); }
    });
  });
}

function esMapa(u) { return u.hostname === 'www.ign.es' && /^\/wmts\/(pnoa-ma|ign-base)\b/.test(u.pathname); }
function esPieza(u) { return u.origin === location.origin && /\/pdf\.json$/.test(u.pathname); }
function clave(u) { return u.origin + u.pathname; }

/* pide la página a internet (sin la copia rápida del navegador) y la guarda; devuelve la respuesta */
function renueva(c, k, req) {
  return fetch(req || k, req ? undefined : { cache: 'no-cache', credentials: 'same-origin' }).then(function (r) {
    if (r && r.ok && r.type === 'basic' && !r.redirected) return c.put(k, r.clone()).then(function () { return r; }, function () { return r; });
    return r;
  });
}

function pagina(e) {
  var req = e.request, u = new URL(req.url), k = clave(u);
  return caches.open(APP).then(function (c) {
    return c.match(k).then(function (m) {
      if (m) {
        /* abre ya la copia guardada; la versión nueva se guarda por detrás para la próxima vez */
        e.waitUntil(renueva(c, k).catch(function () {}));
        return m;
      }
      var red = renueva(c, k, req);
      var plazo = new Promise(function (ok) { setTimeout(function () { ok(null); }, ESPERA); });
      return Promise.race([red.catch(function () { return null; }), plazo]).then(function (r) {
        if (r) return r;
        return c.match(clave(new URL(self.registration.scope))).then(function (m2) { return m2 || red; });
      });
    });
  });
}

/* pdf.json?v=<versión>: lo guardado para esa versión; si no, internet (y se guarda, quitando el de versiones viejas);
   sin conexión, el de otra versión si lo hay */
function pieza(req) {
  var u = new URL(req.url);
  return caches.open(PIEZAS).then(function (c) {
    return c.match(req.url).then(function (m) {
      if (m) return m;
      return fetch(req).then(function (r) {
        if (r && r.ok && r.type === 'basic') {
          var cp = r.clone();
          c.put(req.url, cp).then(function () { return c.keys(); }).then(function (ks) {
            return Promise.all(ks.filter(function (q) { return q.url !== req.url && clave(new URL(q.url)) === clave(u); }).map(function (q) { return c.delete(q); }));
          }).catch(function () {});
        }
        return r;
      }, function (err) {
        return c.match(req.url, { ignoreSearch: true }).then(function (m2) { if (m2) return m2; throw err; });
      });
    });
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var u; try { u = new URL(req.url); } catch (x) { return; }
  if (req.mode === 'navigate' && u.origin === location.origin) { e.respondWith(pagina(e)); return; }
  if (esPieza(u)) { e.respondWith(pieza(req)); return; }
  if (esMapa(u)) {
    /* «Actualizar» en la app pide el trozo de nuevo a internet (cache: reload): entonces no se usa lo guardado */
    if (req.cache === 'reload' || req.cache === 'no-store') return;
    e.respondWith(trozo(req.url).then(function (v) {
      if (v && v.d) return new Response(v.d, { headers: { 'Content-Type': v.t || 'image/jpeg' } });
      return caches.has(MAPAS).then(function (h) { return h ? caches.match(req.url, { cacheName: MAPAS }) : null; })
        .then(function (m) { return m || fetch(req); }, function () { return fetch(req); });
    }));
  }
});

/* la app pide guardar su propia página: nada más abrirse la primera vez (aún no pasó por aquí) y al tocar
   «Hay una versión nueva» (Entrega 101: contesta cuando ya está guardada, para abrir la nueva) */
self.addEventListener('message', function (e) {
  var d = e.data || {}, p = e.ports && e.ports[0];
  var fin = function (ok) { if (p) try { p.postMessage({ t: 'guardada', ok: ok }); } catch (x) {} };
  if (d.t !== 'guarda' || typeof d.u !== 'string') return fin(false);
  var u; try { u = new URL(d.u, self.registration.scope); } catch (x) { return fin(false); }
  if (u.origin !== location.origin) return fin(false);
  e.waitUntil(caches.open(APP).then(function (c) { return renueva(c, clave(u)); })
    .then(function (r) { fin(!!(r && r.ok)); }, function () { fin(false); }));
});

/* Entrega 140 · avisos de los recordatorios de las notas (Web Push). Llegan cifrados desde la función «avisos» de Supabase:
   {t: título, b: texto, id: nota, n: número para el icono, u: dirección que abre}. Siempre se enseña el aviso (el iPhone lo exige). */
self.addEventListener('push', function (e) {
  var d = {};
  try { d = e.data ? e.data.json() : {}; } catch (x) { try { d = { b: e.data ? e.data.text() : '' }; } catch (y) { d = {}; } }
  var p = self.registration.showNotification(String(d.t || '⏰ Recordatorio'), {
    body: String(d.b || ''), tag: d.id ? 'nota-' + d.id : 'mf-aviso', data: d, icon: 'apple-touch-icon.png', badge: 'apple-touch-icon.png'
  });
  if (typeof d.n === 'number' && self.navigator && self.navigator.setAppBadge) {
    p = p.then(function () { return d.n > 0 ? self.navigator.setAppBadge(d.n) : (self.navigator.clearAppBadge ? self.navigator.clearAppBadge() : null); }).catch(function () {});
  }
  e.waitUntil(p);
});
self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  var d = e.notification.data || {}, sc = self.registration.scope, u = sc;
  try { var x = new URL(d.u || sc, sc); if (x.origin === location.origin) u = x.href; } catch (y) {}
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.url.indexOf(sc) === 0 && 'focus' in c) { try { c.postMessage({ mf: 'nota', id: d.id || '' }); } catch (z) {} return c.focus(); }
    }
    return self.clients.openWindow ? self.clients.openWindow(u) : null;
  }));
});
