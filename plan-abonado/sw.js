/* Funciona sin cobertura: guarda la herramienta en el aparato. Los datos del usuario no pasan por aquí. */
var V='202611171759', C='plan-abonado-'+V;
var F=['./','index.html','app.js?v='+V,'style.css?v='+V,'manifest.webmanifest','iconos/favicon.png','iconos/apple-touch-icon.png','iconos/icono-192.png','iconos/icono-512.png',
  'fonts/barlow-latin-400-normal.woff2','fonts/barlow-latin-500-normal.woff2','fonts/barlow-latin-600-normal.woff2','fonts/barlow-latin-700-normal.woff2',
  'lib/pdf.js','lib/pdfjs.js','lib/pdfjs-worker.js','privacidad.html','condiciones.html'];
self.addEventListener('install',function(e){e.waitUntil(caches.open(C).then(function(c){return c.addAll(F)}).then(function(){return self.skipWaiting()}))});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==C&&/^plan-abonado-/.test(k)}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener('fetch',function(e){var u=new URL(e.request.url);if(e.request.method!=='GET'||u.origin!==location.origin)return;
  var nav=e.request.mode==='navigate';
  if(nav){e.respondWith(fetch(e.request).then(function(r){var cp=r.clone();caches.open(C).then(function(c){c.put('index.html',cp)});return r}).catch(function(){return caches.match('index.html')}));return}
  e.respondWith(caches.match(e.request).then(function(r){return r||fetch(e.request).then(function(n){if(n.ok){var cp=n.clone();caches.open(C).then(function(c){c.put(e.request,cp)})}return n})}));
});
