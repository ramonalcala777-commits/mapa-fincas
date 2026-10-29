const fs=require('fs');
const t=fs.readFileSync('template.html','utf8');
let css=fs.readFileSync('package/dist/leaflet.css','utf8').replace(/url\(\s*['"]?images\/[^)]*\)/g,'none');
const js=fs.readFileSync('package/dist/leaflet.js','utf8');
if(/<\/script/i.test(js)) throw new Error('leaflet.js contiene </script');
const icon=require('./icon.js');
const pdf=fs.readFileSync('vendor/node_modules/jspdf/dist/jspdf.umd.min.js','utf8').replace(/<\/script/gi,'<\\/script');
const pjs=fs.readFileSync('vendor/node_modules/pdfjs-dist/build/pdf.min.js','utf8').replace(/<\/script/gi,'<\\/script'), pjw=fs.readFileSync('vendor/node_modules/pdfjs-dist/build/pdf.worker.min.js','utf8').replace(/<\/script/gi,'<\\/script');
const ENTREGA='Entrega 113';
const FR=fs.readFileSync('fonts/barlow-pdf-400.ttf').toString('base64'), FB=fs.readFileSync('fonts/barlow-pdf-600.ttf').toString('base64'), FCR=fs.readFileSync('fonts/barlowc-pdf-500.ttf').toString('base64'), FCB=fs.readFileSync('fonts/barlowc-pdf-600.ttf').toString('base64');
const BW=[['Barlow',400,'barlow-latin-400'],['Barlow',500,'barlow-latin-500'],['Barlow',600,'barlow-latin-600'],['Barlow',700,'barlow-latin-700'],['Barlow Semi Condensed',500,'barlow-semi-condensed-latin-500'],['Barlow Semi Condensed',600,'barlow-semi-condensed-latin-600']];
const BCSS=BW.map(([f,w,n])=>"@font-face{font-family:'"+f+"';font-style:normal;font-weight:"+w+";font-display:swap;src:url(data:font/woff2;base64,"+fs.readFileSync('fonts/'+n+'-normal.woff2').toString('base64')+") format('woff2')}").join('\n');
const now=process.env.MF_HORA?new Date(process.env.MF_HORA):new Date(), pr=Object.fromEntries(new Intl.DateTimeFormat('es-ES',{timeZone:'Europe/Madrid',day:'numeric',month:'short',hour:'2-digit',minute:'2-digit',hour12:false}).formatToParts(now).filter(p=>p.type!=='literal').map(p=>[p.type,p.value]));
const BV=now.toISOString().slice(0,19)+'Z', BN=ENTREGA+' · '+pr.day+' '+pr.month.replace('.','')+' · '+pr.hour+':'+pr.minute;
fs.writeFileSync('version.json',JSON.stringify({v:BV,n:BN}));
const INI=require('./inicio.js')();
if(!t.includes('<!--INICIO-->')) throw new Error('falta <!--INICIO--> en template.html');
const out=t.replace('<!--INICIO-->',()=>INI.html).replace('/*BUILDV*/',BV).replace('/*BUILDN*/',BN).replace('/*PDFLIB*/',()=>pdf).replace('/*PDFJS*/',()=>pjs).replace('/*PDFJSW*/',()=>pjw).replace('/*BARLOWCSS*/',()=>BCSS).replace('/*FONTR*/',()=>FR).replace('/*FONTB*/',()=>FB).replace('/*FONTCR*/',()=>FCR).replace('/*FONTCB*/',()=>FCB).replace(/\/\*ICON\*\//g,()=>icon).replace('/*LEAFLET_CSS*/',()=>css).replace('/*LEAFLET_JS*/',()=>js);
// Entrega 60: cabecera de seguridad (CSP). Huella de cada programa interno, calculada en cada montaje.
const crypto=require('crypto');
const hashes=[...out.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m=>"'sha256-"+crypto.createHash('sha256').update(m[1],'utf8').digest('base64')+"'");
const MAPAS='https://server.arcgisonline.com https://services.arcgisonline.com https://*.arcgis.com https://www.ign.es https://www.cartociudad.es https://sigpac.mapa.gob.es https://sigpac-hubcloud.es https://ovc.catastro.meh.es https://www1.sedecatastro.gob.es https://tile.openstreetmap.org https://tilecache.rainviewer.com https://rbrvokfrbvcsqmskshul.supabase.co';
const CSP=["default-src 'self'","script-src 'self' 'unsafe-eval' "+hashes.join(' '),"style-src 'self' 'unsafe-inline'","img-src 'self' data: blob: "+MAPAS,"font-src 'self' data:","connect-src 'self' blob: data: https://rbrvokfrbvcsqmskshul.supabase.co https://nominatim.openstreetmap.org https://api.rainviewer.com https://s3-us-west-2.amazonaws.com https://archive-api.open-meteo.com https://api.open-meteo.com https://api.met.no https://servicios.idee.es "+MAPAS,"worker-src 'self' blob:","frame-src 'self'","object-src 'none'","media-src 'none'","base-uri 'self'","form-action 'none'"].join('; ');
const out2=out.replace('<!--CSP-->','<meta http-equiv="Content-Security-Policy" content="'+CSP+'">');
if(out2===out) throw new Error('no se pudo insertar la CSP');
fs.writeFileSync('mapa-fincas.html',out2);
// Entrega 101: arranque más rápido. index.html (lo que se publica) va sin las librerías de PDF ni sus letras (casi 2 MB):
// la app las trae de pdf.json unos segundos después de abrir. mapa-fincas.html las lleva dentro (lo usan las pruebas).
// También las pantallas de inicio del iPhone (inicio/*.png). piezas.json: lo que se publica además de index.html y version.json.
const PZ={pdflib:pdf,pdfjslib:pjs,pdffontr:FR,pdffontb:FB,pdffontcr:FCR,pdffontcb:FCB,pdfjsworker:pjw};
let lite=out2;
for(const id of Object.keys(PZ)){
  const a='<script type="text/plain" id="'+id+'">', i=lite.indexOf(a), j=lite.indexOf('</script>',i);
  if(i<0||j<0||lite.indexOf(a,i+1)>=0) throw new Error('no encuentro la pieza '+id);
  if(lite.slice(i+a.length,j)!==PZ[id]) throw new Error('la pieza '+id+' no es la esperada');
  lite=lite.slice(0,i+a.length)+lite.slice(j);
}
fs.writeFileSync('index.html',lite);
fs.writeFileSync('pdf.json',JSON.stringify(PZ));
// Entrega 106: geotiff.js (lee el MDT del IGN para la aptitud climática) se publica aparte y solo se carga al calcular
fs.writeFileSync('geotiff-2.1.3.js',fs.readFileSync('vendor/geotiff/geotiff.js'));
fs.mkdirSync('inicio',{recursive:true});
for(const [n,b] of Object.entries(INI.archivos)) fs.writeFileSync(n,b);
fs.writeFileSync('piezas.json',JSON.stringify(['pdf.json','geotiff-2.1.3.js'].concat(Object.keys(INI.archivos).sort())));
console.log('index.html KB:',(lite.length/1024).toFixed(0),'· pdf.json KB:',(fs.statSync('pdf.json').size/1024).toFixed(0),'· pantallas de inicio:',Object.keys(INI.archivos).length);
// script de la app (sin Leaflet) para pruebas de sintaxis
const m=[...t.matchAll(/<script>([\s\S]*?)<\/script>/g)].sort((a,b)=>b[1].length-a[1].length); m.unshift(null);
fs.writeFileSync('app.js',m[1][1]);
console.log('tamaño final KB:',(out2.length/1024).toFixed(0),'· app.js KB:',(m[1][1].length/1024).toFixed(0));
