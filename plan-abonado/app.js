/* Piezas comunes (formato de números, Excel, ZIP, tablas en el móvil) */
var esc=function(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})};
var nf=function(n,d){var t=Number(n).toLocaleString('es-ES',{minimumFractionDigits:d,maximumFractionDigits:d}); return t.replace(/^(-?)(\d)(\d{3})(?=,|$)/,'$1$2.$3')};
var ha2=function(n){return nf(n,2)}, ha4=function(n){return nf(n,4)};
var uid=function(){var s='';try{var a=new Uint8Array(8);(window.crypto||window.msCrypto).getRandomValues(a);for(var i=0;i<8;i++)s+=(a[i]%36).toString(36)}catch(e){s=(Math.random().toString(36)+'00000000').slice(2,10)}return s+Date.now().toString(36).slice(-4)};
var NOCROP='Sin clasificar';
var noAcc=function(s){return String(s==null?'':s).normalize('NFD').replace(/[̀-ͯ]/g,'')};
var safeName=function(s){return noAcc(s).replace(/[^\w-]+/g,'_').replace(/^_+|_+$/g,'')||'archivo'};
var u8s=function(s){return new TextEncoder().encode(s)};
var zipStore=function(files){return new Blob(zipBytes(files),{type:'application/zip'})};
var CRCT=null;
var diaMes=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
var NUMDOT=0;
var XC={ink:'172126',gray:'5A6765',hd:'E9EFEC',wk:'F6F8F7',on:'E3F1E8',pe:'ECEFEE',yel:'F5C400',yel2:'FFF7D6',alt:'F3F5F4',box:'F7F9F8',bd:'B0BAB7',grn:'2E7D4F',red:'B04028'};
var TABVAR={Pistacho:/^(kerman|larnaka|sirora|golden hills|lost hills|aegina|avdat|kastel|mateur|napoletana|bianca|sfax|joley|peters|red aleppo|c-special|c special|batoury|kalehghouchi|ohadi)/,Almendro:/^(marcona|guara|lauranne|vairo|soleta|penta|tuono|ferragnes|ferraduel|belona|constanti|avijor|antonieta|marinada|makako|masbovera|desmayo|ramillete|nonpareil|monterey|carmel|ruby|tarraco|glorieta|mardia|isabelona|texas)/,Olivo:/^(arbequina|picual|cornicabra|hojiblanca|manzanilla|koroneiki|arbosana|empeltre|frantoio|picudo|lechin|verdial|gordal)/,'Viña':/^(tempranillo|airen|bobal|garnacha|monastrell|cabernet|syrah|merlot|macabeo|verdejo|chardonnay|sauvignon|petit verdot|tintorera|moscatel|pinot|viura)/};
function refLabel(r){var a=r.split(':');return String(a[0]).padStart(2,'0')+':'+String(a[1]).padStart(3,'0')+':'+a[2]+':'+a[3]+':'+a[4]+':'+a[5]+':'+a[6]}
function haC(r){return r?(+r.hc>0?+r.hc:(r.ha||0)):0}
function haPair(rs){var s=0,c=0,d=false;(rs||[]).forEach(function(r){s+=r.ha||0;c+=haC(r);if(+r.hc>0)d=true});return {s:s,c:c,d:d&&Math.abs(s-c)>0.00005}}
function haTxt(rs){var x=haPair(rs);return x.d?ha2(x.c)+' ha cultivadas (SIGPAC '+ha2(x.s)+' ha)':ha2(x.s)+' ha'}
function campLbl(k){k=String(k==null?'':k);return /^\d{4}$/.test(k)?(+k-1)+'/'+k.slice(2):k}
function campFile(k){return campLbl(k).replace('/','-')}
function abHas(v){return v!==null&&v!==undefined&&v!==''}
function abEd(v){return abHas(v)?String(v).replace('.',','):''}
function numEs(v){
  var t=String(v==null?'':v).trim().replace(/\s/g,''); if(!t)return null;
  if(/^[1-9]\d{0,2}(\.\d{3})+(,\d+)?$/.test(t))t=t.replace(/\./g,'').replace(',','.');
  else if(/^\d+(,\d+)?$/.test(t))t=t.replace(',','.');
  else{if(/^[\d.,]+$/.test(t)&&t.indexOf('.')>=0)NUMDOT=Date.now(); return NaN}
  var n=Number(t); return isFinite(n)?n:NaN;
}
function trNum(v){return numEs(v)}
function trTxt(n,d,trim){if(n===null||n===undefined||n===''||isNaN(n))return '';var t=nf(n,d==null?2:d);if(trim&&t.indexOf(',')>=0)t=t.replace(/0+$/,'').replace(/,$/,'');return t}
function diaFmt(iso){var m=/^(\d{4})-(\d\d)-(\d\d)$/.exec(iso||'');return m?(+m[3])+' '+diaMes[+m[2]-1]+' '+m[1]:'Sin fecha'}
function daLbl(L,k){var x=L.filter(function(y){return y[0]===k})[0];return x?x[1]:''}
function daRepTxt(t){t=t||{};return t.rn?t.rn+(t.rnif?' (NIF '+t.rnif+')':''):''}
function xlSheet(name,o){o=o||{};return {name:String(name).replace(/[\\\/\?\*\[\]:]/g,' ').slice(0,31),cells:{},merges:[],widths:o.widths||[],heights:{},charts:[],foot:o.foot||'',land:o.land!==false,maxR:0,maxC:0}}
function xlSet(sh,r,c,v,st){sh.cells[r+','+c]={v:v,st:st||null};if(r>sh.maxR)sh.maxR=r;if(c>sh.maxC)sh.maxC=c}
function xlMerge(sh,r1,c1,r2,c2,v,st){for(var r=r1;r<=r2;r++)for(var c=c1;c<=c2;c++)xlSet(sh,r,c,r===r1&&c===c1?v:null,st);if(r2>r1||c2>c1)sh.merges.push([r1,c1,r2,c2])}
function xlH(sh,r,h){sh.heights[r]=Math.max(sh.heights[r]||0,h)}
function xlCol(i){var s='';i++;while(i>0){var m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
function xlFmt(v,maxd,suf){var d=0;if(typeof v==='number'&&isFinite(v)){for(d=0;d<(maxd==null?2:maxd);d++){if(Math.abs(v*Math.pow(10,d)-Math.round(v*Math.pow(10,d)))<1e-7)break}}return '#,##0'+(d?'.'+new Array(d+1).join('0'):'')+(suf?'"'+suf+'"':'')}
function xlBuild(sheets){
  var X=function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]}).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'')};
  var fonts=['<font><sz val="10"/><color rgb="FF172126"/><name val="Calibri"/><family val="2"/></font>'], fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'], borders=['<border><left/><right/><top/><bottom/><diagonal/></border>'], fmts=[], xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], K={};
  var idx=function(arr,s){var i=arr.indexOf(s);if(i<0){arr.push(s);i=arr.length-1}return i};
  var fontX=function(st){return '<font>'+(st.b?'<b/>':'')+(st.i?'<i/>':'')+'<sz val="'+(st.sz||10)+'"/><color rgb="FF'+(st.col||'172126')+'"/><name val="Calibri"/><family val="2"/></font>'};
  var side=function(n,v,cl){if(!v)return '<'+n+'/>';return '<'+n+' style="'+v+'"><color rgb="FF'+cl+'"/></'+n+'>'};
  var styleId=function(st){
    if(!st)return 0; var key=JSON.stringify(st); if(K[key]!==undefined)return K[key];
    var fo=idx(fonts,fontX(st)), fi=st.fill?idx(fills,'<fill><patternFill patternType="solid"><fgColor rgb="FF'+st.fill+'"/><bgColor indexed="64"/></patternFill></fill>'):0;
    var bd=0; if(st.bd){var B=typeof st.bd==='string'?{l:st.bd,r:st.bd,t:st.bd,b:st.bd}:st.bd, bc=st.bc||'B0BAB7';
      bd=idx(borders,'<border>'+side('left',B.l,B.lc||bc)+side('right',B.r,B.rc||bc)+side('top',B.t,B.tc||bc)+side('bottom',B.b,B.bc||bc)+'<diagonal/></border>')}
    var nf=0; if(st.fmt){var fi2=fmts.indexOf(st.fmt);if(fi2<0){fmts.push(st.fmt);fi2=fmts.length-1}nf=164+fi2}
    var al=(st.ha||st.va||st.wrap||st.ind)?'<alignment'+(st.ha?' horizontal="'+st.ha+'"':(st.ind?' horizontal="left"':''))+' vertical="'+(st.va||'center')+'"'+(st.wrap?' wrapText="1"':'')+(st.ind&&st.ha!=='center'?' indent="'+st.ind+'"':'')+'/>':'<alignment vertical="center"/>';
    xfs.push('<xf numFmtId="'+nf+'" fontId="'+fo+'" fillId="'+fi+'" borderId="'+bd+'" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"'+(nf?' applyNumberFormat="1"':'')+'>'+al+'</xf>');
    return (K[key]=xfs.length-1);
  };
  var runs=function(rs){return rs.map(function(p){return '<r><rPr>'+(p.b?'<b/>':'')+(p.i?'<i/>':'')+'<sz val="'+(p.sz||10)+'"/><color rgb="FF'+(p.col||'172126')+'"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve">'+X(p.t)+'</t></r>'}).join('')};
  var files=[], ct=[], wbRels=[], nDr=0, nCh=0;
  sheets.forEach(function(sh,si){
    var rowsX='';
    for(var r=0;r<=sh.maxR;r++){
      var cs='', any=false;
      for(var c=0;c<=sh.maxC;c++){var o=sh.cells[r+','+c]; if(!o)continue; var s=styleId(o.st), ref=xlCol(c)+(r+1), v=o.v; any=true;
        if(v===null||v===undefined||v===''){if(s)cs+='<c r="'+ref+'" s="'+s+'"/>';continue}
        if(typeof v==='number'){cs+='<c r="'+ref+'" s="'+s+'"><v>'+(isFinite(v)?+v.toFixed(10):0)+'</v></c>';continue}
        if(v&&v.rich){cs+='<c r="'+ref+'" s="'+s+'" t="inlineStr"><is>'+runs(v.rich)+'</is></c>';continue}
        cs+='<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t xml:space="preserve">'+X(v)+'</t></is></c>'}
      if(any||sh.heights[r])rowsX+='<row r="'+(r+1)+'"'+(sh.heights[r]?' ht="'+sh.heights[r].toFixed(1)+'" customHeight="1"':'')+'>'+cs+'</row>';
    }
    var cols=sh.widths.length?'<cols>'+sh.widths.map(function(w,i){return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+(+w).toFixed(2)+'" customWidth="1"/>'}).join('')+'</cols>':'';
    var mg=sh.merges.length?'<mergeCells count="'+sh.merges.length+'">'+sh.merges.map(function(m){return '<mergeCell ref="'+xlCol(m[1])+(m[0]+1)+':'+xlCol(m[3])+(m[2]+1)+'"/>'}).join('')+'</mergeCells>':'';
    var foot=sh.foot?'<headerFooter><oddFooter>'+X('&L&8'+sh.foot.replace(/&/g,'&&')+'&R&8&P / &N')+'</oddFooter></headerFooter>':'';
    var dr=''; if(sh.charts.length){nDr++;dr='<drawing r:id="rId1"/>';
      files.push({n:'xl/worksheets/_rels/sheet'+(si+1)+'.xml.rels',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'+nDr+'.xml"/></Relationships>')});
      var anc='', dRel='';
      sh.charts.forEach(function(chs,k){nCh++;
        anc+='<xdr:oneCellAnchor><xdr:from><xdr:col>'+chs.c+'</xdr:col><xdr:colOff>'+(chs.dx||0)+'</xdr:colOff><xdr:row>'+chs.r+'</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="'+Math.round(chs.w*360000)+'" cy="'+Math.round(chs.h*360000)+'"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="'+(k+2)+'" name="Gráfica '+(k+1)+'"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId'+(k+1)+'"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>';
        dRel+='<Relationship Id="rId'+(k+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart'+nCh+'.xml"/>';
        files.push({n:'xl/charts/chart'+nCh+'.xml',d:u8s(xlChartXML(chs,X))}); ct.push('<Override PartName="/xl/charts/chart'+nCh+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
      });
      files.push({n:'xl/drawings/drawing'+nDr+'.xml',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+anc+'</xdr:wsDr>')});
      files.push({n:'xl/drawings/_rels/drawing'+nDr+'.xml.rels',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+dRel+'</Relationships>')});
      ct.push('<Override PartName="/xl/drawings/drawing'+nDr+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
    }
    files.push({n:'xl/worksheets/sheet'+(si+1)+'.xml',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0"'+(si===0?' tabSelected="1"':'')+'/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'+cols+'<sheetData>'+rowsX+'</sheetData>'+mg+'<pageMargins left="0.4" right="0.4" top="0.45" bottom="0.55" header="0.3" footer="0.25"/><pageSetup paperSize="9" orientation="'+(sh.land?'landscape':'portrait')+'" fitToWidth="1" fitToHeight="0"/>'+foot+dr+'</worksheet>')});
    ct.push('<Override PartName="/xl/worksheets/sheet'+(si+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    wbRels.push('<Relationship Id="rId'+(si+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(si+1)+'.xml"/>');
  });
  var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+(fmts.length?'<numFmts count="'+fmts.length+'">'+fmts.map(function(f,i){return '<numFmt numFmtId="'+(164+i)+'" formatCode="'+X(f)+'"/>'}).join('')+'</numFmts>':'')+'<fonts count="'+fonts.length+'">'+fonts.join('')+'</fonts><fills count="'+fills.length+'">'+fills.join('')+'</fills><borders count="'+borders.length+'">'+borders.join('')+'</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="'+xfs.length+'">'+xfs.join('')+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  files.push({n:'xl/styles.xml',d:u8s(styles)});
  files.push({n:'xl/workbook.xml',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>'+sheets.map(function(s,i){return '<sheet name="'+X(s.name)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>'}).join('')+'</sheets></workbook>')});
  files.push({n:'xl/_rels/workbook.xml.rels',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+wbRels.join('')+'<Relationship Id="rId'+(sheets.length+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')});
  files.push({n:'_rels/.rels',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')});
  files.unshift({n:'[Content_Types].xml',d:u8s('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+ct.join('')+'</Types>')});
  return zipStore(files).slice(0,undefined,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
function xlSpans(W,mm,c0,c1){c0=c0||0;c1=c1===undefined?W.length-1:c1;var out=[],start=c0;
  mm.forEach(function(w,k){var left=mm.slice(k).reduce(function(a,b){return a+b},0),room=0,i;for(i=start;i<=c1;i++)room+=W[i];var tg=w/left*room,acc=0,end=start,bd=1e9;
    for(i=start;i<=c1-(mm.length-1-k);i++){acc+=W[i];var d=Math.abs(acc-tg);if(d<bd){bd=d;end=i}}if(k===mm.length-1)end=c1;out.push([start,end]);start=end+1});return out}
function xlLines(t,chars){return String(t||'').split('\n').reduce(function(a,p){return a+Math.max(1,Math.ceil(p.length/Math.max(8,chars)))},0)}
function xlDoc(name,foot){
  var N=26, MM=182/N, CW=[], i; for(i=0;i<N;i++)CW.push(MM/1.9);
  var sh=xlSheet(name,{widths:CW,foot:foot,land:false}), L=N-1, D={sh:sh,r:0,L:L};
  var chars=function(s){var w=0;for(var c=s[0];c<=s[1];c++)w+=CW[c];return w*1.1};
  D.band=function(top,title,right,sub){
    xlMerge(sh,0,0,0,L-8,top,{sz:9,col:'FFFFFF',fill:XC.ink,ind:1,va:'bottom'});xlMerge(sh,0,L-7,0,L,right,{sz:9,col:'FFFFFF',fill:XC.ink,ha:'right',va:'bottom'});xlH(sh,0,18);
    xlMerge(sh,1,0,1,L,title,{b:true,sz:18,col:'FFFFFF',fill:XC.ink,ind:1});xlH(sh,1,30);
    var r=2; if(sub){xlMerge(sh,r,0,r,L,sub,{sz:10,col:'FFFFFF',fill:XC.ink,ind:1,va:'top'});xlH(sh,r,18);r++}
    xlMerge(sh,r,0,r,L,null,{fill:XC.yel});xlH(sh,r,4);D.r=r+2;
  };
  D.title=function(t){D.r++;xlMerge(sh,D.r,0,D.r,L,t,{b:true,sz:12,bd:{b:'medium'},bc:XC.ink,va:'bottom'});xlH(sh,D.r,21);D.r+=2};
  D.sub=function(t){xlMerge(sh,D.r,0,D.r,L,t,{b:true,sz:11});xlH(sh,D.r,18);D.r++};
  D.para=function(t,o){o=o||{};var s=[0,L];xlMerge(sh,D.r,0,D.r,L,t,{sz:o.sz||10,b:!!o.b,col:o.col,wrap:true,va:'top'});xlH(sh,D.r,Math.max(15,xlLines(t,chars(s))*(o.sz||10)*1.45+3));D.r++};
  D.kv=function(k,v){var nk=7,a=xlAuto(v),st={sz:10,wrap:true,va:'top'};if(a&&typeof a==='object'){st.fmt=a.f;st.ha='left';v=a.v}xlMerge(sh,D.r,0,D.r,nk-1,k,{b:true,sz:10,va:'top'});xlMerge(sh,D.r,nk,D.r,L,v,st);xlH(sh,D.r,Math.max(15,xlLines(v,chars([nk,L]))*14.5+2));D.r++};
  D.gap=function(n){D.r+=n||1};
  /* cols: [{h,w (mm),a:'r'}]; celdas: texto, número ({v,f}: formato) o {t,s} (dos líneas, la segunda en gris); o.tot: fila de total; o.bold(i): fila destacada */
  D.tbl=function(cols,rows,o){o=o||{};var sp=xlSpans(CW,cols.map(function(c){return c.w}));
    sp.forEach(function(s,k){xlMerge(sh,D.r,s[0],D.r,s[1],cols[k].h,{b:true,sz:9,col:XC.ink,ha:cols[k].a==='r'?'right':'left',wrap:true,va:'bottom',bd:{b:'thin'},bc:XC.ink,ind:cols[k].a==='r'?0:1})});
    xlH(sh,D.r,Math.max(17,Math.max.apply(null,sp.map(function(s,k){return xlLines(cols[k].h,chars(s))}))*12+5));D.r++;
    var z=0;
    rows.forEach(function(rw,i){var bold=o.bold&&o.bold(rw,i), fill=bold?(o.boldFill||"FFF4C2"):(z++%2===0?XC.alt:null); if(bold)z=0; var hl=1;
      sp.forEach(function(s,k){var v=rw[k], c=cols[k], st={sz:9.5,b:!!bold,fill:fill,ha:c.a==='r'?'right':'left',wrap:true,va:'top',ind:c.a==='r'?0:1};
        if(v&&typeof v==='object'&&v.t!==undefined){var lines=xlLines(v.t,chars(s))+(v.s?xlLines(v.s,chars(s)*1.15):0);hl=Math.max(hl,lines);xlMerge(sh,D.r,s[0],D.r,s[1],v.s?{rich:[{t:String(v.t),b:!!bold,sz:9.5},{t:'\n'+v.s,sz:8,col:XC.gray}]}:String(v.t),st);return}
        if(v&&typeof v==='object'&&v.v!==undefined){if(v.v===null||v.v===''){xlMerge(sh,D.r,s[0],D.r,s[1],'',st);return}st.fmt=v.f||xlFmt(v.v,2);xlMerge(sh,D.r,s[0],D.r,s[1],v.v,st);return}
        hl=Math.max(hl,xlLines(v==null?'':v,chars(s)));xlMerge(sh,D.r,s[0],D.r,s[1],v==null?'':v,st)});
      xlH(sh,D.r,Math.max(16,hl*12.5+4));D.r++});
    if(o.tot){sp.forEach(function(s,k){var v=o.tot[k],c=cols[k],st={b:true,sz:9.5,ha:c.a==='r'?'right':'left',bd:{t:'thin'},bc:XC.ink,ind:c.a==='r'?0:1};if(v&&typeof v==='object'){st.fmt=v.f||xlFmt(v.v,2);v=v.v}xlMerge(sh,D.r,s[0],D.r,s[1],v==null?'':v,st)});xlH(sh,D.r,17);D.r++}
    D.r++;
  };
  return D;
}
function xlN(v,d,suf){if(v===null||v===undefined||v===''||isNaN(+v))return {v:null};var x=Math.round(+v*Math.pow(10,d||0))/Math.pow(10,d||0);return {v:x,f:'#,##0'+(d?'.'+new Array(d+1).join('0'):'')+(suf?'"'+suf+'"':'')}}
function xlAuto(v){if(typeof v!=='string')return v;var m=/^\s*(-?\d{1,3}(?:\.\d{3})+|-?\d+)(?:,(\d+))?\s*(kg\/ha|L\/ha|t\/ha|kg\/t|kg|L|%|ha|m³\/ha|m³|€\/ha|€)?\s*$/.exec(v);if(!m||(!m[3]&&!m[2]&&m[1].indexOf('.')<0))return v;var n=+(m[1].replace(/\./g,'')+(m[2]?'.'+m[2]:'')),d=m[2]?m[2].length:0;if(!isFinite(n))return v;return {v:n,f:'#,##0'+(d?'.'+new Array(d+1).join('0'):'')+(m[3]?'" '+m[3]+'"':'')}}
function crc32(u8){
  if(!CRCT){CRCT=new Uint32Array(256);for(var n=0;n<256;n++){var c=n;for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);CRCT[n]=c>>>0}}
  var c2=0xFFFFFFFF; for(var i=0;i<u8.length;i++)c2=CRCT[(c2^u8[i])&255]^(c2>>>8); return (c2^0xFFFFFFFF)>>>0;
}
function zipBytes(files){
  var parts=[], cen=[], off=0, dt=new Date();
  var dosT=(dt.getHours()<<11)|(dt.getMinutes()<<5)|(dt.getSeconds()>>1), dosD=((dt.getFullYear()-1980)<<9)|((dt.getMonth()+1)<<5)|dt.getDate();
  files.forEach(function(f){
    var nm=u8s(f.n), crc=crc32(f.d), L=new DataView(new ArrayBuffer(30)), C=new DataView(new ArrayBuffer(46));
    L.setUint32(0,0x04034b50,true);L.setUint16(4,20,true);L.setUint16(6,0x0800,true);L.setUint16(8,0,true);L.setUint16(10,dosT,true);L.setUint16(12,dosD,true);
    L.setUint32(14,crc,true);L.setUint32(18,f.d.length,true);L.setUint32(22,f.d.length,true);L.setUint16(26,nm.length,true);L.setUint16(28,0,true);
    C.setUint32(0,0x02014b50,true);C.setUint16(4,20,true);C.setUint16(6,20,true);C.setUint16(8,0x0800,true);C.setUint16(10,0,true);C.setUint16(12,dosT,true);C.setUint16(14,dosD,true);
    C.setUint32(16,crc,true);C.setUint32(20,f.d.length,true);C.setUint32(24,f.d.length,true);C.setUint16(28,nm.length,true);C.setUint32(42,off,true);
    parts.push(new Uint8Array(L.buffer),nm,f.d); cen.push(new Uint8Array(C.buffer),nm);
    off+=30+nm.length+f.d.length;
  });
  var cs=0; cen.forEach(function(a){cs+=a.length});
  var E=new DataView(new ArrayBuffer(22)); E.setUint32(0,0x06054b50,true);E.setUint16(8,files.length,true);E.setUint16(10,files.length,true);E.setUint32(12,cs,true);E.setUint32(16,off,true);
  return parts.concat(cen,[new Uint8Array(E.buffer)]);
}
function canShareFile(blob,name){try{return typeof File==='function'&&navigator.canShare&&navigator.canShare({files:[new File([blob],name,{type:blob.type||'application/octet-stream'})]})}catch(e){return false}}
function stkLab(){
  var L=document.querySelectorAll('table.stk:not([data-stk])');
  for(var i=0;i<L.length;i++){var tb=L[i], hr=tb.tHead&&tb.tHead.rows.length?tb.tHead.rows[tb.tHead.rows.length-1]:null; tb.setAttribute('data-stk','1'); if(!hr)continue;
    var hs=[]; for(var j=0;j<hr.cells.length;j++){var c=hr.cells[j], tx=(c.textContent||'').replace(/\s+/g,' ').trim(); for(var s=0;s<(c.colSpan||1);s++)hs.push(tx)}
    var rs=tb.querySelectorAll('tbody tr, tfoot tr');
    for(var k=0;k<rs.length;k++){var pos=0; for(var m=0;m<rs[k].cells.length;m++){var td=rs[k].cells[m]; if(!td.hasAttribute('data-l'))td.setAttribute('data-l',hs[pos]||''); if(m===0&&td.querySelector('input,select,textarea'))td.classList.add('stkin'); pos+=td.colSpan||1}}
  }
}
function inThou(){[].forEach.call(document.querySelectorAll('textarea.autoh:not([data-th])'),function(e){e.setAttribute('data-th','1');taAuto(e)});[].forEach.call(document.querySelectorAll('input[inputmode="decimal"]:not([data-th])'),function(i){i.setAttribute('data-th','1');if(i===document.activeElement||i.type==='number')return;var v=String(i.value||''),m=/^([1-9]\d{3,})(,\d+)?$/.exec(v);if(m)i.value=m[1].replace(/\B(?=(\d{3})+(?!\d))/g,'.')+(m[2]||'')})}
function taAuto(e){if(!e||!e.style)return;e.style.height='auto';var h=e.scrollHeight+2;if(h>40)e.style.height=Math.min(h,window.innerHeight*0.6)+'px'}
function pacCrop(prod){
  var p=noAcc(String(prod||'')).toUpperCase().replace(/^(FS|FY|FL|FC|CI)\s+/,'').trim();
  if(/^ALMENDR/.test(p))return 'Almendro'; if(/^PISTACH/.test(p))return 'Pistacho'; if(/^OLIV/.test(p))return 'Olivo';
  if(/^(VINED|VID|VINA|UVA|VINEDO)/.test(p))return 'Viña'; if(/^(CEBADA|TRIGO|AVENA|CENTENO|TRITICALE|CEREAL)/.test(p))return 'Cereal'; if(/^BARBECHO/.test(p))return 'Barbecho';
  return '';
}
function pacCropVar(v){var p=noAcc(String(v||'')).toLowerCase().trim(); if(!p)return ''; for(var c in TABVAR){if(TABVAR[c].test(p))return c} return ''}
function pacParseRow(txt){
  var t=String(txt).trim().split(/\s+/);
  if(t.length<14||!/^\d+$/.test(t[0])||!/^\d{2}$/.test(t[1])||!/^\d{1,3}$/.test(t[2]))return null;
  var i=3, name=[]; while(i<t.length&&!/^\d+$/.test(t[i])){name.push(t[i]);i++}
  if(!name.length)return null;
  var dn=/^\d+[.,]\d+$/, it=/^\d+$/;
  var a=t.slice(i,i+9); if(a.length<9||!(it.test(a[0])&&it.test(a[1])&&it.test(a[2])&&it.test(a[3])&&it.test(a[4])&&dn.test(a[5])&&dn.test(a[6])&&dn.test(a[7])&&it.test(a[8])))return null;
  var rest=t.slice(i+9), k=-1; for(var j=0;j<rest.length;j++){if(/^[SR]$/i.test(rest[j])){k=j;break}}
  var pv=k>=0?rest.slice(0,k):rest.slice(0,4), riego=k>=0?(/^R$/i.test(rest[k])?'regadio':'secano'):'';
  var prod,vari;
  if(pv.length&&/^(FS|FY|FL|FC|CI)$/.test(pv[0])&&pv.length>1){prod=pv[0]+' '+pv[1];vari=pv.slice(2).join(' ')}else{prod=pv[0]||'';vari=pv.slice(1).join(' ')}
  if(/^SIN\s+VARIEDAD$/i.test(vari))vari='';
  var f=function(s){return parseFloat(String(s).replace(',','.'))};
  return {prov:+t[1],muni:+t[2],mname:name.join(' '),agr:+a[0],zona:+a[1],pol:+a[2],par:+a[3],rec:+a[4],sigHa:f(a[5]),decHa:f(a[6]),prod:prod,vari:vari,riego:riego};
}
function pacXY(vt,m){return [vt[0]*m[4]+vt[2]*m[5]+vt[4],vt[1]*m[4]+vt[3]*m[5]+vt[5]]}
function pacRowsFromItems(pages){
  var rows=[], seen={}, year='';
  pages.forEach(function(pg){
    var items=pg.items, vt=pg.vt||[1,0,0,-1,0,0];
    var ys=[]; items.forEach(function(it){var s=it.str; if(!s||!s.trim())return; var xy=pacXY(vt,it.transform), y=xy[1], g=null; for(var i=0;i<ys.length;i++){if(Math.abs(ys[i].y-y)<=2.5){g=ys[i];break}} if(!g){g={y:y,w:[]};ys.push(g)} g.w.push({x:xy[0],s:s.trim()})});
    ys.sort(function(a,b){return a.y-b.y}).forEach(function(g){
      var txt=g.w.sort(function(a,b){return a.x-b.x}).map(function(w){return w.s}).join(' ');
      if(!year){var m=txt.match(/Campa[ñn]a\s*(\d{4})/);if(m)year=m[1]}
      var r=pacParseRow(txt); if(!r)return;
      var key=[r.prov,r.muni,r.agr,r.zona,r.pol,r.par,r.rec].join(':'); if(seen[key])return; seen[key]=1; r.key=key; rows.push(r);
    });
  });
  return {rows:rows,year:year};
}
var titleCase2=function(s){return String(s||'').toLowerCase().replace(/(^|\s)(\S)/g,function(m,a,b){return a+b.toUpperCase()})};
var isDesk=function(){return window.matchMedia&&window.matchMedia('(min-width:900px)').matches};
function xlChartXML(){return ''}
/* ═══ Plan de abonado (RD 1051/2022 y Decreto 80/2025 de CLM) ═══ */
/* ───────── Entrega 78: plan de abonado (RD 1051/2022) — primera parte ───────── */
/* f.pab = [{id, y:'2026' (campaña 2026/2027), cl:{a,o,i,p,v,src} (clima), ud:[unidad], c, m}]
   unidad = {id, n, ck (clave de cultivo para las tablas), rg:'secano'|'regadio', refs:[ref], zv:0|1, tp:1|2|0 (0 = según el suelo), inh:0|1,
     rp:ref representativo, rc:{ref:{pe,dc}} (pendiente %, distancia a cauce m), su:{f,lab,inf,prof,mo,ph,n,p,k,tex,cont},
     ag:{m3,no3}, ne:{mo:'cos'|'edad'|'man', pr:kg/ha, ex:{n,p,k} kg/t, ed:1..5, nb:{n,p,k}}, ap:{s:{},o:{},x:{}}, pg:{src:'plan'|'own', pid, rows:[{id,mo,n,d,rn,rp,rk,fo,eq,mi}]}} */
var PAB={cur:{},open:{}};
/* Decreto 80/2025 de CLM, anejo III: kg N/ha·año en zona vulnerable [tipo 1, tipo 2] */
var PAB_LIM={regadio:{almendro:[60,100],pistacho:[60,85],nogal:[90,120],melocotonero:[100,120],olivo:[85,115],vid:[70,90],frutal:[90,100],avena:[80,100],cebada:[115,135],trigo:[140,150],trigoduro:[150,170],maiz:[200,210]},
  secano:{almendro:[45,60],pistacho:[35,65],melocotonero:[65,70],olivo:[50,60],vid:[50,60],frutal:[65,70],avena:[50,60],cebada:[50,60],centeno:[25,30],trigo:[70,80],triticale:[45,50]}};
var PAB_CK=[['almendro','Almendro'],['pistacho','Pistacho'],['olivo','Olivo'],['nogal','Nogal'],['vid','Vid'],['melocotonero','Melocotonero'],['frutal','Otros frutales'],['trigo','Trigo'],['trigoduro','Trigo duro'],['cebada','Cebada'],['avena','Avena / cereal forrajero'],['centeno','Centeno'],['triticale','Triticale'],['maiz','Maíz'],['otro','Otro']];
/* extracciones por tonelada de cosecha [N, P₂O₅, K₂O] y de dónde salen */
/* Entrega 83: fuentes comprobadas en la guía del MAPA (Partes I y II, 2010) */
var PAB_EX={almendro:[[32,12,48],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 25, tabla 25.8, pág. 210: 80-30-120 kg/ha para 2.500 kg/ha de almendra en cáscara'],
  olivo:[[17.5,4.5,22.5],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 27, tabla 27.4, pág. 226: extracción de 15-20 de N, 4-5 de P₂O₅ y 20-25 de K₂O por t de aceituna (valor medio)'],
  pistacho:[[40,30,70],'Criterio técnico (no hay tabla oficial de extracciones del pistacho)'],
  nogal:[[35,15,40],'Criterio técnico (no hay tabla oficial de extracciones del nogal)'],
  trigo:[[28,12,20],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 16, tabla 16.2, pág. 125: extracciones medias por t de grano (grano + paja)'],trigoduro:[[28,12,20],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 16, tabla 16.2, pág. 125: extracciones medias por t de grano (grano + paja)'],
  avena:[[28,12,30],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 16, tabla 16.2, pág. 125: extracciones medias por t de grano (grano + paja)'],cebada:[[24,11,24],'Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010), cap. 16, tabla 16.2, pág. 125: extracciones medias por t de grano (grano + paja)']};
/* pistacho en formación (criterio técnico, no hay referencia oficial) [N, P₂O₅, K₂O] por año de plantación */
var PAB_ED=[[15,8,12],[35,15,30],[50,20,45],[65,25,65],[80,30,85]];
/* AEMET, Albacete Base Aérea (8175), normales 1981-2010 */
var PAB_CL={a:353,o:110,i:77,p:110,v:55,src:'AEMET, Albacete Base Aérea (8175), valores normales 1981-2010'};
var PAB_TEX=[['arenosa','Arenosa'],['franco-arenosa','Franco-arenosa'],['franca','Franca'],['franco-limosa','Franco-limosa'],['franco-arcillosa','Franco-arcillosa'],['franco-arcillo-arenosa','Franco-arcillo-arenosa'],['arcillosa','Arcillosa'],['otra','Otra']];
var PAB_NU=[['n','N'],['p','P₂O₅'],['k','K₂O']];
function pabCk(c){c=String(c||'').toLowerCase();
  if(/pistach/.test(c))return 'pistacho'; if(/almend/.test(c))return 'almendro'; if(/oliv/.test(c))return 'olivo'; if(/nogal/.test(c))return 'nogal'; if(/vi[ñn]|\bvid\b|uva/.test(c))return 'vid';
  if(/melocot/.test(c))return 'melocotonero'; if(/duro/.test(c))return 'trigoduro'; if(/trigo/.test(c))return 'trigo'; if(/cebada/.test(c))return 'cebada'; if(/avena|forraj/.test(c))return 'avena';
  if(/centeno/.test(c))return 'centeno'; if(/tritic/.test(c))return 'triticale'; if(/ma[ií]z/.test(c))return 'maiz'; if(/frut|cerez|ciruel|manzan|peral|albaric/.test(c))return 'frutal'; return 'otro'}
function pabCkName(k){return daLbl(PAB_CK,k)||'Otro'}
function pabRg(r){return r.riego==='regadio'?'regadio':'secano'}
function pabList(f){return (f.pab||[]).slice().sort(function(a,b){return String(b.y).localeCompare(String(a.y))})}
function pabCur(f){var l=pabList(f);if(!l.length)return null;var id=PAB.cur[f.id];return l.filter(function(p){return p.id===id})[0]||l[0]}
function pabCamp(y){return campLbl(+y+1)}
function pabY0(){var d=new Date();return d.getMonth()>=8?d.getFullYear():d.getFullYear()-1}
/* Entrega 80: una unidad puede llevar recintos de otra finca del mismo propietario (u.xr) y cada recinto sus «ha en este plan» (u.rc[ref].ph) */
function pabUdRecs(f,u){var m={};fincaRecs(f.id).forEach(function(r){m[r.ref]=r});var o=(u.refs||[]).map(function(x){return m[x]}).filter(Boolean);
  (u.xr||[]).forEach(function(x){var r=S.recintos[x];if(r&&r.fincaId!==f.id&&finca(r.fincaId)&&o.indexOf(r)<0)o.push(r)});return o}
function pabHaR(u,r){var x=((u.rc||{})[r.ref]||{}).ph;return abHas(x)&&+x>0?+x:haC(r)}
function pabUdHa(f,u){return pabUdRecs(f,u).reduce(function(a,r){return a+pabHaR(u,r)},0)}
function pabIsX(f,u,ref){return (u.xr||[]).indexOf(ref)>=0&&!!S.recintos[ref]&&S.recintos[ref].fincaId!==f.id}
/* partes de un recinto en los planes de una campaña: [{f,u,ha}] */
function pabShares(ref,y){var o=[], r=S.recintos[ref]; if(!r)return o;
  S.fincas.forEach(function(f2){(f2.pab||[]).forEach(function(p2){if(String(p2.y)!==String(y))return;(p2.ud||[]).forEach(function(u2){
    var own=(u2.refs||[]).indexOf(ref)>=0&&r.fincaId===f2.id, ext=(u2.xr||[]).indexOf(ref)>=0&&r.fincaId!==f2.id;
    if(own||ext)o.push({f:f2,u:u2,ha:pabHaR(u2,r)})})})});
  return o}
/* recintos repartidos: ref -> campaña más reciente en la que están repartidos */
function pabSplits(){var m={};S.fincas.forEach(function(f2){(f2.pab||[]).forEach(function(p2){(p2.ud||[]).forEach(function(u2){(u2.xr||[]).forEach(function(x){var r=S.recintos[x];if(!r||r.fincaId===f2.id)return;if(!m[x]||+p2.y>+m[x])m[x]=p2.y})})})});return m}
function pabSplitTxt(ref,y){return pabShares(ref,y).map(function(s){return '«'+s.f.nombre+'» '+ha2(s.ha)+' ha'}).join(' · ')}
/* aviso en los formularios del cuaderno (tratamiento, abonado, fertirrigación) */
function pabSplitWarn(f,refs){
  var M=pabSplits(), ks=Object.keys(M); if(!ks.length||!f)return '';
  var sel=(refs||[]).filter(function(x){return S.recintos[x]&&S.recintos[x].fincaId===f.id}); if(!sel.length)sel=fincaRecs(f.id).map(function(r){return r.ref});
  var L=[];
  ks.forEach(function(x){var r=S.recintos[x], y=M[x];
    if(r.fincaId===f.id){if(sel.indexOf(x)>=0)L.push('El recinto '+refLabel(x)+' está repartido ('+pabSplitTxt(x,y)+'). Aquí se cuenta con sus '+ha2(r.ha)+' ha del SIGPAC.')}
    else{var mine=pabShares(x,y).filter(function(s){return s.f.id===f.id});if(mine.length){var o=finca(r.fincaId);L.push('Una parte del recinto '+refLabel(x)+' ('+ha2(mine.reduce(function(a,s){return a+s.ha},0))+' ha) es de esta finca, pero el recinto está en «'+(o?o.nombre:'otra finca')+'»: lo que se haga en él se anota allí.')}}});
  return L.map(function(t){return '⚠ '+t}).join('\n');
}
function pabSplitBox(id,f,refs){var t=pabSplitWarn(f,refs);return '<div id="'+id+'" class="small trwarn" style="margin-top:6px;white-space:pre-line"'+(t?'':' hidden')+'>'+esc(t)+'</div>'}
function pabSplitUpd(id,f,refs){var e=$('#'+id);if(!e)return;var t=pabSplitWarn(f,refs);e.textContent=t;e.hidden=!t}
/* recintos de otras fincas del mismo propietario que se pueden añadir a una unidad */
function pabXCands(f,p){var g=f.grupo&&grupo(f.grupo); if(!g)return [];var used={};(p.ud||[]).forEach(function(u){(u.xr||[]).forEach(function(x){used[x]=1})});
  var o=[];grupoFincas(g.id).forEach(function(f2){if(f2.id===f.id)return;fincaRecs(f2.id).forEach(function(r){if(!used[r.ref])o.push({r:r,f:f2})})});
  return o.sort(function(a,b){return a.f.nombre.localeCompare(b.f.nombre,'es')||a.r.ref.localeCompare(b.r.ref,'es',{numeric:true})})}
function pabAddX(f,ui,ref){var p=pabCur(f); if(!p)return; var u=p.ud[ui], r=S.recintos[ref]; if(!u||!r||r.fincaId===f.id){renderFincas();return}
  undoPush(); (u.xr||(u.xr=[])).push(ref); if(!u.rp)u.rp=ref; p.m=Date.now(); save(); renderFincas();
  undoToast('Recinto '+refLabel(ref)+' añadido a «'+u.n+'»: pon sus hectáreas en «ha en este plan»')}
/* unidades propuestas: recintos agrupados por cultivo y régimen (el titular puede juntarlas o separarlas: art. 3 v) */
function pabAuto(f){
  var m={},o=[];
  fincaRecs(f.id).slice().sort(function(a,b){return a.ref.localeCompare(b.ref,'es',{numeric:true})}).forEach(function(r){
    var c=r.cultivo||NOCROP, g=pabRg(r), k=c+'§'+g, u=m[k];
    if(!u){u=m[k]={id:uid(),n:c+' '+(g==='regadio'?'regadío':'secano'),ck:pabCk(c),rg:g,refs:[],zv:0,tp:0,inh:0,rc:{},su:{},ag:{},ne:{mo:'cos',pr:'',ex:null,ed:1},ap:{s:{},o:{},x:{}},pg:{src:'own',rows:[]}};o.push(u)}
    u.refs.push(r.ref);
  });
  o.forEach(function(u){var rs=pabUdRecs(f,u);if(rs.length){var big=rs.slice().sort(function(a,b){return haC(b)-haC(a)})[0];u.rp=big.ref}});
  return o;
}
/* tipo de parcela del anejo III: tipo 1 = suelo ligero (arenosa, franco-arenosa o franca); pasa a tipo 2 si se fertirriga */
function pabTipo(f,u,p){
  if(u.tp===1||u.tp===2)return {t:u.tp,why:'elegido a mano'};
  if(u.leg)return {t:1,why:'el cultivo precedente es una leguminosa'};
  var tx=(u.su||{}).tex, fr=u.rg==='regadio'&&pabFert(f,u,p);
  if(!tx)return {t:1,why:'sin textura del suelo: se toma tipo 1 por prudencia'};
  var lig=tx==='arenosa'||tx==='franco-arenosa'||tx==='franca';
  /* Entrega 143: «suelo franco» (no «suelo franca») */
  var sn=pabSuelo(tx);
  if(lig&&fr)return {t:2,why:'suelo '+sn+', pero se fertirriga: pasa a tipo 2'};
  return {t:lig?1:2,why:'suelo '+sn+(lig?' (ligero)':' (no es ligero)')};
}
function pabSuelo(tx){return daLbl(PAB_TEX,tx).toLowerCase().replace(/a$/,'o').replace(/a-/g,'o-')}
function pabFert(f,u,p){var pg=u.pg||{};return (pg.rows||[]).some(function(q){return /fertirr/i.test(q.fo||'')})}
function pabLim(f,u,p){
  var T=PAB_LIM[u.rg]||{}, ck=u.ck, ty=pabTipo(f,u,p), v=T[ck], via='';
  if(!v&&u.rg==='secano'&&ck==='nogal'){v=T.frutal;via=' (el anejo III solo recoge el nogal en regadío: se usa «otros frutales» de secano)'}
  if(!v)return {ok:false,t:ty};
  var b=v[ty.t-1], x=u.inh?Math.round(b*1.1):b;
  return {ok:true,t:ty,base:b,lim:x,via:via,txt:'Decreto 80/2025 de Castilla-La Mancha, anejo III: '+pabCkName(ck).toLowerCase()+' '+(u.rg==='regadio'?'regadío':'secano')+' '+v[0]+' kg N/ha·año en parcela tipo 1 y '+v[1]+' en tipo 2'+via+'. Se aplica tipo '+ty.t+': '+ty.why+'.'+(u.inh?' +10 % por usar fertilizantes con inhibidores.':'')};
}
function pabPlan(f,u){var pg=u.pg||{};return pg.pid?(f.plan||[]).filter(function(p){return p.id===pg.pid})[0]:null}
function pabN3(o){o=o||{};return {n:+o.n||0,p:+o.p||0,k:+o.k||0}}
/* cálculo de una unidad */
function pabCalc(f,u,P){
  var rs=pabUdRecs(f,u), ha=rs.reduce(function(a,r){return a+pabHaR(u,r)},0), sig=rs.reduce(function(a,r){return a+(r.ha||0)},0), ne=u.ne||{}, nb={n:0,p:0,k:0}, src='';
  if(ne.mo==='edad'){var e=PAB_ED[Math.max(1,Math.min(5,+ne.ed||1))-1];nb={n:e[0],p:e[1],k:e[2]};src='Plantación en formación ('+(+ne.ed||1)+'º verde): criterio técnico para pistacho, sin referencia oficial'}
  else if(ne.mo==='man'){nb=pabN3(ne.nb);src=ne.src?String(ne.src):'Necesidad fijada por el técnico'}
  else{var ex=ne.ex||(PAB_EX[u.ck]?{n:PAB_EX[u.ck][0][0],p:PAB_EX[u.ck][0][1],k:PAB_EX[u.ck][0][2]}:null), pr=+ne.pr||0;
    if(ex){nb={n:pr/1000*(+ex.n||0),p:pr/1000*(+ex.p||0),k:pr/1000*(+ex.k||0)}} var ref=PAB_EX[u.ck], igual=!!(ne.ex&&ref&&['n','p','k'].every(function(k,j){return Math.abs((+ne.ex[k]||0)-ref[0][j])<0.01}));
    src=ne.ex&&!igual?'Criterio técnico (extracciones fijadas por el técnico)':(ref?ref[1]:'Sin extracciones de referencia: pon las del cultivo')}
  var ag=u.ag||{}, nAg=u.rg==='regadio'&&abHas(ag.m3)&&abHas(ag.no3)?Math.round(ag.m3*ag.no3*14/62/1000*10)/10:0;
  var ap=u.ap||{}, S0=pabN3(ap.s), O0=pabN3(ap.o), X0=pabN3(ap.x);
  /* programa */
  var pg=u.pg||{}, pl={n:0,p:0,k:0}, rows=[], pln=null;
  if(u.d0===1){}
  else (pg.rows||[]).forEach(function(q){var d=+q.d||0,uf={n:d*(+q.rn||0)/100,p:d*(+q.rp||0)/100,k:d*(+q.rk||0)/100};pl.n+=uf.n;pl.p+=uf.p;pl.k+=uf.k;rows.push({id:q.id,org:q.org===1,mo:q.mo||'',n:q.n||'',d:d,u:'kg',r:[q.rn,q.rp,q.rk].map(function(x){return trTxt(+x||0,1)}).join('-'),uf:uf,fo:q.fo||'',eq:q.eq||'',mi:q.mi||''})});
  var B={};
  PAB_NU.forEach(function(x){var k=x[0], ag0=k==='n'?nAg:0, net=Math.max(0,nb[k]-S0[k]-ag0-O0[k]-X0[k]), dv=pl[k]-net;
    B[k]={nb:nb[k],s:S0[k],a:ag0,o:O0[k],x:X0[k],net:net,pl:pl[k],dv:dv,st:Math.abs(dv)<=1?'Ajustado':(dv<0?'Déficit':'Exceso')}});
  var L=pabLim(f,u,P), lst=!u.zv?'No está en zona vulnerable: el anejo III no se aplica':(!L.ok?'Elige el cultivo para conocer el límite':(pl.n<=L.lim+0.05?'Cumple':'Supera el límite: hace falta balance de N justificado (y conservarlo 3 años)'));
  var ex=u.rg==='secano'&&ha<=10;
  return {rs:rs,ha:ha,sig:sig,nb:nb,src:src,nAg:nAg,pl:pl,rows:rows,pln:pln,B:B,L:L,lst:lst,lok:!u.zv||!L.ok||pl.n<=L.lim+0.05,ex:ex};
}
/* Entrega 84: caracterización de los abonos orgánicos del programa (anejo I del Decreto 80/2025 de CLM) */
function pabOrg(u,C){
  var lim=u.zv&&C.L&&C.L.ok?Math.min(C.L.lim,170):170, rows=(u.pg||{}).rows||[];
  return (C.rows||[]).filter(function(o){return o.org}).map(function(o){
    var q=rows.filter(function(x){return x.id===o.id})[0]||{}, rn=+q.rn||0, rp=+q.rp||0, rk=+q.rk||0;
    return {n:o.n||'',rn:rn,rp:rp,rk:rk,kn:rn*10,kp:rp*10,kk:rk*10,d:(+o.d||0)/1000,nN:o.uf?o.uf.n:0,max:rn>0?lim/(rn*10):null,lim:lim}});
}
function pabOrgHTML(u,C){var L=pabOrg(u,C);if(!L.length)return '';
  return '<div class="lbl" style="margin-top:8px">Caracterización del abono orgánico (riqueza garantizada en etiqueta)</div><div class="xs" style="overflow-x:auto"><table class="patb stk"><thead><tr><th>Abono</th><th>Riqueza N-P₂O₅-K₂O</th><th>kg/t</th><th>Dosis prevista</th><th>N aportado</th><th>Dosis máxima por N</th></tr></thead><tbody>'+
    L.map(function(x){return '<tr><td>'+esc(x.n)+'</td><td>'+trTxt(x.rn,1)+'-'+trTxt(x.rp,1)+'-'+trTxt(x.rk,1)+' %</td><td>'+trTxt(x.kn,0)+'-'+trTxt(x.kp,0)+'-'+trTxt(x.kk,0)+'</td><td class="r">'+trTxt(x.d,2)+' t/ha</td><td class="r">'+trTxt(x.nN,1)+' kg N/ha</td><td class="r">'+(x.max===null?'—':trTxt(x.max,2)+' t/ha')+'</td></tr>'}).join('')+
    '</tbody></table></div><div class="small">Contenido en kg/t = riqueza garantizada (%) × 10. Dosis máxima: la que llega al límite de N de '+nf(L[0].lim,0)+' kg/ha (anejo III y máximo de 170 kg N/ha de origen orgánico).</div>'}
/* Entrega 85: pendientes (apdo. 4.3.2 del Decreto 80/2025) y perímetros de protección (apdo. 4.3.3) */
function pabSlope(u,C){
  var rc=u.rc||{}, x15=[], b8=[], pmax=-1, isB=u.ck==='otro'&&/barbech/i.test(String(u.n||''));
  (C.rs||[]).forEach(function(r){var pe=(rc[r.ref]||{}).pe; if(!abHas(pe))return; pe=+pe; if(pe>pmax)pmax=pe;
    if(pe>15)x15.push({ref:r.ref,pe:pe}); else if(pe>8&&(isB||/barbech/i.test(String(r.cultivo||''))))b8.push({ref:r.ref,pe:pe})});
  return {x15:x15,b8:b8,pmax:pmax};
}
function pabSlopeTxt(L,lim){return L.map(function(x){return refLabel(x.ref)+' ('+trTxt(x.pe,1)+' %)'}).join(', ')}
function pabPerim(u,C){var rc=u.rc||{};return (C.rs||[]).filter(function(r){return (rc[r.ref]||{}).pp===1})}
function pabStTxt(f,u,C){if(C.ex)return 'Exenta de plan';if(u.d0===1)return 'Sin abonado';if(!C.lok)return 'Revisar límite de N';var s=PAB_NU.map(function(x){return C.B[x[0]].st});if(s.indexOf('Déficit')>=0)return 'Déficit';if(s.indexOf('Exceso')>=0)return 'Exceso';return 'Ajustado'}
/* ---- pantalla ---- */
function pabIn(path,val,ph,w,mode){return '<input class="in pa-in" data-p="'+esc(path)+'"'+(mode?' inputmode="'+mode+'"':' inputmode="decimal"')+' value="'+esc(val===undefined||val===null?'':(typeof val==='number'?abEd(val):val))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+(w?' style="width:'+w+'"':'')+' autocomplete="off">'}
function pabTx(path,val,ph){return '<input class="in pa-in" data-p="'+esc(path)+'" value="'+esc(val||'')+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off">'}
function pabChips(path,cur,L){return '<div class="chips">'+L.map(function(x){return '<button type="button" class="chip'+(String(cur)===String(x[0])?' on':'')+'" data-act="pa-set" data-p="'+esc(path)+'" data-v="'+esc(x[0])+'">'+x[1]+'</button>'}).join('')+'</div>'}
function pabF(lbl,inner){return '<div class="lbl" style="margin-top:8px">'+lbl+'</div>'+inner}
function pabNum(v,d){return abHas(v)||typeof v==='number'?trTxt(+v,d==null?1:d):''}
function pabHTML(f){
  if(GTOK)return '<div class="small">El plan de abonado no está disponible en esta vista.</div>';
  pabSync(f); var l=pabList(f), p=pabCur(f), h='';
  h+='<div class="acts" style="margin-top:0">'+l.map(function(x){return '<button type="button" class="chip'+(p&&x.id===p.id?' on':'')+'" data-act="pa-pick" data-id="'+esc(x.id)+'">Campaña '+esc(pabCamp(x.y))+'</button>'}).join('')+'<button class="btn'+(l.length?'':' primary')+'" data-act="pa-new">+ Nuevo plan de abonado</button></div>';
  /* Entrega 143: antes de crear el plan, se elige la campaña */
  if(PAB.ask===f.id){var y0=pabY0(), ex={}; (f.pab||[]).forEach(function(x){ex[+x.y]=1}); var yo=[]; for(var yy=y0-2;yy<=y0+2;yy++)if(!ex[yy])yo.push(yy);
    var ys=yo.indexOf(y0)>=0?y0:(yo.filter(function(v){return v>y0})[0]||yo[yo.length-1]);
    h+='<div class="note diaf" style="margin-top:8px"><b>Nuevo plan de abonado</b>'+(yo.length?'<div class="lbl" style="margin-top:8px">¿De qué campaña?</div><select id="pa-ny" class="in" aria-label="Campaña del plan de abonado">'+yo.map(function(v){return '<option value="'+v+'"'+(v===ys?' selected':'')+'>Campaña '+esc(pabCamp(v))+'</option>'}).join('')+'</select><div class="acts"><button class="btn primary" data-act="pa-new-ok">Crear el plan</button><button class="btn" data-act="pa-new-no">Cancelar</button></div>':'<div class="small" style="margin-top:6px">Ya hay planes de todas las campañas cercanas.</div><div class="acts"><button class="btn" data-act="pa-new-no">Cerrar</button></div>')+'</div>'}
  if(!p)return h+'<div class="small">El plan de abonado (Real Decreto 1051/2022) es obligatorio desde el 1 de septiembre de 2026 para cada unidad de producción, salvo las de secano de 10 ha o menos. Al crearlo, los recintos se agrupan por cultivo y riego; luego puedes moverlos de una unidad a otra.</div>';
  var egb=expGGet(f); if(egb){var gr0={};grRecs(f,egb).forEach(function(r){gr0[r.ref]=1});var us=(p.ud||[]).filter(function(u){return (u.refs||[]).some(function(x){return gr0[x]})}).map(function(u){return u.n});h+='<div class="note" style="margin-top:0">El plan de abonado va por unidades de producción (cultivo y riego), no por grupos. Los recintos de «'+esc(grOf(f,egb).n)+'» están en: <b>'+(us.length?esc(us.join(', ')):'ninguna unidad todavía')+'</b>.</div>'}
  var cl=p.cl||PAB_CL, sum=(+cl.o||0)+(+cl.i||0)+(+cl.p||0)+(+cl.v||0);
  var Cs=(p.ud||[]).map(function(u){return pabCalc(f,u,p)});
  h+='<div class="plhead"><div><div class="exptt" style="font-size:20px">Plan de abonado · campaña '+esc(pabCamp(p.y))+'</div><div class="small">'+(p.ud||[]).length+((p.ud||[]).length===1?' unidad de producción':' unidades de producción')+' · '+haTxt(fincaRecs(f.id))+'</div></div><div class="acts" style="margin:6px 0 0"><button class="btn primary sm" data-act="pa-pdf">PDF</button><button class="btn sm" data-act="pa-xlsx">Excel</button><button class="btn sm" data-act="pa-ud-new">+ Unidad</button><button class="btn sm lk danger" data-act="pa-del">Borrar plan</button></div></div>';
  /* resumen */
  if((p.ud||[]).length){
    h+='<h3 style="margin-top:12px">Resumen</h3><div class="xs" style="overflow-x:auto"><table class="patb stk"><thead><tr><th>Unidad</th><th>ha</th><th>Producción objetivo</th><th>N</th><th>P₂O₅</th><th>K₂O</th><th>Estado</th></tr></thead><tbody>';
    (p.ud||[]).forEach(function(u,i){var C=Cs[i], st=pabStTxt(f,u,C);h+='<tr><td><b>'+esc(u.n)+'</b></td><td class="r">'+ha2(C.ha)+'</td><td class="r">'+((u.ne||{}).mo==='cos'&&abHas((u.ne||{}).pr)?nf(+u.ne.pr,0)+' kg/ha':((u.ne||{}).mo==='edad'?'en formación':'—'))+'</td><td class="r">'+trTxt(C.pl.n,1)+'</td><td class="r">'+trTxt(C.pl.p,1)+'</td><td class="r">'+trTxt(C.pl.k,1)+'</td><td><span class="pa-st pa-'+(st==='Ajustado'?'ok':(st==='Exenta de plan'||st==='Sin abonado')?'ex':'bad')+'">'+st+'</span></td></tr>'});
    h+='</tbody></table></div><div class="small">U.F./ha planificadas. «Ajustado» = lo planificado está a ±1 U.F. de la necesidad neta.</div>';
  }
  /* clima (art. 6.3) */
  h+='<details class="more pa-box"'+(PAB.open.cl?' open':'')+' data-pa-o="cl"><summary><b>Precipitación de la zona</b> <span class="small">(art. 6.3)</span></summary>';
  h+='<div class="pa-g">'+[['a','Media anual (mm)'],['o','Otoño, sep-nov (mm)'],['i','Invierno, dic-feb (mm)'],['p','Primavera, mar-may (mm)'],['v','Verano, jun-ago (mm)']].map(function(x){return '<div>'+pabF(x[1],pabIn('cl.'+x[0],cl[x[0]]))+'</div>'}).join('')+'</div>'+pabF('Fuente',pabTx('cl.src',cl.src));
  h+='<div class="small" style="margin-top:6px">Suma de las estaciones: '+nf(sum,0)+' mm'+(abHas(cl.a)&&Math.abs(sum-cl.a)>2?' · <b>no cuadra con la media anual</b>':'')+'.</div></details>';
  h+=pabPlanMore(f,p);
  /* unidades */
  (p.ud||[]).forEach(function(u,i){h+=pabUdHTML(f,p,u,Cs[i],i)});
  if(!(p.ud||[]).length)h+='<div class="note">No hay unidades. Pulsa «+ Unidad».</div>';
  return h;
}
function pabUdHTML(f,p,u,C,i){
  var P='ud.'+i+'.', op=PAB.open[u.id]!==undefined?PAB.open[u.id]:(i===0&&(p.ud||[]).length===1), st=pabStTxt(f,u,C), h='';
  h+='<details class="more pa-box"'+(op?' open':'')+' data-pa-o="'+esc(u.id)+'"><summary><b>'+esc(u.n)+'</b> <span class="small">· '+ha2(C.ha)+' ha · '+C.rs.length+(C.rs.length===1?' recinto':' recintos')+'</span> <span class="pa-st pa-'+(st==='Ajustado'?'ok':(st==='Exenta de plan'||st==='Sin abonado')?'ex':'bad')+'">'+st+'</span></summary>';
  if(C.ex)h+='<div class="note">Unidad de secano de '+ha2(C.ha)+' ha: <b>no necesita plan de abonado</b> (art. 4.2 b del RD 1051/2022, secano de 10 ha o menos). Puedes rellenarlo igualmente.</div>';
  else if(u.rg==='regadio'&&C.ha<=10)h+='<div class="small" style="margin-top:6px">Aunque tenga menos de 10 ha, al ser de regadío necesita plan.</div>';
  /* identificación */
  h+='<div class="pa-g"><div>'+pabF('Nombre de la unidad',pabTx(P+'n',u.n))+'</div><div>'+pabF('Cultivo (para las tablas de límites y extracciones)','<select class="trsel pa-in" data-p="'+P+'ck" style="width:100%">'+PAB_CK.map(function(x){return '<option value="'+x[0]+'"'+(u.ck===x[0]?' selected':'')+'>'+x[1]+'</option>'}).join('')+'</select>')+'</div></div>';
  h+=pabF('Régimen',pabChips(P+'rg',u.rg,[['secano','Secano'],['regadio','Regadío']]));
  h+=pabF('¿Está en zona vulnerable a nitratos?',pabChips(P+'zv',u.zv?1:0,[[1,'Sí'],[0,'No']]));
  if(u.zv){var ty=pabTipo(f,u,p);h+=pabF('Tipo de parcela (anejo III)',pabChips(P+'tp',u.tp||0,[[0,'Según el suelo (tipo '+ty.t+')'],[1,'Tipo 1'],[2,'Tipo 2']])+'<div class="small">'+esc(ty.why)+'. Tipo 1 = suelo ligero (arenoso, franco-arenoso o franco) o precedente leguminosa; tipo 2 = el resto, o si se fertirriga.</div>');
    h+=pabF('¿Fertilizantes con inhibidores de la nitrificación o de la ureasa? (+10 % de N)',pabChips(P+'inh',u.inh?1:0,[[1,'Sí'],[0,'No']]))}
  /* recintos (art. 6.1) */
  h+='<h4 style="margin:14px 0 4px">Recintos <span class="small">(art. 6.1)</span></h4><div class="xs" style="overflow-x:auto"><table class="patb stk"><thead><tr><th>Recinto SIGPAC</th><th>ha cultiv.</th><th>ha en este plan</th><th>Cultivo</th><th>Pendiente %</th><th>Perímetro de protección</th><th>Repr.</th><th>Unidad</th></tr></thead><tbody>';
  C.rs.forEach(function(r){var rc=(u.rc||{})[r.ref]||{}, xo=pabIsX(f,u,r.ref), of=xo?finca(r.fincaId):null, sh=pabShares(r.ref,p.y), oth=sh.filter(function(s){return s.u!==u}), tot=sh.reduce(function(a,s){return a+s.ha},0);
    var nt=(xo?'<div class="small">Recinto de la finca «'+esc(of?of.nombre:'')+'»</div>':'')+(oth.length?'<div class="small">También en '+oth.map(function(s){return (s.f.id===f.id?'«'+esc(s.u.n)+'»':'«'+esc(s.f.nombre)+'»')+' ('+ha2(s.ha)+' ha)'}).join(', ')+'</div>':'')+(tot>(r.ha||0)+0.00005?'<div class="small trwarn">Las partes suman '+ha2(tot)+' ha y el recinto tiene '+ha2(r.ha||0)+' ha en el SIGPAC</div>':'');
    h+='<tr><td>'+esc(refLabel(r.ref))+(r.nombre?' <span class="small">'+esc(r.nombre)+'</span>':'')+nt+'</td><td class="r">'+ha4(haC(r))+'</td><td>'+pabIn(P+'rc.'+r.ref+'.ph',rc.ph,ha4(haC(r)),'80px')+'</td><td>'+esc(r.cultivo||'-')+'</td><td>'+pabIn(P+'rc.'+r.ref+'.pe',rc.pe,'','64px')+(abHas(rc.pe)&&+rc.pe>15?'<div class="small"><b>más del 15 %: no se abona</b></div>':'')+'</td><td>'+pabChips(P+'rc.'+r.ref+'.pp',rc.pp===1?1:(rc.pp===0?0:''),[[0,'No'],[1,'Sí']])+'</td><td><button type="button" class="chip'+(u.rp===r.ref?' on':'')+'" data-act="pa-set" data-p="'+P+'rp" data-v="'+esc(r.ref)+'" aria-label="Recinto representativo">'+(u.rp===r.ref?'Sí':'—')+'</button></td><td><select class="trsel pa-mv" data-u="'+i+'" data-ref="'+esc(r.ref)+'">'+(p.ud||[]).map(function(x,j){return '<option value="'+j+'"'+(j===i?' selected':'')+'>'+esc(x.n)+'</option>'}).join('')+'<option value="new">Unidad nueva…</option>'+(xo?'<option value="out">Quitar de este plan</option>':'')+'</select></td></tr>'});
  h+='</tbody></table></div><div class="small">«Repr.» es el recinto representativo cuyos datos de suelo se usan (art. 6.2). «ha en este plan»: déjalo vacío para usar las ha cultivadas; pon otra cifra si solo una parte del recinto entra en este plan (por ejemplo, un recinto repartido entre dos titulares).</div>';
  var xc=pabXCands(f,p), gq=f.grupo&&grupo(f.grupo);
  if(xc.length)h+='<div style="margin-top:6px"><select class="trsel pa-add" data-u="'+i+'" style="max-width:100%"><option value="">+ Añadir recinto de otra finca de «'+esc(gq.nombre)+'»…</option>'+xc.map(function(o){return '<option value="'+esc(o.r.ref)+'">'+esc(refLabel(o.r.ref)+' · '+o.f.nombre+' · '+(o.r.cultivo||NOCROP)+' · '+ha2(haC(o.r))+' ha')+'</option>'}).join('')+'</select></div>';
  /* suelo (art. 6.2) */
  var su=u.su||{}, rp=u.rp&&S.recintos[u.rp];
  h+='<h4 style="margin:14px 0 4px">Suelo del recinto representativo <span class="small">(art. 6.2)'+(rp?' · '+esc(refLabel(rp.ref)):'')+'</span></h4><div class="pa-g">'+
    [['f','Fecha de muestreo','','',2],['lab','Laboratorio','','',1],['inf','Nº de informe','','',1],['prof','Profundidad (cm)','0-30',''],['mo','Materia orgánica (%)'],['ph','pH'],['n','N total (mg/kg)'],['p','P asimilable, Olsen (mg/kg)'],['k','K asimilable (mg/kg)'],['cont','Contaminantes analizados','No analizados','',1]].map(function(x){return '<div>'+pabF(x[1],x[4]===2?'<input type="date" class="in pa-in" data-p="'+esc(P+'su.'+x[0])+'" value="'+esc(/^\s*\d{4}\s*$/.test(String(su[x[0]]||''))?'':(pabIso(su[x[0]])||''))+'" aria-label="'+esc(x[1])+'">'+(su[x[0]]&&(!pabIso(su[x[0]])||/^\s*\d{4}\s*$/.test(String(su[x[0]])))?'<div class="small">Apuntado antes como «'+esc(su[x[0]])+'»</div>':''):x[4]?pabTx(P+'su.'+x[0],su[x[0]],x[2]):pabIn(P+'su.'+x[0],su[x[0]],x[2]))+'</div>'}).join('')+
    '<div>'+pabF('Textura','<select class="trsel pa-in" data-p="'+P+'su.tex" style="width:100%"><option value="">—</option>'+PAB_TEX.map(function(x){return '<option value="'+x[0]+'"'+(su.tex===x[0]?' selected':'')+'>'+x[1]+'</option>'}).join('')+'</select>')+'</div></div>';
  /* agua (art. 6.3) */
  if(u.rg==='regadio'){var ag=u.ag||{};h+='<h4 style="margin:14px 0 4px">Agua de riego <span class="small">(art. 6.3)</span></h4><div class="pa-g"><div>'+pabF('Dotación prevista (m³/ha)',pabIn(P+'ag.m3',ag.m3))+'</div><div>'+pabF('Dotación máxima concedida (m³/ha·año)',pabIn(P+'ag.max',ag.max))+'</div><div>'+pabF('Nitratos del agua (mg NO₃/L)',pabIn(P+'ag.no3',ag.no3))+'</div><div>'+pabF('Análisis del agua (laboratorio, nº, fecha)',pabTx(P+'ag.src',ag.src))+'</div></div><div class="small" style="margin-top:4px">N que aporta el agua: <b>'+trTxt(C.nAg,1)+' kg N/ha</b> (m³ × mg NO₃/L × 14/62 ÷ 1.000). Se descuenta solo en el balance.</div>'}
  /* necesidades */
  var ne=u.ne||{}, ex=ne.ex||(PAB_EX[u.ck]?{n:PAB_EX[u.ck][0][0],p:PAB_EX[u.ck][0][1],k:PAB_EX[u.ck][0][2]}:{});
  h+='<h4 style="margin:14px 0 4px">Necesidades del cultivo</h4>'+pabChips(P+'ne.mo',ne.mo||'cos',[['cos','Por cosecha esperada'],['edad','En formación (pistacho, por edad)'],['man','A mano']]);
  if((ne.mo||'cos')==='cos')h+='<div class="pa-g"><div>'+pabF('Producción objetivo (kg/ha)',pabIn(P+'ne.pr',ne.pr))+'</div>'+PAB_NU.map(function(x){return '<div>'+pabF('Extracción '+x[1]+' (kg por t)',pabIn(P+'ne.ex.'+x[0],ex[x[0]]))+'</div>'}).join('')+'</div>';
  else if(ne.mo==='edad')h+=pabF('Años de la plantación (verdes)',pabChips(P+'ne.ed',ne.ed||1,[[1,'1º'],[2,'2º'],[3,'3º'],[4,'4º'],[5,'5º']]));
  else h+='<div class="pa-g">'+PAB_NU.map(function(x){return '<div>'+pabF('Necesidad '+x[1]+' (kg/ha)',pabIn(P+'ne.nb.'+x[0],(ne.nb||{})[x[0]]))+'</div>'}).join('')+'<div>'+pabF('Producción esperada (kg/ha; 0 = sin cosecha)',pabIn(P+'ne.pr',ne.pr))+'</div></div>'+pabF('Justificación de la necesidad',pabTx(P+'ne.src',ne.src,'Ej. Plantación en formación (2º verde): criterio técnico'));
  h+='<div class="small" style="margin-top:4px">'+esc(C.src)+'.</div>';
  /* aportes */
  var ap=u.ap||{};
  h+='<h4 style="margin:14px 0 4px">Aportes que ya hay (U.F./ha)</h4><div class="xs" style="overflow-x:auto"><table class="patb stk"><thead><tr><th></th>'+PAB_NU.map(function(x){return '<th>'+x[1]+'</th>'}).join('')+'</tr></thead><tbody>'+[['s','Suelo'],['o','Orgánico (estiércol, restos…)'],['x','Otros']].map(function(y){return '<tr><td>'+y[1]+'</td>'+PAB_NU.map(function(x){return '<td>'+pabIn(P+'ap.'+y[0]+'.'+x[0],(ap[y[0]]||{})[x[0]],'0','70px')+'</td>'}).join('')+'</tr>'}).join('')+'</tbody></table></div><div class="small">El abono orgánico que pongas en el programa no va aquí: ya cuenta como planificado.</div>';
  /* programa (art. 6.4) */
  var pg=u.pg||{}, pls=(f.plan||[]).filter(function(x){return String(x.y)===String(+p.y+1)||String(x.y)===String(p.y)});
  h+='<h4 style="margin:14px 0 4px">Programa previsto <span class="small">(art. 6.4)</span></h4>'+pabF('¿Sin abonado esta campaña?',pabChips(P+'d0',u.d0===1?1:0,[[1,'Sí, dosis 0'],[0,'No, hay programa']]));
  if(u.d0===1)h+='<div class="note">No se prevé fertilizar esta campaña (programa a dosis 0). No hace falta programa, análisis de suelo, medidas de amoniaco ni restricciones de aplicación; si se decide abonar, quita esta opción y rellena el programa.</div>';
  if(u.d0===1){}
  else{
    h+='<div style="overflow-x:auto"><table class="patb pa-prog stk"><thead><tr><th>Momento</th><th>Producto</th><th>Dosis kg/ha</th><th>N %</th><th>P₂O₅ %</th><th>K₂O %</th><th>U.F. N-P-K</th><th>Forma</th><th>Equipo</th><th>Medida de mitigación</th><th>Orgánico</th><th></th></tr></thead><tbody>';
    (pg.rows||[]).forEach(function(q,j){var R=P+'pg.rows.'+j+'.',o=C.rows[j];h+='<tr><td>'+pabTx(R+'mo',q.mo,'Marzo-abril')+'</td><td>'+pabTx(R+'n',q.n,'Producto')+'</td><td>'+pabIn(R+'d',q.d,'','70px')+'</td><td>'+pabIn(R+'rn',q.rn,'','52px')+'</td><td>'+pabIn(R+'rp',q.rp,'','52px')+'</td><td>'+pabIn(R+'rk',q.rk,'','52px')+'</td><td class="r small">'+(o?trTxt(o.uf.n,1)+'-'+trTxt(o.uf.p,1)+'-'+trTxt(o.uf.k,1):'')+'</td><td>'+pabTx(R+'fo',q.fo,'Suelo')+'</td><td>'+pabTx(R+'eq',q.eq,'Abonadora')+'</td><td>'+pabTx(R+'mi',q.mi,'')+'</td><td>'+pabChips(R+'org',q.org===1?1:0,[[1,'Sí'],[0,'No']])+'</td><td><button type="button" class="btn sm danger" data-act="pa-row-del" data-u="'+i+'" data-j="'+j+'" aria-label="Quitar">✕</button></td></tr>'});
    h+='</tbody></table></div><div class="acts" style="margin:6px 0 0"><button class="btn sm" data-act="pa-row-add" data-u="'+i+'">+ Añadir aporte</button></div>';
  }
  /* balance */
  h+='<h4 style="margin:14px 0 4px">Balance de nutrientes (U.F./ha)</h4><div class="xs" style="overflow-x:auto"><table class="patb stk"><thead><tr><th>Nutriente</th><th>Necesidad bruta</th><th>Suelo</th><th>Agua</th><th>Orgánico</th><th>Otros</th><th>Necesidad neta</th><th>Planificado</th><th>Desviación</th><th>Estado</th></tr></thead><tbody>'+
    PAB_NU.map(function(x){var b=C.B[x[0]];return '<tr><td><b>'+x[1]+'</b></td><td class="r">'+trTxt(b.nb,1)+'</td><td class="r">'+trTxt(b.s,1)+'</td><td class="r">'+trTxt(b.a,1)+'</td><td class="r">'+trTxt(b.o,1)+'</td><td class="r">'+trTxt(b.x,1)+'</td><td class="r"><b>'+trTxt(b.net,1)+'</b></td><td class="r"><b>'+trTxt(b.pl,1)+'</b></td><td class="r">'+(b.dv>0?'+':'')+trTxt(b.dv,1)+'</td><td><span class="pa-st pa-'+(b.st==='Ajustado'?'ok':'bad')+'">'+b.st+'</span></td></tr>'}).join('')+'</tbody></table></div>';
  h+=pabF('Observaciones del balance',pabTx(P+'bobs',u.bobs,'Ej. Exceso de P₂O₅: aporte de fondo que se compensa en las campañas siguientes'));
  if(u.zv&&C.L.ok)h+='<div class="note" style="margin-top:6px"><b>Límite de N: '+nf(C.L.lim,0)+' kg N/ha·año</b> · planificado '+trTxt(C.pl.n,1)+' · <span class="pa-st pa-'+(C.lok?'ok':'bad')+'">'+esc(C.lst)+'</span><div class="small" style="margin-top:4px">'+esc(C.L.txt)+'</div></div>';
  else h+='<div class="small" style="margin-top:6px">'+esc(C.lst)+'.</div>';
  h+=pabUdMore(f,p,u,C,i);
  h+='<div class="acts" style="margin:10px 0 0"><button class="btn sm lk danger" data-act="pa-ud-del" data-u="'+i+'">Borrar unidad</button></div></details>';
  return h;
}
/* ---- guardar ---- */
function pabP(f){return pabCur(f)}
function pabSetPath(o,path,v){var ks=path.split('.'), x=o;
  /* las referencias SIGPAC llevan «:» pero no «.»; los índices son números */
  for(var i=0;i<ks.length-1;i++){var k=ks[i];if(x[k]===undefined||x[k]===null||typeof x[k]!=='object')x[k]=/^\d+$/.test(ks[i+1])?[]:{};x=x[k]}
  var last=ks[ks.length-1]; if(v===''||v===null||v===undefined)delete x[last]; else x[last]=v}
function pabIsTxt(path){return /^ud\.\d+\.(rot|bobs)$/.test(path)||/\.mo\.cub$/.test(path)||/\.ne\.src$/.test(path)||/^ud\.\d+\.prec$/.test(path)||/\.mo\.(obj|enm)$/.test(path)||/\.(mit|res)\.\w+\.d$/.test(path)||/^(fe|fa|ft)$/.test(path)||/^mods\.\d+\.(f|u|mo|ch|re)$/.test(path)||/^ud\.\d+\.(n|ck|rg|rp)$/.test(path)||/\.pg\.(src|pid)$/.test(path)||/\.pg\.rows\.\d+\.(mo|n|fo|eq|mi)$/.test(path)||/\.su\.(f|lab|inf|cont|tex)$/.test(path)||/\.ag\.src$/.test(path)||path==='cl.src'||/\.ne\.mo$/.test(path)}
function pabChange(f,path,raw,rerender){
  var p=pabCur(f); if(!p)return; var v=raw;
  if(/^AS\./.test(path)){S.x.asesor=Object.assign({},S.x.asesor||{});var s0=String(raw==null?'':raw).trim(),k0=path.slice(3);if(s0)S.x.asesor[k0]=s0;else delete S.x.asesor[k0];save();if(rerender)pabRedraw();return}
  if(pabIsTxt(path))v=String(raw==null?'':raw).trim();
  else{var s=String(raw==null?'':raw).trim(); if(s===''){v=''}else{v=trNum(s);if(v===null||isNaN(v)){toast('Pon un número (los decimales con coma)');renderFincas();return}}}
  if(/\.(zv|inh|tp|ed|d0)$/.test(path))v=+raw;
  if(path==='cl.src'||/^cl\./.test(path)){if(!p.cl)p.cl=Object.assign({},PAB_CL)}
  if(/\.ne\.ex\.[npk]$/.test(path)){var u=p.ud[+path.split('.')[1]];if(u&&!u.ne.ex){var d=PAB_EX[u.ck];u.ne.ex=d?{n:d[0][0],p:d[0][1],k:d[0][2]}:{}}}
  pabSetPath(p,path,v);
  if(/\.ck$/.test(path)){var u2=p.ud[+path.split('.')[1]];if(u2&&u2.ne)delete u2.ne.ex}
  p.m=Date.now(); save(); if(rerender)pabRedraw();
}
/* vuelve a dibujar sin perder la casilla en la que estás escribiendo */
function pabRedraw(){setTimeout(function(){var a=document.activeElement, k=a&&a.dataset?a.dataset.p:null, sy=window.scrollY; renderFincas();
  if(k){var n=null;$$('[data-p]').forEach(function(x){if(x.dataset.p===k&&x.tagName!=='BUTTON')n=x});if(n){try{n.focus({preventScroll:true})}catch(e){n.focus()}}}},0)}
function pabClick(b){
  var f=finca(EXP.fid); if(!f||GTOK)return; var a=b.dataset.act, p=pabCur(f);
  if(a==='pa-pick'){PAB.cur[f.id]=b.dataset.id;renderFincas();return}
  if(a==='pa-new'){if(!fincaRecs(f.id).length){toast('La finca no tiene recintos');return} PAB.ask=PAB.ask===f.id?null:f.id; renderFincas(); return}
  if(a==='pa-new-no'){PAB.ask=null;renderFincas();return}
  if(a==='pa-new-ok'){var sy=$('#pa-ny'), y=sy?+sy.value:pabY0(); if(!(y>1900))return;
    if((f.pab||[]).some(function(x){return +x.y===y})){toast('Ya hay un plan de abonado de la campaña '+pabCamp(y));return}
    if(!fincaRecs(f.id).length){toast('La finca no tiene recintos');return}
    PAB.ask=null; undoPush(); var n={id:uid(),y:String(y),cl:Object.assign({},PAB_CL),ud:pabAuto(f),c:Date.now(),m:Date.now()};
    var prev=pabList(f)[0]; if(prev){/* copia lo que no cambia de un año a otro: suelo, recintos, zona vulnerable y agua */
      n.cl=Object.assign({},prev.cl||PAB_CL); n.ud.forEach(function(u){var o=(prev.ud||[]).filter(function(x){return x.n===u.n||(x.ck===u.ck&&x.rg===u.rg)})[0];if(o){['zv','tp','inh','rp'].forEach(function(k){if(o[k]!==undefined)u[k]=o[k]});u.su=JSON.parse(JSON.stringify(o.su||{}));u.rc=JSON.parse(JSON.stringify(o.rc||{}));u.ag=JSON.parse(JSON.stringify(o.ag||{}))}})}
    (f.pab||(f.pab=[])).push(n); PAB.cur[f.id]=n.id; save(); renderFincas(); undoToast('Plan de abonado '+pabCamp(y)+' creado con '+n.ud.length+(n.ud.length===1?' unidad':' unidades')); return}
  if(!p)return;
  if(a==='pa-del'){if(typeof confirm==='function'&&!confirm('¿Borrar el plan de abonado '+pabCamp(p.y)+'?'))return;undoPush();f.pab=(f.pab||[]).filter(function(x){return x.id!==p.id});if(!f.pab.length)delete f.pab;save();renderFincas();undoToast('Plan de abonado borrado');return}
  if(a==='pa-set'){pabChange(f,b.dataset.p,b.dataset.v,true);return}
  if(a==='pa-pdf'||a==='pa-xlsx'){var go=function(){pabOut(f,p,a==='pa-pdf'?'pdf':'xlsx')};if(a==='pa-pdf')Promise.resolve(loadPdfLib()).then(go,function(){pdfNoListo()});else go();return}
  if(a==='pa-mod-add'){(p.mods||(p.mods=[])).push({id:uid(),f:new Date().toISOString().slice(0,10),u:'',mo:'',ch:'',re:''});PAB.open.md=true;p.m=Date.now();save();renderFincas();return}
  if(a==='pa-mod-del'){undoPush();p.mods.splice(+b.dataset.j,1);p.m=Date.now();save();renderFincas();undoToast('Modificación quitada');return}
  if(a==='pa-ud-new'){undoPush();p.ud=p.ud||[];var u={id:uid(),n:'Unidad '+(p.ud.length+1),ck:'otro',rg:'secano',refs:[],zv:0,tp:0,inh:0,rc:{},su:{},ag:{},ne:{mo:'cos',ed:1},ap:{s:{},o:{},x:{}},pg:{src:'own',rows:[]}};p.ud.push(u);PAB.open[u.id]=true;save();renderFincas();undoToast('Unidad añadida: mueve a ella los recintos desde la tabla de recintos');return}
  var ui=+b.dataset.u, ud=(p.ud||[])[ui];
  if(a==='pa-ud-del'&&ud){undoPush();if(ud.refs.length&&p.ud.length>1){var to=p.ud[ui===0?1:0];to.refs=to.refs.concat(ud.refs)}if((ud.xr||[]).length&&p.ud.length>1){var to2=p.ud[ui===0?1:0];to2.xr=(to2.xr||[]).concat(ud.xr);to2.rc=to2.rc||{};ud.xr.forEach(function(x){if(ud.rc&&ud.rc[x])to2.rc[x]=ud.rc[x]})}p.ud.splice(ui,1);save();renderFincas();undoToast('Unidad borrada'+(ud.refs.length&&p.ud.length?' (sus recintos pasan a «'+p.ud[0].n+'»)':''));return}
  if(a==='pa-row-add'&&ud){ud.pg=ud.pg||{src:'own',rows:[]};(ud.pg.rows||(ud.pg.rows=[])).push({id:uid(),mo:'',n:'',fo:'Suelo'});p.m=Date.now();save();renderFincas();return}
  if(a==='pa-row-del'&&ud){undoPush();ud.pg.rows.splice(+b.dataset.j,1);p.m=Date.now();save();renderFincas();undoToast('Aporte quitado');return}
}
function pabMove(f,ui,ref,to){
  var p=pabCur(f); if(!p)return; var u=p.ud[ui]; if(!u)return; undoPush();
  if(to==='new'){var r=S.recintos[ref], c=(r&&r.cultivo)||NOCROP, g=r?pabRg(r):'secano'; p.ud.push({id:uid(),n:c+' '+(g==='regadio'?'regadío':'secano')+' ('+(p.ud.length+1)+')',ck:pabCk(c),rg:g,refs:[],zv:u.zv,tp:0,inh:0,rc:{},su:{},ag:{},ne:{mo:'cos',ed:1},ap:{s:{},o:{},x:{}},pg:{src:'own',rows:[]}});to=p.ud.length-1}
  var xo=pabIsX(f,u,ref);
  if(to==='out'){u.xr=(u.xr||[]).filter(function(x){return x!==ref});if(u.rc)delete u.rc[ref];if(u.rp===ref)u.rp=u.refs[0]||(u.xr||[])[0]||'';p.m=Date.now();save();renderFincas();undoToast('Recinto quitado de este plan');return}
  var d=p.ud[+to]; if(!d||d===u){renderFincas();return}
  if(xo){u.xr=(u.xr||[]).filter(function(x){return x!==ref});(d.xr||(d.xr=[])).push(ref)}
  else{u.refs=u.refs.filter(function(x){return x!==ref}); d.refs.push(ref)} if(u.rc&&u.rc[ref]){d.rc=d.rc||{};d.rc[ref]=u.rc[ref];delete u.rc[ref]}
  if(u.rp===ref)u.rp=u.refs[0]||''; if(!d.rp)d.rp=ref;
  p.m=Date.now(); save(); renderFincas(); undoToast('Recinto movido a «'+d.n+'»');
}
/* recintos que se han añadido a la finca después de crear el plan: van a la unidad de su cultivo y riego */
function pabSync(f){
  (f.pab||[]).forEach(function(p){var inU={};(p.ud||[]).forEach(function(u){if(u.pg&&u.pg.src==='plan'){u.pg.src='own';u.pg.rows=u.pg.rows||[];delete u.pg.pid}u.refs=(u.refs||[]).filter(function(x){var r=S.recintos[x];return r&&r.fincaId===f.id});
      if(u.xr){u.xr=u.xr.filter(function(x){var r=S.recintos[x];if(r&&r.fincaId===f.id){if(u.refs.indexOf(x)<0)u.refs.push(x);return false}return !!r});if(!u.xr.length)delete u.xr}
      u.refs.forEach(function(x){inU[x]=1})});
    fincaRecs(f.id).forEach(function(r){if(inU[r.ref])return;var u=(p.ud||[]).filter(function(x){return x.ck===pabCk(r.cultivo)&&x.rg===pabRg(r)})[0]||(p.ud||[])[0];if(u)u.refs.push(r.ref)})});
}
document.addEventListener('change',function(ev){
  var t=ev.target; if(!t||!t.classList)return; var f=finca(EXP.fid); if(!f||GTOK)return;
  if(t.classList.contains('pa-mv')){pabMove(f,+t.dataset.u,t.dataset.ref,t.value);return}
  if(t.classList.contains('pa-add')){if(t.value)pabAddX(f,+t.dataset.u,t.value);return}
  if(t.classList.contains('pa-in')&&t.dataset.p){pabChange(f,t.dataset.p,t.value,true)}
});
document.addEventListener('toggle',function(ev){var t=ev.target;if(t&&t.dataset&&t.dataset.paO)PAB.open[t.dataset.paO]=t.open},true);
/* ───────── Entrega 79: plan de abonado completo (art. 4, 5, 6 y anexo V del RD 1051/2022; Decreto 80/2025 de CLM) ───────── */
/* anexo V del RD 1051/2022: parte A (estiércoles y materiales orgánicos) y parte B (urea y soluciones ureicas), más buenas prácticas */
var PAB_MIT=[
  ['B1','B','Ajustar la dosis y el momento a la fase en que el cultivo aprovecha antes el nutriente','Se ajusta la dosis al balance y se aplica cuando el cultivo absorbe el N.'],
  ['B2','B','Incorporar el fertilizante al suelo al aplicarlo o enterrarlo en las 4 horas siguientes',''],
  ['B3','B','Urea en gránulos recubiertos de polímero',''],
  ['B4','B','Riego inmediatamente después de abonar, o aplicación en fertirrigación','Se aplica con humedad del suelo, antes de lluvia útil o con un riego que lo incorpore.'],
  ['B5','B','Inhibidores de la ureasa',''],
  ['B6','B','Dosificación variable dentro de la parcela',''],
  ['A1','A','Aplicación en bandas con mangueras o tubos rígidos (purines)',''],
  ['A2','A','Aplicación en bandas con discos o rejas',''],
  ['A3','A','Inyección',''],
  ['A4','A','Enterrado de purines y materiales líquidos en las primeras 4 horas',''],
  ['A5','A','Dilución de los purines mediante el riego',''],
  ['A6','A','Purines acidificados o con inhibidores de la ureasa',''],
  ['A7','A','Inhibidores de la ureasa o de la nitrificación, con supervisión profesional',''],
  ['A8','A','Compost o digerido con N amoniacal inferior al 0,6 % sobre peso fresco',''],
  ['A9','A','Enterrado de estiércoles en las primeras 12 horas (salvo siembra directa, conservación o pastos)',''],
  ['O1','O','Evitar aplicar con viento fuerte o en las horas de más calor','No se aplica con viento fuerte ni en las horas de mayor temperatura.'],
  ['O2','O','Cubierta vegetal o reducción del laboreo',''],
  ['O3','O','Otras medidas','']];
var PAB_MITG={A:'Anexo V, parte A (estiércoles, purines y materiales orgánicos)',B:'Anexo V, parte B (urea y soluciones ureicas)',O:'Otras buenas prácticas'};
/* restricciones (art. 4.4 del RD 1051/2022 y apartado 4.3 del Decreto 80/2025) */
var PAB_RES=[
  ['r1','No se aplica sobre suelo helado, nevado, hidromorfo, encharcado o inundado (art. 4.4 a)','No se aplica sobre suelo helado, nevado, hidromorfo, encharcado ni inundado.'],
  ['r2','No se aplica con aviso rojo por lluvia de AEMET o Protección Civil (art. 4.4 b)','Se consulta el aviso de AEMET antes de abonar.'],
  ['r3','No se aplica donde pueda haber arrastre a humedales, barrancos o saladares (art. 4.4 c)','Ningún recinto vierte a humedales, barrancos ni saladares.'],
  ['r4','No se aplica con previsión de lluvia de más de 30 mm/h (zona vulnerable, Decreto 80/2025)','Se evita abonar con previsión de lluvia intensa.'],
  ['r5','No se aplica dentro de perímetros de protección del plan hidrológico, de aguas minerales o de captaciones de consumo humano (apdo. 4.3.3 del Decreto 80/2025)','Ningún recinto está dentro de un perímetro de protección.'],
  ['r6','Abonadoras y aperos calibrados y en buen estado (art. 4.6)','La abonadora se revisa y calibra antes de la campaña.']];
/* anexo II del RD 1051/2022: meses sin abono nitrogenado (salvo riego localizado o agricultura de precisión) */
var PAB_PER={almendro:[11,12,1],pistacho:[11,12,1],nogal:[11,12,1],olivo:[11,12,1],vid:[11,12,1,2],trigo:[6,7,8,9],trigoduro:[6,7,8,9],cebada:[6,7,8,9],avena:[6,7,8,9],centeno:[6,7,8,9],triticale:[6,7,8,9]};
var PAB_MES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
function pabPerTxt(ck){var m=PAB_PER[ck];if(m)return m.map(function(x){return PAB_MES[x-1]}).join(', ').replace(/, ([^,]*)$/,' y $1');if(ck==='frutal'||ck==='melocotonero')return 'de la caída de la hoja al inicio de la brotación';return ''}
function pabCampRange(p){var y=+p.y;return [y+'-11-01',(y+1)+'-10-31']}
function pabIso(s){s=String(s||'').trim();var m=/^(\d{4})-(\d\d)-(\d\d)$/.exec(s);if(m)return s;m=/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/.exec(s);if(m){var y=+m[3];if(y<100)y+=2000;return y+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2)}m=/^(\d{4})$/.exec(s);return m?s+'-06-30':''}
function pabDays(a,b){return Math.round((new Date(b+'T12:00:00')-new Date(a+'T12:00:00'))/864e5)}
function pabAs(){return (S.x&&S.x.asesor)||{}}
/* lo aplicado de verdad (Abonado y Fertirrigación hechos) en la campaña y en los recintos de la unidad */
/* Entrega 82: el plan de abonado es independiente de lo aplicado en la herramienta (no lo lee) */
function pabReal(f,u,p){return {rows:[],tot:{n:0,p:0,k:0}}}
/* aplicaciones de N en meses prohibidos (anexo II) */
var PAB_MES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
function pabMeses(s){s=String(s||'').toLowerCase();var o=[];PAB_MES.forEach(function(m,i){if(s.indexOf(m)>=0||new RegExp('(^|[^a-zñáéíóú])'+m.slice(0,3)+'([^a-zñáéíóú]|$)').test(s))o.push(i+1)});return o}
/* aportes de N del programa previsto que caen en los meses prohibidos (anexo II); solo se comprueban los que dicen el mes */
function pabPerHits(f,u,p,C){
  var m=PAB_PER[u.ck]; if(!m)return [];
  if(u.rg==='regadio'&&u.loc!==0)return [];
  var hits=[];
  (C.rows||[]).forEach(function(o){if(!(o.uf&&o.uf.n>0.05))return;var ms=pabMeses(o.mo);if(ms.some(function(x){return m.indexOf(x)>=0}))hits.push((o.mo||'')+' · '+(o.n||''))});
  return hits;
}
function pabPerSinMes(u,C){return (C.rows||[]).filter(function(o){return o.uf&&o.uf.n>0.05&&!pabMeses(o.mo).length}).length}
function pabNorg(C){return C.rows.reduce(function(a,o){return a+(o.org?o.uf.n:0)},0)}
function pabSoilAge(u){var f=pabIso((u.su||{}).f);if(!f)return null;return Math.floor(pabDays(f,new Date().toISOString().slice(0,10))/365.25)}
/* comprobación del contenido del plan */
function pabCheck(f,p,u,C,RL){
  var su=u.su||{}, pg=u.pg||{}, ne=u.ne||{}, mit=u.mit||{}, res=u.res||{}, A=pabAs(), L=[];
  var ok=function(c,lbl,why,na){L.push({c:c,l:lbl,w:why,na:!!na})}, d0=u.d0===1;
  ok(C.rs.length>0,'6.1 Recintos identificados',C.rs.length+(C.rs.length===1?' recinto':' recintos'));
  var hasSu=['mo','n','p','k','ph'].some(function(k){return abHas(su[k])}), age=pabSoilAge(u), mx=u.rg==='regadio'?5:10;
  ok(d0||(hasSu&&!!u.rp),'6.2 Datos de suelo del recinto representativo',d0&&!hasSu?'sin abonado esta campaña (decisión justificada): no se hace análisis':hasSu?((age!==null&&age>mx)?'el análisis tiene '+age+' años (más de '+mx+')':'materia orgánica y nutrientes'):'faltan los datos del análisis',d0&&!hasSu);
  var cl=p.cl||PAB_CL; ok(abHas(cl.a)&&(u.rg!=='regadio'||abHas((u.ag||{}).m3)||abHas((u.ag||{}).max)),'6.3 Precipitación'+(u.rg==='regadio'?' y agua de riego':''),u.rg==='regadio'&&!abHas((u.ag||{}).m3)&&!abHas((u.ag||{}).max)?'falta la dotación de riego':'');
  var rows=C.rows, rowsOk=rows.length&&(rows.every(function(o){return o.n&&o.d>0&&o.mo&&o.fo}));
  ok(d0||!!rowsOk,'6.4 Dosis, momento, tipo de abono, forma de aplicación y maquinaria',d0?'no se prevé fertilizar esta campaña (dosis 0)':rows.length?(rowsOk?'':'falta el momento, la dosis o la forma en algún aporte'):'no hay programa',d0);
  var anyM=PAB_MIT.some(function(x){return (mit[x[0]]||{}).a===1});
  ok(d0||anyM,'6.5 Medidas de reducción de amoniaco y GEI (anexo V)',anyM?'':(d0?'sin aplicaciones previstas':'marca al menos una medida'),d0&&!anyM);
  ok(!!(u.mo&&u.mo.obj),'4.3 Materia orgánica: mantener o aumentar',(u.mo&&u.mo.obj)?'':'elige el objetivo');
  var allR=PAB_RES.every(function(x){return (res[x[0]]||{}).a===1});
  ok(d0||allR,'4.4 Restricciones de aplicación comprobadas',allR?'':(d0?'sin aplicaciones previstas':'faltan restricciones por comprobar'),d0&&!allR);
  ok(abHas(ne.pr)||ne.mo==='edad'||ne.mo==='man','5.1 a) Rendimiento esperado y necesidades','');
  ok(!!u.prec,'5.1 a) Cultivo precedente',u.prec?'':'falta');
  ok(!!pabIso(p.fe),'5.1 a) Fecha de elaboración del plan',pabIso(p.fe)?'':'falta');
  var ob=u.zv?'2027-09-01':'2028-09-01', need=pabCampRange(p)[1]>=ob, hasA=!!(A.n&&A.regfer);
  ok(hasA||!need,'6.6 Asesoramiento técnico',hasA?'asesor: '+A.n:(need?'obligatorio desde el '+diaFmt(ob):'obligatorio desde el '+diaFmt(ob)+' (aún no)'));
  if(u.zv){ok(C.lok,'Zona vulnerable: límite de N (anejo III)',C.L.ok?'límite '+nf(C.L.lim,0)+' kg N/ha':(d0?'sin abonado esta campaña':'elige el cultivo'),d0&&!C.L.ok);var no=pabNorg(C);ok(no<=170,'Zona vulnerable: N orgánico ≤ 170 kg/ha',trTxt(no,1)+' kg N/ha de abonos orgánicos')}
  if(!d0){var SL=pabSlope(u,C), cub=!!((u.mo||{}).cub)||(mit.O2||{}).a===1, ex=SL.x15.concat(SL.b8);
    ok(!(u.zv&&SL.pmax>2&&!cub),'Pendientes (apdo. 4.3.2 del Decreto 80/2025)',SL.pmax<0?'sin dato de pendiente':(ex.length?'no se abonan: '+pabSlopeTxt(ex)+'. ':'')+(SL.pmax>2&&u.zv?(cub?'más del 2 %: con medida de precaución (cubierta vegetal)':'más del 2 %: falta la medida de precaución (cubierta vegetal, franjas…)'):(ex.length?'':'ningún recinto con restricciones')));
    var PP=pabPerim(u,C); ok(!PP.length,'Perímetros de protección (apdo. 4.3.3 del Decreto 80/2025)',PP.length?'dentro de un perímetro: '+PP.map(function(r){return refLabel(r.ref)}).join(', '):'ningún recinto dentro de un perímetro de protección')}
  var ph=pabPerHits(f,u,p,C,RL); ok(!ph.length,'Anexo II: sin N en los meses prohibidos'+(pabPerTxt(u.ck)?' ('+pabPerTxt(u.ck)+')':''),ph.length?ph.join(' · '):(u.rg==='regadio'&&u.loc!==0?'riego localizado: no se aplica':(PAB_PER[u.ck]&&pabPerSinMes(u,C)?'según el programa previsto; pon el mes en el momento de cada aporte para poder comprobarlo':'según el programa previsto')));
  return L;
}
function pabChkHTML(L){return '<div class="xs" style="overflow-x:auto"><table class="patb"><tbody>'+L.map(function(x){return '<tr><td>'+esc(x.l)+(x.w?'<div class="small">'+esc(x.w)+'</div>':'')+'</td><td><span class="pa-st pa-'+(x.na||x.c?'ok':'bad')+'">'+(x.na?'No procede':(x.c?'Completo':'Pendiente'))+'</span></td></tr>'}).join('')+'</tbody></table></div>'}
/* secciones nuevas de cada unidad */
function pabUdMore(f,p,u,C,i){
  var P='ud.'+i+'.', h='', mo=u.mo||{}, mit=u.mit||{}, res=u.res||{}, RL=pabReal(f,u,p);
  /* inicio de campaña */
  h+='<h4 style="margin:14px 0 4px">Datos de inicio de campaña <span class="small">(art. 5.1 a)</span></h4><div class="pa-g"><div>'+pabF('Cultivo precedente',pabTx(P+'prec',u.prec,'Ej. Pistacho'))+pabF('Rotación de cultivos (herbáceos)',pabTx(P+'rot',u.rot,'Ej. Cereal - barbecho'))+'</div><div>'+pabF('¿El precedente es una leguminosa?',pabChips(P+'leg',u.leg?1:0,[[1,'Sí'],[0,'No']]))+'</div></div>';
  if(u.rg==='regadio')h+=pabF('¿Riego localizado (goteo)?',pabChips(P+'loc',u.loc===0?0:1,[[1,'Sí'],[0,'No']])+'<div class="small">Con riego localizado no se aplican los meses de prohibición del anexo II.</div>');
  /* materia orgánica */
  h+='<h4 style="margin:14px 0 4px">Materia orgánica <span class="small">(art. 4.3)</span></h4><div class="small">MO actual del análisis: <b>'+(abHas((u.su||{}).mo)?trTxt(+u.su.mo,2)+' %':'sin dato')+'</b></div>'+pabF('Objetivo',pabChips(P+'mo.obj',mo.obj||'',[['mantener','Mantener'],['aumentar','Aumentar']]))+'<div class="pa-g"><div>'+pabF('Enmienda o abono orgánico previsto',pabTx(P+'mo.enm',mo.enm,'Ej. Ecofem (pellet orgánico)'))+'</div><div>'+pabF('Dosis (t/ha)',pabIn(P+'mo.dose',mo.dose))+'</div></div>'+pabF('Cubierta vegetal y manejo del suelo',pabTx(P+'mo.cub',mo.cub,'Ej. Cubierta vegetal espontánea en las calles (ecorrégimen PAC)'));
  var no=pabNorg(C); h+='<div class="small" style="margin-top:4px">N de abonos orgánicos en el programa: <b>'+trTxt(no,1)+' kg N/ha</b>'+(u.zv?' (máximo 170 en zona vulnerable'+(no>170?': <b>lo supera</b>':'')+')':'')+'. Marca «Orgánico» en los aportes del programa que lo sean.</div>'+pabOrgHTML(u,C);
  /* medidas */
  h+='<h4 style="margin:14px 0 4px">Medidas para reducir amoniaco y gases de efecto invernadero <span class="small">(art. 6.5 y anexo V)</span></h4>';
  ['B','A','O'].forEach(function(g){h+='<div class="lbl" style="margin-top:8px">'+PAB_MITG[g]+'</div>';PAB_MIT.filter(function(x){return x[1]===g}).forEach(function(x){var m=mit[x[0]]||{};h+='<div class="pa-row"><div class="pa-rl">'+esc(x[2])+'</div>'+pabChips(P+'mit.'+x[0]+'.a',m.a===1?1:(m.a===0?0:''),[[1,'Sí'],[0,'No']])+(m.a===1?pabTx(P+'mit.'+x[0]+'.d',m.d===undefined?x[3]:m.d,'Cómo se aplica'):'')+'</div>'})});
  var urea=C.rows.some(function(o){return /urea/i.test(o.n)}), org=C.rows.some(function(o){return o.org}), hasB=PAB_MIT.some(function(x){return x[1]==='B'&&(mit[x[0]]||{}).a===1}), hasA=PAB_MIT.some(function(x){return x[1]==='A'&&(mit[x[0]]||{}).a===1});
  if(urea&&!hasB)h+='<div class="note">El programa lleva urea: hace falta al menos una medida de la parte B (art. 12.3).</div>';
  if(org&&!hasA)h+='<div class="small" style="margin-top:4px">Si aplicas estiércoles, purines u otros materiales orgánicos sin compostar, hace falta al menos una medida de la parte A (art. 10.4).</div>';
  /* restricciones */
  var pmax=C.rs.reduce(function(a,r){var x=((u.rc||{})[r.ref]||{}).pe;return abHas(x)?Math.max(a,+x):a},-1);
  h+='<h4 style="margin:14px 0 4px">Restricciones y condiciones de aplicación <span class="small">(art. 4.4 y Decreto 80/2025)</span></h4>';
  PAB_RES.forEach(function(x){var m=res[x[0]]||{};h+='<div class="pa-row"><div class="pa-rl">'+esc(x[1])+'</div>'+pabChips(P+'res.'+x[0]+'.a',m.a===1?1:(m.a===0?0:''),[[1,'Comprobado'],[0,'No']])+(m.a!==undefined?pabTx(P+'res.'+x[0]+'.d',m.d===undefined?x[2]:m.d,'Observación'):'')+'</div>'});
  h+='<div class="small" style="margin-top:6px">Pendiente máxima de los recintos: <b>'+(pmax<0?'sin dato':trTxt(pmax,1)+' %')+'</b>'+(pmax>15?' · <b>más del 15 %: prohibido aplicar estiércoles y abonos N y P en zona vulnerable</b>':(pmax>2&&u.zv?' · más del 2 %: tomar precauciones (curvas de nivel, franjas sin abonar…)':''))+'.</div>';
  var pt=pabPerTxt(u.ck); if(pt)h+='<div class="small">Sin abono nitrogenado en '+pt+' (anexo II)'+(u.rg==='regadio'&&u.loc!==0?': no se aplica con riego localizado':'')+'.</div>';
  var ph=pabPerHits(f,u,p,C,RL); if(ph.length)h+='<div class="note">Aplicaciones de N en meses prohibidos: '+esc(ph.join(' · '))+'</div>';
  /* comprobación */
  h+='<h4 style="margin:14px 0 4px">Comprobación del plan</h4>'+pabChkHTML(pabCheck(f,p,u,C,RL));
  return h;
}
/* elaboración, asesor, firmas y modificaciones (a nivel de plan) */
function pabPlanMore(f,p){
  var A=pabAs(), t=((f.datos||{}).tit||{}), h='';
  h+='<details class="more pa-box"'+(PAB.open.fi?' open':'')+' data-pa-o="fi"><summary><b>Elaboración, asesor y firmas</b> <span class="small">(art. 5.1 a, 6.6 y 21)</span></summary><div class="pa-g">';
  h+='<div>'+pabF('Fecha de elaboración del plan','<input type="date" class="in pa-in" data-p="fe" value="'+esc(pabIso(p.fe))+'">')+'</div><div>'+pabF('Fecha del asesoramiento','<input type="date" class="in pa-in" data-p="fa" value="'+esc(pabIso(p.fa))+'">')+'</div>';
  h+='<div>'+pabF('Titular que firma',pabTx('ft',p.ft,t.rn||t.n||'Nombre del titular o representante'))+(t.rn?'<div class="small">Si lo dejas en blanco firma '+esc(t.rn)+' como representante (se cambia en Datos → Titular y contacto).</div>':'')+'</div></div>';
  h+='<div class="lbl" style="margin-top:10px">Asesor en fertilización <span class="small">(se guarda una vez para todas las fincas)</span></div><div class="pa-g">'+[['n','Nombre y apellidos'],['nif','NIF'],['tit','Titulación'],['regfer','Nº de inscripción en el REGFER (sección asesores)'],['acr','Acreditación ante la comunidad autónoma']].map(function(x){return '<div>'+pabF(x[1],pabTx('AS.'+x[0],A[x[0]]))+'</div>'}).join('')+'</div>';
  h+='<div class="small" style="margin-top:6px">El asesoramiento es obligatorio desde el 1 de septiembre de 2027 en zona vulnerable y desde el 1 de septiembre de 2028 en el resto (art. 6.6). Las unidades exentas de plan no lo necesitan.</div></details>';
  h+='<details class="more pa-box"'+(PAB.open.md?' open':'')+' data-pa-o="md"><summary><b>Modificaciones durante la campaña</b> <span class="small">('+(p.mods||[]).length+')</span></summary>';
  if((p.mods||[]).length)h+='<div style="overflow-x:auto"><table class="patb pa-prog stk"><thead><tr><th>Fecha</th><th>Unidad / recintos</th><th>Motivo (evolución del cultivo, clima…)</th><th>Cambio realizado</th><th>Responsable</th><th></th></tr></thead><tbody>'+p.mods.map(function(m,j){var R='mods.'+j+'.';return '<tr><td><input type="date" class="in pa-in" data-p="'+R+'f" value="'+esc(pabIso(m.f))+'"></td><td>'+pabTx(R+'u',m.u)+'</td><td>'+pabTx(R+'mo',m.mo)+'</td><td>'+pabTx(R+'ch',m.ch)+'</td><td>'+pabTx(R+'re',m.re)+'</td><td><button type="button" class="btn sm danger" data-act="pa-mod-del" data-j="'+j+'" aria-label="Quitar">✕</button></td></tr>'}).join('')+'</tbody></table></div>';
  h+='<div class="acts" style="margin:6px 0 0"><button class="btn sm" data-act="pa-mod-add">+ Añadir modificación</button></div><div class="small">El plan se puede cambiar durante la campaña según evolucione el cultivo o el tiempo; apunta aquí cada cambio.</div></details>';
  return h;
}
/* ───── contenido del PDF y del Excel (los dos salen de la misma lista de bloques) ───── */
function pabN(v,d){return {n:+v||0,d:d==null?1:d}}
function pabBlocks(f,p){
  var B=[], A=pabAs(), t=((f.datos||{}).tit||{}), cl=p.cl||PAB_CL, Cs=(p.ud||[]).map(function(u){return pabCalc(f,u,p)});
  var H=function(x){B.push({t:'h',x:x})}, Pp=function(x,o){B.push({t:'p',x:x,o:o||{}})}, K=function(k,v){if(v!==''&&v!==undefined&&v!==null)B.push({t:'kv',k:k,v:String(v)})}, T=function(cols,rows,tot){B.push({t:'tb',cols:cols,rows:rows,tot:tot})};
  H('Identificación de la explotación');
  K('Titular',t.n||'');K('NIF/CIF',t.nif||'');K('Representante',daRepTxt(t));K('Dirección fiscal',[t.dir,[t.cp,t.mun].filter(Boolean).join(' '),t.prov].filter(Boolean).join(', '));K('Finca',f.nombre);K('Campaña',pabCamp(p.y));K('Superficie',haTxt(fincaRecs(f.id)));K('Fecha de elaboración del plan',pabIso(p.fe)?diaFmt(pabIso(p.fe)):'');
  H('Resumen por unidad de producción');
  T([{h:'Unidad',w:52},{h:'Régimen',w:20},{h:'ha',w:16,a:'r'},{h:'Producción objetivo',w:28,a:'r'},{h:'N',w:14,a:'r'},{h:'P₂O₅',w:14,a:'r'},{h:'K₂O',w:14,a:'r'},{h:'Estado',w:24}],(p.ud||[]).map(function(u,i){var C=Cs[i],ne=u.ne||{};return [u.n,u.rg==='regadio'?'Regadío':'Secano',pabN(C.ha,2),ne.mo==='edad'?'en formación':(abHas(ne.pr)&&+ne.pr>0?nf(+ne.pr,0)+' kg/ha':'—'),pabN(C.pl.n),pabN(C.pl.p),pabN(C.pl.k),pabStTxt(f,u,C)]}));
  Pp('U.F./ha planificadas.',{sz:8,gray:1});
  H('Precipitación de la zona (art. 6.3)');
  T([{h:'Media anual',w:36,a:'r'},{h:'Otoño (sep-nov)',w:36,a:'r'},{h:'Invierno (dic-feb)',w:36,a:'r'},{h:'Primavera (mar-may)',w:37,a:'r'},{h:'Verano (jun-ago)',w:37,a:'r'}],[[pabN(cl.a,0),pabN(cl.o,0),pabN(cl.i,0),pabN(cl.p,0),pabN(cl.v,0)]]);
  Pp('mm. Fuente: '+(cl.src||''),{sz:8,gray:1});
  (p.ud||[]).forEach(function(u,i){
    var C=Cs[i], RL=pabReal(f,u,p), su=u.su||{}, ne=u.ne||{}, mo=u.mo||{}, mit=u.mit||{}, res=u.res||{};
    B.push({t:'np'}); B.push({t:'h1',x:'Unidad de producción: '+u.n});
    Pp(ha2(C.ha)+(Math.abs(C.ha-C.sig)>0.00005?' ha cultivadas (SIGPAC '+ha2(C.sig)+' ha)':' ha')+' · '+C.rs.length+(C.rs.length===1?' recinto':' recintos')+' · '+pabCkName(u.ck)+' en '+(u.rg==='regadio'?'regadío':'secano')+(u.zv?' · zona vulnerable':' · fuera de zona vulnerable'),{sz:10});
    if(C.ex)Pp('Unidad de secano de 10 ha o menos: no está obligada a tener plan de abonado (art. 4.2 b del RD 1051/2022).',{b:1,sz:9});
    H('Datos de inicio de campaña (art. 5.1 a)');
    K('Cultivo precedente',(u.prec||'—')+(u.leg?' (leguminosa)':''));if(u.rot)K('Rotación de cultivos',u.rot);K('Rendimiento esperado',ne.mo==='edad'?'Plantación en formación ('+(+ne.ed||1)+'º verde)':(abHas(ne.pr)?(+ne.pr===0?'Sin cosecha prevista':nf(+ne.pr,0)+' kg/ha'):'—'));
    var EX0=(ne.mo||'cos')==='cos'?(ne.ex||(PAB_EX[u.ck]?{n:PAB_EX[u.ck][0][0],p:PAB_EX[u.ck][0][1],k:PAB_EX[u.ck][0][2]}:null)):null;
    if(EX0&&abHas(ne.pr)){K('Extracción por t de cosecha','N '+trTxt(+EX0.n||0,1)+' · P₂O₅ '+trTxt(+EX0.p||0,1)+' · K₂O '+trTxt(+EX0.k||0,1)+' kg/t');K('Extracción total del cultivo',nf(+ne.pr,0)+' kg/ha × extracción por t = '+trTxt(C.nb.n,1)+' / '+trTxt(C.nb.p,1)+' / '+trTxt(C.nb.k,1)+' kg/ha')}
    K('Necesidades','N '+trTxt(C.nb.n,1)+' · P₂O₅ '+trTxt(C.nb.p,1)+' · K₂O '+trTxt(C.nb.k,1)+' kg/ha · '+C.src);
    if(u.rg==='regadio')K('Riego localizado',u.loc===0?'No':'Sí');
    H('Recintos (art. 6.1)');
    /* Entrega 144: sin la columna «Nombre» si ningún recinto lo tiene; total con 2 decimales */
    var hNm=C.rs.some(function(r){return r.nombre||pabIsX(f,u,r.ref)}), cNm=function(a){if(!hNm)a.splice(1,1);return a};
    T(cNm([{h:'Referencia SIGPAC',w:36},{h:'Nombre',w:22},{h:'ha cultiv.',w:16,a:'r'},{h:'ha en el plan',w:18,a:'r'},{h:'Cultivo',w:24},{h:'Pendiente %',w:18,a:'r'},{h:'Perím. protección',w:20},{h:'Repr.',w:14}]).map(function(c,i,A){if(!hNm&&c.h==='Cultivo')c.w+=22;return c}),C.rs.map(function(r){var rc=(u.rc||{})[r.ref]||{}, xo=pabIsX(f,u,r.ref), of=xo?finca(r.fincaId):null;return [refLabel(r.ref),(r.nombre||(xo?'':'-'))+(xo?(r.nombre?' ':'')+'(finca «'+(of?of.nombre:'')+'»)':''),pabN(haC(r),4),pabN(pabHaR(u,r),4),r.cultivo||'-',abHas(rc.pe)?pabN(rc.pe,1):'',rc.pp===1?'Sí':(rc.pp===0?'No':'—'),u.rp===r.ref?'Sí':''].filter(function(_,i){return hNm||i!==1})}),cNm(['Total','','',ha2(C.ha),'','','','']));
    C.rs.forEach(function(r){var oth=pabShares(r.ref,p.y).filter(function(s){return s.f.id!==f.id});if(oth.length)Pp('Recinto repartido: '+refLabel(r.ref)+' ('+ha2(r.ha||0)+' ha en el SIGPAC) entra en este plan con '+ha2(pabHaR(u,r))+' ha; el resto figura en el plan de '+oth.map(function(s){return '«'+s.f.nombre+'» ('+ha2(s.ha)+' ha)'}).join(' y ')+'.',{sz:9})});
    H('Datos de suelo del recinto representativo (art. 6.2)');
    K('Recinto representativo',u.rp?refLabel(u.rp):'—');K('Muestreo',[pabIso(su.f)?diaFmt(pabIso(su.f)):su.f,su.lab,su.inf?(/^(informe|an[aá]lisis)/i.test(String(su.inf))?su.inf:'informe '+su.inf):'',abHas(su.prof)?su.prof+' cm':''].filter(Boolean).join(' · '));
    K('Resultados',[abHas(su.mo)?'MO '+trTxt(+su.mo,2)+' %':'',abHas(su.ph)?'pH '+trTxt(+su.ph,1):'',abHas(su.n)?'N '+trTxt(+su.n,0)+' mg/kg':'',abHas(su.p)?'P Olsen '+trTxt(+su.p,0)+' mg/kg':'',abHas(su.k)?'K '+trTxt(+su.k,0)+' mg/kg':'',su.tex?'textura '+daLbl(PAB_TEX,su.tex).toLowerCase():''].filter(Boolean).join(' · ')||'—');
    K('Contaminantes',su.cont||'No analizados');
    if(u.rg==='regadio'){var ag=u.ag||{};H('Agua de riego (art. 6.3)');K('Dotación prevista',abHas(ag.m3)?nf(+ag.m3,0)+' m³/ha':'—');K('Dotación máxima concedida',abHas(ag.max)?nf(+ag.max,0)+' m³/ha·año':'');K('Nitratos del agua',abHas(ag.no3)?trTxt(+ag.no3,1)+' mg NO₃/L':'—');K('N aportado por el agua',trTxt(C.nAg,1)+' kg N/ha');K('Análisis',ag.src||'')}
    H('Balance de nutrientes (U.F./ha)');
    T([{h:'Nutriente',w:20},{h:'Necesidad bruta',w:20,a:'r'},{h:'Suelo',w:16,a:'r'},{h:'Agua',w:16,a:'r'},{h:'Orgánico',w:18,a:'r'},{h:'Otros',w:16,a:'r'},{h:'Neta',w:18,a:'r'},{h:'Planificado',w:20,a:'r'},{h:'Desviación',w:20,a:'r'},{h:'Estado',w:18}],PAB_NU.map(function(x){var b=C.B[x[0]];return [x[1],pabN(b.nb),pabN(b.s),pabN(b.a),pabN(b.o),pabN(b.x),pabN(b.net),pabN(b.pl),(b.dv>0?'+':'')+trTxt(b.dv,1),b.st]}));
    if(u.zv&&C.L.ok)Pp('Límite de N en zona vulnerable: '+nf(C.L.lim,0)+' kg N/ha·año · planificado '+trTxt(C.pl.n,1)+' · '+C.lst+'. '+C.L.txt,{sz:8.5});else Pp(C.lst+'.',{sz:8.5,gray:1});if(u.bobs)Pp('Observaciones del balance: '+u.bobs,{sz:8.5});
    H('Programa previsto de fertilización (art. 6.4)');
    if(C.rows.length)T([{h:'Momento',w:22},{h:'Producto',w:34},{h:'Dosis por ha',w:16,a:'r'},{h:'Riqueza',w:17},{h:'U.F. N-P-K',w:22,a:'r'},{h:'Forma',w:16},{h:'Equipo',w:18},{h:'Mitigación',w:37}],C.rows.map(function(o){return [o.mo||'',o.n+(o.org?' (orgánico)':''),trTxt(o.d,1)+' '+(o.u||'kg'),String(o.r||'').replace(/,0(?=[^\d]|$)/g,''),trTxt(o.uf.n,1)+'-'+trTxt(o.uf.p,1)+'-'+trTxt(o.uf.k,1),o.fo||'',o.eq||'',o.mi||'']}),['Total','','','',trTxt(C.pl.n,1)+'-'+trTxt(C.pl.p,1)+'-'+trTxt(C.pl.k,1),'','','']);
    else Pp(u.d0===1?'No se prevé fertilizar esta campaña: programa a dosis 0. Si se decide abonar, se modificará el plan.':'Sin aportes previstos.',{sz:9,gray:1});
    Pp('Las dosis son previsiones: el plan se puede modificar durante la campaña según la evolución del cultivo y del tiempo.',{sz:8,gray:1});
    H('Materia orgánica (art. 4.3)');
    K('Materia orgánica actual',abHas(su.mo)?trTxt(+su.mo,2)+' %':'sin dato');K('Objetivo',mo.obj==='aumentar'?'Aumentar el contenido de materia orgánica':(mo.obj==='mantener'?'Mantener, al menos, el contenido de materia orgánica':'—'));K('Enmienda o abono orgánico previsto',(mo.enm||'—')+(abHas(mo.dose)?' · '+trTxt(+mo.dose,2)+' t/ha':''));if(mo.cub)K('Cubierta vegetal y manejo del suelo',mo.cub);K('N de abonos orgánicos',trTxt(pabNorg(C),1)+' kg N/ha'+(u.zv?' (máximo 170 kg N/ha en zona vulnerable)':''));
    var OG=pabOrg(u,C); if(OG.length){H('Caracterización del abono orgánico (riqueza garantizada en etiqueta)');
      T([{h:'Abono',w:52},{h:'Riqueza N-P₂O₅-K₂O (%)',w:26},{h:'Contenido (kg/t)',w:24},{h:'Dosis prevista (t/ha)',w:24,a:'r'},{h:'N aportado (kg/ha)',w:24,a:'r'},{h:'Dosis máxima por N (t/ha)',w:26,a:'r'}],OG.map(function(x){return [x.n,trTxt(x.rn,1)+'-'+trTxt(x.rp,1)+'-'+trTxt(x.rk,1),trTxt(x.kn,0)+'-'+trTxt(x.kp,0)+'-'+trTxt(x.kk,0),trTxt(x.d,2),trTxt(x.nN,1),x.max===null?'—':trTxt(x.max,2)]}));
      Pp('Riqueza garantizada en la etiqueta del producto; contenido en kg/t = % × 10. Dosis máxima: la que llega al límite de N de '+nf(OG[0].lim,0)+' kg/ha (anejo III y máximo de 170 kg N/ha de origen orgánico).',{sz:8.5,gray:1})}
    H('Medidas para reducir amoniaco y gases de efecto invernadero (art. 6.5 y anexo V)');
    var ms=PAB_MIT.filter(function(x){return (mit[x[0]]||{}).a===1});
    if(ms.length)T([{h:'Medida',w:90},{h:'Cómo se aplica',w:92}],ms.map(function(x){var m=mit[x[0]];return [x[2],m.d===undefined?x[3]:m.d]}));else Pp(u.d0===1?'No procede: sin aplicaciones esta campaña.':'Sin medidas marcadas.',{sz:9,gray:1});
    H('Restricciones y condiciones de aplicación (art. 4.4)');
    T([{h:'Condición',w:100},{h:'Comprobado',w:22},{h:'Observación',w:60}],PAB_RES.map(function(x){var m=res[x[0]]||{};return [x[1],m.a===1?'Sí':(m.a===0?'No':'—'),m.a!==undefined?(m.d===undefined?x[2]:m.d):'']}));
    var pt=pabPerTxt(u.ck), pm=C.rs.reduce(function(a,r){var x=((u.rc||{})[r.ref]||{}).pe;return abHas(x)?Math.max(a,+x):a},-1);
    if(pt)Pp('Periodo sin abono nitrogenado (anexo II): '+pt+(u.rg==='regadio'&&u.loc!==0?' (no se aplica con riego localizado)':'')+'.',{sz:8.5});
    if(pm>=0)Pp('Pendiente máxima de los recintos: '+trTxt(pm,1)+' %.',{sz:8.5});
    var SL=pabSlope(u,C); if(!u.d0){if(SL.x15.length)Pp('Recintos con más del 15 % de pendiente, que no se abonan (apdo. 4.3.2 c del Decreto 80/2025): '+pabSlopeTxt(SL.x15)+'.',{sz:8.5,b:1});
      if(SL.b8.length)Pp('Barbechos con más del 8 % de pendiente, sin aplicación en superficie (apdo. 4.3.2 b): '+pabSlopeTxt(SL.b8)+'.',{sz:8.5,b:1});
      if(u.zv&&SL.pmax>2)Pp('Pendiente superior al 2 %: la fertilización se combina con medidas de precaución (apdo. 4.3.2 a)'+((u.mo||{}).cub?': '+u.mo.cub:'')+'.',{sz:8.5})}
    H('Comprobación del contenido del plan');
    T([{h:'Apartado',w:100},{h:'Estado',w:22},{h:'Detalle',w:60}],pabCheck(f,p,u,C,RL).map(function(x){return [x.l,x.na?'No procede':(x.c?'Completo':'Pendiente'),x.w||'']}));
  });
  B.push({t:'np'});
  H('Modificaciones del plan durante la campaña');
  if((p.mods||[]).length)T([{h:'Fecha',w:20},{h:'Unidad / recintos',w:50},{h:'Motivo',w:40},{h:'Cambio realizado',w:46},{h:'Responsable',w:26}],p.mods.map(function(m){return [pabIso(m.f)?diaFmt(pabIso(m.f)):'',m.u||'',m.mo||'',m.ch||'',m.re||'']}));else Pp('Sin modificaciones.',{sz:9,gray:1});
  H('Asesoramiento técnico y firmas (art. 6.6 y 21)');
  K('Asesor en fertilización',A.n||'—');K('NIF',A.nif||'');K('Titulación',A.tit||'');K('Nº REGFER (sección asesores)',A.regfer||'');K('Acreditación',A.acr||'');K('Fecha del asesoramiento',pabIso(p.fa)?diaFmt(pabIso(p.fa)):'');
  var fT=pabFirmaT(p,t); B.push({t:'sign',a:A.n||'',b:fT.b,l:fT.l});
  return {B:B,title:f.nombre,top:'Plan de abonado · Campaña '+pabCamp(p.y),foot:'Plan de abonado '+pabCamp(p.y)+' · '+f.nombre};
}
/* Entrega 103: quién firma a la derecha. El campo «Titular que firma» del plan manda; si no, el representante de la finca; si no, el titular */
function pabFirmaT(p,t){t=t||{};var ft=String(p.ft||'').trim();
  if(t.rn&&(!ft||ft.toLowerCase()===String(t.rn).toLowerCase()))return {l:'Firma del representante',b:daRepTxt(t)+(t.n?', en representación de '+t.n:'')};
  return {l:'Firma del titular',b:ft||t.n||''}}
function pabCell(v){if(v&&typeof v==='object'&&v.n!==undefined)return trTxt(v.n,v.d);return v==null?'':String(v)}
/* Entrega 83: que ningún apartado de los PDF empiece al pie de una página y se corte.
   El PDF se construye, se mide dónde empieza y acaba cada apartado y, si uno que cabe en una página ha quedado partido
   (o empieza con menos de 45 mm libres), se vuelve a construir empezándolo en la página siguiente. Se repite hasta que no quede ninguno. */
var KP=null;
function kpT(doc,y,top){if(!KP)return false;var i=KP.n++;KP.rec.push({p:doc.getNumberOfPages(),y:y});return !!KP.brk[i]&&y>top+2}
function kpEnd(doc,y){if(KP&&!KP.end)KP.end={p:doc.getNumberOfPages(),y:y}}
function kpRun(fn,top,bot){
  var brk={},out=null;
  for(var g=0;g<60;g++){
    KP={n:0,rec:[],brk:brk,end:null};
    try{out=fn()}finally{var R=KP.rec,E=KP.end;KP=null}
    if(!E)return out;
    var U=bot-top,add=-1;
    for(var i=0;i<R.length;i++){if(brk[i])continue;var a=R[i],b=R[i+1]||E;if(b.p===a.p||a.y<=top+2)continue;
      var h=(bot-a.y)+(b.p-a.p-1)*U+(b.y-top);
      /* Entrega 144: si queda un tercio de página libre o más, el apartado empieza aquí aunque siga en la página siguiente (antes dejaba medias páginas en blanco) */
      if((h<=U-2&&(bot-a.y)<U*0.3)||(bot-a.y)<45){add=i;break}}
    if(add<0)return out;
    brk[add]=1;
  }
  return out;
}
function buildPabPdf(f,p){return kpRun(function(){return buildPabPdf0(f,p)},18,281)}
function buildPabPdf0(f,p){
  var D=pabBlocks(f,p), jsPDF=window.jspdf.jsPDF, doc=new jsPDF({unit:'mm',format:'a4'}), PW=210, PH=297, M=14, CW=PW-2*M, ink=[23,33,38], gray=[90,103,101], y=0, FN=pdfFont(doc);
  var room=function(h){if(y+h>PH-16){doc.addPage();y=18}};
  var tbH=function(b){return 7+5.8+b.rows.length*6.2+(b.tot?6:0)+4};
  var title=function(t,need,big){if(kpT(doc,y,18)){doc.addPage();y=18}need=Math.max(16,need||16);if(y>40)y+=2.5;room(need>PH-40?40:need);doc.setFont(FN,'bold');doc.setFontSize(big?14:11.5);doc.setTextColor.apply(doc,ink);var ls=doc.splitTextToSize(t,CW);doc.text(ls,M,y);if(!big){doc.setDrawColor.apply(doc,ink);doc.setLineWidth(.4);doc.line(M,y+1.6+(ls.length-1)*5,PW-M,y+1.6+(ls.length-1)*5)}y+=7+(ls.length-1)*5};
  var para=function(t,o){o=o||{};doc.setFont(FN,o.b?'bold':'normal');doc.setFontSize(o.sz||9.5);doc.setTextColor.apply(doc,o.gray?gray:ink);var ls=doc.splitTextToSize(String(t),CW);room(ls.length*4.2+1);doc.text(ls,M,y);y+=ls.length*4.2+1.5};
  var kv=function(k,v){doc.setFontSize(9.5);var ls=doc.splitTextToSize(String(v),CW-58);room(ls.length*4.4+1);doc.setFont(FN,'bold');doc.setTextColor.apply(doc,ink);doc.text(doc.splitTextToSize(k,56),M,y);doc.setFont(FN,'normal');doc.text(ls,M+58,y);y+=Math.max(ls.length,doc.splitTextToSize(k,56).length)*4.4+1};
  var tbl=function(cols,rows,tot){
    var sc=CW/cols.reduce(function(a,c){return a+c.w},0); cols=cols.map(function(c){return {h:c.h,w:c.w*sc,a:c.a}});
    var head=function(){doc.setFont(FN,'bold');doc.setFontSize(8);doc.setTextColor.apply(doc,ink);var x=M,hh=1;cols.forEach(function(c){var l=doc.splitTextToSize(c.h,c.w-2);hh=Math.max(hh,l.length);doc.text(l,c.a==='r'?x+c.w-1.5:x+1,y,{align:c.a==='r'?'right':'left'});x+=c.w});y+=(hh-1)*3.4;doc.setDrawColor.apply(doc,ink);doc.setLineWidth(.3);doc.line(M,y+1.6,PW-M,y+1.6);y+=5.8};
    room(14);head();
    rows.forEach(function(r,i){doc.setFont(FN,'normal');doc.setFontSize(8.3);
      var cells=cols.map(function(c,j){return doc.splitTextToSize(pabCell(r[j]),c.w-2)}), h=Math.max.apply(null,cells.map(function(c){return c.length}))*3.7+1.8;
      if(y+h>PH-16){doc.addPage();y=18;head();doc.setFont(FN,'normal');doc.setFontSize(8.3)}
      if(i%2===0){doc.setFillColor(243,245,244);doc.rect(M,y-3.6,CW,h,'F')}
      var x=M;cols.forEach(function(c,j){doc.setTextColor.apply(doc,ink);doc.text(cells[j],c.a==='r'?x+c.w-1.5:x+1,y,{align:c.a==='r'?'right':'left'});x+=c.w});y+=h});
    if(tot){if(y+8>PH-16){doc.addPage();y=18;head()}doc.setDrawColor.apply(doc,ink);doc.setLineWidth(.3);doc.line(M,y-3.6,PW-M,y-3.6);doc.setFont(FN,'bold');doc.setFontSize(8.3);var xt=M;cols.forEach(function(c,j){var v=pabCell(tot[j]);if(v)doc.text(doc.splitTextToSize(v,c.w-2),c.a==='r'?xt+c.w-1.5:xt+1,y,{align:c.a==='r'?'right':'left'});xt+=c.w});y+=6}
    y+=3};
  /* cabecera */
  doc.setFillColor.apply(doc,ink);doc.rect(0,0,PW,24,'F');doc.setFillColor(245,196,0);doc.rect(0,24,PW,1.4,'F');doc.setTextColor(255,255,255);doc.setFont(FN,'normal');doc.setFontSize(9);doc.text(D.top,M,9);doc.text(new Date().toLocaleDateString('es-ES',{day:'numeric',month:'long',year:'numeric'}),PW-M,9,{align:'right'});
  doc.setFont(FN,'bold');doc.setFontSize(20);var nm=String(D.title);while(nm.length>1&&doc.getTextWidth(nm)>CW-4)nm=nm.slice(0,-1);doc.text(nm,M,19);y=33;
  D.B.forEach(function(b,i){var nx=D.B[i+1];
    /* Entrega 133: cada unidad empieza en página nueva solo si queda menos de media página; si no, una raya y sigue */
    if(b.t==='np'){if(y>PH*0.7){doc.addPage();y=18} /* Entrega 138: la unidad siguiente empieza en la misma página si queda casi un tercio */else if(y>40){y+=4;doc.setDrawColor(200,206,204);doc.setLineWidth(.5);doc.line(M,y,PW-M,y);y+=8}return}
    if(b.t==='h1'){title(b.x,30,true);return}
    if(b.t==='h'){var need=16;if(nx&&nx.t==='tb')need=tbH(nx);else if(nx&&nx.t==='kv'){var n=0;for(var j=i+1;j<D.B.length&&D.B[j].t==='kv';j++)n++;need=7+n*5.4}title(b.x,need);return}
    if(b.t==='p')return para(b.x,b.o);
    if(b.t==='kv')return kv(b.k,b.v);
    if(b.t==='tb')return tbl(b.cols,b.rows,b.tot);
    if(b.t==='sign'){room(40);y+=6;doc.setDrawColor.apply(doc,ink);doc.setLineWidth(.3);doc.line(M,y+22,M+80,y+22);doc.line(PW-M-80,y+22,PW-M,y+22);doc.setFont(FN,'normal');doc.setFontSize(9);doc.setTextColor.apply(doc,ink);doc.text('Firma del asesor',M,y);doc.text(b.l||'Firma del titular',PW-M-80,y);doc.setFontSize(8.5);doc.setTextColor.apply(doc,gray);var la=b.a?doc.splitTextToSize(String(b.a),80):[],lb=b.b?doc.splitTextToSize(String(b.b),80):[];if(la.length)doc.text(la,M,y+27);if(lb.length)doc.text(lb,PW-M-80,y+27);y+=30+Math.max(1,la.length,lb.length)*3.8}
  });
  kpEnd(doc,y);
  var n=doc.getNumberOfPages();for(var pg=1;pg<=n;pg++){doc.setPage(pg);doc.setFont(FN,'normal');doc.setFontSize(8);doc.setTextColor.apply(doc,gray);doc.text(D.foot,M,PH-8);doc.text(pg+' / '+n,PW-M,PH-8,{align:'right'})}
  return doc.output('blob');
}
function buildPabXlsx(f,p){
  var D=pabBlocks(f,p), X=xlDoc('Plan de abonado',D.foot);
  X.band(D.top,D.title,new Date().toLocaleDateString('es-ES',{day:'numeric',month:'long',year:'numeric'}));
  var xc=function(v){if(v&&typeof v==='object'&&v.n!==undefined)return xlN(v.n,v.d);return v==null?'':xlAuto(v)};
  D.B.forEach(function(b){
    if(b.t==='np')return X.gap(2);
    if(b.t==='h1'){X.gap();X.para(b.x,{b:1,sz:14});return}
    if(b.t==='h')return X.title(b.x);
    if(b.t==='p')return X.para(b.x,{sz:b.o.sz||9.5,b:b.o.b,col:b.o.gray?XC.gray:undefined});
    if(b.t==='kv')return X.kv(b.k,b.v);
    if(b.t==='tb'){var sc=182/b.cols.reduce(function(a,c){return a+c.w},0);return X.tbl(b.cols.map(function(c){return {h:c.h,w:c.w*sc,a:c.a}}),b.rows.map(function(r){return r.map(xc)}),{tot:b.tot?b.tot.map(xc):null})}
    if(b.t==='sign'){X.gap(2);X.kv('Firma del asesor',b.a||'');X.gap(3);X.kv(b.l||'Firma del titular',b.b||'');}
  });
  return xlBuild([X.sh]);
}
function pabOut(f,p,kind){
  try{
    if(kind==='pdf'){if(!loadPdfLib()){pdfNoListo();return}}
    var b=kind==='pdf'?buildPabPdf(f,p):buildPabXlsx(f,p);
    showDl(b,'plan-de-abonado-'+safeName(f.nombre)+'-'+campFile(+p.y+1)+'.'+(kind==='pdf'?'pdf':'xlsx'),'Plan de abonado listo',esc(f.nombre)+' · campaña '+pabCamp(p.y));
  }catch(e){toast('No he podido crear el plan de abonado.')}
}
/* Plan de abonado — herramienta independiente (versión 2).
   Todo se guarda en este aparato (localStorage) y, si se quiere, en un archivo de copia. No hay servidor ni cuentas.
   El cálculo, el PDF y el Excel del plan vienen de base.js (RD 1051/2022 y Decreto 80/2025 de Castilla-La Mancha). */
var VERSION="2 · 17 nov 2026";
var KEY='plan-abonado:datos', KEYUI='plan-abonado:vista';
var $=function(s,r){return (r||document).querySelector(s)}, $$=function(s,r){return [].slice.call((r||document).querySelectorAll(s))};

/* ───── datos ───── */
var S={fincas:[],recintos:{},x:{asesor:{},abonos:[]},copia:0,mod:0};
var UI={v:'inicio',fid:null,step:'rec'};
var GTOK=false, EXP={fid:null};
function uiDef(){return {v:'inicio',fid:null,step:'rec',q:'',sort:'ref',cost:0,bu:null,visto:UI&&UI.visto}}
function cargar(){
  try{var t=localStorage.getItem(KEY);if(t){var o=JSON.parse(t);if(o&&Array.isArray(o.fincas)){S=o;S.recintos=S.recintos||{};S.x=S.x||{};S.x.asesor=S.x.asesor||{};S.x.abonos=S.x.abonos||[]}}}catch(e){}
  try{var u=JSON.parse(localStorage.getItem(KEYUI)||'null');if(u&&u.v){UI=Object.assign(uiDef(),u)}}catch(e){}
  if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null}
}
var guardaFallo=false;
function saveRaw(){try{localStorage.setItem(KEY,JSON.stringify(S));guardaFallo=false;return true}catch(e){if(!guardaFallo)toast('Este navegador no deja guardar los datos. Guarda una copia en archivo.');guardaFallo=true;return false}}
function save(){S.mod=Date.now();saveRaw();fsWrite()}
window.addEventListener('storage',function(ev){if(ev.key!==KEY||!ev.newValue)return;try{var o=JSON.parse(ev.newValue);if(!o||!Array.isArray(o.fincas))return;S=o;S.recintos=S.recintos||{};S.x=S.x||{};S.x.asesor=S.x.asesor||{};S.x.abonos=S.x.abonos||[];UNDO.length=0;if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null}if(typeof render==='function'){render();toast('Cargados los cambios hechos en otra pestaña')}}catch(e){}});
function saveUI(){try{localStorage.setItem(KEYUI,JSON.stringify(UI))}catch(e){}}
try{if(navigator.storage&&navigator.storage.persist)navigator.storage.persist()}catch(e){}

/* piezas que el plan espera de su aplicación de origen */
function finca(id){return S.fincas.filter(function(f){return f.id===id})[0]||null}
function fincaRecs(fid){return Object.keys(S.recintos).map(function(k){return S.recintos[k]}).filter(function(r){return r.fincaId===fid}).sort(function(a,b){return a.ref.localeCompare(b.ref,'es',{numeric:true})})}
function grupo(){return null}
function grupoFincas(){return []}
function expGGet(){return null}
function grRecs(){return []}
function renderFincas(){render()}
/* en pantalla las hectáreas van con 2 decimales (el PDF y el Excel no usan esta función) */
ha4=function(n){return nf(n,2)};

/* deshacer */
var UNDO=[];
function undoPush(){try{var t=JSON.stringify({fincas:S.fincas,recintos:S.recintos,x:S.x});if(UNDO[UNDO.length-1]===t)return;UNDO.push(t);if(UNDO.length>40)UNDO.shift()}catch(e){}}
function undoDo(){var t=UNDO.pop();if(!t){toast('No hay nada que deshacer');return}var o=JSON.parse(t);S.fincas=o.fincas;S.recintos=o.recintos;S.x=o.x;save();var mo=!$('#mapa').hidden;if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null;if(mo){mapaCerrar();mo=false}}render();if(mo){mRestyle();mSide()}toast('Deshecho')}
function undoToast(msg){toast(msg,'Deshacer',undoDo)}

/* avisos: cortos, arriba a la derecha en el ordenador y encima de la barra en el móvil; no tapan botones */
var toastT=null;
function toast(msg,btn,fn){var el=$('#toast');if(!el)return;clearTimeout(toastT);el.innerHTML='<div role="status"><span>'+esc(msg)+'</span>'+(btn?'<button type="button" id="tbtn">'+esc(btn)+'</button>':'')+'<button type="button" class="tx" aria-label="Cerrar aviso">×</button></div>';
  if(btn)$('#tbtn').onclick=function(){el.innerHTML='';fn()};
  $('.tx',el).onclick=function(){el.innerHTML=''};
  toastT=setTimeout(function(){el.innerHTML=''},btn?5500:3200)}
function toastOff(){clearTimeout(toastT);var el=$('#toast');if(el)el.innerHTML=''}

/* ───── librerías que se cargan solo al usarlas (PDF, lector de la PAC y mapa) ───── */
function cargaScript(src){return new Promise(function(res,rej){var s=document.createElement('script');s.src=src;s.onload=function(){res()};s.onerror=function(){s.remove();rej(new Error('net'))};document.head.appendChild(s)})}
var pdfP=null, pdfjsP=null, mapP=null;
function ensurePdf(){if(window.jspdf&&window.jspdf.jsPDF)return Promise.resolve();if(!pdfP)pdfP=cargaScript('lib/pdf.js').catch(function(e){pdfP=null;throw e});return pdfP}
function ensurePdfJs(){if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);if(!pdfjsP)pdfjsP=cargaScript('lib/pdfjs.js').then(function(){window.pdfjsLib.GlobalWorkerOptions.workerSrc='lib/pdfjs-worker.js';return window.pdfjsLib}).catch(function(e){pdfjsP=null;throw e});return pdfjsP}
function ensureMap(){if(window.L&&window.L.map)return Promise.resolve();if(!mapP){var l=document.createElement('link');l.rel='stylesheet';l.href='lib/leaflet.css';document.head.appendChild(l);mapP=cargaScript('lib/leaflet.js').catch(function(e){mapP=null;throw e})}return mapP}
function loadPdfLib(){return !!(window.jspdf&&window.jspdf.jsPDF)}
function pdfNoListo(){toast('No he podido preparar el PDF. Ábrela una vez con internet y prueba otra vez.')}
function pdfFont(doc){
  try{var F=window.PA_FONTS;if(!F||!F.r)return 'helvetica';
    if(!doc.__bw){doc.addFileToVFS('bw.ttf',F.r);doc.addFont('bw.ttf','Barlow','normal','Identity-H');doc.addFileToVFS('bw-b.ttf',F.b);doc.addFont('bw-b.ttf','Barlow','bold','Identity-H');
      if(F.cr&&F.cb){doc.addFileToVFS('bwc.ttf',F.cr);doc.addFont('bwc.ttf','BarlowC','normal','Identity-H');doc.addFileToVFS('bwc-b.ttf',F.cb);doc.addFont('bwc-b.ttf','BarlowC','bold','Identity-H');doc.__bwc=1}
      doc.__bw=1}
    return 'Barlow'}catch(e){return 'helvetica'}
}

/* ───── ventanas (hojas) y descargas ───── */
function sheet(h,cls){var el=$('#sheet');el.className=cls||'';el.innerHTML='<div class="box" role="dialog" aria-modal="true">'+h+'</div>';el.hidden=false;
  setTimeout(function(){var x=$('#sheet .box [autofocus]')||$('#sheet .box .btn.primary');if(x&&isDesk())try{x.focus()}catch(e){}},30)}
function sheetClose(){var el=$('#sheet');el.hidden=true;el.innerHTML='';el.className='';PACI=null;COPIA=null}
var DL=null;
function showDl(blob,name,title,extra){
  DL={blob:blob,name:name}; var sh=canShareFile(blob,name), phone=!isDesk()&&sh;
  sheet('<h2>'+esc(title||'Archivo listo')+'</h2><p class="small">'+esc(name)+' · '+Math.max(1,Math.round(blob.size/1024))+' KB</p>'+(extra?'<p class="small">'+extra+'</p>':'')+
    (phone?'<p class="small">Pulsa «Guardar / Enviar…» y elige «Guardar en Archivos», el correo o WhatsApp.</p>':'<p class="small">Se guarda en la carpeta de Descargas.</p>')+
    '<div class="acts">'+(phone?'<button class="btn primary xl" data-act="ap-dl-share">Guardar / Enviar…</button><button class="btn" data-act="ap-dl-save">Descargar</button>':'<button class="btn primary xl" data-act="ap-dl-save">Guardar archivo</button>'+(sh?'<button class="btn" data-act="ap-dl-share">Enviar…</button>':''))+
    (/\.pdf$/.test(name)&&isDesk()?'<button class="btn" data-act="ap-dl-view">Abrir en otra pestaña</button>':'')+'<button class="btn" data-act="ap-close">Cerrar</button></div>');
}
function dlSave(){if(!DL)return;var u=URL.createObjectURL(DL.blob),a=document.createElement('a');a.href=u;a.download=DL.name;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);a.remove()},4000)}
function dlShare(){if(!DL)return;var fl=new File([DL.blob],DL.name,{type:DL.blob.type||'application/octet-stream'});navigator.share({files:[fl],title:DL.name}).catch(function(e){if(e&&e.name!=='AbortError')dlSave()})}
function dlView(){if(!DL)return;var u=URL.createObjectURL(DL.blob);window.open(u,'_blank')}

/* ───── lluvia de la zona: valores normales de AEMET (1981-2010) de las estaciones principales de Castilla-La Mancha ───── */
var CLIMA={2:{a:353,o:110,i:77,p:110,v:55,src:'AEMET, Albacete Base Aérea (8175), valores normales 1981-2010',ll:[38.952,-1.863],n:'Albacete Base Aérea'},
  13:{a:402,o:124,i:124,p:117,v:36,src:'AEMET, Ciudad Real (4121), valores normales 1981-2010',ll:[38.989,-3.920],n:'Ciudad Real'},
  16:{a:501,o:150,i:136,p:145,v:71,src:'AEMET, Cuenca (8096), valores normales 1981-2010',ll:[40.067,-2.138],n:'Cuenca'},
  19:{a:468,o:136,i:90,p:144,v:96,src:'AEMET, Molina de Aragón (3013), valores normales 1981-2010',ll:[40.844,-1.885],n:'Molina de Aragón'},
  45:{a:342,o:105,i:92,p:106,v:40,src:'AEMET, Toledo (3260B), valores normales 1982-2010',ll:[39.885,-4.045],n:'Toledo'}};
var PROV_N={2:'Albacete',13:'Ciudad Real',16:'Cuenca',19:'Guadalajara',45:'Toledo'};
function provDe(f){var m={};fincaRecs(f.id).forEach(function(r){var p=+String(r.ref).split(':')[0];m[p]=(m[p]||0)+haC(r)});var b=Object.keys(m).sort(function(a,c){return m[c]-m[a]})[0];return b?+b:null}
/* centro de la finca (de los recintos con coordenadas, pesado por superficie) */
function fincaLL(f){var la=0,lo=0,w=0;fincaRecs(f.id).forEach(function(r){if(r.ll){var h=haC(r)||1;la+=r.ll[0]*h;lo+=r.ll[1]*h;w+=h}});return w?[la/w,lo/w]:null}
function kmEntre(a,b){var R=6371,d=Math.PI/180,x=(b[1]-a[1])*d*Math.cos((a[0]+b[0])/2*d),y=(b[0]-a[0])*d;return Math.sqrt(x*x+y*y)*R}
/* estación más cercana a los recintos; si no hay coordenadas, la de la provincia */
function climaKey(f){var c=fincaLL(f);if(c){var best=null,bd=1e9;Object.keys(CLIMA).forEach(function(k){var d=kmEntre(c,CLIMA[k].ll);if(d<bd){bd=d;best=k}});return {k:+best,km:Math.round(bd)}}var p=provDe(f);return {k:CLIMA[p]?p:2,km:null}}
function climaDe(f){var c=CLIMA[climaKey(f).k];return {a:c.a,o:c.o,i:c.i,p:c.p,v:c.v,src:c.src}}

/* ───── catálogo general de abonos (riqueza típica del tipo de abono; la de la etiqueta manda) ───── */
var CAT=[
  ['Nitrogenados',[['Urea 46 %',46,0,0],['Nitrato amónico 33,5 %',33.5,0,0],['Nitrato amónico cálcico (NAC) 27 %',27,0,0],['Nitrosulfato amónico 26 %',26,0,0],['Sulfato amónico 21 %',21,0,0],['Nitrato cálcico 15,5 %',15.5,0,0],['Nitrato magnésico 11 %',11,0,0]]],
  ['Fosfatados',[['Superfosfato simple 18 %',0,18,0],['Superfosfato triple 46 %',0,46,0],['Fosfato monoamónico (MAP) 12-61-0',12,61,0],['Fosfato diamónico (DAP) 18-46-0',18,46,0],['Fosfato monopotásico (MKP) 0-52-34',0,52,34]]],
  ['Potásicos',[['Cloruro potásico 60 %',0,0,60],['Sulfato potásico 50 %',0,0,50],['Nitrato potásico 13-0-46',13,0,46]]],
  ['Complejos',[['Complejo 8-15-15',8,15,15],['Complejo 8-24-8',8,24,8],['Complejo 9-18-27',9,18,27],['Complejo 12-12-17',12,12,17],['Complejo 12-24-12',12,24,12],['Complejo 15-15-15',15,15,15],['Complejo 20-10-10',20,10,10]]],
  ['Líquidos',[['Solución nitrogenada N32',32,0,0,0,1.32],['Solución nitrogenada N20',20,0,0,0,1.26],['Ácido fosfórico 54 %',0,54,0,0,1.58],['Ácido nítrico 59 % (13 % N)',13,0,0,0,1.36],['Tiosulfato amónico 12 %',12,0,0,0,1.32]]],
  ['Orgánicos (riqueza orientativa)',[['Estiércol de oveja',0.8,0.5,1,1],['Estiércol de vacuno',0.5,0.3,0.5,1],['Gallinaza',2,1.8,1.5,1],['Compost',1.2,0.8,1,1],['Abono orgánico granulado',3,2,2,1]]]];
function catGen(){var o=[];CAT.forEach(function(g){g[1].forEach(function(x){o.push({n:x[0],rn:x[1],rp:x[2],rk:x[3],org:x[4]?1:0,dens:x[5]||0,g:g[0]})})});return o}
/* «Mis abonos»: los que el usuario guarda (con su riqueza y, si quiere, su precio) y los que ya ha usado en algún plan */
function catMios(){var gen={},o=[],seen={};catGen().forEach(function(x){gen[noAcc(x.n).toLowerCase()]=1});
  (S.x.abonos||[]).forEach(function(a){var k=noAcc(a.n||'').toLowerCase().trim();if(!k||seen[k])return;seen[k]=1;o.push({n:a.n,rn:+a.rn||0,rp:+a.rp||0,rk:+a.rk||0,org:a.org?1:0,pz:a.pz,dens:+a.dens||0,g:'Mis abonos',mio:1})});
  S.fincas.forEach(function(f){(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){((u.pg||{}).rows||[]).forEach(function(q){var k=noAcc(q.n||'').toLowerCase().trim();if(!k||gen[k]||seen[k])return;if(!(+q.rn||+q.rp||+q.rk))return;seen[k]=1;o.push({n:q.n,rn:+q.rn||0,rp:+q.rp||0,rk:+q.rk||0,org:q.org===1?1:0,g:'Usados antes'})})})})});
  return o.sort(function(a,b){return (b.mio||0)-(a.mio||0)||a.n.localeCompare(b.n,'es')})}
function catTodo(){return catMios().concat(catGen())}
function catBusca(q){var t=noAcc(q||'').toLowerCase().trim(), L=catTodo();if(!t)return L;var ws=t.split(/\s+/);
  return L.filter(function(x){var n=noAcc(x.n).toLowerCase()+' '+[x.rn,x.rp,x.rk].join('-')+' '+noAcc(x.g).toLowerCase();return ws.every(function(w){return n.indexOf(w)>=0})})}
function catDe(n){var k=noAcc(n||'').toLowerCase().trim();return catTodo().filter(function(x){return noAcc(x.n).toLowerCase()===k})[0]||null}

/* ───── cultivos ───── */
var CULTS=['Pistacho','Almendro','Olivo','Nogal','Viña','Melocotonero','Otros frutales','Trigo','Trigo duro','Cebada','Avena','Centeno','Triticale','Maíz','Barbecho','Otro'];
function cultOpts(cur){var L=CULTS.slice();if(cur&&L.indexOf(cur)<0)L.unshift(cur);return '<option value="">Sin cultivo</option>'+L.map(function(c){return '<option'+(c===cur?' selected':'')+'>'+esc(c)+'</option>'}).join('')}
function cultSel(attr,cur){return '<select class="trsel ap-in" '+attr+'>'+cultOpts(cur)+'</select>'}
var USOS={TA:'Tierra arable',TH:'Huerta',FS:'Frutos secos',FY:'Frutales',OV:'Olivar',VI:'Viñedo',FL:'Frutos secos y olivar',FV:'Frutos secos y viñedo',FF:'Frutos secos y frutal',OF:'Olivar y frutal',VO:'Viñedo y olivar',VF:'Viñedo y frutal',OP:'Otros cultivos permanentes',PS:'Pastizal',PR:'Pasto arbustivo',PA:'Pasto con arbolado',FO:'Forestal',MT:'Matorral',IM:'Improductivo',ED:'Edificaciones',CA:'Viales',AG:'Agua',ZU:'Zona urbana',ZC:'Zona concentrada',IV:'Invernaderos',CI:'Cítricos',EP:'Elemento del paisaje'};
function usoTxt(u){return u?(USOS[u]?u+' · '+USOS[u]:u):''}
/* cultivo propuesto por el uso del SIGPAC (solo cuando el uso lo dice claro) */
function cultDeUso(u){return {OV:'Olivo',VI:'Viña',FY:'Otros frutales'}[u]||''}

/* ───── recintos: código SIGPAC ───── */
function refParse(s){var a=String(s||'').trim().split(/[^0-9]+/).filter(Boolean).map(Number);
  if(a.length===5)a=[a[0],a[1],0,0,a[2],a[3],a[4]];
  if(a.length===4)a=[a[0],a[1],0,0,a[2],a[3],1];
  if(a.length!==7||a.some(function(x){return !isFinite(x)})||!a[0]||a[0]>52||!a[1])return null;return a.join(':')}
function recNew(ref,o){o=o||{};var r={ref:ref,ha:+o.ha||0,hc:+o.hc>0?+o.hc:undefined,cultivo:o.cultivo||'',riego:o.riego||'',variedad:o.variedad||'',nombre:o.nombre||'',fincaId:o.fincaId};
  ['hs','pend','uso','coef','ll','alt','adm','inc'].forEach(function(k){if(o[k]!==undefined&&o[k]!==null)r[k]=o[k]});return r}
function addRecs(f,list){var added=0,skip=[];list.forEach(function(o){var ex=S.recintos[o.ref];if(ex&&ex.fincaId!==f.id&&finca(ex.fincaId)){skip.push(o.ref);return}if(ex&&ex.fincaId===f.id){skip.push(o.ref);return}o.fincaId=f.id;S.recintos[o.ref]=recNew(o.ref,o);added++});return {added:added,skip:skip}}
/* geometría WKT del SIGPAC (lon lat, ETRS89) → punto central [lat, lon] */
function wktLL(w){var m=String(w||'').match(/-?\d+\.?\d*\s+-?\d+\.?\d*/g);if(!m||!m.length)return null;var x=0,y=0;m.forEach(function(p){var a=p.trim().split(/\s+/);x+=+a[0];y+=+a[1]});x/=m.length;y/=m.length;
  if(!(Math.abs(y)<=90&&Math.abs(x)<=180))return null;return [Math.round(y*1e5)/1e5,Math.round(x*1e5)/1e5]}
/* SIGPAC: superficie, pendiente media, uso y situación del recinto (consulta pública, necesita internet) */
var SIGB={};
function sigInfo(ref){return fetch('https://sigpac-hubcloud.es/servicioconsultassigpac/query/recinfo/'+ref.split(':').join('/')+'.json').then(function(r){if(!r.ok)throw 1;return r.json()}).then(function(j){var o=Array.isArray(j)?j[0]:j;if(!o)throw 1;return o})}
function sigAplica(r,o){var nn=function(v){return v==null||v===''?null:+v};
  if(o.superficie!=null&&!(r.ha>0))r.ha=+o.superficie; if(o.superficie!=null)r.hs=+o.superficie; if(o.pendiente_media!=null)r.pend=+o.pendiente_media; if(o.uso_sigpac)r.uso=o.uso_sigpac; if(o.coef_regadio!=null)r.coef=+o.coef_regadio;
  if(nn(o.altitud))r.alt=+o.altitud; if(o.admisibilidad!=null)r.adm=+o.admisibilidad; if(o.incidencias!=null)r.inc=String(o.incidencias); var ll=wktLL(o.wkt); if(ll)r.ll=ll}
function sigLeer(f,refs,cb){
  refs=refs.filter(function(x){return S.recintos[x]&&!SIGB[x]}); if(!refs.length){if(cb)cb(0,0);return}
  if(navigator.onLine===false){toast('Sin internet no puedo leer el SIGPAC. Prueba luego.');if(cb)cb(0,refs.length);return}
  var idx=0,ok=0,ko=0; refs.forEach(function(x){SIGB[x]=1});
  var one=function(){if(idx>=refs.length)return Promise.resolve();var x=refs[idx++];
    return sigInfo(x).then(function(o){var r=S.recintos[x];if(!r)return;sigAplica(r,o);ok++}).catch(function(){ko++}).then(function(){delete SIGB[x];return one()})};
  Promise.all([one(),one(),one(),one()]).then(function(){pendAPlanes(f);climaAuto(f);save();if(cb)cb(ok,ko);render()});
}
/* la pendiente del SIGPAC pasa a la tabla de recintos de cada unidad si allí está vacía */
function pendAPlanes(f){(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){(u.refs||[]).forEach(function(x){var r=S.recintos[x];if(!r||r.pend==null)return;u.rc=u.rc||{};var c=u.rc[x]=u.rc[x]||{};if(!abHas(c.pe))c.pe=r.pend})})})}
/* si la lluvia del plan es la puesta sola y ahora se sabe la estación más cercana, se cambia */
function climaAuto(f){(f.pab||[]).forEach(function(p){if(!p.clAuto)return;var c=climaDe(f);if(p.cl&&p.cl.src!==c.src){p.cl=c}})}

/* ───── PAC (PDF de la solicitud) y tablas CSV ───── */
var PACI=null;
function pacLeer(file){
  toast('Leyendo la PAC…');
  var rd=new FileReader();
  rd.onload=function(){
    var txt0='';
    ensurePdfJs().then(function(lib){return lib.getDocument({data:new Uint8Array(rd.result),isEvalSupported:false}).promise}).then(function(pdf){
      var ps=[];for(var p=1;p<=pdf.numPages;p++)ps.push(pdf.getPage(p).then(function(pg){var vt=pg.getViewport({scale:1}).transform;return pg.getTextContent().then(function(tc){return {items:tc.items,vt:vt}})}));
      return Promise.all(ps);
    }).then(function(pages){
      var o=pacRowsFromItems(pages); if(!o.rows.length)throw new Error('vacio');
      pages.slice(0,2).forEach(function(pg){txt0+=' '+pg.items.map(function(i){return i.str}).join(' ')});
      var nifs={};(txt0.match(/\b(\d{8}[A-HJ-NP-TV-Z]|[XYZ]\d{7}[A-HJ-NP-TV-Z]|[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J])\b/g)||[]).forEach(function(n){nifs[n]=1});
      var nl=Object.keys(nifs);
      PACI={kind:'pac',rows:o.rows.map(function(r){return {ref:[r.prov,r.muni,r.agr,r.zona,r.pol,r.par,r.rec].join(':'),mname:r.mname,ha:r.sigHa,hc:Math.abs(r.sigHa-r.decHa)>0.00005&&r.decHa>0?r.decHa:undefined,cultivo:pacCrop(r.prod)||pacCropVar(r.vari)||titleCase2(r.prod),riego:r.riego||'',variedad:titleCase2(r.vari)}}),year:o.year,nif:nl.length===1?nl[0]:'',name:''};
      var mn=PACI.rows.map(function(r){return r.mname}).filter(Boolean);PACI.name=mn.length?'Finca de '+titleCase2(mn[0]):'Mi finca';
      pacHoja();
    }).catch(function(e){PACI=null;toast(e&&e.message==='net'?'Para leer el PDF hace falta internet la primera vez.':'No he encontrado la tabla de recintos en ese PDF.')});
  };
  rd.readAsArrayBuffer(file);
}
function csvLeer(file){
  var rd=new FileReader();
  rd.onload=function(){
    var t=String(rd.result||'').replace(/^﻿/,''), lines=t.split(/\r?\n/).filter(function(l){return l.trim()});
    var l0=lines[0]||'',cnt=function(c){return l0.split(c).length},sep=cnt('\t')>Math.max(cnt(';'),cnt(','))?'\t':(cnt(';')>=cnt(',')?';':',');
    var hi=-1,H=null;for(var i=0;i<Math.min(lines.length,10);i++){var c=lines[i].split(sep).map(function(x){return noAcc(x).toLowerCase().trim()});if(c.some(function(x){return /^pol/.test(x)})&&c.some(function(x){return /^parc/.test(x)})){hi=i;H=c;break}}
    if(hi<0){toast('No encuentro las columnas «polígono» y «parcela» en la tabla.');return}
    var col=function(re){for(var k=0;k<H.length;k++)if(re.test(H[k]))return k;return -1};
    var cP=col(/^prov/),cM=col(/^(muni|mun\b|cod.*mun|termino)/),cA=col(/^agr/),cZ=col(/^zona/),cPo=col(/^pol/),cPa=col(/^parc/),cR=col(/^rec/),cH=col(/^(sup|ha\b|hect)/),cC=col(/^(cult|uso|prod)/),cRg=col(/^(rieg|regad|sist|s\/r)/),cV=col(/^vari/);
    var num=function(v){v=String(v||'').trim();if(/^\d+\.\d+$/.test(v))return parseFloat(v);var n=numEs(v);if(n===null||isNaN(n))n=parseFloat(String(v).replace(',','.'));return isFinite(n)?n:0};
    var rows=[];lines.slice(hi+1).forEach(function(l){var c=l.split(sep).map(function(x){return x.replace(/^"|"$/g,'').trim()});var g=function(k){return k>=0?c[k]||'':''};
      var ref=[g(cP),g(cM),g(cA)||0,g(cZ)||0,g(cPo),g(cPa),g(cR)||1].map(function(x){return parseInt(x,10)||0}).join(':');if(!refParse(ref))return;
      var rg=noAcc(g(cRg)).toLowerCase();rows.push({ref:refParse(ref),ha:num(g(cH)),cultivo:titleCase2(g(cC)),riego:/^r|regad/.test(rg)?'regadio':(/^s|secan/.test(rg)?'secano':''),variedad:titleCase2(g(cV))})});
    if(!rows.length){toast('La tabla no tiene filas con provincia, municipio, polígono y parcela.');return}
    PACI={kind:'csv',rows:rows,name:'Mi finca'};pacHoja();
  };
  rd.readAsText(file);
}
/* «Esto es lo que he entendido»: resumen de lo leído y, de un golpe, la finca y el plan */
function pacGrupos(rows){var m={},o=[];rows.forEach(function(r){if(r.off)return;var c=r.cultivo||'Sin cultivo',g=r.riego==='regadio'?'regadío':'secano',k=c+'§'+g;if(!m[k]){m[k]={n:c+' '+g,ha:0,k:0,rg:g};o.push(m[k])}m[k].ha+=(r.hc||r.ha||0);m[k].k++});return o}
function pacHoja(){toastOff();
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, rows=PACI.rows, on=rows.filter(function(r){return !r.off}), n=on.length, tot=on.reduce(function(a,r){return a+(r.ha||0)},0);
  var G=pacGrupos(rows), y0=pabY0(), conPlan=!f||!(f.pab||[]).length;
  var h='<h2>Esto es lo que he entendido</h2>';
  h+='<div class="und">'+
    (PACI.kind==='pac'?'<div><span class="k">Solicitud</span><b>PAC '+esc(PACI.year||'')+'</b></div>':'<div><span class="k">Origen</span><b>Tabla CSV</b></div>')+
    (PACI.nif?'<div><span class="k">NIF del titular</span><b>'+esc(PACI.nif)+'</b></div>':'')+
    '<div><span class="k">Recintos</span><b>'+n+' de '+rows.length+'</b></div><div><span class="k">Superficie</span><b>'+ha2(tot)+' ha</b></div></div>';
  h+='<div class="lbl">Unidades de producción que salen (cultivo y riego)</div><div class="ugl">'+G.map(function(g){return '<div><b>'+esc(g.n)+'</b><span>'+g.k+(g.k===1?' recinto':' recintos')+' · '+ha2(g.ha)+' ha'+(g.rg==='secano'&&g.ha<=10?' · <i>exenta (secano ≤ 10 ha)</i>':'')+'</span></div>'}).join('')+'</div>';
  if(rows.some(function(r){return !r.riego}))h+='<p class="small">Los recintos sin «R» ni «S» en la PAC se toman como secano: cámbialo en la tabla de recintos si hace falta.</p>';
  h+='<details class="more"'+(isDesk()?' open':'')+'><summary>Ver los '+rows.length+' recintos (toca uno para quitarlo o ponerlo)</summary><div class="pacl">'+rows.map(function(r,i){var ex=S.recintos[r.ref];return '<div class="pacrow'+(r.off?' off':'')+'" data-act="ap-pac-tog" data-i="'+i+'" role="checkbox" tabindex="0" aria-checked="'+(!r.off)+'"><i>'+(r.off?'':'✓')+'</i><div><b>'+esc(refLabel(r.ref))+'</b>'+(r.mname?' · '+esc(titleCase2(r.mname)):'')+'<small>'+esc([r.cultivo||'Sin cultivo',r.variedad,r.riego==='regadio'?'Regadío':(r.riego==='secano'?'Secano':''),ha2(r.ha)+' ha'+(r.hc?' (declaradas '+ha2(r.hc)+')':'')].filter(Boolean).join(' · '))+(ex?' · ya está en «'+esc((finca(ex.fincaId)||{}).nombre||'')+'»':'')+'</small></div></div>'}).join('')+'</div></details>';
  if(!f)h+='<label class="lbl" for="pac-name">Nombre de la finca</label><input id="pac-name" class="in" value="'+esc(PACI.name)+'" autocomplete="off">';
  if(conPlan)h+='<label class="ck"><input type="checkbox" id="pac-plan" checked> Crear también el plan de abonado de la campaña <select id="pac-y" class="trsel sm">'+[y0-1,y0,y0+1].map(function(v){return '<option value="'+v+'"'+(v===y0?' selected':'')+'>'+pabCamp(v)+'</option>'}).join('')+'</select></label>';
  h+='<div class="acts"><button class="btn primary xl" data-act="ap-pac-go"'+(n?'':' disabled')+'>'+(f?'Añadir '+n+(n===1?' recinto':' recintos'):'Crear la finca')+(conPlan?' y el plan':'')+'</button><button class="btn" data-act="ap-close">Cancelar</button></div>';
  sheet(h,'wide');
}
function pacGo(){
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, ni=$('#pac-name'), pc=$('#pac-plan'), py=$('#pac-y'), mk=pc&&pc.checked, y=py?+py.value:pabY0();
  undoPush();
  if(!f){f=fincaNueva((ni&&ni.value.trim())||PACI.name||'Mi finca')}
  if(PACI.nif){f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};if(!f.datos.tit.nif)f.datos.tit.nif=PACI.nif}
  var R=addRecs(f,PACI.rows.filter(function(r){return !r.off}).map(function(r){return Object.assign({},r)}));
  var refs=fincaRecs(f.id).map(function(r){return r.ref});
  var p=null; if(mk&&fincaRecs(f.id).length&&!(f.pab||[]).some(function(x){return +x.y===y}))p=planNuevo(f,y);
  sheetClose(); UI.v='finca';UI.fid=f.id;UI.step=p?'ud':'rec'; save(); saveUI(); window.__scrollTop=true; render();
  toast(R.added+(R.added===1?' recinto añadido':' recintos añadidos')+(R.skip.length?' · '+R.skip.length+' ya estaban':'')+(p?' y plan '+pabCamp(p.y)+' creado':'')+'. Leo el SIGPAC…');
  sigLeer(f,refs,function(ok,ko){if(ok||ko)toast('SIGPAC: '+ok+' leídos'+(ko?', '+ko+' sin respuesta':''))});
}
function fincaNueva(nm){var f={id:uid(),nombre:nm||'Mi finca',datos:{tit:{}},c:Date.now()};S.fincas.push(f);return f}

/* ───── planes: nuevo, y el de la campaña siguiente copiando lo que vale ───── */
var LEN=/^(pistacho|almendro|olivo|nogal|vid|melocotonero|frutal)$/;
function udPrefill(u){
  if(LEN.test(u.ck)&&!u.prec)u.prec=pabCkName(u.ck);
  if(!u.mit||!Object.keys(u.mit).length)mitPropuesta(u,true);
  if(!u.res||!Object.keys(u.res).length){u.res={};PAB_RES.forEach(function(x){u.res[x[0]]={a:1}});u.resProp=1}
}
function planNuevo(f,y,prev){
  var n={id:uid(),y:String(y),cl:climaDe(f),clAuto:1,ud:pabAuto(f),c:Date.now(),m:Date.now(),fe:hoyIso()};
  prev=prev||pabList(f)[0];
  if(prev){n.cl=Object.assign({},prev.cl||PAB_CL);delete n.clAuto;if(prev.clAuto)n.clAuto=1;
    n.ud.forEach(function(u){var o=(prev.ud||[]).filter(function(x){return x.n===u.n||(x.ck===u.ck&&x.rg===u.rg)})[0];if(!o)return;
      ['zv','tp','inh','rp','loc','leg','bobs'].forEach(function(k){if(o[k]!==undefined)u[k]=o[k]});
      ['su','rc','ag','ne','ap','mo','mit','res'].forEach(function(k){if(o[k]!==undefined)u[k]=JSON.parse(JSON.stringify(o[k]))});
      u.pg={src:'own',rows:((o.pg||{}).rows||[]).map(function(q){var c=JSON.parse(JSON.stringify(q));c.id=uid();return c})};
      if(o.d0===1)u.d0=1; if(LEN.test(u.ck))u.prec=pabCkName(u.ck); u.n=o.n||u.n})}
  n.ud.forEach(udPrefill);
  (f.pab||(f.pab=[])).push(n); PAB.cur[f.id]=n.id; PAB.ask=null; pendAPlanes(f); return n;
}
function hoyIso(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2)}

/* ───── copias en archivo ───── */
function copiaBlob(fs,sinPz){
  var ids={};fs.forEach(function(f){ids[f.id]=1});var rec={};Object.keys(S.recintos).forEach(function(k){if(ids[S.recintos[k].fincaId])rec[k]=S.recintos[k]});
  return new Blob([JSON.stringify({app:'plan-abonado',v:1,fecha:new Date().toISOString(),datos:{fincas:fs,recintos:rec,x:{asesor:S.x.asesor||{},abonos:(S.x.abonos||[]).map(function(a){if(!sinPz)return a;var c=Object.assign({},a);delete c.pz;return c})}}},null,1)],{type:'application/json'})}
function copiaTodo(){if(!S.fincas.length){toast('Aún no hay nada que guardar.');return}var t=Date.now();S.copia=t;S.fincas.forEach(function(f){f.cp=t});saveRaw();showDl(copiaBlob(S.fincas),'plan-abonado-copia-'+hoyIso()+'.json','Copia lista','Con este archivo recuperas tus fincas y planes en este u otro aparato («Abrir una copia»).');render()}
function copiaFinca(f){f.cp=Date.now();saveRaw();showDl(copiaBlob([f],1),'plan-abonado-'+safeName(f.nombre)+'-'+hoyIso()+'.json','Copia de la finca lista','Para pasar esta finca a otro aparato o a otra persona: «Abrir una copia».')}
var COPIA=null;
function copiaAbrir(file){var rd=new FileReader();rd.onload=function(){try{var o=JSON.parse(rd.result);if(!o||o.app!=='plan-abonado'||!o.datos||!Array.isArray(o.datos.fincas))throw 1;var D=o.datos;
    var nf0=D.fincas.length, ya=D.fincas.filter(function(f){return finca(f.id)}).length;
    sheet('<h2>Abrir copia</h2><p class="small">'+nf0+(nf0===1?' finca':' fincas')+': '+esc(D.fincas.map(function(f){return f.nombre}).join(', '))+'.</p>'+(ya?'<p class="small">'+ya+(ya===1?' ya está':' ya están')+' en este aparato: se cambian por las de la copia.</p>':'')+(function(){var ids={},n=0;D.fincas.forEach(function(f){ids[f.id]=1});Object.keys(D.recintos||{}).forEach(function(k){var r=S.recintos[k];if(r&&!ids[r.fincaId]&&finca(r.fincaId))n++});return n?'<p class="small">'+n+(n===1?' recinto de la copia ya está en otra finca de este aparato: si añades, se queda donde está.':' recintos de la copia ya están en otra finca de este aparato: si añades, se quedan donde están.')+'</p>':''})()+
      '<div class="acts"><button class="btn primary xl" data-act="ap-copia-join">Añadir a lo que hay</button>'+(S.fincas.length?'<button class="btn danger" data-act="ap-copia-repl">Cambiar todo por la copia</button>':'')+'<button class="btn" data-act="ap-close">Cancelar</button></div>');
    COPIA=D;
  }catch(e){toast('Ese archivo no es una copia de esta herramienta.')}};rd.readAsText(file)}
function copiaAplicar(repl){if(!COPIA)return;undoPush();
  if(repl){S.fincas=[];S.recintos={}}
  COPIA.fincas.forEach(function(f){S.fincas=S.fincas.filter(function(x){return x.id!==f.id});Object.keys(S.recintos).forEach(function(k){if(S.recintos[k].fincaId===f.id)delete S.recintos[k]});S.fincas.push(f)});
  var ids={},otra=0;COPIA.fincas.forEach(function(f){ids[f.id]=1});
  Object.keys(COPIA.recintos||{}).forEach(function(k){var r=S.recintos[k];if(r&&!ids[r.fincaId]&&finca(r.fincaId)){otra++;return}S.recintos[k]=COPIA.recintos[k]});
  var X=COPIA.x||{};
  if(X.asesor&&(repl||!Object.keys(S.x.asesor||{}).length))S.x.asesor=X.asesor;
  (X.abonos||[]).forEach(function(a){if(!(S.x.abonos||[]).some(function(b){return noAcc(b.n).toLowerCase()===noAcc(a.n).toLowerCase()}))(S.x.abonos=S.x.abonos||[]).push(a)});
  COPIA=null;S.copia=Date.now();save();sheetClose();UI.v='inicio';UI.fid=null;saveUI();render();undoToast('Copia abierta'+(otra?' · '+otra+(otra===1?' recinto ya estaba en otra finca de este aparato y se ha dejado allí':' recintos ya estaban en otra finca de este aparato y se han dejado allí'):''))}

/* ───── guardado automático en un archivo (Chrome y Edge en el ordenador) ───── */
var FS={h:null,ok:false,name:'',t:null,last:0,err:false};
function fsCan(){return typeof window.showSaveFilePicker==='function'}
function idbDo(mode,fn){return new Promise(function(res,rej){try{var q=indexedDB.open('plan-abonado',1);q.onupgradeneeded=function(){q.result.createObjectStore('kv')};q.onerror=function(){rej(q.error)};
  q.onsuccess=function(){var db=q.result,tx=db.transaction('kv',mode),st=tx.objectStore('kv'),r=fn(st);tx.oncomplete=function(){res(r&&r.result);db.close()};tx.onerror=function(){rej(tx.error)}}}catch(e){rej(e)}})}
function fsInit(){if(!fsCan()||!window.indexedDB)return;idbDo('readonly',function(st){return st.get('archivo')}).then(function(h){if(!h)return;FS.h=h;FS.name=h.name;
  return h.queryPermission({mode:'readwrite'}).then(function(s){FS.ok=s==='granted';renderHdr()})}).catch(function(){})}
function fsPick(){if(!fsCan())return;window.showSaveFilePicker({suggestedName:'plan-abonado.json',types:[{description:'Copia del plan de abonado',accept:{'application/json':['.json']}}]}).then(function(h){FS.h=h;FS.name=h.name;FS.ok=true;FS.err=false;
    idbDo('readwrite',function(st){return st.put(h,'archivo')}).catch(function(){});fsWrite(true);toast('Desde ahora cada cambio se guarda solo en «'+h.name+'»')}).catch(function(e){if(e&&e.name!=='AbortError')toast('No he podido usar ese archivo.')})}
function fsResume(){if(!FS.h)return;FS.h.requestPermission({mode:'readwrite'}).then(function(s){FS.ok=s==='granted';if(FS.ok)fsWrite(true);renderHdr()}).catch(function(){})}
function fsStop(){FS.h=null;FS.ok=false;FS.name='';idbDo('readwrite',function(st){return st.delete('archivo')}).catch(function(){});renderHdr();toast('Guardado automático quitado')}
function fsWrite(now){if(!FS.h||!FS.ok)return;clearTimeout(FS.t);FS.t=setTimeout(function(){
  FS.h.createWritable().then(function(w){return w.write(copiaBlob(S.fincas)).then(function(){return w.close()})}).then(function(){var t=Date.now();S.copia=t;FS.last=t;FS.err=false;S.fincas.forEach(function(f){f.cp=t});saveRaw();renderHdr();if(now===2)toast('Guardado en «'+FS.name+'»')})
    .catch(function(){FS.ok=false;FS.err=true;renderHdr();toast('No he podido guardar en «'+FS.name+'». Pulsa «Guardar copia» para volver a darle permiso.')})},now?10:1200)}

/* ───── ejemplo inventado (para probar) ───── */
function ejemplo(){
  undoPush(); var f=fincaNueva('Finca de ejemplo'); f.datos.tit={n:'Titular de ejemplo',nif:'00000000T',mun:'Municipio',prov:'Albacete'};
  addRecs(f,[{ref:'2:900:0:0:1:1:1',ha:5.2,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman',nombre:'La Hoya',pend:2.1},{ref:'2:900:0:0:1:2:1',ha:3.1,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman',pend:1.4},{ref:'2:900:0:0:3:7:2',ha:6.4,cultivo:'Almendro',riego:'secano',variedad:'Guara',pend:4.8}]);
  var p=planNuevo(f,pabY0());
  p.ud.forEach(function(u){u.mo={obj:'mantener'};
    u.su={f:hoyIso(),lab:'Laboratorio de ejemplo',inf:'0001',prof:'0-30',mo:1.2,ph:8.1,n:900,p:14,k:260,tex:'franca'};
    if(u.ck==='pistacho'){u.ne={mo:'cos',pr:1500,ed:1};u.ag={m3:2500,no3:20};u.pg={src:'own',rows:[{id:uid(),mo:'Marzo',n:'Nitrato potásico 13-0-46',d:80,rn:13,rp:0,rk:46,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Abril-junio',n:'Nitrato amónico 33,5 %',d:120,rn:33.5,rp:0,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Abril',n:'Fosfato monoamónico (MAP) 12-61-0',d:40,rn:12,rp:61,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'}]};mitPropuesta(u,true)}
    else{u.ne={mo:'cos',pr:600,ed:1};u.pg={src:'own',rows:[{id:uid(),mo:'Febrero-marzo',n:'Complejo 15-15-15',d:150,rn:15,rp:15,rk:15,fo:'Suelo',eq:'Abonadora'}]};mitPropuesta(u,true)}});
  UI.v='finca';UI.fid=f.id;UI.step='rec'; save(); saveUI(); window.__scrollTop=true; render(); undoToast('Finca de ejemplo creada (datos inventados)');
}

/* ───── ayudas agronómicas: medidas propuestas, meses, programa propuesto, arreglos del balance, semáforo del suelo ───── */
var MESES=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
/* medidas del anexo V que encajan con el riego y los abonos de la unidad */
function mitPropuesta(u,silencio){
  var rows=(u.pg||{}).rows||[], fert=rows.some(function(q){return /fertirr/i.test(q.fo||'')})||(u.rg==='regadio'&&!rows.length), urea=rows.some(function(q){return /urea|n32|n20|soluci[oó]n nitrog/i.test(noAcc(q.n||''))}), org=rows.some(function(q){return q.org===1});
  var m=u.mit=u.mit||{}, add=function(k){if(!m[k]||m[k].a!==1){m[k]={a:1};return 1}return 0}, n=0;
  n+=add('B1'); n+=add('O1'); if(fert)n+=add('B4'); if(urea&&!fert)n+=add('B2'); if(org)n+=add('A9');
  return n}
/* meses de un momento escrito («Marzo», «Abril-junio», «oct. a dic.», «Febrero y marzo»): lista de números 1-12 */
function mesesDe(s){var T=(noAcc(String(s||'')).toLowerCase().match(/[a-z]+|[-–]/g)||[]), K=MESES.map(function(m){return noAcc(m).toLowerCase()});
  var mes=function(t){if(t.length<3)return 0;for(var i=0;i<12;i++)if(K[i].indexOf(t)===0)return i+1;return 0};
  var o=[];for(var j=0;j<T.length;j++){var m=mes(T[j]);if(!m)continue;
    if(j+2<T.length&&/^(-|–|a|hasta|al)$/.test(T[j+1])&&mes(T[j+2])){var e=mes(T[j+2]),m0=m;for(var g=0;g<12;g++){o.push(m0);if(m0===e)break;m0=m0%12+1}j+=2;continue}
    o.push(m)}
  return o.filter(function(v,i,a){return a.indexOf(v)===i})}
pabMeses=function(s){return mesesDe(s)};
function rango(a,b){return MESES[a-1]+(b&&b!==a?'-'+MESES[b-1].toLowerCase():'')}
/* U.F./ha de cada mes de la campaña (noviembre → octubre) */
function calMeses(u,C){var M={};for(var i=1;i<=12;i++)M[i]={n:0,p:0,k:0,L:[]};var sin=0;
  (C.rows||[]).forEach(function(o){var ms=mesesDe(o.mo);if(!ms.length){if(o.d>0)sin++;return}ms.forEach(function(m){['n','p','k'].forEach(function(k){M[m][k]+=o.uf[k]/ms.length});M[m].L.push(o.n||'Aporte')})});
  return {M:M,sin:sin}}
function perMeses(u){if(u.rg==='regadio'&&u.loc!==0)return [];return PAB_PER[u.ck]||[]}

/* ───── programa propuesto con el catálogo, repartido por meses y sin N en los meses prohibidos ───── */
var r5=function(x){return Math.max(5,Math.round(x/5)*5)};
function catN(n){return catGen().filter(function(x){return x.n===n})[0]}
function propuesta(f,u,C){
  var B=C.B, need={n:B.n.net,p:B.p.net,k:B.k.net}, rg=u.rg==='regadio', cer=/^(trigo|trigoduro|cebada|avena|centeno|triticale|maiz)$/.test(u.ck), rows=[];
  if(!(need.n>1||need.p>1||need.k>1))return rows;
  var fo=rg?'Fertirrigación':'Suelo', eq=rg?'Cabezal de riego':'Abonadora';
  var put=function(nm,dose,mo){var x=catN(nm);if(!x||!(dose>0))return;var d=r5(dose),mx=Math.max(x.rn,x.rp,x.rk)/100;if(Math.abs(d-dose)*mx>0.9)d=Math.max(1,Math.round(dose));rows.push({id:uid(),mo:mo,n:x.n,d:d,rn:x.rn,rp:x.rp,rk:x.rk,fo:fo,eq:eq});need.n-=d*x.rn/100;need.p-=d*x.rp/100;need.k-=d*x.rk/100};
  var M=rg?{p:'Marzo-abril',k:'Mayo-julio',n:'Marzo-junio'}:(cer?{f:'Octubre',n:'Febrero-marzo'}:(u.ck==='vid'?{f:'Marzo',n:'Abril'}:{f:'Febrero',n:'Marzo'}));
  if(rg){
    if(need.p>1)put('Fosfato monoamónico (MAP) 12-61-0',need.p/0.61,M.p);
    if(need.k>1){var dk=need.k/0.46; if(dk*0.13<=Math.max(0,need.n)+2)put('Nitrato potásico 13-0-46',dk,M.k); else put('Sulfato potásico 50 %',need.k/0.5,M.k)}
    if(need.n>2)put('Nitrato amónico 33,5 %',need.n/0.335,M.n);
  }else{
    /* fondo con un complejo que cubra el fósforo sin pasarse de N ni de K */
    if(need.p>1){var cs=['Complejo 8-24-8','Complejo 8-15-15','Complejo 9-18-27','Complejo 15-15-15'].map(catN), best=null,bs=1e9;
      cs.forEach(function(x){var d=need.p/(x.rp/100),en=d*x.rn/100-Math.max(0,need.n),ek=d*x.rk/100-Math.max(0,need.k),sc=Math.max(0,en)*2+Math.max(0,ek)+Math.abs(Math.min(0,ek))*0.5;if(sc<bs){bs=sc;best=x}});
      put(best.n,need.p/(best.rp/100),M.f)}
    if(need.k>1)put('Sulfato potásico 50 %',need.k/0.5,M.f);
    if(need.n>2)put(cer?'Nitrato amónico cálcico (NAC) 27 %':'Nitrato amónico cálcico (NAC) 27 %',need.n/0.27,M.n);
  }
  /* ningún aporte con N en los meses prohibidos del anexo II */
  var pm=perMeses(u); rows.forEach(function(q){if(!(+q.rn>0))return;var ms=mesesDe(q.mo);if(ms.some(function(m){return pm.indexOf(m)>=0}))q.mo=rg?'Marzo':'Febrero'});
  return rows}
/* arreglo de cada desviación del balance: qué abono y cuánto, o qué aporte bajar */
function arreglo(f,u,C,k){var b=C.B[k], rg=u.rg==='regadio', nm={n:rg?'Nitrato amónico 33,5 %':'Nitrato amónico cálcico (NAC) 27 %',p:rg?'Fosfato monoamónico (MAP) 12-61-0':'Superfosfato triple 46 %',k:'Sulfato potásico 50 %'}[k], R={n:'rn',p:'rp',k:'rk'}[k];
  if(b.st==='Ajustado'||C.ex||u.d0===1)return null;
  if(b.dv<0){var x=catN(nm), d0=-b.dv/(x[R]/100), d=r5(d0);if(Math.abs(d-d0)*x[R]/100>0.9)d=Math.max(1,Math.round(d0));return {t:'add',txt:'unos '+nf(d,0)+' kg/ha de '+x.n,q:{id:uid(),mo:k==='n'?(rg?'Abril':'Marzo'):(k==='p'?(rg?'Marzo':'Febrero'):(rg?'Junio':'Febrero')),n:x.n,d:d,rn:x.rn,rp:x.rp,rk:x.rk,fo:rg?'Fertirrigación':'Suelo',eq:rg?'Cabezal de riego':'Abonadora'}}}
  var rows=(u.pg||{}).rows||[], best=-1, bv=0; rows.forEach(function(q,j){var v=(+q.d||0)*(+q[R]||0)/100;if(v>bv){bv=v;best=j}});
  if(best<0)return null; var q=rows[best], nx=(+q.d||0)-b.dv/((+q[R]||0)/100), nd=Math.max(0,Math.round(nx/5)*5);if(Math.abs(nd-nx)*(+q[R]||0)/100>0.9)nd=Math.max(0,Math.round(nx));if(nd===(+q.d||0))return null;
  return {t:'set',j:best,d:nd,txt:'baja «'+(q.n||'aporte '+(best+1))+'» de '+nf(+q.d||0,0)+' a '+nf(nd,0)+' kg/ha'}}

/* ───── semáforo del suelo (niveles orientativos, criterio técnico) ───── */
var SUELO=[
  ['mo','Materia orgánica','%',[0.8,1.5,2.5,3.5],2],
  ['p','Fósforo (Olsen)','mg/kg',[7,12,20,35],0],
  ['k','Potasio','mg/kg',[100,175,300,450],0],
  ['ph','pH','',[5.5,6.5,7.5,8.5],1]];
var NIV=['Muy bajo','Bajo','Medio','Alto','Muy alto'], NIVPH=['Muy ácido','Ácido','Neutro','Básico','Muy básico'];
function nivel(v,cut){for(var i=0;i<cut.length;i++)if(v<cut[i])return i;return 4}
/* crédito del suelo propuesto: alto = 25 % de la necesidad; muy alto = 50 % (P₂O₅ y K₂O) */
function creditoSuelo(u,C){var su=u.su||{},o={};[['p','p'],['k','k']].forEach(function(x){var v=su[x[0]];if(!abHas(v))return;var l=nivel(+v,SUELO.filter(function(s){return s[0]===x[0]})[0][3]);if(l>=3)o[x[1]]=Math.round(C.nb[x[1]]*(l===4?0.5:0.25))});return o}

/* ───── ayudas (?) con el artículo de la norma ───── */
var GLOS={
  ud:['Unidad de producción','Conjunto de recintos con el mismo cultivo y el mismo régimen (secano o regadío) que se abonan igual. El plan se hace por unidad.','RD 1051/2022, art. 3 v) y art. 6'],
  ex:['Unidad exenta','Las unidades de secano de 10 ha o menos no necesitan plan de abonado. Se pueden rellenar igualmente.','RD 1051/2022, art. 4.2 b)'],
  zv:['Zona vulnerable a nitratos','Zonas declaradas por la comunidad autónoma por el riesgo de contaminación por nitratos. En ellas hay un límite de N por cultivo y más restricciones. Consulta el mapa de zonas vulnerables de tu comunidad.','Decreto 80/2025 de Castilla-La Mancha, anejo III'],
  tp:['Tipo de parcela','Tipo 1: suelo ligero (arenoso, franco-arenoso o franco) o precedente leguminosa. Tipo 2: el resto, y también si se fertirriga. Cambia el límite de N.','Decreto 80/2025 de Castilla-La Mancha, anejo III'],
  rp:['Recinto representativo','El recinto cuyo análisis de suelo vale para toda la unidad.','RD 1051/2022, art. 6.2'],
  uf:['U.F. (unidades fertilizantes)','Kilos de nutriente puro por hectárea: dosis (kg/ha) × riqueza (%) ÷ 100. Por ejemplo, 100 kg/ha de un 15-15-15 son 15 U.F. de N, de P₂O₅ y de K₂O.',''],
  net:['Necesidad neta','Lo que necesita el cultivo (necesidad bruta) menos lo que ya aportan el suelo, el agua de riego, el estiércol y otros. Es lo que tiene que cubrir el programa.','RD 1051/2022, art. 5'],
  suelo:['Aporte del suelo','Lo que el suelo puede dar según el análisis. Si no se pone, cuenta 0 y no se descuenta nada (es lo prudente). Con niveles altos de P o K se puede descontar una parte: es criterio técnico.','RD 1051/2022, art. 5.1'],
  prec:['Cultivo precedente','El cultivo de la campaña anterior en esos recintos. En leñosos es el mismo cultivo.','RD 1051/2022, art. 5.1 a)'],
  mo:['Materia orgánica','El plan debe buscar mantener o aumentar la materia orgánica del suelo.','RD 1051/2022, art. 4.3'],
  pg:['Programa previsto','Abonos que se piensan aplicar: momento, producto, dosis, forma de aplicación y maquinaria.','RD 1051/2022, art. 6.4'],
  mit:['Medidas del anexo V','Medidas para reducir las emisiones de amoniaco y gases de efecto invernadero. Con urea hace falta al menos una de la parte B; con estiércoles sin compostar, una de la parte A.','RD 1051/2022, art. 6.5, 10.4 y 12.3, anexo V'],
  res:['Restricciones de aplicación','Situaciones en las que no se puede abonar (suelo helado o encharcado, aviso rojo por lluvia, arrastre a cauces…).','RD 1051/2022, art. 4.4; Decreto 80/2025, apdo. 4.3'],
  per:['Meses sin abono nitrogenado','Meses en los que no se puede aplicar N según el cultivo. No se aplica con riego localizado ni con agricultura de precisión.','RD 1051/2022, anexo II'],
  as:['Asesor en fertilización','Técnico inscrito en el REGFER. Obligatorio desde el 1 de septiembre de 2027 en zona vulnerable y desde el 1 de septiembre de 2028 en el resto.','RD 1051/2022, art. 6.6'],
  cl:['Lluvia de la zona','Precipitación media anual y por estaciones. Se pone la de la estación principal de AEMET más cercana a los recintos; cámbiala si tienes datos mejores.','RD 1051/2022, art. 6.3'],
  pr:['Producción objetivo','Cosecha que esperas en kg/ha. Con las extracciones por tonelada da la necesidad del cultivo.','RD 1051/2022, art. 5.1 a)'],
  lim:['Límite de N','Máximo de N por hectárea y año en zona vulnerable, según cultivo, riego y tipo de parcela. Con inhibidores se puede subir un 10 %.','Decreto 80/2025 de Castilla-La Mancha, anejo III'],
  pend:['Pendiente','En zona vulnerable no se abona con N ni P en recintos de más del 15 % de pendiente, y con más del 2 % hay que tomar precauciones.','Decreto 80/2025 de Castilla-La Mancha, apdo. 4.3.2']};
function hlp(k){return '<button type="button" class="hlp" data-act="ap-help" data-k="'+k+'" aria-label="Qué es esto">?</button>'}

/* ───── comprobación en vivo de cifras raras (se marcan en rojo suave, no se bloquea) ───── */
var VAL=[
  [/\.pg\.rows\.\d+\.d$/,0,3000,'Dosis muy alta para un abono mineral: revisa que esté en kg/ha',function(q){return q&&q.org!==1}],
  [/\.pg\.rows\.\d+\.d$/,0,80000,'Dosis muy alta: revisa que esté en kg/ha'],
  [/\.pg\.rows\.\d+\.r[npk]$/,0,100,'Una riqueza no puede pasar del 100 %'],
  [/\.su\.ph$/,3,10,'Un pH de suelo está entre 4 y 9'],[/\.su\.mo$/,0,15,'Materia orgánica muy alta: revisa que esté en %'],
  [/\.su\.p$/,0,300,'Fósforo Olsen muy alto: revisa que esté en mg/kg'],[/\.su\.k$/,0,3000,'Potasio muy alto: revisa que esté en mg/kg'],
  [/\.ne\.pr$/,0,60000,'Producción muy alta: revisa que esté en kg/ha'],[/\.ag\.m3$/,0,15000,'Dotación muy alta: revisa que esté en m³/ha'],
  [/\.ag\.no3$/,0,500,'Nitratos del agua muy altos: revisa que esté en mg/L'],[/\.rc\.[\d:]+\.pe$/,0,100,'Una pendiente va de 0 a 100 %'],
  [/^cl\.[aoipv]$/,0,2500,'Lluvia muy alta: revisa que esté en mm']];
function valida(path,v,q){if(v===''||v==null||!VAL.some(function(x){return x[0].test(path)}))return '';var n=typeof v==='number'?v:numEs(String(v));if(n===null||isNaN(n))return 'No es un número (los decimales con coma)';
  for(var i=0;i<VAL.length;i++){var x=VAL[i];if(x[0].test(path)&&(!x[4]||x[4](q))&&(n<x[1]||n>x[2]))return x[3]}return ''}
function filaDe(path){var m=/^ud\.(\d+)\.pg\.rows\.(\d+)\./.exec(path);if(!m)return null;var f=finca(UI.fid),p=f&&pabCur(f),u=p&&p.ud[+m[1]];return u&&((u.pg||{}).rows||[])[+m[2]]}
function marcaValida(root){$$('input.pa-in[data-p],input.ap-in[data-num]',root).forEach(function(i){var path=i.dataset.p||'';var w=valida(path,i.value,filaDe(path),i.getAttribute('inputmode')==='decimal');i.classList.toggle('rare',!!w);if(w)i.title=w;else if(i.title&&i.classList.contains('was-rare'))i.removeAttribute('title');i.classList.toggle('was-rare',!!w)})}

/* ───── lo que falta en cada paso (para el ✓ / «faltan N» y para «Ir a lo que falta») ───── */
function faltas(f,p,Cs){var L=[], add=function(s,t,sel,u){L.push({s:s,t:t,sel:sel||'',u:u==null?null:u})}, rs=fincaRecs(f.id), t=(f.datos||{}).tit||{};
  if(!rs.length)add('rec','Añadir los recintos','[data-act="ap-addrec"]');
  rs.forEach(function(r){if(!r.cultivo)add('rec','Cultivo del recinto '+refLabel(r.ref),'[data-r="'+r.ref+'"][data-k="cultivo"]')});
  if(!t.n)add('rec','Nombre del titular','[data-e="tit.n"]'); if(!t.nif)add('rec','NIF del titular','[data-e="tit.nif"]');
  if(!p){if(rs.length)add('ud','Crear el plan de la campaña','[data-act="pa-new-ok"]');return L}
  if(!abHas((p.cl||{}).a))add('ud','Lluvia media anual','[data-p="cl.a"]');
  (p.ud||[]).forEach(function(u,i){var C=Cs[i],P='ud.'+i+'.';if(C.ex)return; var d0=u.d0===1, su=u.su||{}, ne=u.ne||{}, rows=(u.pg||{}).rows||[];
    if(!C.rs.length)add('ud','«'+u.n+'» no tiene recintos','#u'+i,i);
    if(!d0){if(!['mo','n','p','k','ph'].some(function(k){return abHas(su[k])}))add('su','Análisis de suelo de «'+u.n+'»','[data-p="'+P+'su.mo"]',i);
      if(!u.rp)add('ud','Recinto representativo de «'+u.n+'»','#u'+i,i);
      if(u.rg==='regadio'&&!abHas((u.ag||{}).m3)&&!abHas((u.ag||{}).max))add('su','Dotación de riego de «'+u.n+'»','[data-p="'+P+'ag.m3"]',i)}
    if((ne.mo||'cos')==='cos'&&!abHas(ne.pr))add('ne','Producción objetivo de «'+u.n+'»','[data-p="'+P+'ne.pr"]',i);
    if(!u.prec)add('ne','Cultivo precedente de «'+u.n+'»','[data-p="'+P+'prec"]',i);
    if(!(u.mo&&u.mo.obj))add('ne','Objetivo de materia orgánica de «'+u.n+'»','[data-p="'+P+'mo.obj"]',i);
    if(!d0){if(!rows.length)add('pg','Programa de abonos de «'+u.n+'»','[data-act="ap-prop"][data-u="'+i+'"]',i);
      rows.forEach(function(q,j){var R=P+'pg.rows.'+j+'.';if(!q.n)add('pg','Abono del aporte '+(j+1)+' de «'+u.n+'»','[data-p="'+R+'n"]',i);else if(!(+q.d>0))add('pg','Dosis del aporte '+(j+1)+' de «'+u.n+'»','[data-p="'+R+'d"]',i);else if(!q.mo)add('pg','Mes del aporte '+(j+1)+' de «'+u.n+'»','[data-p="'+R+'mo"]',i)});
      if(!PAB_MIT.some(function(x){return ((u.mit||{})[x[0]]||{}).a===1}))add('pg','Medidas del anexo V de «'+u.n+'»','#u'+i+'-mit',i);
      if(!PAB_RES.every(function(x){return ((u.res||{})[x[0]]||{}).a===1}))add('pg','Restricciones sin comprobar en «'+u.n+'»','#u'+i+'-res',i);
      if(u.zv&&!C.lok)add('pg','«'+u.n+'» pasa del límite de N','#u'+i+'-tot',i);
      if(pabPerHits(f,u,p,C).length)add('pg','«'+u.n+'» lleva N en meses prohibidos','#u'+i+'-tot',i)}});
  if(!pabIso(p.fe))add('rv','Fecha de elaboración del plan','[data-p="fe"]');
  var nec=(p.ud||[]).some(function(u,i){if(Cs[i].ex)return false;var ob=u.zv?'2027-09-01':'2028-09-01';return pabCampRange(p)[1]>=ob}), A=pabAs();
  if(nec&&!(A.n&&A.regfer))add('rv','Asesor en fertilización (nombre y nº REGFER)','[data-p="AS.n"]');
  return L}

/* ───── lista de la compra y coste (solo en pantalla) ───── */
function compra(f,p,Cs){var m={},o=[];(p.ud||[]).forEach(function(u,i){var C=Cs[i];if(u.d0===1)return;((u.pg||{}).rows||[]).forEach(function(q){var d=+q.d||0;if(!q.n||!(d>0))return;var k=noAcc(q.n).toLowerCase().trim();
    if(!m[k]){var c=catDe(q.n);m[k]={n:q.n,kg:0,eur:0,sinPz:0,dens:c&&c.dens||0,uds:{}};o.push(m[k])}var kg=d*C.ha;m[k].kg+=kg;m[k].uds[u.n]=1;if(abHas(q.pz))m[k].eur+=kg*q.pz;else m[k].sinPz=1})});
  return o.sort(function(a,b){return b.kg-a.kg})}
function costeUd(u,C){var e=0,sin=0;((u.pg||{}).rows||[]).forEach(function(q){if(!(+q.d>0))return;if(abHas(q.pz))e+=q.d*q.pz;else sin=1});return {ha:e,tot:e*C.ha,sin:sin}}

/* ───── iconos (trazo, del color del texto) ───── */
var ICP={back:'M15 18l-6-6 6-6',undo:'M9 14L4 9l5-5M4 9h10a6 6 0 010 12h-3',save:'M5 3h11l3 3v15H5zM8 3v6h8M8 21v-7h8v7',pdf:'M14 3H6v18h12V7zM14 3v4h4M9 13h6M9 17h6',xls:'M4 4h16v16H4zM4 10h16M4 15h16M10 4v16',
  map:'M9 4L3 6v14l6-2 6 2 6-2V4l-6 2zM9 4v14M15 6v14',plus:'M12 5v14M5 12h14',cal:'M4 6h16v14H4zM4 10h16M8 3v5M16 3v5',check:'M5 12l5 5 9-10',alert:'M12 4l9 16H3zM12 10v4M12 17v.5',
  search:'M11 4a7 7 0 100 14 7 7 0 000-14zM20 20l-4-4',folder:'M3 6h6l2 2h10v11H3z',up:'M12 16V4M7 9l5-5 5 5M4 20h16',leaf:'M5 19C5 11 11 5 19 5c0 8-6 14-14 14zM5 19l8-8',trash:'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13',
  spark:'M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5L18 18M6 18l2.5-2.5M15.5 8.5L18 6',copy:'M8 8h12v12H8zM4 16V4h12',next:'M9 6l6 6-6 6',close:'M6 6l12 12M18 6L6 18',loc:'M12 3v3M12 18v3M3 12h3M18 12h3M12 8a4 4 0 100 8 4 4 0 000-8z',
  list:'M8 6h12M8 12h12M8 18h12M4 6h.5M4 12h.5M4 18h.5',drop:'M12 3c4 5 6 8 6 11a6 6 0 01-12 0c0-3 2-6 6-11z',euro:'M17 6a7 7 0 100 12M4 10h9M4 14h9'};
function ic(n){return '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path d="'+(ICP[n]||'')+'"/></svg>'}

var STEPS=[['rec','Recintos','La finca, el titular y los recintos SIGPAC.'],['ud','Unidades','Recintos juntados por cultivo y riego, zona vulnerable y lluvia de la zona.'],['su','Suelo y agua','Análisis de suelo del recinto representativo y, en regadío, el agua de riego.'],
  ['ne','Cosecha','Producción esperada, lo que ya aportan el suelo, el estiércol u otros, y la materia orgánica.'],['pg','Programa','Abonos previstos, medidas del anexo V y restricciones de aplicación.'],['rv','Revisar y PDF','Balance, comprobación del plan, firmas y el plan en PDF o Excel.']];
var CTX=null, SEL={}, RM={};
function isWide(){return window.matchMedia&&window.matchMedia('(min-width:1200px)').matches}
function stepIdx(){for(var i=0;i<STEPS.length;i++)if(STEPS[i][0]===UI.step)return i;return 0}
function ctxOf(f){var p=pabCur(f),Cs=p?(p.ud||[]).map(function(u){return pabCalc(f,u,p)}):[];return {p:p,Cs:Cs,F:faltas(f,p,Cs)}}
function nFalta(ctx,s){return ctx.F.filter(function(x){return x.s===s}).length}
function stCls(st){return st==='Ajustado'?'ok':(st==='Déficit'?'warn':(st==='Exenta de plan'||st==='Sin abonado'?'ex':'bad'))}
function stChip(st){var c=stCls(st);return '<span class="pa-st pa-'+c+'">'+(c==='ok'?'✓ ':c==='warn'?'▼ ':c==='bad'?'▲ ':'')+esc(st)+'</span>'}

/* ───── dibujo general ───── */
function render(){
  var f=UI.v==='finca'?finca(UI.fid):null; if(!f){UI.v='inicio';UI.fid=null}
  EXP.fid=UI.fid; var B=document.body; B.classList.toggle('vf',!!f); B.classList.toggle('vi',!f);
  var sy=window.scrollY;
  if(f){pabSync(f);CTX=ctxOf(f);
    $('#side').innerHTML=sideHTML(f,CTX); $('#app').innerHTML=fixHTML(fincaHTML(f,CTX)); $('#bal').innerHTML=balHTML(f,CTX); $('#bar').innerHTML=barHTML(); $('#bar').hidden=false}
  else{CTX=null;$('#side').innerHTML='';$('#bal').innerHTML='';$('#app').innerHTML=inicioHTML();$('#bar').hidden=true;$('#bar').innerHTML=''}
  renderHdr();
  document.title=f?f.nombre+' · Plan de abonado':'Plan de abonado';
  if(window.__scrollTop){window.scrollTo(0,0);window.__scrollTop=false}else window.scrollTo(0,sy);
  stkLab(); inThou(); marcaValida(); chipCentro(); prevAuto();
}
function renderHdr(){var el=$('#hdr');if(!el)return;var f=UI.v==='finca'?finca(UI.fid):null,p=f&&pabCur(f),sinCopia=S.fincas.length&&S.mod>(S.copia||0)+1000;
  var h='<div class="hl">'+(f?'<button class="hback" data-act="ap-home" aria-label="Volver a mis fincas">'+ic('back')+'</button>':'')+'<button class="brand" data-act="ap-home" title="Mis fincas">'+ic('leaf')+'<b>Plan de abonado</b></button>';
  if(f){h+='<span class="sep">/</span><span class="fn" title="'+esc(f.nombre)+'">'+esc(f.nombre)+'</span>';
    var l=pabList(f);if(l.length)h+='<select class="trsel sm ap-camp" aria-label="Campaña">'+l.map(function(x){return '<option value="'+esc(x.id)+'"'+(p&&x.id===p.id?' selected':'')+'>'+(isDesk()?'Campaña ':'')+esc(pabCamp(x.y))+'</option>'}).join('')+'<option value="new">+ Plan de otra campaña…</option></select>'}
  h+='</div><div class="hr">';
  if(FS.ok)h+='<span class="fs" title="Cada cambio se guarda solo en ese archivo">'+ic('check')+'Guardado en '+esc(FS.name)+(FS.last?' · '+new Date(FS.last).toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'}):'')+'</span>';
  else if(FS.h)h+='<button class="btn sm warnb" data-act="ap-fs-res" title="El navegador pide permiso otra vez para escribir en el archivo">'+ic('save')+'Reactivar guardado en '+esc(FS.name)+'</button>';
  h+='<button class="btn sm lk" data-act="ap-undo"'+(UNDO.length?'':' disabled')+' title="Deshacer (Ctrl+Z)">'+ic('undo')+'<span>Deshacer</span></button>';
  if(S.fincas.length&&!FS.ok)h+='<button class="btn sm'+(sinCopia?' warnb':'')+'" data-act="ap-copias" title="Guardar copia (Ctrl+S)">'+ic('save')+'<span>'+(sinCopia?'Guardar copia':'Copias')+'</span></button>';
  if(f&&p)h+='<button class="btn sm primary" data-act="ap-pdf" title="PDF del plan (Ctrl+P)">'+ic('pdf')+'<span>PDF</span></button><button class="btn sm" data-act="ap-xlsx">'+ic('xls')+'<span>Excel</span></button>';
  el.innerHTML=h+'</div>'}
function barHTML(){var i=stepIdx();return '<button class="btn" data-act="ap-prev">'+(i===0?'‹ Fincas':'‹ Atrás')+'</button>'+(i<STEPS.length-1?'<button class="btn primary" data-act="ap-next">Siguiente ›</button>':'<button class="btn primary" data-act="ap-pdf">PDF del plan</button>')}
function chipCentro(){var s=$('.steps'),c=s&&$('.chip.on',s);if(!s||!c)return;var x=c.offsetLeft-(s.clientWidth-c.offsetWidth)/2;s.scrollLeft=Math.max(0,x)}

/* columna izquierda: pasos con su estado y accesos de la finca */
function sideHTML(f,ctx){var p=ctx.p,y1=p?+p.y+1:pabY0(),ex=p&&(f.pab||[]).some(function(x){return +x.y===y1});
  var h='<div class="sbox"><div class="k">Pasos</div><ol class="snav">'+STEPS.map(function(s,i){var n=nFalta(ctx,s[0]),dis=!p&&i>1;
    return '<li><button class="sn'+(s[0]===UI.step?' on':'')+'" data-act="ap-step" data-s="'+s[0]+'"'+(dis?' disabled':'')+'><span class="no">'+(i+1)+'</span><span class="nm">'+s[1]+'</span>'+(dis?'':n?'<span class="st warn">faltan '+n+'</span>':'<span class="st ok">'+ic('check')+'</span>')+'</button></li>'}).join('')+'</ol>';
  h+=ctx.F.length?'<button class="btn sm wide" data-act="ap-falta">'+ic('alert')+'Ir a lo que falta ('+ctx.F.length+')</button>':(p?'<div class="okall">'+ic('check')+'Plan completo</div>':'');
  h+='</div><div class="sbox"><div class="k">Esta finca</div>'+
    '<button class="sl" data-act="ap-titular">Titular y finca'+ic('next')+'</button>'+
    '<button class="sl" data-act="ap-mapa">Ver y añadir en el mapa'+ic('next')+'</button>'+
    (p?'<button class="sl" data-act="ap-sigcamp">'+(ex?'Ir al plan de '+pabCamp(y1):'Hacer el plan de '+pabCamp(y1))+ic('next')+'</button>':'')+
    '<button class="sl" data-act="ap-mis">Mis abonos'+ic('next')+'</button>'+
    '<button class="sl" data-act="ap-copiaf">Copia de esta finca'+ic('next')+'</button></div>';
  h+='<div class="kbd small">Atajos: <b>Ctrl+S</b> copia · <b>Ctrl+Z</b> deshacer · <b>Ctrl+P</b> PDF · <b>Alt+←/→</b> pasos</div>';
  return h}

function fincaHTML(f,ctx){
  var p=ctx.p, i0=stepIdx(), st=STEPS[i0], h='';
  h+='<div class="steps" role="tablist">'+STEPS.map(function(s,i){var n=nFalta(ctx,s[0]);return '<button type="button" class="chip'+(s[0]===UI.step?' on':'')+'" data-act="ap-step" data-s="'+s[0]+'" role="tab" aria-selected="'+(s[0]===UI.step)+'"><b>'+(i+1)+'</b>'+s[1]+(p||i<2?(n?'<em class="dw">'+n+'</em>':'<em class="dk">✓</em>'):'')+'</button>'}).join('')+'</div>';
  var nF=nFalta(ctx,st[0]);
  h+='<div class="sh"><div><div class="k">Paso '+(i0+1)+' de 6</div><h2>'+st[1]+(UI.step==='ud'?hlp('ud'):'')+'</h2><div class="small">'+st[2]+'</div></div>'+
    (nF?'<button class="falta" data-act="ap-falta" data-s="'+st[0]+'" title="'+esc(ctx.F.filter(function(x){return x.s===st[0]}).map(function(x){return x.t}).join('\n'))+'">'+ic('alert')+'Faltan '+nF+' · Ir a lo que falta</button>':((p||i0===0)?'<span class="okst">'+ic('check')+'Completo</span>':''))+'</div>';
  if(UI.step==='rec')h+=recHTML(f,ctx);
  else if(!p)h+=nuevoPlanHTML(f,true);
  else{
    if(PAB.ask===f.id)h+=nuevoPlanHTML(f,false);
    var Cs=ctx.Cs, ord=udsOrden(p,Cs);
    if(UI.step==='ud')h+=udStepHTML(f,p,Cs,ord);
    else if(UI.step==='rv')h+=rvStepHTML(f,p,Cs,ord);
    else{ord.forEach(function(i){h+=udBox(f,p,p.ud[i],Cs[i],i,UI.step)});if(!(p.ud||[]).length)h+='<div class="note">No hay unidades: añádelas en el paso 2.</div>'}
  }
  h+='<div class="snext">'+(i0>0?'<button class="btn" data-act="ap-prev">‹ '+STEPS[i0-1][1]+'</button>':'<button class="btn" data-act="ap-home">‹ Mis fincas</button>')+(i0<5?'<button class="btn primary" data-act="ap-next">Siguiente: '+STEPS[i0+1][1]+' ›</button>':'<button class="btn primary" data-act="ap-pdf">'+ic('pdf')+'PDF del plan</button>')+'</div>';
  return h}
function udsOrden(p,Cs){return (p.ud||[]).map(function(u,i){return i}).sort(function(a,b){return (Cs[a].ex?1:0)-(Cs[b].ex?1:0)||a-b})}
function nuevoPlanHTML(f,first){
  var y0=pabY0(), ex={}; (f.pab||[]).forEach(function(x){ex[+x.y]=1}); var yo=[]; for(var yy=y0-2;yy<=y0+2;yy++)if(!ex[yy])yo.push(yy);
  var ys=yo.indexOf(y0)>=0?y0:(yo.filter(function(v){return v>y0})[0]||yo[yo.length-1]), n=fincaRecs(f.id).length;
  if(!n)return '<div class="note">Primero mete los recintos en el paso 1.</div>';
  return '<div class="card"><h3>'+(first?'Crear el plan de abonado':'Plan de otra campaña')+'</h3><p class="small" style="margin:6px 0 0">Los '+n+' recintos se juntan solos en unidades de producción por cultivo y riego. '+((f.pab||[]).length?'Del plan anterior se copian el suelo, el agua, la cosecha, el programa y las medidas: revísalos.':'Se marcan como propuesta las medidas del anexo V y las restricciones habituales.')+'</p>'+
    '<label class="lbl" for="pa-ny">Campaña (de noviembre a octubre)</label><select id="pa-ny" class="trsel">'+yo.map(function(v){return '<option value="'+v+'"'+(v===ys?' selected':'')+'>'+pabCamp(v)+'</option>'}).join('')+'</select>'+
    '<div class="acts"><button class="btn primary xl" data-act="pa-new-ok">Crear el plan</button>'+(first?'':'<button class="btn" data-act="ap-newplan-no">Cancelar</button>')+'</div></div>';
}

/* ───── paso 1: titular y recintos ───── */
function inp(attr,val,ph,mode,cls){return '<input class="in ap-in'+(cls?' '+cls:'')+'" '+attr+(mode?' inputmode="'+mode+'"':'')+' value="'+esc(val==null?'':(typeof val==='number'?abEd(val):val))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off">'}
function recOrden(rs){var k=(UI.sort||'ref').replace(/^-/,''),d=/^-/.test(UI.sort||'')?-1:1;
  return rs.slice().sort(function(a,b){var v=k==='ha'?haC(a)-haC(b):k==='cultivo'?String(a.cultivo||'~').localeCompare(String(b.cultivo||'~'),'es'):k==='riego'?String(a.riego||'').localeCompare(String(b.riego||'')):0;return (v||a.ref.localeCompare(b.ref,'es',{numeric:true}))*d})}
function thS(k,t){var on=(UI.sort||'ref').replace(/^-/,'')===k,d=/^-/.test(UI.sort||'');return '<th><button class="ths'+(on?' on':'')+'" data-act="ap-sort" data-k="'+k+'">'+t+(on?(d?' ▾':' ▴'):'')+'</button></th>'}
function recHTML(f,ctx){
  var t=(f.datos||{}).tit||{}, rs=recOrden(fincaRecs(f.id)), h='', nsel=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id}).length;
  var op=UI.titOpen||!t.n||!t.nif;
  h+='<details class="more pa-box tit" id="titular" data-pa-o="tit"'+(op?' open':'')+'><summary><b>Titular y finca</b><span class="small">'+esc([f.nombre,t.n||'sin titular',t.nif].filter(Boolean).join(' · '))+'</span></summary><div class="pa-g">';
  h+='<div>'+pabF('Nombre de la finca',inp('data-e="nombre"',f.nombre))+'</div>';
  [['n','Titular (nombre o razón social)'],['nif','NIF / CIF'],['dir','Dirección fiscal'],['cp','Código postal'],['mun','Municipio'],['prov','Provincia'],['rn','Representante (si firma otro)'],['rnif','NIF del representante']].forEach(function(x){h+='<div>'+pabF(x[1],inp('data-e="tit.'+x[0]+'"',t[x[0]]))+'</div>'});
  h+='</div></details>';
  h+='<div class="rh2"><h3>Recintos <span class="small">'+rs.length+(rs.length===1?' recinto':' recintos')+' · '+esc(haTxt(rs))+'</span></h3><div class="acts">'+
    '<button class="btn primary" data-act="ap-addrec">'+ic('plus')+'Añadir recintos</button><button class="btn" data-act="ap-mapa">'+ic('map')+'Mapa SIGPAC</button>'+(rs.length?'<button class="btn" data-act="ap-sig" title="Superficie, pendiente y uso del SIGPAC">Leer el SIGPAC</button>':'')+'</div></div>';
  if(!rs.length)return h+empiezaRec();
  if(nsel)h+='<div class="bulk"><b>'+nsel+(nsel===1?' seleccionado':' seleccionados')+'</b><label>Cultivo <select class="trsel sm ap-bulk" data-k="cultivo"><option value="">—</option>'+cultOpts('').replace('<option value="">Sin cultivo</option>','')+'</select></label><label>Riego <select class="trsel sm ap-bulk" data-k="riego"><option value="">—</option><option value="secano">Secano</option><option value="regadio">Regadío</option></select></label><button class="btn sm danger" data-act="ap-bulk-del">Quitar de la finca</button><button class="btn sm lk" data-act="ap-bulk-none">Quitar la selección</button></div>';
  if(isDesk()){
    h+='<div class="xs tw"><table class="rtab"><colgroup><col style="width:40px"><col style="width:200px"><col style="width:72px"><col style="width:82px"><col style="width:150px"><col><col style="width:112px"><col style="width:66px"><col style="width:44px"></colgroup><thead><tr><th class="c"><input type="checkbox" class="rsel-all" aria-label="Seleccionar todos"'+(nsel&&nsel===rs.length?' checked':'')+'></th>'+thS('ref','Recinto SIGPAC · paraje')+thS('ha','ha')+'<th>ha cultiv.</th>'+thS('cultivo','Cultivo')+'<th>Variedad</th>'+thS('riego','Riego')+'<th class="r">Pend.</th><th></th></tr></thead><tbody>'+
      rs.map(function(r){var k=esc(r.ref);return '<tr class="'+(SEL[r.ref]?'sel':'')+'"><td class="c"><input type="checkbox" class="rsel" data-ref="'+k+'"'+(SEL[r.ref]?' checked':'')+' aria-label="Seleccionar"></td><td><div class="rref"><b>'+esc(refLabel(r.ref))+'</b>'+(r.uso?'<span class="uso" title="Uso SIGPAC: '+esc(usoTxt(r.uso))+'">'+esc(r.uso)+'</span>':'')+'</div>'+inp('data-r="'+k+'" data-k="nombre"',r.nombre,'Paraje (opcional)','','cell par')+(r.hs!=null&&Math.abs(r.hs-(r.ha||0))>0.00005?'<div class="small">SIGPAC '+ha2(r.hs)+' ha</div>':'')+'</td>'+
        '<td>'+inp('data-r="'+k+'" data-k="ha" data-num="1"',r.ha,'','decimal','cell n')+'</td><td>'+inp('data-r="'+k+'" data-k="hc" data-num="1"',r.hc,ha2(r.ha||0),'decimal','cell n')+'</td>'+
        '<td>'+cultSel('data-r="'+k+'" data-k="cultivo"',r.cultivo).replace('class="trsel ap-in"','class="in ap-in cell'+(r.cultivo?'':' miss')+'"')+'</td><td>'+inp('data-r="'+k+'" data-k="variedad"',r.variedad,'','','cell')+'</td>'+
        '<td><select class="in ap-in cell" data-r="'+k+'" data-k="riego"><option value="secano">Secano</option><option value="regadio"'+(r.riego==='regadio'?' selected':'')+'>Regadío</option></select></td>'+
        '<td class="r">'+(r.pend!=null?trTxt(r.pend,1)+' %':'—')+'</td><td><button class="btn sm lk danger ib" data-act="ap-recdel" data-ref="'+k+'" aria-label="Quitar recinto" title="Quitar de la finca">'+ic('trash')+'</button></td></tr>'}).join('')+'</tbody></table></div>';
    h+='<div class="small" style="margin-top:6px">Marca varios recintos para cambiarles el cultivo o el riego de una vez. «ha cultiv.»: déjalo vacío si se cultiva todo el recinto.</div>';
  }else{
    h+='<div class="recs">'+rs.map(function(r){var k=esc(r.ref);
      return '<div class="rec card'+(SEL[r.ref]?' sel':'')+'"><div class="rh"><label class="ckr"><input type="checkbox" class="rsel" data-ref="'+k+'"'+(SEL[r.ref]?' checked':'')+'> <b>'+esc(refLabel(r.ref))+'</b></label><button class="btn sm lk danger ib" data-act="ap-recdel" data-ref="'+k+'" aria-label="Quitar recinto">'+ic('trash')+'</button></div>'+
        '<div class="small">'+esc([r.uso?'Uso '+usoTxt(r.uso):'',r.pend!=null?'pendiente '+trTxt(r.pend,1)+' %':'pendiente sin leer',r.hs!=null&&Math.abs(r.hs-(r.ha||0))>0.00005?'SIGPAC '+ha2(r.hs)+' ha':''].filter(Boolean).join(' · '))+'</div>'+
        '<div class="pa-g"><div>'+pabF('Cultivo',cultSel('data-r="'+k+'" data-k="cultivo"',r.cultivo))+'</div><div>'+pabF('Riego','<select class="trsel ap-in" data-r="'+k+'" data-k="riego"><option value="secano">Secano</option><option value="regadio"'+(r.riego==='regadio'?' selected':'')+'>Regadío</option></select>')+'</div>'+
        '<div>'+pabF('ha',inp('data-r="'+k+'" data-k="ha" data-num="1"',r.ha,'','decimal'))+'</div><div>'+pabF('ha cultivadas',inp('data-r="'+k+'" data-k="hc" data-num="1"',r.hc,ha2(r.ha||0),'decimal'))+'</div><div>'+pabF('Variedad',inp('data-r="'+k+'" data-k="variedad"',r.variedad))+'</div><div>'+pabF('Paraje',inp('data-r="'+k+'" data-k="nombre"',r.nombre))+'</div></div></div>'}).join('')+'</div>';
  }
  if((f.pab||[]).length)h+='<div class="small" style="margin-top:6px">Los recintos que añadas pasan solos a la unidad de su cultivo y riego del plan.</div>';
  h+='<div class="acts" style="margin-top:22px"><button class="btn sm danger lk" data-act="ap-fdel">'+ic('trash')+'Borrar la finca</button></div>';
  return h}
function empiezaRec(){return '<div class="opts">'+
  '<button class="opt" data-act="ap-mapa">'+ic('map')+'<b>En el mapa SIGPAC</b><span>Busca tu parcela en la foto aérea y toca los recintos.</span></button>'+
  '<button class="opt" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Del PDF de la PAC</b><span>Salen todos con su cultivo, superficie y riego.</span></button>'+
  '<button class="opt" data-act="ap-csv-pick">'+ic('xls')+'<b>De una tabla (CSV)</b><span>Con columnas de provincia, municipio, polígono, parcela y recinto.</span></button>'+
  '<button class="opt" data-act="ap-addrec">'+ic('plus')+'<b>A mano</b><span>Con el código SIGPAC de cada recinto.</span></button></div><p class="small center">También puedes arrastrar aquí el PDF de la PAC o la tabla.</p>'}
function addRecHoja(){
  sheet('<h2>Añadir recintos</h2><div class="opts two"><button class="opt" data-act="ap-mapa">'+ic('map')+'<b>En el mapa SIGPAC</b><span>Foto aérea y recintos: toca y añade.</span></button><button class="opt" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Del PDF de la PAC</b><span>Todos de una vez.</span></button><button class="opt" data-act="ap-csv-pick">'+ic('xls')+'<b>De una tabla (CSV)</b><span>Provincia, municipio, polígono, parcela y recinto.</span></button></div>'+
    '<h3>A mano</h3><label class="lbl" for="ap-ref">Código SIGPAC</label><input id="ap-ref" class="in" inputmode="numeric" placeholder="02:012:0:0:1:16:5" autocomplete="off" autofocus><div class="small">Provincia, municipio, agregado, zona, polígono, parcela y recinto (o solo provincia, municipio, polígono, parcela y recinto).</div>'+
    '<div class="pa-g"><div>'+pabF('Cultivo','<select id="ap-rc" class="trsel">'+cultOpts('')+'</select>')+'</div><div>'+pabF('Riego','<select id="ap-rr" class="trsel"><option value="secano">Secano</option><option value="regadio">Regadío</option></select>')+'</div><div>'+pabF('ha (vacío: las del SIGPAC)','<input id="ap-rh" class="in" inputmode="decimal" autocomplete="off">')+'</div></div>'+
    '<div class="acts"><button class="btn primary" data-act="ap-ref-add">Añadir este recinto</button><button class="btn" data-act="ap-close">Cerrar</button></div>','wide')}
function refAdd(){var f=finca(UI.fid), ref=refParse(($('#ap-ref')||{}).value);if(!ref){toast('Ese código SIGPAC no está completo.');return}
  var hv=($('#ap-rh')||{}).value, ha=hv?numEs(hv):0; if(hv&&(ha===null||isNaN(ha))){toast('Pon las ha con coma decimal, por ejemplo 2,35.');return}
  undoPush(); var R=addRecs(f,[{ref:ref,ha:ha||0,cultivo:($('#ap-rc')||{}).value||'',riego:($('#ap-rr')||{}).value||'secano'}]);
  if(!R.added){toast('Ese recinto ya está'+(S.recintos[ref]&&S.recintos[ref].fincaId!==f.id?' en otra finca':'')+'.');return}
  save(); $('#ap-ref').value=''; $('#ap-rh').value=''; $('#ap-ref').focus(); toast('Recinto '+refLabel(ref)+' añadido. Leo el SIGPAC…'); render();
  sigLeer(f,[ref],function(ok){if(!ok)toast('No he podido leer ese recinto en el SIGPAC: revisa el código o pon las ha a mano.')});
}

/* ───── unidades: los apartados de cada unidad se reparten entre los pasos ───── */
var SEG={ud:[/^(Recintos)/],su:[/^Suelo del recinto/,/^Agua de riego/],ne:[/^Necesidades del cultivo/,/^Aportes que ya hay/,/^Datos de inicio de campa/,/^Materia org/],pg:[/^Medidas para reducir/,/^Restricciones y condiciones/],rv:[/^Balance de nutrientes/,/^Comprobaci/]};
function udBox(f,p,u,C,i,step){
  var st=pabStTxt(f,u,C), firstNx=udsOrden(p,CTX?CTX.Cs:p.ud.map(function(x){return pabCalc(f,x,p)}))[0];
  if(PAB.open[u.id]===undefined)PAB.open[u.id]=!C.ex&&i===firstNx;
  var html=pabUdHTML(f,p,u,C,i), e=html.indexOf('</summary>')+10, body=html.slice(e,html.lastIndexOf('</details>'));
  body=body.replace(/<div class="acts" style="margin:10px 0 0"><button class="btn sm lk danger" data-act="pa-ud-del"[^]*?<\/div>$/,'');
  var parts=body.split(/(?=<h4 )/), first=parts[0], segs={};
  parts.slice(1).forEach(function(s){var t=s.replace(/^<h4[^>]*>/,'').replace(/<[^>]+>/g,'');for(var k in SEG)SEG[k].forEach(function(re,ix){if(re.test(t))(segs[k]||(segs[k]=[]))[ix]=s})});
  var h='<details class="more pa-box ud'+(C.ex?' exu':'')+'" id="u'+i+'" data-pa-o="'+esc(u.id)+'"'+(PAB.open[u.id]?' open':'')+'><summary><b>'+esc(u.n)+'</b><span class="small">'+ha2(C.ha)+' ha · '+C.rs.length+(C.rs.length===1?' recinto':' recintos')+'</span>'+stChip(st)+(C.ex?hlp('ex'):'')+'</summary>';
  if(C.ex&&step!=='ud')h+='<div class="note">Unidad de secano de '+ha2(C.ha)+' ha: no necesita plan de abonado (art. 4.2 b del RD 1051/2022). Puedes rellenarla igualmente.</div>';
  if(step==='ud')h+=first+(segs.ud||[]).join('')+'<div class="acts"><button class="btn sm danger lk" data-act="pa-ud-del" data-u="'+i+'">'+ic('trash')+'Borrar unidad</button></div>';
  else if(step==='pg')h+=progHTML(f,p,u,C,i)+(u.d0===1?'':calHTML(u,C,i)+mitResHTML(f,p,u,C,i)+pabOrgHTML(u,C));
  else if(step==='su')h+=((segs.su||[]).filter(Boolean).join('')||'')+sueloHTML(u,C,i);
  else h+=(segs[step]||[]).filter(Boolean).join('')||'<div class="small">Nada que rellenar aquí para esta unidad.</div>';
  return h+'</details>';
}
function udStepHTML(f,p,Cs,ord){
  var h=climaHTML(f,p);
  ord.forEach(function(i){h+=udBox(f,p,p.ud[i],Cs[i],i,'ud')});
  h+='<div class="acts"><button class="btn" data-act="pa-ud-new">'+ic('plus')+'Unidad</button></div>';
  return h}
function climaHTML(f,p){var cl=p.cl||PAB_CL, sum=(+cl.o||0)+(+cl.i||0)+(+cl.p||0)+(+cl.v||0), ck=climaKey(f);
  var h='<details class="more pa-box"'+(PAB.open.cl?' open':'')+' data-pa-o="cl"><summary><b>Lluvia de la zona</b>'+hlp('cl')+'<span class="small">'+nf(+cl.a||0,0)+' mm al año · '+esc(String(cl.src||'').replace(/^AEMET, /,'').replace(/ \(\w+\), valores.*/,''))+'</span></summary>';
  h+='<div class="pa-g">'+[['a','Media anual (mm)'],['o','Otoño, sep-nov (mm)'],['i','Invierno, dic-feb (mm)'],['p','Primavera, mar-may (mm)'],['v','Verano, jun-ago (mm)']].map(function(x){return '<div>'+pabF(x[1],pabIn('cl.'+x[0],cl[x[0]]))+'</div>'}).join('')+'</div>'+pabF('Fuente',pabTx('cl.src',cl.src));
  h+='<div class="small" style="margin-top:6px">Suma de las estaciones: '+nf(sum,0)+' mm'+(abHas(cl.a)&&Math.abs(sum-cl.a)>2?' · <b>no cuadra con la media anual</b>':'')+'.'+(ck.km!=null?' La estación de AEMET con valores normales más cercana a los recintos es '+esc(CLIMA[ck.k].n)+' (a unos '+ck.km+' km).':'')+'</div>';
  h+=pabF('Poner los valores de AEMET de','<select class="trsel ap-clprov"><option value="">Elegir estación…</option>'+Object.keys(CLIMA).map(function(k){return '<option value="'+k+'">'+CLIMA[k].n+(+k===ck.k?' (la más cercana)':'')+'</option>'}).join('')+'</select>');
  return h+'</details>'}

/* ───── paso 3: semáforo del suelo ───── */
function sueloHTML(u,C,i){var su=u.su||{}, any=SUELO.some(function(s){return abHas(su[s[0]])});
  var h='<h4 style="margin:14px 0 4px">Cómo está el suelo <span class="small">(niveles orientativos, criterio técnico)</span></h4>';
  if(!any)return h+'<div class="small">Cuando pongas el análisis verás aquí si cada valor es bajo, medio o alto.</div>';
  h+='<div class="sem">'+SUELO.map(function(s){var v=su[s[0]];if(!abHas(v))return '<div class="sm0"><span>'+s[1]+'</span><b>—</b></div>';var l=nivel(+v,s[3]),ph=s[0]==='ph';
    return '<div class="lv'+(ph?(l===2?' l2':l===0||l===4?' l4':' l3'):' l'+l)+'"><span>'+s[1]+'</span><b>'+trTxt(+v,s[4])+(s[2]?' '+s[2]:'')+'</b><i>'+(ph?NIVPH:NIV)[l]+'</i><div class="lvb">'+[0,1,2,3,4].map(function(k){return '<em'+(k===l?' class="on"':'')+'></em>'}).join('')+'</div></div>'}).join('')+'</div>';
  var cr=creditoSuelo(u,C), ap=(u.ap||{}).s||{};
  if(cr.p||cr.k){h+='<div class="note">Con fósforo o potasio altos se puede descontar parte de la necesidad (criterio técnico: nivel alto, el 25 %; muy alto, el 50 %). Propuesta: '+[cr.p?'<b>'+cr.p+' U.F. de P₂O₅</b>':'',cr.k?'<b>'+cr.k+' U.F. de K₂O</b>':''].filter(Boolean).join(' y ')+' por ha.'+
    '<div class="acts"><button class="btn sm" data-act="ap-cred" data-u="'+i+'">Ponerlo como aporte del suelo</button>'+(abHas(ap.p)||abHas(ap.k)?'<span class="small">Ahora: P₂O₅ '+trTxt(+ap.p||0,0)+' · K₂O '+trTxt(+ap.k||0,0)+'</span>':'')+'</div></div>'}
  else h+='<div class="small" style="margin-top:6px">El suelo no descuenta nada del balance mientras no lo pongas en «Cosecha › Aportes que ya hay» (lo prudente con niveles medios o bajos).</div>';
  return h}

/* ───── paso 5: programa como una hoja de cálculo ───── */
var FO=['Suelo','Localizado','Fertirrigación','Foliar','Enterrado'];
function ufTxt(o){return o?trTxt(o.uf.n,1,true)+' · '+trTxt(o.uf.p,1,true)+' · '+trTxt(o.uf.k,1,true):''}
function progHTML(f,p,u,C,i){
  var P='ud.'+i+'.', pg=u.pg||{}, rows=pg.rows||[], co=!!UI.cost, desk=isDesk();
  var h='<h4 style="margin:14px 0 4px">Programa previsto <span class="small">(art. 6.4)</span>'+hlp('pg')+'</h4>'+pabF('¿Sin abonado esta campaña?',pabChips(P+'d0',u.d0===1?1:0,[[1,'Sí, dosis 0'],[0,'No, hay programa']]));
  if(u.d0===1)return h+'<div class="note">No se prevé abonar esta campaña (programa a dosis 0). No hace falta programa, análisis de suelo, medidas de amoniaco ni restricciones.</div>';
  var fo=function(q,R){var L=FO.slice();if(q.fo&&L.indexOf(q.fo)<0)L.unshift(q.fo);return '<select class="in pa-in cell" data-p="'+esc(R+'fo')+'" aria-label="Forma de aplicación">'+L.map(function(x){return '<option'+(x===(q.fo||'')?' selected':'')+'>'+esc(x)+'</option>'}).join('')+'</select>'};
  var cin=function(R,v,ph,cls,lbl){return '<input class="in pa-in cell'+(cls?' '+cls:'')+'" data-p="'+esc(R)+'"'+(cls&&/n/.test(cls)?' inputmode="decimal"':'')+' value="'+esc(v===undefined||v===null?'':(typeof v==='number'?abEd(v):v))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off" aria-label="'+esc(lbl||'')+'">'};
  var sg=function(R,v,kind,ph,lbl){return '<div class="sgw">'+cin(R,v,ph,'sgi',lbl).replace('class="in pa-in cell sgi"','class="in pa-in cell sgi" data-sg="'+kind+'"')+'</div>'};
  var Cst=costeUd(u,C);
  if(desk){
    h+='<div class="tw xs"><table class="sheet pgt'+(co?' co':'')+'"><colgroup><col class="c-mes"><col class="c-ab"><col class="c-d"><col class="c-r"><col class="c-r"><col class="c-r"><col class="c-uf"><col class="c-fo"><col class="c-org">'+(co?'<col class="c-pz"><col class="c-eh">':'')+'<col class="c-ac"></colgroup><thead><tr><th>Mes</th><th>Abono</th><th class="r">kg/ha</th><th class="r">N %</th><th class="r">P₂O₅ %</th><th class="r">K₂O %</th><th class="r">U.F./ha '+hlp('uf')+'</th><th>Forma</th><th class="c" title="Abono orgánico">Org.</th>'+(co?'<th class="r">€/kg</th><th class="r">€/ha</th>':'')+'<th></th></tr></thead><tbody>';
    rows.forEach(function(q,j){var R=P+'pg.rows.'+j+'.', o=C.rows[j];
      h+='<tr data-j="'+j+'"><td>'+sg(R+'mo',q.mo,'mes','Marzo','Mes')+'</td><td title="'+esc(q.n||'')+'">'+sg(R+'n',q.n,'ab','Buscar abono…','Abono')+'</td><td>'+cin(R+'d',q.d,'','n','Dosis kg/ha')+'</td><td>'+cin(R+'rn',q.rn,'0','n s','N %')+'</td><td>'+cin(R+'rp',q.rp,'0','n s','P₂O₅ %')+'</td><td>'+cin(R+'rk',q.rk,'0','n s','K₂O %')+'</td>'+
        '<td class="r uf">'+ufTxt(o)+'</td><td>'+fo(q,R)+'</td><td class="c"><input type="checkbox" class="pa-chk" data-p="'+esc(R+'org')+'"'+(q.org===1?' checked':'')+' aria-label="Abono orgánico"></td>'+
        (co?'<td>'+cin(R+'pz',q.pz,'0,00','n s','Precio €/kg')+'</td><td class="r">'+(abHas(q.pz)&&q.d?nf(q.d*q.pz,2):'')+'</td>':'')+
        '<td class="c nw"><button type="button" class="btn sm lk ib" data-act="ap-rowmore" data-id="'+esc(q.id)+'" aria-label="Equipo y medida de mitigación" title="Equipo y medida de mitigación de este aporte">⋯</button><button type="button" class="btn sm lk danger ib" data-act="pa-row-del" data-u="'+i+'" data-j="'+j+'" aria-label="Quitar aporte">'+ic('close')+'</button></td></tr>';
      if(RM[q.id])h+='<tr class="sub"><td></td><td colspan="'+(co?10:8)+'"><div class="subg"><label class="inl">Equipo '+cin(R+'eq',q.eq,'Abonadora','','Equipo')+'</label><label class="inl">Medida de mitigación '+cin(R+'mi',q.mi,'Ej. Riego justo después de abonar','','Medida de mitigación')+'</label></div></td><td></td></tr>'});
    if(!rows.length)h+='<tr><td colspan="'+(co?12:10)+'" class="small empty">Aún no hay aportes. Pulsa «Proponer programa» o «+ Añadir aporte».</td></tr>';
    h+='</tbody>'+totHTML(u,C,i,co)+'</table></div>';
  }else{
    rows.forEach(function(q,j){var R=P+'pg.rows.'+j+'.', o=C.rows[j];
      h+='<div class="pa-card"><div class="hd"><span>Aporte '+(j+1)+(o&&o.d?' · <span class="small">U.F. '+ufTxt(o)+'</span>':'')+'</span><button type="button" class="btn sm lk danger ib" data-act="pa-row-del" data-u="'+i+'" data-j="'+j+'" aria-label="Quitar aporte">'+ic('close')+'</button></div>'+
        pabF('Abono',sg(R+'n',q.n,'ab','Buscar abono…','Abono'))+'<div class="pa-g"><div>'+pabF('Mes',sg(R+'mo',q.mo,'mes','Marzo','Mes'))+'</div><div>'+pabF('Dosis (kg/ha)',cin(R+'d',q.d,'','n'))+'</div>'+
        '<div>'+pabF('N (%)',cin(R+'rn',q.rn,'0','n'))+'</div><div>'+pabF('P₂O₅ (%)',cin(R+'rp',q.rp,'0','n'))+'</div><div>'+pabF('K₂O (%)',cin(R+'rk',q.rk,'0','n'))+'</div>'+
        '<div>'+pabF('Forma',fo(q,R))+'</div><div>'+pabF('Equipo',cin(R+'eq',q.eq,'Abonadora'))+'</div><div>'+pabF('Mitigación',cin(R+'mi',q.mi,''))+'</div>'+(co?'<div>'+pabF('Precio (€/kg)',cin(R+'pz',q.pz,'0,00','n'))+'</div>':'')+'</div>'+
        '<label class="ckr" style="margin-top:8px"><input type="checkbox" class="pa-chk" data-p="'+esc(R+'org')+'"'+(q.org===1?' checked':'')+'> Abono orgánico</label></div>'});
    if(!rows.length)h+='<div class="small">Aún no hay aportes.</div>';
    h+='<table class="sheet mini">'+totHTML(u,C,i,co)+'</table>';
  }
  var otras=CTX&&CTX.p?CTX.p.ud.map(function(x,k){return {x:x,k:k}}).filter(function(o){return o.k!==i&&((o.x.pg||{}).rows||[]).length}):[];
  h+='<div class="acts pgacts"><button class="btn sm" data-act="pa-row-add" data-u="'+i+'">'+ic('plus')+'Añadir aporte</button><button class="btn sm primary" data-act="ap-prop" data-u="'+i+'" title="Abonos del catálogo que cubren la necesidad, repartidos por meses">'+ic('spark')+'Proponer programa</button>'+
    (otras.length?'<select class="trsel sm ap-copyud" data-u="'+i+'" aria-label="Copiar de otra unidad"><option value="">Copiar de otra unidad…</option>'+otras.map(function(o){return '<option value="'+o.k+'">'+esc(o.x.n)+'</option>'}).join('')+'</select>':'')+
    '<label class="ckr sm"><input type="checkbox" class="ap-cost"'+(co?' checked':'')+'> Ver costes</label><button class="btn sm lk" data-act="ap-mis">Mis abonos</button></div>';
  h+='<div class="small" style="margin-top:6px">La riqueza del catálogo es la típica de ese tipo de abono: pon la de la etiqueta si es otra.'+(desk?' Tabulador e Intro para moverse, como en Excel.':'')+(co?' Los precios solo se ven aquí: no salen en el PDF ni en el Excel.':'')+'</div>';
  if(co&&rows.length)h+='<div class="small">Coste: <b>'+nf(Cst.ha,2)+' €/ha</b> · <b>'+nf(Cst.tot,0)+' €</b> en las '+ha2(C.ha)+' ha'+(Cst.sin?' (hay abonos sin precio)':'')+'.</div>';
  return h}
function totHTML(u,C,i,co){var B=C.B,cell=function(k){var b=B[k],c=stCls(b.st);return '<span class="tn t-'+c+'" title="'+esc(b.st)+'">'+trTxt(b.pl,0)+'</span>'};
  var cols=co?12:10;
  return '<tfoot><tr class="tot" id="u'+i+'-tot"><td colspan="6"><b>Total planificado</b> <span class="small">U.F./ha</span></td><td class="r"><b>'+cell('n')+'-'+cell('p')+'-'+cell('k')+'</b></td><td colspan="'+(cols-7)+'" class="small">'+
    ['n','p','k'].map(function(k){var b=B[k];return b.st==='Ajustado'?'':(b.dv<0?'faltan '+trTxt(-b.dv,0):'sobran '+trTxt(b.dv,0))+' de '+daLbl(PAB_NU,k)}).filter(Boolean).join(' · ')+(u.zv&&C.L.ok?(C.lok?'':' · <b class="t-bad">pasa del límite de N ('+nf(C.L.lim,0)+')</b>'):'')+'</td></tr>'+
    '<tr class="need"><td colspan="6">Necesidad neta '+hlp('net')+'</td><td class="r">'+trTxt(B.n.net,0)+'-'+trTxt(B.p.net,0)+'-'+trTxt(B.k.net,0)+'</td><td colspan="'+(cols-7)+'"></td></tr></tfoot>'}

/* medidas del anexo V y restricciones, como lista para marcar */
function mitResHTML(f,p,u,C,i){var P='ud.'+i+'.', mit=u.mit||{}, res=u.res||{};
  var org=C.rows.some(function(o){return o.org}), urea=C.rows.some(function(o){return /urea/i.test(o.n)}), hasA=PAB_MIT.some(function(x){return x[1]==='A'&&(mit[x[0]]||{}).a===1}), hasB=PAB_MIT.some(function(x){return x[1]==='B'&&(mit[x[0]]||{}).a===1});
  var row=function(path,x,on,dpath,dv,ph){return '<div class="ckrow'+(on?' on':'')+'"><label class="ckr"><input type="checkbox" class="pa-chk" data-p="'+esc(path)+'"'+(on?' checked':'')+'> <span>'+esc(x)+'</span></label>'+(on?'<input class="in pa-in sm" data-p="'+esc(dpath)+'" value="'+esc(dv||'')+'" placeholder="'+esc(ph)+'" autocomplete="off">':'')+'</div>'};
  var h='<div class="chkl" id="u'+i+'-mit"><div class="chh"><h4>Medidas para reducir amoniaco y gases <span class="small">(art. 6.5 y anexo V)</span>'+hlp('mit')+'</h4><button class="btn sm" data-act="ap-mitprop" data-u="'+i+'">Marcar las que encajan con el programa</button></div>';
  var gr=function(g){return '<div class="k gk">'+PAB_MITG[g]+'</div>'+PAB_MIT.filter(function(x){return x[1]===g}).map(function(x){var m=mit[x[0]]||{};return row(P+'mit.'+x[0]+'.a',x[2],m.a===1,P+'mit.'+x[0]+'.d',m.d===undefined?x[3]:m.d,'Cómo se aplica')}).join('')};
  h+=gr('B')+gr('O')+((org||hasA)?gr('A'):'<details class="more"><summary>Parte A: estiércoles, purines y materiales orgánicos</summary>'+gr('A')+'</details>');
  if(urea&&!hasB)h+='<div class="note">El programa lleva urea: hace falta al menos una medida de la parte B (art. 12.3).</div>';
  if(org&&!hasA)h+='<div class="small" style="margin-top:4px">Si aplicas estiércoles, purines u otros materiales orgánicos sin compostar, hace falta al menos una medida de la parte A (art. 10.4).</div>';
  h+='</div><div class="chkl" id="u'+i+'-res"><div class="chh"><h4>Restricciones y condiciones de aplicación <span class="small">(art. 4.4 y Decreto 80/2025)</span>'+hlp('res')+'</h4><button class="btn sm" data-act="ap-resall" data-u="'+i+'">Marcar todas</button></div>';
  if(u.resProp)h+='<div class="note">Vienen marcadas como propuesta: comprueba que se cumplen en tu finca y desmarca la que no.</div>';
  h+=PAB_RES.map(function(x){var m=res[x[0]]||{};return row(P+'res.'+x[0]+'.a',x[1],m.a===1,P+'res.'+x[0]+'.d',m.d===undefined?x[2]:m.d,'Observación')}).join('');
  var pmax=C.rs.reduce(function(a,r){var x=((u.rc||{})[r.ref]||{}).pe;return abHas(x)?Math.max(a,+x):a},-1);
  h+='<div class="small" style="margin-top:6px">Pendiente máxima de los recintos: <b>'+(pmax<0?'sin dato':trTxt(pmax,1)+' %')+'</b>'+hlp('pend')+(pmax>15?' · <b>más del 15 %: prohibido aplicar estiércoles y abonos N y P en zona vulnerable</b>':(pmax>2&&u.zv?' · más del 2 %: tomar precauciones (curvas de nivel, franjas sin abonar…)':''))+'.</div>';
  var pt=pabPerTxt(u.ck); if(pt)h+='<div class="small">Sin abono nitrogenado en '+pt+' (anexo II)'+hlp('per')+(u.rg==='regadio'&&u.loc!==0?': no se aplica con riego localizado':'')+'.</div>';
  var ph=pabPerHits(f,u,p,C); if(ph.length)h+='<div class="note bad">Aportes de N en meses prohibidos: '+esc(ph.join(' · '))+'</div>';
  return h+'</div>'}

/* ───── calendario de abonado: U.F./ha por mes, con los meses sin N sombreados ───── */
var CAMP=[11,12,1,2,3,4,5,6,7,8,9,10];
function calHTML(u,C,i){var K=calMeses(u,C), M=K.M, pm=perMeses(u), mx=0;CAMP.forEach(function(m){mx=Math.max(mx,M[m].n+M[m].p+M[m].k)});
  var W=680,H=230,l=44,r=10,t=34,b=34,pw=W-l-r,ph=H-t-b,bw=pw/12;
  var step=mx<=20?5:mx<=50?10:mx<=100?25:mx<=200?50:100, top=Math.max(step,Math.ceil(mx/step)*step), Y=function(v){return t+ph-v/top*ph};
  var s='<svg class="cal" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Calendario de abonado: unidades fertilizantes por hectárea en cada mes"><defs><pattern id="hz'+i+'" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="6" class="hzl"/></pattern></defs>';
  CAMP.forEach(function(m,k){if(pm.indexOf(m)>=0)s+='<rect x="'+(l+k*bw)+'" y="'+t+'" width="'+bw+'" height="'+ph+'" fill="url(#hz'+i+')" class="hz"/>'});
  for(var g=0;g<=top;g+=step){var y=Y(g);s+='<line x1="'+l+'" x2="'+(W-r)+'" y1="'+y+'" y2="'+y+'" class="'+(g?'gr':'ax')+'"/><text x="'+(l-6)+'" y="'+(y+4)+'" text-anchor="end" class="tk">'+g+'</text>'}
  s+='<text x="'+l+'" y="14" class="ttl">U.F./ha por mes</text>';
  var lg=[['n','N'],['p','P₂O₅'],['k','K₂O']], lx=W-r-190; lg.forEach(function(x,j){s+='<rect x="'+(lx+j*64)+'" y="6" width="10" height="10" rx="2" class="s-'+x[0]+'"/><text x="'+(lx+j*64+14)+'" y="15" class="lg">'+x[1]+'</text>'});
  CAMP.forEach(function(m,k){var x=l+k*bw+bw*0.2,w=bw*0.6,y0=t+ph,o=M[m],tot=o.n+o.p+o.k;
    ['n','p','k'].forEach(function(kk){var v=o[kk];if(v<=0.05)return;var hh=v/top*ph;var y=y0-hh;var last=['n','p','k'].filter(function(z){return o[z]>0.05}).pop()===kk;
      s+='<path class="s-'+kk+' bar" d="'+(last?'M'+x+','+(y0)+'V'+(y+Math.min(4,hh))+'Q'+x+','+y+' '+(x+Math.min(4,w/2))+','+y+'H'+(x+w-Math.min(4,w/2))+'Q'+(x+w)+','+y+' '+(x+w)+','+(y+Math.min(4,hh))+'V'+y0+'Z':'M'+x+','+y0+'V'+y+'H'+(x+w)+'V'+y0+'Z')+'"/>';
      y0=y-2});
    s+='<rect x="'+(l+k*bw)+'" y="'+t+'" width="'+bw+'" height="'+ph+'" class="hit" data-tip="'+esc(MESES[m-1]+(tot>0.05?': N '+trTxt(o.n,1)+' · P₂O₅ '+trTxt(o.p,1)+' · K₂O '+trTxt(o.k,1)+' U.F./ha'+(o.L.length?' ('+o.L.join(', ')+')':''):': sin abono')+(pm.indexOf(m)>=0?' · mes sin N (anexo II)':''))+'"/>';
    s+='<text x="'+(l+k*bw+bw/2)+'" y="'+(H-b+16)+'" text-anchor="middle" class="tk">'+MESES[m-1].slice(0,3)+'</text>'});
  s+='</svg>';
  var nota=[pm.length?'Sombreado: meses sin abono nitrogenado (anexo II).':(PAB_PER[u.ck]&&u.rg==='regadio'?'Con riego localizado no hay meses prohibidos.':''),K.sin?K.sin+(K.sin===1?' aporte sin mes no sale':' aportes sin mes no salen')+' en el calendario.':''].filter(Boolean).join(' ');
  return '<div class="calw"><h4 style="margin:14px 0 4px">Calendario de abonado</h4>'+s+(nota?'<div class="small">'+nota+'</div>':'')+'</div>'}

/* ───── paso 6: revisar, lista de la compra, firmas y PDF ───── */
function rvStepHTML(f,p,Cs,ord){
  var full=pabHTML(f), a=full.indexOf('<h3 style="margin-top:12px">Resumen</h3>'), b=full.indexOf('<details class="more pa-box"',a), y1=+p.y+1, ex=(f.pab||[]).some(function(x){return +x.y===y1});
  var h='<div class="acts rvacts"><button class="btn primary xl" data-act="ap-pdf">'+ic('pdf')+'PDF del plan</button><button class="btn xl" data-act="ap-xlsx">'+ic('xls')+'Excel</button><button class="btn xl" data-act="ap-sigcamp">'+(ex?'Ir al plan de '+pabCamp(y1):'Hacer el plan de '+pabCamp(y1))+'</button></div>';
  var rvl='<div class="rvl">';
  if(a>=0&&b>a)rvl+=full.slice(a,b).replace('<h3 style="margin-top:12px">Resumen</h3>','<h3>Resumen</h3>');
  rvl+=compraHTML(f,p,Cs);
  ord.forEach(function(i){rvl+=udBox(f,p,p.ud[i],Cs[i],i,'rv')});
  if(!pabIso(p.fe)||ctxFalta('[data-p="AS.n"]'))PAB.open.fi=true;
  rvl+=pabPlanMore(f,p);
  rvl+='<div class="acts" style="margin-top:22px"><button class="btn sm danger lk" data-act="pa-del">'+ic('trash')+'Borrar el plan de esta campaña</button></div></div>';
  if(isWide()||isDesk())h+='<div class="rvg">'+rvl+'<div class="prev"><div class="prevh"><b>Vista previa del PDF</b><button class="btn sm" data-act="ap-prevupd">Actualizar</button></div><div class="prevb" id="prevb"><div class="small">Preparando la vista previa…</div></div></div></div>';
  else h+=rvl;
  return h}
function ctxFalta(sel){return CTX&&CTX.F.some(function(x){return x.sel===sel})}
function compraHTML(f,p,Cs){var L=compra(f,p,Cs);if(!L.length)return '';var co=!!UI.cost, tot=L.reduce(function(a,x){return a+x.eur},0);
  return '<h3>Lista de la compra <span class="small">(toda la finca, campaña '+esc(pabCamp(p.y))+')</span></h3><div class="xs"><table class="patb shop"><thead><tr><th>Abono</th><th class="r">Total</th><th class="r">Equivale a</th>'+(co?'<th class="r">Coste</th>':'')+'<th>Unidades</th></tr></thead><tbody>'+
    L.map(function(x){var kg=x.kg;return '<tr><td><b>'+esc(x.n)+'</b></td><td class="r">'+(kg>=1000?nf(kg/1000,2)+' t':nf(kg,0)+' kg')+'</td><td class="r small">'+(x.dens?'≈ '+nf(kg/x.dens,0)+' L':'≈ '+nf(Math.ceil(kg/25),0)+' sacos de 25 kg')+'</td>'+(co?'<td class="r">'+(x.eur?nf(x.eur,0)+' €':'—')+'</td>':'')+'<td class="small">'+esc(Object.keys(x.uds).join(', '))+'</td></tr>'}).join('')+
    '</tbody>'+(co&&tot?'<tfoot><tr><td colspan="3"><b>Total</b></td><td class="r"><b>'+nf(tot,0)+' €</b></td><td></td></tr></tfoot>':'')+'</table></div><div class="acts"><button class="btn sm" data-act="ap-compra-copy">'+ic('copy')+'Copiar la lista</button></div><div class="small">Solo en pantalla: no sale en el PDF ni en el Excel.</div>'}
/* vista previa del PDF dentro de la página (ordenador) */
var PREV={url:null,mod:-1,t:null,busy:false,id:null};
function prevAuto(){var el=$('#prevb');if(!el)return;var f=finca(UI.fid),p=f&&pabCur(f);if(!p)return;
  if(PREV.url&&PREV.mod===S.mod&&PREV.id===p.id){prevShow();return}
  clearTimeout(PREV.t);PREV.t=setTimeout(prevMake,PREV.url?900:150)}
function prevShow(){var el=$('#prevb');if(!el||!PREV.url)return;if(el.dataset.u===PREV.url)return;el.dataset.u=PREV.url;el.innerHTML='<iframe title="Vista previa del PDF del plan" src="'+PREV.url+'#toolbar=0&navpanes=0&view=FitH"></iframe>'}
function prevMake(){var f=finca(UI.fid),p=f&&pabCur(f);if(!p||PREV.busy)return;PREV.busy=true;var mod=S.mod;
  ensurePdf().then(function(){var b=buildPabPdf(f,p);if(PREV.url)URL.revokeObjectURL(PREV.url);PREV.url=URL.createObjectURL(b);PREV.mod=mod;PREV.id=p.id;PREV.busy=false;prevShow()},function(){PREV.busy=false;var el=$('#prevb');if(el)el.innerHTML='<div class="small">Para la vista previa hace falta abrir la herramienta una vez con internet.</div>'})
  .catch(function(){PREV.busy=false})}

/* ───── panel de balance (derecha en el ordenador; barra plegable en pantallas pequeñas) ───── */
function balHTML(f,ctx){var p=ctx.p;
  if(!p)return '<div class="bp nop"><div class="bh"><span class="k">Balance (U.F./ha)</span></div><p class="small">Cuando crees el plan verás aquí, al momento, si los abonos cubren lo que necesita cada unidad.</p></div>';
  var uds=p.ud||[]; if(!uds.length)return '<div class="bp"><p class="small">El plan no tiene unidades.</p></div>';
  var bi=-1;uds.forEach(function(u,i){if(u.id===UI.bu)bi=i});if(bi<0)bi=udsOrden(p,ctx.Cs)[0];
  var u=uds[bi], C=ctx.Cs[bi], h='<div class="bp'+(UI.balOpen?' open':'')+'">';
  var mini=['n','p','k'].map(function(k){var b=C.B[k],c=stCls(b.st);return '<span class="t-'+c+'">'+daLbl(PAB_NU,k)+' '+(c==='ok'?'✓':c==='warn'?'▼'+trTxt(-b.dv,0):'▲'+trTxt(b.dv,0))+'</span>'}).join(' · ');
  h+='<button class="bcol" data-act="ap-baltog" aria-expanded="'+(!!UI.balOpen)+'"><b>Balance</b> <span class="small">'+esc(u.n)+'</span> <span class="bm">'+(C.ex?'Exenta':u.d0===1?'Sin abonado':mini)+'</span><i>'+(UI.balOpen?'▾':'▴')+'</i></button>';
  h+='<div class="bbody"><div class="bh"><span class="k">Balance (U.F./ha)</span>'+hlp('net')+'</div>';
  if(uds.length>1)h+='<div class="btabs">'+udsOrden(p,ctx.Cs).map(function(i){var st=pabStTxt(f,uds[i],ctx.Cs[i]);return '<button class="bt'+(i===bi?' on':'')+'" data-act="ap-bu" data-u="'+i+'"><i class="dot d-'+stCls(st)+'"></i>'+esc(uds[i].n)+'</button>'}).join('')+'</div>';
  h+='<div class="bun"><b>'+esc(u.n)+'</b><span class="small">'+ha2(C.ha)+' ha</span></div>';
  if(C.ex)h+='<div class="note">Secano de '+ha2(C.ha)+' ha: exenta de plan (art. 4.2 b). El balance es orientativo.</div>';
  if(u.d0===1)h+='<div class="note">Sin abonado esta campaña (dosis 0).</div>';
  else{
    ['n','p','k'].forEach(function(k){var b=C.B[k], c=stCls(b.st), lim=k==='n'&&u.zv&&C.L.ok?C.L.lim:0, mx=Math.max(b.net,b.pl,lim,1)*1.12, w=function(v){return Math.min(100,v/mx*100).toFixed(1)+'%'};
      var fx=arreglo(f,u,C,k);
      h+='<div class="bn"><div class="bnh"><b>'+daLbl(PAB_NU,k)+'</b><span>'+trTxt(b.pl,0)+' de '+trTxt(b.net,0)+'</span></div>'+
        '<div class="bar2" title="Planificado '+trTxt(b.pl,1)+' · necesidad neta '+trTxt(b.net,1)+'"><span class="fl t-'+c+'" style="width:'+w(b.pl)+'"></span><span class="tg" style="left:'+w(b.net)+'"></span>'+(lim?'<span class="lm" style="left:'+w(lim)+'" title="Límite de N: '+nf(lim,0)+'"></span>':'')+'</div>'+
        '<div class="bst t-'+c+'">'+(c==='ok'?'✓ Ajustado':c==='warn'?'▼ Faltan '+trTxt(-b.dv,0):'▲ Sobran '+trTxt(b.dv,0))+'</div>'+
        (fx?'<div class="fix">'+(fx.t==='add'?'Para cubrirlo: '+esc(fx.txt):'Para ajustarlo: '+esc(fx.txt))+' <button class="btn sm lk" data-act="ap-fix" data-u="'+bi+'" data-k="'+k+'">'+(fx.t==='add'?'Añadir':'Aplicar')+'</button></div>':'')+'</div>'});
    if(u.zv&&C.L.ok)h+='<div class="blim'+(C.lok?'':' t-bad')+'">Límite de N (zona vulnerable)'+hlp('lim')+': <b>'+nf(C.L.lim,0)+' kg/ha</b> · planificado '+trTxt(C.pl.n,0)+(C.lok?' ✓':' ▲')+'</div>';
    var B=C.B; if(!(B.n.s||B.p.s||B.k.s))h+='<div class="small bsu">El suelo cuenta 0 en el balance'+hlp('suelo')+'. <button class="lnk" data-act="ap-goto" data-s="su" data-u="'+bi+'">Ver el semáforo del suelo</button></div>';
    if(UI.cost){var Cst=costeUd(u,C);if(Cst.ha)h+='<div class="bcost">'+ic('euro')+'<b>'+nf(Cst.ha,2)+' €/ha</b> · '+nf(Cst.tot,0)+' € en la unidad'+(Cst.sin?' <span class="small">(faltan precios)</span>':'')+'</div>'}
    if(UI.step!=='pg')h+='<button class="btn sm wide" data-act="ap-goto" data-s="pg" data-u="'+bi+'">Ir al programa de «'+esc(u.n)+'»</button>';
  }
  return h+'</div></div>'}

/* ───── pantalla de inicio: tabla de fincas ───── */
function inicioHTML(){
  var h='', sinCopia=S.fincas.length&&S.mod>(S.copia||0)+1000;
  if(!S.fincas.length){
    h+='<div class="hero"><div class="ht"><h1>Plan de abonado</h1><p class="lead">Del PDF de la PAC (o del mapa SIGPAC) al plan de abonado en PDF y Excel, paso a paso. Real Decreto 1051/2022 y Decreto 80/2025 de Castilla-La Mancha.</p><p class="small">Todo se queda en este ordenador: no hay cuentas ni servidor. Para no perder nada, guarda una copia o activa el guardado automático en un archivo.</p></div>'+
      '<div class="k" style="margin-top:18px">Empieza aquí</div><div class="opts">'+
      '<button class="opt big" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Leer el PDF de la PAC</b><span>Salen solos los recintos con su cultivo, superficie y riego.</span></button>'+
      '<button class="opt" data-act="ap-mapa-nueva">'+ic('map')+'<b>Elegir en el mapa SIGPAC</b><span>Foto aérea y recintos: toca los tuyos.</span></button>'+
      '<button class="opt" data-act="ap-csv-pick">'+ic('xls')+'<b>Tabla de recintos (CSV)</b><span>Provincia, municipio, polígono, parcela y recinto.</span></button>'+
      '<button class="opt" data-act="ap-mano">'+ic('plus')+'<b>A mano</b><span>Con el código SIGPAC de cada recinto.</span></button>'+
      '<button class="opt" data-act="ap-ejemplo">'+ic('spark')+'<b>Probar con un ejemplo</b><span>Una finca inventada para verlo todo funcionando.</span></button>'+
      '<button class="opt" data-act="ap-abrir">'+ic('folder')+'<b>Abrir una copia</b><span>Un archivo guardado antes con esta herramienta.</span></button></div>'+
      '<p class="small center dropi">'+ic('up')+'O arrastra aquí el PDF de la PAC, una tabla CSV o una copia.</p></div>';
  }else{
    var q=noAcc(UI.q||'').toLowerCase().trim(), L=S.fincas.slice().sort(function(a,b){return String(a.nombre).localeCompare(String(b.nombre),'es')}).filter(function(f){if(!q)return true;var t=(f.datos||{}).tit||{};return noAcc([f.nombre,t.n,t.nif,t.mun].concat(fincaRecs(f.id).map(function(r){return refLabel(r.ref)+' '+(r.cultivo||'')})).join(' ')).toLowerCase().indexOf(q)>=0});
    h+='<div class="ih"><h1>Mis fincas</h1><div class="srch">'+ic('search')+'<input id="ap-q" class="in" placeholder="Buscar finca, titular, recinto o cultivo" value="'+esc(UI.q||'')+'" autocomplete="off" aria-label="Buscar"></div><div class="acts"><button class="btn primary" data-act="ap-nueva">'+ic('plus')+'Nueva finca</button><button class="btn" data-act="ap-abrir">'+ic('folder')+'Abrir una copia</button></div></div>';
    if(sinCopia&&!FS.ok)h+='<div class="warn">Hay cambios sin copia. Lo que metes se guarda solo en este navegador.<div class="acts" style="margin-top:8px">'+(fsCan()?'<button class="btn sm" data-act="ap-fs-pick">'+ic('save')+'Guardar solo en un archivo (recomendado)</button>':'')+'<button class="btn sm" data-act="ap-copia">Descargar una copia</button></div></div>';
    if(isDesk()){
      h+='<div class="xs tw"><table class="ftab"><thead><tr><th>Finca</th><th>Titular</th><th class="r">Recintos</th><th class="r">ha</th><th>Planes</th><th>Estado del plan</th><th>Última copia</th><th></th></tr></thead><tbody>'+L.map(function(f){var rs=fincaRecs(f.id),t=(f.datos||{}).tit||{},c=ctxOf(f),p=c.p;
        return '<tr class="tap" data-act="ap-open" data-id="'+esc(f.id)+'" tabindex="0"><td><b>'+esc(f.nombre)+'</b></td><td>'+esc(t.n||'—')+(t.nif?'<div class="small">'+esc(t.nif)+'</div>':'')+'</td><td class="r">'+rs.length+'</td><td class="r">'+ha2(rs.reduce(function(a,r){return a+haC(r)},0))+'</td>'+
          '<td>'+(pabList(f).map(function(x){return esc(pabCamp(x.y))}).join(', ')||'<span class="small">ninguno</span>')+'</td><td>'+(p?(c.F.length?'<span class="pa-st pa-warn">Faltan '+c.F.length+'</span> <span class="small">'+esc(pabCamp(p.y))+'</span>':'<span class="pa-st pa-ok">✓ Completo</span> <span class="small">'+esc(pabCamp(p.y))+'</span>'):'<span class="small">Sin plan</span>')+'</td>'+
          '<td class="small">'+(f.cp?new Date(f.cp).toLocaleDateString('es-ES')+' '+new Date(f.cp).toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'}):'nunca')+'</td>'+
          '<td class="nw"><button class="btn sm lk ib" data-act="ap-copiaf-id" data-id="'+esc(f.id)+'" title="Guardar copia de esta finca" aria-label="Copia">'+ic('save')+'</button><button class="btn sm lk danger ib" data-act="ap-fdel-id" data-id="'+esc(f.id)+'" title="Borrar la finca" aria-label="Borrar">'+ic('trash')+'</button></td></tr>'}).join('')+
        (L.length?'':'<tr><td colspan="8" class="small empty">Ninguna finca coincide con «'+esc(UI.q)+'».</td></tr>')+'</tbody></table></div>';
    }else{
      L.forEach(function(f){var rs=fincaRecs(f.id),t=(f.datos||{}).tit||{},c=ctxOf(f),p=c.p;
        h+='<div class="card tap" data-act="ap-open" data-id="'+esc(f.id)+'" role="button" tabindex="0"><h3>'+esc(f.nombre)+'</h3><div class="small">'+esc([t.n,rs.length+(rs.length===1?' recinto':' recintos'),haTxt(rs)].filter(Boolean).join(' · '))+'</div>'+
          '<div class="small" style="margin-top:4px">'+(p?'Plan '+esc(pabCamp(p.y))+': '+(c.F.length?'<span class="pa-st pa-warn">faltan '+c.F.length+'</span>':'<span class="pa-st pa-ok">✓ completo</span>'):'Sin plan todavía')+'</div></div>'});
    }
    h+='<p class="small dropi">'+ic('up')+'Arrastra aquí el PDF de la PAC, una tabla CSV o una copia para abrirlos.</p>';
  }
  h+='<details class="more"><summary>Cómo funciona</summary><div class="small"><p>1. <b>Recintos:</b> del PDF de la PAC, del mapa SIGPAC, de una tabla (CSV) o a mano con su código SIGPAC.</p><p>2. <b>Unidades:</b> los recintos se juntan solos por cultivo y riego; marca si están en zona vulnerable.</p><p>3. <b>Suelo y agua</b>, 4. <b>Cosecha</b> y 5. <b>Programa</b> de abonos (te lo puedo proponer yo).</p><p>6. <b>Revisar:</b> balance, comprobación del plan y PDF o Excel.</p><p>Los límites de nitrógeno son los del Decreto 80/2025 de Castilla-La Mancha. Para pistacho y nogal las extracciones son de criterio técnico (no hay tabla oficial). El plan lo firma el titular y, cuando sea obligatorio, el asesor.</p><p>En el iPhone: Compartir › «Añadir a pantalla de inicio» para tenerla como una app y que funcione sin cobertura.</p></div></details>';
  h+='<div class="foot"><a href="privacidad.html">Privacidad</a> · <a href="condiciones.html">Condiciones</a> · © 2026 Ramón Alcalá<br>Versión '+esc(VERSION)+'</div>';
  return h}

/* ───── retoques al HTML del plan: fechas propias, ayudas (?) y textos ───── */
function isoEs(v){var m=/^(\d{4})-(\d\d)-(\d\d)$/.exec(String(v||''));return m?m[3]+'/'+m[2]+'/'+m[1]:String(v||'')}
var HLPS=[['>¿Está en zona vulnerable a nitratos?</div>','zv'],['>Tipo de parcela (anejo III)</div>','tp'],['<th>Repr.</th>','rp'],['>Cultivo precedente</div>','prec'],['Materia orgánica <span class="small">(art. 4.3)</span>','mo'],['>Producción objetivo (kg/ha)</div>','pr'],['Aportes que ya hay (U.F./ha)','suelo'],['<th>Necesidad neta</th>','net'],['Asesor en fertilización ','as']];
function fixHTML(h){
  h=h.replace(/<input type="date" class="in pa-in" data-p="([^"]+)" value="([^"]*)"( aria-label="[^"]*")?>/g,function(m,p,v,al){return '<span class="dt"><input class="in pa-in dti" data-p="'+p+'" data-date="1" value="'+esc(isoEs(v))+'" placeholder="dd/mm/aaaa" inputmode="numeric" autocomplete="off"'+(al||' aria-label="Fecha"')+'><button type="button" class="dtb" data-act="ap-cal" aria-label="Elegir la fecha en el calendario">'+ic('cal')+'</button></span>'});
  h=h.replace(/pa-st pa-bad">(Déficit|Pendiente)/g,'pa-st pa-warn">$1');
  h=h.split('Ej. Ecofem (pellet orgánico)').join('Ej. Estiércol de oveja o compost');
  HLPS.forEach(function(x){var k=x[1],s=x[0];if(s.slice(-6)==='</div>')h=h.split(s).join(s.slice(0,-6)+hlp(k)+'</div>');else if(s.slice(-5)==='</th>')h=h.split(s).join(s.slice(0,-5)+' '+hlp(k)+'</th>');else h=h.split(s).join(s+hlp(k))});
  h=h.split('El abono orgánico que pongas en el programa no va aquí: ya cuenta como planificado.').join('El abono orgánico que pongas en el programa no va aquí: ya cuenta como planificado. <b>Suelo:</b> si lo dejas en 0 no se descuenta nada; en «Suelo y agua» te propongo una cifra cuando el análisis da P o K altos (criterio técnico).');
  return h}

/* ───── visor SIGPAC para añadir recintos: foto aérea PNOA de máxima actualidad, recintos y datos de cada recinto ───── */
var IGNW='https://www.ign.es/wmts/', MZ=15, MNZ=Math.pow(2,MZ), RT=6378137;
var MAP={m:null,lay:null,tiles:{},sel:null,info:{},busy:0,fail:0,places:null};
var refKey=function(p){return [p.provincia,p.municipio,p.agregado||0,p.zona||0,p.poligono,p.parcela,p.recinto].join(':')};
function m2ll(p){return [p[0]/RT*180/Math.PI,(2*Math.atan(Math.exp(p[1]/RT))-Math.PI/2)*180/Math.PI]}
function projG(g){var cv=function(r){return r.map(m2ll)};if(!g)return null;if(g.type==='Polygon')return {type:'Polygon',coordinates:g.coordinates.map(cv)};if(g.type==='MultiPolygon')return {type:'MultiPolygon',coordinates:g.coordinates.map(function(p){return p.map(cv)})};return null}
function wktGeo(w){var t=String(w||'').trim(),multi=/^MULTIPOLYGON/i.test(t),s=t.replace(/^[A-Za-z]+\s*/,'').replace(/\(/g,'[').replace(/\)/g,']').replace(/(-?[\d.eE+-]+)\s+(-?[\d.eE+-]+)/g,'[$1,$2]');try{var a=JSON.parse(s);return multi?{type:'MultiPolygon',coordinates:a}:{type:'Polygon',coordinates:a}}catch(e){return null}}
function mapaAbrir(){var f=finca(UI.fid);if(!f)return;sheetClose();
  var el=$('#mapa');el.hidden=false;document.body.classList.add('mapon');
  el.innerHTML='<div class="mtop"><div class="mtt"><b>Añadir recintos desde el mapa</b><span class="small">'+esc(f.nombre)+'</span></div><form class="msr" data-f="m-go"><input id="m-q" class="in" placeholder="Código SIGPAC (2:81:1:1:1) o municipio" autocomplete="off" aria-label="Buscar"><button class="btn" type="submit">'+ic('search')+'<span>Buscar</span></button></form>'+
    '<button class="btn" data-act="ap-m-loc" title="Mi ubicación">'+ic('loc')+'<span>Mi ubicación</span></button><button class="btn primary" data-act="ap-m-close">Hecho</button></div><div id="m-res" class="mres" hidden></div>'+
    '<div class="mbody"><div id="mmap" role="application" aria-label="Mapa"></div><div id="mhint" class="mhint" hidden></div><aside id="mside" class="mside"></aside></div>';
  mSide();
  ensureMap().then(mInit,function(){$('#mmap').innerHTML='<div class="note" style="margin:20px">Para el mapa hace falta internet.</div>'})}
function mapaCerrar(){var el=$('#mapa');if(MAP.m){try{MAP.m.remove()}catch(e){}}MAP.m=null;MAP.lay=null;MAP.tiles={};MAP.sel=null;MAP.fail=0;MAP.gen=(MAP.gen||0)+1;el.hidden=true;el.innerHTML='';document.body.classList.remove('mapon');render()}
function mInit(){var f=finca(UI.fid);if(!f||!$('#mmap'))return;
  var m=MAP.m=L.map('mmap',{minZoom:6,maxZoom:20,zoomControl:true,attributionControl:true,preferCanvas:true});m.attributionControl.setPrefix(false);
  L.tileLayer(IGNW+'pnoa-ma?service=WMTS&request=GetTile&version=1.0.0&format=image/jpeg&layer=OI.OrthoimageCoverage&style=default&tilematrixset=GoogleMapsCompatible&TileMatrix={z}&TileRow={y}&TileCol={x}',{maxNativeZoom:19,maxZoom:20,attribution:'Foto aérea: PNOA máxima actualidad © IGN · Recintos: SIGPAC (FEGA)'}).addTo(m);
  MAP.lay=L.geoJSON(null,{renderer:L.canvas({padding:0.4,tolerance:3}),style:mEstilo,onEachFeature:function(ft,l){l.on('click',function(ev){L.DomEvent.stop(ev);mSel(ft.properties,ev.latlng)})}}).addTo(m);
  m.on('moveend',mTiles);
  var pts=fincaRecs(f.id).filter(function(r){return r.ll}).map(function(r){return r.ll});
  if(pts.length){if(pts.length===1)m.setView(pts[0],17);else m.fitBounds(L.latLngBounds(pts).pad(0.25),{maxZoom:17})}
  else{var r0=fincaRecs(f.id)[0];m.setView([39.3,-2.6],8);if(r0)sigInfo(r0.ref).then(function(o){var ll=wktLL(o.wkt);if(ll&&MAP.m)MAP.m.setView(ll,16)}).catch(function(){});else mHint('Busca tu municipio o un código SIGPAC, o acércate en el mapa. Los recintos se ven a partir del zoom 15.')}
  setTimeout(function(){if(MAP.m)MAP.m.invalidateSize()},60)}
function mHint(t){var e=$('#mhint');if(!e)return;e.hidden=!t;e.textContent=t||''}
function mEstilo(ft){var p=ft.properties||{},k=refKey(p),r=S.recintos[k],mine=r&&r.fincaId===UI.fid,oth=r&&!mine&&finca(r.fincaId),sel=MAP.sel&&MAP.sel.ref===k;
  return {color:sel?'#ffd400':(mine?'#7cff9a':'#ffffff'),weight:sel?3.2:(mine?2.2:1.2),opacity:sel?1:0.9,fill:true,fillColor:mine?'#2fbf5a':(oth?'#8fa3ff':'#ffffff'),fillOpacity:mine?0.35:(oth?0.25:(sel?0.12:0.02))}}
function mTiles(){var m=MAP.m;if(!m)return;var z=m.getZoom();
  if(z<14){if(m.hasLayer(MAP.lay))m.removeLayer(MAP.lay);mHint('Acércate un poco más para ver los recintos.');return}
  if(!m.hasLayer(MAP.lay))MAP.lay.addTo(m);
  var b=m.getBounds(),tx=function(lon){return Math.floor((lon+180)/360*MNZ)},ty=function(lat){var r=lat*Math.PI/180;return Math.floor((1-Math.log(Math.tan(r)+1/Math.cos(r))/Math.PI)/2*MNZ)};
  var x0=tx(b.getWest()),x1=tx(b.getEast()),y0=ty(b.getNorth()),y1=ty(b.getSouth()),n=0,todo=[];
  for(var x=x0;x<=x1;x++)for(var y=y0;y<=y1;y++){var k=x+'.'+y;if(!MAP.tiles[k])todo.push([x,y,k])}
  if(todo.length>80){mHint('Acércate un poco más para ver los recintos.');return}
  mHint(MAP.fail?'El SIGPAC no responde ahora: prueba en un rato.':'');
  var gen=MAP.gen||0;todo.forEach(function(t){MAP.tiles[t[2]]=1;MAP.busy++;n++;
    fetch('https://sigpac.mapa.gob.es/vectorsdg/vector/recinto@3857/'+MZ+'.'+t[0]+'.'+(MNZ-1-t[1])+'.geojson').then(function(r){if(!r.ok)throw new Error(r.status);return r.json()})
      .then(function(j){if(!MAP.lay||(MAP.gen||0)!==gen)return;(j.features||[]).forEach(function(ft){var g=projG(ft.geometry);if(g)MAP.lay.addData({type:'Feature',properties:ft.properties||{},geometry:g})});MAP.fail=0})
      .catch(function(){delete MAP.tiles[t[2]];MAP.fail=1;mHint('El SIGPAC no responde ahora: prueba en un rato.')})
      .then(function(){MAP.busy--})})}
function mRestyle(){if(MAP.lay)MAP.lay.setStyle(mEstilo)}
function mSel(p,ll){var k=refKey(p);MAP.sel={ref:k,ha:(+p.dn_surface||0)/10000,uso:p.uso_sigpac||'',ll:ll?[Math.round(ll.lat*1e5)/1e5,Math.round(ll.lng*1e5)/1e5]:null};mRestyle();mSide();
  if(!MAP.info[k])sigInfo(k).then(function(o){MAP.info[k]=o;if(MAP.sel&&MAP.sel.ref===k)mSide()}).catch(function(){MAP.info[k]={err:1};if(MAP.sel&&MAP.sel.ref===k)mSide()})}
function mSide(){var el=$('#mside');if(!el)return;var st0=MAP.sel?MAP.sel.ref:'';if(el.dataset.ref!==st0){el.dataset.ref=st0;setTimeout(function(){el.scrollTop=0},0)}var f=finca(UI.fid),rs=fincaRecs(f.id),s=MAP.sel,h='';
  if(!s){h='<div class="k">Recintos</div><p class="small">Toca un recinto en el mapa para ver sus datos y añadirlo a la finca. Los de esta finca se ven en verde.</p>'}
  else{var o=MAP.info[s.ref]||{}, r=S.recintos[s.ref], mine=r&&r.fincaId===f.id, oth=r&&!mine?finca(r.fincaId):null, uso=o.uso_sigpac||s.uso, ha=o.superficie!=null?+o.superficie:s.ha;
    var dl=function(k,v){return v===''||v==null?'':'<div><dt>'+k+'</dt><dd>'+v+'</dd></div>'};
    h='<div class="k">Recinto SIGPAC</div><h3>'+esc(refLabel(s.ref))+'</h3><dl class="mdl">'+dl('Uso',esc(usoTxt(uso)))+dl('Superficie',ha2(ha)+' ha')+
      (o.err?'<div class="small">No he podido leer el resto de datos del SIGPAC.</div>':!o.provincia?'<div class="small">Leyendo los datos del SIGPAC…</div>':
        dl('Pendiente media',o.pendiente_media!=null?trTxt(+o.pendiente_media,1)+' %':'')+dl('Coef. de regadío',o.coef_regadio!=null?trTxt(+o.coef_regadio,0)+' %':'')+dl('Admisibilidad',o.admisibilidad!=null?trTxt(+o.admisibilidad,0)+' %':'')+dl('Altitud',+o.altitud?nf(+o.altitud,0)+' m':'')+dl('Incidencias',o.incidencias?esc(o.incidencias):''))+'</dl>';
    if(mine)h+='<div class="okst">'+ic('check')+'Ya está en esta finca</div><div class="acts"><button class="btn sm danger lk" data-act="ap-m-del">Quitarlo de la finca</button></div>';
    else if(oth)h+='<div class="note">Está en la finca «'+esc(oth.nombre)+'».</div>';
    else{var rg=+o.coef_regadio>0?'regadio':'secano';
      h+='<div class="pa-g one"><div>'+pabF('Cultivo','<select id="m-cult" class="trsel">'+cultOpts(cultDeUso(uso))+'</select>')+'</div><div>'+pabF('Riego','<select id="m-rg" class="trsel"><option value="secano">Secano</option><option value="regadio"'+(rg==='regadio'?' selected':'')+'>Regadío</option></select>')+'</div></div>'+
        '<div class="acts"><button class="btn primary wide" data-act="ap-m-add">'+ic('plus')+'Añadir a la finca</button></div>'}}
  h+='<div class="msum"><b>'+rs.length+(rs.length===1?' recinto':' recintos')+'</b> en «'+esc(f.nombre)+'» · '+esc(haTxt(rs))+'</div>';
  el.innerHTML=h}
function mAdd(){var f=finca(UI.fid),s=MAP.sel;if(!f||!s)return;var o=MAP.info[s.ref]||{};
  undoPush();var x={ref:s.ref,ha:o.superficie!=null?+o.superficie:s.ha,cultivo:($('#m-cult')||{}).value||'',riego:($('#m-rg')||{}).value||'secano',uso:s.uso,ll:s.ll};
  var R=addRecs(f,[x]);if(!R.added){toast('Ese recinto ya está en otra finca.');return}
  var r=S.recintos[s.ref];if(o.provincia)sigAplica(r,o);if(!r.ll&&s.ll)r.ll=s.ll;pendAPlanes(f);climaAuto(f);save();mRestyle();mSide();toast('Recinto '+refLabel(s.ref)+' añadido','Deshacer',function(){undoDo();mRestyle();mSide()});
  if(!o.provincia)sigLeer(f,[s.ref])}
function mDel(){var f=finca(UI.fid),s=MAP.sel;if(!f||!s)return;recQuitar(f,[s.ref]);mRestyle();mSide();toast('Recinto quitado','Deshacer',function(){undoDo();mRestyle();mSide()})}
function mGo(t){t=String(t||'').trim();if(!t)return;var ref=refParse(t),res=$('#m-res');
  if(ref){toast('Buscando el recinto…');sigInfo(ref).then(function(o){var g=wktGeo(o.wkt);if(!g||!MAP.m)throw 1;toastOff();var lay=L.geoJSON(g);MAP.m.fitBounds(lay.getBounds().pad(0.6),{maxZoom:18});MAP.info[ref]=o;mSel({provincia:o.provincia,municipio:o.municipio,agregado:o.agregado,zona:o.zona,poligono:o.poligono,parcela:o.parcela,recinto:o.recinto,dn_surface:(+o.superficie||0)*10000,uso_sigpac:o.uso_sigpac},null);var ll=wktLL(o.wkt);if(ll)MAP.sel.ll=ll})
    .catch(function(){toast('No encuentro ese recinto en el SIGPAC. Revisa los números.')});return}
  res.hidden=false;res.innerHTML='<div class="small">Buscando «'+esc(t)+'»…</div>';
  fetch('https://www.cartociudad.es/geocoder/api/geocoder/candidates?limit=8&q='+encodeURIComponent(t)).then(function(r){return r.json()}).then(function(a){
    var L0=(Array.isArray(a)?a:[]).filter(function(c){return c&&c.address});MAP.places=L0;
    res.innerHTML=L0.length?L0.map(function(c,i){return '<button class="mr" data-act="ap-m-place" data-i="'+i+'"><b>'+esc(c.address)+'</b><span class="small">'+esc([c.type,c.province].filter(Boolean).join(' · '))+'</span></button>'}).join(''):'<div class="small">No encuentro ese sitio.</div>'})
    .catch(function(){res.innerHTML='<div class="small">No puedo buscar sitios ahora. Prueba con un código SIGPAC.</div>'})}
function mPlace(i){var c=(MAP.places||[])[i],res=$('#m-res');if(!c||!MAP.m)return;res.hidden=true;var ty=String(c.type||''),z=/provinc/i.test(ty)?10:/municip|poblac/i.test(ty)?14:17;
  if(c.lat&&c.lng){MAP.m.setView([+c.lat,+c.lng],z);return}
  fetch('https://www.cartociudad.es/geocoder/api/geocoder/find?outputformat=json&q='+encodeURIComponent(c.address)+'&type='+encodeURIComponent(c.type||'')+'&id='+encodeURIComponent(c.id||'')).then(function(r){return r.json()}).then(function(j){if(j&&j.lat&&j.lng)MAP.m.setView([+j.lat,+j.lng],z);else throw 1}).catch(function(){toast('No he podido situar ese lugar.')})}
function mLoc(){if(!navigator.geolocation||!MAP.m){toast('Este navegador no da la ubicación.');return}toast('Buscando tu ubicación…');
  navigator.geolocation.getCurrentPosition(function(p){toastOff();MAP.m.setView([p.coords.latitude,p.coords.longitude],17)},function(){toast('No he podido ver tu ubicación. Mira el permiso del navegador.')},{enableHighAccuracy:true,timeout:15000,maximumAge:30000})}

/* ───── acciones ───── */
function recQuitar(f,refs){undoPush();refs.forEach(function(ref){delete S.recintos[ref];delete SEL[ref];(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){u.refs=(u.refs||[]).filter(function(x){return x!==ref});if(u.rc)delete u.rc[ref];if(u.rp===ref)u.rp=u.refs[0]||''})})});save();render()}
function fincaBorrar(f){undoPush();Object.keys(S.recintos).forEach(function(k){if(S.recintos[k].fincaId===f.id)delete S.recintos[k]});S.fincas=S.fincas.filter(function(x){return x.id!==f.id});if(UI.fid===f.id){UI.v='inicio';UI.fid=null}save();saveUI();render();undoToast('Finca «'+f.nombre+'» borrada')}
function irPaso(s){UI.step=s;saveUI();window.__scrollTop=true;render()}
function irFalta(s){if(!CTX)return;var x=s?CTX.F.filter(function(y){return y.s===s})[0]:CTX.F[0];if(!x){toast('No falta nada en este paso');return}
  UI.step=x.s;if(x.u!=null&&CTX.p){var u=CTX.p.ud[x.u];if(u){PAB.open[u.id]=true;UI.bu=u.id}}
  if(/tit\./.test(x.sel))UI.titOpen=true; if(/cl\./.test(x.sel))PAB.open.cl=true; if(/AS\.|"fe"/.test(x.sel))PAB.open.fi=true;
  saveUI();render();setTimeout(function(){var el=x.sel?$(x.sel):null;if(!el&&x.u!=null)el=$('#u'+x.u);if(!el)return;el.scrollIntoView({block:'center',behavior:'smooth'});flash(el);if(el.matches('input,select,textarea'))try{el.focus({preventScroll:true})}catch(e){}},40)}
function flash(el){el.classList.remove('flash');void el.offsetWidth;el.classList.add('flash');setTimeout(function(){el.classList.remove('flash')},1600)}
function renderBal(){var f=finca(UI.fid);if(!f||!CTX)return;$('#bal').innerHTML=balHTML(f,CTX)}
function pdfGo(f,p){toast('Preparando el PDF…');ensurePdf().then(function(){toastOff();pabOut(f,p,'pdf')},function(){pdfNoListo()})}
function nuevaHoja(){sheet('<h2>Nueva finca</h2><div class="opts two">'+
  '<button class="opt" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Leer el PDF de la PAC</b><span>Los recintos salen solos.</span></button><button class="opt" data-act="ap-mapa-nueva">'+ic('map')+'<b>En el mapa SIGPAC</b><span>Toca tus recintos en la foto aérea.</span></button>'+
  '<button class="opt" data-act="ap-csv-pick">'+ic('xls')+'<b>Tabla de recintos (CSV)</b><span>Provincia, municipio, polígono, parcela y recinto.</span></button><button class="opt" data-act="ap-mano">'+ic('plus')+'<b>A mano</b><span>Con el código SIGPAC.</span></button>'+
  '<button class="opt" data-act="ap-ejemplo">'+ic('spark')+'<b>Finca de ejemplo</b><span>Inventada, para probar.</span></button></div><div class="acts"><button class="btn lk" data-act="ap-close">Cancelar</button></div>','wide')}
function copiasHoja(){var h='<h2>Copias</h2><p class="small">Lo que metes se guarda en este navegador. Una copia en archivo te protege si se borran los datos del navegador o cambias de ordenador.</p><div class="opts two">';
  if(fsCan())h+=FS.h?'<button class="opt" data-act="ap-fs-stop">'+ic('close')+'<b>Dejar de guardar en «'+esc(FS.name)+'»</b><span>Se sigue guardando en el navegador.</span></button>':'<button class="opt" data-act="ap-fs-pick">'+ic('save')+'<b>Guardar solo en un archivo</b><span>Eliges el archivo una vez y cada cambio se guarda ahí (recomendado).</span></button>';
  h+='<button class="opt" data-act="ap-copia">'+ic('up')+'<b>Descargar una copia ahora</b><span>Un archivo con todas tus fincas.</span></button><button class="opt" data-act="ap-abrir">'+ic('folder')+'<b>Abrir una copia</b><span>Recuperar fincas de un archivo.</span></button></div><div class="acts"><button class="btn" data-act="ap-close">Cerrar</button></div>';sheet(h,'wide')}
function misHoja(){var L=S.x.abonos||[];
  var h='<h2>Mis abonos</h2><p class="small">Los abonos que usas, con la riqueza de su etiqueta y, si quieres, el precio. Salen los primeros al buscar en el programa. El precio solo se ve en pantalla.</p><div class="xs"><table class="sheet ma"><thead><tr><th>Nombre</th><th class="r">N %</th><th class="r">P₂O₅ %</th><th class="r">K₂O %</th><th class="c">Org.</th><th class="r">€/kg</th><th class="r">Densidad (kg/L)</th><th></th></tr></thead><tbody>'+
    L.map(function(a,i){var c=function(k,ph,n){return '<td><input class="in cell ma-in'+(n?' n':'')+'" data-i="'+i+'" data-k="'+k+'" value="'+esc(a[k]==null?'':(typeof a[k]==='number'?abEd(a[k]):a[k]))+'" placeholder="'+(ph||'')+'"'+(n?' inputmode="decimal"':'')+' autocomplete="off"></td>'};
      return '<tr>'+c('n','Nombre del abono')+c('rn','0',1)+c('rp','0',1)+c('rk','0',1)+'<td class="c"><input type="checkbox" class="ma-in" data-i="'+i+'" data-k="org"'+(a.org?' checked':'')+'></td>'+c('pz','0,00',1)+c('dens','',1)+'<td><button class="btn sm lk danger ib" data-act="ap-ma-del" data-i="'+i+'" aria-label="Quitar">'+ic('close')+'</button></td></tr>'}).join('')+
    (L.length?'':'<tr><td colspan="8" class="small empty">Aún no hay abonos tuyos.</td></tr>')+'</tbody></table></div><div class="acts"><button class="btn" data-act="ap-ma-add">'+ic('plus')+'Añadir abono</button><button class="btn primary" data-act="ap-ma-ok">Hecho</button></div>';
  sheet(h,'wide')}
function popAt(btn,html,cls){var p=$('#pop');p.className=cls||'';p.innerHTML=html;p.hidden=false;var r=btn.getBoundingClientRect(),w=p.offsetWidth,hh=p.offsetHeight,x=Math.min(window.innerWidth-w-8,Math.max(8,r.left+r.width/2-w/2)),y=r.bottom+6;if(y+hh>window.innerHeight-8)y=Math.max(8,r.top-hh-6);p.style.left=x+'px';p.style.top=y+'px';POPB=btn}
var POPB=null;function popClose(){var p=$('#pop');if(p){p.hidden=true;p.innerHTML=''}POPB=null}
/* calendario propio (fecha siempre dd/mm/aaaa) */
var CAL={inp:null,y:0,m:0};
function calHTML2(){var y=CAL.y,m=CAL.m,d1=new Date(y,m,1),off=(d1.getDay()+6)%7,nd=new Date(y,m+1,0).getDate(),sel=pabIso(CAL.inp&&CAL.inp.value),hoy=hoyIso(),h='<div class="calp"><div class="calh"><button type="button" data-act="ap-calm" data-d="-1" aria-label="Mes anterior">‹</button><b>'+MESES[m]+' '+y+'</b><button type="button" data-act="ap-calm" data-d="1" aria-label="Mes siguiente">›</button></div><div class="calg">'+['L','M','X','J','V','S','D'].map(function(x){return '<i>'+x+'</i>'}).join('');
  for(var k=0;k<off;k++)h+='<span></span>';for(var d=1;d<=nd;d++){var iso=y+'-'+('0'+(m+1)).slice(-2)+'-'+('0'+d).slice(-2);h+='<button type="button" data-act="ap-cald" data-v="'+iso+'" class="'+(iso===sel?'on':'')+(iso===hoy?' hoy':'')+'">'+d+'</button>'}
  return h+'</div><div class="acts"><button type="button" class="btn sm" data-act="ap-cald" data-v="'+hoy+'">Hoy</button><button type="button" class="btn sm lk" data-act="ap-cald" data-v="">Borrar</button></div></div>'}

document.addEventListener('click',function(ev){
  var pp=$('#pop'); if(pp&&!pp.hidden&&!pp.contains(ev.target)&&!(POPB&&POPB.contains(ev.target)))popClose();
  if(!(ev.target.closest&&ev.target.closest('.sgw')))sgClose();
  var b=ev.target.closest&&ev.target.closest('[data-act]'); if(!b||b.disabled)return; var a=b.dataset.act, f=finca(UI.fid);
  if(a==='pa-new-ok'){if(!f)return;var sy=$('#pa-ny'),y=sy?+sy.value:pabY0();if((f.pab||[]).some(function(x){return +x.y===y})){toast('Ya hay un plan de la campaña '+pabCamp(y));return}
    undoPush();var n=planNuevo(f,y);if(UI.step==='rec')UI.step='ud';save();saveUI();window.__scrollTop=true;render();undoToast('Plan '+pabCamp(n.y)+' creado con '+n.ud.length+(n.ud.length===1?' unidad':' unidades'));return}
  if(/^pa-/.test(a)){if(!f)return;EXP.fid=f.id;pabClick(b);return}
  if(a==='ap-help'){ev.preventDefault();var g=GLOS[b.dataset.k];if(g){if(POPB===b){popClose();return}popAt(b,'<b>'+esc(g[0])+'</b><p>'+esc(g[1])+'</p>'+(g[2]?'<div class="small">'+esc(g[2])+'</div>':''),'hlpp')}return}
  if(a==='ap-cal'){var inp=b.previousElementSibling;if(POPB===b){popClose();return}CAL.inp=inp;var iso=pabIso(inp.value)||hoyIso();CAL.y=+iso.slice(0,4);CAL.m=+iso.slice(5,7)-1;popAt(b,calHTML2(),'calpop');return}
  if(a==='ap-calm'){CAL.m+=+b.dataset.d;if(CAL.m<0){CAL.m=11;CAL.y--}if(CAL.m>11){CAL.m=0;CAL.y++}$('#pop').innerHTML=calHTML2();return}
  if(a==='ap-cald'){var ci=CAL.inp;popClose();if(!ci)return;ci.value=b.dataset.v?isoEs(b.dataset.v):'';ci.dispatchEvent(new Event('change',{bubbles:true}));return}
  if(a==='ap-close'){sheetClose();return}
  if(a==='ap-undo'){undoDo();return}
  if(a==='ap-open'){if(ev.target.closest('button')&&ev.target.closest('button')!==b)return;UI=Object.assign(UI,{v:'finca',fid:b.dataset.id,step:'rec',bu:null});var f2=finca(b.dataset.id);if(f2&&(f2.pab||[]).length)UI.step=ctxOf(f2).F.length?ctxOf(f2).F[0].s:'rv';SEL={};saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-home'){UI.v='inicio';UI.fid=null;saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-step'){irPaso(b.dataset.s);return}
  if(a==='ap-prev'){var i=stepIdx();if(i===0){UI.v='inicio';UI.fid=null;saveUI();render()}else irPaso(STEPS[i-1][0]);return}
  if(a==='ap-next'){irPaso(STEPS[Math.min(5,stepIdx()+1)][0]);return}
  if(a==='ap-nueva'){nuevaHoja();return}
  if(a==='ap-mano'){undoPush();var nf1=fincaNueva('Mi finca');UI.v='finca';UI.fid=nf1.id;UI.step='rec';UI.titOpen=true;save();saveUI();sheetClose();render();addRecHoja();return}
  if(a==='ap-mapa-nueva'){undoPush();var nf2=fincaNueva('Mi finca');UI.v='finca';UI.fid=nf2.id;UI.step='rec';save();saveUI();sheetClose();render();mapaAbrir();return}
  if(a==='ap-ejemplo'){sheetClose();ejemplo();return}
  if(a==='ap-pdf-pick'){pick('application/pdf,.pdf',pacLeer);return}
  if(a==='ap-csv-pick'){pick('.csv,.txt,text/csv,text/plain',csvLeer);return}
  if(a==='ap-abrir'){pick('.json,application/json',copiaAbrir);return}
  if(a==='ap-copias'){copiasHoja();return}
  if(a==='ap-copia'){sheetClose();copiaTodo();return}
  if(a==='ap-fs-pick'){sheetClose();fsPick();return}
  if(a==='ap-fs-res'){fsResume();return}
  if(a==='ap-fs-stop'){sheetClose();fsStop();return}
  if(a==='ap-copia-join'){copiaAplicar(false);return}
  if(a==='ap-copia-repl'){copiaAplicar(true);return}
  if(a==='ap-pac-tog'){var r=PACI&&PACI.rows[+b.dataset.i];if(r){r.off=!r.off;var sc=$('#sheet .box').scrollTop;pacHoja();$('#sheet .box').scrollTop=sc}return}
  if(a==='ap-pac-go'){pacGo();return}
  if(a==='ap-dl-save'){dlSave();return}
  if(a==='ap-dl-share'){dlShare();return}
  if(a==='ap-dl-view'){dlView();return}
  if(a==='ap-copiaf-id'){ev.stopPropagation();var fc=finca(b.dataset.id);if(fc)copiaFinca(fc);return}
  if(a==='ap-fdel-id'){ev.stopPropagation();var fd=finca(b.dataset.id);if(fd)fincaBorrar(fd);return}
  if(a==='ap-mis'){misHoja();return}
  if(a==='ap-ma-add'){(S.x.abonos=S.x.abonos||[]).push({n:'',rn:0,rp:0,rk:0});saveRaw();misHoja();var ins=$$('#sheet .ma-in[data-k="n"]');if(ins.length)ins[ins.length-1].focus();return}
  if(a==='ap-ma-del'){undoPush();S.x.abonos.splice(+b.dataset.i,1);save();misHoja();return}
  if(a==='ap-ma-ok'){S.x.abonos=(S.x.abonos||[]).filter(function(x){return String(x.n||'').trim()});save();sheetClose();render();return}
  if(a==='ap-baltog'){UI.balOpen=!UI.balOpen;saveUI();renderBal();return}
  if(a==='ap-sort'){var k=b.dataset.k;UI.sort=(UI.sort===k?'-'+k:k);saveUI();render();return}
  if(a==='ap-compra-copy'){var fp=f&&pabCur(f);if(!fp)return;var L=compra(f,fp,CTX.Cs);var txt=L.map(function(x){return x.n+': '+(x.kg>=1000?nf(x.kg/1000,2)+' t':nf(x.kg,0)+' kg')}).join('\n');
    (navigator.clipboard?navigator.clipboard.writeText(txt):Promise.reject()).then(function(){toast('Lista copiada')},function(){toast('No he podido copiar la lista')});return}
  if(a==='ap-m-close'){mapaCerrar();return}
  if(a==='ap-m-add'){mAdd();return}
  if(a==='ap-m-del'){mDel();return}
  if(a==='ap-m-loc'){mLoc();return}
  if(a==='ap-m-place'){mPlace(+b.dataset.i);return}
  if(!f)return;
  var p=pabCur(f);
  if(a==='ap-addrec'){addRecHoja();return}
  if(a==='ap-ref-add'){refAdd();return}
  if(a==='ap-mapa'){mapaAbrir();return}
  if(a==='ap-titular'){UI.step='rec';UI.titOpen=true;saveUI();render();var tt=$('#titular');if(tt){tt.scrollIntoView({block:'start',behavior:'smooth'});flash(tt)}return}
  if(a==='ap-sig'){toast('Leyendo el SIGPAC…');sigLeer(f,fincaRecs(f.id).map(function(r){return r.ref}),function(ok,ko){toast('SIGPAC: '+ok+' recintos leídos'+(ko?', '+ko+' sin respuesta':''))});return}
  if(a==='ap-recdel'){recQuitar(f,[b.dataset.ref]);undoToast('Recinto '+refLabel(b.dataset.ref)+' quitado');return}
  if(a==='ap-bulk-del'){var ks=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id});if(!ks.length)return;recQuitar(f,ks);undoToast(ks.length+(ks.length===1?' recinto quitado':' recintos quitados'));return}
  if(a==='ap-bulk-none'){SEL={};render();return}
  if(a==='ap-copiaf'){copiaFinca(f);return}
  if(a==='ap-fdel'){fincaBorrar(f);return}
  if(a==='ap-newplan-no'){PAB.ask=null;render();return}
  if(a==='ap-falta'){irFalta(b.dataset.s||'');return}
  if(a==='ap-goto'){var gu=p&&p.ud[+b.dataset.u];if(gu){PAB.open[gu.id]=true;UI.bu=gu.id}UI.step=b.dataset.s;saveUI();render();var ge=$('#u'+b.dataset.u);if(ge){ge.scrollIntoView({block:'start',behavior:'smooth'})}return}
  if(a==='ap-sigcamp'){if(!p)return;var y1=+p.y+1,ex=(f.pab||[]).filter(function(x){return +x.y===y1})[0];if(ex){PAB.cur[f.id]=ex.id;render();toast('Plan de la campaña '+pabCamp(y1));return}
    undoPush();var np=planNuevo(f,y1,p);UI.step='ud';save();saveUI();window.__scrollTop=true;render();undoToast('Plan '+pabCamp(y1)+' hecho a partir del de '+pabCamp(p.y)+': revisa la cosecha y el programa');return}
  if(a==='ap-pdf'||a==='ap-xlsx'){if(!p){toast('Primero crea el plan');return}if(a==='ap-xlsx'){pabOut(f,p,'xlsx');return}pdfGo(f,p);return}
  if(a==='ap-prevupd'){PREV.mod=-1;var pb=$('#prevb');if(pb){pb.dataset.u='';pb.innerHTML='<div class="small">Preparando la vista previa…</div>'}prevMake();return}
  if(!p)return;
  var ui=+b.dataset.u, u=p.ud[ui], C=u&&pabCalc(f,u,p);
  if(a==='ap-bu'){if(u){UI.bu=u.id;saveUI();renderBal()}return}
  if(!u)return;
  if(a==='ap-prop'){var rows=propuesta(f,u,C);if(!rows.length){toast('Con lo que ya aporta, esta unidad no necesita abono (necesidad neta 0).');return}
    var had=((u.pg||{}).rows||[]).length;undoPush();u.pg={src:'own',rows:rows};delete u.d0;mitPropuesta(u);p.m=Date.now();save();render();undoToast('Programa propuesto con '+rows.length+(rows.length===1?' abono':' abonos')+(had?' (cambia el que había)':'')+'. Revísalo y cambia lo que quieras');return}
  if(a==='ap-fix'){var fx=arreglo(f,u,C,b.dataset.k);if(!fx)return;undoPush();u.pg=u.pg||{src:'own',rows:[]};u.pg.rows=u.pg.rows||[];if(fx.t==='add')u.pg.rows.push(fx.q);else u.pg.rows[fx.j].d=fx.d;p.m=Date.now();save();render();undoToast(fx.t==='add'?'Aporte añadido: '+fx.txt:'Hecho: '+fx.txt);return}
  if(a==='ap-mitprop'){undoPush();var nm=mitPropuesta(u);p.m=Date.now();save();render();toast(nm?nm+(nm===1?' medida marcada':' medidas marcadas'):'Ya estaban marcadas las que encajan');return}
  if(a==='ap-resall'){undoPush();u.res=u.res||{};PAB_RES.forEach(function(x){u.res[x[0]]=Object.assign({},u.res[x[0]],{a:1})});delete u.resProp;p.m=Date.now();save();render();return}
  if(a==='ap-cred'){var cr=creditoSuelo(u,C);undoPush();u.ap=u.ap||{};u.ap.s=u.ap.s||{};if(cr.p)u.ap.s.p=cr.p;if(cr.k)u.ap.s.k=cr.k;p.m=Date.now();save();render();toast('Puesto como aporte del suelo (criterio técnico)');return}
  if(a==='ap-rowmore'){RM[b.dataset.id]=!RM[b.dataset.id];render();return}
});
document.addEventListener('submit',function(ev){var fm=ev.target;if(fm&&fm.dataset&&fm.dataset.f==='m-go'){ev.preventDefault();mGo(($('#m-q')||{}).value)}});
function pick(accept,fn){var i=document.createElement('input');i.type='file';i.accept=accept;i.style.display='none';document.body.appendChild(i);i.onchange=function(){var fl=i.files&&i.files[0];i.remove();if(fl)fn(fl)};i.click()}

/* deshacer también lo escrito en las casillas: se guarda el estado justo antes de cada cambio */
document.addEventListener('change',function(ev){var t=ev.target;if(!t||!t.closest||!t.closest('#app,#sheet,#bal'))return;if(t.id==='ap-q'||(t.classList&&/\b(rsel|rsel-all|ap-camp|ap-cost|ap-bulk|ap-clprov|ap-copyud)\b/.test(t.className)))return;if(t.dataset.p||t.dataset.e||t.dataset.r||t.classList.contains('ap-in')||t.classList.contains('ma-in'))undoPush()},true);
document.addEventListener('click',function(ev){var b=ev.target.closest&&ev.target.closest('[data-act^="pa-"]');if(b)undoPush()},true);
/* fechas: lo escrito como dd/mm/aaaa se guarda como fecha */
document.addEventListener('change',function(ev){var t=ev.target;if(!t||!t.dataset||!t.dataset.date)return;var v=String(t.value||'').trim();if(!v)return;
  var m=/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/.exec(v),iso='';
  if(m){var y=+m[3];if(y<100)y+=2000;var d=new Date(y,+m[2]-1,+m[1]);if(d.getFullYear()===y&&d.getMonth()===+m[2]-1&&d.getDate()===+m[1])iso=y+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2)}
  else if(/^\d{4}-\d\d-\d\d$/.test(v)){var dd=new Date(v+'T12:00:00');if(!isNaN(dd)&&dd.toISOString().slice(0,10)===v)iso=v}
  if(!iso){ev.stopImmediatePropagation();t.classList.add('bad');toast('Pon la fecha como día/mes/año, por ejemplo 15/03/2027.');return}
  t.classList.remove('bad');t.value=iso},true);
/* casillas propias (finca, titular, recintos), campaña, lluvia, marcas, costes */
document.addEventListener('change',function(ev){
  var t=ev.target; if(!t||!t.classList)return;
  if(t.classList.contains('ma-in')){var a=(S.x.abonos||[])[+t.dataset.i];if(!a)return;var k=t.dataset.k;if(k==='org')a.org=t.checked?1:0;else if(k==='n')a.n=t.value.trim();else{var vn=t.value.trim()===''?undefined:numEs(t.value);if(vn===null||isNaN(vn)){toast('Pon un número (decimales con coma).');t.classList.add('bad');return}t.classList.remove('bad');if(vn===undefined)delete a[k];else a[k]=vn}save();return}
  var f=finca(UI.fid); if(!f)return;
  if(t.classList.contains('ap-in')){
    var v=String(t.value==null?'':t.value).trim();
    if(t.dataset.num){if(v===''){v=undefined}else{var n=numEs(v);if(n===null||isNaN(n)||n<0){toast('Pon un número (los decimales con coma, por ejemplo 2,35).');t.classList.add('bad');return}v=n}}
    if(t.dataset.e){var ks=t.dataset.e.split('.');if(ks[0]==='nombre'){if(!v){toast('La finca necesita un nombre.');render();return}f.nombre=v}else{f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};if(v)f.datos.tit[ks[1]]=v;else delete f.datos.tit[ks[1]]}}
    else if(t.dataset.r){var r=S.recintos[t.dataset.r];if(!r)return;if(v===undefined||v==='')delete r[t.dataset.k];else r[t.dataset.k]=v}
    save(); pabRedraw(); return}
  if(t.classList.contains('rsel')){SEL[t.dataset.ref]=t.checked;render();return}
  if(t.classList.contains('rsel-all')){var on=t.checked;SEL={};if(on)fincaRecs(f.id).forEach(function(r){SEL[r.ref]=true});render();return}
  if(t.classList.contains('ap-bulk')){var ks2=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id});if(!ks2.length||!t.value)return;undoPush();ks2.forEach(function(k){S.recintos[k][t.dataset.k]=t.value});save();render();undoToast(ks2.length+(ks2.length===1?' recinto cambiado':' recintos cambiados'));return}
  if(t.classList.contains('ap-camp')){if(t.value==='new'){PAB.ask=f.id;if(UI.step==='rec')UI.step='ud'}else{PAB.cur[f.id]=t.value;PAB.ask=null}render();return}
  var p=pabCur(f);
  if(t.classList.contains('ap-cost')){UI.cost=t.checked?1:0;saveUI();render();return}
  if(!p)return;
  if(t.classList.contains('ap-clprov')){var c=CLIMA[+t.value];if(c){undoPush();p.cl={a:c.a,o:c.o,i:c.i,p:c.p,v:c.v,src:c.src};delete p.clAuto;p.m=Date.now();save();render();toast('Lluvia de '+c.n+' puesta')}return}
  if(t.classList.contains('pa-chk')){var path=t.dataset.p;pabChange(f,path,t.checked?1:0,false);var mm=/^ud\.(\d+)\.res\./.exec(path);if(mm&&p.ud[+mm[1]])delete p.ud[+mm[1]].resProp;save();pabRedraw();return}
  if(t.classList.contains('ap-copyud')){var src=p.ud[+t.value],dst=p.ud[+t.dataset.u];if(!src||!dst)return;undoPush();dst.pg=dst.pg||{src:'own',rows:[]};dst.pg.rows=(dst.pg.rows||[]).concat(((src.pg||{}).rows||[]).map(function(q){var c2=JSON.parse(JSON.stringify(q));c2.id=uid();return c2}));delete dst.d0;p.m=Date.now();save();render();undoToast('Programa de «'+src.n+'» copiado');return}
  if(t.classList.contains('pa-in')&&/^cl\./.test(t.dataset.p||'')&&p.clAuto){delete p.clAuto;saveRaw()}
});
document.addEventListener('input',function(ev){var t=ev.target;if(!t)return;
  if(t.id==='ap-q'){UI.q=t.value;saveUI();var pos=t.selectionStart;render();var q=$('#ap-q');if(q){q.focus();try{q.setSelectionRange(pos,pos)}catch(e){}}return}
  if(t.classList&&t.classList.contains('sgi')){sgOpen(t);return}
  if(t.classList&&(t.classList.contains('pa-in')||t.dataset.num)){var w=valida(t.dataset.p||'',t.value,filaDe(t.dataset.p||''),t.getAttribute('inputmode')==='decimal');t.classList.toggle('rare',!!w);if(w)t.title=w}});
/* la unidad en la que estás manda en el panel de balance */
document.addEventListener('focusin',function(ev){var t=ev.target;if(!t||!t.dataset)return;var m=/^ud\.(\d+)\./.exec(t.dataset.p||'');
  if(m&&CTX&&CTX.p){var u=CTX.p.ud[+m[1]];if(u&&UI.bu!==u.id){UI.bu=u.id;renderBal()}}
  if(t.classList.contains('sgi')&&!t.value)sgOpen(t)});
document.addEventListener('toggle',function(ev){var t=ev.target;if(!t||!t.id||!/^u\d+$/.test(t.id)||!t.open||!CTX||!CTX.p)return;var u=CTX.p.ud[+t.id.slice(1)];if(u&&UI.bu!==u.id){UI.bu=u.id;renderBal()}if(t.id&&t.dataset.paO==='tit')UI.titOpen=t.open},true);
document.addEventListener('toggle',function(ev){var t=ev.target;if(t&&t.id==='titular'){UI.titOpen=t.open;saveUI()}},true);

/* ───── lista de sugerencias propia (abonos y meses), dentro de la página bajo la casilla ───── */
var SG={box:null,inp:null,L:[],hi:-1};
var MES_SG=MESES.concat(['Febrero-marzo','Marzo-abril','Abril-junio','Marzo-junio','Mayo-julio','Octubre-noviembre']);
function sgOpen(inp){var kind=inp.dataset.sg,q=noAcc(inp.value||'').toLowerCase().trim(),L;
  if(kind==='ab')L=catBusca(inp.value).slice(0,60);else L=MES_SG.filter(function(m){return !q||noAcc(m).toLowerCase().indexOf(q)===0}).map(function(m){return {n:m}});
  if(SG.inp!==inp)sgClose();SG.inp=inp;SG.L=L;SG.hi=-1;
  if(!L.length){if(SG.box){SG.box.remove();SG.box=null}return}
  if(!SG.box){SG.box=document.createElement('div');SG.box.className='sgl';SG.box.setAttribute('role','listbox');inp.parentNode.appendChild(SG.box)}
  var g0='';SG.box.innerHTML=L.map(function(x,i){var gh='';if(kind==='ab'&&x.g!==g0){g0=x.g;gh='<div class="sgg">'+esc(x.g)+'</div>'}
    return gh+'<div class="sgo" role="option" data-i="'+i+'"><span>'+esc(x.n)+'</span>'+(kind==='ab'?'<em>'+trTxt(x.rn,1)+'-'+trTxt(x.rp,1)+'-'+trTxt(x.rk,1)+(x.org?' · orgánico':'')+'</em>':'')+'</div>'}).join('')}
function sgClose(){if(SG.box){SG.box.remove()}SG.box=null;SG.inp=null;SG.L=[];SG.hi=-1}
function sgHi(d){if(!SG.box)return;var n=SG.L.length;SG.hi=SG.hi<0?(d>0?0:n-1):(SG.hi+d+n)%n;$$('.sgo',SG.box).forEach(function(o){o.classList.toggle('on',+o.dataset.i===SG.hi)});var on=$('.sgo.on',SG.box);if(on)on.scrollIntoView({block:'nearest'})}
function sgPick(i){var inp=SG.inp,x=SG.L[i];if(!inp||!x)return;var kind=inp.dataset.sg,path=inp.dataset.p;sgClose();inp.value=x.n;
  var nxt=null;
  if(kind==='ab'){var q=filaDe(path),f=finca(UI.fid),p=f&&pabCur(f);if(q&&p){q.n=x.n;q.rn=x.rn;q.rp=x.rp;q.rk=x.rk;if(x.org)q.org=1;else delete q.org;if(abHas(x.pz))q.pz=x.pz;
      var mu=/^ud\.(\d+)\./.exec(path),u=mu&&p.ud[+mu[1]];if(!q.fo)q.fo=u&&u.rg==='regadio'?'Fertirrigación':'Suelo';if(!q.eq)q.eq=u&&u.rg==='regadio'?'Cabezal de riego':'Abonadora';p.m=Date.now();save()}
    nxt=$('[data-p="'+path.replace(/\.n$/,'.d')+'"]')}
  else{var q2=filaDe(path),f2=finca(UI.fid),p2=f2&&pabCur(f2);if(q2&&p2){q2.mo=x.n;p2.m=Date.now();save()}nxt=$('[data-p="'+path.replace(/\.mo$/,'.n')+'"]')}
  if(nxt)nxt.focus();else inp.dispatchEvent(new Event('change',{bubbles:true}));
  pabRedraw()}
document.addEventListener('mousedown',function(ev){var o=ev.target.closest&&ev.target.closest('.sgo');if(o){ev.preventDefault();sgPick(+o.dataset.i)}});
var SGT=null;document.addEventListener('touchstart',function(ev){var t=ev.touches&&ev.touches[0];SGT=t?{x:t.clientX,y:t.clientY}:null},{passive:true});
document.addEventListener('touchend',function(ev){var o=ev.target.closest&&ev.target.closest('.sgo');if(!o)return;var t=ev.changedTouches&&ev.changedTouches[0];if(SGT&&t&&(Math.abs(t.clientX-SGT.x)>8||Math.abs(t.clientY-SGT.y)>8))return;ev.preventDefault();sgPick(+o.dataset.i)},{passive:false});

/* ───── teclado: atajos y moverse por el programa como en Excel ───── */
document.addEventListener('keydown',function(ev){var t=ev.target,k=ev.key,mod=ev.ctrlKey||ev.metaKey,ed=t&&(/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName)||t.isContentEditable);
  if(SG.box&&t===SG.inp){if(k==='ArrowDown'){ev.preventDefault();sgHi(1);return}if(k==='ArrowUp'){ev.preventDefault();sgHi(-1);return}if(k==='Escape'){ev.preventDefault();sgClose();return}
    if((k==='Enter'||k==='Tab')&&SG.hi>=0){ev.preventDefault();sgPick(SG.hi);return}if(k==='Tab')sgClose()}
  else if(t&&t.classList&&t.classList.contains('sgi')&&k==='ArrowDown'){ev.preventDefault();sgOpen(t);sgHi(1);return}
  if(k==='Escape'){if(!$('#pop').hidden){popClose();return}if(!$('#sheet').hidden){sheetClose();return}if(!$('#mapa').hidden){mapaCerrar();return}}
  if(mod&&(k==='s'||k==='S'||k==='p'||k==='P')&&ed&&t.blur)t.blur();
  if(mod&&(k==='s'||k==='S')){ev.preventDefault();if(FS.ok)fsWrite(2);else if(S.fincas.length)copiaTodo();return}
  if(mod&&(k==='z'||k==='Z')&&!ev.shiftKey&&!ed){ev.preventDefault();undoDo();return}
  if(mod&&(k==='p'||k==='P')){var f=finca(UI.fid),p=f&&pabCur(f);if(p){ev.preventDefault();pdfGo(f,p)}return}
  if(ev.altKey&&(k==='ArrowLeft'||k==='ArrowRight')&&UI.v==='finca'){ev.preventDefault();var i=stepIdx()+(k==='ArrowRight'?1:-1);if(i>=0&&i<6)irPaso(STEPS[i][0]);return}
  if((k==='Enter'||k===' ')&&t.classList&&t.classList.contains('tap')&&t===document.activeElement&&!ed){ev.preventDefault();t.click();return}
  if((k==='Enter'||k===' ')&&t.classList&&t.classList.contains('pacrow')){ev.preventDefault();t.click();return}
  if(k==='Enter'&&t.tagName==='INPUT'&&t.dataset.p){var tr=t.closest('table.sheet tbody tr');if(tr){ev.preventDefault();var key=t.dataset.p.replace(/^.*\.pg\.rows\.\d+\./,''),rows=$$('tr[data-j]',tr.parentNode),j=rows.indexOf(tr),nx=rows[j+(ev.shiftKey?-1:1)];
      var dst=nx?$('[data-p$=".'+key+'"]',nx):$('[data-act="pa-row-add"][data-u="'+(t.dataset.p.split('.')[1])+'"]');if(dst){dst.focus();if(dst.select)dst.select()}else t.blur();return}
    if(t.classList.contains('pa-in')||t.classList.contains('ap-in')){ev.preventDefault();t.blur();t.focus();return}}
});
/* globo de los datos del calendario */
document.addEventListener('mousemove',function(ev){var h=ev.target.closest&&ev.target.closest('[data-tip]'),tp=$('#tip');if(!tp)return;if(!h){tp.hidden=true;return}tp.textContent=h.dataset.tip;tp.hidden=false;var x=Math.min(window.innerWidth-tp.offsetWidth-8,ev.clientX+14),y=ev.clientY+16;if(y+tp.offsetHeight>window.innerHeight)y=ev.clientY-tp.offsetHeight-10;tp.style.left=x+'px';tp.style.top=y+'px'});

/* ───── arrastrar y soltar archivos en la ventana ───── */
(function(){var n=0,hasF=function(e){return e.dataTransfer&&[].indexOf.call(e.dataTransfer.types||[],'Files')>=0};
  window.addEventListener('dragenter',function(e){if(!hasF(e))return;e.preventDefault();n++;$('#drop').hidden=false});
  window.addEventListener('dragover',function(e){if(!hasF(e))return;e.preventDefault();e.dataTransfer.dropEffect='copy'});
  window.addEventListener('dragleave',function(e){if(!hasF(e))return;n=Math.max(0,n-1);if(!n)$('#drop').hidden=true});
  window.addEventListener('drop',function(e){if(!hasF(e))return;e.preventDefault();n=0;$('#drop').hidden=true;var fl=e.dataTransfer.files&&e.dataTransfer.files[0];if(!fl)return;soltar(fl)})})();
function soltar(fl){var nm=String(fl.name||'').toLowerCase(),ty=fl.type||'';
  if(/\.pdf$/.test(nm)||ty==='application/pdf'){pacLeer(fl);return}
  if(/\.json$/.test(nm)||ty==='application/json'){copiaAbrir(fl);return}
  if(/\.(csv|txt)$/.test(nm)||/text\//.test(ty)){csvLeer(fl);return}
  toast('Ese archivo no lo sé leer: suelta el PDF de la PAC, una tabla CSV o una copia (.json).')}

/* ───── arranque ───── */
cargar();
var MQ=[window.matchMedia('(min-width:900px)'),window.matchMedia('(min-width:1200px)')];
MQ.forEach(function(q){var fn=function(){if($('#app'))render()};if(q.addEventListener)q.addEventListener('change',fn);else if(q.addListener)q.addListener(fn)});
document.addEventListener('DOMContentLoaded',function(){render();fsInit()});
(function(){var tm=null;try{new MutationObserver(function(){if(tm)return;tm=setTimeout(function(){tm=null;stkLab();inThou()},0)}).observe(document.documentElement,{childList:true,subtree:true})}catch(e){}})();
if('serviceWorker' in navigator&&(location.protocol==='https:'||location.hostname==='localhost'))window.addEventListener('load',function(){navigator.serviceWorker.register('sw.js').catch(function(){})});
window.__pa={S:function(){return S},UI:function(){return UI},render:render,ejemplo:ejemplo,pabCalc:pabCalc,pabBlocks:pabBlocks,buildPabXlsx:buildPabXlsx,ensurePdf:ensurePdf,buildPabPdf:buildPabPdf,climaDe:climaDe,climaKey:climaKey,refParse:refParse,catGen:catGen,mesesDe:mesesDe,propuesta:propuesta,faltas:faltas,ctxOf:ctxOf,mapaAbrir:mapaAbrir,MAP:function(){return MAP}};
