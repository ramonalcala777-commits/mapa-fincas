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
/* Plan de abonado — herramienta independiente (versión 3: tres pasos, Finca · Datos · Plan y PDF).
   Todo se guarda en este aparato (localStorage) y, si se quiere, en un archivo de copia. No hay servidor ni cuentas.
   El cálculo, el PDF y el Excel del plan vienen de base.js (RD 1051/2022 y Decreto 80/2025 de Castilla-La Mancha). */
var VERSION="3.2 · 19 nov 2026";
var KEY='plan-abonado:datos', KEYUI='plan-abonado:vista', KEYGEO='plan-abonado:contornos';
var $=function(s,r){return (r||document).querySelector(s)}, $$=function(s,r){return [].slice.call((r||document).querySelectorAll(s))};

/* ───── datos ───── */
var S={fincas:[],recintos:{},x:{asesor:{},abonos:[]},copia:0,mod:0};
var UI={v:'inicio',fid:null,step:'fin'};
var GTOK=false, EXP={fid:null};
function uiDef(){return {v:'inicio',fid:null,step:'fin',q:'',sort:'ref',cost:0,det:null,mLast:null}}
/* los pasos de la versión 2 (rec, ud, su, ne, pg, rv) pasan a los tres de la versión 3 */
function stepOf(s){return {rec:'fin',ud:'dat',su:'dat',ne:'dat',pg:'plan',rv:'plan'}[s]||(/^(fin|dat|plan)$/.test(s)?s:'fin')}
function cargar(){
  try{var t=localStorage.getItem(KEY);if(t){var o=JSON.parse(t);if(o&&Array.isArray(o.fincas)){S=o;S.recintos=S.recintos||{};S.x=S.x||{};S.x.asesor=S.x.asesor||{};S.x.abonos=S.x.abonos||[]}}}catch(e){}
  try{var u=JSON.parse(localStorage.getItem(KEYUI)||'null');if(u&&u.v){UI=Object.assign(uiDef(),u)}}catch(e){}
  if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null}
  UI.step=stepOf(UI.step); UI.det=null;
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
function undoDo(){var t=UNDO.pop();if(!t){toast('No hay nada que deshacer');return}var o=JSON.parse(t);S.fincas=o.fincas;S.recintos=o.recintos;S.x=o.x;save();if(typeof RMODE==='object')RMODE={};var mo=!$('#mapa').hidden;if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null;if(mo){mapaCerrar();mo=false}}render();if(mo){mRestyle();mSide()}toast('Deshecho')}
function undoToast(msg){toast(msg,'Deshacer',undoDo)}

/* avisos: cortos, arriba a la derecha (punto 36); no tapan casillas */
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
/* PDF y Excel con un clic (punto 34): en el ordenador se descarga directamente; en el móvil, la hoja de «Guardar / Enviar» */
function dlDirecto(blob,name,what,pdf){DL={blob:blob,name:name};if(!isDesk()&&canShareFile(blob,name)){showDl(blob,name,what+' listo');return}dlSave();toast(what+' descargado: '+name,pdf?'Abrir':null,pdf?dlView:null)}
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
/* v3.2: «Abono tipo» (el listado de arriba, con su riqueza típica) o «Mi abono» (nombre comercial: «Mis abonos» y «Usados antes») */
function abKey(n){return noAcc(String(n||'')).toLowerCase().trim()}
function catEsTipo(n){var k=abKey(n);return !!k&&catGen().some(function(x){return abKey(x.n)===k})}
function catGrupo(x){return x.g==='Mis abonos'||x.g==='Usados antes'?x.g:'Abonos tipo: '+x.g}
function catBusca(q,de){var t=noAcc(q||'').toLowerCase().trim(), L=de==='tipo'?catGen():de==='mio'?catMios():catTodo();if(!t)return L;var ws=t.split(/\s+/);
  return L.filter(function(x){var n=noAcc(x.n).toLowerCase()+' '+[x.rn,x.rp,x.rk].join('-')+' '+noAcc(catGrupo(x)).toLowerCase();return ws.every(function(w){return n.indexOf(w)>=0})})}
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
  ['hs','pend','uso','coef','ll','alt','adm','inc','hpac'].forEach(function(k){if(o[k]!==undefined&&o[k]!==null)r[k]=o[k]});return r}
function addRecs(f,list){var p0=pabCur(f),man=p0&&udMan(f,p0),added=0,skip=[];list.forEach(function(o){var ex=S.recintos[o.ref];if(ex&&ex.fincaId!==f.id&&finca(ex.fincaId)){skip.push(o.ref);return}if(ex&&ex.fincaId===f.id){skip.push(o.ref);return}o.fincaId=f.id;S.recintos[o.ref]=recNew(o.ref,o);added++});if(p0&&added)udTrasCambio(f,p0,man);return {added:added,skip:skip}}
/* geometría WKT del SIGPAC (lon lat, ETRS89) → punto central [lat, lon] */
function wktLL(w){var m=String(w||'').match(/-?\d+\.?\d*\s+-?\d+\.?\d*/g);if(!m||!m.length)return null;var x=0,y=0;m.forEach(function(p){var a=p.trim().split(/\s+/);x+=+a[0];y+=+a[1]});x/=m.length;y/=m.length;
  if(!(Math.abs(y)<=90&&Math.abs(x)<=180))return null;return [Math.round(y*1e5)/1e5,Math.round(x*1e5)/1e5]}
/* SIGPAC: superficie, pendiente media, uso y situación del recinto (consulta pública, necesita internet) */
var SIGB={};
function sigInfo(ref){return fetch('https://sigpac-hubcloud.es/servicioconsultassigpac/query/recinfo/'+ref.split(':').join('/')+'.json').then(function(r){if(!r.ok)throw 1;return r.json()}).then(function(j){var o=Array.isArray(j)?j[0]:j;if(!o)throw 1;if(o.wkt&&typeof geoPon==='function')geoPon(ref,o.wkt);return o})}
function sigAplica(r,o){var nn=function(v){return v==null||v===''?null:+v};
  /* v3.1 punto 5: la PAC trae las ha con 2 decimales y el SIGPAC con 4. Si la superficie es la de la PAC (no escrita a mano) y coincide
     con la del SIGPAC redondeada (menos de 0,006 ha de diferencia), se queda la del SIGPAC; si difiere más, no se pisa (el recinto pudo cambiar).
     Las ha declaradas de la PAC que solo se distinguen del SIGPAC por el redondeo se quitan. */
  if(o.superficie!=null){var sg=+o.superficie,dePac=r.hpac!=null&&r.ha>0&&Math.abs(r.ha-r.hpac)<1e-9;
    if(!(r.ha>0))r.ha=sg;
    else if(dePac&&Math.abs(sg-r.ha)<0.006){r.ha=sg;if(r.hc!=null&&Math.abs(r.hc-sg)<0.006)delete r.hc;delete r.hpac}
    r.hs=sg} if(o.pendiente_media!=null)r.pend=+o.pendiente_media; if(o.uso_sigpac)r.uso=o.uso_sigpac; if(o.coef_regadio!=null)r.coef=+o.coef_regadio;
  if(nn(o.altitud))r.alt=+o.altitud; if(o.admisibilidad!=null)r.adm=+o.admisibilidad; if(o.incidencias!=null)r.inc=String(o.incidencias); var ll=wktLL(o.wkt); if(ll)r.ll=ll}
function sigLeer(f,refs,cb){
  refs=refs.filter(function(x){return S.recintos[x]&&!SIGB[x]}); if(!refs.length){if(cb)cb(0,0);return}
  if(navigator.onLine===false){toast('Sin internet no puedo leer el SIGPAC. Prueba luego.');if(cb)cb(0,refs.length);return}
  var idx=0,ok=0,ko=0; refs.forEach(function(x){SIGB[x]=1});
  var one=function(){if(idx>=refs.length)return Promise.resolve();var x=refs[idx++];
    /* v3.2: la zona vulnerable solo se pregunta si el SIGPAC ha dado los datos del recinto (un recinto que no existe también sale «fuera») */
    return sigInfo(x).then(function(o){var r=S.recintos[x];if(!r)return;sigAplica(r,o);ok++;return zvLee(x)}).catch(function(){ko++}).then(function(){delete SIGB[x];return one()})};
  Promise.all([one(),one(),one(),one()]).then(function(){pendAPlanes(f);climaAuto(f);save();if(cb)cb(ok,ko);render()});
}
/* ───── v3.2: zona vulnerable a nitratos según el SIGPAC (FEGA) ─────
   Consulta «Intersección de recinto con nitratos»: % del recinto dentro de la capa de nitratos, o [] si no toca.
   Se guarda en el recinto r.zvn={p:%,m2,f:fecha}. Si falla, no se guarda nada (nunca «No» por un fallo). */
var ZVQ={}, ZVKO={};
function sigNitr(ref){return fetch('https://sigpac-hubcloud.es/servicioconsultassigpac/intersection/nitratos/'+ref.split(':').join('/')+'.json').then(function(r){if(!r.ok)throw 1;return r.json()}).then(function(j){if(!Array.isArray(j))throw 1;var p=0,m2=0;j.forEach(function(o){p+=+o.surface_tpc||0;m2+=+o.surface_intersection||0});return {p:Math.min(100,Math.round(p*1000)/1000),m2:Math.round(m2),f:hoyIso()}})}
function zvLee(ref){ZVQ[ref]=1;return sigNitr(ref).then(function(o){var r=S.recintos[ref];if(r)r.zvn=o;delete ZVKO[ref];return o}).catch(function(){ZVKO[ref]=1;return null}).then(function(o){delete ZVQ[ref];return o})}
/* al abrir una finca: se pregunta lo que falte, sin avisos. Solo los recintos que el SIGPAC ya ha confirmado (tienen sus datos);
   los demás, si el SIGPAC los encuentra ahora. No cambia nada más del recinto. */
/* la finca de ejemplo tiene códigos inventados (municipio 900): no se pregunta al SIGPAC */
function fEj(f){return !!(f&&(f.ej===1||(f.datos&&f.datos.tit&&f.datos.tit.nif==='00000000T'&&f.nombre==='Finca de ejemplo')))}
function zvComprobar(f){if(!f||fEj(f)||navigator.onLine===false)return;var rs=fincaRecs(f.id).filter(function(r){return !r.zvn&&!ZVQ[r.ref]});if(!rs.length)return;
  var L=rs.slice(),n=0,one=function(){var r=L.shift();if(!r)return Promise.resolve();var go=r.hs!=null?zvLee(r.ref):sigInfo(r.ref).then(function(){return zvLee(r.ref)},function(){ZVKO[r.ref]=1;return null});
    return go.then(function(o){if(o)n++;return one()})};
  rs.forEach(function(r){ZVQ[r.ref]=1});
  Promise.all([one(),one(),one(),one()]).then(function(){if(n){saveRaw();if(UI.v==='finca'&&UI.fid===f.id)render()}else if(UI.v==='finca'&&UI.fid===f.id)renderLuego()})}
/* «Sí, 100 %», «Sí, 2 % (borde)», «No» */
function zvPct(p){return p>=99.5?'100 %':(p<0.1?'menos de 0,1 %':(p<10?trTxt(p,1).replace(/,0$/,''):trTxt(p,0))+' %')}
function zvTxt(o){if(!o)return '';return o.p>0?'Sí, '+zvPct(o.p)+(o.p<99.5?(o.p<10?' (borde)':' (en parte)'):''):'No'}
/* la unidad: lo que dice el SIGPAC de sus recintos */
function zvMan(u){return u.zvMan===1||(u.zv===1&&u.zvA!==1)}
function zvSig(f,u){var R=(u.refs||[]).map(function(x){return S.recintos[x]}).filter(function(r){return r&&r.fincaId===f.id}),con=R.filter(function(r){return r.zvn}),den=con.filter(function(r){return r.zvn.p>0}),fe='';
  con.forEach(function(r){if(r.zvn.f>fe)fe=r.zvn.f});
  return {R:R,con:con,den:den,fue:con.filter(function(r){return !(r.zvn.p>0)}),sin:R.filter(function(r){return !r.zvn}),pend:R.some(function(r){return ZVQ[r.ref]}),v:den.length?1:(R.length&&con.length===R.length?0:null),f:fe}}
/* pone la unidad en «Sí» si algún recinto cae dentro y en «No» si ninguno cae y se han podido mirar todos; lo puesto a mano no se toca.
   Las unidades de antes de la 3.2 marcadas en «Sí» cuentan como puestas a mano (no se bajan a «No» solas). Solo escribe si cambia algo. */
function zvAplicar(f){var ch=0;(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){if(zvMan(u))return;var z=zvSig(f,u);if(z.v===null)return;
  if(u.zv!==z.v||u.zvA!==1||u.zvF!==z.f){u.zv=z.v;u.zvA=1;u.zvF=z.f;ch=1}})});return ch}
/* «Separar en dos unidades»: los recintos que caen en zona vulnerable (aunque sea por el borde) y los que no */
function zvSeparar(f,p,i){var u=p.ud[i];if(!u)return;var z=zvSig(f,u);if(!z.den.length||!z.fue.length)return;undoPush();
  var fuera=z.fue.map(function(r){return r.ref}),base=String(u.n||'').replace(/ \((fuera de )?zona vulnerable\)$/,'');
  var nu=JSON.parse(JSON.stringify(u));nu.id=uid();nu.refs=fuera;nu.n=base+' (fuera de zona vulnerable)';nu.zv=0;nu.zvA=1;nu.zvF=z.f;delete nu.zvMan;delete nu.suDe;
  if(nu.rc)Object.keys(nu.rc).forEach(function(x){if(fuera.indexOf(x)<0)delete nu.rc[x]});
  ((nu.pg||{}).rows||[]).forEach(function(q){q.id=uid()});
  u.refs=(u.refs||[]).filter(function(x){return fuera.indexOf(x)<0});u.n=base+' (zona vulnerable)';u.zv=1;u.zvA=1;u.zvF=z.f;delete u.zvMan;
  if(u.rc)fuera.forEach(function(x){delete u.rc[x]});
  var big=function(L){return L.map(function(x){return S.recintos[x]}).filter(Boolean).sort(function(a,b){return haC(b)-haC(a)})[0]};
  if(u.refs.indexOf(u.rp)<0){var b1=big(u.refs);u.rp=b1?b1.ref:''}if(nu.refs.indexOf(nu.rp)<0){var b2=big(nu.refs);nu.rp=b2?b2.ref:''}
  p.ud.splice(i+1,0,nu);p.udMan=1;p.m=Date.now();save();render();undoToast('«'+base+'» separada en dos unidades: dentro y fuera de zona vulnerable')}
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
      var nl=Object.keys(nifs), tit={};try{tit=pacTitular(pages)}catch(e){tit={}}
      /* v3.1: si no se encuentra el bloque del solicitante, como antes: el NIF solo si hay uno */
      if(!tit.nif&&nl.length===1)tit.nif=nl[0];
      PACI={kind:'pac',rows:o.rows.map(function(r){var cu=r.cultivo!==undefined?r.cultivo:(pacCrop(r.prod)||pacCropVar(r.vari)||titleCase2(r.prod));
        return {ref:[r.prov,r.muni,r.agr,r.zona,r.pol,r.par,r.rec].join(':'),mname:r.mname,ha:r.sigHa,hc:Math.abs(r.sigHa-r.decHa)>0.00005&&r.decHa>0?r.decHa:undefined,cultivo:cu,riego:r.riego||'',variedad:titleCase2(r.vari),uso:r.uso||undefined,hpac:r.sigHa,mix:r.mix}}),year:o.year,nif:tit.nif||'',tit:tit,name:''};
      var mn=PACI.rows.map(function(r){return r.mname}).filter(Boolean);PACI.name=mn.length?'Finca de '+titleCase2(mn[0]):'Mi finca';
      pacHoja();
    }).catch(function(e){PACI=null;toast(e&&e.message==='net'?'Para leer el PDF hace falta internet la primera vez.':'No he encontrado la tabla de recintos en ese PDF.')});
  };
  rd.readAsArrayBuffer(file);
}
/* v3.1 punto 4: titular, NIF, domicilio y representante desde el bloque del solicitante de la PAC.
   Se lee por posición: una fila de rótulos («CIF/NIF», «Nombre / Razón Social»…) y debajo sus valores, dentro del apartado
   («DATOS PERSONALES DEL SOLICITANTE», «DIRECCIÓN RESIDENCIA», «REPRESENTANTE»). Lo que no se encuentre así se deja vacío:
   mejor vacío que con el NIF del cónyuge, del gerente o de un socio. Si la PAC tiene otra maqueta, no se lee nada y no falla. */
var PAC_LBL=/^(cif\/nif|nif\/cif|nif|cif|dni\/nif|nombre( \/ razon social| o razon social)?|razon social|apellidos|primer apellido|segundo apellido|f\. nacimiento|fecha de nacimiento|genero|sexo|estado civil|telefono|movil|e-?mail|correo electronico|tipo (de )?via|nombre (de la )?via|domicilio|numero|n\.?º|bloque|portal|escalera|planta|puerta|piso|km|provincia|municipio|localidad|codigo postal|c\.? ?p\.?|apartado de correos|oficina de correos)$/;
function pacNorm(s){return noAcc(s).toLowerCase().replace(/\s+/g,' ').trim()}
function pacTitular(pages){
  var L=[];pages.slice(0,4).forEach(function(pg,pi){pacLineas(pg).forEach(function(g){g.pg=pi;L.push(g)})});
  var lbl=function(w){return PAC_LBL.test(pacNorm(w.s))}, hdr=function(g){return g.w.length&&g.w.every(lbl)},
    sec=function(g){return g.w.length===1&&!lbl(g.w[0])&&/^[A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ .\/,()-]{4,}$/.test(g.w[0].s)};
  var find=function(re){for(var i=0;i<L.length;i++)if(sec(L[i])&&re.test(pacNorm(L[i].w[0].s)))return i;return -1};
  /* rótulo → valor, en las filas del apartado que empieza en la línea i */
  var vals=function(i,desde){var V={};if(i<0)return V;var j0=desde?i:i+1;for(var j=j0;j<L.length&&j<j0+24&&!sec(L[j])&&L[j].pg===L[i].pg;j++){var h=L[j],d=L[j+1];if(!hdr(h)||!d||d.pg!==h.pg||hdr(d)||sec(d)||d.y-h.y>16)continue;
      d.w.forEach(function(w){var c=null;h.w.forEach(function(x){if(x.x<=w.x+6)c=x});if(!c)return;var k=pacNorm(c.s);V[k]=V[k]?V[k]+' '+w.s:w.s});j++}return V};
  var nifOk=function(v){v=String(v||'').toUpperCase().replace(/[\s.\-]/g,'');return /^(\d{8}[A-Z]|[XYZ]\d{7}[A-Z]|[A-HJNP-SUVW]\d{7}[0-9A-J])$/.test(v)&&!nifMal(v)?v:''};
  var nifDe=function(V){return nifOk(V['cif/nif']||V['nif/cif']||V['nif']||V['cif']||V['dni/nif'])};
  var nomDe=function(V){var n=V['nombre / razon social']||V['nombre o razon social']||V['razon social']||V['nombre']||'';var ap=V['apellidos']||[V['primer apellido'],V['segundo apellido']].filter(Boolean).join(' ');return [n,ap].filter(Boolean).join(' ').replace(/\s+/g,' ').trim()};
  var T={}, V=vals(find(/^(datos (personales )?(del|de la) (persona )?(solicitante|titular|beneficiari[oa])|solicitante|titular|beneficiari[oa])$/));
  if(!nifDe(V)&&!nomDe(V)){/* sin el apartado: la fila «CIF/NIF · Nombre / Razón Social» de la cabecera de las hojas */
    for(var i=0;i<L.length-1;i++){var h=L[i];if(!hdr(h))continue;var ks=h.w.map(function(w){return pacNorm(w.s)});if(ks.indexOf('cif/nif')<0||ks.indexOf('nombre / razon social')<0)continue;
      var W=vals(i,1);if(nifDe(W)){V=W;break}}}
  T.nif=nifDe(V);T.n=T.nif?nomDe(V):'';if(!T.n)delete T.n;if(!T.nif)delete T.nif;
  if(T.nif){
    var iD=find(/^(direccion|domicilio)( de)? (residencia|fiscal|social)$/);if(iD<0)iD=find(/^(direccion|domicilio)( de)? notificacion$/);if(iD<0)iD=find(/^(direccion|domicilio)$/);
    var D=vals(iD),via=D['nombre via']||D['nombre de la via']||D['domicilio']||'';
    if(via){var tv=D['tipo via']||D['tipo de via']||'',nu=D['numero']||D['n.º']||D['nº']||'';T.dir=[titleCase2(tv),titleVia(via)].filter(Boolean).join(' ')+(nu?', '+nu:'')}
    var cp=String(D['codigo postal']||D['c.p.']||D['cp']||'').trim();if(/^\d{5}$/.test(cp))T.cp=cp;
    if(D['municipio']||D['localidad'])T.mun=D['municipio']||D['localidad'];if(D['provincia'])T.prov=D['provincia'];
    var R=vals(find(/^(datos del )?representante( legal)?$/)),rnif=nifDe(R),rn=nomDe(R);
    if(rnif&&rn&&rnif!==T.nif){T.rn=rn;T.rnif=rnif}}
  return T}
function titleVia(s){return titleCase2(s).replace(/ (De|Del|La|Las|Los|El|Y|En)(?= )/g,function(m){return m.toLowerCase()})}
/* lo leído del titular, en una línea para la hoja de la PAC */
function pacTitTxt(t){t=t||{};var L=[];if(t.n||t.nif)L.push('Titular: '+(t.n||'sin nombre')+(t.nif?' · NIF '+t.nif:''));
  var d=[t.dir,[t.cp,t.mun].filter(Boolean).join(' '),t.prov].filter(Boolean).join(', ');if(d)L.push('Domicilio: '+d);if(t.rn)L.push('Representante: '+t.rn+(t.rnif?' · NIF '+t.rnif:''));return L}
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
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, rows=PACI.rows, on=rows.filter(function(r){return !r.off}), n=on.length, tot=on.reduce(function(a,r){return a+(r.hc||r.ha||0)},0), totS=on.reduce(function(a,r){return a+(r.ha||0)},0);
  var G=pacGrupos(rows), y0=pabY0(), conPlan=!f||!(f.pab||[]).length;
  var h='<h2>Esto es lo que he entendido</h2>';
  h+='<div class="und">'+
    (PACI.kind==='pac'?'<div><span class="k">Solicitud</span><b>PAC '+esc(PACI.year||'')+'</b></div>':'<div><span class="k">Origen</span><b>Tabla CSV</b></div>')+
    (PACI.nif&&!(PACI.tit&&PACI.tit.n)?'<div><span class="k">NIF del titular</span><b>'+esc(PACI.nif)+'</b></div>':'')+
    '<div><span class="k">Recintos</span><b>'+n+' de '+rows.length+'</b></div><div><span class="k">'+(Math.abs(tot-totS)>0.005?'Superficie declarada':'Superficie')+'</span><b>'+ha2(tot)+' ha</b>'+(Math.abs(tot-totS)>0.005?'<span class="small">SIGPAC: '+ha2(totS)+' ha</span>':'')+'</div></div>';
  /* v3.1 punto 4: lo leído del titular, para verlo antes de aceptar */
  var TL=pacTitTxt(PACI.tit),ft=f&&((f.datos||{}).tit||{}),tOtro=ft&&PACI.tit&&PACI.tit.nif&&ft.nif&&ft.nif.toUpperCase()!==PACI.tit.nif;
  if(TL.length)h+='<div class="pactit">'+TL.map(function(t){return '<div>'+esc(t)+'</div>'}).join('')+(tOtro?'<div class="small">«'+esc(f.nombre)+'» ya tiene otro titular: no lo cambio.</div>':(ft&&(ft.n||ft.nif)?'<div class="small">Lo que ya está escrito del titular en «'+esc(f.nombre)+'» no se cambia.</div>':''))+'</div>';
  else if(PACI.kind==='pac')h+='<p class="small">No he encontrado los datos del titular en la PAC: escríbelos luego en «Finca».</p>';
  /* v3.1 punto 3: recintos declarados con dos cultivos (solo en pantalla) */
  var MX=rows.filter(function(r){return r.mix&&!r.off});
  if(MX.length)h+='<div class="warn pacdos">'+ic('alert')+'<div><b>'+(MX.length===1?'1 recinto declarado':MX.length+' recintos declarados')+' con más de un cultivo.</b> Aquí cada recinto lleva un cultivo: pongo el de más hectáreas.'+MX.map(function(r){return '<div class="pmx"><b>'+esc(refLabel(r.ref))+'</b>: '+esc(pacMixTxt(r))+'</div>'}).join('')+'</div></div>';
  h+='<div class="lbl">Unidades de producción que salen (cultivo y riego'+(Math.abs(tot-totS)>0.005?', con las ha declaradas':'')+')</div><div class="ugl">'+G.map(function(g){return '<div><b>'+esc(g.n)+'</b><span>'+g.k+(g.k===1?' recinto':' recintos')+' · '+ha2(g.ha)+' ha'+(g.rg==='secano'&&g.ha<=10?' · <i>exenta (secano ≤ 10 ha)</i>':'')+'</span></div>'}).join('')+'</div>';
  if(rows.some(function(r){return !r.riego}))h+='<p class="small">Los recintos sin «R» ni «S» en la PAC se toman como secano: cámbialo en la tabla de recintos si hace falta.</p>';
  h+='<details class="more"'+(isDesk()?' open':'')+'><summary>Ver los '+rows.length+' recintos (toca uno para quitarlo o ponerlo)</summary><div class="pacl">'+rows.map(function(r,i){var ex=S.recintos[r.ref];return '<div class="pacrow'+(r.off?' off':'')+'" data-act="ap-pac-tog" data-i="'+i+'" role="checkbox" tabindex="0" aria-checked="'+(!r.off)+'"><i>'+(r.off?'':'✓')+'</i><div><b>'+esc(refLabel(r.ref))+'</b>'+(r.mname?' · '+esc(titleCase2(r.mname)):'')+'<small>'+esc([r.cultivo||'Sin cultivo',r.variedad,r.riego==='regadio'?'Regadío':(r.riego==='secano'?'Secano':''),ha2(r.ha)+' ha'+(r.hc?' (declaradas '+ha2(r.hc)+')':'')].filter(Boolean).join(' · '))+(ex?' · ya está en «'+esc((finca(ex.fincaId)||{}).nombre||'')+'»':'')+(r.mix?' · <b>'+(r.mix.length===2?'dos':r.mix.length)+' cultivos</b>':'')+'</small></div></div>'}).join('')+'</div></details>';
  if(!f)h+='<label class="lbl" for="pac-name">Nombre de la finca</label><input id="pac-name" class="in" value="'+esc(PACI.name)+'" autocomplete="off">';
  if(conPlan)h+='<label class="ckr pacplan"><input type="checkbox" id="pac-plan" checked> Crear también el plan de abonado de la campaña <select id="pac-y" class="trsel sm">'+[y0-1,y0,y0+1].map(function(v){return '<option value="'+v+'"'+(v===y0?' selected':'')+'>'+pabCamp(v)+'</option>'}).join('')+'</select></label>';
  h+='<div class="acts"><button class="btn primary xl" data-act="ap-pac-go"'+(n?'':' disabled')+'>'+(f?'Añadir '+n+(n===1?' recinto':' recintos'):'Crear la finca')+(conPlan?' y el plan':'')+'</button><button class="btn" data-act="ap-close">Cancelar</button></div>';
  sheet(h,'wide');
}
function pacGo(){
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, ni=$('#pac-name'), pc=$('#pac-plan'), py=$('#pac-y'), mk=pc&&pc.checked, y=py?+py.value:pabY0();
  undoPush();
  if(!f){f=fincaNueva((ni&&ni.value.trim())||PACI.name||'Mi finca')}
  /* v3.1 punto 4: los datos del titular leídos se ponen solo donde la finca no tiene nada; si la finca ya es de otro NIF, no se toca */
  f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};var T0=f.datos.tit,TP=PACI.tit||(PACI.nif?{nif:PACI.nif}:{});
  if(!(TP.nif&&T0.nif&&T0.nif.toUpperCase()!==TP.nif))['n','nif','dir','cp','mun','prov','rn','rnif'].forEach(function(k){if(TP[k]&&!T0[k])T0[k]=TP[k]});
  var R=addRecs(f,PACI.rows.filter(function(r){return !r.off}).map(function(r){return Object.assign({},r)}));
  var refs=fincaRecs(f.id).map(function(r){return r.ref});
  var p=null; if(mk&&fincaRecs(f.id).length&&!(f.pab||[]).some(function(x){return +x.y===y}))p=planNuevo(f,y);
  sheetClose(); UI.v='finca';UI.fid=f.id;planAuto(f);UI.step=ctxOf(f).F.some(function(x){return x.s==='fin'})?'fin':'dat'; save(); saveUI(); window.__scrollTop=true; render();
  toast(R.added+(R.added===1?' recinto añadido':' recintos añadidos')+(R.skip.length?' · '+R.skip.length+' ya estaban':'')+(p?' y plan '+pabCamp(p.y)+' creado':'')+'. Leo el SIGPAC…');
  sigLeer(f,refs,function(ok,ko){if(ok||ko)toast('SIGPAC: '+ok+' leídos'+(ko?', '+ko+' sin respuesta':''))});
}
/* «Declarado con dos cultivos: Trigo 12,50 ha y Barbecho 3,50 ha. Pongo Trigo con 12,50 ha.» */
function pacMixTxt(r){var M=r.mix||[],nn=['','un','dos','tres','cuatro','cinco'][M.length]||String(M.length),li=M.map(function(m){return m.c+' '+ha2(m.ha)+' ha'});
  return 'Declarado con '+nn+' cultivos: '+(li.length>1?li.slice(0,-1).join(', ')+' y '+li[li.length-1]:li[0])+'. Pongo '+(r.cultivo||'Sin cultivo')+' con '+ha2(r.hc||r.ha)+' ha.'}
function fincaNueva(nm){var f={id:uid(),nombre:nm||'Mi finca',datos:{tit:{}},c:Date.now()};S.fincas.push(f);return f}

/* ───── planes: nuevo, y el de la campaña siguiente copiando lo que vale ───── */
var LEN=/^(pistacho|almendro|olivo|nogal|vid|melocotonero|frutal)$/;
function udPrefill(u){
  if(LEN.test(u.ck)&&!u.prec)u.prec=pabCkName(u.ck);
  if(!u.mit||!Object.keys(u.mit).length)mitPropuesta(u,true);
  if(!u.res||!Object.keys(u.res).length){u.res={};PAB_RES.forEach(function(x){u.res[x[0]]={a:1}});u.resProp=1}
  if(!u.mo||!u.mo.obj)u.mo=Object.assign({},u.mo,{obj:'mantener'});
}
function planNuevo(f,y,prev){
  var n={id:uid(),y:String(y),cl:climaDe(f),clAuto:1,ud:pabAuto(f),c:Date.now(),m:Date.now(),fe:hoyIso()};
  prev=prev||pabList(f)[0];
  if(prev){n.cl=Object.assign({},prev.cl||PAB_CL);delete n.clAuto;if(prev.clAuto)n.clAuto=1;
    n.ud.forEach(function(u){var o=(prev.ud||[]).filter(function(x){return x.n===u.n||(x.ck===u.ck&&x.rg===u.rg)})[0];if(!o)return;
      ['zv','zvMan','zvA','zvF','suNo','tp','inh','rp','loc','leg','bobs'].forEach(function(k){if(o[k]!==undefined)u[k]=o[k]});
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

/* ───── unidades que se hacen solas (punto 22): cultivo + riego ─────
   Se rehacen al cambiar el cultivo o el riego de un recinto, o al añadir o quitar recintos, solo en el plan de la campaña que se ve.
   Las unidades que siguen existiendo conservan sus datos y su programa. Si el usuario las había cambiado a mano (nombre, reparto,
   unidades nuevas), no se tocan: se avisa y se ofrece «Rehacer las unidades». */
function udClave(r){return (r.cultivo||NOCROP)+'§'+pabRg(r)}
function udNombre(k){var a=k.split('§');return a[0]+' '+(a[1]==='regadio'?'regadío':'secano')}
function udMan(f,p){if(!p)return false;if(p.udMan===1)return true;var seen={},man=false;
  (p.ud||[]).forEach(function(u){if((u.xr||[]).length)man=true;var ks={};(u.refs||[]).forEach(function(x){var r=S.recintos[x];if(r&&r.fincaId===f.id)ks[udClave(r)]=1});var kl=Object.keys(ks);
    if(kl.length>1)man=true;else if(kl.length===1){if(seen[kl[0]])man=true;seen[kl[0]]=1;if(u.n!==udNombre(kl[0]))man=true}else if(!/ (regadío|secano)$/.test(String(u.n||'')))man=true});
  return man}
function udNueva(f,k,refs){var a=k.split('§'),u={id:uid(),n:udNombre(k),ck:pabCk(a[0]===NOCROP?'':a[0]),rg:a[1],refs:refs.slice(),zv:0,tp:0,inh:0,rc:{},su:{},ag:{},ne:{mo:'cos',pr:'',ex:null,ed:1},ap:{s:{},o:{},x:{}},pg:{src:'own',rows:[]}};
  var big=refs.map(function(x){return S.recintos[x]}).filter(Boolean).sort(function(a,b){return haC(b)-haC(a)})[0];if(big)u.rp=big.ref;return u}
function udRehacer(f,p){if(!p)return;var G={},orden=[];
  fincaRecs(f.id).forEach(function(r){var k=udClave(r);if(!G[k]){G[k]=[];orden.push(k)}G[k].push(r.ref)});
  var used={},nu=[],ref0=(p.ud||[]).filter(function(u){return u.res&&Object.keys(u.res).length})[0];
  (p.ud||[]).forEach(function(u){var k=orden.filter(function(k){return udNombre(k)===u.n&&!used[k]})[0];if(!k)return;used[k]=1;
    u.refs=G[k].slice();if(u.rc)Object.keys(u.rc).forEach(function(x){if(u.refs.indexOf(x)<0)delete u.rc[x]});
    if(u.refs.indexOf(u.rp)<0){var big=u.refs.map(function(x){return S.recintos[x]}).sort(function(a,b){return haC(b)-haC(a)})[0];u.rp=big?big.ref:''}
    var a=k.split('§');u.ck=pabCk(a[0]===NOCROP?'':a[0]);u.rg=a[1];delete u.xr;nu.push(u)});
  orden.forEach(function(k){if(used[k])return;var u=udNueva(f,k,G[k]);udPrefill(u);
    /* medidas y restricciones de toda la finca (punto 24): las mismas que ya tienen las demás unidades */
    if(ref0){u.res=JSON.parse(JSON.stringify(ref0.res));if(ref0.resProp)u.resProp=1;else delete u.resProp;if(ref0.mo&&ref0.mo.obj)u.mo.obj=ref0.mo.obj}
    nu.push(u)});
  p.ud=nu;delete p.udMan;p.m=Date.now();pendAPlanes(f)}
/* después de cambiar los recintos: se rehacen si no estaban hechas a mano; si lo estaban, se avisa */
var UDAV={};
function udTrasCambio(f,p,man){if(!p)return;if(man){UDAV[p.id]=1;return}udRehacer(f,p)}

/* ───── ejemplo inventado (para probar) ───── */
function ejemplo(){
  undoPush(); var f=fincaNueva('Finca de ejemplo'); f.ej=1; f.datos.tit={n:'Titular de ejemplo',nif:'00000000T',mun:'Municipio',prov:'Albacete'};
  addRecs(f,[{ref:'2:900:0:0:1:1:1',ha:5.2,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman',nombre:'La Hoya',pend:2.1},{ref:'2:900:0:0:1:2:1',ha:3.1,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman',pend:1.4},{ref:'2:900:0:0:3:7:2',ha:6.4,cultivo:'Almendro',riego:'secano',variedad:'Guara',pend:4.8}]);
  var p=planNuevo(f,pabY0());
  /* punto 37: el programa del pistacho de regadío sale cuadrado (49,2 / 45,8 / 105,0 frente a 48,7 / 45 / 105); el almendro de secano (6,40 ha) está exento y se deja vacío */
  p.ud.forEach(function(u){
    if(u.ck==='pistacho'){u.su={f:hoyIso(),lab:'Laboratorio de ejemplo',inf:'0001',prof:'0-30',mo:1.2,ph:8.1,n:900,p:14,k:260,tex:'franca'};u.ne={mo:'cos',pr:1500,ed:1};u.ag={m3:2500,no3:20};
      u.pg={src:'own',rows:[{id:uid(),mo:'Marzo-abril',n:'Fosfato monoamónico (MAP) 12-61-0',d:75,rn:12,rp:61,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Abril-junio',n:'Nitrato amónico 33,5 %',d:120,rn:33.5,rp:0,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Mayo-julio',n:'Sulfato potásico 50 %',d:210,rn:0,rp:0,rk:50,fo:'Fertirrigación',eq:'Cabezal de riego'}]};mitPropuesta(u,true)}});
  UI.v='finca';UI.fid=f.id;UI.step='plan';UI.det=null; save(); saveUI(); window.__scrollTop=true; render(); undoToast('Finca de ejemplo creada (datos inventados)');
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

/* ───── v3.2 (propuesta 2.4): el semáforo del suelo (bajo, medio, alto…) y el «crédito del suelo» se han quitado: sus cortes no
   tenían fuente y no he encontrado una publicación oficial que los dé (ver INFORME.md). Lo que aporta el suelo se sigue pudiendo
   escribir a mano en «Aportes»; los planes guardados con un aporte del suelo puesto con el botón lo conservan tal cual. ───── */
/* «No tengo análisis de suelo» (u.suNo, opcional): 'nd' no dispongo; 'no20' no obligatorio (unidad de 20 ha o menos sin hoja de 5 ha).
   Solo vale si la unidad no tiene análisis; 'no20' solo si la unidad tiene 20 ha o menos (si crece, pasa a «No dispongo»). */
var SUNO_TXT={nd:'Sin análisis de suelo',no20:'Sin análisis de suelo: no obligatorio (RD 1311/2012, anexo III, parte I, A.3, datos generales, letra d)'};
function suNoMot(u,C){if(!u||!u.suNo||suTiene(u.su))return '';return u.suNo==='no20'&&C&&C.ha<=20+1e-9?'no20':'nd'}

/* ───── ayudas (?) con el artículo de la norma ───── */
/* cultivos para los que la guía del MAPA no da extracciones por tonelada (viña, triticale, otros frutales…): aviso en ámbar */
var SIN_EX_TXT='La guía del MAPA no da extracciones por tonelada para este cultivo';
var GLOS={
  ud:['Unidad de producción','Conjunto de recintos con el mismo cultivo y el mismo régimen (secano o regadío) que se abonan igual. El plan se hace por unidad.','RD 1051/2022, art. 3 v) y art. 6'],
  ex:['Unidad exenta','Las unidades de secano de 10 ha o menos no necesitan plan de abonado. Se pueden rellenar igualmente.','RD 1051/2022, art. 4.2 b)'],
  zv:['Zona vulnerable a nitratos','Zonas que designa la comunidad autónoma por el riesgo de contaminación por nitratos. En ellas hay un límite de N por cultivo y más restricciones. La herramienta lo mira sola en el SIGPAC del FEGA (consulta «Intersección de recinto con nitratos»), que da qué parte de cada recinto cae en la capa de zonas vulnerables. Es una ayuda: la referencia legal es la designación de la Junta de Comunidades de Castilla-La Mancha. Si un recinto toca solo por el borde, se enseña su porcentaje para que lo mires. Puedes cambiarlo a mano y se queda como lo pongas.','Decreto 80/2025 de Castilla-La Mancha, anejo III · SIGPAC (FEGA), consulta de nitratos'],
  cos:['Cosecha esperada','Lo que esperas recoger este año, en kg por ha. Con las extracciones por tonelada da lo que necesita el cultivo.','RD 1051/2022, art. 5.1 a)'],
  ag:['Agua de riego','Lo que piensas regar en la campaña, en m³ por ha. Si pones en el detalle los nitratos del agua, se descuenta el N que trae.','RD 1051/2022, art. 6.3'],
  su:['Análisis de suelo','El plan lleva datos del suelo de los recintos o, al menos, de un recinto representativo por cada hoja de cultivo: materia orgánica, pH, fósforo y potasio asimilables (los primeros 30 cm). Según la norma salen de mapas o registros provinciales y, si no los hay, de análisis cada 5 años como mínimo en regadío y cada 10 en secano. No hace falta el análisis si los recintos no pueden agruparse en una hoja de cultivo de al menos 5 ha y la unidad tiene 20 ha o menos. No hay mapa de suelos oficial de Castilla-La Mancha que la herramienta pueda consultar. Si no tienes análisis, marca «No tengo análisis de suelo». Si otra unidad tiene el mismo suelo, puedes usar su análisis.','RD 1051/2022, art. 6.2; RD 1311/2012, anexo III, parte I, A.3 (redacción del RD 92/2024)'],
  sinex:['Extracciones de referencia',SIN_EX_TXT+'. Escribe las tuyas: los kilos de N, P₂O₅ y K₂O que se lleva cada tonelada de cosecha (pídeselas a tu técnico si no las sabes).',''],
  tp:['Tipo de parcela','Tipo 1: suelo ligero (arenoso, franco-arenoso o franco) o precedente leguminosa. Tipo 2: el resto, y también si se fertirriga. Cambia el límite de N.','Decreto 80/2025 de Castilla-La Mancha, anejo III'],
  rp:['Recinto representativo','El recinto cuyo análisis de suelo vale para toda la unidad.','RD 1051/2022, art. 6.2'],
  abt:['Abono tipo o mi abono','«Abono tipo»: uno del listado, con la riqueza típica de ese tipo de abono; si tu etiqueta dice otra, cámbiala. «Mi abono»: el nombre comercial y la riqueza de la etiqueta; al pulsar «Hecho» se guarda en «Mis abonos» y sale el primero la próxima vez. En el PDF y en el Excel sale el nombre que pongas con su riqueza.',''],
  uf:['U.F. (unidades fertilizantes)','Kilos de nutriente puro por hectárea: dosis (kg/ha) × riqueza (%) ÷ 100. Por ejemplo, 100 kg/ha de un 15-15-15 son 15 U.F. de N, de P₂O₅ y de K₂O.',''],
  net:['Necesidad neta','Lo que necesita el cultivo (necesidad bruta) menos lo que ya aportan el suelo, el agua de riego, el estiércol y otros. Es lo que tiene que cubrir el programa.','RD 1051/2022, art. 5'],
  suelo:['Aporte del suelo','Lo que el suelo puede dar. Si no se pone, cuenta 0 y no se descuenta nada (es lo prudente). Si tu técnico calcula lo que aporta el suelo, escríbelo aquí.','RD 1051/2022, art. 5.1'],
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

/* ───── lo que falta (para el indicador único, «Hacer el plan» y «Ir a lo que falta») ─────
   s: 'fin' (Finca), 'dat' (tabla de Datos), 'det' (detalle del plan, ya relleno con lo habitual), 'plan' (filas del programa) */
function noExtr(u){var ne=u.ne||{};if((ne.mo||'cos')!=='cos'||PAB_EX[u.ck])return false;var ex=ne.ex||{};return !(+ex.n>0||+ex.p>0||+ex.k>0)}
function suTiene(su){su=su||{};return ['mo','n','p','k','ph'].some(function(k){return abHas(su[k])})}
function faltas(f,p,Cs){var L=[],seen={},add=function(s,t,sel,u,det){if(seen[t])return;seen[t]=1;L.push({s:s,t:t,sel:sel||'',u:u==null?null:u,det:det||null})},rs=fincaRecs(f.id),t=(f.datos||{}).tit||{};
  if(!rs.length)add('fin','los recintos de la finca','[data-act="ap-addrec"]');
  rs.forEach(function(r){if(!r.cultivo)add('fin','el cultivo del recinto '+refLabel(r.ref),'[data-r="'+r.ref+'"][data-k="cultivo"]')});
  if(!t.n)add('fin','el nombre del titular','[data-e="tit.n"]'); if(!t.nif)add('fin','el NIF del titular','[data-e="tit.nif"]');
  if(!p)return L;
  (p.ud||[]).forEach(function(u,i){var C=Cs[i],P='ud.'+i+'.';if(C.ex)return;var d0=u.d0===1,ne=u.ne||{},rows=(u.pg||{}).rows||[],q='«'+u.n+'»';
    if(!C.rs.length)add('det','algún recinto para '+q,'',i,'ud');
    if(!d0){
      if((ne.mo||'cos')==='cos'&&!abHas(ne.pr))add('dat','la cosecha esperada de '+q,'[data-p="'+P+'ne.pr"]',i);
      if(noExtr(u))add('dat','las extracciones de '+q,'[data-p="'+P+'ne.ex.n"]',i);
      if(u.rg==='regadio'&&!abHas((u.ag||{}).m3)&&!abHas((u.ag||{}).max))add('dat','el agua de riego de '+q,'[data-p="'+P+'ag.m3"]',i);
      if(!suTiene(u.su)&&!suNoMot(u,C))add('dat','el análisis de suelo de '+q,'[data-p="'+P+'su.mo"]',i);
      if(!u.rp&&C.rs.length)add('det','el recinto representativo de '+q,'',i,'su')}
    if(!u.prec)add('det','el cultivo precedente de '+q,'[data-p="'+P+'prec"]',i,'ne');
    if(!(u.mo&&u.mo.obj))add('det','el objetivo de materia orgánica','',i,'mo');
    if(!d0){
      rows.forEach(function(r,j){var R=P+'pg.rows.'+j+'.';if(!r.n)add('plan','el abono de una fila de '+q,'[data-p="'+R+'n"]',i);else if(!(+r.d>0))add('plan','la dosis de «'+r.n+'» en '+q,'[data-p="'+R+'d"]',i);else if(!r.mo)add('plan','el mes de «'+r.n+'» en '+q,'[data-p="'+R+'mo"]',i)});
      if(!PAB_MIT.some(function(x){return ((u.mit||{})[x[0]]||{}).a===1}))add('det','alguna medida del anexo V','',i,'mit');
      if(!PAB_RES.every(function(x){return ((u.res||{})[x[0]]||{}).a===1}))add('det','comprobar las restricciones de aplicación','',i,'res')}});
  if(!pabIso(p.fe))add('det','la fecha del plan','[data-p="fe"]',null,'fi');
  var nec=(p.ud||[]).some(function(u,i){if(Cs[i].ex)return false;return pabCampRange(p)[1]>=(u.zv?'2027-09-01':'2028-09-01')}), A=pabAs();
  if(nec&&!(A.n&&A.regfer))add('det','el asesor en fertilización (nombre y nº REGFER)','[data-p="AS.n"]',null,'fi');
  return L}
function capi(s){s=String(s||'');return s.charAt(0).toUpperCase()+s.slice(1)}
/* estado de cada unidad: «Cuadra» solo si hay cosecha, hay abonos y los tres nutrientes cuadran (fallos 3 y 4) */
function udEst(f,p,u,C){var rows=(u.pg||{}).rows||[], ne=u.ne||{};
  if(C.ex&&!rows.length)return {k:'ex',t:'No necesita plan'};
  if(u.d0===1)return {k:'ok',t:'Sin abonado'};
  if((ne.mo||'cos')==='cos'&&!(+ne.pr>0))return {k:'w',t:'Falta la cosecha'};
  if(noExtr(u))return {k:'w',t:'Faltan las extracciones'};
  var nec=['n','p','k'].some(function(k){return C.B[k].net>1});
  if(!rows.length)return nec?{k:'w',t:'Sin abonos'}:{k:'ok',t:'No necesita abono'};
  if(!C.lok)return {k:'bad',t:'Pasa del límite de N'};
  if(pabPerHits(f,u,p,C).length)return {k:'bad',t:'N en meses prohibidos'};
  if(['n','p','k'].some(function(k){return C.B[k].st!=='Ajustado'}))return {k:'w',t:'Revisa el balance'};
  return {k:'ok',t:'Cuadra'}}
/* el indicador único (punto 19) */
function estGlobal(f,p,Cs,F,E){
  if(!fincaRecs(f.id).length)return {k:'w',t:'Empieza por los recintos',s:'Del PDF de la PAC, del mapa o a mano'};
  if(F.length)return {k:'w',f:1,t:F.length===1?'Falta 1 dato':'Faltan '+F.length+' datos',s:capi(F[0].t)};
  var idx=(p?p.ud||[]:[]).map(function(u,i){return i}).filter(function(i){return !Cs[i].ex});
  if(!idx.length)return {k:'ok',t:'Plan listo',s:'Ninguna unidad necesita plan'};
  if(idx.every(function(i){var u=p.ud[i];return u.d0!==1&&!((u.pg||{}).rows||[]).length&&E[i].t!=='No necesita abono'}))return {k:'w',t:'Falta hacer el plan',s:'Pulsa «Hacer el plan» en Datos'};
  var bad=idx.filter(function(i){return E[i].k!=='ok'});
  if(bad.length)return {k:'w',t:'Revisa el balance',s:p.ud[bad[0]].n+': '+E[bad[0]].t.toLowerCase()};
  return {k:'ok',t:'Plan listo',s:'Las cuentas cuadran'}}

/* ───── «Ajustar a lo que falta» (punto 17): los tres nutrientes a la vez, subiendo o bajando las dosis que ya hay ─────
   Se busca la combinación de dosis (≥ 0) más cercana a la necesidad neta de N, P₂O₅ y K₂O; solo se añade un abono si ningún
   abono mineral del programa lleva ese nutriente. Nunca se repite una fila. En zona vulnerable el N no pasa del límite. */
function ajustar(f,p,u){
  var C=pabCalc(f,u,p), rg=u.rg==='regadio', K=['n','p','k'], R={n:'rn',p:'rp',k:'rk'}, rows=((u.pg||{}).rows||[]).map(function(q){return Object.assign({},q)}), added=[], quit=[];
  var tgt={};K.forEach(function(k){tgt[k]=C.B[k].net}); var lim=u.zv&&C.L.ok?C.L.lim:null; if(lim!=null)tgt.n=Math.min(tgt.n,lim);
  var fo=rg?'Fertirrigación':'Suelo', eq=rg?'Cabezal de riego':'Abonadora', pm=perMeses(u);
  K.forEach(function(k){if(tgt[k]>1&&!rows.some(function(q){return q.org!==1&&+q[R[k]]>0})){
    var nm={n:rg?'Nitrato amónico 33,5 %':'Nitrato amónico cálcico (NAC) 27 %',p:rg?'Fosfato monoamónico (MAP) 12-61-0':'Superfosfato triple 46 %',k:'Sulfato potásico 50 %'}[k], x=catN(nm);
    var mo={n:rg?'Abril-junio':'Febrero-marzo',p:rg?'Marzo-abril':'Febrero',k:rg?'Mayo-julio':'Febrero'}[k];
    if(+x.rn>0&&mesesDe(mo).some(function(m){return pm.indexOf(m)>=0}))mo=rg?'Marzo':'Febrero';
    rows.push({id:uid(),mo:mo,n:x.n,d:0,rn:x.rn,rp:x.rp,rk:x.rk,fo:fo,eq:eq,nuevo:1});added.push(x.n)}});
  var V=rows.map(function(q,j){return j}).filter(function(j){return rows[j].org!==1&&(+rows[j].rn>0||+rows[j].rp>0||+rows[j].rk>0)});
  var fix={n:0,p:0,k:0};rows.forEach(function(q,j){if(V.indexOf(j)<0)K.forEach(function(k){fix[k]+=(+q.d||0)*(+q[R[k]]||0)/100})});
  var A=function(k,j){return (+rows[j][R[k]]||0)/100}, d={}, d0={};V.forEach(function(j){d[j]=+rows[j].d||0;d0[j]=d[j]});
  var dev=function(k){var s=fix[k];V.forEach(function(j){s+=d[j]*A(k,j)});return s-tgt[k]};
  /* mínimos cuadrados con dosis no negativas (descenso por coordenadas) */
  var lam=1e-5;
  for(var it=0;it<4000;it++){var mv=0;V.forEach(function(j){var g=2*lam*(d[j]-d0[j]),h=2*lam;K.forEach(function(k){var a=A(k,j);g+=2*a*dev(k);h+=2*a*a});if(h<=0)return;var nv=Math.max(0,d[j]-g/h);mv=Math.max(mv,Math.abs(nv-d[j]));d[j]=nv});if(mv<1e-7)break}
  /* redondeo a 5 kg y retoque fino para que cada nutriente quede a ±1 U.F. */
  V.forEach(function(j){d[j]=d[j]<2.5?0:Math.round(d[j]/5)*5});
  var cost=function(){var c=0;K.forEach(function(k){var v=Math.abs(dev(k));c+=Math.max(0,v-0.9)*Math.max(0,v-0.9)*100+v*v*0.001});if(lim!=null){var nn=dev('n')+tgt.n;if(nn>lim+0.05)c+=1e6}V.forEach(function(j){if(d[j]%5)c+=0.0005});return c};
  for(var g=0;g<400;g++){var best=cost(),bj=-1,bv=0;V.forEach(function(j){[5,-5,1,-1,10,-10].forEach(function(s){var o=d[j],n=o+s;if(n<0)return;d[j]=n;var c=cost();if(c<best-1e-9){best=c;bj=j;bv=n}d[j]=o})});if(bj<0)break;d[bj]=bv}
  V.forEach(function(j){rows[j].d=d[j]});
  var out=rows.filter(function(q,j){if(V.indexOf(j)>=0&&!(d[j]>0)){if(!q.nuevo)quit.push(q.n);return false}return true}).map(function(q){delete q.nuevo;return q});
  var off=K.filter(function(k){return Math.abs(dev(k))>1});
  return {rows:out,added:added.filter(function(n){return out.some(function(q){return q.n===n})}),quit:quit,off:off,limN:lim!=null&&C.B.n.net>lim}}

/* ───── NIF, NIE o CIF: aviso si la letra o el dígito de control no cuadran (punto 39; no bloquea) ───── */
function nifMal(v){v=String(v||'').toUpperCase().replace(/[\s.\-]/g,'');if(!v)return '';var L='TRWAGMYFPDXBNJZSQVHLCKE',m;
  if((m=/^(\d{8})([A-Z])$/.exec(v)))return L[+m[1]%23]===m[2]?'':'La letra del NIF no cuadra con los números: revisa que esté bien escrito.';
  if((m=/^([XYZ])(\d{7})([A-Z])$/.exec(v)))return L[+('XYZ'.indexOf(m[1])+m[2])%23]===m[3]?'':'La letra del NIE no cuadra con los números: revisa que esté bien escrito.';
  if((m=/^([ABCDEFGHJNPQRSUVW])(\d{7})([0-9A-J])$/.exec(v))){var s=0;for(var i=0;i<7;i++){var x=+m[2][i];if(i%2===0){x*=2;s+=Math.floor(x/10)+x%10}else s+=x}var c=(10-s%10)%10,ch='JABCDEFGHI'[c];
    var ok=/[NPQRSW]/.test(m[1])?m[3]===ch:(/[ABEH]/.test(m[1])?m[3]===String(c):(m[3]===String(c)||m[3]===ch));return ok?'':'El dígito de control del CIF no cuadra: revisa que esté bien escrito.'}
  return 'No parece un NIF, NIE ni CIF: revisa que esté bien escrito.'}
/* ───── lista de la compra y coste (solo en pantalla) ───── */
function compra(f,p,Cs){var m={},o=[];(p.ud||[]).forEach(function(u,i){var C=Cs[i];if(u.d0===1)return;((u.pg||{}).rows||[]).forEach(function(q){var d=+q.d||0;if(!q.n||!(d>0))return;var k=noAcc(q.n).toLowerCase().trim();
    if(!m[k]){var c=catDe(q.n);m[k]={n:q.n,kg:0,eur:0,sinPz:0,dens:c&&c.dens||0,uds:{}};o.push(m[k])}var kg=d*C.ha;m[k].kg+=kg;m[k].uds[u.n]=1;if(abHas(q.pz))m[k].eur+=kg*q.pz;else m[k].sinPz=1})});
  return o.sort(function(a,b){return b.kg-a.kg})}
function costeUd(u,C){var e=0,sin=0;((u.pg||{}).rows||[]).forEach(function(q){if(!(+q.d>0))return;if(abHas(q.pz))e+=q.d*q.pz;else sin=1});return {ha:e,tot:e*C.ha,sin:sin}}

/* ───── iconos (trazo, del color del texto; los mismos que el diseño aprobado) ───── */
var ICP={
  check:'<path d="M20 6 9 17l-5-5"/>',
  leaf:'<path d="M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.48 19 2c1 2 2 4.18 2 8 0 5.5-4.78 10-10 10Z"/><path d="M2 21c0-3 1.85-5.36 5.08-6C9.5 14.52 12 13 13 12"/>',
  download:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/>',
  upload:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/>',
  plus:'<path d="M5 12h14M12 5v14"/>',
  scale:'<path d="m16 16 3-8 3 8c-.87.65-1.920 1-3 1s-2.130-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.920 1-3 1s-2.130-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>',
  refresh:'<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>',
  grid:'<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
  folder:'<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.900a2 2 0 0 1-1.690-.9L9.600 3.900A2 2 0 0 0 7.930 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
  bag:'<path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4Z"/><path d="M3 6h18"/><path d="M16 10a4 4 0 0 1-8 0"/>',
  save:'<path d="M15.200 3a2 2 0 0 1 1.400.6l3.800 3.800a2 2 0 0 1 .6 1.400V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7"/><path d="M7 3v4a1 1 0 0 0 1 1h7"/>',
  next:'<path d="m9 18 6-6-6-6"/>', back:'<path d="m15 18-6-6 6-6"/>', down:'<path d="m6 9 6 6 6-6"/>',
  alert:'<path d="m21.730 18-8-14a2 2 0 0 0-3.480 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.730-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
  pdf:'<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M8 13h8M8 17h8"/>',
  map:'<path d="M14.106 5.553a2 2 0 0 0 1.788 0l3.659-1.830A1 1 0 0 1 21 4.619v12.764a1 1 0 0 1-.553.894l-4.553 2.277a2 2 0 0 1-1.788 0l-4.212-2.106a2 2 0 0 0-1.788 0l-3.659 1.830A1 1 0 0 1 3 19.381V6.618a1 1 0 0 1 .553-.894l4.553-2.277a2 2 0 0 1 1.788 0z"/><path d="M15 5.764v15"/><path d="M9 3.236v15"/>',
  cart:'<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.050 2.050h2l2.660 12.420a2 2 0 0 0 2 1.580h9.780a2 2 0 0 0 1.950-1.570l1.650-7.430H5.120"/>',
  undo:'<path d="M9 14 4 9l5-5"/><path d="M4 9h10.500a5.500 5.500 0 0 1 0 11H11"/>',
  trash:'<path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
  close:'<path d="M18 6 6 18M6 6l12 12"/>', search:'<circle cx="11" cy="11" r="7"/><path d="m21 21-4.300-4.300"/>',
  loc:'<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/>', copy:'<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M4 16V5a2 2 0 0 1 2-2h11"/>',
  cal:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>', euro:'<path d="M17 6a7 7 0 1 0 0 12"/><path d="M4 10h9M4 14h9"/>',
  more:'<circle cx="5" cy="12" r="1.300"/><circle cx="12" cy="12" r="1.300"/><circle cx="19" cy="12" r="1.300"/>', edit:'<path d="M12 20h9"/><path d="M16.500 3.500a2.120 2.120 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  spark:'<path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.500 2.500M15.500 15.500 18 18M6 18l2.500-2.500M15.500 8.500 18 6"/>'};
ICP.xls=ICP.grid; ICP.up=ICP.upload;
function ic(n){return '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true">'+(ICP[n]||'')+'</svg>'}

/* ───── glifo y color de cada cultivo (punto 43) ───── */
var GL={
  pistacho:{c:'#8fb23a',bg:'#eaf2d3',s:'<path d="M12 2.5c4.2 3 6.5 6.6 6.5 10.6a6.5 6.5 0 0 1-13 0C5.5 9.1 7.8 5.5 12 2.5Z" fill="#8fb23a"/><path d="M12 7.5c1.8 2.2 2.2 6 0 11.5C9.8 13.5 10.2 9.7 12 7.5Z" fill="#4e6a17"/>'},
  almendro:{c:'#d9965a',bg:'#f8e8d8',s:'<g transform="rotate(24 12 12)"><path d="M12 2.5c3.8 3.5 5.5 7 5.5 10.5 0 4.2-2.5 8.5-5.5 8.5s-5.5-4.3-5.5-8.5C6.5 9.5 8.2 6 12 2.5Z" fill="#d9965a"/><path d="M10 10.5c-.5 2-.5 4 0 6M13.2 8.5c.8 3 .8 6.5 0 9.5" stroke="#9a5f2c" stroke-width="1.2" fill="none" stroke-linecap="round"/></g>'},
  olivo:{c:'#6f7d3a',bg:'#e9ecd8',s:'<path d="M4 20C8 16 13 10 20 4" stroke="#5b6b2e" stroke-width="1.5" fill="none" stroke-linecap="round"/><ellipse cx="8.5" cy="11" rx="3.2" ry="4.4" transform="rotate(-35 8.5 11)" fill="#6f7d3a"/><ellipse cx="15" cy="16" rx="3" ry="4.1" transform="rotate(-35 15 16)" fill="#3f4a1f"/><path d="M12.5 10.5c2-3 5-4.2 7.2-4-1 3-4 5-7.2 4Z" fill="#a3b465"/>'},
  nogal:{c:'#9a7348',bg:'#efe4d6',s:'<circle cx="12" cy="12.5" r="8" fill="#9a7348"/><path d="M12 4.5v16" stroke="#6b4c2a" stroke-width="1.4"/><path d="M8 8.2c1.6 2.2 1.6 6.6 0 8.8M16 8.2c-1.6 2.2-1.6 6.6 0 8.8" stroke="#6b4c2a" stroke-width="1.1" fill="none" stroke-linecap="round"/>'},
  vid:{c:'#7a4a9c',bg:'#efe6f5',s:'<path d="M12 2.5v3.5M12 4c2-1.6 4.4-1.6 6 0-2 1.2-4 1.4-6 0Z" stroke="#4e7a2a" stroke-width="1.3" fill="#7fae4a"/><g fill="#7a4a9c"><circle cx="8" cy="8.6" r="2.4"/><circle cx="12.5" cy="8.6" r="2.4"/><circle cx="17" cy="8.6" r="2.4"/><circle cx="10.2" cy="12.6" r="2.4"/><circle cx="14.8" cy="12.6" r="2.4"/><circle cx="12.5" cy="16.6" r="2.4"/><circle cx="12.5" cy="20.2" r="1.8"/></g>'},
  melocotonero:{c:'#ee9a55',bg:'#fdebdc',s:'<circle cx="12" cy="13.5" r="7.5" fill="#ee9a55"/><path d="M12 6.5c-1.6 3-1.6 9 0 14" stroke="#c96a2f" stroke-width="1.2" fill="none"/><path d="M12 6c1.2-2.6 3.6-3.6 6-3.2-.8 2.4-3.2 3.6-6 3.2Z" fill="#6a9a3a"/>'},
  frutal:{c:'#d0473b',bg:'#f9e1de',s:'<path d="M12 7.5c-2.5-1.8-7-1-7 4.5 0 4.5 3 9 5 9 1 0 1.3-.5 2-.5s1 .5 2 .5c2 0 5-4.5 5-9 0-5.5-4.5-6.3-7-4.5Z" fill="#d0473b"/><path d="M12 7.5V4" stroke="#6b4c2a" stroke-width="1.4" stroke-linecap="round"/><path d="M12.5 5.5c1-2 3-2.8 5-2.5-.6 2-2.6 3-5 2.5Z" fill="#6a9a3a"/>'},
  trigo:{c:'#c9a227',bg:'#f7efd3',s:'<path d="M12 22V8" stroke="#a5841c" stroke-width="1.4" stroke-linecap="round"/><g fill="#c9a227"><ellipse cx="9.6" cy="9" rx="1.9" ry="3" transform="rotate(-25 9.6 9)"/><ellipse cx="14.4" cy="9" rx="1.9" ry="3" transform="rotate(25 14.4 9)"/><ellipse cx="9.6" cy="13.4" rx="1.9" ry="3" transform="rotate(-25 9.6 13.4)"/><ellipse cx="14.4" cy="13.4" rx="1.9" ry="3" transform="rotate(25 14.4 13.4)"/><ellipse cx="12" cy="5" rx="1.7" ry="2.8"/></g>'},
  trigoduro:{c:'#b5832a',bg:'#f5e8d0',s:'<path d="M12 22V8" stroke="#8a6420" stroke-width="1.4" stroke-linecap="round"/><path d="M9 7 6.5 2.5M15 7l2.5-4.5M12 4V1" stroke="#8a6420" stroke-width="1" stroke-linecap="round"/><g fill="#b5832a"><ellipse cx="9.6" cy="9.5" rx="1.9" ry="3" transform="rotate(-25 9.6 9.5)"/><ellipse cx="14.4" cy="9.5" rx="1.9" ry="3" transform="rotate(25 14.4 9.5)"/><ellipse cx="9.6" cy="14" rx="1.9" ry="3" transform="rotate(-25 9.6 14)"/><ellipse cx="14.4" cy="14" rx="1.9" ry="3" transform="rotate(25 14.4 14)"/><ellipse cx="12" cy="5.6" rx="1.7" ry="2.6"/></g>'},
  cebada:{c:'#cdb25a',bg:'#f6f0d8',s:'<path d="M12 22V9" stroke="#9c8636" stroke-width="1.4" stroke-linecap="round"/><path d="M10 10 4 4M14 10l6-6M10 14l-6-5M14 14l6-5M12 6V1" stroke="#b29a45" stroke-width=".9" stroke-linecap="round"/><g fill="#cdb25a"><ellipse cx="10.2" cy="10.5" rx="1.6" ry="2.6" transform="rotate(-30 10.2 10.5)"/><ellipse cx="13.8" cy="10.5" rx="1.6" ry="2.6" transform="rotate(30 13.8 10.5)"/><ellipse cx="10.2" cy="14.5" rx="1.6" ry="2.6" transform="rotate(-30 10.2 14.5)"/><ellipse cx="13.8" cy="14.5" rx="1.6" ry="2.6" transform="rotate(30 13.8 14.5)"/><ellipse cx="12" cy="7" rx="1.4" ry="2.3"/></g>'},
  avena:{c:'#a9a35e',bg:'#f0efdc',s:'<path d="M12 22V4" stroke="#7d7a3e" stroke-width="1.3" stroke-linecap="round"/><path d="M12 7c-2 1-4 3-4.5 5M12 7c2 1 4 3 4.5 5M12 12c-2 1-3.6 3-4 5M12 12c2 1 3.6 3 4 5" stroke="#7d7a3e" stroke-width="1" fill="none"/><g fill="#a9a35e"><ellipse cx="7.5" cy="13" rx="1.4" ry="2.4"/><ellipse cx="16.5" cy="13" rx="1.4" ry="2.4"/><ellipse cx="8" cy="18" rx="1.3" ry="2.2"/><ellipse cx="16" cy="18" rx="1.3" ry="2.2"/></g>'},
  centeno:{c:'#8f9160',bg:'#eceddf',s:'<path d="M12 22V7" stroke="#6b6d43" stroke-width="1.3" stroke-linecap="round"/><path d="M12 5V1M10.5 6 9 2M13.5 6 15 2" stroke="#6b6d43" stroke-width=".9" stroke-linecap="round"/><g fill="#8f9160"><ellipse cx="10.6" cy="8" rx="1.3" ry="2.4" transform="rotate(-20 10.6 8)"/><ellipse cx="13.4" cy="8" rx="1.3" ry="2.4" transform="rotate(20 13.4 8)"/><ellipse cx="10.6" cy="11.6" rx="1.3" ry="2.4" transform="rotate(-20 10.6 11.6)"/><ellipse cx="13.4" cy="11.6" rx="1.3" ry="2.4" transform="rotate(20 13.4 11.6)"/><ellipse cx="10.6" cy="15.2" rx="1.3" ry="2.4" transform="rotate(-20 10.6 15.2)"/><ellipse cx="13.4" cy="15.2" rx="1.3" ry="2.4" transform="rotate(20 13.4 15.2)"/></g>'},
  triticale:{c:'#b39b4e',bg:'#f3eed9',s:'<path d="M12 22V8" stroke="#86723a" stroke-width="1.4" stroke-linecap="round"/><path d="M10 8 8 3M14 8l2-5" stroke="#86723a" stroke-width=".9" stroke-linecap="round"/><g fill="#b39b4e"><rect x="8.2" y="8" width="3.4" height="4.2" rx="1.6"/><rect x="12.4" y="8" width="3.4" height="4.2" rx="1.6"/><rect x="8.2" y="12.6" width="3.4" height="4.2" rx="1.6"/><rect x="12.4" y="12.6" width="3.4" height="4.2" rx="1.6"/><ellipse cx="12" cy="5.6" rx="1.6" ry="2.5"/></g>'},
  maiz:{c:'#e2b93b',bg:'#fbf1d2',s:'<path d="M12 3c3 0 4.5 4 4.5 9s-1.5 9-4.5 9-4.5-4-4.5-9S9 3 12 3Z" fill="#e2b93b"/><path d="M9 8h6M8.6 11h6.8M8.6 14h6.8M9 17h6M12 3.5v17" stroke="#c49a24" stroke-width=".8"/><path d="M7.8 21C5 18 4 13 5 8c1.5 3 2.5 6 4.5 9M16.2 21C19 18 20 13 19 8c-1.5 3-2.5 6-4.5 9" fill="#6f9e3c"/>'},
  barbecho:{c:'#9b7a55',bg:'#eee5da',s:'<rect x="3" y="5" width="18" height="14" rx="2.5" fill="#b89a74"/><path d="M5 9c2.5-1.2 4.5 1.2 7 0s4.5-1.2 7 0M5 13c2.5-1.2 4.5 1.2 7 0s4.5-1.2 7 0M5 17c2.5-1.2 4.5 1.2 7 0s4.5-1.2 7 0" stroke="#7d5e3b" stroke-width="1.2" fill="none" stroke-linecap="round"/>'},
  otro:{c:'#5f8f45',bg:'#e5efdd',s:'<path d="M12 21v-9" stroke="#4d6f37" stroke-width="1.6" stroke-linecap="round"/><path d="M12 13C12 8 8.5 5.5 4 6c0 4.5 3.5 7.5 8 7Z" fill="#5f8f45"/><path d="M12 11c0-4 3-6.5 7.5-6 0 4-3 6.5-7.5 6Z" fill="#86b35f"/>'}};
function glKey(c){if(/barbech/i.test(String(c||'')))return 'barbecho';var k=pabCk(c);return GL[k]?k:'otro'}
function glOf(c){return GL[glKey(c)]}
function glyph(c,cls){var g=glOf(c);return '<span class="gl'+(cls?' '+cls:'')+'" style="--glb:'+g.bg+'"><svg viewBox="0 0 24 24" aria-hidden="true">'+g.s+'</svg></span>'}
function udCult(f,u){var r=(u.refs||[]).map(function(x){return S.recintos[x]}).filter(Boolean)[0];return r?r.cultivo:pabCkName(u.ck)}

/* ───── pasos ───── */
var STEPS=[['fin','Finca'],['dat','Datos'],['plan','Plan y PDF']];
var CTX=null, SEL={}, RM={}, RMR={}, RMODE={}, HL=[];
/* modo de la fila que se edita: lo elegido en el desplegable o, si no, «tipo» si el nombre es del listado (o no hay nombre) */
function abModo(q){return RMODE[q.id]||(!q.n||catEsTipo(q.n)?'tipo':'mio')}
function isWide(){return window.matchMedia&&window.matchMedia('(min-width:1200px)').matches}
function stepIdx(){for(var i=0;i<STEPS.length;i++)if(STEPS[i][0]===UI.step)return i;return 0}
function ctxOf(f){var p=pabCur(f),Cs=p?(p.ud||[]).map(function(u){return pabCalc(f,u,p)}):[],F=faltas(f,p,Cs),E=Cs.map(function(C,i){return udEst(f,p,p.ud[i],C)});return {p:p,Cs:Cs,F:F,E:E,G:estGlobal(f,p,Cs,F,E)}}
function nFalta(ctx,s){return ctx.F.filter(function(x){return s.indexOf(x.s)>=0}).length}
function udsOrden(p,Cs){return (p.ud||[]).map(function(u,i){return i}).sort(function(a,b){return (Cs[a].ex?1:0)-(Cs[b].ex?1:0)||a-b})}
function nRec(n){return n+(n===1?' recinto':' recintos')}
function haSum(rs){return rs.reduce(function(a,r){return a+haC(r)},0)}
function chipEst(e){var c=e.k==='ok'?'ok':(e.k==='ex'?'m':(e.k==='bad'?'r':'w'));return '<span class="chip '+c+'">'+(c==='ok'?ic('check'):(c==='m'?'':ic('alert')))+esc(e.t)+'</span>'}
/* el plan de la campaña se crea solo cuando la finca ya tiene recintos (no hay paso «crear el plan») */
function planAuto(f){if((f.pab||[]).length||!fincaRecs(f.id).length)return;planNuevo(f,pabY0());saveRaw()}

/* ───── dibujo general: se rehace todo, pero sin perder la casilla, lo tecleado ni el cursor (fallo 1, punto 33) ───── */
var RTO=null;
function renderLuego(){if(RTO)return;RTO=setTimeout(function(){RTO=null;render()},0)}
function fkey(el){if(!el||el===document.body||!el.dataset)return null;var d=el.dataset,q=function(n,v){return '['+n+'="'+String(v).replace(/\\/g,'\\\\').replace(/"/g,'\\"')+'"]'};
  if(el.tagName==='BUTTON'&&d.act){var s='button'+q('data-act',d.act);['u','j','s','k','id','ref','i','m','v','p'].forEach(function(x){if(d[x]!=null)s+=q('data-'+x,d[x])});return s}
  if(d.p)return el.tagName.toLowerCase()+q('data-p',d.p);
  if(d.e)return q('data-e',d.e); if(d.r)return q('data-r',d.r)+q('data-k',d.k||'');
  if(el.id)return '#'+(window.CSS&&CSS.escape?CSS.escape(el.id):el.id);
  if(d.act){var t=q('data-act',d.act);['u','j','s','k','id','ref','i','m'].forEach(function(x){if(d[x]!=null)t+=q('data-'+x,d[x])});return t}
  if(d.i!=null&&d.k)return el.tagName.toLowerCase()+'.'+String(el.className).split(' ')[0]+q('data-i',d.i)+q('data-k',d.k);
  if(el.tagName==='SUMMARY'&&el.parentNode.dataset&&el.parentNode.dataset.o)return 'details'+q('data-o',el.parentNode.dataset.o)+'>summary';
  /* lo demás: su camino dentro de la zona (sirve mientras no cambie la forma de la página) */
  var path=[],x=el;while(x&&x.parentNode&&!/^(app|side|desk|sheet)$/.test(x.id||'')){var i=1,y=x;while((y=y.previousElementSibling))i++;path.unshift(x.tagName.toLowerCase()+':nth-child('+i+')');x=x.parentNode}
  return x&&x.id&&path.length?'#'+x.id+'>'+path.join('>'):null}
function focoGuarda(){var a=document.activeElement;if(!a||a===document.body)return null;var k=fkey(a);if(!k)return null;var o={k:k,el:a};
  if(a.tagName==='INPUT'||a.tagName==='TEXTAREA'){try{o.s=a.selectionStart;o.e=a.selectionEnd}catch(e){}if(a.type!=='checkbox'&&a.value!==a.defaultValue)o.v=a.value}return o}
function focoPon(o){if(!o)return;if(document.body.contains(o.el)&&document.activeElement===o.el)return;var el=null;try{el=document.querySelector(o.k)}catch(e){}if(!el)return;
  if(o.v!=null&&el.value!==o.v)el.value=o.v;try{el.focus({preventScroll:true})}catch(e){el.focus()}if(o.s!=null)try{el.setSelectionRange(o.s,o.e)}catch(e){}}
function render(){
  if(RTO){clearTimeout(RTO);RTO=null}
  var fo=focoGuarda(), sy=window.scrollY, ss=$('#side')?$('#side').scrollTop:0;
  var f=UI.v==='finca'?finca(UI.fid):null; if(!f){UI.v='inicio';UI.fid=null}
  EXP.fid=UI.fid; var B=document.body; B.classList.toggle('vf',!!f); B.classList.toggle('vi',!f);
  if(f){UI.step=stepOf(UI.step);pabSync(f);planAuto(f);if(zvAplicar(f))saveRaw();CTX=ctxOf(f);B.dataset.step=UI.step;
    $('#side').innerHTML=sideHTML(f,CTX); $('#app').innerHTML=fixHTML(pasoHTML(f,CTX)); $('#desk').innerHTML=deskHTML(f,CTX)}
  else{CTX=null;delete B.dataset.step;$('#side').innerHTML='';$('#desk').innerHTML='';$('#app').innerHTML=inicioHTML()}
  document.title=f?f.nombre+' · Plan de abonado':'Plan de abonado';
  if(window.__scrollTop){window.scrollTo(0,0);window.__scrollTop=false;ss=0}else window.scrollTo(0,sy);
  if($('#side'))$('#side').scrollTop=ss;
  stkLab(); inThou(); marcaValida(); aerMonta(); prevAuto(); focoPon(fo); hlAuto();
}
function renderHdr(){var e=$('#saved');if(e)e.innerHTML=savedHTML();var u=$('#side [data-act="ap-undo"]');if(u)u.disabled=!UNDO.length}
function savedHTML(){return FS.ok?ic('check')+'Guardado en «'+esc(FS.name)+'»':(FS.h?'<button class="lnk" data-act="ap-fs-res">'+ic('alert')+'Volver a guardar en «'+esc(FS.name)+'»</button>':ic('check')+'Guardado en este ordenador')}

/* ───── barra lateral: marca, finca, los tres pasos y el indicador único ───── */
function campSel(f,p){var l=pabList(f);if(!l.length)return '<div class="camp">Campaña '+esc(pabCamp(pabY0()))+'</div>';
  return '<label class="camp"><span class="vh">Campaña</span><select class="ap-camp" aria-label="Campaña">'+l.map(function(x){return '<option value="'+esc(x.id)+'"'+(p&&x.id===p.id?' selected':'')+'>Campaña '+esc(pabCamp(x.y))+'</option>'}).join('')+'<option value="new">Plan de otra campaña…</option></select>'+ic('down')+'</label>'}
function sideHTML(f,ctx){var p=ctx.p,rs=fincaRecs(f.id),G=ctx.G,nf0=nFalta(ctx,['fin']),nd=nFalta(ctx,['dat']),hasPg=p&&(p.ud||[]).some(function(u){return ((u.pg||{}).rows||[]).length}),nu=p?(p.ud||[]).filter(function(u,i){return !ctx.Cs[i].ex}).length:0;
  var sinCopia=S.fincas.length&&S.mod>(S.copia||0)+1000&&!FS.ok;
  var h='<button class="brand" data-act="ap-home" title="Mis fincas"><span class="lg">'+ic('leaf')+'</span>Plan de abonado</button>';
  h+='<div class="fin"><div class="aer" data-aer="side" data-fid="'+esc(f.id)+'"></div><div class="tx"><b>'+esc(f.nombre)+'</b><span>'+ha2(haSum(rs))+' ha · '+nRec(rs.length)+'</span>'+campSel(f,p)+'</div></div>';
  var st=[{s:'fin',ok:!nf0&&rs.length,sm:nf0?(nf0===1?'Falta 1':'Faltan '+nf0):nRec(rs.length)},{s:'dat',ok:!nf0&&!nd&&rs.length,sm:nd?(nd===1?'Falta 1':'Faltan '+nd):'Todo puesto'},{s:'plan',ok:G.k==='ok',sm:hasPg?(nu===1?'1 unidad':nu+' unidades'):'Cuando estén los datos'}];
  h+='<div class="cap">Pasos</div><div class="steps">'+STEPS.map(function(s,i){var o=st[i],on=UI.step===s[0];return '<button class="s'+(on?' on':(o.ok?' ok':(i===2&&!hasPg?' todo':'')))+'" data-act="ap-step" data-s="'+s[0]+'"'+(on?' aria-current="step"':'')+'><i>'+(o.ok&&!on?ic('check'):(i+1))+'</i><div>'+s[1]+'<small>'+esc(o.sm)+'</small></div></button>'}).join('')+'</div>';
  h+='<button class="state'+(G.k==='ok'?'':' w')+'" id="estado" data-act="ap-falta"'+(G.f||G.k!=='ok'?'':' tabindex="-1"')+'><span class="dot"></span><div><b>'+esc(G.t)+'</b><span>'+esc(G.s)+'</span></div></button>';
  h+='<div class="grow"></div><div class="nav"><button data-act="ap-home">'+ic('folder')+'Mis fincas</button><button data-act="ap-mis">'+ic('bag')+'Mis abonos</button><button data-act="ap-copias" title="Guardar copia (Ctrl+S)">'+ic('save')+'Guardar copia'+(sinCopia?'<i class="wd" title="Hay cambios sin copia en archivo"></i>':'')+'</button><button data-act="ap-undo" title="Deshacer (Ctrl+Z)"'+(UNDO.length?'':' disabled')+'>'+ic('undo')+'Deshacer</button></div>';
  h+='<div class="saved" id="saved">'+savedHTML()+'</div>';
  return h}

function pasoHTML(f,ctx){var s=UI.step;
  if(s==='dat')return datHTML(f,ctx); if(s==='plan')return planHTML(f,ctx); return finHTML(f,ctx)}
function cabeza(n,t,l){return '<div class="eyebrow">Paso '+n+' de 3 · '+STEPS[n-1][1]+'</div><h1>'+t+'</h1>'+(l?'<p class="lead">'+l+'</p>':'')}

/* ───── paso 1: Finca (titular y recintos) ───── */
function inp(attr,val,ph,mode,cls){return '<input class="in ap-in'+(cls?' '+cls:'')+'" '+attr+(mode?' inputmode="'+mode+'"':'')+' value="'+esc(val==null?'':(typeof val==='number'?abEd(val):val))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off">'}
function fld(lbl,inner,cls){return '<label class="fld'+(cls?' '+cls:'')+'"><span>'+lbl+'</span>'+inner+'</label>'}
function recOrden(rs){var k=(UI.sort||'ref').replace(/^-/,''),d=/^-/.test(UI.sort||'')?-1:1;
  return rs.slice().sort(function(a,b){var v=k==='ha'?haC(a)-haC(b):k==='cultivo'?String(a.cultivo||'~').localeCompare(String(b.cultivo||'~'),'es'):k==='riego'?String(a.riego||'').localeCompare(String(b.riego||'')):0;return (v||a.ref.localeCompare(b.ref,'es',{numeric:true}))*d})}
function thS(k,t,cls){var on=(UI.sort||'ref').replace(/^-/,'')===k,d=/^-/.test(UI.sort||'');return '<th'+(cls?' class="'+cls+'"':'')+'><button class="ths'+(on?' on':'')+'" data-act="ap-sort" data-k="'+k+'">'+t+(on?(d?' ▾':' ▴'):'')+'</button></th>'}
function recNum(f){var m={};fincaRecs(f.id).forEach(function(r,i){m[r.ref]=i+1});return m}
function finHTML(f,ctx){
  var t=(f.datos||{}).tit||{}, rs0=fincaRecs(f.id), rs=recOrden(rs0), N=recNum(f), nsel=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id}).length, p=ctx.p;
  var h=cabeza(1,'Tu finca','Mira que el cultivo y el riego de cada recinto estén bien: con eso se forman solas las unidades del plan.');
  var nm=nifMal(t.nif);
  h+='<section class="blk" id="titular"><div class="g3">'+fld('Nombre de la finca',inp('data-e="nombre"',f.nombre))+fld('Titular (nombre o razón social)',inp('data-e="tit.n"',t.n,'Falta','',t.n?'':'e'))+
    fld('NIF / CIF',inp('data-e="tit.nif"',t.nif,'Falta','',(t.nif?'':'e')+(nm?' rare':''))+(nm?'<em class="nifw">'+ic('alert')+esc(nm)+'</em>':''))+'</div>';
  var mas=['dir','cp','mun','prov','rn','rnif'].filter(function(k){return t[k]}).length;
  h+='<details class="mas" data-o="tit"'+(UI.titOpen?' open':'')+'><summary>'+ic('next')+'<b>Más datos del titular</b><span>'+(mas?esc([t.dir,t.mun,t.prov,t.rn?'representante: '+t.rn:''].filter(Boolean).join(' · ')):'Dirección, municipio, provincia y representante')+'</span></summary><div class="g3">'+
    [['dir','Dirección fiscal'],['cp','Código postal'],['mun','Municipio'],['prov','Provincia'],['rn','Representante (si firma otro)'],['rnif','NIF del representante']].map(function(x){var w=x[0]==='rnif'?nifMal(t.rnif):'';return fld(x[1],inp('data-e="tit.'+x[0]+'"',t[x[0]],'','',w?'rare':'')+(w?'<em class="nifw">'+ic('alert')+esc(w)+'</em>':''))}).join('')+'</div></details></section>';
  h+='<div class="rh"><h2>Recintos <span>'+nRec(rs.length)+' · '+esc(haTxt(rs))+'</span></h2><span class="sp"></span><button class="btn p" data-act="ap-addrec">'+ic('plus')+'Añadir recintos</button><button class="btn" data-act="ap-mapa">'+ic('map')+'Mapa SIGPAC</button>'+(rs.length?'<button class="btn q" data-act="ap-sig" title="Superficie, pendiente y uso del SIGPAC">'+ic('refresh')+'Leer el SIGPAC</button>':'')+'</div>';
  if(!rs.length)return h+empiezaRec()+finPie(f);
  if(nsel)h+='<div class="bulk"><b>'+nsel+(nsel===1?' seleccionado':' seleccionados')+'</b><label>Cultivo <select class="ap-bulk" data-k="cultivo"><option value="">—</option>'+cultOpts('').replace('<option value="">Sin cultivo</option>','')+'</select></label><label>Riego <select class="ap-bulk" data-k="riego"><option value="">—</option><option value="secano">Secano</option><option value="regadio">Regadío</option></select></label><button class="btn q danger" data-act="ap-bulk-del">'+ic('trash')+'Quitar de la finca</button><button class="btn q" data-act="ap-bulk-none">Quitar la selección</button></div>';
  h+='<table class="rt"><thead><tr><th class="c"><input type="checkbox" class="rsel-all" aria-label="Seleccionar todos"'+(nsel&&nsel===rs.length?' checked':'')+'></th>'+thS('ref','Recinto SIGPAC')+thS('ha','ha','r')+thS('cultivo','Cultivo')+thS('riego','Riego')+'<th></th></tr></thead><tbody>'+
    rs.map(function(r){var k=esc(r.ref),op=!!RMR[r.ref],dif=r.hs!=null&&Math.abs(r.hs-(r.ha||0))>0.00005;
      var x='<tr class="'+(SEL[r.ref]?'sel':'')+'" data-ref="'+k+'"><td class="c"><input type="checkbox" class="rsel" data-ref="'+k+'"'+(SEL[r.ref]?' checked':'')+' aria-label="Seleccionar"></td>'+
        '<td><div class="rref">'+glyph(r.cultivo,'sm')+'<div><b><i>'+N[r.ref]+'</i>'+esc(refLabel(r.ref))+'</b><span>'+esc([r.nombre,r.variedad,r.uso?'uso '+r.uso:'',r.pend!=null?'pendiente '+trTxt(r.pend,1)+' %':'',+r.hc>0?ha2(r.hc)+' ha cultivadas':''].filter(Boolean).join(' · '))+'</span></div></div></td>'+
        '<td class="r">'+inp('data-r="'+k+'" data-k="ha" data-num="1"',r.ha,'','decimal','n w6')+(dif?'<small class="sg2">SIGPAC '+ha2(r.hs)+'</small>':'')+'</td>'+
        '<td>'+cultSel('data-r="'+k+'" data-k="cultivo"',r.cultivo).replace('class="trsel ap-in"','class="in ap-in'+(r.cultivo?'':' e')+'"')+'</td>'+
        '<td><select class="in ap-in" data-r="'+k+'" data-k="riego"><option value="secano">Secano</option><option value="regadio"'+(r.riego==='regadio'?' selected':'')+'>Regadío</option></select></td>'+
        '<td class="nw"><button type="button" class="btn q ib'+(op?' on':'')+'" data-act="ap-recmas" data-ref="'+k+'" aria-expanded="'+op+'" title="Variedad, paraje y ha cultivadas">más</button><button type="button" class="btn q ib danger" data-act="ap-recdel" data-ref="'+k+'" aria-label="Quitar recinto" title="Quitar de la finca">'+ic('trash')+'</button></td></tr>';
      if(op)x+='<tr class="sub"><td></td><td colspan="5"><div class="g3">'+fld('Variedad',inp('data-r="'+k+'" data-k="variedad"',r.variedad))+fld('Paraje',inp('data-r="'+k+'" data-k="nombre"',r.nombre,'Opcional'))+fld('ha cultivadas (vacío: todo el recinto)',inp('data-r="'+k+'" data-k="hc" data-num="1"',r.hc,ha2(r.ha||0),'decimal','n'))+'</div>'+
        '<div class="small">'+esc([r.uso?'Uso SIGPAC: '+usoTxt(r.uso):'',r.hs!=null?'Superficie SIGPAC: '+ha2(r.hs)+' ha':'',r.pend!=null?'Pendiente media: '+trTxt(r.pend,1)+' %':'Pendiente sin leer del SIGPAC'].filter(Boolean).join(' · '))+'</div></td></tr>';
      return x}).join('')+'</tbody></table>';
  if(p){var Cs=ctx.Cs;h+='<div class="uds"><span>Unidades del plan:</span>'+udsOrden(p,Cs).map(function(i){var u=p.ud[i],C=Cs[i];return '<span class="udc">'+glyph(udCult(f,u),'xs')+'<b>'+esc(u.n)+'</b> '+ha2(C.ha)+' ha'+(C.ex?' · no necesita plan':'')+'</span>'}).join('')+hlp('ud')+'</div>';
    if(UDAV[p.id]||udMan(f,p)&&udDesajuste(f,p))h+='<div class="warn">Las unidades de este plan están cambiadas a mano, así que no las rehago solas. '+(udDesajuste(f,p)?'Hay recintos cuyo cultivo o riego no es el de su unidad. ':'')+'<button class="lnk" data-act="ap-udrehacer">Rehacer las unidades por cultivo y riego</button> <span class="small">(se conservan los datos de las que sigan igual)</span></div>'}
  return h+finPie(f)}
function udDesajuste(f,p){return (p.ud||[]).some(function(u){var ks={};(u.refs||[]).forEach(function(x){var r=S.recintos[x];if(r)ks[udClave(r)]=1});return Object.keys(ks).length>1})}
function finPie(f){return '<div class="foot">'+(fincaRecs(f.id).length?'<button class="btn p big" data-act="ap-next">Siguiente: los datos'+ic('next')+'</button>':'')+'<span class="sp"></span><button class="btn q danger" data-act="ap-fdel">'+ic('trash')+'Borrar la finca</button></div>'}
function empiezaRec(){return '<div class="opts">'+
  '<button class="opt" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Del PDF de la PAC</b><span>Salen todos con su cultivo, superficie y riego.</span></button>'+
  '<button class="opt" data-act="ap-mapa">'+ic('map')+'<b>En el mapa SIGPAC</b><span>Busca tu parcela en la foto aérea y toca los recintos.</span></button>'+
  '<button class="opt" data-act="ap-csv-pick">'+ic('grid')+'<b>De una tabla (CSV)</b><span>Provincia, municipio, polígono, parcela y recinto.</span></button>'+
  '<button class="opt" data-act="ap-addrec">'+ic('plus')+'<b>A mano</b><span>Con el código SIGPAC de cada recinto.</span></button></div>'}
function addRecHoja(){var f=finca(UI.fid);
  sheet('<h2>Añadir recintos</h2>'+(f?'<div class="opts two"><button class="opt" data-act="ap-mapa">'+ic('map')+'<b>En el mapa SIGPAC</b><span>Toca y añade.</span></button><button class="opt" data-act="ap-pdf-pick">'+ic('pdf')+'<b>Del PDF de la PAC</b><span>Todos de una vez.</span></button><button class="opt" data-act="ap-csv-pick">'+ic('grid')+'<b>De una tabla (CSV)</b><span>Polígono, parcela…</span></button></div><h3>A mano</h3>':'<p class="small">Escribe el código SIGPAC de cada recinto. La finca se crea al añadir el primero.</p>'+fld('Nombre de la finca','<input id="ap-fn" class="in" value="Mi finca" autocomplete="off">'))+
    fld('Código SIGPAC','<input id="ap-ref" class="in" inputmode="numeric" placeholder="02:081:0:0:1:1:1" autocomplete="off"'+(f?' autofocus':'')+'>')+'<div class="small">Provincia, municipio, agregado, zona, polígono, parcela y recinto (o solo provincia, municipio, polígono, parcela y recinto).</div>'+
    '<div class="g3">'+fld('Cultivo','<select id="ap-rc" class="in">'+cultOpts(UI.mLast?UI.mLast.c:'')+'</select>')+fld('Riego','<select id="ap-rr" class="in"><option value="secano">Secano</option><option value="regadio"'+(UI.mLast&&UI.mLast.r==='regadio'?' selected':'')+'>Regadío</option></select>')+fld('ha (vacío: las del SIGPAC)','<input id="ap-rh" class="in" inputmode="decimal" autocomplete="off">')+'</div>'+
    '<div class="acts"><button class="btn p" data-act="ap-ref-add">'+ic('plus')+'Añadir este recinto</button><button class="btn" data-act="ap-close">'+(f?'Hecho':'Cancelar')+'</button></div>','wide');
  if(!f)setTimeout(function(){var r=$('#ap-ref');if(r)r.focus()},40)}
function refAdd(){var f=finca(UI.fid), ref=refParse(($('#ap-ref')||{}).value);if(!ref){toast('Ese código SIGPAC no está completo.');var r0=$('#ap-ref');if(r0)r0.focus();return}
  var hv=($('#ap-rh')||{}).value, ha=hv?numEs(hv):0; if(hv&&(ha===null||isNaN(ha))){toast('Pon las ha con coma decimal, por ejemplo 2,35.');return}
  if(S.recintos[ref]&&finca(S.recintos[ref].fincaId)&&(!f||S.recintos[ref].fincaId!==f.id)){toast('Ese recinto ya está en otra finca.');return}
  undoPush();
  /* la finca se crea al entrar el primer recinto (punto 31, fallo 9) */
  if(!f){f=fincaNueva(String(($('#ap-fn')||{}).value||'').trim()||'Mi finca');UI.v='finca';UI.fid=f.id;UI.step='fin';saveUI();var fn=$('#ap-fn');if(fn)fn.closest('.fld').remove()}
  var c=($('#ap-rc')||{}).value||'', rg=($('#ap-rr')||{}).value||'secano';
  var R=addRecs(f,[{ref:ref,ha:ha||0,cultivo:c,riego:rg}]);
  if(!R.added){toast('Ese recinto ya está en la finca.');return}
  UI.mLast={c:c,r:rg};saveUI();
  save(); $('#ap-ref').value=''; $('#ap-rh').value=''; $('#ap-ref').focus(); toast('Recinto '+refLabel(ref)+' añadido. Leo el SIGPAC…'); render();
  sigLeer(f,[ref],function(ok){if(!ok)toast('No he podido leer ese recinto en el SIGPAC: revisa el código o pon las ha a mano.')});
}

/* ───── paso 2: Datos (una fila por unidad, solo lo imprescindible) ───── */
function nin(path,val,cls,ph,lbl){return '<input class="in pa-in'+(cls?' '+cls:'')+'" data-p="'+esc(path)+'" inputmode="decimal" value="'+esc(val===undefined||val===null?'':(typeof val==='number'?abEd(val):val))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off"'+(lbl?' aria-label="'+esc(lbl)+'"':'')+'>'}
/* v3.2: nota de la zona vulnerable bajo el desplegable y aviso de unidad con recintos dentro y fuera */
function zvNota(f,u){var z=zvSig(f,u),o={};if(fEj(f)&&!zvMan(u))return o;
  if(zvMan(u))o.t='Cambiado a mano'+(z.v!==null&&z.v!==(u.zv?1:0)?' (el SIGPAC dice «'+(z.v?'Sí':'No')+'»)':'');
  else if(z.pend&&z.v===null)o.t='Comprobando en el SIGPAC…';
  else if(u.zvA===1&&z.v!==null)o.t='Según el SIGPAC (FEGA), consultado el '+isoEs(z.f||u.zvF);
  else if(z.R.length){o.t='No he podido comprobarlo en el SIGPAC';o.w=1}
  if(z.den.length&&z.fue.length){o.mix=z.den.length+' de '+z.R.length+' recintos están en zona vulnerable.';
    var bo=z.den.filter(function(r){return r.zvn.p<99.5});if(bo.length)o.bor=bo.map(function(r){return refLabel(r.ref)+': '+zvPct(r.zvn.p)+' dentro'}).join(' · ')+'.'}
  return o}
function suFuente(p,u){return u.suDe?(p.ud||[]).filter(function(x){return x.id===u.suDe&&x!==u})[0]||null:null}
function datHTML(f,ctx){var p=ctx.p,Cs=ctx.Cs;
  var h=cabeza(2,'Lo que hace falta para el plan','Solo esto. Lo demás ya está puesto con lo habitual y lo puedes revisar después.');
  if(!p)return h+'<div class="note">Primero mete los recintos en «Finca».</div>';
  var ord=udsOrden(p,Cs), conSu=ord.filter(function(i){return !Cs[i].ex&&suTiene(p.ud[i].su)&&!p.ud[i].suDe});
  h+='<table class="q"><thead><tr><th class="grp" colspan="4"></th><th class="grp sepc" colspan="5">Análisis de suelo'+hlp('su')+'</th></tr><tr><th>Unidad</th><th>Zona vulnerable'+hlp('zv')+'</th><th>Cosecha esperada<small>kg/ha</small></th><th>Agua de riego'+hlp('ag')+'<small>m³/ha al año</small></th><th class="sepc">Materia orgánica<small>%</small></th><th>pH</th><th>Fósforo<small>Olsen, mg/kg</small></th><th>Potasio<small>mg/kg</small></th><th>Fecha</th></tr></thead><tbody>';
  ord.forEach(function(i){var u=p.ud[i],C=Cs[i],P='ud.'+i+'.',ne=u.ne||{},su=u.su||{},ag=u.ag||{},cu=udCult(f,u),uc='<div class="u">'+glyph(cu)+'<div><b>'+esc(u.n)+'</b><span>'+ha2(C.ha)+' ha · '+nRec(C.rs.length)+'</span></div></div>';
    if(C.ex){h+='<tr class="ex" data-u="'+i+'"><td>'+uc+'</td><td colspan="8">No necesita plan por ser secano de 10 ha o menos. No hay que rellenar nada.'+hlp('ex')+'</td></tr>';return}
    if(u.d0===1){h+='<tr class="ex" data-u="'+i+'"><td>'+uc+'</td><td colspan="8">Sin abonado esta campaña (programa a dosis 0). <button class="lnk" data-act="ap-det" data-k="ne">Cambiar</button></td></tr>';return}
    var mo=ne.mo||'cos', e=function(v){return abHas(v)?'':' e'};
    var cos=mo==='cos'?nin(P+'ne.pr',ne.pr,'w7'+e(ne.pr),'Falta','Cosecha esperada de '+u.n):'<span class="tx2">'+(mo==='edad'?'En formación ('+(+ne.ed||1)+'º verde)':'Necesidad a mano')+' <button class="lnk" data-act="ap-det" data-k="ne">Cambiar</button></span>';
    var agua=u.rg==='regadio'?nin(P+'ag.m3',ag.m3,'w7'+(abHas(ag.m3)||abHas(ag.max)?'':' e'),'Falta','Agua de riego de '+u.n):'<span class="na">Secano</span>';
    var src=suFuente(p,u), suc;
    if(src){var s2=src.su||{};suc='<td class="sepc" colspan="5"><div class="lnkd">'+ic('copy')+'<span>El mismo análisis que «'+esc(src.n)+'»: '+esc([abHas(s2.mo)?trTxt(+s2.mo,2)+' % MO':'',abHas(s2.ph)?'pH '+trTxt(+s2.ph,1):'',abHas(s2.p)?'P '+trTxt(+s2.p,0):'',abHas(s2.k)?'K '+trTxt(+s2.k,0):'',pabIso(s2.f)?isoEs(pabIso(s2.f)):''].filter(Boolean).join(' · '))+'</span></div></td>'}
    else if(suNoMot(u,C)){var mot=suNoMot(u,C);suc='<td class="sepc" colspan="5"><div class="suno"><b>Sin análisis de suelo</b><select class="in ap-sunomot" data-u="'+i+'" aria-label="Motivo" title="'+esc(mot==='no20'?'No obligatorio: unidad de 20 ha o menos sin hoja de cultivo de 5 ha':'No dispongo de análisis')+'">'+
        '<option value="nd">No dispongo de análisis</option>'+(C.ha<=20+1e-9?'<option value="no20"'+(mot==='no20'?' selected':'')+'>No obligatorio: unidad de 20 ha o menos sin hoja de cultivo de 5 ha</option>':'')+'</select></div></td>'}
    else{var any=suTiene(su), ee=function(v){return any||abHas(v)?'':' e'};
      suc='<td class="sepc">'+nin(P+'su.mo',su.mo,'w5'+ee(su.mo),any?'':'Falta','Materia orgánica')+'</td><td>'+nin(P+'su.ph',su.ph,'w4'+ee(su.ph),'','pH')+'</td><td>'+nin(P+'su.p',su.p,'w4'+ee(su.p),'','Fósforo Olsen')+'</td><td>'+nin(P+'su.k',su.k,'w5'+ee(su.k),'','Potasio')+'</td>'+
        '<td><input class="in pa-in w8'+(pabIso(su.f)?'':' soft')+'" data-p="'+P+'su.f" data-date="1" value="'+esc(pabIso(su.f)?isoEs(pabIso(su.f)):'')+'" placeholder="dd/mm/aaaa" inputmode="numeric" autocomplete="off" aria-label="Fecha del análisis"></td>'}
    var zv=zvNota(f,u);
    h+='<tr data-u="'+i+'"><td>'+uc+'</td><td><select class="in pa-in sel" data-p="'+P+'zv" aria-label="Zona vulnerable"><option value="0">No</option><option value="1"'+(u.zv?' selected':'')+'>Sí</option></select></td><td>'+cos+'</td><td>'+agua+'</td>'+suc+'</tr>';
    if(noExtr(u)||(!PAB_EX[u.ck]&&mo==='cos')){var ex=ne.ex||{};h+='<tr class="sub ex2" data-u="'+i+'"><td></td><td colspan="8"><div class="exw'+(noExtr(u)?' on':'')+'">'+ic('alert')+'<span><b>No tengo extracciones de referencia para este cultivo: escribe las tuyas.</b> '+esc(SIN_EX_TXT)+'.'+hlp('sinex')+'</span><span class="exi">'+['n','p','k'].map(function(k){return '<label>'+{n:'N',p:'P₂O₅',k:'K₂O'}[k]+' '+nin(P+'ne.ex.'+k,ex[k],'w4'+(noExtr(u)?' e':''),'','Extracción '+k)+'</label>'}).join('')+'<small>kg por t de cosecha</small></span></div></td></tr>'}
    var otras=conSu.filter(function(j){return j!==i});
    /* v3.2: «No tengo análisis de suelo», a la vista en la fila del análisis */
    var sno=!src&&!suTiene(su)?'<label class="ckr suck"><input type="checkbox" class="ap-suno" data-u="'+i+'"'+(u.suNo?' checked':'')+'> No tengo análisis de suelo</label>':'';
    var slk=src||otras.length?'<select class="in sulk ap-sulink'+(src?' on':'')+'" data-u="'+i+'" aria-label="¿Mismo análisis de suelo que otra unidad?"><option value="">'+(src?'Escribir su propio análisis':'Mismo suelo que otra unidad…')+'</option>'+otras.concat(src?[p.ud.indexOf(src)]:[]).filter(function(v,ix,a){return a.indexOf(v)===ix}).map(function(j){return '<option value="'+esc(p.ud[j].id)+'"'+(src&&src.id===p.ud[j].id?' selected':'')+'>El mismo análisis que «'+esc(p.ud[j].n)+'»</option>'}).join('')+'</select>':'';
    /* la nota de la zona vulnerable va debajo de su desplegable, en la fila de abajo (no ensancha la tabla) */
    var zvt=zv.t?'<small class="zvn'+(zv.w?' w':'')+'">'+esc(zv.t)+'</small>':'';
    if(sno||slk||zvt)h+='<tr class="sub sul" data-u="'+i+'"><td></td><td colspan="3">'+zvt+'</td><td class="sepc" colspan="5">'+(sno||slk?'<div class="sul2">'+sno+slk+'</div>':'')+'</td></tr>';
    if(zv.mix)h+='<tr class="sub zvm" data-u="'+i+'"><td></td><td colspan="8"><div class="zvw">'+ic('alert')+'<span><b>'+esc(zv.mix)+'</b>'+(zv.bor?' '+esc(zv.bor):'')+'</span><button class="btn q" data-act="ap-zvsep" data-u="'+i+'">Separar en dos unidades</button></div></td></tr>';});
  h+='</tbody></table>';
  /* revisión: «Hacer el plan» solo se bloquea por lo que falta en esta tabla; lo que falta de Finca se dice al lado y lleva allí */
  var Fd=ctx.F.filter(function(x){return x.s==='dat'}), Ff=ctx.F.filter(function(x){return x.s==='fin'}), bloq=Fd.length>0;
  h+='<div class="foot"><button class="btn p big'+(bloq?' off':'')+'" data-act="ap-hacer"'+(bloq?' aria-disabled="true"':'')+'>Hacer el plan'+ic('next')+'</button>'+(bloq?'<button class="hint" data-act="ap-falta" data-s="dat">'+ic('alert')+'Falta'+(Fd.length===1?' 1 dato: ':'n '+Fd.length+' datos: ')+esc(Fd[0].t)+(Fd.length>1?' y '+(Fd.length-1)+' más':'')+'</button>':'<span class="hint ok">'+ic('check')+(Ff.length?'La tabla está completa':'Están todos los datos')+'</span>')+(Ff.length?'<button class="hint hfin" data-act="ap-falta" data-s="fin">'+ic('alert')+'Falta '+esc(Ff[0].t)+(Ff.length>1?' y '+(Ff.length-1)+' más':'')+' (paso Finca)</button>':'')+'</div>';
  h+='<details class="more" data-o="datmas"'+(UI.datMas?' open':'')+'><summary>'+ic('next')+'<b>Más detalle</b><span>Lluvia de la zona, cosecha y extracciones, aportes de estiércol, plantación joven, recinto representativo… ya rellenos con lo habitual.</span></summary>'+detHTML(f,ctx,['cl','ne','ap','su','ag','ud'])+'</details>';
  return h}

/* ───── paso 3: Plan y PDF ───── */
var CAMP=[11,12,1,2,3,4,5,6,7,8,9,10], MESL='NDEFMAMJJASO';
var FO=['Suelo','Localizado','Fertirrigación','Foliar','Enterrado'];
function foTxt(fo){return {'Fertirrigación':'por el riego','Suelo':'al suelo','Localizado':'localizado','Foliar':'foliar','Enterrado':'enterrado'}[fo]||(fo?String(fo).toLowerCase():'')}
function mesTxt(mo){var ms=mesesDe(mo);if(!ms.length)return mo||'';var ix=ms.map(function(m){return CAMP.indexOf(m)}).sort(function(a,b){return a-b}),seg=true;for(var i=1;i<ix.length;i++)if(ix[i]!==ix[i-1]+1)seg=false;
  if(!seg)return mo;var a=MESES[CAMP[ix[0]]-1],b=MESES[CAMP[ix[ix.length-1]]-1];if(ix.length===1)return a;if(ix.length===2)return a+' y '+b.toLowerCase();return 'De '+a.toLowerCase()+' a '+b.toLowerCase()}
function calCab(){return '<div class="cal">'+MESL.split('').map(function(x){return '<i>'+x+'</i>'}).join('')+'</div>'}
function calFila(q,o,pm){var ms=mesesDe(q.mo),ix=ms.map(function(m){return CAMP.indexOf(m)}).sort(function(a,b){return a-b}),uf=o?o.uf:{n:0,p:0,k:0},mx=Math.max(uf.n,uf.p,uf.k),nu=mx<=0.05?'z':(uf.n===mx?'n':(uf.p===mx?'p':'k'));
  var h='<div class="cal" title="'+esc((q.mo||'Sin mes')+(pm.length?' · sin N en '+pm.map(function(m){return MESES[m-1].toLowerCase()}).join(', ')+' (anexo II)':''))+'">';
  CAMP.forEach(function(m){h+='<u'+(pm.indexOf(m)>=0?' class="pm"':'')+'></u>'});
  var runs=[];ix.forEach(function(x){var r=runs[runs.length-1];if(r&&x===r[1]+1)r[1]=x;else runs.push([x,x])});
  var bad=uf.n>0.05&&ms.some(function(m){return pm.indexOf(m)>=0});
  runs.forEach(function(r){h+='<em class="'+nu+(bad?' bad':'')+'" style="grid-column:'+(r[0]+1)+'/'+(r[1]+2)+'"></em>'});
  return h+'</div>'}
function uf0(v){return v>0.05?nf(v,0):'–'}
function frase(u,C,k){var b=C.B[k],ne=u.ne||{},d=[];
  if(b.s>0.05)d.push('el suelo aporta '+nf(b.s,0));if(b.a>0.05)d.push('el riego trae '+nf(b.a,0));if(b.o>0.05)d.push('el estiércol aporta '+nf(b.o,0));if(b.x>0.05)d.push('otros aportes, '+nf(b.x,0));
  if(d.length)return 'Necesita '+nf(b.nb,0)+'; '+d.join(', ')+'.';
  if(ne.mo==='edad')return 'Es lo que necesita la plantación joven.';if(ne.mo==='man')return 'Es la necesidad puesta a mano.';
  return b.nb>0.05?'Es lo que se lleva la cosecha.':'La cosecha no se lleva nada.'}
function balBarras(f,p,u,C){var ex=noExtr(u),lim=u.zv&&C.L.ok?C.L.lim:0;
  return '<div class="under">'+[['n','Nitrógeno'],['p','Fósforo'],['k','Potasio']].map(function(x){var k=x[0],b=C.B[k],mx=Math.max(b.net,b.pl,k==='n'?lim:0,1)*1.25,w=function(v){return Math.min(100,v/mx*100).toFixed(1)+'%'};
    var dv=b.dv, st, cls;
    var ne=u.ne||{},sinCos=(ne.mo||'cos')==='cos'&&!(+ne.pr>0)&&u.d0!==1;
    /* fallo 4: sin cosecha o sin extracciones no puede salir «cuadra» con 0 de 0 */
    if(ex){st='<b class="w">Faltan las extracciones.</b> No se puede saber lo que necesita.';cls='w'}
    else if(sinCos){st='<b class="w">Falta la cosecha esperada.</b> Sin ella no se sabe lo que necesita.';cls='w'}
    else if(b.st==='Ajustado'){st='<b>✓ Cuadra.</b> '+frase(u,C,k);cls='ok'}
    else{var big=Math.abs(dv)>Math.max(10,b.net*0.25)||(k==='n'&&!C.lok);cls=big?'r':'w';st='<b class="'+cls+'">'+(dv<0?'Faltan '+nf(-dv,0):'Sobran '+nf(dv,0))+'.</b> '+frase(u,C,k)}
    return '<div class="mt '+k+'"><div class="t"><b>'+x[1]+'</b><span class="sp"></span><strong>'+nf(b.pl,0)+'</strong> de '+nf(b.net,0)+'</div><div class="bar"><i style="width:'+w(b.pl)+'"></i><s style="left:'+w(b.net)+'"></s>'+(k==='n'&&lim?'<s class="lim" style="left:'+w(lim)+'" title="Límite de N en zona vulnerable: '+nf(lim,0)+'"></s>':'')+'</div><small class="'+cls+'">'+st+(k==='n'&&lim?' Límite en zona vulnerable: '+nf(lim,0)+hlp('lim')+'.':'')+'</small></div>'}).join('')+'</div>'}
function progTabla(f,p,u,C,i){var P='ud.'+i+'.',rows=(u.pg||{}).rows||[],pm=perMeses(u),co=!!UI.cost;
  var h='<table class="pg"><thead><tr><th>Abono</th><th class="r">kg/ha</th><th class="r">Para la unidad</th><th class="r tn">N</th><th class="r tp">P₂O₅</th><th class="r tk">K₂O</th><th class="calh">'+calCab()+'</th></tr></thead><tbody>';
  rows.forEach(function(q,j){var R=P+'pg.rows.'+j+'.',o=C.rows[j]||{uf:{n:0,p:0,k:0}},op=!!RM[q.id],sub=[mesTxt(q.mo),foTxt(q.fo)].filter(Boolean).join(' · ');
    h+='<tr data-j="'+j+'" class="'+(op?'on':'')+'"><td class="ab"><button type="button" class="abn" data-act="ap-rowed" data-id="'+esc(q.id)+'" aria-expanded="'+op+'" title="Cambiar abono, mes, riqueza o forma"><b>'+(q.n?esc(q.n):'<em>Elige el abono</em>')+'</b><span>'+esc(sub||'Sin mes')+(q.org===1?' · orgánico':'')+'</span></button></td>'+
      '<td class="r"><input class="in box pa-in n" data-p="'+esc(R+'d')+'" inputmode="decimal" value="'+esc(abHas(q.d)?abEd(q.d):'')+'" autocomplete="off" aria-label="Dosis en kg/ha de '+esc(q.n||'abono')+'"></td>'+
      '<td class="r nw">'+(+q.d>0?nf(q.d*C.ha,0)+' kg':'–')+'</td><td class="r tn'+(o.uf.n>0.05?'':' z')+'">'+uf0(o.uf.n)+'</td><td class="r tp'+(o.uf.p>0.05?'':' z')+'">'+uf0(o.uf.p)+'</td><td class="r tk'+(o.uf.k>0.05?'':' z')+'">'+uf0(o.uf.k)+'</td><td>'+calFila(q,o,pm)+'</td></tr>';
    if(op){var cin=function(k,v,ph,cls,lbl){return '<input class="in pa-in'+(cls?' '+cls:'')+'" data-p="'+esc(R+k)+'"'+(/\bn\b/.test(cls||'')?' inputmode="decimal"':'')+' value="'+esc(v===undefined||v===null?'':(typeof v==='number'?abEd(v):v))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off" aria-label="'+esc(lbl)+'">'};
      var sg=function(k,v,kind,ph,lbl){return '<div class="sgw">'+cin(k,v,ph,'sgi',lbl).replace('class="in pa-in sgi"','class="in pa-in sgi" data-sg="'+kind+'"')+'</div>'};
      var L=FO.slice();if(q.fo&&L.indexOf(q.fo)<0)L.unshift(q.fo);
      var tipo=abModo(q)==='tipo',msel='<select class="in ap-abmodo" data-id="'+esc(q.id)+'" aria-label="Abono tipo o mi abono"><option value="tipo"'+(tipo?' selected':'')+'>Abono tipo (del listado)</option><option value="mio"'+(tipo?'':' selected')+'>Mi abono (nombre comercial)</option></select>';
      h+='<tr class="sub ed"><td colspan="7"><div class="edg">'+fld('Abono'+hlp('abt'),msel,'abm')+fld(tipo?'Abono tipo':'Nombre comercial',sg('n',q.n,tipo?'abt':'abm',tipo?'Buscar en los abonos tipo…':'Nombre de la etiqueta',tipo?'Abono tipo':'Nombre comercial'),'w2')+fld('Mes o meses',sg('mo',q.mo,'mes','Marzo-abril','Mes'))+fld('N %',cin('rn',q.rn,'0','n','N %'),'nn')+fld('P₂O₅ %',cin('rp',q.rp,'0','n','P₂O₅ %'),'nn')+fld('K₂O %',cin('rk',q.rk,'0','n','K₂O %'),'nn')+
        fld('Forma','<select class="in pa-in" data-p="'+esc(R+'fo')+'" aria-label="Forma de aplicación">'+L.map(function(x){return '<option'+(x===(q.fo||'')?' selected':'')+'>'+esc(x)+'</option>'}).join('')+'</select>','fo')+fld('Equipo',cin('eq',q.eq,'Abonadora','','Equipo'),'eq')+fld('Medida de mitigación',cin('mi',q.mi,'Opcional','','Medida de mitigación'),'w2')+
        (co?fld('Precio €/kg',cin('pz',q.pz,'0,00','n','Precio €/kg')):'')+'<label class="ckr"><input type="checkbox" class="pa-chk" data-p="'+esc(R+'org')+'"'+(q.org===1?' checked':'')+'> Abono orgánico</label></div>'+
        '<div class="acts"><button type="button" class="btn q danger" data-act="pa-row-del" data-u="'+i+'" data-j="'+j+'">'+ic('trash')+'Quitar este abono</button><span class="sp"></span><span class="small">'+(tipo?'La riqueza es la típica de ese tipo de abono: pon la de tu etiqueta si es otra.':'Pon la riqueza de la etiqueta. Al pulsar «Hecho» queda en «Mis abonos».')+'</span><button type="button" class="btn" data-act="ap-rowed" data-id="'+esc(q.id)+'">Hecho</button></div></td></tr>'}});
  if(!rows.length)h+='<tr><td colspan="7" class="empty">Aún no hay abonos. Pulsa «Proponer programa» o «Añadir abono».</td></tr>';
  var B=C.B;
  h+='<tr class="tot" id="u'+i+'-tot"><td colspan="3">Total del plan <span class="z">U.F./ha</span></td><td class="r">'+nf(B.n.pl,0)+'</td><td class="r">'+nf(B.p.pl,0)+'</td><td class="r">'+nf(B.k.pl,0)+'</td><td></td></tr>';
  h+='<tr class="nec"><td colspan="3">Lo que hay que poner'+hlp('net')+'</td><td class="r">'+nf(B.n.net,0)+'</td><td class="r">'+nf(B.p.net,0)+'</td><td class="r">'+nf(B.k.net,0)+'</td><td></td></tr></tbody></table>';
  var ph=pabPerHits(f,u,p,C);if(ph.length)h+='<div class="note bad">'+ic('alert')+'Lleva nitrógeno en meses en los que no se puede (anexo II): '+esc(ph.join(' · '))+hlp('per')+'</div>';
  return h}
function udSec(f,p,u,C,i,E){var ne=u.ne||{},cu=udCult(f,u),rows=(u.pg||{}).rows||[];
  var meta=[ha2(C.ha)+' ha',nRec(C.rs.length)];
  if(C.ex)meta.push('secano de 10 ha o menos');else if((ne.mo||'cos')==='cos'&&+ne.pr>0)meta.push('cosecha esperada '+nf(+ne.pr,0)+' kg/ha');else if(ne.mo==='edad')meta.push('plantación en formación ('+(+ne.ed||1)+'º verde)');
  if(u.zv)meta.push('zona vulnerable');
  var h='<section class="ud'+(C.ex?' x':'')+'" id="u'+i+'" data-u="'+i+'"><div class="udh">'+glyph(cu)+'<div><h2>'+esc(u.n)+'</h2><div class="meta">'+esc(meta.join(' · '))+'</div></div><span class="sp"></span>'+chipEst(E)+(C.ex?hlp('ex'):'')+'</div>';
  if(C.ex&&!rows.length&&!UI['exo'+u.id])return h+'<div class="exn">No hay que rellenar nada. <button class="lnk" data-act="ap-exopen" data-u="'+i+'">Hacer su plan igualmente</button></div></section>';
  if(u.d0===1)return h+'<div class="note">Sin abonado esta campaña (programa a dosis 0): no hacen falta abonos, análisis, medidas ni restricciones. <button class="lnk" data-act="ap-det" data-k="ne">Cambiar</button></div></section>';
  h+=progTabla(f,p,u,C,i)+balBarras(f,p,u,C);
  h+='<div class="acts"><button class="btn" data-act="pa-row-add" data-u="'+i+'">'+ic('plus')+'Añadir abono</button><button class="btn" data-act="ap-ajustar" data-u="'+i+'"'+(rows.length?'':' disabled')+'>'+ic('scale')+'Ajustar a lo que falta</button><button class="btn" data-act="ap-prop" data-u="'+i+'">'+ic('refresh')+(rows.length?'Proponer otra vez':'Proponer programa')+'</button><span class="sp"></span><button class="btn q" data-act="ap-det" data-k="compra">'+ic('cart')+'Lista de la compra</button></div>';
  if(UI.cost&&rows.length){var Cst=costeUd(u,C);h+='<div class="small cost">'+ic('euro')+'Coste: <b>'+nf(Cst.ha,2)+' €/ha</b> · <b>'+nf(Cst.tot,0)+' €</b> en las '+ha2(C.ha)+' ha'+(Cst.sin?' (hay abonos sin precio)':'')+'. Solo en pantalla.</div>'}
  return h+'</section>'}
function planHTML(f,ctx){var p=ctx.p,Cs=ctx.Cs;
  var h=cabeza(3,'El plan de esta campaña','Propuesto con abonos corrientes. Cambia lo que quieras y las cuentas se rehacen solas.');
  if(!p)return h+'<div class="note">Primero mete los recintos en «Finca».</div>';
  var Fd=ctx.F.filter(function(x){return x.s==='fin'||x.s==='dat'});
  if(Fd.length)h+='<div class="warn">'+ic('alert')+'Falta'+(Fd.length===1?' 1 dato':'n '+Fd.length+' datos')+(Fd[0].s==='fin'?': ':' para el plan: ')+esc(Fd[0].t)+(Fd.length>1?' y '+(Fd.length-1)+' más':'')+(Fd[0].s==='fin'?' (paso Finca)':'')+'. <button class="lnk" data-act="ap-falta" data-s="'+Fd[0].s+'">Ir a ponerlo</button></div>';
  udsOrden(p,Cs).forEach(function(i){h+=udSec(f,p,p.ud[i],Cs[i],i,ctx.E[i])});
  if(!(p.ud||[]).length)h+='<div class="note">El plan no tiene unidades: añade recintos en «Finca».</div>';
  h+='<section class="det" id="detalle"><h3>Detalle del plan <span class="chip m">Ya relleno con lo habitual</span></h3>'+detHTML(f,ctx,null)+'</section>';
  return h}

/* ───── v3.1 punto 7: medidas y restricciones por unidad. «Aplicar a»: todas las unidades (como hasta ahora) o una ───── */
function aplSel(p,U){var id=UI.aplU;if(!id||U.length<2)return U;var k=U.filter(function(i){return p.ud[i].id===id});return k.length?k:U}
function aplHTML(p,UA,U,k){if(UA.length<2)return '';var una=U.length===1,dis=!udsIguales(p,UA,k);
  return '<div class="apl">'+fld('Aplicar a','<select class="in ap-aplu" aria-label="Aplicar a">'+'<option value="">Todas las unidades</option>'+UA.map(function(i){return '<option value="'+esc(p.ud[i].id)+'"'+(una&&U[0]===i?' selected':'')+'>'+esc(p.ud[i].n)+'</option>'}).join('')+'</select>')+
    (dis&&!una?'<div class="small dsu">'+ic('alert')+'Distinto según la unidad: elige una en «Aplicar a» para ver las suyas.</div>':'')+'</div>'}
/* ¿todas las unidades llevan las mismas medidas (o restricciones) marcadas, con el mismo texto? */
function udsIguales(p,U,k){var L=k==='mit'?PAB_MIT:PAB_RES,sig=function(u){var o=u[k]||{};return L.map(function(x){var m=o[x[0]]||{};return x[0]+':'+(m.a===1?1:(m.a===0?0:''))+':'+(m.a===1||m.a===0?(m.d===undefined?'':m.d):'')}).join('|')};
  if(U.length<2)return true;var s0=sig(p.ud[U[0]]);return U.every(function(i){return sig(p.ud[i])===s0})}
/* ───── detalle del plan: cada cosa en una línea con su «Cambiar» (puntos 18, 22, 24, 26, 30) ───── */
function udsNoEx(p,Cs){return (p.ud||[]).map(function(u,i){return i}).filter(function(i){return !Cs[i].ex})}
function detLines(f,ctx){var p=ctx.p,Cs=ctx.Cs,U=udsNoEx(p,Cs),ud=p.ud||[],L=[],t=(f.datos||{}).tit||{};
  var mitOn=PAB_MIT.filter(function(x){return U.some(function(i){return ((ud[i].mit||{})[x[0]]||{}).a===1})}).length, mitDis=!udsIguales(p,U,'mit');
  L.push({k:'mit',l:'Medidas del anexo V',v:mitDis?'Distinto según la unidad':(mitOn?(mitOn===1?'1 marcada':mitOn+' marcadas'):'Ninguna'),w:!mitOn&&U.length});
  var resAll=U.every(function(i){return PAB_RES.every(function(x){return ((ud[i].res||{})[x[0]]||{}).a===1})}), resProp=U.some(function(i){return ud[i].resProp}), resDis=!udsIguales(p,U,'res');
  L.push({k:'res',l:'Restricciones de aplicación',v:resDis?'Distinto según la unidad':(resAll?(resProp?'Las habituales':'Todas comprobadas'):'Hay alguna sin marcar'),w:!resAll});
  var mos={};U.forEach(function(i){mos[(ud[i].mo||{}).obj||'']=1});var mk=Object.keys(mos);
  L.push({k:'mo',l:'Materia orgánica',v:mk.length>1?'Distinta por unidad':(mk[0]==='aumentar'?'Aumentar':(mk[0]==='mantener'?'Mantener':'Sin elegir')),w:mos['']});
  L.push({k:'fi',l:'Fecha y firma',v:pabIso(p.fe)?diaFmt(pabIso(p.fe)):'Sin fecha',w:ctx.F.some(function(x){return x.det==='fi'})});
  var su0=U.map(function(i){return ud[i].su||{}}).filter(function(s){return pabIso(s.f)})[0];
  L.push({k:'su',l:'Análisis de suelo',v:su0?diaFmt(pabIso(su0.f)):(U.some(function(i){return suTiene(ud[i].su)})?'Sin fecha':'Sin análisis'),w:ctx.F.some(function(x){return x.det==='su'})});
  var rgU=U.filter(function(i){return ud[i].rg==='regadio'});
  if(rgU.length){var m3=rgU.map(function(i){return (ud[i].ag||{}).m3}).filter(abHas);L.push({k:'ag',l:'Agua de riego',v:m3.length?nf(+m3[0],0)+' m³/ha'+(m3.length>1&&m3.some(function(x){return +x!==+m3[0]})?' y otras':''):'Sin poner'})}
  var cl=p.cl||PAB_CL;L.push({k:'cl',l:'Lluvia de la zona',v:abHas(cl.a)?nf(+cl.a,0)+' mm al año':'Sin poner',w:!abHas(cl.a)});
  var pz=0,nr=0;U.forEach(function(i){((ud[i].pg||{}).rows||[]).forEach(function(q){nr++;if(abHas(q.pz))pz++})});
  L.push({k:'pz',l:'Precios de los abonos',v:!nr?'Sin abonos':(pz?(pz===nr?'Puestos':pz+' de '+nr):'Sin poner')});
  var mods={};U.forEach(function(i){var n=ud[i].ne||{};mods[(n.mo||'cos')]=1});
  L.push({k:'ne',l:'Cosecha, extracciones y precedente',v:mods.edad?'Plantación joven':(mods.man?'Necesidad a mano':'Por cosecha esperada'),w:ctx.F.some(function(x){return x.det==='ne'})||U.some(function(i){return noExtr(ud[i])})});
  var apN=U.filter(function(i){var a=ud[i].ap||{};return ['s','o','x'].some(function(y){return ['n','p','k'].some(function(k){return +((a[y]||{})[k])>0})})}).length;
  L.push({k:'ap',l:'Aportes de estiércol, suelo y otros',v:apN?(apN===1?'En 1 unidad':'En '+apN+' unidades'):'Ninguno'});
  L.push({k:'ud',l:'Unidades y recintos',v:(ud.length===1?'1 unidad':ud.length+' unidades')+(udMan(f,p)?' (a mano)':''),w:ctx.F.some(function(x){return x.det==='ud'})});
  var cp=compra(f,p,Cs);L.push({k:'compra',l:'Lista de la compra',v:cp.length?(cp.length===1?'1 abono':cp.length+' abonos'):'Vacía'});
  L.push({k:'md',l:'Cambios durante la campaña',v:(p.mods||[]).length?String(p.mods.length):'Ninguno'});
  return L}
function detHTML(f,ctx,set){if(!ctx.p)return '';var L=detLines(f,ctx).filter(function(x){return !set||set.indexOf(x.k)>=0});
  var h='<div class="dl">'+L.map(function(x){var on=UI.det===x.k;return '<div class="'+(on?'on':'')+(x.w?' w':'')+'"><span>'+esc(x.l)+'</span><b>'+(x.w?ic('alert'):'')+esc(x.v)+'</b><button class="lnk" data-act="ap-det" data-k="'+x.k+'" aria-expanded="'+on+'">'+(on?'Cerrar':'Cambiar')+'</button></div>'}).join('')+'</div>';
  if(UI.det&&L.some(function(x){return x.k===UI.det}))h+='<div class="dped" id="det-'+UI.det+'">'+detEd(f,ctx,UI.det)+'<div class="acts"><button class="btn" data-act="ap-det" data-k="'+UI.det+'">Hecho</button></div></div>';
  if(!set){var y1=+ctx.p.y+1,ex=(f.pab||[]).some(function(x){return +x.y===y1});h+='<div class="acts dacts"><button class="btn q" data-act="ap-sigcamp">'+ic('copy')+(ex?'Ir al plan de '+pabCamp(y1):'Hacer el plan de '+pabCamp(y1))+'</button><button class="btn q" data-act="ap-copiaf">'+ic('save')+'Copia de esta finca</button><span class="sp"></span><button class="btn q danger" data-act="pa-del">'+ic('trash')+'Borrar el plan de esta campaña</button></div>'}
  return h}
function udTit(f,u,C){return '<h4>'+glyph(udCult(f,u),'xs')+esc(u.n)+' <span>'+ha2(C.ha)+' ha</span></h4>'}
function chipsP(path,cur,L){return pabChips(path,cur,L)}
function detEd(f,ctx,k){var p=ctx.p,Cs=ctx.Cs,U=udsNoEx(p,Cs),ud=p.ud||[],h='';
  var each=function(fn,all){(all?(p.ud||[]).map(function(u,i){return i}):U).forEach(function(i){h+='<div class="dpu">'+udTit(f,ud[i],Cs[i])+fn(ud[i],Cs[i],i,'ud.'+i+'.')+'</div>'})};
  if(k==='mit'){var UA=U;U=aplSel(p,U);var una=U.length===1&&UA.length>1;
    h+=aplHTML(p,UA,U,'mit')+'<p class="small">'+(una?'Solo para «'+esc(ud[U[0]].n)+'».':(UA.length>1?'Lo que marques aquí se aplica a todas las unidades.':'Se aplican a la unidad del plan.'))+hlp('mit')+'</p>';
    ['B','O','A'].forEach(function(g){h+='<div class="gk">'+esc(PAB_MITG[g])+'</div>';PAB_MIT.filter(function(x){return x[1]===g}).forEach(function(x){var n=U.filter(function(i){return ((ud[i].mit||{})[x[0]]||{}).a===1}).length,ds=U.map(function(i){var m=(ud[i].mit||{})[x[0]]||{};return m.a===1?(m.d===undefined?x[3]:m.d):null}).filter(function(v){return v!==null}),d=ds[0],dd=ds.some(function(v){return v!==d});
      h+='<div class="ckrow'+(n?' on':'')+'"><label class="ckr"><input type="checkbox" class="ap-mitall" data-m="'+x[0]+'"'+(n?' checked':'')+(n&&n<U.length?' data-mix="1"':'')+'> <span>'+esc(x[2])+(n&&n<U.length?' <em class="small">(en '+n+' de '+U.length+' unidades)</em>':'')+'</span></label>'+(n?'<input class="in ap-mitd" data-m="'+x[0]+'" value="'+esc(d||'')+'" placeholder="Cómo se aplica" autocomplete="off">'+(dd?'<em class="small dsu">El texto es distinto según la unidad: si lo cambias aquí, se pone en todas.</em>':''):'')+'</div>'})});
    var urea=U.some(function(i){return Cs[i].rows.some(function(o){return /urea/i.test(o.n)})}),hasB=U.every(function(i){return PAB_MIT.some(function(x){return x[1]==='B'&&((ud[i].mit||{})[x[0]]||{}).a===1})});
    if(urea&&!hasB)h+='<div class="note">El programa lleva urea: hace falta al menos una medida de la parte B.</div>';
    h+='<div class="acts"><button class="btn q" data-act="ap-mitprop">Marcar las que encajan con el programa</button></div>'}
  else if(k==='res'){var UA2=U;U=aplSel(p,U);var una2=U.length===1&&UA2.length>1;
    h+=aplHTML(p,UA2,U,'res')+'<p class="small">'+(una2?'Solo para «'+esc(ud[U[0]].n)+'». ':(UA2.length>1?'Lo que marques aquí se aplica a todas las unidades. ':''))+'Desmarca la que no se cumpla.'+hlp('res')+'</p>';
    PAB_RES.forEach(function(x){var n=U.filter(function(i){return ((ud[i].res||{})[x[0]]||{}).a===1}).length,ds=U.map(function(i){var m=(ud[i].res||{})[x[0]]||{};return m.a!==undefined?(m.d===undefined?x[2]:m.d):null}).filter(function(v){return v!==null}),d=ds[0],dd=ds.some(function(v){return v!==d});
      h+='<div class="ckrow'+(n?' on':'')+'"><label class="ckr"><input type="checkbox" class="ap-resall2" data-m="'+x[0]+'"'+(n===U.length&&U.length?' checked':'')+'> <span>'+esc(x[1])+(n&&n<U.length?' <em class="small">(en '+n+' de '+U.length+' unidades)</em>':'')+'</span></label>'+(n?'<input class="in ap-resd" data-m="'+x[0]+'" value="'+esc(d||'')+'" placeholder="Observación" autocomplete="off">'+(dd?'<em class="small dsu">El texto es distinto según la unidad: si lo cambias aquí, se pone en todas.</em>':''):'')+'</div>'});
    var pmax=-1;U.forEach(function(i){Cs[i].rs.forEach(function(r){var x=((ud[i].rc||{})[r.ref]||{}).pe;if(abHas(x))pmax=Math.max(pmax,+x)})});
    h+='<div class="small">Pendiente máxima de los recintos: <b>'+(pmax<0?'sin dato':trTxt(pmax,1)+' %')+'</b>'+hlp('pend')+'.</div>'}
  else if(k==='mo'){var mos={};U.forEach(function(i){mos[(ud[i].mo||{}).obj||'']=1});var cur=Object.keys(mos).length===1?Object.keys(mos)[0]:'';
    h+=fld('Objetivo de materia orgánica (para toda la finca)'+hlp('mo'),'<select class="in ap-moall"><option value=""'+(cur?'':' selected')+'>'+(Object.keys(mos).length>1?'Distinto en cada unidad':'Elegir…')+'</option><option value="mantener"'+(cur==='mantener'?' selected':'')+'>Mantener</option><option value="aumentar"'+(cur==='aumentar'?' selected':'')+'>Aumentar</option></select>');
    each(function(u,C,i,P){var mo=u.mo||{};return '<div class="g3">'+fld('Enmienda o abono orgánico',pabTx(P+'mo.enm',mo.enm,'Ej. Estiércol de oveja o compost'))+fld('Dosis (t/ha)',pabIn(P+'mo.dose',mo.dose))+fld('Cubierta vegetal y manejo del suelo',pabTx(P+'mo.cub',mo.cub,'Ej. Cubierta espontánea en las calles'))+'</div><div class="small">MO del análisis: '+(abHas((u.su||{}).mo)?trTxt(+u.su.mo,2)+' %':'sin dato')+' · N de abonos orgánicos: '+trTxt(pabNorg(C),1)+' kg/ha'+(u.zv?' (máximo 170 en zona vulnerable)':'')+'.</div>'+pabOrgHTML(u,C)})}
  else if(k==='fi'){var A=pabAs(),t=(f.datos||{}).tit||{};
    h+='<div class="g3">'+fld('Fecha del plan','<input type="date" class="in pa-in" data-p="fe" value="'+esc(pabIso(p.fe))+'">')+fld('Fecha del asesoramiento','<input type="date" class="in pa-in" data-p="fa" value="'+esc(pabIso(p.fa))+'">')+fld('Titular que firma',pabTx('ft',p.ft,t.rn||t.n||'Nombre del titular o representante'))+'</div>';
    h+='<div class="gk">Asesor en fertilización'+hlp('as')+' <span class="small">(se guarda una vez para todas las fincas)</span></div><div class="g3">'+[['n','Nombre y apellidos'],['nif','NIF'],['tit','Titulación'],['regfer','Nº REGFER (sección asesores)'],['acr','Acreditación']].map(function(x){return fld(x[1],pabTx('AS.'+x[0],A[x[0]]))}).join('')+'</div>'}
  else if(k==='su'){each(function(u,C,i,P){var su=u.su||{},src=suFuente(p,u);if(src)return '<div class="small">Usa el mismo análisis que «'+esc(src.n)+'» (se cambia en la tabla de Datos).</div>';
      return '<div class="g3">'+fld('Recinto representativo'+hlp('rp'),'<select class="in pa-in" data-p="'+P+'rp">'+C.rs.map(function(r){return '<option value="'+esc(r.ref)+'"'+(u.rp===r.ref?' selected':'')+'>'+esc(refLabel(r.ref)+(r.nombre?' · '+r.nombre:''))+'</option>'}).join('')+'</select>')+
        fld('Fecha de muestreo','<span class="dt"><input type="date" class="in pa-in" data-p="'+P+'su.f" value="'+esc(pabIso(su.f)||'')+'" aria-label="Fecha de muestreo"></span>')+fld('Laboratorio',pabTx(P+'su.lab',su.lab))+fld('Nº de informe',pabTx(P+'su.inf',su.inf))+fld('Profundidad (cm)',pabTx(P+'su.prof',su.prof,'0-30'))+fld('N total (mg/kg)',pabIn(P+'su.n',su.n))+
        fld('Textura','<select class="in pa-in" data-p="'+P+'su.tex"><option value="">—</option>'+PAB_TEX.map(function(x){return '<option value="'+x[0]+'"'+(su.tex===x[0]?' selected':'')+'>'+x[1]+'</option>'}).join('')+'</select>')+fld('Contaminantes analizados',pabTx(P+'su.cont',su.cont,'No analizados'))+'</div>'+sueloHTML(u,C,i)})}
  else if(k==='ag'){each(function(u,C,i,P){if(u.rg!=='regadio')return '<div class="small">Secano.</div>';var ag=u.ag||{};
      return '<div class="g3">'+fld('Dotación prevista (m³/ha)',pabIn(P+'ag.m3',ag.m3))+fld('Dotación máxima concedida (m³/ha·año)',pabIn(P+'ag.max',ag.max))+fld('Nitratos del agua (mg NO₃/L)',pabIn(P+'ag.no3',ag.no3))+fld('Análisis del agua (laboratorio, nº, fecha)',pabTx(P+'ag.src',ag.src),'w2')+'</div>'+fld('¿Riego localizado (goteo)?',pabChips(P+'loc',u.loc===0?0:1,[[1,'Sí'],[0,'No']]))+'<div class="small">N que trae el agua: <b>'+trTxt(C.nAg,1)+' kg N/ha</b>. Con riego localizado no hay meses prohibidos para el N.</div>'})}
  else if(k==='cl'){var cl=p.cl||PAB_CL,sum=(+cl.o||0)+(+cl.i||0)+(+cl.p||0)+(+cl.v||0),ck=climaKey(f);
    h+='<p class="small">'+hlp('cl')+'Valores normales de AEMET de la estación principal más cercana; cámbialos si tienes datos mejores.</p><div class="g5">'+[['a','Media anual (mm)'],['o','Otoño, sep-nov'],['i','Invierno, dic-feb'],['p','Primavera, mar-may'],['v','Verano, jun-ago']].map(function(x){return fld(x[1],pabIn('cl.'+x[0],cl[x[0]]))}).join('')+'</div>'+fld('Fuente',pabTx('cl.src',cl.src));
    h+='<div class="small">Suma de las estaciones: '+nf(sum,0)+' mm'+(abHas(cl.a)&&Math.abs(sum-cl.a)>2?' · <b>no cuadra con la media anual</b>':'')+'.'+(ck.km!=null?' La estación más cercana a los recintos es '+esc(CLIMA[ck.k].n)+' (a unos '+ck.km+' km).':'')+'</div>';
    h+=fld('Poner los valores de AEMET de','<select class="in ap-clprov"><option value="">Elegir estación…</option>'+Object.keys(CLIMA).map(function(k){return '<option value="'+k+'">'+CLIMA[k].n+(+k===ck.k?' (la más cercana)':'')+'</option>'}).join('')+'</select>')}
  else if(k==='pz'){h+='<label class="ckr"><input type="checkbox" class="ap-cost"'+(UI.cost?' checked':'')+'> Ver los costes en el programa</label><p class="small">Los precios solo se ven en pantalla: no salen en el PDF ni en el Excel.</p>';
    each(function(u,C,i,P){var rows=(u.pg||{}).rows||[];if(!rows.length)return '<div class="small">Sin abonos.</div>';var Cst=costeUd(u,C);
      return '<table class="mini"><tbody>'+rows.map(function(q,j){return '<tr><td>'+esc(q.n||'Abono')+'</td><td class="r">'+trTxt(+q.d||0,0)+' kg/ha</td><td class="r">'+pabIn(P+'pg.rows.'+j+'.pz',q.pz,'0,00','90px')+' €/kg</td></tr>'}).join('')+'</tbody></table><div class="small">Coste: <b>'+nf(Cst.ha,2)+' €/ha</b> · '+nf(Cst.tot,0)+' € en la unidad'+(Cst.sin?' (faltan precios)':'')+'.</div>'});
    h+='<div class="acts"><button class="btn q" data-act="ap-mis">'+ic('bag')+'Mis abonos</button></div>'}
  else if(k==='ne'){each(function(u,C,i,P){var ne=u.ne||{},mo=ne.mo||'cos',ex=ne.ex||(PAB_EX[u.ck]?{n:PAB_EX[u.ck][0][0],p:PAB_EX[u.ck][0][1],k:PAB_EX[u.ck][0][2]}:{});
      var x=fld('Cómo se calcula lo que necesita',pabChips(P+'ne.mo',mo,[['cos','Por cosecha esperada'],['edad','Plantación joven (por edad)'],['man','A mano']]));
      if(mo==='cos')x+='<div class="g4">'+fld('Cosecha esperada (kg/ha)',pabIn(P+'ne.pr',ne.pr))+PAB_NU.map(function(y){return fld('Extracción '+y[1]+' (kg por t)',pabIn(P+'ne.ex.'+y[0],ex[y[0]]))}).join('')+'</div>';
      else if(mo==='edad')x+=fld('Años de la plantación (verdes)',pabChips(P+'ne.ed',ne.ed||1,[[1,'1º'],[2,'2º'],[3,'3º'],[4,'4º'],[5,'5º']]));
      else x+='<div class="g4">'+PAB_NU.map(function(y){return fld('Necesidad '+y[1]+' (kg/ha)',pabIn(P+'ne.nb.'+y[0],(ne.nb||{})[y[0]]))}).join('')+fld('Cosecha esperada (kg/ha; 0 = sin cosecha)',pabIn(P+'ne.pr',ne.pr))+'</div>'+fld('Justificación de la necesidad',pabTx(P+'ne.src',ne.src,'Ej. Criterio técnico del asesor'));
      x+='<div class="small'+(noExtr(u)?' wtx':'')+'">'+(noExtr(u)?SIN_EX_TXT+': escribe las tuyas.':esc(C.src)+'.')+'</div>';
      x+='<div class="g3">'+fld('Cultivo precedente'+hlp('prec'),pabTx(P+'prec',u.prec,'Ej. Cebada'))+fld('¿El precedente es una leguminosa?',pabChips(P+'leg',u.leg?1:0,[[1,'Sí'],[0,'No']]))+fld('Rotación de cultivos (herbáceos)',pabTx(P+'rot',u.rot,'Ej. Cereal - barbecho'))+'</div>';
      x+='<div class="g3">'+fld('¿Sin abonado esta campaña?',pabChips(P+'d0',u.d0===1?1:0,[[1,'Sí, dosis 0'],[0,'No, hay programa']]))+fld('Observaciones del balance',pabTx(P+'bobs',u.bobs,'Ej. Exceso de P₂O₅ de fondo que se compensa'),'w2')+'</div>';
      return x})}
  else if(k==='ap'){h+='<p class="small">Lo que ya aportan el suelo, el estiércol u otros, en U.F./ha: se descuenta de lo que hay que poner. El abono orgánico del programa no va aquí.'+hlp('suelo')+'</p>';
    each(function(u,C,i,P){var ap=u.ap||{};return '<table class="mini"><thead><tr><th></th>'+PAB_NU.map(function(x){return '<th class="r">'+x[1]+'</th>'}).join('')+'</tr></thead><tbody>'+[['s','Suelo'],['o','Estiércol y restos orgánicos'],['x','Otros']].map(function(y){return '<tr><td>'+y[1]+'</td>'+PAB_NU.map(function(x){return '<td class="r">'+pabIn(P+'ap.'+y[0]+'.'+x[0],(ap[y[0]]||{})[x[0]],'0','70px')+'</td>'}).join('')+'</tr>'}).join('')+'</tbody></table>'})}
  else if(k==='ud'){var man=udMan(f,p);
    h+='<p class="small">Las unidades se forman solas por cultivo y riego'+hlp('ud')+'. '+(man?'<b>Este plan tiene las unidades cambiadas a mano</b>, así que ya no se rehacen solas. ':'Si cambias aquí el nombre o mueves un recinto, dejan de rehacerse solas. ')+'<button class="lnk" data-act="ap-udrehacer">Rehacer por cultivo y riego</button></p>';
    each(function(u,C,i,P){var x='<div class="g3">'+fld('Nombre de la unidad',pabTx(P+'n',u.n))+fld('Régimen',pabChips(P+'rg',u.rg,[['secano','Secano'],['regadio','Regadío']]))+(u.zv?fld('Tipo de parcela'+hlp('tp'),pabChips(P+'tp',u.tp||0,[[0,'Según el suelo (tipo '+pabTipo(f,u,p).t+')'],[1,'Tipo 1'],[2,'Tipo 2']]))+fld('¿Inhibidores? (+10 % de N)',pabChips(P+'inh',u.inh?1:0,[[1,'Sí'],[0,'No']])):'')+'</div>';
      x+='<table class="mini rcs"><thead><tr><th>Recinto</th><th class="r">ha cultiv.</th><th>ha en este plan</th><th>Pendiente %</th><th>Perímetro de protección</th><th>Repr.'+hlp('rp')+'</th><th>Unidad</th></tr></thead><tbody>'+C.rs.map(function(r){var rc=(u.rc||{})[r.ref]||{};
        return '<tr><td>'+esc(refLabel(r.ref))+(r.nombre?' <span class="small">'+esc(r.nombre)+'</span>':'')+'</td><td class="r">'+ha2(haC(r))+'</td><td>'+pabIn(P+'rc.'+r.ref+'.ph',rc.ph,ha2(haC(r)),'80px')+'</td><td>'+pabIn(P+'rc.'+r.ref+'.pe',rc.pe,'','64px')+'</td><td>'+pabChips(P+'rc.'+r.ref+'.pp',rc.pp===1?1:(rc.pp===0?0:''),[[0,'No'],[1,'Sí']])+'</td><td><button type="button" class="chip'+(u.rp===r.ref?' on':'')+'" data-act="pa-set" data-p="'+P+'rp" data-v="'+esc(r.ref)+'" aria-label="Recinto representativo">'+(u.rp===r.ref?'Sí':'—')+'</button></td><td><select class="in pa-mv" data-u="'+i+'" data-ref="'+esc(r.ref)+'">'+(p.ud||[]).map(function(y,j){return '<option value="'+j+'"'+(j===i?' selected':'')+'>'+esc(y.n)+'</option>'}).join('')+'<option value="new">Unidad nueva…</option></select></td></tr>'}).join('')+'</tbody></table>';
      return x+'<div class="acts"><button class="btn q danger" data-act="pa-ud-del" data-u="'+i+'">'+ic('trash')+'Borrar esta unidad</button></div>'},true);
    h+='<div class="acts"><button class="btn" data-act="pa-ud-new">'+ic('plus')+'Unidad nueva</button></div>'}
  else if(k==='compra')h+=compraHTML(f,p,Cs)||'<div class="small">Aún no hay abonos en el programa.</div>';
  else if(k==='md'){h+='<p class="small">El plan se puede cambiar durante la campaña según evolucione el cultivo o el tiempo: apunta aquí cada cambio.</p>';
    if((p.mods||[]).length)h+='<table class="mini"><thead><tr><th>Fecha</th><th>Unidad / recintos</th><th>Motivo</th><th>Cambio</th><th>Responsable</th><th></th></tr></thead><tbody>'+p.mods.map(function(m,j){var R='mods.'+j+'.';return '<tr><td><input type="date" class="in pa-in" data-p="'+R+'f" value="'+esc(pabIso(m.f))+'"></td><td>'+pabTx(R+'u',m.u)+'</td><td>'+pabTx(R+'mo',m.mo)+'</td><td>'+pabTx(R+'ch',m.ch)+'</td><td>'+pabTx(R+'re',m.re)+'</td><td><button type="button" class="btn q ib danger" data-act="pa-mod-del" data-j="'+j+'" aria-label="Quitar">'+ic('close')+'</button></td></tr>'}).join('')+'</tbody></table>';
    h+='<div class="acts"><button class="btn" data-act="pa-mod-add">'+ic('plus')+'Añadir un cambio</button></div>'}
  return h}

/* v3.2 (propuesta 2.4): sin semáforo ni crédito del suelo (cortes sin fuente). Queda la función para no cambiar quien la llama. */
function sueloHTML(u,C,i){return ''}

function compraHTML(f,p,Cs){var L=compra(f,p,Cs);if(!L.length)return '';var co=!!UI.cost, tot=L.reduce(function(a,x){return a+x.eur},0);
  return '<table class="mini shop"><thead><tr><th>Abono</th><th class="r">Total</th><th class="r">Equivale a</th>'+(co?'<th class="r">Coste</th>':'')+'<th>Unidades</th></tr></thead><tbody>'+
    L.map(function(x){var kg=x.kg;return '<tr><td><b>'+esc(x.n)+'</b></td><td class="r">'+(kg>=1000?nf(kg/1000,2)+' t':nf(kg,0)+' kg')+'</td><td class="r small">'+(x.dens?'≈ '+nf(kg/x.dens,0)+' L':'≈ '+nf(Math.ceil(kg/25),0)+' sacos de 25 kg')+'</td>'+(co?'<td class="r">'+(x.eur?nf(x.eur,0)+' €':'—')+'</td>':'')+'<td class="small">'+esc(Object.keys(x.uds).join(', '))+'</td></tr>'}).join('')+
    '</tbody>'+(co&&tot?'<tfoot><tr><td colspan="3"><b>Total</b></td><td class="r"><b>'+nf(tot,0)+' €</b></td><td></td></tr></tfoot>':'')+'</table><div class="acts"><button class="btn q" data-act="ap-compra-copy">'+ic('copy')+'Copiar la lista</button><span class="small">Toda la finca, campaña '+esc(pabCamp(p.y))+'. Solo en pantalla.</span></div>'}

/* ───── la mesa de la derecha: foto de la finca (Finca y Datos) o el PDF (Plan y PDF) ───── */
function deskHTML(f,ctx){
  if(UI.step==='plan')return '<div class="dh"><b>El PDF</b><button class="btn p" data-act="ap-pdf" title="Descargar el PDF (Ctrl+P)">'+ic('download')+'Descargar PDF</button><button class="btn" data-act="ap-xlsx">'+ic('grid')+'Excel</button></div><div id="prev">'+prevInner()+'</div>';
  var rs=fincaRecs(f.id),N=recNum(f);
  return '<div class="dh"><b>Tu finca</b><button class="btn" data-act="ap-mapa">'+ic('map')+'Abrir el mapa</button></div>'+
    '<div class="mapw"><div class="aer big" data-aer="desk" data-fid="'+esc(f.id)+'" data-lab="1"></div><span class="tag">Foto aérea con tus recintos</span></div>'+
    (rs.length?'<div class="leg">'+rs.map(function(r){return '<div data-ref="'+esc(r.ref)+'" class="'+(HL.indexOf(r.ref)>=0?'on':'')+'"><i style="background:'+glOf(r.cultivo).c+'"></i><span><b>'+N[r.ref]+'</b> · '+esc(refLabel(r.ref))+(r.nombre?' · '+esc(r.nombre):'')+'</span><b>'+ha2(haC(r))+' ha</b></div>'}).join('')+'</div>':'')+
    '<div class="cap2">'+(UI.step==='dat'?'En amarillo, los recintos de la unidad que estás rellenando.':'En amarillo, el recinto de la fila en la que estás.')+'</div>'}
/* recintos que se resaltan según la casilla en la que estás */
function hlAuto(){var a=document.activeElement,f=finca(UI.fid);if(!f||!CTX)return;var refs=[];
  if(a&&a.closest){var tr=a.closest('[data-ref]');if(tr&&UI.step==='fin')refs=[tr.dataset.ref];
    var uu=a.closest('[data-u]');if(uu&&CTX.p&&CTX.p.ud[+uu.dataset.u])refs=CTX.Cs[+uu.dataset.u].rs.map(function(r){return r.ref})}
  if(!refs.length&&UI.step==='dat'&&CTX.p){var x=CTX.F.filter(function(y){return y.s==='dat'&&y.u!=null})[0];if(x)refs=CTX.Cs[x.u].rs.map(function(r){return r.ref})}
  hlSet(refs)}
function hlSet(refs){HL=refs||[];$$('#desk .leg [data-ref]').forEach(function(e){e.classList.toggle('on',HL.indexOf(e.dataset.ref)>=0)});aerHL()}

/* ───── vista previa del PDF de verdad, con PDF.js (punto 46): hoja grande y miniaturas; se pone al día sola ───── */
var PREV={key:null,imgs:[],big:0,t:null,busy:false,again:false,err:''};
function prevInner(){if(PREV.err&&!PREV.imgs.length)return '<div class="paper ph">'+esc(PREV.err)+'</div>';
  if(!PREV.imgs.length)return '<div class="paper ph">Preparando la vista previa…</div>';
  var b=Math.min(PREV.big,PREV.imgs.length-1);
  return '<div class="paper"><img src="'+PREV.imgs[b]+'" alt="Hoja '+(b+1)+' del PDF"></div><div class="pgs">'+PREV.imgs.map(function(s,i){return '<button class="th'+(i===b?' on':'')+'" data-act="ap-prevpg" data-i="'+i+'" aria-label="Ver la hoja '+(i+1)+'"><img src="'+s+'" alt=""></button>'}).join('')+'</div><div class="cap2">Hoja '+(b+1)+' de '+PREV.imgs.length+' · se pone al día cada vez que cambias algo'+(PREV.t||PREV.busy?' · <span class="upd">poniéndose al día…</span>':'')+'</div>'}
function prevPinta(){var e=$('#prev');if(e)e.innerHTML=prevInner()}
function prevAuto(){if(!$('#prev'))return;var f=finca(UI.fid),p=f&&pabCur(f);if(!p)return;var key=f.id+':'+p.id+':'+S.mod;if(PREV.key===key&&!PREV.t)return;
  if(PREV.fid!==f.id+p.id){PREV.imgs=[];PREV.big=0;PREV.fid=f.id+p.id;prevPinta()}
  clearTimeout(PREV.t);PREV.t=setTimeout(function(){PREV.t=null;prevMake()},PREV.imgs.length?900:120)}
function prevMake(){var f=finca(UI.fid),p=f&&pabCur(f);if(!p||!$('#prev'))return;if(PREV.busy){PREV.again=true;return}PREV.busy=true;var key=f.id+':'+p.id+':'+S.mod;
  var fin=function(){PREV.busy=false;if(PREV.again){PREV.again=false;prevAuto()}};
  ensurePdf().then(function(){var blob=buildPabPdf(f,p);return Promise.all([blob.arrayBuffer(),ensurePdfJs()])}).then(function(r){return r[1].getDocument({data:new Uint8Array(r[0]),isEvalSupported:false}).promise}).then(function(pdf){
    var out=[],n=pdf.numPages,one=function(i){if(i>n)return Promise.resolve();return pdf.getPage(i).then(function(pg){var v=pg.getViewport({scale:1}),sc=760/v.width,vp=pg.getViewport({scale:sc}),c=document.createElement('canvas');c.width=Math.round(vp.width);c.height=Math.round(vp.height);
      return pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise.then(function(){out.push(c.toDataURL('image/jpeg',0.86));return one(i+1)})})};
    return one(1).then(function(){PREV.imgs=out;PREV.key=key;PREV.err='';if(PREV.big>=out.length)PREV.big=0;prevPinta();fin()})})
  .catch(function(){PREV.err='Para la vista previa hace falta abrir la herramienta una vez con internet.';prevPinta();fin()})}

/* ───── inicio (puntos 31 y 47) ───── */
function fincaEst(f){var rs=fincaRecs(f.id);if(!rs.length)return {k:'m',t:'Sin recintos'};if(!(f.pab||[]).length)return {k:'w',t:'Sin plan'};var c=ctxOf(f),G=c.G;
  return G.k==='ok'?{k:'ok',t:'Plan '+pabCamp(c.p.y)+' listo',b:'Abrir'}:{k:'w',t:G.f?G.t:G.t,b:'Seguir'}}
function inicioHTML(){var sinCopia=S.fincas.length&&S.mod>(S.copia||0)+1000&&!FS.ok;
  var h='<div class="home"><section class="hero"><div class="pad"><div class="brand"><span class="lg">'+ic('leaf')+'</span>Plan de abonado</div><h1>Del recinto al plan de abonado, <em>en tres pasos</em></h1><p>El plan que pide el Real Decreto 1051/2022, listo en PDF y en Excel. Sin cuentas: todo se queda en este ordenador.</p>'+
    '<div class="how"><div><i>1</i><div><b>Tu finca</b><span>Lee el PDF de la PAC o elige los recintos en el mapa.</span></div></div><div><i>2</i><div><b>Los datos justos</b><span>Cosecha esperada, agua de riego y análisis de suelo.</span></div></div><div><i>3</i><div><b>Plan y PDF</b><span>Los abonos salen propuestos y las cuentas, hechas.</span></div></div></div></div>'+
    '<div class="hpic" aria-hidden="true"></div><div class="legal">Real Decreto 1051/2022 · Decreto 80/2025 de Castilla-La Mancha</div></section>';
  h+='<section class="start"><div class="eyebrow">Empieza una finca</div><h2>¿De dónde saco tus recintos?</h2><div class="tiles">'+
    '<div class="tile rec"><button class="drop" data-act="ap-pdf-pick" aria-label="Elegir o arrastrar el PDF de la PAC">'+ic('upload')+'Arrastra aquí el PDF de la PAC</button><div class="tx"><b>Leer el PDF de la PAC</b><span>Salen solos los recintos con su cultivo, su superficie y su riego. Es lo más rápido.</span><button class="btn p" data-act="ap-pdf-pick">'+ic('pdf')+'Elegir el PDF</button></div></div>'+
    '<div class="tile"><div class="top mpic" aria-hidden="true"><svg viewBox="0 0 400 150" preserveAspectRatio="none"><path d="M40 18 252 8 262 138 52 146Z" fill="none" stroke="#e9c838" stroke-width="2.5"/><path d="M268 8 392 4 398 132 276 138Z" fill="none" stroke="#fff" stroke-width="1.5"/></svg></div><div class="tx"><b>Elegir en el mapa</b><span>Foto aérea del SIGPAC. Buscas tu pueblo y tocas tus recintos.</span><button class="btn" data-act="ap-mapa-nueva">'+ic('map')+'Abrir el mapa</button></div></div></div>'+
    '<div class="others">Otras formas:<button class="lnk" data-act="ap-mano">Escribir los códigos a mano</button><button class="lnk" data-act="ap-csv-pick">Tabla CSV</button><button class="lnk" data-act="ap-abrir">Abrir una copia guardada</button><button class="lnk" data-act="ap-ejemplo">Ver una finca de ejemplo</button></div>';
  if(S.fincas.length){var q=noAcc(UI.q||'').toLowerCase().trim(), L=S.fincas.slice().sort(function(a,b){return String(a.nombre).localeCompare(String(b.nombre),'es')}).filter(function(f){if(!q)return true;var t=(f.datos||{}).tit||{};return noAcc([f.nombre,t.n,t.nif,t.mun].concat(fincaRecs(f.id).map(function(r){return refLabel(r.ref)+' '+(r.cultivo||'')})).join(' ')).toLowerCase().indexOf(q)>=0});
    h+='<div class="mine"><div class="mh"><div class="eyebrow">Tus fincas</div><span class="sp"></span>'+(S.fincas.length>1||UI.q?'<label class="srch">'+ic('search')+'<input id="ap-q" class="in" placeholder="Buscar finca, titular, recinto o cultivo" value="'+esc(UI.q||'')+'" autocomplete="off" aria-label="Buscar"></label>':'')+'</div>';
    if(sinCopia)h+='<div class="warn">Hay cambios sin copia en archivo: lo que metes solo se guarda en este navegador. '+(fsCan()?'<button class="lnk" data-act="ap-fs-pick">Guardar solo en un archivo (recomendado)</button> · ':'')+'<button class="lnk" data-act="ap-copia">Descargar una copia</button></div>';
    h+='<div class="frows">'+L.map(function(f){var rs=fincaRecs(f.id),e=fincaEst(f),cs={};rs.forEach(function(r){if(r.cultivo)cs[r.cultivo.toLowerCase()]=1});var cl=Object.keys(cs);
      return '<div class="frow tap" data-act="ap-open" data-id="'+esc(f.id)+'" tabindex="0" role="button" aria-label="Abrir '+esc(f.nombre)+'"><div class="aer th" data-aer="home-'+esc(f.id)+'" data-fid="'+esc(f.id)+'"></div><div class="fx"><b>'+esc(f.nombre)+'</b><span class="m">'+esc([ha2(haSum(rs))+' ha',nRec(rs.length),cl.length?(cl.length>2?cl.slice(0,2).join(', ')+'…':cl.join(' y ')):''].filter(Boolean).join(' · '))+'</span></div><span class="sp"></span>'+chipEst(e)+
        '<button class="btn q ib" data-act="ap-copiaf-id" data-id="'+esc(f.id)+'" title="Guardar copia de esta finca" aria-label="Copia de '+esc(f.nombre)+'">'+ic('save')+'</button><button class="btn q ib danger" data-act="ap-fdel-id" data-id="'+esc(f.id)+'" title="Borrar la finca" aria-label="Borrar '+esc(f.nombre)+'">'+ic('trash')+'</button><span class="btn op">'+(e.b||'Seguir')+ic('next')+'</span></div>'}).join('')+
      (L.length?'':'<div class="empty">Ninguna finca coincide con «'+esc(UI.q)+'».</div>')+'</div></div>'}
  h+='<div class="foot2">Versión '+esc(VERSION)+' · <a href="privacidad.html">Privacidad</a> · <a href="condiciones.html">Condiciones</a> · © 2026 Ramón Alcalá</div></section></div>';
  return h}

/* ───── retoques al HTML: fechas propias (dd/mm/aaaa con calendario) ───── */
function isoEs(v){var m=/^(\d{4})-(\d\d)-(\d\d)$/.exec(String(v||''));return m?m[3]+'/'+m[2]+'/'+m[1]:String(v||'')}
function fixHTML(h){
  h=h.replace(/<input type="date" class="in pa-in" data-p="([^"]+)" value="([^"]*)"( aria-label="[^"]*")?>/g,function(m,p,v,al){return '<span class="dt"><input class="in pa-in dti" data-p="'+p+'" data-date="1" value="'+esc(isoEs(v))+'" placeholder="dd/mm/aaaa" inputmode="numeric" autocomplete="off"'+(al||' aria-label="Fecha"')+'><button type="button" class="dtb" data-act="ap-cal" aria-label="Elegir la fecha en el calendario">'+ic('cal')+'</button></span>'});
  return h}

/* ───── visor SIGPAC para añadir recintos: foto aérea PNOA de máxima actualidad, recintos y datos de cada recinto ───── */
var IGNW='https://www.ign.es/wmts/', MZ=15, MNZ=Math.pow(2,MZ), RT=6378137;
var MAP={m:null,lay:null,grid:null,tiles:new Map(),feat:{},top:{},sel:null,info:{},zvn:{},busy:0,fail:0,places:null};
var refKey=function(p){return [p.provincia,p.municipio,p.agregado||0,p.zona||0,p.poligono,p.parcela,p.recinto].join(':')};
function m2ll(p){return [p[0]/RT*180/Math.PI,(2*Math.atan(Math.exp(p[1]/RT))-Math.PI/2)*180/Math.PI]}
function projG(g){var cv=function(r){return r.map(m2ll)};if(!g)return null;if(g.type==='Polygon')return {type:'Polygon',coordinates:g.coordinates.map(cv)};if(g.type==='MultiPolygon')return {type:'MultiPolygon',coordinates:g.coordinates.map(function(p){return p.map(cv)})};return null}
function wktGeo(w){var t=String(w||'').trim(),multi=/^MULTIPOLYGON/i.test(t),s=t.replace(/^[A-Za-z]+\s*/,'').replace(/\(/g,'[').replace(/\)/g,']').replace(/(-?[\d.eE+-]+)\s+(-?[\d.eE+-]+)/g,'[$1,$2]');try{var a=JSON.parse(s);return multi?{type:'MultiPolygon',coordinates:a}:{type:'Polygon',coordinates:a}}catch(e){return null}}
/* sin finca (inicio): la finca se crea al añadir el primer recinto (punto 31) */
function mapaAbrir(){var f=UI.v==='finca'?finca(UI.fid):null;MAP.nueva=!f;sheetClose();
  var el=$('#mapa');el.hidden=false;document.body.classList.add('mapon');
  el.innerHTML='<div class="mtop"><div class="mtt"><b>Añadir recintos desde el mapa</b><span class="small">'+(f?esc(f.nombre):'La finca se crea al añadir el primero')+'</span></div><form class="msr" data-f="m-go"><input id="m-q" class="in" placeholder="Código SIGPAC (2:81:1:1:1) o municipio" autocomplete="off" aria-label="Buscar"><button class="btn" type="submit">'+ic('search')+'<span>Buscar</span></button></form>'+
    '<label class="mnom" title="Nombres de pueblos, caminos y parajes sobre la foto"><input type="checkbox" id="m-nom"'+(UI.mNom===0?'':' checked')+'> Nombres</label><button class="btn" data-act="ap-m-loc" title="Mi ubicación">'+ic('loc')+'<span>Mi ubicación</span></button><button class="btn p" data-act="ap-m-close">Hecho</button></div><div id="m-res" class="mres" hidden></div>'+
    '<div class="mbody"><div id="mmap" role="application" aria-label="Mapa"></div><div id="mhint" class="mhint" hidden></div><aside id="mside" class="mside"></aside></div>';
  mSide();
  ensureMap().then(mInit,function(){$('#mmap').innerHTML='<div class="note" style="margin:20px">Para el mapa hace falta internet.</div>'})}
/* tesela vacía (PNG transparente de 1 px) para cuando el IGN no tiene la de los nombres */
var PNG0='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII=';
function mNombres(on){UI.mNom=on?1:0;saveUI();var m=MAP.m;if(!m||!MAP.nom)return;if(on&&!m.hasLayer(MAP.nom))MAP.nom.addTo(m);if(!on&&m.hasLayer(MAP.nom))m.removeLayer(MAP.nom)}
document.addEventListener('change',function(ev){var t=ev.target;if(t&&t.id==='m-nom')mNombres(t.checked)});
function mapaCerrar(){var el=$('#mapa');if(MAP.m){try{MAP.m.remove()}catch(e){}}MAP.m=null;MAP.lay=null;MAP.grid=null;MAP.nom=null;MAP.sel=null;MAP.fail=0;sigReset();el.hidden=true;el.innerHTML='';document.body.classList.remove('mapon');render()}
function mFin(){return MAP.nueva&&!UI.fid?null:finca(UI.fid)}
function mInit(){var f=mFin();if(!$('#mmap'))return;
  var m=MAP.m=L.map('mmap',{minZoom:6,maxZoom:20,zoomControl:true,attributionControl:true,preferCanvas:true});m.attributionControl.setPrefix(false);
  L.tileLayer(IGNW+'pnoa-ma?service=WMTS&request=GetTile&version=1.0.0&format=image/jpeg&layer=OI.OrthoimageCoverage&style=default&tilematrixset=GoogleMapsCompatible&TileMatrix={z}&TileRow={y}&TileCol={x}',{maxNativeZoom:19,maxZoom:20,attribution:'Foto aérea: PNOA máxima actualidad © IGN · Recintos: SIGPAC (FEGA)'}).addTo(m);
  /* v3.1 punto 8: los nombres (capa transparente del IGN) encima de la foto y debajo de los recintos; si falta una tesela, sale vacía y no rota */
  MAP.nom=L.tileLayer(IGNW+'ign-base?service=WMTS&request=GetTile&version=1.0.0&format=image/png&layer=IGNBaseOrto&style=default&tilematrixset=GoogleMapsCompatible&TileMatrix={z}&TileRow={y}&TileCol={x}',{maxNativeZoom:19,maxZoom:20,zIndex:2,errorTileUrl:PNG0,className:'mnomt',attribution:'Nombres: © IGN'});
  if(UI.mNom!==0)MAP.nom.addTo(m);
  /* v3.2: los recintos «libres» van por cuadros pintados en un hilo aparte (SigGrid); encima, solo los de las fincas y el seleccionado */
  m.createPane('sigp');m.getPane('sigp').style.zIndex=390;m.getPane('sigp').style.pointerEvents='none';
  MAP.grid=SigGrid().on('tileunload',function(e){var el=e.tile;if(el&&el._id){delete SWel[el._id];if(SW)SW.postMessage({t:'x',id:el._id});el._id=0;el._done=null}});
  MAP.lay=L.geoJSON(null,{renderer:L.canvas({padding:0.4,tolerance:3}),style:mEstilo,interactive:false}).addTo(m);
  m.on('click',function(ev){var h=hitAt(ev.latlng);if(h)mSel(h.p,ev.latlng,h.g)});
  m.on('moveend',mTiles);
  var pts=(f?fincaRecs(f.id):[]).filter(function(r){return r.ll}).map(function(r){return r.ll});
  if(pts.length){if(pts.length===1)m.setView(pts[0],17);else m.fitBounds(L.latLngBounds(pts).pad(0.25),{maxZoom:17})}
  else{var r0=f?fincaRecs(f.id)[0]:null;m.setView([39.3,-2.6],8);if(r0)sigInfo(r0.ref).then(function(o){var ll=wktLL(o.wkt);if(ll&&MAP.m)MAP.m.setView(ll,16)}).catch(function(){});else mHint('Busca tu municipio o un código SIGPAC, o acércate en el mapa. Los recintos se ven a partir del zoom 15.')}
  setTimeout(function(){if(MAP.m)MAP.m.invalidateSize()},60)}
function mHint(t){var e=$('#mhint');if(!e)return;e.hidden=!t;e.textContent=t||''}
function mEstilo(ft){var p=ft.properties||{},k=refKey(p),r=S.recintos[k],mine=r&&!!UI.fid&&r.fincaId===UI.fid,oth=r&&!mine&&finca(r.fincaId),sel=MAP.sel&&MAP.sel.ref===k;
  return {color:sel?'#ffd400':(mine?'#7cff9a':'#ffffff'),weight:sel?3.2:(mine?2.2:1.2),opacity:sel?1:0.9,fill:true,fillColor:mine?'#2fbf5a':(oth?'#8fa3ff':'#ffffff'),fillOpacity:mine?0.35:(oth?0.25:(sel?0.12:0.02))}}
/* ───── v3.2: recintos SIGPAC por cuadros, como en la app de fincas (SigGrid, SIGW, sigReq, sigWMsg, sigRepaint, sigSkipSync,
   refreshSigpac, hitAt) ─────
   - SIGCORE guarda cada recinto UNA vez (con cuántos trozos lo traen) y pinta un cuadro de 256 px buscando en los trozos de
     alrededor; sirve igual dentro del Worker (OffscreenCanvas, lienzo 2D con willReadFrequently: lo pinta el procesador) que en
     el hilo del mapa si el navegador no tiene OffscreenCanvas.
   - En el hilo del mapa: MAP.tiles (trozo 'x,y' → {t, refs}) y MAP.feat (recinto → {p, g, bb, n}) para tocar (hitAt) y para
     dibujar encima los de las fincas y el seleccionado. Los trozos lejanos se sueltan con el tope de tres pantallas. */
function SIGCORE(){
  var R=6378137, TZ=15, SC=256*Math.pow(2,TZ), W=2*Math.PI*R, G={}, T={}, SK={}, ST={w:1.2,c:'#ffffff',o:0.9,fc:'#ffffff',fo:0.02};
  function add(k,j){
    if(T[k])return {refs:T[k],nu:[],c:new Float64Array(0)};
    var refs=[], nu=[], cs=[], seen={};
    (j&&j.features||[]).forEach(function(f){
      var g=f&&f.geometry, p=f&&f.properties||{}; if(!g||(g.type!=='Polygon'&&g.type!=='MultiPolygon'))return;
      var ref=[p.provincia,p.municipio,p.agregado||0,p.zona||0,p.poligono,p.parcela,p.recinto].join(':');
      if(seen[ref])return; seen[ref]=1; refs.push(ref);
      if(G[ref]){G[ref].n++;return}
      var P=g.type==='Polygon'?[g.coordinates]:g.coordinates, r=[], st=[], b=[Infinity,Infinity,-Infinity,-Infinity];
      P.forEach(function(pl){var rl=[];pl.forEach(function(ring){
        var a=new Float64Array(ring.length*2);
        for(var n=0;n<ring.length;n++){var X=ring[n][0], Y=ring[n][1], x=SC*(X/W+0.5), y=SC*(0.5-Y/W);
          a[2*n]=x; a[2*n+1]=y; if(x<b[0])b[0]=x; if(y<b[1])b[1]=y; if(x>b[2])b[2]=x; if(y>b[3])b[3]=y;
          cs.push(X/R*180/Math.PI,(2*Math.atan(Math.exp(Y/R))-Math.PI/2)*180/Math.PI)}
        rl.push(ring.length); r.push(a)}); st.push(rl)});
      G[ref]={b:b,r:r,n:1};
      nu.push({ref:ref,m:g.type==='MultiPolygon'?1:0,s:st,p:{provincia:p.provincia,municipio:p.municipio,agregado:p.agregado||0,zona:p.zona||0,poligono:p.poligono,parcela:p.parcela,recinto:p.recinto,dn_surface:p.dn_surface,uso_sigpac:p.uso_sigpac}});
    });
    T[k]=refs; return {refs:refs,nu:nu,c:new Float64Array(cs)};
  }
  function drop(ks){ks.forEach(function(k){var L=T[k];if(!L)return;delete T[k];L.forEach(function(ref){var o=G[ref];if(o&&--o.n<=0)delete G[ref]})})}
  function ring(ctx,a,ox,oy,k){var n=a.length, lx=0, ly=0;
    for(var q=0;q<n;q+=2){var X=(a[q]-ox)*k, Y=(a[q+1]-oy)*k;
      if(!q){ctx.moveTo(X,Y);lx=X;ly=Y}else if(q>=n-2||Math.abs(X-lx)+Math.abs(Y-ly)>=0.7){ctx.lineTo(X,Y);lx=X;ly=Y}}
    ctx.closePath()}
  /* m = {z,x,y,d}; ctx: lienzo de 256·d px */
  function paint(ctx,m){
    var d=m.d||1, w=Math.round(256*d), k=Math.pow(2,m.z-TZ), ox=m.x*256/k, oy=m.y*256/k, ex=ox+256/k, ey=oy+256/k, mg=4/k, px=Math.floor(m.x/k), py=Math.floor(m.y/k), n=0, seen={};
    ctx.setTransform(1,0,0,1,0,0); ctx.clearRect(0,0,w,w); ctx.setTransform(d,0,0,d,0,0);
    ctx.lineWidth=ST.w; ctx.strokeStyle=ST.c; ctx.fillStyle=ST.fc; ctx.lineCap='round'; ctx.lineJoin='round';
    var nx=Math.max(1,Math.ceil(1/k));
    for(var i=-1;i<=nx;i++)for(var j=-1;j<=nx;j++){
      var L=T[(px+i)+','+(py+j)]; if(!L)continue;
      for(var e=0;e<L.length;e++){var ref=L[e], f=G[ref]; if(!f||SK[ref]||seen[ref])continue; var b=f.b;
        if(b[2]<ox-mg||b[0]>ex+mg||b[3]<oy-mg||b[1]>ey+mg)continue;
        seen[ref]=1; ctx.beginPath(); for(var h=0;h<f.r.length;h++)ring(ctx,f.r[h],ox,oy,k);
        ctx.globalAlpha=ST.fo; ctx.fill('evenodd'); ctx.globalAlpha=ST.o; ctx.stroke(); n++}}
    ctx.globalAlpha=1; return n}
  function stats(){var g=0,t=0,s=0,k;for(k in G)g++;for(k in T){t++;s+=T[k].length}return {g:g,t:t,s:s}}
  return {add:add,drop:drop,paint:paint,stats:stats,skip:function(o){SK=o||{}},style:function(o){if(o)ST=o},reset:function(){G={};T={}}};
}
/* el hilo aparte: lee cada trozo (JSON de hasta 1 MB), guarda y pinta; al mapa solo le llega la imagen hecha */
function SIGW(CORE){
  var C=CORE(), Q=[], CV=null, CX=null, busy=0, gen=0, mc=new MessageChannel();
  function run(){busy=0;var t0=Date.now();
    while(Q.length&&Date.now()-t0<12){var m=Q.shift();
      try{var d=m.d||1, w=Math.round(256*d); if(!CV||CV.width!==w){CV=new OffscreenCanvas(w,w);CX=CV.getContext('2d',{willReadFrequently:true})}
        var n=C.paint(CX,m); if(n){var bm=CV.transferToImageBitmap();postMessage({t:'p',id:m.id,n:n,bmp:bm},[bm])}else postMessage({t:'p',id:m.id,n:0,bmp:null})}catch(e){postMessage({t:'p',id:m.id,err:1})}}
    if(Q.length)kick()}
  function kick(){if(!busy){busy=1;mc.port2.postMessage(0)}}
  mc.port1.onmessage=run;
  onmessage=function(e){var m=e.data;
    if(m.t==='a'){if(m.g!==gen){postMessage({t:'f',q:m.q,err:1});return}
      try{var r=C.add(m.k,JSON.parse(new TextDecoder().decode(m.b)));postMessage({t:'f',q:m.q,refs:r.refs,nu:r.nu,c:r.c},[r.c.buffer])}catch(err){postMessage({t:'f',q:m.q,err:1})}}
    else if(m.t==='p'){Q.push(m);kick()}
    else if(m.t==='x'){for(var i=0;i<Q.length;i++)if(Q[i].id===m.id){Q.splice(i,1);break}}
    else if(m.t==='d')C.drop(m.ks);
    else if(m.t==='s')C.skip(m.s);
    else if(m.t==='st')C.style(m.st);
    else if(m.t==='r'){gen=m.g;C.reset();Q.length=0}
    else if(m.t==='n')postMessage({t:'n',st:C.stats()});
  };
  var ok=false;try{var oc=new OffscreenCanvas(4,4),x=oc.getContext('2d',{willReadFrequently:true});x.fillRect(0,0,1,1);ok=!!(x&&oc.transferToImageBitmap())}catch(e){}
  postMessage({t:'ok',ok:ok});
}
var SIG_ST={w:1.2,c:'#ffffff',o:0.9,fc:'#ffffff',fo:0.02};
var SW=null, SWok=false, SWn=0, SWq=0, SWel={}, SWcb={}, SWsk=null, SWp=null, SCORE=null;
function sigArranca(){if(SWp)return SWp;
  SWp=new Promise(function(res){
    var fin=function(ok){if(!ok&&SW){try{SW.terminate()}catch(e){}SW=null}SWok=!!ok;if(!SWok&&!SCORE){SCORE=SIGCORE();SCORE.style(SIG_ST)}res()};
    try{
      if(window.__sinHilo||typeof Worker==='undefined'||typeof OffscreenCanvas==='undefined'||!window.URL||!URL.createObjectURL){fin(false);return}
      SW=new Worker(URL.createObjectURL(new Blob(['('+SIGW.toString()+')('+SIGCORE.toString()+')'],{type:'text/javascript'})));
      var to=setTimeout(function(){fin(false)},3000);
      SW.onmessage=function(e){var m=e.data;if(m.t==='ok'){clearTimeout(to);if(m.ok){SW.postMessage({t:'st',st:SIG_ST});SW.postMessage({t:'r',g:MAP.gen||0})}fin(m.ok);if(SW)SW.onmessage=sigWMsg}};
      SW.onerror=function(){clearTimeout(to);sigWOff();fin(false)};
    }catch(e){fin(false)}});
  return SWp}
/* si el hilo aparte falla: se sigue en el hilo del mapa y se vuelve a pedir lo que había */
function sigWOff(){if(!SWok)return;SWok=false;try{SW.terminate()}catch(e){}SW=null;
  Object.keys(SWcb).forEach(function(q){var f=SWcb[q];delete SWcb[q];f({err:1})});
  Object.keys(SWel).forEach(function(id){var el=SWel[id];delete SWel[id];el._id=0;var d=el._done;el._done=null;sigQueue(el,el._co,d)});
  if(!SCORE){SCORE=SIGCORE();SCORE.style(SIG_ST)}MAP.tiles.clear();MAP.feat={};SWsk=null;setTimeout(function(){if(MAP.grid&&MAP.grid._map)MAP.grid.redraw();mTiles()},0)}
function sigReset(){MAP.gen=(MAP.gen||0)+1;MAP.busy=0;MAP.tiles=new Map();MAP.feat={};MAP.top={};SWsk=null;SWel={};SQ.length=0;
  if(SW&&SWok)SW.postMessage({t:'r',g:MAP.gen});if(SCORE)SCORE.reset()}
/* respaldo sin OffscreenCanvas: los cuadros se pintan aquí por turnos (como mucho 10 ms seguidos por fotograma) */
var SQ=[], SQt=0;
function sigQueue(el,co,done){SQ.push([el,co,done]);if(!SQt)SQt=(window.requestAnimationFrame||setTimeout)(sigRun)}
function sigRun(){SQt=0;var t0=Date.now();
  while(SQ.length&&Date.now()-t0<10){var q=SQ.shift();try{sigPaint(q[0],q[1])}catch(e){}if(q[2])q[2](null,q[0])}
  if(SQ.length)SQt=(window.requestAnimationFrame||setTimeout)(sigRun)}
function sigPaint(el,co){var d=Math.min(2,window.devicePixelRatio||1),w=Math.round(256*d);if(el.width!==w){el.width=w;el.height=w}
  var ctx=el._c2||(el._c2=el.getContext('2d',{willReadFrequently:true}));if(!ctx||!SCORE)return 0;el._n=SCORE.paint(ctx,{z:co.z,x:co.x,y:co.y,d:d});return el._n}
function sigWMsg(e){var m=e.data;
  if(m.t==='f'){var f=SWcb[m.q];delete SWcb[m.q];if(f)f(m);return}
  if(m.t==='n'){var g=SWcb.n;delete SWcb.n;if(g)g(m.st);return}
  if(m.t!=='p')return;
  var el=SWel[m.id];delete SWel[m.id];
  if(!el||el._id!==m.id){if(m.bmp&&m.bmp.close)m.bmp.close();return}
  el._id=0;var done=el._done;el._done=null;
  if(m.err){if(!SCORE){SCORE=SIGCORE();SCORE.style(SIG_ST)}if(done)done(null,el);return}
  var d=Math.min(2,window.devicePixelRatio||1),w=Math.round(256*d);
  try{if(el.width!==w){el.width=w;el.height=w}
    var ctx=el._ctx||(el._ctx=el.getContext('bitmaprenderer')||el.getContext('2d'));
    if(ctx.transferFromImageBitmap)ctx.transferFromImageBitmap(m.bmp||null);
    else{ctx.setTransform(1,0,0,1,0,0);ctx.clearRect(0,0,w,w);if(m.bmp){ctx.drawImage(m.bmp,0,0);m.bmp.close()}}}catch(er){if(m.bmp&&m.bmp.close)m.bmp.close()}
  el._n=m.n;if(done)done(null,el)}
/* los que se dibujan encima (fincas y seleccionado) no se pintan en los cuadros; al hilo solo se manda cuando cambia */
function sigSkipSync(o){var s=Object.keys(o).sort().join('|');if(s===SWsk)return false;SWsk=s;if(SWok&&SW)SW.postMessage({t:'s',s:o});else if(SCORE)SCORE.skip(o);return true}
function sigReq(el,co,done){
  if(!SWok){sigQueue(el,co,done);return}
  if(el._id){delete SWel[el._id];SW.postMessage({t:'x',id:el._id})}
  var id=++SWn;el._id=id;el._co=co;if(done)el._done=done;SWel[id]=el;
  SW.postMessage({t:'p',id:id,z:co.z,x:co.x,y:co.y,d:Math.min(2,window.devicePixelRatio||1)})}
/* vuelve a pintar los cuadros que tocan los trozos de keys ({'x,y':1}); sin keys, todos */
function sigRepaint(keys){var g=MAP.grid;if(!g||!g._map||!g._tiles)return;
  Object.keys(g._tiles).forEach(function(id){var t=g._tiles[id],co=t.coords;
    if(keys){var k=Math.pow(2,co.z-MZ),px=Math.floor(co.x/k),py=Math.floor(co.y/k),nx=Math.max(1,Math.ceil(1/k)),hit=false;
      for(var i=-1;i<=nx&&!hit;i++)for(var j=-1;j<=nx&&!hit;j++)if(keys[(px+i)+','+(py+j)])hit=true;
      if(!hit)return}
    sigReq(t.el,co,null)})}
/* Leaflet se carga al abrir el mapa: la clase se hace entonces */
var SigGridC=null;function SigGrid(){if(!SigGridC)SigGridC=L.GridLayer.extend({options:{pane:'sigp',minZoom:14,maxZoom:20,className:'sigcuad',updateWhenZooming:false},
  createTile:function(co,done){var el=document.createElement('canvas');el.width=el.height=Math.round(256*Math.min(2,window.devicePixelRatio||1));sigReq(el,co,done);return el}});return new SigGridC()}
/* geometrías: lo que llega del hilo (o del respaldo) se guarda una vez por recinto en MAP.feat */
function sigFeats(k,m){var c=m.c,o=0,nu=0;
  m.nu.forEach(function(f){var P=f.s.map(function(pl){return pl.map(function(n){var ring=new Array(n);for(var i=0;i<n;i++,o+=2)ring[i]=[c[o],c[o+1]];return ring})});
    var g=f.m?{type:'MultiPolygon',coordinates:P}:{type:'Polygon',coordinates:P[0]},bb=[Infinity,Infinity,-Infinity,-Infinity];
    P.forEach(function(pl){pl[0].forEach(function(q){if(q[0]<bb[0])bb[0]=q[0];if(q[1]<bb[1])bb[1]=q[1];if(q[0]>bb[2])bb[2]=q[0];if(q[1]>bb[3])bb[3]=q[1]})});
    if(!MAP.feat[f.ref]){MAP.feat[f.ref]={p:f.p,g:g,bb:bb,n:0};nu++}
    if(S.recintos[f.ref]&&!GEO[f.ref])geoPonG(f.ref,g)});
  m.refs.forEach(function(r){var x=MAP.feat[r];if(x)x.n++});
  return nu}
function sigSuelta(ks){var fu={};ks.forEach(function(k){var rec=MAP.tiles.get(k);if(!rec)return;MAP.tiles.delete(k);fu[k]=1;
    (rec.refs||[]).forEach(function(r){var x=MAP.feat[r];if(x&&--x.n<=0)delete MAP.feat[r]})});
  if(SWok&&SW)SW.postMessage({t:'d',ks:ks});else if(SCORE)SCORE.drop(ks);
  return fu}
var tX=function(lon){return Math.floor((lon+180)/360*MNZ)}, tY=function(lat){var r=lat*Math.PI/180;return Math.floor((1-Math.log(Math.tan(r)+1/Math.cos(r))/Math.PI)/2*MNZ)};
function sigCarga(x,y){var k=x+','+y;if(MAP.tiles.has(k)){MAP.tiles.get(k).t=Date.now();return}
  var rec={t:Date.now(),refs:null},gen=MAP.gen||0,hilo=false;MAP.tiles.set(k,rec);MAP.busy++;
  sigArranca().then(function(){return fetch('https://sigpac.mapa.gob.es/vectorsdg/vector/recinto@3857/'+MZ+'.'+x+'.'+(MNZ-1-y)+'.geojson')})
    .then(function(r){if(!r.ok)throw new Error(r.status);hilo=SWok;return hilo?r.arrayBuffer():r.json()})
    .then(function(j){if((MAP.gen||0)!==gen)return null;
      if(hilo)return new Promise(function(res,rej){var q=++SWq;SWcb[q]=function(m){if(m.err)rej(new Error('hilo'));else res(m)};SW.postMessage({t:'a',q:q,k:k,b:j,g:gen},[j])});
      return SCORE.add(k,j)})
    .then(function(m){if(!m||(MAP.gen||0)!==gen||MAP.tiles.get(k)!==rec)return;
      rec.refs=m.refs;sigFeats(k,m);MAP.fail=0;var o={};o[k]=1;mTopSync(true);sigRepaint(o)})
    .catch(function(e){if((MAP.gen||0)!==gen)return;if(MAP.tiles.get(k)===rec)MAP.tiles.delete(k);if(!(e&&e.message==='hilo')){MAP.fail=1;mHint('El SIGPAC no responde ahora: prueba en un rato.')}})
    .then(function(){if((MAP.gen||0)===gen)MAP.busy--})}
function refreshSigpac(){var m=MAP.m;if(!m||!MAP.grid)return;var z=m.getZoom();
  if(z<14){if(m.hasLayer(MAP.grid))m.removeLayer(MAP.grid);if(m.hasLayer(MAP.lay))m.removeLayer(MAP.lay);mHint('Acércate un poco más para ver los recintos.');return}
  var b=m.getBounds(),x0=tX(b.getWest()),x1=tX(b.getEast()),y0=tY(b.getNorth()),y1=tY(b.getSouth()),want=[];
  for(var x=x0;x<=x1;x++)for(var y=y0;y<=y1;y++)want.push([x,y]);
  if(want.length>80){mHint('Acércate un poco más para ver los recintos.');return}
  if(!m.hasLayer(MAP.grid))MAP.grid.addTo(m);if(!m.hasLayer(MAP.lay))MAP.lay.addTo(m);
  mHint(MAP.fail?'El SIGPAC no responde ahora: prueba en un rato.':'');
  want.forEach(function(w){sigCarga(w[0],w[1])});
  /* tope: los trozos de unas tres pantallas; se sueltan los más viejos que no se ven (solo los ya leídos) */
  var sz=m.getSize(),tmax=Math.max(40,3*(Math.ceil(sz.x/256)+1)*(Math.ceil(sz.y/256)+1));MAP.tmax=tmax;
  if(MAP.tiles.size>tmax){var keep={},olds=[],ks=[];want.forEach(function(w){keep[w[0]+','+w[1]]=1});
    MAP.tiles.forEach(function(rec,k){if(!keep[k]&&rec.refs)olds.push([k,rec.t])});olds.sort(function(p,q){return p[1]-q[1]});
    while(MAP.tiles.size-ks.length>tmax&&olds.length)ks.push(olds.shift()[0]);
    if(ks.length){var fu=sigSuelta(ks);mTopSync(true);sigRepaint(fu)}}}
function mTiles(){refreshSigpac()}
/* tocar: cuenta propia sobre las geometrías guardadas; primero los de encima */
function inRing(pt,r){var x=pt[0],y=pt[1],ins=false;for(var i=0,j=r.length-1;i<r.length;j=i++){var xi=r[i][0],yi=r[i][1],xj=r[j][0],yj=r[j][1];if(((yi>y)!==(yj>y))&&(x<(xj-xi)*(y-yi)/(yj-yi)+xi))ins=!ins}return ins}
function inGeo(pt,g){var P=g.type==='Polygon'?[g.coordinates]:g.type==='MultiPolygon'?g.coordinates:[];return P.some(function(pl){var n=0;pl.forEach(function(r){if(inRing(pt,r))n++});return n%2===1})}
function hitAt(ll){var pt=[ll.lng,ll.lat],c=function(ref){var x=MAP.feat[ref];return x&&pt[0]>=x.bb[0]&&pt[0]<=x.bb[2]&&pt[1]>=x.bb[1]&&pt[1]<=x.bb[3]&&inGeo(pt,x.g)?{ref:ref,p:x.p,g:x.g}:null},h=null;
  Object.keys(MAP.top).some(function(r){return (h=c(r))});if(h)return h;
  for(var r in MAP.feat){if((h=c(r)))return h}return null}
/* encima, con sus colores de ahora: los recintos de esta finca, los de otras fincas y el seleccionado */
function mTopGeo(ref){var x=MAP.feat[ref];if(x)return x.g;var s=MAP.sel;if(s&&s.ref===ref){if(s.g)return s.g;if(s.wkt){var g=wktGeo(s.wkt);if(g)return g}}return null}
function mTopSync(noRepaint){if(!MAP.lay)return;var want={},ch={},r;
  for(r in S.recintos){var x=S.recintos[r];if(x&&finca(x.fincaId)&&MAP.feat[r])want[r]=1}
  if(MAP.sel&&mTopGeo(MAP.sel.ref))want[MAP.sel.ref]=1;
  for(r in MAP.top)if(!want[r]){MAP.lay.removeLayer(MAP.top[r]);delete MAP.top[r];ch[r]=1}
  for(r in want)if(!MAP.top[r]){var g=mTopGeo(r);var p=(MAP.feat[r]||{}).p||refProps(r);var l=L.geoJSON({type:'Feature',properties:p,geometry:g},{style:mEstilo,interactive:false,renderer:MAP.lay.options.renderer});
    l.eachLayer(function(y){MAP.lay.addLayer(y);MAP.top[r]=y});ch[r]=1}
  MAP.lay.setStyle(mEstilo);if(MAP.sel&&MAP.top[MAP.sel.ref]&&MAP.top[MAP.sel.ref].bringToFront)MAP.top[MAP.sel.ref].bringToFront();
  var cam=sigSkipSync(want);
  if(cam&&!noRepaint){var keys={};Object.keys(ch).forEach(function(r){var x=MAP.feat[r];if(!x)return;for(var tx=tX(x.bb[0]);tx<=tX(x.bb[2]);tx++)for(var ty=tY(x.bb[3]);ty<=tY(x.bb[1]);ty++)keys[tx+','+ty]=1});sigRepaint(keys)}}
function refProps(ref){var a=String(ref).split(':').map(Number);return {provincia:a[0],municipio:a[1],agregado:a[2],zona:a[3],poligono:a[4],parcela:a[5],recinto:a[6]}}
function mRestyle(){mTopSync()}
/* para las pruebas: lo que hay guardado (aquí y en el hilo) */
function sigStats(){var o={tiles:MAP.tiles.size,feat:Object.keys(MAP.feat).length,top:Object.keys(MAP.top).length,hilo:SWok,tmax:MAP.tmax||0,cuadros:MAP.grid&&MAP.grid._tiles?Object.keys(MAP.grid._tiles).length:0};
  if(SWok&&SW)return new Promise(function(res){SWcb.n=function(st){o.w=st;res(o)};SW.postMessage({t:'n'})});o.w=SCORE?SCORE.stats():null;return Promise.resolve(o)}
function mSel(p,ll,g){var k=refKey(p);MAP.sel={ref:k,g:g||null,ha:(+p.dn_surface||0)/10000,uso:p.uso_sigpac||'',ll:ll?[Math.round(ll.lat*1e5)/1e5,Math.round(ll.lng*1e5)/1e5]:null};mRestyle();mSide();
  if(!MAP.info[k])sigInfo(k).then(function(o){MAP.info[k]=o;mZv(k);if(MAP.sel&&MAP.sel.ref===k)mSide()}).catch(function(){MAP.info[k]={err:1};if(MAP.sel&&MAP.sel.ref===k)mSide()});
  else if(!MAP.info[k].err)mZv(k)}
/* v3.2: zona vulnerable del recinto tocado (solo si el SIGPAC ha dado sus datos). 0 = preguntando */
function mZv(k){if(MAP.zvn[k]!==undefined&&!(MAP.zvn[k]&&MAP.zvn[k].err))return;var r=S.recintos[k];if(r&&r.zvn){MAP.zvn[k]=r.zvn;return}MAP.zvn[k]=0;
  sigNitr(k).then(function(o){MAP.zvn[k]=o;var r2=S.recintos[k];if(r2&&!r2.zvn){r2.zvn=o;saveRaw()}},function(){MAP.zvn[k]={err:1}}).then(function(){if(MAP.sel&&MAP.sel.ref===k)mSide()})}
function mZvTxt(k){var r=S.recintos[k],z=r&&r.zvn?r.zvn:MAP.zvn[k];if(z===0)return 'Consultando…';if(!z)return '';if(z.err)return 'No he podido comprobarlo';return zvTxt(z)}
function mSide(){var el=$('#mside');if(!el)return;var st0=MAP.sel?MAP.sel.ref:'';if(el.dataset.ref!==st0){el.dataset.ref=st0;setTimeout(function(){el.scrollTop=0},0)}var f=mFin(),rs=f?fincaRecs(f.id):[],s=MAP.sel,h='';
  if(!s){h='<div class="k">Recintos</div><p class="small">Toca un recinto en el mapa para ver sus datos y añadirlo'+(f?' a la finca. Los de esta finca se ven en verde.':'. La finca se crea al añadir el primero.')+'</p>'}
  else{var o=MAP.info[s.ref]||{}, r=S.recintos[s.ref], mine=f&&r&&r.fincaId===f.id, oth=r&&!mine?finca(r.fincaId):null, uso=o.uso_sigpac||s.uso, ha=o.superficie!=null?+o.superficie:s.ha;
    var dl=function(k,v){return v===''||v==null?'':'<div><dt>'+k+'</dt><dd>'+v+'</dd></div>'};
    h='<div class="k">Recinto SIGPAC</div><h3>'+esc(refLabel(s.ref))+'</h3><dl class="mdl">'+dl('Uso',esc(usoTxt(uso)))+dl('Superficie',ha2(ha)+' ha')+
      (o.err?'<div class="small">No he podido leer el resto de datos del SIGPAC.</div>':!o.provincia?'<div class="small">Leyendo los datos del SIGPAC…</div>':
        dl('Pendiente media',o.pendiente_media!=null?trTxt(+o.pendiente_media,1)+' %':'')+dl('Coef. de regadío',o.coef_regadio!=null?trTxt(+o.coef_regadio,0)+' %':'')+dl('Admisibilidad',o.admisibilidad!=null?trTxt(+o.admisibilidad,0)+' %':'')+dl('Altitud',+o.altitud?nf(+o.altitud,0)+' m':'')+dl('Incidencias',o.incidencias?esc(o.incidencias):'')+dl('Zona vulnerable (SIGPAC)',esc(mZvTxt(s.ref))))+'</dl>';
    if(mine)h+='<div class="okst">'+ic('check')+'Ya está en esta finca</div><div class="acts"><button class="btn sm danger lk" data-act="ap-m-del">Quitarlo de la finca</button></div>';
    else if(oth)h+='<div class="note">Está en la finca «'+esc(oth.nombre)+'».</div>';
    else{/* punto 32: el cultivo y el riego del último recinto añadido */
      var ml=UI.mLast||null, rg=ml&&ml.r?ml.r:(+o.coef_regadio>0?'regadio':'secano'), cu=ml&&ml.c!=null?ml.c:cultDeUso(uso);
      h+='<div class="mf">'+fld('Cultivo','<select id="m-cult" class="in">'+cultOpts(cu)+'</select>')+fld('Riego','<select id="m-rg" class="in"><option value="secano">Secano</option><option value="regadio"'+(rg==='regadio'?' selected':'')+'>Regadío</option></select>')+'</div>'+
        (ml?'<div class="small">Puestos como el último recinto que añadiste.</div>':'')+
        '<div class="acts"><button class="btn p wide" data-act="ap-m-add">'+ic('plus')+'Añadir a la finca</button></div>'}}
  h+='<div class="msum">'+(f?'<b>'+rs.length+(rs.length===1?' recinto':' recintos')+'</b> en «'+esc(f.nombre)+'» · '+esc(haTxt(rs)):'Aún no hay finca: se crea con el primer recinto.')+'</div>';
  el.innerHTML=h}
function mAdd(){var f=mFin(),s=MAP.sel;if(!s)return;var o=MAP.info[s.ref]||{};
  undoPush();
  if(!f){if(S.recintos[s.ref]&&finca(S.recintos[s.ref].fincaId)){toast('Ese recinto ya está en otra finca.');return}f=fincaNueva('Mi finca');UI.v='finca';UI.fid=f.id;UI.step='fin';MAP.nueva=false}
  var x={ref:s.ref,ha:o.superficie!=null?+o.superficie:s.ha,cultivo:($('#m-cult')||{}).value||'',riego:($('#m-rg')||{}).value||'secano',uso:s.uso,ll:s.ll};
  var R=addRecs(f,[x]);if(!R.added){toast('Ese recinto ya está en otra finca.');return}
  UI.mLast={c:x.cultivo,r:x.riego};saveUI();if(s.g&&!GEO[s.ref])geoPonG(s.ref,s.g);
  var r=S.recintos[s.ref];if(o.provincia)sigAplica(r,o);
  /* v3.2: la zona vulnerable ya leída al tocarlo; si no, se pregunta ahora */
  var zk=MAP.zvn[s.ref];if(zk&&zk.p!=null)r.zvn=zk;else if(o.provincia)zvLee(s.ref).then(function(z){if(z){saveRaw();if(!$('#mapa').hidden)mSide();else render()}});if(!r.ll&&s.ll)r.ll=s.ll;pendAPlanes(f);climaAuto(f);save();mRestyle();mSide();toast('Recinto '+refLabel(s.ref)+' añadido','Deshacer',function(){undoDo();mRestyle();mSide()});
  if(!o.provincia)sigLeer(f,[s.ref])}
function mDel(){var f=mFin(),s=MAP.sel;if(!f||!s)return;recQuitar(f,[s.ref]);mRestyle();mSide();toast('Recinto quitado','Deshacer',function(){undoDo();mRestyle();mSide()})}
function mGo(t){t=String(t||'').trim();if(!t)return;var ref=refParse(t),res=$('#m-res');
  if(ref){toast('Buscando el recinto…');sigInfo(ref).then(function(o){var g=wktGeo(o.wkt);if(!g||!MAP.m)throw 1;toastOff();var lay=L.geoJSON(g);MAP.m.fitBounds(lay.getBounds().pad(0.6),{maxZoom:18});MAP.info[ref]=o;mSel({provincia:o.provincia,municipio:o.municipio,agregado:o.agregado,zona:o.zona,poligono:o.poligono,parcela:o.parcela,recinto:o.recinto,dn_surface:(+o.superficie||0)*10000,uso_sigpac:o.uso_sigpac},null);MAP.sel.wkt=o.wkt;var ll=wktLL(o.wkt);if(ll)MAP.sel.ll=ll;mRestyle()})
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

/* ───── v3: contornos de los recintos (se guardan aparte de los datos, para no cambiar el formato de las copias) ───── */
var GEO={}, GEOT={}, GEOS=null;
(function(){try{var o=JSON.parse(localStorage.getItem(KEYGEO)||'{}');if(o&&typeof o==='object')GEO=o}catch(e){}})();
function geoRed(g){var rd=function(c){return typeof c[0]==='number'?[Math.round(c[0]*1e6)/1e6,Math.round(c[1]*1e6)/1e6]:c.map(rd)};
  var cut=function(ring){if(ring.length<=240)return ring;var st=Math.ceil(ring.length/240),o=ring.filter(function(p,i){return i%st===0});o.push(ring[ring.length-1]);return o};
  if(!g||!g.coordinates)return null;var c=g.type==='Polygon'?g.coordinates.map(cut):g.type==='MultiPolygon'?g.coordinates.map(function(p){return p.map(cut)}):null;return c?{type:g.type,coordinates:rd(c)}:null}
function geoSave(){if(GEOS)return;GEOS=setTimeout(function(){GEOS=null;var o={};Object.keys(GEO).forEach(function(k){if(S.recintos[k])o[k]=GEO[k]});try{localStorage.setItem(KEYGEO,JSON.stringify(o))}catch(e){}},300)}
function geoPonG(ref,g){g=geoRed(g);if(!g)return;GEO[ref]=g;geoSave();aerLuego()}
function geoPon(ref,wkt){geoPonG(ref,wktGeo(wkt))}
/* pide al SIGPAC los contornos que faltan (una vez por sesión y recinto) */
var GEOQ=[], GEOB=0;
function geoPide(f){if(navigator.onLine===false)return;fincaRecs(f.id).forEach(function(r){if(!GEO[r.ref]&&!GEOT[r.ref]){GEOT[r.ref]=1;GEOQ.push(r.ref)}});geoSigue()}
function geoSigue(){while(GEOB<3&&GEOQ.length){var ref=GEOQ.shift();GEOB++;sigInfo(ref).catch(function(){}).then(function(){GEOB--;geoSigue()})}}

/* ───── v3: foto aérea de la finca con sus recintos (barra lateral, mesa de la derecha y lista de fincas) ─────
   Mapas fijos (sin arrastrar ni zoom). Sin contornos o sin internet, un esquema con el color de cada cultivo. */
var AER={}, AERT=null, AERL=false;
function aerLuego(){if(AERT)return;AERT=setTimeout(function(){AERT=null;aerMonta()},60)}
function aerEsq(rs,lab,N,ar){var H=Math.round(100*(ar>0.2&&ar<3?ar:0.7)*10)/10;var it=rs.map(function(r){return {r:r,a:Math.max(haC(r),0.05)}}).sort(function(a,b){return b.a-a.a}),out=[];
  var cut=function(L0,x,y,w,h){if(!L0.length)return;if(L0.length===1){out.push([L0[0].r,x,y,w,h]);return}var t=L0.reduce(function(a,q){return a+q.a},0),s=0,i=0;while(i<L0.length-1&&(s+L0[i].a)<=t/2){s+=L0[i].a;i++}if(i===0){s=L0[0].a;i=1}var fr=s/t,A=L0.slice(0,i),B=L0.slice(i);
    if(w>=h){cut(A,x,y,w*fr,h);cut(B,x+w*fr,y,w*(1-fr),h)}else{cut(A,x,y,w,h*fr);cut(B,x,y+h*fr,w,h*(1-fr))}};
  cut(it,4,4,92,H-8);
  return '<svg class="esq" viewBox="0 0 100 '+H+'" aria-hidden="true"><rect width="100" height="'+H+'" fill="#2c4a33"/>'+out.map(function(o){var r=o[0];return '<rect data-ref="'+esc(r.ref)+'" x="'+(o[1]+.8).toFixed(2)+'" y="'+(o[2]+.8).toFixed(2)+'" width="'+Math.max(o[3]-1.6,.5).toFixed(2)+'" height="'+Math.max(o[4]-1.6,.5).toFixed(2)+'" rx="1.2" fill="'+glOf(r.cultivo).c+'" class="'+(HL.indexOf(r.ref)>=0?'on':'')+'"/>'+(function(){var g=glOf(r.cultivo),z=Math.min(o[3],o[4])*.42;if(lab)z=Math.min(z,16);if(z<4)return '';var cx=o[1]+o[3]/2,cy=o[2]+o[4]/2-(lab?z*.35:0);
      return '<rect x="'+(cx-z*.62).toFixed(2)+'" y="'+(cy-z*.62).toFixed(2)+'" width="'+(z*1.24).toFixed(2)+'" height="'+(z*1.24).toFixed(2)+'" rx="'+(z*.3).toFixed(2)+'" fill="'+g.bg+'"/><svg x="'+(cx-z/2).toFixed(2)+'" y="'+(cy-z/2).toFixed(2)+'" width="'+z.toFixed(2)+'" height="'+z.toFixed(2)+'" viewBox="0 0 24 24">'+g.s+'</svg>'+(lab&&o[4]>z*1.9?'<text x="'+cx.toFixed(2)+'" y="'+(cy+z*.62+6).toFixed(2)+'" text-anchor="middle">'+N[r.ref]+' · '+ha2(haC(r))+' ha</text>':'')})()}).join('')+'</svg>'}
function aerSig(f,rs,lab){return f.id+'|'+(lab?1:0)+'|'+rs.map(function(r){return r.ref+':'+(r.cultivo||'')+':'+(GEO[r.ref]?1:0)+':'+ha2(haC(r))}).join(',')}
function aerQuita(A){if(A&&A.m){try{A.m.off();A.m.remove()}catch(e){}}}
function aerMonta(){var els=$$('.aer[data-aer]');if(!els.length)return;
  els.forEach(function(el,ix){var slot=el.dataset.aer,f=finca(el.dataset.fid);if(!f)return;var rs=fincaRecs(f.id),lab=!!el.dataset.lab,N=recNum(f),sig=aerSig(f,rs,lab),A=AER[slot];
    var con=rs.filter(function(r){return GEO[r.ref]}).length;
    if(rs.length&&con<rs.length)geoPide(f);
    var hasL=!!(window.L&&window.L.map), map=con&&hasL&&!(slot.indexOf('home-')===0&&ix>12);
    var tg=el.parentNode&&el.parentNode.querySelector('.tag');if(tg)tg.textContent=map?'Foto aérea con tus recintos':'Esquema de tus recintos';
    if(con&&!hasL&&!AERL){AERL=true;ensureMap().then(aerLuego,function(){})}
    if(!map){if(A){aerQuita(A);delete AER[slot]}el.classList.add('sq');el.innerHTML=rs.length?aerEsq(rs,lab,N,el.clientWidth?el.clientHeight/el.clientWidth:0):'<div class="aer0">Sin recintos</div>';return}
    if(A&&A.m&&A.fid===f.id){if(A.el!==el)el.replaceWith(A.el);if(A.sig!==sig)aerDibuja(A,f,rs,lab,N,sig);setTimeout(function(){try{A.m.invalidateSize({animate:false});if(A.b)A.m.fitBounds(A.b,{padding:A.pad,animate:false})}catch(e){}},0);return}
    if(A)aerQuita(A);
    el.classList.remove('sq');el.innerHTML='';
    var m=L.map(el,{zoomControl:false,attributionControl:slot==='desk',dragging:false,scrollWheelZoom:false,doubleClickZoom:false,boxZoom:false,keyboard:false,touchZoom:false,tap:false,zoomSnap:0.25,inertia:false,fadeAnimation:false,zoomAnimation:false});
    if(m.attributionControl)m.attributionControl.setPrefix(false);
    L.tileLayer(IGNW+'pnoa-ma?service=WMTS&request=GetTile&version=1.0.0&format=image/jpeg&layer=OI.OrthoimageCoverage&style=default&tilematrixset=GoogleMapsCompatible&TileMatrix={z}&TileRow={y}&TileCol={x}',{maxNativeZoom:19,maxZoom:20,attribution:'PNOA © IGN · SIGPAC'}).addTo(m);
    A=AER[slot]={el:el,m:m,fid:f.id,pl:{},pad:slot==='desk'?[26,26]:[6,6]};
    aerDibuja(A,f,rs,lab,N,sig)});
  aerHL()}
function aerDibuja(A,f,rs,lab,N,sig){var m=A.m;if(A.lay)m.removeLayer(A.lay);A.lay=L.featureGroup().addTo(m);A.pl={};A.sig=sig;
  rs.forEach(function(r){var g=GEO[r.ref];if(!g)return;var l=L.geoJSON(g,{interactive:false,style:function(){return aerEst(r,false)}});l.addTo(A.lay);A.pl[r.ref]={l:l,r:r};
    if(lab)l.bindTooltip(N[r.ref]+' · '+ha2(haC(r))+' ha',{permanent:true,direction:'center',className:'aerl',interactive:false})});
  try{A.b=A.lay.getBounds();if(A.b.isValid())m.fitBounds(A.b,{padding:A.pad,animate:false});else A.b=null}catch(e){A.b=null}}
function aerEst(r,on){return on?{color:'#e9c838',weight:3,opacity:1,fillColor:glOf(r.cultivo).c,fillOpacity:.55}:{color:'#ffffff',weight:1.3,opacity:.95,fillColor:glOf(r.cultivo).c,fillOpacity:.42}}
function aerHL(){Object.keys(AER).forEach(function(k){var A=AER[k];if(!A.pl||k.indexOf('home-')===0)return;Object.keys(A.pl).forEach(function(ref){var o=A.pl[ref];o.l.setStyle(aerEst(o.r,HL.indexOf(ref)>=0))})});
  $$('.aer svg.esq rect[data-ref]').forEach(function(e){e.classList.toggle('on',HL.indexOf(e.dataset.ref)>=0)})}

/* ───── acciones (versión 3: tres pasos, sin confirm(), el foco no se pierde) ───── */
function recQuitar(f,refs){var p=pabCur(f),man=p&&udMan(f,p);undoPush();refs.forEach(function(ref){delete S.recintos[ref];delete SEL[ref];delete RMR[ref];(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){u.refs=(u.refs||[]).filter(function(x){return x!==ref});if(u.rc)delete u.rc[ref];if(u.rp===ref)u.rp=u.refs[0]||''})})});
  if(p)udTrasCambio(f,p,man);save();render()}
function fincaBorrar(f){undoPush();Object.keys(S.recintos).forEach(function(k){if(S.recintos[k].fincaId===f.id)delete S.recintos[k]});S.fincas=S.fincas.filter(function(x){return x.id!==f.id});delete AER['home-'+f.id];if(UI.fid===f.id){UI.v='inicio';UI.fid=null}sheetClose();save();saveUI();window.__scrollTop=true;render();undoToast('Finca «'+f.nombre+'» borrada')}
/* punto 38: se pregunta dentro de la página */
function fincaBorrarPide(f){var rs=fincaRecs(f.id),np=(f.pab||[]).length;
  sheet('<h2>¿Borrar la finca «'+esc(f.nombre)+'»?</h2><p>Se borran sus '+nRec(rs.length)+(np?' y '+(np===1?'su plan de abonado':'sus '+np+' planes de abonado'):'')+'. Justo después puedes deshacerlo.</p><div class="acts"><button class="btn p danger" data-act="ap-fdel-ok" data-id="'+esc(f.id)+'" autofocus>'+ic('trash')+'Sí, borrar la finca</button><button class="btn" data-act="ap-close">No, dejarla</button></div>')}
function planBorrarPide(f,p){sheet('<h2>¿Borrar el plan de la campaña '+esc(pabCamp(p.y))+'?</h2><p>Se borra el programa, los datos y el detalle de este plan. Los recintos de la finca se quedan. Justo después puedes deshacerlo.</p><div class="acts"><button class="btn p danger" data-act="ap-pdel-ok" autofocus>'+ic('trash')+'Sí, borrar el plan</button><button class="btn" data-act="ap-close">No, dejarlo</button></div>')}
function irPaso(s){UI.step=stepOf(s);UI.det=null;saveUI();window.__scrollTop=true;render()}
function focoEn(sel,center){setTimeout(function(){var el=null;try{el=sel?$(sel):null}catch(e){}if(!el)return;el.scrollIntoView({block:center===false?'start':'center',behavior:'smooth'});flash(el);if(el.matches('input,select,textarea,button'))try{el.focus({preventScroll:true})}catch(e){}},40)}
function irFalta(s){if(!CTX)return;var L=CTX.F,x=s?L.filter(function(y){return y.s===s})[0]:L[0];
  if(!x){if(s==='dat'||!L.length){toast(CTX.G.k==='ok'?'No falta nada':CTX.G.s||CTX.G.t);return}x=L[0]}
  if(x.s==='det'){UI.step='plan';UI.det=x.det;saveUI();render();var d=$('#det-'+x.det);if(d){var e=x.sel?$(x.sel,d):null;if(!e)e=$('input:not([type=checkbox]),select,button.chip,input[type=checkbox]',d);d.scrollIntoView({block:'start',behavior:'smooth'});flash(d);if(e)setTimeout(function(){try{e.focus({preventScroll:true})}catch(er){}},60)}return}
  if(x.s==='plan'){var m=/^\[data-p="ud\.(\d+)\.pg\.rows\.(\d+)\.(\w+)"\]$/.exec(x.sel||''),p=CTX.p;if(m&&p){var q=((p.ud[+m[1]].pg||{}).rows||[])[+m[2]];if(q&&m[3]!=='d')RM[q.id]=true}}
  UI.step=x.s==='plan'?'plan':x.s;if(x.s==='fin'&&/tit\./.test(x.sel)&&!/tit\.(n|nif)"/.test(x.sel))UI.titOpen=true;saveUI();render();
  focoEn(x.sel||(x.u!=null?'#u'+x.u:''))}
function flash(el){el.classList.remove('flash');void el.offsetWidth;el.classList.add('flash');setTimeout(function(){el.classList.remove('flash')},1600)}
/* punto 34: PDF y Excel con un clic */
function outGo(f,p,kind){var go=function(){try{var b=kind==='pdf'?buildPabPdf(f,p):buildPabXlsx(f,p);dlDirecto(b,'plan-de-abonado-'+safeName(f.nombre)+'-'+campFile(+p.y+1)+'.'+(kind==='pdf'?'pdf':'xlsx'),kind==='pdf'?'PDF':'Excel',kind==='pdf')}catch(e){toast('No he podido crear el plan de abonado.')}};
  if(kind==='pdf'){toast('Preparando el PDF…');ensurePdf().then(function(){toastOff();go()},function(){pdfNoListo()})}else go()}
function pdfGo(f,p){outGo(f,p,'pdf')}
function abrirFinca(id){var f=finca(id);if(!f)return;UI=Object.assign(UI,{v:'finca',fid:id,det:null});SEL={};RM={};RMR={};RMODE={};planAuto(f);var c=ctxOf(f),hasPg=c.p&&(c.p.ud||[]).some(function(u){return ((u.pg||{}).rows||[]).length});
  var x=c.F.filter(function(y){return y.s==='fin'||y.s==='dat'})[0];UI.step=!fincaRecs(f.id).length?'fin':(x?x.s:(hasPg?'plan':'dat'));saveUI();window.__scrollTop=true;render();zvComprobar(f)}
function copiasHoja(){var h='<h2>Copias</h2><p class="small">Lo que metes se guarda en este navegador. Una copia en archivo te protege si se borran los datos del navegador o cambias de ordenador.</p><div class="opts two">';
  if(fsCan())h+=FS.h?'<button class="opt" data-act="ap-fs-stop">'+ic('close')+'<b>Dejar de guardar en «'+esc(FS.name)+'»</b><span>Se sigue guardando en el navegador.</span></button>':'<button class="opt" data-act="ap-fs-pick">'+ic('save')+'<b>Guardar solo en un archivo</b><span>Eliges el archivo una vez y cada cambio se guarda ahí (recomendado).</span></button>';
  h+='<button class="opt" data-act="ap-copia">'+ic('download')+'<b>Descargar una copia ahora</b><span>Un archivo con todas tus fincas.</span></button><button class="opt" data-act="ap-abrir">'+ic('folder')+'<b>Abrir una copia</b><span>Recuperar fincas de un archivo.</span></button></div><div class="acts"><button class="btn" data-act="ap-close">Cerrar</button></div>';sheet(h,'wide')}
/* v3.2: al terminar una fila con un abono que no es del listado, se guarda en «Mis abonos» (sin duplicar; si cambia la riqueza, se actualiza y se avisa) */
function filaPorId(p,id){var o=null;((p&&p.ud)||[]).forEach(function(u,i){((u.pg||{}).rows||[]).forEach(function(q,j){if(q.id===id)o={u:u,q:q,i:i,j:j}})});return o}
function abRq(a){return trTxt(+a.rn||0,1,1)+'-'+trTxt(+a.rp||0,1,1)+'-'+trTxt(+a.rk||0,1,1)+(a.org?' orgánico':'')}
function abGuardaMio(q){var n=String(q.n||'').trim(),k=abKey(n);if(!k||catEsTipo(n)||!(+q.rn||+q.rp||+q.rk||q.org===1))return;
  var L=S.x.abonos=S.x.abonos||[],a=L.filter(function(x){return abKey(x.n)===k})[0],nu={rn:+q.rn||0,rp:+q.rp||0,rk:+q.rk||0,org:q.org===1?1:0};
  if(!a){a={n:n,rn:nu.rn,rp:nu.rp,rk:nu.rk};if(nu.org)a.org=1;if(abHas(q.pz))a.pz=+q.pz;L.push(a);save();toast('«'+n+'» guardado en «Mis abonos»');return}
  var ch=false;if(abHas(q.pz)&&+a.pz!==+q.pz){a.pz=+q.pz;ch=true}
  if(['rn','rp','rk'].some(function(x){return Math.abs((+a[x]||0)-nu[x])>1e-9})||(a.org?1:0)!==nu.org){var antes=abRq(a);a.rn=nu.rn;a.rp=nu.rp;a.rk=nu.rk;if(nu.org)a.org=1;else delete a.org;save();toast('Actualizada la riqueza de «'+a.n+'» en «Mis abonos»: antes '+antes+', ahora '+abRq(a));return}
  if(ch)save()}
function misHoja(){var L=S.x.abonos||[];
  var h='<h2>Mis abonos</h2><p class="small">Tus abonos con nombre comercial y la riqueza de su etiqueta (y, si quieres, el precio). Se guardan solos al poner «Mi abono» en el programa; aquí puedes corregirlos o añadir más. El precio solo se ve en pantalla.</p><div class="xs"><table class="mini ma"><thead><tr><th>Nombre</th><th class="r">N %</th><th class="r">P₂O₅ %</th><th class="r">K₂O %</th><th class="c">Org.</th><th class="r">€/kg</th><th class="r">Densidad (kg/L)</th><th></th></tr></thead><tbody>'+
    L.map(function(a,i){var c=function(k,ph,n){return '<td><input class="in ma-in'+(n?' n':'')+'" data-i="'+i+'" data-k="'+k+'" value="'+esc(a[k]==null?'':(typeof a[k]==='number'?abEd(a[k]):a[k]))+'" placeholder="'+(ph||'')+'"'+(n?' inputmode="decimal"':'')+' autocomplete="off"></td>'};
      return '<tr>'+c('n','Nombre del abono')+c('rn','0',1)+c('rp','0',1)+c('rk','0',1)+'<td class="c"><input type="checkbox" class="ma-in" data-i="'+i+'" data-k="org"'+(a.org?' checked':'')+'></td>'+c('pz','0,00',1)+c('dens','',1)+'<td><button class="btn q ib danger" data-act="ap-ma-del" data-i="'+i+'" aria-label="Quitar">'+ic('close')+'</button></td></tr>'}).join('')+
    (L.length?'':'<tr><td colspan="8" class="small empty">Aún no hay abonos tuyos.</td></tr>')+'</tbody></table></div><div class="acts"><button class="btn" data-act="ap-ma-add">'+ic('plus')+'Añadir abono</button><button class="btn p" data-act="ap-ma-ok">Hecho</button></div>';
  sheet(h,'wide')}
function campHoja(f){var ys={};(f.pab||[]).forEach(function(x){ys[+x.y]=1});var y0=pabY0(),L=[y0-1,y0,y0+1,y0+2].filter(function(y){return !ys[y]});
  sheet('<h2>Plan de otra campaña</h2><p class="small">Las campañas van de noviembre a octubre. Se copia lo que no cambia de un año a otro del plan más reciente: suelo, agua, zona vulnerable y programa.</p>'+(L.length?'<div class="opts two">'+L.map(function(y){return '<button class="opt" data-act="ap-newplan" data-y="'+y+'">'+ic('cal')+'<b>Campaña '+esc(pabCamp(y))+'</b><span>De noviembre de '+y+' a octubre de '+(y+1)+'</span></button>'}).join('')+'</div>':'<p>Ya están hechos los planes de las campañas cercanas.</p>')+'<div class="acts"><button class="btn" data-act="ap-close">Cancelar</button></div>','wide')}
function popAt(btn,html,cls){var p=$('#pop');p.className=cls||'';p.innerHTML=html;p.hidden=false;var r=btn.getBoundingClientRect(),w=p.offsetWidth,hh=p.offsetHeight,x=Math.min(window.innerWidth-w-8,Math.max(8,r.left+r.width/2-w/2)),y=r.bottom+6;if(y+hh>window.innerHeight-8)y=Math.max(8,r.top-hh-6);p.style.left=x+'px';p.style.top=y+'px';POPB=btn}
var POPB=null;function popClose(){var p=$('#pop');if(p){p.hidden=true;p.innerHTML=''}POPB=null}
/* calendario propio (fecha siempre dd/mm/aaaa) */
var CAL={inp:null,y:0,m:0};
function calHTML2(){var y=CAL.y,m=CAL.m,d1=new Date(y,m,1),off=(d1.getDay()+6)%7,nd=new Date(y,m+1,0).getDate(),sel=pabIso(CAL.inp&&CAL.inp.value),hoy=hoyIso(),h='<div class="calp"><div class="calh"><button type="button" data-act="ap-calm" data-d="-1" aria-label="Mes anterior">‹</button><b>'+MESES[m]+' '+y+'</b><button type="button" data-act="ap-calm" data-d="1" aria-label="Mes siguiente">›</button></div><div class="calg">'+['L','M','X','J','V','S','D'].map(function(x){return '<i>'+x+'</i>'}).join('');
  for(var k=0;k<off;k++)h+='<span></span>';for(var d=1;d<=nd;d++){var iso=y+'-'+('0'+(m+1)).slice(-2)+'-'+('0'+d).slice(-2);h+='<button type="button" data-act="ap-cald" data-v="'+iso+'" class="'+(iso===sel?'on':'')+(iso===hoy?' hoy':'')+'">'+d+'</button>'}
  return h+'</div><div class="acts"><button type="button" class="btn" data-act="ap-cald" data-v="'+hoy+'">Hoy</button><button type="button" class="btn q" data-act="ap-cald" data-v="">Borrar</button></div></div>'}

/* «Hacer el plan» (punto 15): propone el programa de todas las unidades y lleva a «Plan y PDF» */
function hacerPlan(f,p,todas){var n=0,sin=[];undoPush();
  (p.ud||[]).forEach(function(u){var C=pabCalc(f,u,p);if(C.ex||u.d0===1)return;var had=((u.pg||{}).rows||[]).length;if(had&&!todas)return;
    var ne=u.ne||{};if((ne.mo||'cos')==='cos'&&!(+ne.pr>0)||noExtr(u)){sin.push(u.n);return}
    var rows=propuesta(f,u,C);u.pg={src:'own',rows:rows};mitPropuesta(u,true);n++});
  p.m=Date.now();save();sheetClose();UI.step='plan';UI.det=null;saveUI();window.__scrollTop=true;render();
  undoToast(n?'Programa propuesto para '+(n===1?'1 unidad':n+' unidades')+'. Revísalo y cambia lo que quieras':'No había que proponer nada nuevo')}
function hacerPide(f,p){var con=(p.ud||[]).filter(function(u){var C=pabCalc(f,u,p);return !C.ex&&u.d0!==1&&((u.pg||{}).rows||[]).length});
  if(!con.length){hacerPlan(f,p,false);return}
  var nn=(p.ud||[]).filter(function(u){var C=pabCalc(f,u,p);return !C.ex&&u.d0!==1}).length;
  sheet('<h2>'+(con.length===1?'«'+esc(con[0].n)+'» ya tiene programa':con.length+' unidades ya tienen programa')+'</h2><p>¿Lo dejo como está o lo propongo otra vez con los datos de ahora?</p><div class="acts"><button class="btn p" data-act="ap-hacer-ok" data-m="nuevas" autofocus>Dejar lo que ya hay'+(nn>con.length?' y proponer el resto':'')+'</button><button class="btn" data-act="ap-hacer-ok" data-m="todas">Proponer todo otra vez</button><button class="btn q" data-act="ap-close">Cancelar</button></div>')}
function sinCosecha(f,p,u,i){var ne=u.ne||{};if((ne.mo||'cos')==='cos'&&!(+ne.pr>0)){toast('Falta la cosecha esperada de «'+u.n+'»: ponla aquí y propongo el programa.');UI.step='dat';saveUI();render();focoEn('[data-p="ud.'+i+'.ne.pr"]');return true}
  if(noExtr(u)){toast('Faltan las extracciones de «'+u.n+'»: escríbelas aquí y propongo el programa.');UI.step='dat';saveUI();render();focoEn('[data-p="ud.'+i+'.ne.ex.n"]');return true}return false}
function nuevaFila(f,p,u,i){u.pg=u.pg||{src:'own',rows:[]};var rows=u.pg.rows||(u.pg.rows=[]),rg=u.rg==='regadio',q={id:uid(),mo:'',n:'',fo:rg?'Fertirrigación':'Suelo',eq:rg?'Cabezal de riego':'Abonadora'};
  rows.push(q);delete u.d0;RM[q.id]=true;p.m=Date.now();save();render();focoEn('[data-p="ud.'+i+'.pg.rows.'+(rows.length-1)+'.n"]')}
/* las medidas del anexo V y las restricciones se marcan una vez para toda la finca (punto 24) */
function todasUd(f,p,fn){var n=0;(p.ud||[]).forEach(function(u,i){if(pabCalc(f,u,p).ex)return;fn(u,i);n++});p.m=Date.now();return n}
/* v3.1 punto 7: como todasUd, pero si en «Aplicar a» hay una unidad elegida, solo esa */
function udsAplica(f,p,fn){var U=(p.ud||[]).map(function(u,i){return i}).filter(function(i){return !pabCalc(f,p.ud[i],p).ex}),S1=aplSel(p,U);if(S1.length===U.length)return todasUd(f,p,fn);S1.forEach(function(i){fn(p.ud[i],i)});p.m=Date.now();return S1.length}
function suCopia(p,src){(p.ud||[]).forEach(function(u){if(u.suDe===src.id&&u!==src)u.su=JSON.parse(JSON.stringify(src.su||{}))})}

document.addEventListener('click',function(ev){
  var pp=$('#pop'); if(pp&&!pp.hidden&&!pp.contains(ev.target)&&!(POPB&&POPB.contains(ev.target)))popClose();
  if(!(ev.target.closest&&ev.target.closest('.sgw,.sgl')))sgClose();
  var b=ev.target.closest&&ev.target.closest('[data-act]'); if(!b||b.disabled)return; var a=b.dataset.act, f=UI.v==='finca'?finca(UI.fid):null;
  if(a==='pa-del'){ev.stopImmediatePropagation();if(f&&pabCur(f))planBorrarPide(f,pabCur(f));return}
  if(a==='pa-row-add'){var p0=f&&pabCur(f),u0=p0&&p0.ud[+b.dataset.u];if(u0)nuevaFila(f,p0,u0,+b.dataset.u);return}
  if(/^pa-/.test(a)){if(!f)return;EXP.fid=f.id;var pp0=pabCur(f);if(pp0&&/^pa-ud-(new|del)$/.test(a))pp0.udMan=1;if(a==='pa-set'&&pp0&&/^ud\.\d+\.(rg|ck)$/.test(b.dataset.p||''))pp0.udMan=1;pabClick(b);return}
  if(a==='ap-help'){ev.preventDefault();var g=GLOS[b.dataset.k];if(g){if(POPB===b){popClose();return}popAt(b,'<b>'+esc(g[0])+'</b><p>'+esc(g[1])+'</p>'+(g[2]?'<div class="small">'+esc(g[2])+'</div>':''),'hlpp')}return}
  if(a==='ap-cal'){var inp=b.previousElementSibling;if(POPB===b){popClose();return}CAL.inp=inp;var iso=pabIso(inp.value)||hoyIso();CAL.y=+iso.slice(0,4);CAL.m=+iso.slice(5,7)-1;popAt(b,calHTML2(),'calpop');return}
  if(a==='ap-calm'){CAL.m+=+b.dataset.d;if(CAL.m<0){CAL.m=11;CAL.y--}if(CAL.m>11){CAL.m=0;CAL.y++}$('#pop').innerHTML=calHTML2();return}
  if(a==='ap-cald'){var ci=CAL.inp;popClose();if(!ci)return;ci.value=b.dataset.v?isoEs(b.dataset.v):'';ci.dispatchEvent(new Event('change',{bubbles:true}));return}
  if(a==='ap-close'){sheetClose();return}
  if(a==='ap-undo'){undoDo();return}
  if(a==='ap-open'){if(ev.target.closest('button')&&ev.target.closest('button')!==b)return;abrirFinca(b.dataset.id);return}
  if(a==='ap-home'){UI.v='inicio';UI.fid=null;UI.det=null;saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-step'){irPaso(b.dataset.s);return}
  if(a==='ap-next'){irPaso(STEPS[Math.min(2,stepIdx()+1)][0]);return}
  if(a==='ap-mano'){sheetClose();if(!f){UI.v='inicio';UI.fid=null}addRecHoja();return}
  if(a==='ap-mapa-nueva'){sheetClose();UI.v='inicio';UI.fid=null;saveUI();mapaAbrir();return}
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
  if(a==='ap-fdel-id'){ev.stopPropagation();var fd=finca(b.dataset.id);if(fd)fincaBorrarPide(fd);return}
  if(a==='ap-fdel-ok'){var fo=finca(b.dataset.id);if(fo)fincaBorrar(fo);return}
  if(a==='ap-mis'){misHoja();return}
  if(a==='ap-ma-add'){(S.x.abonos=S.x.abonos||[]).push({n:'',rn:0,rp:0,rk:0});saveRaw();misHoja();var ins=$$('#sheet .ma-in[data-k="n"]');if(ins.length)ins[ins.length-1].focus();return}
  if(a==='ap-ma-del'){undoPush();S.x.abonos.splice(+b.dataset.i,1);save();misHoja();return}
  if(a==='ap-ma-ok'){S.x.abonos=(S.x.abonos||[]).filter(function(x){return String(x.n||'').trim()});save();sheetClose();render();return}
  if(a==='ap-sort'){var k=b.dataset.k;UI.sort=(UI.sort===k?'-'+k:k);saveUI();render();return}
  if(a==='ap-m-close'){mapaCerrar();return}
  if(a==='ap-m-add'){mAdd();return}
  if(a==='ap-m-del'){mDel();return}
  if(a==='ap-m-loc'){mLoc();return}
  if(a==='ap-m-place'){mPlace(+b.dataset.i);return}
  if(a==='ap-addrec'){addRecHoja();return}
  if(a==='ap-ref-add'){refAdd();return}
  if(a==='ap-mapa'){mapaAbrir();return}
  if(a==='ap-prevpg'){PREV.big=+b.dataset.i;prevPinta();return}
  if(!f)return;
  var p=pabCur(f);
  if(a==='ap-compra-copy'){if(!p)return;var L=compra(f,p,CTX.Cs);var txt=L.map(function(x){return x.n+': '+(x.kg>=1000?nf(x.kg/1000,2)+' t':nf(x.kg,0)+' kg')}).join('\n');
    (navigator.clipboard?navigator.clipboard.writeText(txt):Promise.reject()).then(function(){toast('Lista copiada')},function(){toast('No he podido copiar la lista')});return}
  if(a==='ap-sig'){toast('Leyendo el SIGPAC…');sigLeer(f,fincaRecs(f.id).map(function(r){return r.ref}),function(ok,ko){toast('SIGPAC: '+ok+' recintos leídos'+(ko?', '+ko+' sin respuesta':''))});return}
  if(a==='ap-recdel'){recQuitar(f,[b.dataset.ref]);undoToast('Recinto '+refLabel(b.dataset.ref)+' quitado');return}
  if(a==='ap-recmas'){RMR[b.dataset.ref]=!RMR[b.dataset.ref];render();if(RMR[b.dataset.ref])focoEn('[data-r="'+b.dataset.ref+'"][data-k="variedad"]');return}
  if(a==='ap-bulk-del'){var ks=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id});if(!ks.length)return;recQuitar(f,ks);undoToast(ks.length+(ks.length===1?' recinto quitado':' recintos quitados'));return}
  if(a==='ap-bulk-none'){SEL={};render();return}
  if(a==='ap-copiaf'){copiaFinca(f);return}
  if(a==='ap-fdel'){fincaBorrarPide(f);return}
  if(a==='ap-newplan'){var y=+b.dataset.y;if((f.pab||[]).some(function(x){return +x.y===y})){toast('Ya hay un plan de la campaña '+pabCamp(y));return}undoPush();var n=planNuevo(f,y);sheetClose();UI.step='dat';save();saveUI();window.__scrollTop=true;render();undoToast('Plan '+pabCamp(n.y)+' creado');return}
  if(a==='ap-falta'){irFalta(b.dataset.s||'');return}
  if(a==='ap-pdel-ok'){if(!p)return;undoPush();f.pab=(f.pab||[]).filter(function(x){return x.id!==p.id});if(!f.pab.length)delete f.pab;sheetClose();save();UI.step=f.pab?UI.step:'fin';saveUI();render();undoToast('Plan de abonado borrado');return}
  if(a==='ap-sigcamp'){if(!p)return;var y1=+p.y+1,ex=(f.pab||[]).filter(function(x){return +x.y===y1})[0];if(ex){PAB.cur[f.id]=ex.id;render();toast('Plan de la campaña '+pabCamp(y1));return}
    undoPush();planNuevo(f,y1,p);UI.step='dat';save();saveUI();window.__scrollTop=true;render();undoToast('Plan '+pabCamp(y1)+' hecho a partir del de '+pabCamp(p.y)+': revisa la cosecha y el programa');return}
  if(a==='ap-pdf'||a==='ap-xlsx'){if(!p){toast('Primero mete los recintos');return}outGo(f,p,a==='ap-pdf'?'pdf':'xlsx');return}
  if(a==='ap-det'){var dk=b.dataset.k;UI.det=UI.det===dk?null:dk;UI.aplU=null;/* v3.1: «Aplicar a» empieza siempre en «Todas las unidades» */if(UI.step==='dat'&&b.closest('#app')&&!b.closest('.more'))UI.datMas=true;if(UI.det&&UI.step==='dat'&&!b.closest('.more')&&['cl','ne','ap','su','ag','ud'].indexOf(dk)<0)UI.step='plan';saveUI();render();if(UI.det){var de=$('#det-'+dk);if(de){var r0=de.getBoundingClientRect();if(r0.top<0||r0.top>window.innerHeight-120)de.scrollIntoView({block:'start',behavior:'smooth'});var fi=$('input:not([type=checkbox]),select',de);if(fi&&b.closest('.dl'))try{fi.focus({preventScroll:true})}catch(e){}}}return}
  if(a==='ap-udrehacer'){if(!p)return;undoPush();delete p.udMan;udRehacer(f,p);delete UDAV[p.id];p.m=Date.now();save();render();undoToast('Unidades rehechas por cultivo y riego');return}
  if(a==='ap-hacer'){if(!p)return;if(b.getAttribute('aria-disabled')==='true'){irFalta('dat');return}hacerPide(f,p);return}
  if(a==='ap-hacer-ok'){if(p)hacerPlan(f,p,b.dataset.m==='todas');return}
  if(!p)return;
  var ui=+b.dataset.u, u=p.ud[ui], C=u&&pabCalc(f,u,p);
  if(a==='ap-rowed'){var id=b.dataset.id;if(RM[id]){var fq=filaPorId(p,id);if(fq)abGuardaMio(fq.q)}RM[id]=!RM[id];render();if(RM[id]){var jj=-1;p.ud.forEach(function(x,i2){((x.pg||{}).rows||[]).forEach(function(q,j){if(q.id===id){jj=[i2,j]}})});if(jj!==-1)focoEn('[data-p="ud.'+jj[0]+'.pg.rows.'+jj[1]+'.n"]')}else focoEn('button[data-act="ap-rowed"][data-id="'+id+'"]');return}
  if(a==='ap-mitprop'){undoPush();var nm=0;udsAplica(f,p,function(x){nm+=mitPropuesta(x,true)||0});save();render();toast(nm?'Marcadas las medidas que encajan con el programa':'Ya estaban marcadas las que encajan');return}
  if(!u)return;
  if(a==='ap-zvsep'){zvSeparar(f,p,ui);return}
  if(a==='ap-exopen'){UI['exo'+u.id]=1;saveUI();render();return}
  if(a==='ap-prop'){if(sinCosecha(f,p,u,ui))return;var rows=propuesta(f,u,C);if(!rows.length){toast('Con lo que ya aporta, «'+u.n+'» no necesita abono (necesidad neta 0).');return}
    var had=((u.pg||{}).rows||[]).length;undoPush();u.pg={src:'own',rows:rows};delete u.d0;mitPropuesta(u,true);p.m=Date.now();save();render();undoToast('Programa propuesto con '+rows.length+(rows.length===1?' abono':' abonos')+(had?' (cambia el que había)':'')+'. Revísalo y cambia lo que quieras');return}
  if(a==='ap-ajustar'){if(sinCosecha(f,p,u,ui))return;var R=ajustar(f,p,u);undoPush();u.pg=u.pg||{src:'own'};u.pg.rows=R.rows;delete u.d0;p.m=Date.now();save();render();
    var C2=pabCalc(f,u,p),ok=['n','p','k'].every(function(k){return C2.B[k].st==='Ajustado'});
    undoToast((ok?'Ajustado: los tres nutrientes cuadran':'Ajustado lo que se puede'+(R.off.length?' (no cuadra '+R.off.map(function(k){return {n:'el N',p:'el P₂O₅',k:'el K₂O'}[k]}).join(' ni ')+')':''))+(R.added.length?'. Añadido: '+R.added.join(', '):'')+(R.quit.length?'. Quitado: '+R.quit.join(', '):'')+(R.limN?'. El N se queda en el límite de zona vulnerable':''));return}
});
document.addEventListener('submit',function(ev){var fm=ev.target;if(fm&&fm.dataset&&fm.dataset.f==='m-go'){ev.preventDefault();mGo(($('#m-q')||{}).value)}});
function pick(accept,fn){var i=document.createElement('input');i.type='file';i.accept=accept;i.style.display='none';document.body.appendChild(i);i.onchange=function(){var fl=i.files&&i.files[0];i.parentNode.removeChild(i);if(fl)fn(fl)};i.click()}

/* deshacer también lo escrito en las casillas: se guarda el estado justo antes de cada cambio */
document.addEventListener('change',function(ev){var t=ev.target;if(!t||!t.closest||!t.closest('#app,#sheet,#side,#desk'))return;if(t.id==='ap-q'||(t.classList&&/\b(rsel|rsel-all|ap-camp|ap-cost|ap-bulk|ap-clprov|ap-aplu)\b/.test(t.className)))return;if(t.dataset.p||t.dataset.e||t.dataset.r||t.dataset.m||t.classList.contains('ap-in')||t.classList.contains('ma-in')||t.classList.contains('ap-sulink')||t.classList.contains('ap-suno')||t.classList.contains('ap-sunomot')||t.classList.contains('ap-moall'))undoPush()},true);
document.addEventListener('click',function(ev){var b=ev.target.closest&&ev.target.closest('[data-act^="pa-"]');if(b&&b.dataset.act!=='pa-del')undoPush()},true);
/* fechas: lo escrito como dd/mm/aaaa se guarda como fecha */
document.addEventListener('change',function(ev){var t=ev.target;if(!t||!t.dataset||!t.dataset.date)return;var v=String(t.value||'').trim();if(!v)return;
  var m=/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/.exec(v),iso='';
  if(m){var y=+m[3];if(y<100)y+=2000;var d=new Date(y,+m[2]-1,+m[1]);if(d.getFullYear()===y&&d.getMonth()===+m[2]-1&&d.getDate()===+m[1])iso=y+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2)}
  else if(/^\d{4}-\d\d-\d\d$/.test(v)){var dd=new Date(v+'T12:00:00');if(!isNaN(dd)&&dd.toISOString().slice(0,10)===v)iso=v}
  if(!iso){ev.stopImmediatePropagation();t.classList.add('bad');toast('Pon la fecha como día/mes/año, por ejemplo 15/03/2027.');return}
  t.classList.remove('bad');t.value=iso},true);
/* casillas propias (finca, titular, recintos), campaña, lluvia, medidas para toda la finca, costes */
document.addEventListener('change',function(ev){
  var t=ev.target; if(!t||!t.classList)return;
  if(t.classList.contains('ma-in')){var a=(S.x.abonos||[])[+t.dataset.i];if(!a)return;var k=t.dataset.k;if(k==='org')a.org=t.checked?1:0;else if(k==='n')a.n=t.value.trim();else{var vn=t.value.trim()===''?undefined:numEs(t.value);if(vn===null||isNaN(vn)){toast('Pon un número (decimales con coma).');t.classList.add('bad');return}t.classList.remove('bad');if(vn===undefined)delete a[k];else a[k]=vn}save();return}
  var f=UI.v==='finca'?finca(UI.fid):null; if(!f)return;
  var p=pabCur(f);
  if(t.classList.contains('ap-in')){
    var v=String(t.value==null?'':t.value).trim();
    if(t.dataset.num){if(v===''){v=undefined}else{var n=numEs(v);if(n===null||isNaN(n)||n<0){toast('Pon un número (los decimales con coma, por ejemplo 2,35).');t.classList.add('bad');return}v=n}}
    if(t.dataset.e){var ks=t.dataset.e.split('.');if(ks[0]==='nombre'){if(!v){toast('La finca necesita un nombre.');renderLuego();return}f.nombre=v}else{f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};if(v)f.datos.tit[ks[1]]=v;else delete f.datos.tit[ks[1]]}}
    else if(t.dataset.r){var r=S.recintos[t.dataset.r];if(!r)return;var kk=t.dataset.k,un=kk==='cultivo'||kk==='riego',man=un&&p&&udMan(f,p);if(v===undefined||v==='')delete r[kk];else r[kk]=v;if(kk==='ha')delete r.hpac;
      if(un&&p){udTrasCambio(f,p,man);p.m=Date.now()}if(kk==='cultivo'&&v)UI.mLast={c:v,r:r.riego||'secano'};}
    save(); renderLuego(); return}
  if(t.classList.contains('rsel')){SEL[t.dataset.ref]=t.checked;render();return}
  if(t.classList.contains('rsel-all')){var on=t.checked;SEL={};if(on)fincaRecs(f.id).forEach(function(r){SEL[r.ref]=true});render();return}
  if(t.classList.contains('ap-bulk')){var ks2=Object.keys(SEL).filter(function(k){return SEL[k]&&S.recintos[k]&&S.recintos[k].fincaId===f.id});if(!ks2.length||!t.value)return;var man2=p&&udMan(f,p);undoPush();ks2.forEach(function(k){S.recintos[k][t.dataset.k]=t.value});if(p)udTrasCambio(f,p,man2);save();render();undoToast(ks2.length+(ks2.length===1?' recinto cambiado':' recintos cambiados'));return}
  if(t.classList.contains('ap-camp')){if(t.value==='new'){render();campHoja(f)}else{PAB.cur[f.id]=t.value;PREV.imgs=[];render()}return}
  if(t.classList.contains('ap-cost')){UI.cost=t.checked?1:0;saveUI();render();return}
  if(!p)return;
  if(t.classList.contains('pa-mv')){p.udMan=1;saveRaw();return}
  if(t.classList.contains('ap-clprov')){var c=CLIMA[+t.value];if(c){undoPush();p.cl={a:c.a,o:c.o,i:c.i,p:c.p,v:c.v,src:c.src};delete p.clAuto;p.m=Date.now();save();render();toast('Lluvia de '+c.n+' puesta')}return}
  if(t.classList.contains('pa-chk')){var path=t.dataset.p;pabChange(f,path,t.checked?1:0,false);var mm=/^ud\.(\d+)\.res\./.exec(path);if(mm&&p.ud[+mm[1]])delete p.ud[+mm[1]].resProp;save();renderLuego();return}
  if(t.classList.contains('ap-mitall')){var mk=t.dataset.m,ch=t.checked,dd=(PAB_MIT.filter(function(x){return x[0]===mk})[0]||[])[3];udsAplica(f,p,function(u){u.mit=u.mit||{};var o=u.mit[mk]||{};u.mit[mk]=ch?Object.assign({},o,{a:1}):Object.assign({},o,{a:0})});save();renderLuego();return}
  if(t.classList.contains('ap-mitd')){var mk2=t.dataset.m,tx=t.value.trim();udsAplica(f,p,function(u){var o=(u.mit||{})[mk2];if(o&&o.a===1){if(tx)o.d=tx;else delete o.d}});save();renderLuego();return}
  if(t.classList.contains('ap-resall2')){var rk=t.dataset.m,rc=t.checked;udsAplica(f,p,function(u){u.res=u.res||{};var o=u.res[rk]||{};u.res[rk]=Object.assign({},o,{a:rc?1:0});delete u.resProp});save();renderLuego();return}
  if(t.classList.contains('ap-resd')){var rk2=t.dataset.m,tx2=t.value.trim();udsAplica(f,p,function(u){var o=(u.res||{})[rk2];if(o&&o.a!==undefined){if(tx2)o.d=tx2;else delete o.d}});save();renderLuego();return}
  if(t.classList.contains('ap-aplu')){UI.aplU=t.value||null;saveUI();render();return}
  if(t.classList.contains('ap-moall')){var ob=t.value;todasUd(f,p,function(u){u.mo=u.mo||{};if(ob)u.mo.obj=ob;else delete u.mo.obj});save();renderLuego();return}
  if(t.classList.contains('ap-abmodo')){var fr=filaPorId(p,t.dataset.id);if(!fr)return;var md=t.value==='mio'?'mio':'tipo',qq=fr.q,esT=catEsTipo(qq.n);RMODE[qq.id]=md;
    if(qq.n&&(md==='mio'?esT:!esT)){undoPush();qq.n='';p.m=Date.now();save()}render();focoEn('[data-p="ud.'+fr.i+'.pg.rows.'+fr.j+'.n"]');return}
  if(t.classList.contains('sgi')&&t.dataset.sg==='abm'){var qm=filaDe(t.dataset.p||'');if(qm&&!(+qm.rn||+qm.rp||+qm.rk)){var km=abKey(qm.n),cm=catMios().filter(function(x){return abKey(x.n)===km})[0];if(cm){qm.rn=cm.rn;qm.rp=cm.rp;qm.rk=cm.rk;if(cm.org)qm.org=1;if(abHas(cm.pz)&&!abHas(qm.pz))qm.pz=cm.pz;p.m=Date.now();save();renderLuego()}}return}
  if(t.classList.contains('ap-suno')){var su0=p.ud[+t.dataset.u];if(!su0)return;if(t.checked)su0.suNo='nd';else delete su0.suNo;p.m=Date.now();save();renderLuego();return}
  if(t.classList.contains('ap-sunomot')){var su1=p.ud[+t.dataset.u];if(!su1)return;su1.suNo=t.value==='no20'?'no20':'nd';p.m=Date.now();save();renderLuego();return}
  if(t.classList.contains('ap-sulink')){var du=p.ud[+t.dataset.u];if(!du)return;var src=(p.ud||[]).filter(function(x){return x.id===t.value})[0];
    if(src&&src!==du){delete du.suNo;du.suDe=src.id;du.su=JSON.parse(JSON.stringify(src.su||{}));if(src.rp&&!du.rp)du.rp=du.refs[0]||''}else{delete du.suDe;du.su={}}p.m=Date.now();save();renderLuego();return}
  if(t.classList.contains('pa-in')&&t.dataset.p){var pa=t.dataset.p,m1=/^ud\.(\d+)\.(n|ck|rg)$/.exec(pa);if(m1){p.udMan=1;saveRaw()}
    /* v3.2: la zona vulnerable cambiada a mano se respeta siempre */
    var mz=/^ud\.(\d+)\.zv$/.exec(pa);if(mz&&p.ud[+mz[1]]){var uz=p.ud[+mz[1]];uz.zvMan=1;delete uz.zvA;delete uz.zvF;saveRaw()}
    var m2=/^ud\.(\d+)\.su\./.exec(pa);if(m2&&p.ud[+m2[1]]){if(suTiene(p.ud[+m2[1]].su))delete p.ud[+m2[1]].suNo;suCopia(p,p.ud[+m2[1]]);saveRaw()}
    if(/^cl\./.test(pa)&&p.clAuto){delete p.clAuto;saveRaw()}}
});
document.addEventListener('input',function(ev){var t=ev.target;if(!t)return;
  if(t.id==='ap-q'){UI.q=t.value;saveUI();render();return}
  if(t.classList&&t.classList.contains('sgi')){sgOpen(t);return}
  if(t.classList&&(t.classList.contains('pa-in')||t.dataset.num)){var w=valida(t.dataset.p||'',t.value,filaDe(t.dataset.p||''),t.getAttribute('inputmode')==='decimal');t.classList.toggle('rare',!!w);if(w)t.title=w}});
/* la fila en la que estás manda en lo que se resalta en la foto */
document.addEventListener('focusin',function(ev){var t=ev.target;if(!t||!t.classList)return;hlAuto();if(t.classList.contains('sgi')&&!t.value)sgOpen(t)});
document.addEventListener('mouseover',function(ev){var e=ev.target.closest&&ev.target.closest('#desk .leg [data-ref]');if(e)hlSet([e.dataset.ref])});
document.addEventListener('mouseout',function(ev){var e=ev.target.closest&&ev.target.closest('#desk .leg');if(e&&!(ev.relatedTarget&&e.contains(ev.relatedTarget)))hlAuto()});
document.addEventListener('toggle',function(ev){var t=ev.target;if(!t||!t.dataset)return;if(t.dataset.o==='tit'){UI.titOpen=t.open;saveUI()}if(t.dataset.o==='datmas'){UI.datMas=t.open;saveUI()}},true);

/* ───── lista de sugerencias propia (abonos y meses): flota sobre la página, no la corta la tabla (fallo 7) ───── */
var SG={box:null,inp:null,L:[],hi:-1};
var MES_SG=MESES.concat(['Febrero-marzo','Marzo-abril','Abril-junio','Marzo-junio','Mayo-julio','Octubre-noviembre']);
function sgPos(){if(!SG.box||!SG.inp)return;var r=SG.inp.getBoundingClientRect(),w=Math.max(r.width,340),x=Math.min(window.innerWidth-w-8,r.left+window.scrollX),below=window.innerHeight-r.bottom;
  SG.box.style.width=w+'px';SG.box.style.left=Math.max(8,x)+'px';var mh=Math.min(320,Math.max(160,below>220?below-16:r.top-16));SG.box.style.maxHeight=mh+'px';
  if(below>220||below>r.top)SG.box.style.top=(r.bottom+window.scrollY+4)+'px';else SG.box.style.top=(r.top+window.scrollY-Math.min(mh,SG.box.scrollHeight)-4)+'px'}
function sgOpen(inp){var kind=inp.dataset.sg,q=noAcc(inp.value||'').toLowerCase().trim(),L;
  if(/^ab/.test(kind))L=catBusca(inp.value,kind==='abt'?'tipo':kind==='abm'?'mio':'').slice(0,60);else L=MES_SG.filter(function(m){return !q||noAcc(m).toLowerCase().indexOf(q)===0}).map(function(m){return {n:m}});
  if(SG.inp!==inp)sgClose();SG.inp=inp;SG.L=L;SG.hi=-1;
  if(!L.length){if(SG.box){SG.box.parentNode.removeChild(SG.box);SG.box=null}return}
  if(!SG.box){SG.box=document.createElement('div');SG.box.className='sgl';SG.box.setAttribute('role','listbox');document.body.appendChild(SG.box)}
  var g0='';SG.box.innerHTML=L.map(function(x,i){var gh='';if(/^ab/.test(kind)&&catGrupo(x)!==g0){g0=catGrupo(x);gh='<div class="sgg">'+esc(g0)+'</div>'}
    return gh+'<div class="sgo" role="option" data-i="'+i+'"><span>'+esc(x.n)+'</span>'+(/^ab/.test(kind)?'<em>'+trTxt(x.rn,1)+'-'+trTxt(x.rp,1)+'-'+trTxt(x.rk,1)+(x.org?' · orgánico':'')+'</em>':'')+'</div>'}).join('');sgPos()}
function sgClose(){if(SG.box&&SG.box.parentNode)SG.box.parentNode.removeChild(SG.box);SG.box=null;SG.inp=null;SG.L=[];SG.hi=-1}
window.addEventListener('resize',sgPos);window.addEventListener('scroll',function(){if(SG.box)sgPos()},true);
function sgHi(d){if(!SG.box)return;var n=SG.L.length;SG.hi=SG.hi<0?(d>0?0:n-1):(SG.hi+d+n)%n;$$('.sgo',SG.box).forEach(function(o){o.classList.toggle('on',+o.dataset.i===SG.hi)});var on=$('.sgo.on',SG.box);if(on)on.scrollIntoView({block:'nearest'})}
function sgPick(i){var inp=SG.inp,x=SG.L[i];if(!inp||!x)return;var kind=inp.dataset.sg,path=inp.dataset.p;sgClose();inp.value=x.n;inp.defaultValue=x.n;
  var nxt=null;
  if(/^ab/.test(kind)){var q=filaDe(path),f=finca(UI.fid),p=f&&pabCur(f);if(q&&p){q.n=x.n;q.rn=x.rn;q.rp=x.rp;q.rk=x.rk;if(x.org)q.org=1;else delete q.org;if(abHas(x.pz))q.pz=x.pz;
      var mu=/^ud\.(\d+)\./.exec(path),u=mu&&p.ud[+mu[1]];if(!q.fo)q.fo=u&&u.rg==='regadio'?'Fertirrigación':'Suelo';if(!q.eq)q.eq=u&&u.rg==='regadio'?'Cabezal de riego':'Abonadora';p.m=Date.now();save()}
    nxt=$('[data-p="'+path.replace(/\.n$/,'.mo')+'"]')}
  else{var q2=filaDe(path),f2=finca(UI.fid),p2=f2&&pabCur(f2);if(q2&&p2){q2.mo=x.n;p2.m=Date.now();save()}nxt=$('[data-p="'+path.replace(/\.mo$/,'.d')+'"]')}
  if(nxt){nxt.focus();if(nxt.select)nxt.select()}
  renderLuego()}
document.addEventListener('mousedown',function(ev){var o=ev.target.closest&&ev.target.closest('.sgo');if(o){ev.preventDefault();sgPick(+o.dataset.i)}});
var SGT=null;document.addEventListener('touchstart',function(ev){var t=ev.touches&&ev.touches[0];SGT=t?{x:t.clientX,y:t.clientY}:null},{passive:true});
document.addEventListener('touchend',function(ev){var o=ev.target.closest&&ev.target.closest('.sgo');if(!o)return;var t=ev.changedTouches&&ev.changedTouches[0];if(SGT&&t&&(Math.abs(t.clientX-SGT.x)>8||Math.abs(t.clientY-SGT.y)>8))return;ev.preventDefault();sgPick(+o.dataset.i)},{passive:false});

/* ───── teclado: Intro pasa a la casilla siguiente como Tab (punto 33); atajos ───── */
function focables(root){return $$('input:not([type=hidden]):not([disabled]),select:not([disabled]),textarea:not([disabled])',root).filter(function(e){return e.offsetParent!==null&&e.type!=='checkbox'&&e.type!=='file'&&!e.closest('details:not([open])')})}
function enterSigue(t,back){
  /* en la tabla del programa, la dosis pasa a la dosis de la fila siguiente; en la última, a «Añadir abono» */
  var tr=t.closest('table.pg tbody tr[data-j]');
  if(tr&&/\.pg\.rows\.\d+\.d$/.test(t.dataset.p||'')){var rows=$$('tr[data-j]',tr.parentNode),j=rows.indexOf(tr),nx=rows[j+(back?-1:1)];var dst=nx?$('input.box',nx):$('[data-act="pa-row-add"][data-u="'+t.dataset.p.split('.')[1]+'"]');if(dst){dst.focus();if(dst.select)dst.select();return true}}
  var root=t.closest('#sheet,#mapa,#app')||document.body,L=focables(root),i=L.indexOf(t);if(i<0)return false;var n=L[i+(back?-1:1)];
  if(n){n.focus();if(n.select&&n.tagName==='INPUT')try{n.select()}catch(e){}return true}
  t.blur();t.focus();return true}
document.addEventListener('keydown',function(ev){var t=ev.target,k=ev.key,mod=ev.ctrlKey||ev.metaKey,ed=t&&(/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName)||t.isContentEditable);
  if(SG.box&&t===SG.inp){if(k==='ArrowDown'){ev.preventDefault();sgHi(1);return}if(k==='ArrowUp'){ev.preventDefault();sgHi(-1);return}if(k==='Escape'){ev.preventDefault();sgClose();return}
    if((k==='Enter'||k==='Tab')&&SG.hi>=0){ev.preventDefault();sgPick(SG.hi);return}sgClose()}
  else if(t&&t.classList&&t.classList.contains('sgi')&&k==='ArrowDown'){ev.preventDefault();sgOpen(t);sgHi(1);return}
  if(k==='Escape'){if(!$('#pop').hidden){popClose();return}if(!$('#sheet').hidden){sheetClose();return}if(!$('#mapa').hidden){mapaCerrar();return}}
  if(mod&&(k==='s'||k==='S'||k==='p'||k==='P')&&ed&&t.blur)t.blur();
  if(mod&&(k==='s'||k==='S')){ev.preventDefault();if(FS.ok)fsWrite(2);else if(S.fincas.length)copiaTodo();return}
  if(mod&&(k==='z'||k==='Z')&&!ev.shiftKey&&!ed){ev.preventDefault();undoDo();return}
  if(mod&&(k==='p'||k==='P')){var f=UI.v==='finca'&&finca(UI.fid),p=f&&pabCur(f);if(p){ev.preventDefault();outGo(f,p,'pdf')}return}
  if(ev.altKey&&(k==='ArrowLeft'||k==='ArrowRight')&&UI.v==='finca'){ev.preventDefault();var i=stepIdx()+(k==='ArrowRight'?1:-1);if(i>=0&&i<STEPS.length)irPaso(STEPS[i][0]);return}
  if((k==='Enter'||k===' ')&&t.classList&&t.classList.contains('tap')&&t===document.activeElement&&!ed){ev.preventDefault();t.click();return}
  if((k==='Enter'||k===' ')&&t.classList&&t.classList.contains('pacrow')){ev.preventDefault();t.click();return}
  if(k==='Enter'&&!mod&&!ev.altKey&&t.tagName==='INPUT'&&t.type!=='checkbox'){
    if(t.id==='ap-ref'||t.id==='ap-rh'){ev.preventDefault();refAdd();return}
    if(t.id==='m-q'||t.id==='ap-q')return;
    if(t.closest('#app,#sheet')){ev.preventDefault();enterSigue(t,ev.shiftKey);return}}
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
var MQ=[window.matchMedia('(min-width:1200px)')];
MQ.forEach(function(q){var fn=function(){if($('#app'))render()};if(q.addEventListener)q.addEventListener('change',fn);else if(q.addListener)q.addListener(fn)});
document.addEventListener('DOMContentLoaded',function(){render();fsInit();if(UI.v==='finca')zvComprobar(finca(UI.fid))});
(function(){var tm=null;try{new MutationObserver(function(){if(tm)return;tm=setTimeout(function(){tm=null;stkLab();inThou()},0)}).observe(document.documentElement,{childList:true,subtree:true})}catch(e){}})();
if('serviceWorker' in navigator&&(location.protocol==='https:'||location.hostname==='localhost'))window.addEventListener('load',function(){navigator.serviceWorker.register('sw.js').catch(function(){})});
window.__pa={S:function(){return S},UI:function(){return UI},render:render,ejemplo:ejemplo,pabCalc:pabCalc,pabBlocks:pabBlocks,buildPabXlsx:buildPabXlsx,ensurePdf:ensurePdf,buildPabPdf:buildPabPdf,climaDe:climaDe,climaKey:climaKey,refParse:refParse,catGen:catGen,mesesDe:mesesDe,propuesta:propuesta,faltas:faltas,ctxOf:ctxOf,mapaAbrir:mapaAbrir,MAP:function(){return MAP},
  ajustar:ajustar,udRehacer:udRehacer,udMan:udMan,udEst:udEst,estGlobal:estGlobal,nifMal:nifMal,GEO:function(){return GEO},PREV:function(){return PREV},abrirFinca:abrirFinca};

/* ───── v3: parches pequeños sobre base.js ─────
   base.js sale de otra aplicación (herramientas/extraer.js) y no se edita. Aquí se cambian solo cuatro cosas, cada una con su motivo:
   1. Punto 27: extracciones de la guía del MAPA que faltaban (maíz, centeno y melocotonero), leídas en EXTRACCIONES.md.
      Viña, triticale y «otros frutales» se quedan sin cifra a propósito: la guía no da extracciones por tonelada para ellos.
   2. Fallo 10: en el PDF y el Excel la columna «Forma» del programa era tan estrecha que «Fertirrigación» salía partida en dos líneas.
   3. Punto 28: una unidad exenta (secano de 10 ha o menos) que está vacía ocupa una sola línea en el PDF, no página y media.
   4. Revisión: «Sin datos» en la columna «Estado» del PDF y el Excel de una unidad sin cosecha esperada o sin extracciones.
   El cálculo del balance, los límites de N del Decreto 80/2025 y las fuentes de los demás cultivos no se tocan. */
var GUIA_MAPA='Guía práctica de la fertilización racional de los cultivos en España (MAPA, 2010)';
PAB_EX.maiz=[[27,10,20],GUIA_MAPA+', cap. 17, tabla 17.6, pág. 140: necesidades por t de grano (grano y resto de la parte aérea)'];
PAB_EX.centeno=[[19,13,18],GUIA_MAPA+', cap. 16, tabla 16.2, pág. 125: valor medio del rango de extracciones por t de grano (N 18-20, P₂O₅ 12-14 y K₂O 16-20; grano y paja)'];
PAB_EX.melocotonero=[[3.5,1.71,3.84],GUIA_MAPA+', cap. 25, tabla 25.7, pág. 209 (N, extracción neta) y tabla 25.4, pág. 207 (P₂O₅ y K₂O): kg por t de fruto'];

/* revisión: unidad que necesita plan y no tiene cosecha esperada (por cosecha) o no tiene extracciones */
var SIN_DATOS='Sin datos';
function udSinDatos(u,C){if(C.ex||u.d0===1)return false;var ne=u.ne||{};return ((ne.mo||'cos')==='cos'&&!(+ne.pr>0))||noExtr(u)}
/* unidad exenta sin nada rellenado */
function udVacia(u,C){var su=u.su||{};return !!C.ex&&!((u.pg||{}).rows||[]).length&&!abHas((u.ne||{}).pr)&&!['mo','ph','n','p','k','f','lab'].some(function(k){return abHas(su[k])})}
var pabBlocks0=pabBlocks;
pabBlocks=function(f,p){
  var D=pabBlocks0(f,p);
  /* fallo 10: «Forma» más ancha (la «Mitigación» cede lo mismo; el total de la tabla no cambia) */
  D.B.forEach(function(b){if(b.t==='tb'&&b.cols.length===8&&b.cols[0].h==='Momento'&&b.cols[5].h==='Forma'&&b.cols[7].h==='Mitigación'){b.cols[5].w=23;b.cols[7].w=30}});
  /* punto 28: las unidades exentas vacías pasan a una línea debajo del resumen */
  var uds=p.ud||[], Cs=uds.map(function(u){return pabCalc(f,u,p)}), vac=uds.map(function(u,i){return udVacia(u,Cs[i])});
  /* revisión: una unidad que necesita plan pero no tiene cosecha esperada o no tiene extracciones dice «Sin datos» en la columna
     «Estado» (resumen y balance), no «Ajustado» con 0 de 0. Un plan completo no cambia: solo toca las unidades a las que les falta eso. */
  var sd=uds.map(function(u,i){return udSinDatos(u,Cs[i])});
  if(sd.some(Boolean)){var uj=-1;
    D.B.forEach(function(b){
      if(b.t==='h1'&&/^Unidad de producción: /.test(b.x))uj++;
      if(b.t!=='tb')return;var c=b.cols,L=c.length-1;if(c[L].h!=='Estado')return;
      if(c[0].h==='Unidad'&&b.rows.length===uds.length)b.rows.forEach(function(r,i){if(sd[i])r[L]=SIN_DATOS});
      else if(c[0].h==='Nutriente'&&uj>=0&&sd[uj])b.rows.forEach(function(r){r[L]=SIN_DATOS})})}
  /* v3.2: unidad sin análisis de suelo («No tengo análisis»): donde iba el análisis y en la comprobación 6.2 dice «Sin análisis de suelo» */
  var sn=uds.map(function(u,i){return suNoMot(u,Cs[i])});
  if(sn.some(Boolean)){var us=-1;D.B.forEach(function(b){if(b.t==='h1'&&/^Unidad de producción: /.test(b.x)){us++;return}var m=us>=0&&sn[us];if(!m)return;
    if(b.t==='kv'&&b.k==='Resultados')b.v=SUNO_TXT[m];
    if(b.t==='tb'&&b.cols[0].h==='Apartado')b.rows.forEach(function(r){if(/^6\.2 /.test(r[0])){r[1]='Sin análisis de suelo';r[2]=m==='no20'?'no obligatorio (RD 1311/2012, anexo III, parte I, A.3, datos generales, letra d)':'no se dispone de análisis de suelo'}})})}
  /* v3.2: de dónde sale la zona vulnerable de cada unidad (lo que no se sabe de dónde sale, como antes) */
  var uz=-1;D.B.forEach(function(b){if(b.t==='h1'&&/^Unidad de producción: /.test(b.x)){uz++;b.zvU=uz;return}
    if(b.t==='p'&&uz>=0&&!b.zvHecho&&/ · (fuera de )?zona vulnerable$/.test(b.x)){var u=uds[uz];b.zvHecho=1;if(!u)return;
      if(u.zvMan===1)b.x+=' (indicado a mano)';else if(u.zvA===1&&!zvMan(u)&&u.zvF)b.x+=' (según el SIGPAC del FEGA, consultado el '+isoEs(u.zvF)+')'}});
  if(!vac.some(Boolean))return D;
  /* v3.1 punto 6: de la unidad exenta vacía se quedan su título, su línea de superficie, la frase de exenta y la tabla de sus recintos
     (igual que en las demás unidades); lo que se quita es lo vacío: datos de inicio, suelo, balance, programa, medidas y comprobación. */
  var out=[], lines=[], ui=-1, vu=false, keep=true;
  D.B.forEach(function(b,i){var nx=D.B[i+1];
    if(b.t==='np'){if(nx&&nx.t==='h1'&&/^Unidad de producción: /.test(nx.x)){ui++;vu=!!vac[ui];keep=true;if(vu){var C=Cs[ui];lines.push(uds[ui].n+' ('+ha2(C.ha)+' ha, '+C.rs.length+(C.rs.length===1?' recinto':' recintos')+'): secano de 10 ha o menos, exenta de plan de abonado (art. 4.2 b del RD 1051/2022).')}}else{vu=false;keep=true}}
    else if(vu&&b.t==='h')keep=b.x==='Recintos (art. 6.1)';
    if(!vu||keep)out.push(b)});
  var at=-1;out.forEach(function(b,i){if(at<0&&b.t==='p'&&b.x==='U.F./ha planificadas.')at=i});
  var add=lines.map(function(t){return {t:'p',x:t,o:{sz:9}}});
  if(at>=0)out.splice.apply(out,[at+1,0].concat(add));else out=add.concat(out);
  D.B=out;return D};

/* el redibujado de base.js pasa por el de la versión 3, que conserva la casilla, lo tecleado y el cursor */
pabRedraw=function(){renderLuego()};

/* ───── v3.1: lectura de la PAC (puntos 1, 2 y 3) ─────
   base.js lee cada renglón del PDF por separado y se queda con la primera línea de cada recinto. En la PAC de Castilla-La Mancha:
   1. el producto va detrás del uso SIGPAC («TA TRIGO BLANDO», «FS PISTACHO», o solo «TA» con el producto en otro renglón);
   2. el producto a veces ocupa dos renglones dentro de la misma fila («BARBECHO» arriba y «TRADICIONAL» abajo; el uso «O»/«V» igual);
   3. un recinto puede ir declarado en varias líneas (trigo + trigo + barbecho).
   Aquí se redefinen pacParseRow, pacCrop y pacRowsFromItems. Si la maqueta es otra, se lee como antes (renglón a renglón). */
var PAC_USO=(function(){var o={};'TA TH PS PR PA OV VI FS FY FL FV FF OF OC CF CS CV CI IV IM ZU ZV ZC CA ED AG FO EP MT OP VO VF'.split(' ').concat(Object.keys(USOS)).forEach(function(k){o[k]=1});return o})();
/* palabras que van pegadas al producto, no a la variedad */
var PAC_QUAL=/^(BLANDO|DURO|TRADICIONAL|MEDIOAMBIENTAL|SEMILLADO|VERDE|GRANO|FORRAJERO|FORRAJERA|DE)$/;
/* quita los códigos de uso SIGPAC del principio: «TA TRIGO BLANDO» → uso TA, producto «TRIGO BLANDO» */
function pacSinUso(tk){var uso='';tk=tk.slice();while(tk.length&&PAC_USO[noAcc(tk[0]).toUpperCase()]){if(!uso)uso=noAcc(tk[0]).toUpperCase();tk.shift()}return {uso:uso,tk:tk}}
var pacParseRow0=pacParseRow;
/* area: opcional, [{x,s}] con las casillas de producto y variedad de la fila (con los trozos de los renglones de arriba y abajo) y vx, la x donde empieza la variedad */
pacParseRow=function(txt,area){
  var r=pacParseRow0(txt); if(!r)return null;
  var prod=[],vari=[],u0='';
  if(area&&area.cells&&area.cells.length){area.cells.forEach(function(c){(c.x>=area.vx-3?vari:prod).push.apply(c.x>=area.vx-3?vari:prod,c.s.split(/\s+/).filter(Boolean))})}
  else{var t=String(txt).trim().split(/\s+/),i=3;while(i<t.length&&!/^\d+$/.test(t[i]))i++;var rest=t.slice(i+9),k=-1;for(var j=0;j<rest.length;j++){if(/^[SR]$/i.test(rest[j])){k=j;break}}
    var pv=k>=0?rest.slice(0,k):rest.slice(0,4),su=pacSinUso(pv);pv=su.tk;u0=su.uso;
    if(pv.length>=2&&/^SIN$/i.test(pv[pv.length-2])&&/^VARIEDAD$/i.test(pv[pv.length-1]))pv=pv.slice(0,-2);
    if(pv.length){prod=[pv[0]];var q=1;while(q<pv.length&&PAC_QUAL.test(noAcc(pv[q]).toUpperCase())&&!(q+1<pv.length&&/^VARIEDAD$/i.test(pv[q+1])))prod.push(pv[q++]);vari=pv.slice(q)}}
  var s=pacSinUso(prod);prod=s.tk;
  /* «SIN VARIEDAD» que se haya quedado pegado al producto */
  var pj=prod.join(' ').replace(/\s*\bSIN\s+VARIEDAD\b\s*$/i,'').trim(), vj=vari.join(' ').trim();
  if(/^SIN\s+VARIEDAD$/i.test(vj))vj='';
  r.prod=pj; r.vari=vj; r.uso=u0||s.uso||''; return r};
/* cultivo de la herramienta a partir del producto de la PAC (los cereales, con su nombre: «Trigo», no «Cereal») */
var pacCrop0=pacCrop;
pacCrop=function(prod){var tk=noAcc(String(prod||'')).toUpperCase().trim().split(/\s+/).filter(Boolean),p=pacSinUso(tk).tk.join(' ');
  if(!p)return '';
  if(/^TRIGO\s+DURO/.test(p))return 'Trigo duro'; if(/^TRIGO/.test(p))return 'Trigo'; if(/^CEBADA/.test(p))return 'Cebada'; if(/^AVENA/.test(p))return 'Avena';
  if(/^CENTENO/.test(p))return 'Centeno'; if(/^TRITICALE/.test(p))return 'Triticale'; if(/^MAIZ/.test(p))return 'Maíz'; if(/^NOGAL/.test(p))return 'Nogal'; if(/^MELOCOT/.test(p))return 'Melocotonero';
  return pacCrop0(p)};
/* cultivo de una línea: un producto vacío no da cultivo (nunca «Ta» ni «Fs») */
function pacCultivo(r){var c=pacCrop(r.prod)||pacCropVar(r.vari);if(c)return c;var p=String(r.prod||'').trim();if(!p||/^[A-Za-z]{1,2}$/.test(p))return '';return titleCase2(p)}
/* renglones de una hoja: [{y, w:[{x,s,i}]}] ordenados de arriba abajo y de izquierda a derecha (como base.js: ±2,5 de y) */
function pacLineas(pg){var vt=pg.vt||[1,0,0,-1,0,0],ys=[];
  pg.items.forEach(function(it){var s=it.str;if(!s||!s.trim())return;var xy=pacXY(vt,it.transform),y=xy[1],g=null;for(var i=0;i<ys.length;i++){if(Math.abs(ys[i].y-y)<=2.5){g=ys[i];break}}if(!g){g={y:y,w:[]};ys.push(g)}g.w.push({x:xy[0],s:s.trim()})});
  ys.sort(function(a,b){return a.y-b.y});ys.forEach(function(g){g.w.sort(function(a,b){return a.x-b.x})});return ys}
/* casillas de producto y variedad de una fila de recinto: entre la del CSP (la última cifra) y la de «S»/«R» */
function pacArea(g,r,txt){var t=String(txt).trim().split(/\s+/),i=3;while(i<t.length&&!/^\d+$/.test(t[i]))i++;var tCsp=i+8,acc=0,iC=-1,iR=-1;
  for(var j=0;j<g.w.length;j++){var n=g.w[j].s.split(/\s+/).length;if(iC<0&&acc+n>tCsp){iC=j}else if(iC>=0&&/^[SR]$/i.test(g.w[j].s)){iR=j;break}acc+=n}
  if(iC<0||iR<0)return null;var own=g.w.slice(iC+1,iR);return {x0:g.w[iC].x+1,x1:g.w[iR].x-1,own:own,y:g.y}}
pacRowsFromItems=function(pages){
  var lines=[],year='',byKey={},order=[];
  pages.forEach(function(pg,pi){
    var L=pacLineas(pg),main=[];
    L.forEach(function(g){var txt=g.w.map(function(w){return w.s}).join(' ');if(!year){var m=txt.match(/Campa[ñn]a\s*(\d{4})/);if(m)year=m[1]}
      var r=pacParseRow0(txt);if(r){g.r=r;g.txt=txt;g.A=pacArea(g,r,txt);main.push(g)}});
    /* punto 2: los trozos de los renglones de arriba y abajo que caen en la columna del producto de una fila (a menos de 6,5 de su y) */
    L.forEach(function(g){if(g.r)return;g.w.forEach(function(w){var best=null,bd=6.5;main.forEach(function(m){if(!m.A)return;var d=Math.abs(m.y-g.y);if(d<bd&&w.x>m.A.x0&&w.x<m.A.x1){bd=d;best=m}});if(best)(best.frag=best.frag||[]).push({x:w.x,y:g.y,s:w.s})})});
    main.forEach(function(g){var area=null;
      if(g.A){var P=g.A.own.map(function(w){return {x:w.x,y:g.y,s:w.s}}).concat(g.frag||[]),cl=[];
        P.sort(function(a,b){return a.x-b.x||a.y-b.y}).forEach(function(p){var c=cl.filter(function(c){return Math.abs(c.x-p.x)<=3})[0];if(!c){c={x:p.x,p:[]};cl.push(c)}c.p.push(p)});
        var cells=cl.map(function(c){c.p.sort(function(a,b){return a.y-b.y});var one=c.p.every(function(p){return /^[A-Za-zÑñ]$/.test(p.s)});return {x:c.x,s:c.p.map(function(p){return p.s}).join(one?'':' ')}});
        var vx=g.A.own.length>=2?g.A.own[g.A.own.length-1].x:Infinity;area={cells:cells,vx:vx}}
      var r=pacParseRow(g.txt,area);if(!r)return;r.key=[r.prov,r.muni,r.agr,r.zona,r.pol,r.par,r.rec].join(':');r.cultivo=pacCultivo(r);
      if(!byKey[r.key]){byKey[r.key]=[];order.push(r.key)}byKey[r.key].push(r)})});
  /* punto 3: un recinto en varias líneas. Las del mismo cultivo se suman; con dos cultivos se queda el de más hectáreas y se avisa */
  var rows=order.map(function(k){var L=byKey[k],G={},go=[];
    /* si lo declarado suma más que la superficie SIGPAC del recinto, es que la PAC repite la misma línea en otra tabla: las repetidas (mismo cultivo y mismas ha) cuentan una vez */
    var tot=L.reduce(function(a,r){return a+(r.decHa||0)},0);
    if(L.length>1&&L[0].sigHa>0&&tot>L[0].sigHa+0.006){var visto={};L=L.filter(function(r){var q=(r.cultivo||'')+'|'+(r.decHa||0);if(visto[q])return false;visto[q]=1;return true})}
    L.forEach(function(r){var c=r.cultivo||'';if(!G[c]){G[c]={c:c,ha:0,r:r};go.push(G[c])}G[c].ha+=r.decHa||0});
    /* se queda el cultivo de más hectáreas; una línea sin cultivo solo gana si no hay otra */
    var top=go.slice().sort(function(a,b){return (a.c?0:1)-(b.c?0:1)||b.ha-a.ha})[0],o=Object.assign({},top.r);
    o.decHa=Math.round(top.ha*100)/100;o.lineas=L.length;
    if(go.length>1)o.mix=go.map(function(g){return {c:g.c||'Sin cultivo',ha:Math.round(g.ha*100)/100}}).sort(function(a,b){return b.ha-a.ha});
    return o});
  return {rows:rows,year:year}};
