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
/* Plan de abonado — herramienta independiente.
   Todo se guarda en este aparato (localStorage) y en las copias que el usuario descarga. No hay servidor ni cuentas.
   El cálculo, el PDF y el Excel del plan vienen de base.js (RD 1051/2022 y Decreto 80/2025 de Castilla-La Mancha). */
var VERSION="1 · 17 nov 2026";
var KEY='plan-abonado:datos', KEYUI='plan-abonado:vista';
var $=function(s,r){return (r||document).querySelector(s)}, $$=function(s,r){return [].slice.call((r||document).querySelectorAll(s))};

/* ───── datos ───── */
var S={fincas:[],recintos:{},x:{asesor:{}},copia:0,mod:0};
var UI={v:'inicio',fid:null,step:'rec'};
var GTOK=false, EXP={fid:null};
function cargar(){
  try{var t=localStorage.getItem(KEY);if(t){var o=JSON.parse(t);if(o&&Array.isArray(o.fincas)){S=o;S.recintos=S.recintos||{};S.x=S.x||{};S.x.asesor=S.x.asesor||{}}}}catch(e){}
  try{var u=JSON.parse(localStorage.getItem(KEYUI)||'null');if(u&&u.v){UI=u}}catch(e){}
  if(UI.fid&&!finca(UI.fid))UI={v:'inicio',fid:null,step:'rec'};
}
var guardaFallo=false;
function save(){
  S.mod=Date.now();
  try{localStorage.setItem(KEY,JSON.stringify(S));guardaFallo=false}catch(e){if(!guardaFallo)toast('Este navegador no deja guardar los datos. Guarda una copia en archivo.');guardaFallo=true}
}
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
var ha4=function(n){return nf(n,4)};

/* deshacer */
var UNDO=[];
function undoPush(){try{UNDO.push(JSON.stringify({fincas:S.fincas,recintos:S.recintos,x:S.x}));if(UNDO.length>25)UNDO.shift()}catch(e){}}
function undoDo(){var t=UNDO.pop();if(!t)return;var o=JSON.parse(t);S.fincas=o.fincas;S.recintos=o.recintos;S.x=o.x;save();if(UI.fid&&!finca(UI.fid)){UI.v='inicio';UI.fid=null}render();toast('Deshecho')}
function undoToast(msg){toast(msg,'Deshacer',undoDo)}

/* avisos */
var toastT=null;
function toast(msg,btn,fn){var el=$('#toast');clearTimeout(toastT);el.innerHTML='<div><span>'+esc(msg)+'</span>'+(btn?'<button type="button" id="tbtn">'+esc(btn)+'</button>':'')+'</div>';
  if(btn)$('#tbtn').onclick=function(){el.innerHTML='';fn()};
  if(NUMDOT&&Date.now()-NUMDOT<4000&&/n[uú]mero/.test(msg)){NUMDOT=0}
  toastT=setTimeout(function(){el.innerHTML=''},btn?7000:4000)}

/* ───── librerías que se cargan solo al usarlas (PDF y lector de la PAC) ───── */
function cargaScript(src){return new Promise(function(res,rej){var s=document.createElement('script');s.src=src;s.onload=function(){res()};s.onerror=function(){rej(new Error('net'))};document.head.appendChild(s)})}
var pdfP=null, pdfjsP=null;
function ensurePdf(){if(window.jspdf&&window.jspdf.jsPDF)return Promise.resolve();if(!pdfP)pdfP=cargaScript('lib/pdf.js').catch(function(e){pdfP=null;throw e});return pdfP}
function ensurePdfJs(){if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);if(!pdfjsP)pdfjsP=cargaScript('lib/pdfjs.js').then(function(){window.pdfjsLib.GlobalWorkerOptions.workerSrc='lib/pdfjs-worker.js';return window.pdfjsLib}).catch(function(e){pdfjsP=null;throw e});return pdfjsP}
function loadPdfLib(){return !!(window.jspdf&&window.jspdf.jsPDF)}
function pdfNoListo(){toast('No he podido preparar el PDF. Ábrela una vez con internet y prueba otra vez.')}
function pdfFont(doc){
  try{var F=window.PA_FONTS;if(!F||!F.r)return 'helvetica';
    if(!doc.__bw){doc.addFileToVFS('bw.ttf',F.r);doc.addFont('bw.ttf','Barlow','normal','Identity-H');doc.addFileToVFS('bw-b.ttf',F.b);doc.addFont('bw-b.ttf','Barlow','bold','Identity-H');
      if(F.cr&&F.cb){doc.addFileToVFS('bwc.ttf',F.cr);doc.addFont('bwc.ttf','BarlowC','normal','Identity-H');doc.addFileToVFS('bwc-b.ttf',F.cb);doc.addFont('bwc-b.ttf','BarlowC','bold','Identity-H');doc.__bwc=1}
      doc.__bw=1}
    return 'Barlow'}catch(e){return 'helvetica'}
}

/* ───── ventana de abajo (hojas) y descargas ───── */
function sheet(h){var el=$('#sheet');el.innerHTML='<div class="box" role="dialog" aria-modal="true">'+h+'</div>';el.hidden=false}
function sheetClose(){var el=$('#sheet');el.hidden=true;el.innerHTML='';PACI=null}
var DL=null;
function showDl(blob,name,title,extra){
  DL={blob:blob,name:name}; var sh=canShareFile(blob,name), phone=!isDesk()&&sh;
  sheet('<h2>'+esc(title||'Archivo listo')+'</h2><p class="small">'+esc(name)+' · '+Math.max(1,Math.round(blob.size/1024))+' KB</p>'+(extra?'<p class="small">'+extra+'</p>':'')+
    (phone?'<p class="small">Pulsa «Guardar / Enviar…» y elige «Guardar en Archivos», el correo o WhatsApp.</p>':'<p class="small">Se guarda en la carpeta de Descargas.</p>')+
    '<div class="acts">'+(phone?'<button class="btn primary xl" data-act="ap-dl-share">Guardar / Enviar…</button><button class="btn" data-act="ap-dl-save">Descargar</button>':'<button class="btn primary xl" data-act="ap-dl-save">Guardar archivo</button>'+(sh?'<button class="btn" data-act="ap-dl-share">Enviar…</button>':''))+
    (/\.pdf$/.test(name)&&isDesk()?'<button class="btn" data-act="ap-dl-view">Ver</button>':'')+'<button class="btn" data-act="ap-close">Cerrar</button></div>');
}
function dlSave(){if(!DL)return;var u=URL.createObjectURL(DL.blob),a=document.createElement('a');a.href=u;a.download=DL.name;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);a.remove()},4000)}
function dlShare(){if(!DL)return;var fl=new File([DL.blob],DL.name,{type:DL.blob.type||'application/octet-stream'});navigator.share({files:[fl],title:DL.name}).catch(function(e){if(e&&e.name!=='AbortError')dlSave()})}
function dlView(){if(!DL)return;var u=URL.createObjectURL(DL.blob);window.open(u,'_blank')}

/* ───── lluvia de la zona: valores normales de AEMET (1981-2010) de una estación de cada provincia de Castilla-La Mancha ───── */
var CLIMA={2:{a:353,o:110,i:77,p:110,v:55,src:'AEMET, Albacete Base Aérea (8175), valores normales 1981-2010'},
  13:{a:402,o:124,i:124,p:117,v:36,src:'AEMET, Ciudad Real (4121), valores normales 1981-2010'},
  16:{a:501,o:150,i:136,p:145,v:71,src:'AEMET, Cuenca (8096), valores normales 1981-2010'},
  19:{a:468,o:136,i:90,p:144,v:96,src:'AEMET, Molina de Aragón (3013), valores normales 1981-2010'},
  45:{a:342,o:105,i:92,p:106,v:40,src:'AEMET, Toledo (3260B), valores normales 1982-2010'}};
var PROV_N={2:'Albacete',13:'Ciudad Real',16:'Cuenca',19:'Guadalajara',45:'Toledo'};
function provDe(f){var m={};fincaRecs(f.id).forEach(function(r){var p=+String(r.ref).split(':')[0];m[p]=(m[p]||0)+haC(r)});var b=Object.keys(m).sort(function(a,c){return m[c]-m[a]})[0];return b?+b:null}
function climaDe(f){var p=provDe(f);return Object.assign({},CLIMA[p]||CLIMA[2])}

/* ───── catálogo general de abonos (riqueza típica del tipo de abono; la de la etiqueta manda) ───── */
var CAT=[
  ['Nitrogenados',[['Urea 46 %',46,0,0],['Nitrato amónico 33,5 %',33.5,0,0],['Nitrato amónico cálcico (NAC) 27 %',27,0,0],['Nitrosulfato amónico 26 %',26,0,0],['Sulfato amónico 21 %',21,0,0],['Solución nitrogenada N32',32,0,0],['Nitrato cálcico 15,5 %',15.5,0,0]]],
  ['Fosfatados',[['Superfosfato simple 18 %',0,18,0],['Superfosfato triple 46 %',0,46,0],['Fosfato monoamónico (MAP) 12-61-0',12,61,0],['Fosfato diamónico (DAP) 18-46-0',18,46,0],['Ácido fosfórico 54 %',0,54,0],['Fosfato monopotásico (MKP) 0-52-34',0,52,34]]],
  ['Potásicos',[['Cloruro potásico 60 %',0,0,60],['Sulfato potásico 50 %',0,0,50],['Nitrato potásico 13-0-46',13,0,46]]],
  ['Complejos',[['Complejo 8-15-15',8,15,15],['Complejo 9-18-27',9,18,27],['Complejo 12-12-17',12,12,17],['Complejo 15-15-15',15,15,15],['Complejo 20-10-10',20,10,10]]]];
function catGen(){var o=[];CAT.forEach(function(g){g[1].forEach(function(x){o.push({n:x[0],rn:x[1],rp:x[2],rk:x[3],org:0})})});return o}
/* «Mis abonos»: los que ya se han usado en algún plan de este aparato con otro nombre */
function catMios(){var gen={},o=[],seen={};catGen().forEach(function(x){gen[noAcc(x.n).toLowerCase()]=1});
  S.fincas.forEach(function(f){(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){((u.pg||{}).rows||[]).forEach(function(q){var k=noAcc(q.n||'').toLowerCase().trim();if(!k||gen[k]||seen[k])return;if(!(+q.rn||+q.rp||+q.rk))return;seen[k]=1;o.push({n:q.n,rn:+q.rn||0,rp:+q.rp||0,rk:+q.rk||0,org:q.org===1?1:0})})})})});
  return o.sort(function(a,b){return a.n.localeCompare(b.n,'es')})}
function catSel(i,j,q){var mi=catMios(),cur=noAcc(q.n||'').toLowerCase().trim();
  var opt=function(x,k){var on=noAcc(x.n).toLowerCase()===cur;return '<option value="'+k+'"'+(on?' selected':'')+'>'+esc(x.n)+'</option>'};
  var h='<select class="trsel pa-cat" data-u="'+i+'" data-j="'+j+'" style="width:100%" aria-label="Elegir abono"><option value="">Elegir un abono…</option>';
  if(mi.length)h+='<optgroup label="Mis abonos">'+mi.map(function(x,k){return opt(x,'m'+k)}).join('')+'</optgroup>';
  var k0=0;CAT.forEach(function(g){h+='<optgroup label="'+esc(g[0])+'">'+g[1].map(function(x){var o=opt({n:x[0]},'g'+k0);k0++;return o}).join('')+'</optgroup>'});
  return h+'<option value="otro">Otro: lo escribo abajo</option></select>'}

/* ───── cultivos ───── */
var CULTS=['Pistacho','Almendro','Olivo','Nogal','Viña','Melocotonero','Otros frutales','Trigo','Trigo duro','Cebada','Avena','Centeno','Triticale','Maíz','Barbecho','Otro'];
function cultSel(attr,cur){var L=CULTS.slice();if(cur&&L.indexOf(cur)<0)L.unshift(cur);return '<select class="trsel ap-in" '+attr+' style="width:100%"><option value="">Sin cultivo</option>'+L.map(function(c){return '<option'+(c===cur?' selected':'')+'>'+esc(c)+'</option>'}).join('')+'</select>'}

/* ───── recintos: código SIGPAC ───── */
function refParse(s){var a=String(s||'').trim().split(/[^0-9]+/).filter(Boolean).map(Number);
  if(a.length===5)a=[a[0],a[1],0,0,a[2],a[3],a[4]];
  if(a.length!==7||a.some(function(x){return !isFinite(x)})||!a[0]||a[0]>52||!a[1])return null;return a.join(':')}
function recNew(ref,o){o=o||{};return {ref:ref,ha:+o.ha||0,hc:+o.hc>0?+o.hc:undefined,cultivo:o.cultivo||'',riego:o.riego||'',variedad:o.variedad||'',nombre:o.nombre||'',fincaId:o.fincaId}}
function addRecs(f,list){var added=0,skip=[];list.forEach(function(o){var ex=S.recintos[o.ref];if(ex&&ex.fincaId!==f.id&&finca(ex.fincaId)){skip.push(o.ref);return}if(ex&&ex.fincaId===f.id){skip.push(o.ref);return}o.fincaId=f.id;S.recintos[o.ref]=recNew(o.ref,o);added++});return {added:added,skip:skip}}
/* SIGPAC: superficie y pendiente media del recinto (consulta pública, necesita internet) */
var SIGB={};
function sigLeer(f,refs,cb){
  refs=refs.filter(function(x){return S.recintos[x]&&!SIGB[x]}); if(!refs.length){if(cb)cb(0,0);return}
  if(navigator.onLine===false){toast('Sin internet no puedo leer el SIGPAC. Prueba luego.');if(cb)cb(0,refs.length);return}
  var idx=0,ok=0,ko=0; refs.forEach(function(x){SIGB[x]=1});
  var one=function(){if(idx>=refs.length)return Promise.resolve();var x=refs[idx++];
    return fetch('https://sigpac-hubcloud.es/servicioconsultassigpac/query/recinfo/'+x.split(':').join('/')+'.json').then(function(r){if(!r.ok)throw 1;return r.json()}).then(function(j){var o=Array.isArray(j)?j[0]:j;if(!o)throw 1;var r=S.recintos[x];if(!r)return;
        if(o.superficie!=null&&!(r.ha>0))r.ha=+o.superficie; if(o.superficie!=null)r.hs=+o.superficie; if(o.pendiente_media!=null)r.pend=+o.pendiente_media; if(o.uso_sigpac)r.uso=o.uso_sigpac; if(o.coef_regadio!=null)r.coef=+o.coef_regadio; ok++})
      .catch(function(){ko++}).then(function(){delete SIGB[x];return one()})};
  Promise.all([one(),one(),one(),one()]).then(function(){pendAPlanes(f);save();if(cb)cb(ok,ko);render()});
}
/* la pendiente del SIGPAC pasa a la tabla de recintos de cada unidad si allí está vacía */
function pendAPlanes(f){(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){(u.refs||[]).forEach(function(x){var r=S.recintos[x];if(!r||r.pend==null)return;u.rc=u.rc||{};var c=u.rc[x]=u.rc[x]||{};if(!abHas(c.pe))c.pe=r.pend})})})}

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
      PACI={kind:'pac',rows:o.rows.map(function(r){return {ref:[r.prov,r.muni,r.agr,r.zona,r.pol,r.par,r.rec].join(':'),mname:r.mname,ha:r.sigHa,hc:Math.abs(r.sigHa-r.decHa)>0.00005&&r.decHa>0?r.decHa:undefined,cultivo:pacCrop(r.prod)||pacCropVar(r.vari)||titleCase2(r.prod),riego:r.riego||'',variedad:titleCase2(r.vari)}}),year:o.year,nif:nl.length===1?nl[0]:'',name:'PAC '+(o.year||'')};
      pacHoja();
    }).catch(function(e){PACI=null;toast(e&&e.message==='net'?'Para leer el PDF hace falta internet la primera vez.':'No he encontrado la tabla de recintos en ese PDF.')});
  };
  rd.readAsArrayBuffer(file);
}
function csvLeer(file){
  var rd=new FileReader();
  rd.onload=function(){
    var t=String(rd.result||'').replace(/^﻿/,''), lines=t.split(/\r?\n/).filter(function(l){return l.trim()});
    var sep=(lines[0]||'').split(';').length>=(lines[0]||'').split(',').length?';':(/\t/.test(lines[0])?'\t':',');
    var hi=-1,H=null;for(var i=0;i<Math.min(lines.length,10);i++){var c=lines[i].split(sep).map(function(x){return noAcc(x).toLowerCase().trim()});if(c.some(function(x){return /^pol/.test(x)})&&c.some(function(x){return /^parc/.test(x)})){hi=i;H=c;break}}
    if(hi<0){toast('No encuentro las columnas «polígono» y «parcela» en la tabla.');return}
    var col=function(re){for(var k=0;k<H.length;k++)if(re.test(H[k]))return k;return -1};
    var cP=col(/^prov/),cM=col(/^(muni|mun\b|cod.*mun|termino)/),cA=col(/^agr/),cZ=col(/^zona/),cPo=col(/^pol/),cPa=col(/^parc/),cR=col(/^rec/),cH=col(/^(sup|ha\b|hect)/),cC=col(/^(cult|uso|prod)/),cRg=col(/^(rieg|regad|sist|s\/r)/),cV=col(/^vari/);
    var num=function(v){var n=numEs(v);if(n===null||isNaN(n))n=parseFloat(String(v).replace(',','.'));return isFinite(n)?n:0};
    var rows=[];lines.slice(hi+1).forEach(function(l){var c=l.split(sep).map(function(x){return x.replace(/^"|"$/g,'').trim()});var g=function(k){return k>=0?c[k]||'':''};
      var ref=[g(cP),g(cM),g(cA)||0,g(cZ)||0,g(cPo),g(cPa),g(cR)||1].map(function(x){return parseInt(x,10)||0}).join(':');if(!refParse(ref))return;
      var rg=noAcc(g(cRg)).toLowerCase();rows.push({ref:refParse(ref),ha:num(g(cH)),cultivo:titleCase2(g(cC)),riego:/^r|regad/.test(rg)?'regadio':(/^s|secan/.test(rg)?'secano':''),variedad:titleCase2(g(cV))})});
    if(!rows.length){toast('La tabla no tiene filas con provincia, municipio, polígono y parcela.');return}
    PACI={kind:'csv',rows:rows,name:'Mi finca'};pacHoja();
  };
  rd.readAsText(file);
}
function pacHoja(){clearTimeout(toastT);$("#toast").innerHTML="";
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, rows=PACI.rows, n=rows.filter(function(r){return !r.off}).length, tot=rows.reduce(function(a,r){return a+(r.off?0:(r.ha||0))},0);
  var h='<h2>'+(PACI.kind==='pac'?'PAC '+esc(PACI.year||''):'Tabla de recintos')+'</h2><p class="small">He encontrado <b>'+rows.length+'</b> recintos · '+ha2(tot)+' ha marcadas. Toca uno para quitarlo o ponerlo.</p>';
  h+='<div>'+rows.map(function(r,i){var ex=S.recintos[r.ref];return '<div class="pacrow'+(r.off?' off':'')+'" data-act="ap-pac-tog" data-i="'+i+'" role="checkbox" aria-checked="'+(!r.off)+'"><i>'+(r.off?'':'✓')+'</i><div><b>'+esc(refLabel(r.ref))+'</b>'+(r.mname?' · '+esc(titleCase2(r.mname)):'')+'<small>'+esc([r.cultivo||'Sin cultivo',r.variedad,r.riego==='regadio'?'Regadío':(r.riego==='secano'?'Secano':''),ha2(r.ha)+' ha'+(r.hc?' (declaradas '+ha2(r.hc)+')':'')].filter(Boolean).join(' · '))+(ex?' · ya está en «'+esc((finca(ex.fincaId)||{}).nombre||'')+'»':'')+'</small></div></div>'}).join('')+'</div>';
  if(!f)h+='<label class="lbl" for="pac-name">Nombre de la finca</label><input id="pac-name" class="in" value="'+esc(PACI.name)+'" autocomplete="off">';
  if(PACI.nif)h+='<p class="small">NIF leído en el PDF: <b>'+esc(PACI.nif)+'</b> (se pone como NIF del titular si está vacío; compruébalo).</p>';
  h+='<div class="acts"><button class="btn primary xl" data-act="ap-pac-go"'+(n?'':' disabled')+'>'+(f?'Añadir ':'Crear la finca con ')+n+(n===1?' recinto':' recintos')+'</button><button class="btn" data-act="ap-close">Cancelar</button></div>';
  sheet(h);
}
function pacGo(){
  if(!PACI)return; var f=UI.v==='finca'?finca(UI.fid):null, ni=$('#pac-name');
  undoPush();
  if(!f){f=fincaNueva((ni&&ni.value.trim())||PACI.name||'Mi finca')}
  if(PACI.nif){f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};if(!f.datos.tit.nif)f.datos.tit.nif=PACI.nif}
  var R=addRecs(f,PACI.rows.filter(function(r){return !r.off}).map(function(r){return Object.assign({},r)}));
  var refs=fincaRecs(f.id).map(function(r){return r.ref});
  sheetClose(); UI.v='finca';UI.fid=f.id;UI.step='rec'; save(); saveUI(); render();
  toast(R.added+(R.added===1?' recinto añadido':' recintos añadidos')+(R.skip.length?' · '+R.skip.length+' ya estaban':'')+'. Leo la pendiente del SIGPAC…');
  sigLeer(f,refs,function(ok,ko){if(ok||ko)toast('SIGPAC: '+ok+' leídos'+(ko?', '+ko+' sin respuesta':''))});
}
function fincaNueva(nm){var f={id:uid(),nombre:nm||'Mi finca',datos:{tit:{}},c:Date.now()};S.fincas.push(f);return f}

/* ───── copias en archivo ───── */
function copiaBlob(fs){
  var ids={};fs.forEach(function(f){ids[f.id]=1});var rec={};Object.keys(S.recintos).forEach(function(k){if(ids[S.recintos[k].fincaId])rec[k]=S.recintos[k]});
  return new Blob([JSON.stringify({app:'plan-abonado',v:1,fecha:new Date().toISOString(),datos:{fincas:fs,recintos:rec,x:{asesor:S.x.asesor||{}}}},null,1)],{type:'application/json'})}
function hoyIso(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2)}
function copiaTodo(){if(!S.fincas.length){toast('Aún no hay nada que guardar.');return}S.copia=Date.now();save();showDl(copiaBlob(S.fincas),'plan-abonado-copia-'+hoyIso()+'.json','Copia lista','Con este archivo recuperas tus fincas y planes en este u otro aparato («Abrir una copia»).');render()}
function copiaFinca(f){showDl(copiaBlob([f]),'plan-abonado-'+safeName(f.nombre)+'-'+hoyIso()+'.json','Copia de la finca lista','Para pasar esta finca a otro aparato o a otra persona: «Abrir una copia».')}
var COPIA=null;
function copiaAbrir(file){var rd=new FileReader();rd.onload=function(){try{var o=JSON.parse(rd.result);if(!o||o.app!=='plan-abonado'||!o.datos||!Array.isArray(o.datos.fincas))throw 1;COPIA=o.datos;
    var nf0=o.datos.fincas.length, ya=o.datos.fincas.filter(function(f){return finca(f.id)}).length;
    sheet('<h2>Abrir copia</h2><p class="small">'+nf0+(nf0===1?' finca':' fincas')+': '+esc(o.datos.fincas.map(function(f){return f.nombre}).join(', '))+'.</p>'+(ya?'<p class="small">'+ya+(ya===1?' ya está':' ya están')+' en este aparato: se cambian por las de la copia.</p>':'')+
      '<div class="acts"><button class="btn primary xl" data-act="ap-copia-join">Añadir a lo que hay</button>'+(S.fincas.length?'<button class="btn danger" data-act="ap-copia-repl">Cambiar todo por la copia</button>':'')+'<button class="btn" data-act="ap-close">Cancelar</button></div>');
  }catch(e){toast('Ese archivo no es una copia de esta herramienta.')}};rd.readAsText(file)}
function copiaAplicar(repl){if(!COPIA)return;undoPush();
  if(repl){S.fincas=[];S.recintos={}}
  COPIA.fincas.forEach(function(f){S.fincas=S.fincas.filter(function(x){return x.id!==f.id});Object.keys(S.recintos).forEach(function(k){if(S.recintos[k].fincaId===f.id)delete S.recintos[k]});S.fincas.push(f)});
  Object.keys(COPIA.recintos||{}).forEach(function(k){S.recintos[k]=COPIA.recintos[k]});
  if(COPIA.x&&COPIA.x.asesor&&(repl||!Object.keys(S.x.asesor||{}).length))S.x.asesor=COPIA.x.asesor;
  COPIA=null;S.copia=Date.now();save();sheetClose();UI={v:'inicio',fid:null,step:'rec'};saveUI();render();undoToast('Copia abierta')}

/* ───── ejemplo inventado (para probar) ───── */
function ejemplo(){
  undoPush(); var f=fincaNueva('Finca de ejemplo'); f.datos.tit={n:'Titular de ejemplo',mun:'Municipio',prov:'Albacete'};
  addRecs(f,[{ref:'2:900:0:0:1:1:1',ha:5.2,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman',nombre:'La Hoya'},{ref:'2:900:0:0:1:2:1',ha:3.1,cultivo:'Pistacho',riego:'regadio',variedad:'Kerman'},{ref:'2:900:0:0:3:7:2',ha:6.4,cultivo:'Almendro',riego:'secano',variedad:'Guara'}]);
  var y=pabY0(), p={id:uid(),y:String(y),cl:climaDe(f),ud:pabAuto(f),c:Date.now(),m:Date.now(),fe:hoyIso()}; (f.pab=[]).push(p);
  p.ud.forEach(function(u){u.prec=u.ck==='pistacho'?'Pistacho':'Almendro';u.mo={obj:'mantener'};u.res={};PAB_RES.forEach(function(x){u.res[x[0]]={a:1}});u.mit={B1:{a:1}};
    u.su={f:hoyIso(),lab:'Laboratorio de ejemplo',inf:'0001',prof:'0-30',mo:1.2,ph:8.1,n:900,p:14,k:260,tex:'franca'};
    if(u.ck==='pistacho'){u.ne={mo:'cos',pr:1500,ed:1};u.ag={m3:2500,no3:20};u.mit.B4={a:1};u.pg={src:'own',rows:[{id:uid(),mo:'Marzo',n:'Nitrato potásico 13-0-46',d:80,rn:13,rp:0,rk:46,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Abril-junio',n:'Nitrato amónico 33,5 %',d:120,rn:33.5,rp:0,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'},{id:uid(),mo:'Abril',n:'Fosfato monoamónico (MAP) 12-61-0',d:40,rn:12,rp:61,rk:0,fo:'Fertirrigación',eq:'Cabezal de riego'}]}}
    else{u.ne={mo:'cos',pr:600,ed:1};u.pg={src:'own',rows:[{id:uid(),mo:'Febrero-marzo',n:'Complejo 15-15-15',d:150,rn:15,rp:15,rk:15,fo:'Suelo',eq:'Abonadora'}]}}});
  UI={v:'finca',fid:f.id,step:'rec'}; save(); saveUI(); render(); undoToast('Finca de ejemplo creada (datos inventados)');
}

/* ───── pantallas ───── */
var STEPS=[['rec','Recintos'],['ud','Unidades'],['su','Suelo y agua'],['ne','Cosecha'],['pg','Programa'],['rv','Revisar y PDF']];
function render(){
  var a=$('#app'), sy=window.scrollY;
  if(UI.v==='finca'&&finca(UI.fid)){a.innerHTML=fincaHTML(finca(UI.fid));$('#bar').innerHTML=barHTML();$('#bar').hidden=false}
  else{UI.v='inicio';UI.fid=null;a.innerHTML=inicioHTML();$('#bar').hidden=true}
  EXP.fid=UI.fid; document.title=UI.v==='finca'?finca(UI.fid).nombre+' · Plan de abonado':'Plan de abonado';
  if(window.__scrollTop){window.scrollTo(0,0);window.__scrollTop=false}else window.scrollTo(0,sy);
  stkLab(); inThou();
}
function inicioHTML(){
  var h='<h1>Plan de abonado</h1><div class="sub">Real Decreto 1051/2022 y Decreto 80/2025 de Castilla-La Mancha</div>';
  var sinCopia=S.fincas.length&&S.mod>(S.copia||0)+1000;
  if(sinCopia)h+='<div class="warn">Hay cambios sin copia. Lo que metes se guarda solo en este aparato: guarda una copia en archivo de vez en cuando.<div class="acts" style="margin-top:8px"><button class="btn sm" data-act="ap-copia">Guardar copia</button></div></div>';
  if(!S.fincas.length)h+='<div class="card"><h3>Empieza aquí</h3><p class="small" style="margin:6px 0 0">Mete el PDF de la solicitud de la PAC y te salen solos los recintos, con su cultivo, superficie y riego. Después rellenas, paso a paso, la cosecha esperada, el análisis de suelo, el agua y los abonos previstos. Al final sacas el plan en PDF y Excel.</p><p class="small">Todo se queda en este aparato: no hay cuentas ni servidor.</p></div>';
  S.fincas.slice().sort(function(a,b){return String(a.nombre).localeCompare(String(b.nombre),'es')}).forEach(function(f){var rs=fincaRecs(f.id),t=(f.datos||{}).tit||{};
    h+='<div class="card tap" data-act="ap-open" data-id="'+esc(f.id)+'" role="button" tabindex="0"><h3>'+esc(f.nombre)+'</h3><div class="small">'+esc([t.n,rs.length+(rs.length===1?' recinto':' recintos'),haTxt(rs)].filter(Boolean).join(' · '))+'</div>'+
      ((f.pab||[]).length?'<div class="small" style="margin-top:4px">Planes: '+pabList(f).map(function(p){return 'campaña '+pabCamp(p.y)}).join(', ')+'</div>':'<div class="small" style="margin-top:4px">Sin plan todavía</div>')+'</div>'});
  h+='<div class="acts"><button class="btn primary xl wide" data-act="ap-nueva">+ Nueva finca</button></div>';
  h+='<div class="acts"><button class="btn" data-act="ap-abrir">Abrir una copia</button>'+(S.fincas.length?'<button class="btn" data-act="ap-copia">Guardar copia de todo</button>':'')+'</div>';
  h+='<details class="more"><summary>Cómo funciona</summary><div class="small"><p>1. <b>Recintos:</b> del PDF de la PAC, de una tabla (CSV) o a mano con su código SIGPAC.</p><p>2. <b>Unidades:</b> los recintos se juntan solos por cultivo y riego; marca si están en zona vulnerable.</p><p>3. <b>Suelo y agua</b>, 4. <b>Cosecha</b> y 5. <b>Programa</b> de abonos.</p><p>6. <b>Revisar:</b> balance de nutrientes, comprobación del plan y PDF o Excel.</p><p>Los límites de nitrógeno son los del Decreto 80/2025 de Castilla-La Mancha. Para pistacho y nogal las extracciones son de criterio técnico (no hay tabla oficial). El plan lo firma el titular y, cuando sea obligatorio, el asesor.</p><p>En el iPhone: Compartir › «Añadir a pantalla de inicio» para tenerla como una app y que funcione sin cobertura.</p></div></details>';
  h+='<div class="foot"><a href="privacidad.html">Privacidad</a> · <a href="condiciones.html">Condiciones</a> · © 2026 Ramón Alcalá<br>Versión '+esc(VERSION)+'</div>';
  return h;
}
function stepIdx(){for(var i=0;i<STEPS.length;i++)if(STEPS[i][0]===UI.step)return i;return 0}
function barHTML(){var i=stepIdx();return '<button class="btn" data-act="ap-prev">'+(i===0?'‹ Mis fincas':'‹ Atrás')+'</button>'+(i<STEPS.length-1?'<button class="btn primary" data-act="ap-next">Siguiente ›</button>':'<button class="btn primary" data-act="ap-home">Terminar</button>')}
function fincaHTML(f){
  var p=pabCur(f), h='<div class="top"><div class="row"><button class="back" data-act="ap-home" aria-label="Volver a mis fincas">‹ Fincas</button><div class="ttl">'+esc(f.nombre)+'</div>'+(UNDO.length?'<button class="btn sm lk" data-act="ap-undo">Deshacer</button>':'')+'</div>';
  h+='<div class="steps" role="tablist">'+STEPS.map(function(s,i){return '<button type="button" class="chip'+(s[0]===UI.step?' on':'')+'" data-act="ap-step" data-s="'+s[0]+'" role="tab" aria-selected="'+(s[0]===UI.step)+'"><b>'+(i+1)+'</b>'+s[1]+'</button>'}).join('')+'</div></div>';
  if(UI.step==='rec')return h+recHTML(f);
  pabSync(f);
  if(!p)return h+nuevoPlanHTML(f,true);
  h+=campHTML(f,p);
  var Cs=(p.ud||[]).map(function(u){return pabCalc(f,u,p)});
  if(UI.step==='ud')h+=udStepHTML(f,p,Cs);
  else if(UI.step==='rv')h+=rvStepHTML(f,p,Cs);
  else{
    var T={su:['Suelo y agua','Análisis de suelo del recinto representativo de cada unidad y, en regadío, el agua de riego.'],ne:['Cosecha y aportes','Producción esperada (o años de la plantación) y lo que ya aportan el suelo, el estiércol u otros.'],pg:['Programa de abonado','Abonos previstos en la campaña, medidas para reducir el amoniaco y restricciones de aplicación.']}[UI.step];
    h+='<h2>'+T[0]+'</h2><div class="small">'+T[1]+'</div>';
    (p.ud||[]).forEach(function(u,i){h+=udBox(f,p,u,Cs[i],i,UI.step)});
    if(!(p.ud||[]).length)h+='<div class="note">No hay unidades: añádelas en el paso 2.</div>';
  }
  return h;
}
function campHTML(f,p){var l=pabList(f);
  return '<div class="plhead"><div class="small">Plan de la campaña <b>'+esc(pabCamp(p.y))+'</b> · '+(p.ud||[]).length+((p.ud||[]).length===1?' unidad':' unidades')+'</div><div><select class="trsel ap-camp" aria-label="Campaña">'+l.map(function(x){return '<option value="'+esc(x.id)+'"'+(x.id===p.id?' selected':'')+'>Campaña '+esc(pabCamp(x.y))+'</option>'}).join('')+'<option value="new">+ Plan de otra campaña…</option></select></div></div>'+(PAB.ask===f.id?nuevoPlanHTML(f,false):'')}
function nuevoPlanHTML(f,first){
  var y0=pabY0(), ex={}; (f.pab||[]).forEach(function(x){ex[+x.y]=1}); var yo=[]; for(var yy=y0-2;yy<=y0+2;yy++)if(!ex[yy])yo.push(yy);
  var ys=yo.indexOf(y0)>=0?y0:(yo.filter(function(v){return v>y0})[0]||yo[yo.length-1]), n=fincaRecs(f.id).length;
  if(!n)return '<div class="note">Primero mete los recintos en el paso 1.</div>';
  return '<div class="card"><h3>'+(first?'Crear el plan de abonado':'Plan de otra campaña')+'</h3><p class="small" style="margin:6px 0 0">Los '+n+' recintos se juntan solos en unidades de producción por cultivo y riego. '+((f.pab||[]).length?'Del plan anterior se copian el suelo, el agua, la zona vulnerable y la pendiente.':'')+'</p>'+
    '<label class="lbl" for="pa-ny">Campaña (de noviembre a octubre)</label><select id="pa-ny" class="trsel" style="width:100%">'+yo.map(function(v){return '<option value="'+v+'"'+(v===ys?' selected':'')+'>'+pabCamp(v)+'</option>'}).join('')+'</select>'+
    '<div class="acts"><button class="btn primary xl" data-act="pa-new-ok">Crear el plan</button>'+(first?'':'<button class="btn" data-act="ap-newplan-no">Cancelar</button>')+'</div></div>';
}
function inp(attr,val,ph,mode){return '<input class="in ap-in" '+attr+(mode?' inputmode="'+mode+'"':'')+' value="'+esc(val==null?'':(typeof val==='number'?abEd(val):val))+'"'+(ph?' placeholder="'+esc(ph)+'"':'')+' autocomplete="off">'}
function recHTML(f){
  var t=(f.datos||{}).tit||{}, rs=fincaRecs(f.id), h='<h2>Finca y titular</h2><div class="pa-g">';
  h+='<div>'+pabF('Nombre de la finca',inp('data-e="nombre"',f.nombre))+'</div>';
  [['n','Titular (nombre o razón social)'],['nif','NIF / CIF'],['dir','Dirección fiscal'],['cp','Código postal'],['mun','Municipio'],['prov','Provincia'],['rn','Representante (si firma otro)'],['rnif','NIF del representante']].forEach(function(x){h+='<div>'+pabF(x[1],inp('data-e="tit.'+x[0]+'"',t[x[0]]))+'</div>'});
  h+='</div>';
  h+='<h2>Recintos <span class="small">('+rs.length+' · '+esc(haTxt(rs))+')</span></h2>';
  h+='<div class="acts" style="margin-top:4px"><button class="btn primary" data-act="ap-addrec">+ Añadir recintos</button>'+(rs.length?'<button class="btn" data-act="ap-sig">Leer pendiente del SIGPAC</button>':'')+'</div>';
  if(!rs.length)h+='<div class="note">Añade los recintos desde el PDF de la PAC, una tabla CSV o a mano con su código SIGPAC.</div>';
  h+='<div class="card recs">'+rs.map(function(r){var k=esc(r.ref);
    return '<div class="rec"><div class="rh"><b>'+esc(refLabel(r.ref))+'</b><button class="btn sm danger" data-act="ap-recdel" data-ref="'+k+'" aria-label="Quitar recinto">✕</button></div>'+
      '<div class="small">'+esc([r.uso?'Uso SIGPAC '+r.uso:'',r.pend!=null?'pendiente media '+trTxt(r.pend,1)+' %':'pendiente sin leer',r.hs!=null&&Math.abs(r.hs-(r.ha||0))>0.00005?'SIGPAC '+ha4(r.hs)+' ha':''].filter(Boolean).join(' · '))+'</div>'+
      '<div class="pa-g"><div>'+pabF('Cultivo',cultSel('data-r="'+k+'" data-k="cultivo"',r.cultivo))+'</div><div>'+pabF('Variedad',inp('data-r="'+k+'" data-k="variedad"',r.variedad))+'</div>'+
      '<div>'+pabF('Riego',pabChipsAp(r.ref,'riego',r.riego==='regadio'?'regadio':'secano',[['secano','Secano'],['regadio','Regadío']]))+'</div>'+
      '<div>'+pabF('ha SIGPAC',inp('data-r="'+k+'" data-k="ha" data-num="1"',r.ha,'','decimal'))+'</div><div>'+pabF('ha cultivadas (si son menos)',inp('data-r="'+k+'" data-k="hc" data-num="1"',r.hc,ha4(r.ha||0),'decimal'))+'</div><div>'+pabF('Nombre del paraje (opcional)',inp('data-r="'+k+'" data-k="nombre"',r.nombre))+'</div></div></div>'}).join('')+'</div>';
  if((f.pab||[]).length)h+='<div class="small" style="margin-top:6px">Los recintos que añadas pasan solos a la unidad de su cultivo y riego del plan.</div>';
  h+='<h3>Esta finca</h3><div class="acts"><button class="btn" data-act="ap-copiaf">Guardar copia de esta finca</button><button class="btn danger" data-act="ap-fdel">Borrar la finca</button></div>';
  return h;
}
function pabChipsAp(ref,k,cur,L){return '<div class="chips">'+L.map(function(x){return '<button type="button" class="chip'+(String(cur)===String(x[0])?' on':'')+'" data-act="ap-recset" data-ref="'+esc(ref)+'" data-k="'+k+'" data-v="'+esc(x[0])+'">'+x[1]+'</button>'}).join('')+'</div>'}
function addRecHoja(){var f=finca(UI.fid);
  sheet('<h2>Añadir recintos</h2><div class="acts" style="flex-direction:column"><button class="btn primary xl" data-act="ap-pdf-pick">Del PDF de la solicitud de la PAC</button><button class="btn xl" data-act="ap-csv-pick">De una tabla (CSV)</button></div>'+
    '<h3>A mano</h3><label class="lbl" for="ap-ref">Código SIGPAC</label><input id="ap-ref" class="in" inputmode="numeric" placeholder="02:012:0:0:1:16:5" autocomplete="off"><div class="small">Provincia, municipio, agregado, zona, polígono, parcela y recinto (o solo provincia, municipio, polígono, parcela y recinto).</div>'+
    '<div class="pa-g"><div>'+pabF('Cultivo',cultSel('id="ap-rc"',''))+'</div><div>'+pabF('Riego','<select id="ap-rr" class="trsel" style="width:100%"><option value="secano">Secano</option><option value="regadio">Regadío</option></select>')+'</div><div>'+pabF('ha (vacío: las del SIGPAC)','<input id="ap-rh" class="in" inputmode="decimal" autocomplete="off">')+'</div></div>'+
    '<div class="acts"><button class="btn primary" data-act="ap-ref-add">Añadir este recinto</button><button class="btn" data-act="ap-close">Cerrar</button></div>'+
    '<p class="small">La tabla CSV necesita columnas de provincia, municipio, polígono, parcela y recinto (y si quieres superficie, cultivo y riego).</p>')}
function refAdd(){var f=finca(UI.fid), ref=refParse(($('#ap-ref')||{}).value);if(!ref){toast('Ese código SIGPAC no está completo.');return}
  var hv=($('#ap-rh')||{}).value, ha=hv?numEs(hv):0; if(hv&&(ha===null||isNaN(ha))){toast('Pon las ha con coma decimal, por ejemplo 2,35.');return}
  undoPush(); var R=addRecs(f,[{ref:ref,ha:ha||0,cultivo:($('#ap-rc')||{}).value||'',riego:($('#ap-rr')||{}).value||'secano'}]);
  if(!R.added){toast('Ese recinto ya está'+(S.recintos[ref]&&S.recintos[ref].fincaId!==f.id?' en otra finca':'')+'.');return}
  save(); $('#ap-ref').value=''; $('#ap-rh').value=''; toast('Recinto '+refLabel(ref)+' añadido. Leo el SIGPAC…'); render();
  sigLeer(f,[ref],function(ok){if(!ok)toast('No he podido leer ese recinto en el SIGPAC: revisa el código o pon las ha a mano.')});
}
/* unidades: los apartados de cada unidad se reparten entre los pasos */
var SEG={ud:[/^(Recintos)/],su:[/^Suelo del recinto/,/^Agua de riego/],ne:[/^Necesidades del cultivo/,/^Aportes que ya hay/,/^Datos de inicio de campa/,/^Materia org/],pg:[/^Medidas para reducir/,/^Restricciones y condiciones/],rv:[/^Balance de nutrientes/,/^Comprobaci/]};
function udBox(f,p,u,C,i,step){
  if(PAB.open[u.id]===undefined)PAB.open[u.id]=(i===0);
  var html=pabUdHTML(f,p,u,C,i), e=html.indexOf('</summary>')+10, head=html.slice(0,e), body=html.slice(e,html.lastIndexOf('</details>'));
  body=body.replace(/<div class="acts" style="margin:10px 0 0"><button class="btn sm lk danger" data-act="pa-ud-del"[^]*?<\/div>$/,'');
  var parts=body.split(/(?=<h4 )/), first=parts[0], segs={};
  parts.slice(1).forEach(function(s){var t=s.replace(/^<h4[^>]*>/,'').replace(/<[^>]+>/g,'');for(var k in SEG)SEG[k].forEach(function(re,ix){if(re.test(t))(segs[k]||(segs[k]=[]))[ix]=s})});
  var h=head;
  if(step==='ud')h+=first+(segs.ud||[]).join('')+'<div class="acts"><button class="btn sm danger" data-act="pa-ud-del" data-u="'+i+'">Borrar unidad</button></div>';
  else if(step==='pg')h+=progHTML(f,p,u,C,i)+(u.d0===1?'':(segs.pg||[]).join(''));
  else h+=(segs[step]||[]).filter(Boolean).join('')||'<div class="small">Nada que rellenar aquí para esta unidad.</div>';
  return h+'</details>';
}
function progHTML(f,p,u,C,i){
  var P='ud.'+i+'.', pg=u.pg||{}, h='<h4 style="margin:14px 0 4px">Programa previsto <span class="small">(art. 6.4)</span></h4>'+pabF('¿Sin abonado esta campaña?',pabChips(P+'d0',u.d0===1?1:0,[[1,'Sí, dosis 0'],[0,'No, hay programa']]));
  if(u.d0===1)return h+'<div class="note">No se prevé abonar esta campaña (programa a dosis 0). No hace falta programa, análisis de suelo, medidas de amoniaco ni restricciones.</div>';
  var FO=['Suelo','Localizado','Fertirrigación','Foliar','Enterrado'];
  (pg.rows||[]).forEach(function(q,j){var R=P+'pg.rows.'+j+'.', o=C.rows[j], fo=FO.slice(); if(q.fo&&fo.indexOf(q.fo)<0)fo.unshift(q.fo);
    h+='<div class="pa-card"><div class="hd"><span>Aporte '+(j+1)+(o&&o.d?' · <span class="small">U.F. '+trTxt(o.uf.n,1)+'-'+trTxt(o.uf.p,1)+'-'+trTxt(o.uf.k,1)+'</span>':'')+'</span><button type="button" class="btn sm danger" data-act="pa-row-del" data-u="'+i+'" data-j="'+j+'" aria-label="Quitar aporte">✕</button></div>'+
      pabF('Abono',catSel(i,j,q))+'<div class="pa-g"><div>'+pabF('Nombre del producto',pabTx(R+'n',q.n,'Producto'))+'</div><div>'+pabF('Momento',pabTx(R+'mo',q.mo,'Marzo-abril'))+'</div><div>'+pabF('Dosis (kg/ha)',pabIn(R+'d',q.d))+'</div>'+
      '<div>'+pabF('N (%)',pabIn(R+'rn',q.rn))+'</div><div>'+pabF('P₂O₅ (%)',pabIn(R+'rp',q.rp))+'</div><div>'+pabF('K₂O (%)',pabIn(R+'rk',q.rk))+'</div>'+
      '<div>'+pabF('Forma de aplicación','<select class="trsel pa-in" data-p="'+esc(R+'fo')+'" style="width:100%">'+fo.map(function(x){return '<option'+(x===(q.fo||'Suelo')?' selected':'')+'>'+esc(x)+'</option>'}).join('')+'</select>')+'</div><div>'+pabF('Equipo',pabTx(R+'eq',q.eq,'Abonadora'))+'</div><div>'+pabF('Medida de mitigación',pabTx(R+'mi',q.mi,''))+'</div></div>'+
      pabF('¿Abono orgánico?',pabChips(R+'org',q.org===1?1:0,[[1,'Sí'],[0,'No']]))+'</div>'});
  if(!(pg.rows||[]).length)h+='<div class="small">Aún no hay aportes.</div>';
  h+='<div class="acts" style="margin-top:8px"><button class="btn sm primary" data-act="pa-row-add" data-u="'+i+'">+ Añadir aporte</button></div><div class="small" style="margin-top:6px">La riqueza del catálogo es la típica de ese tipo de abono: pon la de la etiqueta si es otra.</div>';
  return h+pabOrgHTML(u,C);
}
function climaHTML(f,p){var cl=p.cl||PAB_CL, sum=(+cl.o||0)+(+cl.i||0)+(+cl.p||0)+(+cl.v||0), pr=provDe(f);
  var h='<details class="more pa-box"'+(PAB.open.cl?' open':'')+' data-pa-o="cl"><summary><b>Lluvia de la zona</b> <span class="small">· '+nf(+cl.a||0,0)+' mm al año (art. 6.3)</span></summary>';
  h+='<div class="pa-g">'+[['a','Media anual (mm)'],['o','Otoño, sep-nov (mm)'],['i','Invierno, dic-feb (mm)'],['p','Primavera, mar-may (mm)'],['v','Verano, jun-ago (mm)']].map(function(x){return '<div>'+pabF(x[1],pabIn('cl.'+x[0],cl[x[0]]))+'</div>'}).join('')+'</div>'+pabF('Fuente',pabTx('cl.src',cl.src));
  h+='<div class="small" style="margin-top:6px">Suma de las estaciones: '+nf(sum,0)+' mm'+(abHas(cl.a)&&Math.abs(sum-cl.a)>2?' · <b>no cuadra con la media anual</b>':'')+'.</div>';
  h+=pabF('Poner los valores de AEMET de','<select class="trsel ap-clprov" style="width:100%"><option value="">Elegir provincia…</option>'+Object.keys(CLIMA).map(function(k){return '<option value="'+k+'">'+PROV_N[k]+(+k===pr?' (la de los recintos)':'')+' · '+CLIMA[k].src.replace(/^AEMET, /,'').replace(/, valores.*/,'')+'</option>'}).join('')+'</select>');
  return h+'</details>'}
function udStepHTML(f,p,Cs){
  var h='<h2>Unidades de producción</h2><div class="small">Recintos juntados por cultivo y riego. Puedes cambiar el nombre, mover recintos de una unidad a otra o añadir unidades. Marca si están en zona vulnerable a nitratos.</div>';
  h+=climaHTML(f,p);
  (p.ud||[]).forEach(function(u,i){h+=udBox(f,p,u,Cs[i],i,'ud')});
  h+='<div class="acts"><button class="btn" data-act="pa-ud-new">+ Unidad</button></div>';
  return h}
function rvStepHTML(f,p,Cs){
  var full=pabHTML(f), a=full.indexOf('<h3 style="margin-top:12px">Resumen</h3>'), b=full.indexOf('<details class="more pa-box"',a);
  var h='<h2>Revisar y sacar el plan</h2>';
  h+='<div class="acts" style="margin-top:6px"><button class="btn primary xl" data-act="ap-pdf">PDF del plan</button><button class="btn xl" data-act="ap-xlsx">Excel</button></div>';
  if(a>=0&&b>a)h+=full.slice(a,b).replace('<h3 style="margin-top:12px">Resumen</h3>','<h3>Resumen</h3>');
  (p.ud||[]).forEach(function(u,i){h+=udBox(f,p,u,Cs[i],i,'rv')});
  h+=pabPlanMore(f,p);
  h+='<h3>Este plan</h3><div class="acts"><button class="btn danger" data-act="pa-del">Borrar el plan de esta campaña</button></div>';
  return h}

/* ───── toques ───── */
document.addEventListener('click',function(ev){
  var b=ev.target.closest&&ev.target.closest('[data-act]'); if(!b||b.disabled)return; var a=b.dataset.act, f=finca(UI.fid);
  if(/^pa-/.test(a)){if(!f)return;EXP.fid=f.id;var had=(f.pab||[]).length;pabClick(b);
    if(a==='pa-new-ok'){var p=pabCur(f);if(p&&(f.pab||[]).length>had){if(!had)p.cl=climaDe(f);pendAPlanes(f);p.fe=p.fe||hoyIso();save();if(UI.step==='rec'){UI.step='ud';saveUI();window.__scrollTop=true}render()}}
    return}
  if(a==='ap-close'){sheetClose();return}
  if(a==='ap-undo'){undoDo();return}
  if(a==='ap-open'){UI={v:'finca',fid:b.dataset.id,step:'rec'};var f2=finca(b.dataset.id);if(f2&&(f2.pab||[]).length)UI.step='ud';saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-home'){UI={v:'inicio',fid:null,step:'rec'};saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-step'){UI.step=b.dataset.s;saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-prev'){var i=stepIdx();if(i===0){UI={v:'inicio',fid:null,step:'rec'}}else UI.step=STEPS[i-1][0];saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-next'){var j=stepIdx();UI.step=STEPS[Math.min(STEPS.length-1,j+1)][0];saveUI();window.__scrollTop=true;render();return}
  if(a==='ap-nueva'){sheet('<h2>Nueva finca</h2><div class="acts" style="flex-direction:column"><button class="btn primary xl" data-act="ap-pdf-pick">Leer el PDF de la solicitud de la PAC</button><button class="btn xl" data-act="ap-csv-pick">Leer una tabla de recintos (CSV)</button><button class="btn xl" data-act="ap-mano">Empezar a mano</button><button class="btn" data-act="ap-ejemplo">Probar con una finca de ejemplo</button><button class="btn lk" data-act="ap-close">Cancelar</button></div>');return}
  if(a==='ap-mano'){undoPush();var nf1=fincaNueva('Mi finca');UI={v:'finca',fid:nf1.id,step:'rec'};save();saveUI();sheetClose();render();addRecHoja();return}
  if(a==='ap-ejemplo'){sheetClose();ejemplo();return}
  if(a==='ap-pdf-pick'){pick('application/pdf,.pdf',pacLeer);return}
  if(a==='ap-csv-pick'){pick('.csv,.txt,text/csv,text/plain',csvLeer);return}
  if(a==='ap-abrir'){pick('.json,application/json',copiaAbrir);return}
  if(a==='ap-copia'){copiaTodo();return}
  if(a==='ap-copia-join'){copiaAplicar(false);return}
  if(a==='ap-copia-repl'){if(confirm('¿Cambiar todas las fincas de este aparato por las de la copia?'))copiaAplicar(true);return}
  if(a==='ap-pac-tog'){var r=PACI&&PACI.rows[+b.dataset.i];if(r){r.off=!r.off;pacHoja()}return}
  if(a==='ap-pac-go'){pacGo();return}
  if(a==='ap-dl-save'){dlSave();return}
  if(a==='ap-dl-share'){dlShare();return}
  if(a==='ap-dl-view'){dlView();return}
  if(!f)return;
  if(a==='ap-addrec'){addRecHoja();return}
  if(a==='ap-ref-add'){refAdd();return}
  if(a==='ap-sig'){toast('Leyendo el SIGPAC…');sigLeer(f,fincaRecs(f.id).map(function(r){return r.ref}),function(ok,ko){toast('SIGPAC: '+ok+' recintos leídos'+(ko?', '+ko+' sin respuesta':''))});return}
  if(a==='ap-recdel'){var ref=b.dataset.ref;if(!confirm('¿Quitar el recinto '+refLabel(ref)+' de esta finca?'))return;undoPush();delete S.recintos[ref];(f.pab||[]).forEach(function(p){(p.ud||[]).forEach(function(u){u.refs=(u.refs||[]).filter(function(x){return x!==ref});if(u.rc)delete u.rc[ref];if(u.rp===ref)u.rp=u.refs[0]||''})});save();render();undoToast('Recinto quitado');return}
  if(a==='ap-recset'){var rr=S.recintos[b.dataset.ref];if(!rr)return;rr[b.dataset.k]=b.dataset.v;save();render();return}
  if(a==='ap-copiaf'){copiaFinca(f);return}
  if(a==='ap-fdel'){if(!confirm('¿Borrar la finca «'+f.nombre+'» con sus recintos y planes?'))return;undoPush();Object.keys(S.recintos).forEach(function(k){if(S.recintos[k].fincaId===f.id)delete S.recintos[k]});S.fincas=S.fincas.filter(function(x){return x.id!==f.id});UI={v:'inicio',fid:null,step:'rec'};save();saveUI();render();undoToast('Finca borrada');return}
  if(a==='ap-newplan-no'){PAB.ask=null;render();return}
  if(a==='ap-pdf'||a==='ap-xlsx'){var p2=pabCur(f);if(!p2)return;if(a==='ap-xlsx'){pabOut(f,p2,'xlsx');return}
    b.disabled=true;toast('Preparando el PDF…');ensurePdf().then(function(){b.disabled=false;$('#toast').innerHTML='';pabOut(f,p2,'pdf')},function(){b.disabled=false;pdfNoListo()});return}
});
document.addEventListener('keydown',function(ev){if((ev.key==='Enter'||ev.key===' ')&&ev.target.classList&&ev.target.classList.contains('tap')){ev.preventDefault();ev.target.click()}});
function pick(accept,fn){var i=document.createElement('input');i.type='file';i.accept=accept;i.style.display='none';document.body.appendChild(i);i.onchange=function(){var fl=i.files&&i.files[0];i.remove();if(fl)fn(fl)};i.click()}
/* casillas propias (finca, titular, recintos), campaña, catálogo y lluvia */
document.addEventListener('change',function(ev){
  var t=ev.target; if(!t||!t.classList)return; var f=finca(UI.fid); if(!f)return;
  if(t.classList.contains('ap-in')){
    var v=String(t.value==null?'':t.value).trim();
    if(t.dataset.num){if(v===''){v=undefined}else{var n=numEs(v);if(n===null||isNaN(n)||n<0){toast('Pon un número (los decimales con coma, por ejemplo 2,35).');t.classList.add('bad');return}v=n}}
    if(t.dataset.e){var ks=t.dataset.e.split('.');if(ks[0]==='nombre'){if(!v){toast('La finca necesita un nombre.');render();return}f.nombre=v}else{f.datos=f.datos||{};f.datos.tit=f.datos.tit||{};if(v)f.datos.tit[ks[1]]=v;else delete f.datos.tit[ks[1]]}}
    else if(t.dataset.r){var r=S.recintos[t.dataset.r];if(!r)return;if(v===undefined||v==='')delete r[t.dataset.k];else r[t.dataset.k]=v}
    save(); pabRedraw(); return}
  if(t.classList.contains('ap-camp')){if(t.value==='new'){PAB.ask=f.id}else{PAB.cur[f.id]=t.value;PAB.ask=null}render();return}
  if(t.classList.contains('ap-clprov')){var p=pabCur(f),c=CLIMA[+t.value];if(p&&c){undoPush();p.cl=Object.assign({},c);p.m=Date.now();save();render();toast('Lluvia de '+PROV_N[+t.value]+' puesta')}return}
  if(t.classList.contains('pa-cat')){var p3=pabCur(f),u=p3&&p3.ud[+t.dataset.u],q=u&&((u.pg||{}).rows||[])[+t.dataset.j];if(!q)return;var val=t.value,x=null;
    if(val==='otro'){q.n='';delete q.rn;delete q.rp;delete q.rk}
    else if(/^g\d+$/.test(val))x=catGen()[+val.slice(1)]; else if(/^m\d+$/.test(val))x=catMios()[+val.slice(1)];
    if(x){q.n=x.n;q.rn=x.rn;q.rp=x.rp;q.rk=x.rk;if(x.org)q.org=1;else delete q.org}
    p3.m=Date.now();save();pabRedraw();return}
});

/* ───── arranque ───── */
cargar();
document.addEventListener('DOMContentLoaded',function(){render()});
(function(){var tm=null;try{new MutationObserver(function(){if(tm)return;tm=setTimeout(function(){tm=null;stkLab();inThou()},0)}).observe(document.documentElement,{childList:true,subtree:true})}catch(e){}})();
if('serviceWorker' in navigator&&(location.protocol==='https:'||location.hostname==='localhost'))window.addEventListener('load',function(){navigator.serviceWorker.register('sw.js').catch(function(){})});
window.__pa={S:function(){return S},UI:function(){return UI},render:render,ejemplo:ejemplo,pabCalc:pabCalc,pabBlocks:pabBlocks,buildPabXlsx:buildPabXlsx,ensurePdf:ensurePdf,buildPabPdf:buildPabPdf,climaDe:climaDe,refParse:refParse,catGen:catGen};
