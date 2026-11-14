/* PPTX export for the WITJ deck.
   Walks the rendered DOM of every slide and writes native PowerPoint elements:
   boxes -> shapes, text -> text boxes (with runs), <img> -> pictures, inline SVG -> pictures (+ native text).
   The .pptx package is assembled by hand (OOXML) with JSZip, so each paragraph gets exactly one pPr. */
(function(){
'use strict';
var JSZIP_URL='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
var W=1280,H=720,PX=9525; /* EMU per CSS px: 12192000 EMU / 1280 px */
var FONT='Noto Sans JP';
var NS_A='http://schemas.openxmlformats.org/drawingml/2006/main',NS_R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',NS_P='http://schemas.openxmlformats.org/presentationml/2006/main';
var REPLACED={IMG:1,SVG:1,CANVAS:1,VIDEO:1,PICTURE:1,IFRAME:1,OBJECT:1};
var cs=function(el,ps){return getComputedStyle(el,ps||null)};
function emu(v){return Math.round(v*PX)}
function pt100(px){return Math.round(px*75)}
function esc(s){return String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function hex2(n){n=Math.max(0,Math.min(255,Math.round(n)));return (n<16?'0':'')+n.toString(16).toUpperCase()}
function color(str){
  if(!str||str==='transparent'||str==='none')return null;
  var m=/rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:[,\s\/]+([\d.]+%?))?\s*\)/.exec(str);
  if(!m){var h=/^#([0-9a-f]{6})$/i.exec(str.trim());return h?{hex:h[1].toUpperCase(),a:1}:null}
  var a=m[4]===undefined?1:(/%$/.test(m[4])?parseFloat(m[4])/100:parseFloat(m[4]));
  if(a<=0.002)return null;
  return {hex:hex2(+m[1])+hex2(+m[2])+hex2(+m[3]),a:a};
}
function fillXml(c,op){
  if(!c)return '<a:noFill/>';
  var a=Math.round(c.a*(op===undefined?1:op)*100000);
  return '<a:solidFill><a:srgbClr val="'+c.hex+'"'+(a<100000?'><a:alpha val="'+Math.max(0,a)+'"/></a:srgbClr>':'/>')+'</a:solidFill>';
}
function lnXml(ln,op){
  if(!ln)return '<a:ln><a:noFill/></a:ln>';
  return '<a:ln w="'+emu(ln.w)+'" cap="flat">'+fillXml(ln.c,op)+(ln.dash?'<a:prstDash val="'+ln.dash+'"/>':'')+'</a:ln>';
}
function shdwXml(sh,op){
  if(!sh)return '';
  return '<a:effectLst><a:outerShdw blurRad="'+emu(sh.blur)+'" dist="'+emu(sh.dist)+'" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="'+sh.c.hex+'"><a:alpha val="'+Math.round(sh.c.a*op*100000)+'"/></a:srgbClr></a:outerShdw></a:effectLst>';
}
function xfrm(o){return '<a:xfrm'+(o.rot?' rot="'+Math.round(o.rot*60000)+'"':'')+'><a:off x="'+emu(o.x)+'" y="'+emu(o.y)+'"/><a:ext cx="'+Math.max(0,emu(o.w))+'" cy="'+Math.max(0,emu(o.h))+'"/></a:xfrm>'}
function geomXml(o){
  if(o.prst==='roundRect')return '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val '+Math.round(Math.max(0,Math.min(50000,o.adj||0)))+'"/></a:avLst></a:prstGeom>';
  return '<a:prstGeom prst="'+(o.prst||'rect')+'"><a:avLst/></a:prstGeom>';
}
function spXml(id,o){
  return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="'+esc(o.name||('Shape '+id))+'"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'+xfrm(o)+geomXml(o)+fillXml(o.fill,o.op)+lnXml(o.ln,o.op)+shdwXml(o.shadow,o.op)+'</p:spPr>'+
    '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="ja-JP"/></a:p></p:txBody></p:sp>';
}
function rprXml(s){
  return '<a:rPr lang="ja-JP" altLang="en-US" sz="'+pt100(s.size)+'" kern="1200"'+(s.bold?' b="1"':'')+(s.italic?' i="1"':'')+(s.u?' u="sng"':'')+(s.spc?' spc="'+Math.round(s.spc*75)+'"':'')+' dirty="0">'+fillXml(s.c,s.op)+'<a:latin typeface="'+FONT+'"/><a:ea typeface="'+FONT+'"/><a:cs typeface="'+FONT+'"/></a:rPr>';
}
function txXml(id,o){
  var sz=pt100(o.size),body='<a:p><a:pPr algn="'+o.align+'"><a:lnSpc><a:spcPts val="'+Math.round(o.lineSp*75)+'"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>';
  o.lines.forEach(function(line,i){
    if(i>0)body+='<a:br>'+rprXml(line[0].st)+'</a:br>';
    line.forEach(function(r){body+='<a:r>'+rprXml(r.st)+'<a:t>'+esc(r.t)+'</a:t></a:r>'});
  });
  body+='<a:endParaRPr lang="ja-JP" sz="'+sz+'" dirty="0"/></a:p>';
  return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="'+esc(o.name||('TextBox '+id))+'"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'+xfrm(o)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'+
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>'+body+'</p:txBody></p:sp>';
}
function pc(f){return Math.round(Math.max(0,Math.min(1,f))*100000)}
function picXml(id,o,rId){
  return '<p:pic><p:nvPicPr><p:cNvPr id="'+id+'" name="'+esc(o.name||('Picture '+id))+'" descr="'+esc(o.descr||'')+'"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'+
    '<p:blipFill><a:blip r:embed="'+rId+'">'+(o.op<0.995?'<a:alphaModFix amt="'+Math.round(o.op*100000)+'"/>':'')+'</a:blip>'+
    (o.crop?'<a:srcRect l="'+pc(o.crop.l)+'" t="'+pc(o.crop.t)+'" r="'+pc(o.crop.r)+'" b="'+pc(o.crop.b)+'"/>':'')+'<a:stretch><a:fillRect/></a:stretch></p:blipFill>'+
    '<p:spPr>'+xfrm(o)+geomXml(o)+shdwXml(o.shadow,1)+'</p:spPr></p:pic>';
}

/* ---------- geometry helpers ---------- */
function radius(v,w,h){
  var m=/^([\d.]+)(px|%)/.exec(v||'');if(!m)return 0;
  var n=parseFloat(m[1]);if(m[2]==='%')n=n/100*Math.min(w,h);
  return Math.max(0,Math.min(n,Math.min(w,h)/2));
}
function radii(st,w,h){return [radius(st.borderTopLeftRadius,w,h),radius(st.borderTopRightRadius,w,h),radius(st.borderBottomRightRadius,w,h),radius(st.borderBottomLeftRadius,w,h)]}
function shapeFor(rad,w,h){ /* -> {prst, adj, extra:[rects covering square sides]} */
  var mn=Math.min(w,h),r=Math.max.apply(null,rad);
  if(r<=0.5)return {prst:'rect'};
  if(rad.every(function(v){return v>=mn/2-0.75}))return {prst:'ellipse'};
  if(rad.every(function(v){return Math.abs(v-r)<0.5}))return {prst:'roundRect',adj:r/mn*100000};
  var tl=rad[0],tr=rad[1],br=rad[2],bl=rad[3],z=function(v){return v<0.5},s=function(a,b){return Math.abs(a-b)<0.5};
  var extra=null;
  if(s(tr,br)&&z(tl)&&z(bl))extra={x:0,y:0,w:r,h:h};
  else if(s(tl,bl)&&z(tr)&&z(br))extra={x:w-r,y:0,w:r,h:h};
  else if(s(tl,tr)&&z(br)&&z(bl))extra={x:0,y:h-r,w:w,h:r};
  else if(s(br,bl)&&z(tl)&&z(tr))extra={x:0,y:0,w:w,h:r};
  return {prst:'roundRect',adj:r/mn*100000,extra:extra};
}
function parseShadows(str){
  var out=[];if(!str||str==='none')return out;
  var re=/(rgba?\([^)]*\)|#[0-9a-fA-F]{3,8})\s+(-?[\d.]+)px\s+(-?[\d.]+)px(?:\s+(-?[\d.]+)px)?(?:\s+(-?[\d.]+)px)?(\s+inset)?/g,m;
  while((m=re.exec(str))){out.push({c:color(m[1]),x:parseFloat(m[2]),y:parseFloat(m[3]),blur:parseFloat(m[4]||0),spread:parseFloat(m[5]||0),inset:!!m[6]})}
  return out.filter(function(s){return s.c&&!s.inset});
}
function softShadow(list){
  var best=null;list.forEach(function(s){if(s.blur>0&&(!best||s.blur*s.c.a>best.blur*best.c.a))best=s});
  if(!best)return null;
  return {blur:Math.max(2,best.blur*0.55),dist:Math.max(0,best.y+best.spread*0.5),c:{hex:best.c.hex,a:Math.min(1,best.c.a*1.2)}};
}
function roundPath(ctx,x,y,w,h,r){
  var tl=r[0],tr=r[1],br=r[2],bl=r[3];
  ctx.beginPath();ctx.moveTo(x+tl,y);ctx.lineTo(x+w-tr,y);ctx.arcTo(x+w,y,x+w,y+tr,tr);ctx.lineTo(x+w,y+h-br);ctx.arcTo(x+w,y+h,x+w-br,y+h,br);
  ctx.lineTo(x+bl,y+h);ctx.arcTo(x,y+h,x,y+h-bl,bl);ctx.lineTo(x,y+tl);ctx.arcTo(x,y,x+tl,y,tl);ctx.closePath();
}
function splitTop(s){var out=[],d=0,cur='';for(var i=0;i<s.length;i++){var ch=s[i];if(ch==='(')d++;if(ch===')')d--;if(ch===','&&d===0){out.push(cur);cur=''}else cur+=ch}if(cur.trim())out.push(cur);return out}
function cssGradient(ctx,s,w,h){
  var m=/linear-gradient\(([\s\S]*)\)\s*$/.exec(s.trim());if(!m)return null;
  var parts=splitTop(m[1]),angle=180,first=parts[0].trim();
  if(/deg\s*$/.test(first)){angle=parseFloat(first);parts.shift()}
  else if(/rad\s*$/.test(first)){angle=parseFloat(first)*180/Math.PI;parts.shift()}
  else if(/turn\s*$/.test(first)){angle=parseFloat(first)*360;parts.shift()}
  else if(/^to\s/.test(first)){var d=first.slice(3).trim().split(/\s+/).sort().join(' ');angle={'top':0,'right':90,'bottom':180,'left':270,'right top':45,'bottom right':135,'bottom left':225,'left top':315}[d];if(angle===undefined)angle=180;parts.shift()}
  var a=angle*Math.PI/180,dx=Math.sin(a),dy=-Math.cos(a),L=Math.abs(w*dx)+Math.abs(h*dy),cx=w/2,cy=h/2;
  var g=ctx.createLinearGradient(cx-dx*L/2,cy-dy*L/2,cx+dx*L/2,cy+dy*L/2);
  parts.forEach(function(p,i){p=p.trim();var mm=/^(rgba?\([^)]*\)|#[0-9a-fA-F]+|[a-zA-Z]+)(?:\s+(-?[\d.]+)%)?/.exec(p);if(!mm)return;var pos=mm[2]!==undefined?parseFloat(mm[2])/100:(parts.length>1?i/(parts.length-1):0);g.addColorStop(Math.max(0,Math.min(1,pos)),mm[1])});
  return g;
}
function rasterGradient(w,h,rad,grad){
  var k=2,c=document.createElement('canvas');c.width=Math.max(1,Math.ceil(w*k));c.height=Math.max(1,Math.ceil(h*k));
  var ctx=c.getContext('2d');ctx.scale(k,k);roundPath(ctx,0,0,w,h,rad);ctx.clip();
  var g=cssGradient(ctx,grad,w,h);if(!g)return null;ctx.fillStyle=g;ctx.fillRect(0,0,w,h);
  return c.toDataURL('image/png');
}

/* ---------- exporter ---------- */
function Exporter(){this.media=[];this.mediaMap=new Map();this.slides=[];this.pending=[]}
Exporter.prototype.addMedia=function(dataUrl){
  var k=this.mediaMap.get(dataUrl);if(k!==undefined)return k;
  var m=/^data:image\/(png|jpeg|jpg|gif)(?:;[^,]*)?;base64,([\s\S]*)$/.exec(dataUrl);if(!m)return -1;
  this.media.push({ext:m[1]==='jpg'?'jpeg':m[1],b64:m[2]});k=this.media.length-1;this.mediaMap.set(dataUrl,k);return k;
};

function SlideCtx(ex,root,base){this.ex=ex;this.root=root;this.base=base;this.ops=[];this.deferred=[]}
SlideCtx.prototype.rect=function(el){var r=el.getBoundingClientRect();return {x:r.left-this.base.left,y:r.top-this.base.top,w:r.width,h:r.height}};
SlideCtx.prototype.run=function(){
  var self=this;
  this.walk(this.root,1,true);
  this.deferred.sort(function(a,b){return a.z-b.z}).forEach(function(d){self.walk(d.el,d.op,true)});
};
SlideCtx.prototype.walk=function(el,op,noDefer){
  if(!el||el.nodeType!==1)return;
  var st=cs(el),tag=el.tagName.toUpperCase();
  if(st.display==='none'||st.visibility==='hidden')return;
  op=op*(parseFloat(st.opacity)||0);if(op<=0.005)return;
  if(!noDefer&&st.position!=='static'&&st.zIndex!=='auto'&&+st.zIndex>0){this.deferred.push({el:el,z:+st.zIndex,op:op/(parseFloat(st.opacity)||1)});return}
  if(tag==='BR'||tag==='SCRIPT'||tag==='STYLE'||tag==='TEMPLATE')return;
  var r=this.rect(el);
  if(tag==='IMG'){this.img(el,st,r,op);return}
  if(tag==='SVG'){this.svg(el,st,r,op);return}
  if(el!==this.root)this.box(el,st,r,op);
  this.pseudo(el,'::before',st,r,op);
  if(this.isTextContainer(el))this.text(el,st,r,op);
  else for(var c=el.firstElementChild;c;c=c.nextElementSibling)this.walk(c,op);
  this.pseudo(el,'::after',st,r,op);
};
SlideCtx.prototype.isTextContainer=function(el){
  for(var c=el.firstChild;c;c=c.nextSibling){
    if(c.nodeType===3){if(/\S/.test(c.nodeValue))return true}
    else if(c.nodeType===1&&!REPLACED[c.tagName.toUpperCase()]&&cs(c).display==='inline'&&/\S/.test(c.textContent))return true;
  }
  return false;
};
function edges(st){return {bt:parseFloat(st.borderTopWidth)||0,br:parseFloat(st.borderRightWidth)||0,bb:parseFloat(st.borderBottomWidth)||0,bl:parseFloat(st.borderLeftWidth)||0,pt:parseFloat(st.paddingTop)||0,pr:parseFloat(st.paddingRight)||0,pb:parseFloat(st.paddingBottom)||0,pl:parseFloat(st.paddingLeft)||0}}

/* backgrounds, borders, shadows of a block */
SlideCtx.prototype.box=function(el,st,r,op){
  if(r.w<=0||r.h<=0)return;
  var bg=color(st.backgroundColor),grad=/linear-gradient/.test(st.backgroundImage)?st.backgroundImage:null;
  var rad=radii(st,r.w,r.h),shadows=parseShadows(st.boxShadow);
  var sides=['Top','Right','Bottom','Left'].map(function(s){var w=parseFloat(st['border'+s+'Width'])||0,c=color(st['border'+s+'Color']),sty=st['border'+s+'Style'];return (w>0&&c&&sty!=='none'&&sty!=='hidden')?{w:w,c:c,dash:sty==='dashed'?'dash':sty==='dotted'?'sysDot':null,key:w+'|'+c.hex+'|'+c.a+'|'+sty}:null});
  var ln=null,bars=[];
  var counts={};sides.forEach(function(s){if(s)counts[s.key]=(counts[s.key]||0)+1});
  var bestKey=null;Object.keys(counts).forEach(function(k){if(counts[k]>=3&&(!bestKey||counts[k]>counts[bestKey]))bestKey=k});
  if(bestKey){ln=sides.filter(function(s){return s&&s.key===bestKey})[0]}
  sides.forEach(function(s,i){if(s&&(!ln||s.key!==ln.key||counts[s.key]<4))bars.push({side:i,s:s})});
  if(!bg&&!grad&&!ln&&!shadows.length&&!bars.length)return;
  var g=shapeFor(rad,r.w,r.h),self=this;
  /* halo-style shadows (no blur, spread only) */
  shadows.forEach(function(sh){
    if(sh.blur===0&&sh.spread>0){var hr={x:r.x-sh.spread+sh.x,y:r.y-sh.spread+sh.y,w:r.w+2*sh.spread,h:r.h+2*sh.spread};
      var hg=shapeFor(rad.map(function(v){return v>0?v+sh.spread:0}),hr.w,hr.h);
      self.ops.push({k:'sp',x:hr.x,y:hr.y,w:hr.w,h:hr.h,prst:hg.prst,adj:hg.adj,fill:sh.c,ln:null,op:op,name:'Halo'})}
  });
  var soft=softShadow(shadows);
  if(grad){
    var png=rasterGradient(r.w,r.h,rad,grad);
    if(png){var key=this.ex.addMedia(png);if(key>=0)this.ops.push({k:'pic',x:r.x,y:r.y,w:r.w,h:r.h,media:key,prst:'rect',shadow:soft,op:op,name:'Gradient fill'})}
  }else if(bg||ln||soft){
    var sr={x:r.x,y:r.y,w:r.w,h:r.h},adj=g.adj;
    if(ln){sr={x:r.x+ln.w/2,y:r.y+ln.w/2,w:r.w-ln.w,h:r.h-ln.w};if(g.prst==='roundRect')adj=Math.max(0,(Math.max.apply(null,rad)-ln.w/2))/Math.min(sr.w,sr.h)*100000}
    this.ops.push({k:'sp',x:sr.x,y:sr.y,w:sr.w,h:sr.h,prst:g.prst,adj:adj,fill:bg,ln:ln?{w:ln.w,c:ln.c,dash:ln.dash}:null,shadow:soft,op:op,name:el.className?String(el.className).split(' ')[0]:'Box'});
    if(g.extra&&bg)this.ops.push({k:'sp',x:r.x+g.extra.x,y:r.y+g.extra.y,w:g.extra.w,h:g.extra.h,prst:'rect',fill:bg,ln:null,op:op,name:'Box corner'});
  }
  bars.forEach(function(b){
    var s=b.s,w=s.w,i=b.side,br;
    var ins=function(a,c){return Math.max(rad[a],rad[c])};
    if(i===0)br={x:r.x+ins(0,0),y:r.y,w:r.w-ins(0,0)-ins(1,1),h:w};
    else if(i===2)br={x:r.x+ins(3,3),y:r.y+r.h-w,w:r.w-ins(3,3)-ins(2,2),h:w};
    else if(i===3)br={x:r.x,y:r.y+ins(0,0),w:w,h:r.h-ins(0,0)-ins(3,3)};
    else br={x:r.x+r.w-w,y:r.y+ins(1,1),w:w,h:r.h-ins(1,1)-ins(2,2)};
    if(br.w>0&&br.h>0)self.ops.push({k:'sp',x:br.x,y:br.y,w:br.w,h:br.h,prst:'rect',fill:s.c,ln:null,op:op,name:'Border'});
  });
};

/* absolutely positioned decorative pseudo-elements (slide accent bar, timeline line, bullets, ...) */
SlideCtx.prototype.pseudo=function(el,which,st,r,op){
  var ps;try{ps=cs(el,which)}catch(e){return}
  if(!ps||ps.content==='none'||ps.content==='normal'||ps.display==='none'||ps.position!=='absolute')return;
  var e=edges(st),cb={x:r.x+e.bl,y:r.y+e.bt,w:r.w-e.bl-e.br,h:r.h-e.bt-e.bb};
  var num=function(v){return v==='auto'?null:parseFloat(v)};
  var L=num(ps.left),R=num(ps.right),T=num(ps.top),B=num(ps.bottom),Wd=num(ps.width),Ht=num(ps.height);
  var w=Wd!==null?Wd:(L!==null&&R!==null?cb.w-L-R:0),h=Ht!==null?Ht:(T!==null&&B!==null?cb.h-T-B:0);
  if(w<=0||h<=0)return;
  var x=L!==null?cb.x+L:(R!==null?cb.x+cb.w-R-w:cb.x),y=T!==null?cb.y+T:(B!==null?cb.y+cb.h-B-h:cb.y);
  var bg=color(ps.backgroundColor),grad=/linear-gradient/.test(ps.backgroundImage)?ps.backgroundImage:null;
  var rot=0,m=/matrix\(([^)]+)\)/.exec(ps.transform||'');
  if(m){var v=m[1].split(',').map(parseFloat);rot=Math.atan2(v[1],v[0])*180/Math.PI;x+=v[4]||0;y+=v[5]||0}
  var o=parseFloat(ps.opacity);if(!isNaN(o))op*=o;
  if(grad){
    var png=rasterGradient(w,h,radii(ps,w,h),grad);
    if(png){var key=this.ex.addMedia(png);if(key>=0)this.ops.push({k:'pic',x:x,y:y,w:w,h:h,media:key,prst:'rect',op:op,name:'Decoration'})}
    return;
  }
  if(!bg)return;
  if(Math.abs(Math.abs(rot)-45)<1.5&&Math.abs(w-h)<0.6){var d=w*Math.SQRT2,cx=x+w/2,cy=y+h/2;this.ops.push({k:'sp',x:cx-d/2,y:cy-d/2,w:d,h:d,prst:'diamond',fill:bg,ln:null,op:op,name:'Bullet'});return}
  var g=shapeFor(radii(ps,w,h),w,h);
  this.ops.push({k:'sp',x:x,y:y,w:w,h:h,prst:g.prst,adj:g.adj,fill:bg,ln:null,rot:rot,op:op,name:'Decoration'});
};

/* text container -> one text box with runs and explicit line breaks at the browser's wrap points */
SlideCtx.prototype.text=function(el,st,r,op){
  var self=this,range=document.createRange(),lines=[],cur=null;
  var styleCache=new Map();
  function styleOf(p){
    var s=styleCache.get(p);if(s)return s;
    var c=cs(p),o=1,q=p;while(q&&q!==el){o*=(parseFloat(cs(q).opacity)||1);q=q.parentNode}
    s={size:parseFloat(c.fontSize)||16,bold:(parseInt(c.fontWeight,10)||400)>=600,italic:c.fontStyle==='italic',u:/underline/.test(c.textDecorationLine||c.textDecoration||''),spc:parseFloat(c.letterSpacing)||0,c:color(c.color)||{hex:'000000',a:1},op:op*o};
    styleCache.set(p,s);return s;
  }
  function trimEnd(line){while(line.runs.length){var last=line.runs[line.runs.length-1];last.t=last.t.replace(/\s+$/,'');if(last.t)break;line.runs.pop()}}
  function addChar(ch,b,p){
    var same=false;
    if(cur){var ov=Math.min(cur.bottom,b.y+b.h)-Math.max(cur.top,b.y);same=ov>0.5*Math.min(b.h,cur.bottom-cur.top)}
    if(!same){if(cur)trimEnd(cur);if(/\s/.test(ch))return;cur={runs:[],top:b.y,bottom:b.y+b.h,left:b.x,right:b.x+b.w};lines.push(cur)}
    else if(!cur.runs.length&&/\s/.test(ch))return;
    cur.top=Math.min(cur.top,b.y);cur.bottom=Math.max(cur.bottom,b.y+b.h);cur.left=Math.min(cur.left,b.x);cur.right=Math.max(cur.right,b.x+b.w);
    var last=cur.runs[cur.runs.length-1];
    if(last&&last.p===p)last.t+=ch;else cur.runs.push({t:ch,p:p,st:styleOf(p)});
  }
  function gather(node){
    for(var c=node.firstChild;c;c=c.nextSibling){
      if(c.nodeType===3){
        var txt=c.nodeValue;
        for(var k=0;k<txt.length;k++){
          var len=1,code=txt.charCodeAt(k);if(code>=0xD800&&code<=0xDBFF&&k+1<txt.length)len=2;
          var ch=txt.substr(k,len);range.setStart(c,k);range.setEnd(c,k+len);var b=range.getBoundingClientRect();k+=len-1;
          if(b.width<=0.01)continue;
          if(/[\n\r\t]/.test(ch))ch=' ';
          addChar(ch,{x:b.left-self.base.left,y:b.top-self.base.top,w:b.width,h:b.height},c.parentNode);
        }
      }else if(c.nodeType===1){
        var t=c.tagName.toUpperCase();if(t==='BR')continue;
        if(REPLACED[t]||cs(c).display!=='inline')self.walk(c,op);else gather(c);
      }
    }
  }
  gather(el);if(cur)trimEnd(cur);
  lines=lines.filter(function(l){return l.runs.length});
  if(!lines.length)return;
  var n=lines.length,e=edges(st);
  var lineH=n>1?(lines[n-1].top-lines[0].top)/(n-1):(parseFloat(st.lineHeight)||(lines[0].bottom-lines[0].top));
  if(!(lineH>0))lineH=lines[0].bottom-lines[0].top;
  var gl=Infinity,gr=-Infinity,size=0;lines.forEach(function(l){gl=Math.min(gl,l.left);gr=Math.max(gr,l.right);l.runs.forEach(function(x){size=Math.max(size,x.st.size)})});
  var gt=lines[0].top,gb=lines[n-1].bottom,cy=(gt+gb)/2,h=n*lineH+2;
  var cl=r.x+e.bl+e.pl,cw=Math.max(0,r.w-e.bl-e.br-e.pl-e.pr);
  var ta=st.textAlign,align=(ta==='center')?'ctr':(ta==='right'||ta==='end')?'r':'l';
  var slack=n===1?Math.max(24,0.3*(gr-gl)):Math.max(8,0.1*Math.max(cw,gr-gl)),x,w;
  if(align==='l'){x=gl;w=Math.max(cl+cw,gr)-gl+slack}
  else if(align==='ctr'){w=Math.max(cw,gr-gl)+slack;x=(gl+gr)/2-w/2}
  else{w=Math.max(cw,gr-gl)+slack;x=gr-w}
  if(x<0){w+=x;x=0}if(x+w>W)w=W-x;
  var strokeW=parseFloat(st.webkitTextStrokeWidth)||0;
  if(strokeW>0){this.rasterText(el,st,lines,{gl:gl,gr:gr,gt:gt,gb:gb},strokeW,op);return}
  this.ops.push({k:'tx',x:x,y:cy-h/2,w:w,h:h,align:align,lineSp:lineH,size:size,op:op,name:'Text: '+lines[0].runs.map(function(x){return x.t}).join('').slice(0,24),
    lines:lines.map(function(l){return l.runs.map(function(x){return {t:x.t,st:x.st}})})});
};
/* outlined text (-webkit-text-stroke) is painted with canvas, using the page's loaded web font */
SlideCtx.prototype.rasterText=function(el,st,lines,g,sw,op){
  var k=3,pad=sw+2,w=g.gr-g.gl+pad*2,h=g.gb-g.gt+pad*2,c=document.createElement('canvas');
  c.width=Math.ceil(w*k);c.height=Math.ceil(h*k);var ctx=c.getContext('2d');ctx.scale(k,k);
  var fill=color(st.color),stroke=color(st.webkitTextStrokeColor);
  lines.forEach(function(l){
    var x=l.left-g.gl+pad;
    l.runs.forEach(function(run){
      var s=run.st,p=run.p,pc=cs(p);
      ctx.font=pc.fontStyle+' '+pc.fontWeight+' '+s.size+'px '+pc.fontFamily;ctx.textBaseline='alphabetic';
      var m=ctx.measureText(run.t),asc=m.fontBoundingBoxAscent||s.size*0.88;
      var y=l.top-g.gt+pad+asc;
      if(fill){ctx.fillStyle='rgba('+parseInt(fill.hex.slice(0,2),16)+','+parseInt(fill.hex.slice(2,4),16)+','+parseInt(fill.hex.slice(4,6),16)+','+fill.a+')';ctx.fillText(run.t,x,y)}
      if(stroke){ctx.lineWidth=sw;ctx.lineJoin='round';ctx.strokeStyle='rgba('+parseInt(stroke.hex.slice(0,2),16)+','+parseInt(stroke.hex.slice(2,4),16)+','+parseInt(stroke.hex.slice(4,6),16)+','+stroke.a+')';ctx.strokeText(run.t,x,y)}
      x+=m.width+s.spc*run.t.length;
    });
  });
  var key=this.ex.addMedia(c.toDataURL('image/png'));
  if(key>=0)this.ops.push({k:'pic',x:g.gl-pad,y:g.gt-pad,w:w,h:h,media:key,prst:'rect',op:op,name:'Outlined text'});
};

SlideCtx.prototype.img=function(el,st,r,op){
  if(r.w<=0||r.h<=0)return;
  var src=el.currentSrc||el.src||'';
  if(!/^data:/.test(src)){
    try{var cv=document.createElement('canvas');cv.width=el.naturalWidth;cv.height=el.naturalHeight;cv.getContext('2d').drawImage(el,0,0);src=cv.toDataURL('image/png')}catch(e){console.warn('pptx-export: skipped cross-origin image',src);return}
  }
  var key=this.ex.addMedia(src);if(key<0)return;
  var nw=el.naturalWidth,nh=el.naturalHeight,fit=st.objectFit,crop=null,d={x:r.x,y:r.y,w:r.w,h:r.h};
  if(nw&&nh){
    if(fit==='cover'){var s=Math.max(r.w/nw,r.h/nh),sw=r.w/s,sh=r.h/s,l=(nw-sw)/2/nw,t=(nh-sh)/2/nh;if(l>0.001||t>0.001)crop={l:l,t:t,r:l,b:t}}
    else if(fit==='contain'||fit==='scale-down'){var s2=Math.min(r.w/nw,r.h/nh);if(fit==='scale-down')s2=Math.min(s2,1);var dw=nw*s2,dh=nh*s2;d={x:r.x+(r.w-dw)/2,y:r.y+(r.h-dh)/2,w:dw,h:dh}}
  }
  var g=shapeFor(radii(st,d.w,d.h),d.w,d.h);
  this.ops.push({k:'pic',x:d.x,y:d.y,w:d.w,h:d.h,media:key,crop:crop,prst:g.prst,adj:g.adj,shadow:softShadow(parseShadows(st.boxShadow)),op:op,descr:el.alt||'',name:el.alt?('Image: '+el.alt):'Image'});
};

/* inline SVG: <text> becomes native text boxes, the rest is rasterised to a transparent PNG */
SlideCtx.prototype.svg=function(el,st,r,op){
  if(r.w<=0||r.h<=0)return;
  var self=this,textOps=[];
  el.querySelectorAll('text').forEach(function(t){
    var ts=cs(t);if(ts.display==='none'||ts.visibility==='hidden')return;
    var txt=t.textContent.replace(/\s+/g,' ').trim();if(!txt)return;
    var tr=self.rect(t),ctm=t.getScreenCTM(),k=ctm?Math.sqrt(ctm.a*ctm.a+ctm.b*ctm.b):1;
    var size=(parseFloat(ts.fontSize)||16)*k,anchor=ts.textAnchor,align=anchor==='middle'?'ctr':anchor==='end'?'r':'l';
    var lineH=size*1.25,slack=Math.max(8,tr.w*0.15),w=tr.w+slack,x=align==='ctr'?tr.x+tr.w/2-w/2:align==='r'?tr.x+tr.w-w:tr.x;
    var stl={size:size,bold:(parseInt(ts.fontWeight,10)||400)>=600,italic:ts.fontStyle==='italic',u:false,spc:(parseFloat(ts.letterSpacing)||0)*k,c:color(ts.fill)||{hex:'000000',a:1},op:op*(parseFloat(ts.opacity)||1)*(parseFloat(ts.fillOpacity)||1)};
    textOps.push({k:'tx',x:x,y:tr.y+tr.h/2-lineH/2-1,w:w,h:lineH+2,align:align,lineSp:lineH,size:size,op:op,name:'Text: '+txt.slice(0,24),lines:[[{t:txt,st:stl}]]});
  });
  var clone=el.cloneNode(true);
  clone.querySelectorAll('text').forEach(function(t){t.remove()});
  if(!clone.getAttribute('xmlns'))clone.setAttribute('xmlns','http://www.w3.org/2000/svg');
  if(!clone.getAttribute('viewBox'))clone.setAttribute('viewBox','0 0 '+r.w+' '+r.h);
  var k=Math.min(8,Math.max(2,Math.ceil(600/Math.max(r.w,r.h))));
  clone.setAttribute('width',String(Math.round(r.w*k)));clone.setAttribute('height',String(Math.round(r.h*k)));
  clone.style.width='';clone.style.height='';clone.style.display='';clone.removeAttribute('class');
  clone.style.color=st.color;clone.style.stroke=st.stroke;clone.style.strokeWidth=st.strokeWidth;clone.style.fill=st.fill;
  clone.style.strokeLinecap=st.strokeLinecap;clone.style.strokeLinejoin=st.strokeLinejoin;clone.style.fillOpacity=st.fillOpacity;clone.style.strokeOpacity=st.strokeOpacity;
  var xml=new XMLSerializer().serializeToString(clone);
  var opn={k:'pic',x:r.x,y:r.y,w:r.w,h:r.h,media:-1,prst:'rect',op:op,name:'Graphic'};
  this.ops.push(opn);textOps.forEach(function(t){self.ops.push(t)});
  this.ex.pending.push(new Promise(function(resolve){
    var im=new Image();
    im.onload=function(){try{var c=document.createElement('canvas');c.width=Math.round(r.w*k);c.height=Math.round(r.h*k);c.getContext('2d').drawImage(im,0,0,c.width,c.height);opn.media=self.ex.addMedia(c.toDataURL('image/png'))}catch(e){console.warn('pptx-export: svg raster failed',e)}resolve()};
    im.onerror=function(){console.warn('pptx-export: svg load failed');resolve()};
    im.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(xml);
  }));
};

/* ---------- package ---------- */
var CT_PRES='application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
function slideXml(sl,rels){
  var id=2,xml='';
  sl.ops.forEach(function(o){
    if(o.k==='sp')xml+=spXml(id++,o);
    else if(o.k==='tx')xml+=txXml(id++,o);
    else if(o.k==='pic'&&o.media>=0){var rId=rels.use(o.media);xml+=picXml(id++,o,rId)}
  });
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'"><p:cSld><p:bg><p:bgPr>'+fillXml(sl.bg||{hex:'FFFFFF',a:1},1)+'<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'+xml+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}
Exporter.prototype.buildZip=function(title){
  var zip=new JSZip(),n=this.slides.length,self=this,XH='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var ct=XH+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>'+
    '<Override PartName="/ppt/presentation.xml" ContentType="'+CT_PRES+'"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
  for(var i=1;i<=n;i++)ct+='<Override PartName="/ppt/slides/slide'+i+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
  ct+='</Types>';
  zip.file('[Content_Types].xml',ct);
  zip.file('_rels/.rels',XH+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
  var now=new Date().toISOString().replace(/\.\d+Z$/,'Z');
  zip.file('docProps/core.xml',XH+'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>'+esc(title)+'</dc:title><dc:creator>WITJ</dc:creator><cp:lastModifiedBy>WITJ</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'+now+'</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">'+now+'</dcterms:modified></cp:coreProperties>');
  zip.file('docProps/app.xml',XH+'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><Slides>'+n+'</Slides><PresentationFormat>Widescreen</PresentationFormat><AppVersion>16.0000</AppVersion></Properties>');
  var sldIds='',pRels='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
  for(i=1;i<=n;i++){sldIds+='<p:sldId id="'+(255+i)+'" r:id="rId'+(i+1)+'"/>';pRels+='<Relationship Id="rId'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+i+'.xml"/>'}
  var k=n+2;
  pRels+='<Relationship Id="rId'+k+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId'+(k+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/><Relationship Id="rId'+(k+2)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId'+(k+3)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>';
  zip.file('ppt/presentation.xml',XH+'<p:presentation xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>'+sldIds+'</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels',XH+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+pRels+'</Relationships>');
  zip.file('ppt/presProps.xml',XH+'<p:presentationPr xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'"/>');
  zip.file('ppt/viewProps.xml',XH+'<p:viewPr xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr></p:viewPr>');
  zip.file('ppt/tableStyles.xml',XH+'<a:tblStyleLst xmlns:a="'+NS_A+'" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
  var ph='<a:schemeClr val="phClr"/>';
  zip.file('ppt/theme/theme1.xml',XH+'<a:theme xmlns:a="'+NS_A+'" name="WITJ"><a:themeElements><a:clrScheme name="WITJ"><a:dk1><a:srgbClr val="0D2340"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="084B8F"/></a:dk2><a:lt2><a:srgbClr val="EEF4FB"/></a:lt2><a:accent1><a:srgbClr val="0A66C2"/></a:accent1><a:accent2><a:srgbClr val="B4232F"/></a:accent2><a:accent3><a:srgbClr val="56677D"/></a:accent3><a:accent4><a:srgbClr val="CFE0F3"/></a:accent4><a:accent5><a:srgbClr val="1F8A5B"/></a:accent5><a:accent6><a:srgbClr val="8A98AA"/></a:accent6><a:hlink><a:srgbClr val="0A66C2"/></a:hlink><a:folHlink><a:srgbClr val="56677D"/></a:folHlink></a:clrScheme><a:fontScheme name="WITJ"><a:majorFont><a:latin typeface="'+FONT+'"/><a:ea typeface="'+FONT+'"/><a:cs typeface="'+FONT+'"/></a:majorFont><a:minorFont><a:latin typeface="'+FONT+'"/><a:ea typeface="'+FONT+'"/><a:cs typeface="'+FONT+'"/></a:minorFont></a:fontScheme><a:fmtScheme name="WITJ"><a:fillStyleLst><a:solidFill>'+ph+'</a:solidFill><a:solidFill>'+ph+'</a:solidFill><a:solidFill>'+ph+'</a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill>'+ph+'</a:solidFill></a:ln><a:ln w="19050"><a:solidFill>'+ph+'</a:solidFill></a:ln><a:ln w="28575"><a:solidFill>'+ph+'</a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill>'+ph+'</a:solidFill><a:solidFill>'+ph+'</a:solidFill><a:solidFill>'+ph+'</a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>');
  var emptyTree='<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';
  zip.file('ppt/slideMasters/slideMaster1.xml',XH+'<p:sldMaster xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'+emptyTree+'</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3600"><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>');
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',XH+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml',XH+'<p:sldLayout xmlns:a="'+NS_A+'" xmlns:r="'+NS_R+'" xmlns:p="'+NS_P+'" type="blank" preserve="1"><p:cSld name="Blank">'+emptyTree+'</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>');
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',XH+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  this.media.forEach(function(m,i){zip.file('ppt/media/image'+(i+1)+'.'+m.ext,m.b64,{base64:true,compression:'STORE'})});
  this.slides.forEach(function(sl,i){
    var map=new Map(),rels='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>';
    var relObj={use:function(mi){var id=map.get(mi);if(id)return id;id='rId'+(map.size+2);map.set(mi,id);rels+='<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+(mi+1)+'.'+self.media[mi].ext+'"/>';return id}};
    var xml=slideXml(sl,relObj);
    zip.file('ppt/slides/slide'+(i+1)+'.xml',xml);
    zip.file('ppt/slides/_rels/slide'+(i+1)+'.xml.rels',XH+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels+'</Relationships>');
  });
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'DEFLATE',compressionOptions:{level:6}});
};

/* ---------- driver ---------- */
function loadScript(url){return new Promise(function(res,rej){if(window.JSZip)return res();var s=document.createElement('script');s.src=url;s.onload=res;s.onerror=function(){rej(new Error('failed to load '+url))};document.head.appendChild(s)})}
function nextFrame(){return new Promise(function(r){requestAnimationFrame(function(){setTimeout(r,0)})})}
async function exportDeck(progress){
  progress=progress||function(){};
  await loadScript(JSZIP_URL);
  var stage=document.getElementById('stage');
  var holder=document.createElement('div');holder.className='pptx-exp';holder.setAttribute('aria-hidden','true');
  holder.style.cssText='position:fixed;left:-20000px;top:0;width:1280px;height:720px;overflow:hidden;z-index:-1;pointer-events:none;';
  var style=document.createElement('style');
  var BASE='.pptx-exp *{animation:none!important;transition:none!important;font-feature-settings:normal!important}';
  style.textContent=BASE+'.pptx-exp .slide{display:flex!important}';
  var clone=stage.cloneNode(true);clone.removeAttribute('id');
  clone.style.cssText='position:absolute;left:0;top:0;width:1280px;height:720px;transform:none;';
  clone.querySelectorAll('[data-step]').forEach(function(e){e.classList.add('shown')});
  holder.appendChild(clone);document.head.appendChild(style);document.body.appendChild(holder);
  var slides=Array.prototype.slice.call(clone.querySelectorAll('.slide')),ex=new Exporter();
  try{
    /* lay every slide out once so all font subsets/weights get requested, then wait for them */
    void holder.offsetHeight;
    await document.fonts.ready;await nextFrame();await document.fonts.ready;
    await Promise.all(Array.prototype.map.call(clone.querySelectorAll('img'),function(im){return im.decode?im.decode().catch(function(){}):null}));
    style.textContent=BASE;
    for(var i=0;i<slides.length;i++){
      slides.forEach(function(s){s.classList.remove('on')});slides[i].classList.add('on');
      void holder.offsetHeight;
      var base=holder.getBoundingClientRect(),ctx=new SlideCtx(ex,slides[i],base);
      ctx.run();
      ex.slides.push({ops:ctx.ops,bg:color(cs(slides[i]).backgroundColor)||{hex:'FFFFFF',a:1}});
      progress(i+1,slides.length);
      await nextFrame();
    }
    await Promise.all(ex.pending);
  }finally{holder.remove();style.remove()}
  return ex.buildZip(document.title);
}
window.__exportPptx=exportDeck;

function init(){
  var btn=document.getElementById('ppt');if(!btn)return;
  var label=btn.querySelector('span'),orig=label?label.textContent:'',busy=false;
  btn.addEventListener('click',async function(){
    if(busy)return;busy=true;btn.disabled=true;
    try{
      var blob=await exportDeck(function(i,n){if(label)label.textContent=i+'/'+n});
      var a=document.createElement('a');a.href=URL.createObjectURL(blob);
      a.download=(document.title.split('｜')[0].trim()||'slides')+'.pptx';
      document.body.appendChild(a);a.click();
      setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},4000);
    }catch(e){console.error(e);alert('PPTXの生成に失敗しました: '+(e&&e.message||e))}
    finally{busy=false;btn.disabled=false;if(label)label.textContent=orig}
  });
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
