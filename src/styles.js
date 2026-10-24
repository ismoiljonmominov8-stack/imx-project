/* IMX style engine.
   Every page style draws its animated background on a 1600x900 canvas from a time t (seconds),
   and lists the text boxes ("slots") that sit on top of it. In the exported PPTX the background
   becomes a video and the slots become real, editable PowerPoint text with entrance animations. */
(function(){
'use strict';
const W=1600,H=900;
const clamp=(v,a=0,b=1)=>Math.max(a,Math.min(b,v));
const seg=(t,a,b)=>clamp((t-a)/(b-a));
const eo=p=>1-Math.pow(1-p,3), eio=p=>p<.5?4*p*p*p:1-Math.pow(-2*p+2,3)/2;
const back=p=>{const c1=1.70158,c3=c1+1;return 1+c3*Math.pow(p-1,3)+c1*Math.pow(p-1,2)};
const bounce=p=>{const n=7.5625,d=2.75;if(p<1/d)return n*p*p;if(p<2/d)return n*(p-=1.5/d)*p+.75;if(p<2.5/d)return n*(p-=2.25/d)*p+.9375;return n*(p-=2.625/d)*p+.984375};
const lerp=(a,b,k)=>a+(b-a)*k;
function hexRgb(h){const n=parseInt(h.slice(1),16);return[(n>>16)&255,(n>>8)&255,n&255]}
function mix(a,b,k){const A=hexRgb(a),B=hexRgb(b);return '#'+A.map((v,i)=>Math.round(v+(B[i]-v)*k).toString(16).padStart(2,'0')).join('')}
function rgba(h,a){const [r,g,b]=hexRgb(h);return `rgba(${r},${g},${b},${a})`}

/* ---------- theme ---------- */
const THEMES={
  navy:{name:'Navy & Gold',bg:'#16264a',accent:'#d9b25f',text:'#ece6d8'},
  emerald:{name:'Emerald',bg:'#0f3b33',accent:'#e3c26f',text:'#eef2ea'},
  burgundy:{name:'Burgundy',bg:'#4a1426',accent:'#e9b86a',text:'#f5e9e2'},
  mono:{name:'Black & White',bg:'#151515',accent:'#f2f2f2',text:'#e6e6e6'},
  ocean:{name:'Ocean',bg:'#0d3a5c',accent:'#5fd3e0',text:'#e8f4f8'},
  sunset:{name:'Sunset',bg:'#3b1f4a',accent:'#ff9a5a',text:'#fbeee6'},
  forest:{name:'Forest',bg:'#1f2e1a',accent:'#b9d46a',text:'#eef0e4'},
  purple:{name:'Royal Purple',bg:'#261a4a',accent:'#c8a2ff',text:'#efeaff'}
};
const FONTS={
  head:['Georgia','Cambria','Constantia','Bahnschrift','Century Gothic','Impact','Times New Roman','Arial Black'],
  body:['Calibri','Segoe UI','Arial','Corbel','Candara','Times New Roman','Georgia','Verdana']
};
const SIZES={S:.86,M:1,L:1.18};
/* resolved colours for the current drawing */
let C={},F={},IMG={};
function setTheme(T){
  const bg=T.bg,ac=T.accent;
  C={bg,bg2:mix(bg,'#000000',.5),bg3:mix(bg,'#000000',.78),panelA:mix(bg,'#ffffff',.08),panelB:mix(bg,'#000000',.45),
     gold:ac,hi:mix(ac,'#ffffff',.45),lo:mix(ac,'#000000',.4),deep:mix(ac,'#000000',.72),text:T.text,muted:mix(T.text,bg,.4),dark:mix(bg,'#000000',.6)};
  F={head:`"${T.headFont||'Georgia'}", Georgia, serif`,body:`"${T.bodyFont||'Calibri'}", Calibri, Arial, sans-serif`};
  IMG={bg:T.bgImg||null,vis:T.bgVis??.35,...(T.images||{})};
}

/* ---------- drawing helpers ---------- */
let x=null;
function rr(X,Y,w,h,r){x.beginPath();x.roundRect(X,Y,w,h,r)}
function goldFill(y0,y1){const g=x.createLinearGradient(0,y0,0,y1);g.addColorStop(0,C.hi);g.addColorStop(.55,C.gold);g.addColorStop(1,C.lo);return g}
function panel(X,Y,w,h,{r=14,alpha=1,stroke=C.gold}={}){
  x.save();x.globalAlpha*=alpha;const g=x.createLinearGradient(X,Y,X+w,Y+h);g.addColorStop(0,C.panelA);g.addColorStop(1,C.panelB);
  x.shadowColor='rgba(0,0,0,.55)';x.shadowBlur=30;x.shadowOffsetY=14;rr(X,Y,w,h,r);x.fillStyle=g;x.fill();
  x.shadowColor='transparent';x.lineWidth=2;x.strokeStyle=stroke;x.stroke();x.restore();
}
function cover(img,X,Y,w,h){const k=Math.max(w/img.width,h/img.height),iw=img.width*k,ih=img.height*k;x.drawImage(img,X+(w-iw)/2,Y+(h-ih)/2,iw,ih)}
const DUST=Array.from({length:70},(_,i)=>({x:(i*397)%1600,y:(i*571)%900,r:.6+(i%5)*.35,v:.15+(i%7)*.05,a:i*1.3}));
function bg(t,{dust=true}={}){
  const g=x.createRadialGradient(800,450,40,800,450,950);g.addColorStop(0,C.bg);g.addColorStop(.5,C.bg2);g.addColorStop(1,C.bg3);
  x.fillStyle=g;x.fillRect(0,0,W,H);
  if(IMG.bg){x.save();x.globalAlpha=IMG.vis;cover(IMG.bg,0,0,W,H);x.restore()}
  if(dust)for(const p of DUST){const y=((p.y-t*p.v*60)%900+900)%900;
    x.fillStyle=rgba(C.hi,.2+.35*Math.abs(Math.sin(t*.7+p.a)));x.beginPath();x.arc(p.x+Math.sin(t+p.a)*6,y,p.r,0,7);x.fill()}
}
/* a placeholder picture, used until the user adds a photo */
function picture(X,Y,w,h,hue=0,img=null){
  x.save();rr(X,Y,w,h,6);x.clip();
  if(img){cover(img,X,Y,w,h);x.restore();return}
  const g=x.createLinearGradient(X,Y,X+w,Y+h);g.addColorStop(0,mix(C.bg,'#ffffff',.12+.06*(hue%3)));g.addColorStop(1,mix(C.bg,C.gold,.35));x.fillStyle=g;x.fillRect(X,Y,w,h);
  x.globalAlpha=.3;x.fillStyle=C.gold;x.beginPath();x.arc(X+w*.72,Y+h*.35,h*.28,0,7);x.fill();
  x.globalAlpha=.18;x.fillStyle='#ffffff';x.beginPath();x.arc(X+w*.3,Y+h*.7,h*.38,0,7);x.fill();
  x.globalAlpha=.8;const s=Math.min(w,h)/6,cx=X+w/2,cy=Y+h/2;x.strokeStyle=C.hi;x.lineWidth=Math.max(2,s/14);
  x.beginPath();x.roundRect(cx-s,cy-s*.75,s*2,s*1.5,s*.15);x.stroke();
  x.beginPath();x.moveTo(cx-s*.8,cy+s*.5);x.lineTo(cx-s*.25,cy-s*.1);x.lineTo(cx+s*.1,cy+s*.25);x.lineTo(cx+s*.4,cy);x.lineTo(cx+s*.8,cy+s*.5);x.stroke();
  x.beginPath();x.arc(cx+s*.45,cy-s*.35,s*.15,0,7);x.stroke();x.restore();
}
const HEAD_PX=64;
function textWidth(s,font,size,weight=700){x.save();x.font=`${weight} ${size}px ${font}`;const w=x.measureText(s).width;x.restore();return w}
/* the gold lines either side of a page heading */
function headDeco(t,at,head,mult){
  const p=eo(seg(t,at,at+.9));if(!p)return;const tw=Math.min(1200,textWidth(head||'',F.head,HEAD_PX*mult)+head.length*5);
  x.save();x.globalAlpha=p;x.strokeStyle=C.gold;x.lineWidth=2;const half=tw/2+40,w=lerp(half,560,p);
  x.beginPath();x.moveTo(800-w,128);x.lineTo(800-half,128);x.moveTo(800+half,128);x.lineTo(800+w,128);x.stroke();x.restore();
}
function headSlot(f,at){return {key:'heading',x:200,y:66,w:1200,h:80,role:'h',align:'c',valign:'m',at,text:f.heading||''}}

/* ================= the styles ================= */
const S={};

/* 9. Split Screen */
S.split={n:9,name:{uz:'Ikki tomon',ru:'Разделённый экран',en:'Split Screen'},use:{uz:'matn va rasm yonma-yon',ru:'текст рядом с картинкой',en:'text beside a picture'},dur:3.4,
  fields:[{k:'heading',t:'text'},{k:'text',t:'area'},{k:'img',t:'img'}],
  sample:{heading:'Kirish',text:'Bu yerda sahifaning asosiy matni joylashadi. Fikringizni qisqa va aniq yozing.\nIkkinchi xatboshi ham shu yerda.'},
  draw(t,f){bg(t);const a=eio(seg(t,.2,1.1)),b=eio(seg(t,.35,1.25));
    x.save();x.translate(0,lerp(-H,0,a));x.beginPath();x.moveTo(0,0);x.lineTo(900,0);x.lineTo(740,H);x.lineTo(0,H);x.closePath();
    const g=x.createLinearGradient(0,0,900,H);g.addColorStop(0,C.panelA);g.addColorStop(1,C.panelB);x.fillStyle=g;x.fill();x.restore();
    x.save();x.translate(0,lerp(H,0,b));x.beginPath();x.moveTo(900,0);x.lineTo(W,0);x.lineTo(W,H);x.lineTo(740,H);x.closePath();x.clip();
    const z=1.12-.12*eo(seg(t,1,6));x.translate(1170,450);x.scale(z,z);x.translate(-1170,-450);picture(700,0,900,H,2,IMG[f.img]);x.restore();
    const d=eo(seg(t,1.2,1.8));x.strokeStyle=C.hi;x.lineWidth=6;x.beginPath();x.moveTo(900,0);x.lineTo(lerp(900,740,d),H*d);x.stroke();
    const ul=eo(seg(t,1.9,2.4));x.fillStyle=C.gold;x.fillRect(100,300,240*ul,4);},
  slots(f){return[{key:'heading',x:100,y:170,w:620,h:120,role:'h',align:'l',valign:'b',at:1.5,text:f.heading||''},
    {key:'text',x:100,y:340,w:600,h:470,role:'b',align:'l',at:2.1,text:f.text||''}]}};

/* 10. Key Points */
function cardsRow(n,cwMax=300,gap=40,Y=270,ch=460){const cw=Math.min(cwMax,(1440-(n-1)*gap)/n),X0=800-(n*cw+(n-1)*gap)/2;return{cw,ch,gap,Y,X0}}
S.points={n:10,name:{uz:'Asosiy fikrlar',ru:'Ключевые пункты',en:'Key Points'},use:{uz:'3–5 asosiy fikr kartalarda',ru:'3–5 главных идей на карточках',en:'3–5 main ideas as cards'},dur:3.6,
  fields:[{k:'heading',t:'text'},{k:'items',t:'list',min:3,max:5,of:[{k:'title',t:'text'},{k:'text',t:'area'}]}],
  sample:{heading:'Asosiy fikrlar',items:[{title:'Birinchi',text:'Qisqa izoh'},{title:'Ikkinchi',text:'Qisqa izoh'},{title:'Uchinchi',text:'Qisqa izoh'},{title:'To\'rtinchi',text:'Qisqa izoh'}]},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const n=clamp((f.items||[]).length,1,5),L=cardsRow(n);
    for(let i=n-1;i>=0;i--){const a=.6+i*.3,p=seg(t,a,a+.7),e=back(p);
      const fx=800-L.cw/2,fy=L.Y+560,tx=L.X0+i*(L.cw+L.gap),cx=lerp(fx,tx,e),cy=lerp(fy,L.Y,e),rot=lerp(.4-i*.15,0,e);
      x.save();x.translate(cx+L.cw/2,cy+L.ch/2);x.rotate(rot);x.translate(-L.cw/2,-L.ch/2);panel(0,0,L.cw,L.ch,{});
      const q=eo(seg(t,a+.5,a+1));x.globalAlpha=q;x.fillStyle=goldFill(40,140);x.beginPath();x.arc(L.cw/2,96,48,0,7);x.fill();
      x.fillStyle=C.dark;x.font=`700 54px ${F.head}`;x.textAlign='center';x.textBaseline='middle';x.fillText(String(i+1),L.cw/2,100);x.restore()}
    const h=t-(1.4+n*.3);if(h>0){const i=Math.floor(h/1.2)%n,pp=(h%1.2)/1.2,tx=L.X0+i*(L.cw+L.gap);
      x.save();x.globalAlpha=Math.sin(pp*Math.PI)*.9;x.strokeStyle=C.hi;x.lineWidth=4;x.shadowColor=C.hi;x.shadowBlur=22;rr(tx-6,L.Y-6,L.cw+12,L.ch+12,18);x.stroke();x.restore()}},
  slots(f){const it=f.items||[],n=clamp(it.length,1,5),L=cardsRow(n),out=[headSlot(f,.2)];
    it.slice(0,5).forEach((o,i)=>{const X=L.X0+i*(L.cw+L.gap),a=.6+i*.3+.9;
      out.push({key:`items.${i}.title`,x:X+22,y:L.Y+170,w:L.cw-44,h:70,role:'t',align:'c',valign:'m',at:a,text:o.title||''});
      out.push({key:`items.${i}.text`,x:X+22,y:L.Y+250,w:L.cw-44,h:190,role:'s',align:'c',at:a+.15,text:o.text||''})});return out}};

/* 11. Puzzle */
function piecePath(X,Y,w,h,tabs){const k=Math.min(w,h)*.16;x.beginPath();x.moveTo(X,Y);
  const edge=(ax,ay,bx,by,s,nx,ny)=>{if(!s){x.lineTo(bx,by);return}const mx=(ax+bx)/2,my=(ay+by)/2,dx=bx-ax,dy=by-ay;
    x.lineTo(ax+dx*.38,ay+dy*.38);x.bezierCurveTo(ax+dx*.32+nx*k*s,ay+dy*.32+ny*k*s,mx-dx*.14+nx*k*1.4*s,my-dy*.14+ny*k*1.4*s,mx+nx*k*1.4*s,my+ny*k*1.4*s);
    x.bezierCurveTo(mx+dx*.14+nx*k*1.4*s,my+dy*.14+ny*k*1.4*s,ax+dx*.68+nx*k*s,ay+dy*.68+ny*k*s,ax+dx*.62,ay+dy*.62);x.lineTo(bx,by)};
  edge(X,Y,X+w,Y,tabs[0],0,-1);edge(X+w,Y,X+w,Y+h,tabs[1],1,0);edge(X+w,Y+h,X,Y+h,tabs[2],0,1);edge(X,Y+h,X,Y,tabs[3],-1,0);x.closePath()}
function puzzleGrid(n){const cols=n<=4?2:3,rows=Math.ceil(n/cols),w=cols===2?520:400,h=rows===1?440:300;return{cols,rows,w,h,X0:800-cols*w/2,Y0:180+(rows===1?60:0)}}
S.puzzle={n:11,name:{uz:'Pazl',ru:'Пазл',en:'Puzzle'},use:{uz:'bir g\'oyaning 4–6 qismi',ru:'4–6 частей одной идеи',en:'4–6 parts of one idea'},dur:3.6,
  fields:[{k:'heading',t:'text'},{k:'items',t:'list',min:4,max:6,of:[{k:'title',t:'text'},{k:'text',t:'area'}]}],
  sample:{heading:'Bosqichlar',items:[1,2,3,4,5,6].map(i=>({title:`${i}-qism`,text:'Qisqa izoh'}))},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const n=clamp((f.items||[]).length,4,6),G=puzzleGrid(n);
    const Hs=(r,c)=>(r+c)%2?1:-1,Vs=(r,c)=>c%2?1:-1;let i=0;
    for(let r=0;r<G.rows;r++)for(let c=0;c<G.cols;c++,i++){if(i>=n)break;
      const a=.5+i*.32,p=seg(t,a,a+.8),e=back(p);if(!p)continue;
      const fx=[-400,1900,-300,1800,700,800][i],fy=[-300,-200,1100,1200,-400,1300][i],rot=[-1.2,1,.8,-.9,1.4,-1.1][i];
      const tx=G.X0+c*G.w,ty=G.Y0+r*G.h,last=r===G.rows-1||(r+1)*G.cols+c>=n;
      const tb=[r?-Vs(r-1,c):0,c<G.cols-1&&i+1<n?Hs(r,c):0,!last?Vs(r,c):0,c?-Hs(r,c-1):0];
      x.save();x.translate(lerp(fx,tx,e)+G.w/2,lerp(fy,ty,e)+G.h/2);x.rotate(rot*(1-e));x.translate(-G.w/2,-G.h/2);
      piecePath(0,0,G.w,G.h,tb);const g=x.createLinearGradient(0,0,G.w,G.h);g.addColorStop(0,C.panelA);g.addColorStop(1,C.panelB);
      x.shadowColor='rgba(0,0,0,.55)';x.shadowBlur=24;x.shadowOffsetY=10;x.fillStyle=g;x.fill();x.shadowColor='transparent';x.lineWidth=3;x.strokeStyle=C.gold;x.stroke();
      x.fillStyle=goldFill(40,100);x.font=`700 46px ${F.head}`;x.textAlign='left';x.textBaseline='alphabetic';x.fillText(String(i+1).padStart(2,'0'),46,92);x.restore();
      const fl=seg(t,a+.8,a+1.2);if(fl>0&&fl<1){x.save();x.globalAlpha=1-fl;x.strokeStyle='#ffffff';x.lineWidth=8;piecePath(tx,ty,G.w,G.h,tb);x.stroke();x.restore()}}},
  slots(f){const it=f.items||[],n=clamp(it.length,4,6),G=puzzleGrid(n),out=[headSlot(f,.2)];
    it.slice(0,n).forEach((o,i)=>{const c=i%G.cols,r=Math.floor(i/G.cols),X=G.X0+c*G.w,Y=G.Y0+r*G.h,a=.5+i*.32+1;
      out.push({key:`items.${i}.title`,x:X+130,y:Y+40,w:G.w-170,h:62,role:'t',align:'l',valign:'m',at:a,text:o.title||''});
      out.push({key:`items.${i}.text`,x:X+46,y:Y+120,w:G.w-92,h:G.h-150,role:'s',align:'l',at:a+.15,text:o.text||''})});return out}};

/* 12. Polaroid Table */
const POL=[[340,500,-.16],[800,520,.05],[1260,480,.15],[570,540,.22],[1030,560,-.2]];
function polLayout(n){const sets={3:[[400,500,-.12],[800,520,.04],[1200,500,.12]],4:[[330,500,-.14],[650,530,.06],[950,500,-.05],[1270,520,.13]],5:POL};return sets[clamp(n,3,5)]}
S.gallery={n:12,name:{uz:'Foto stol',ru:'Фото на столе',en:'Polaroid Table'},use:{uz:'rasmlar galereyasi',ru:'галерея фотографий',en:'photo gallery'},dur:3.4,
  fields:[{k:'heading',t:'text'},{k:'photos',t:'list',min:3,max:5,of:[{k:'img',t:'img'},{k:'caption',t:'text'}]}],
  sample:{heading:'Galereya',photos:[1,2,3,4,5].map(i=>({caption:`Rasm ${i}`}))},
  draw(t,f,m){bg(t,{dust:false});x.fillStyle='rgba(0,0,0,.25)';x.fillRect(0,0,W,H);
    for(let i=0;i<H;i+=36){x.strokeStyle='rgba(0,0,0,.18)';x.beginPath();x.moveTo(0,i+Math.sin(i)*6);x.lineTo(W,i+Math.cos(i)*6);x.stroke()}
    headDeco(t,0,f.heading,m);const ph=f.photos||[],L=polLayout(ph.length);
    L.forEach(([px,py,r],i)=>{const a=.4+i*.4,p=seg(t,a,a+.8);if(!p)return;const e=eo(p),s=lerp(1.4,.74,e);
      x.save();x.translate(px,lerp(py-760,py,e));x.rotate(r+(1-e)*1.4);x.scale(s,s);
      x.shadowColor='rgba(0,0,0,.6)';x.shadowBlur=30;x.shadowOffsetY=16;x.fillStyle='#f3ead6';x.fillRect(-170,-200,340,400);x.shadowColor='transparent';
      picture(-150,-180,300,290,i,IMG[(ph[i]||{}).img]);x.restore()})},
  slots(f){const ph=f.photos||[],L=polLayout(ph.length),out=[headSlot(f,.2)];
    L.forEach(([px,py,r],i)=>{const s=.74,cx=px-Math.sin(r)*150*s,cy=py+Math.cos(r)*150*s;
      out.push({key:`photos.${i}.caption`,x:cx-120,y:cy-24,w:240,h:48,role:'s',align:'c',valign:'m',at:.4+i*.4+.9,rot:r*180/Math.PI,color:'#2b2216',text:(ph[i]||{}).caption||''})});return out}};

/* 13. Domino Process */
function domLayout(n){const gap=Math.min(240,1400/n);return{gap,X0:800-(n-1)*gap/2,floor:760,dw:42,dh:190}}
S.domino={n:13,name:{uz:'Domino jarayon',ru:'Домино-процесс',en:'Domino Process'},use:{uz:'qadamma-qadam jarayon',ru:'пошаговый процесс',en:'step-by-step process'},dur:4,
  fields:[{k:'heading',t:'text'},{k:'items',t:'list',min:3,max:6,of:[{k:'title',t:'area'}]}],
  sample:{heading:'Jarayon',items:[1,2,3,4,5].map(i=>({title:`${i}-qadam qisqa matni`}))},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const n=clamp((f.items||[]).length,1,6),L=domLayout(n),push=1.1;
    x.strokeStyle=rgba(C.gold,.5);x.lineWidth=2;x.beginPath();x.moveTo(80,L.floor);x.lineTo(1520,L.floor);x.stroke();
    for(let i=0;i<n;i++){const a=push+i*.32,p=seg(t,a,a+.42),ang=eio(p)*(i<n-1?.92:1.32),px=L.X0+i*L.gap;
      x.save();x.translate(px+L.dw/2,L.floor);x.rotate(ang);const g=x.createLinearGradient(-L.dw,-L.dh,0,0);g.addColorStop(0,C.hi);g.addColorStop(1,C.lo);
      x.shadowColor='rgba(0,0,0,.5)';x.shadowBlur=16;x.fillStyle=g;rr(-L.dw,-L.dh,L.dw,L.dh,6);x.fill();x.shadowColor='transparent';
      x.fillStyle=C.deep;x.fillRect(-L.dw+6,-L.dh/2-2,L.dw-12,4);x.beginPath();x.arc(-L.dw/2,-L.dh*.75,6,0,7);x.arc(-L.dw/2,-L.dh*.25,6,0,7);x.fill();x.restore();
      const q=back(seg(t,a+.3,a+.8));if(q>0){const cw=L.gap-30;panel(px-cw/2,210+(1-q)*60,cw,300,{alpha:clamp(q)});
        x.save();x.globalAlpha=clamp(q);x.fillStyle=goldFill(240,300);x.font=`700 60px ${F.head}`;x.textAlign='center';x.fillText(String(i+1),px,300+(1-q)*60);x.restore()}}
    const h=seg(t,.4,1.1);if(h<1){x.save();x.globalAlpha=1-seg(t,1,1.3);x.fillStyle=C.hi;rr(lerp(L.X0-300,L.X0-50,eio(h)),L.floor-160,110,30,15);x.fill();x.restore()}},
  slots(f){const it=f.items||[],n=clamp(it.length,1,6),L=domLayout(n),out=[headSlot(f,.2)];
    it.slice(0,6).forEach((o,i)=>{const px=L.X0+i*L.gap,cw=L.gap-30;out.push({key:`items.${i}.title`,x:px-cw/2+14,y:330,w:cw-28,h:165,role:'s',align:'c',at:1.1+i*.32+.8,text:o.title||''})});return out}};

/* 14. Rocket Stats: the numbers count up inside the animation */
function parseVal(v){const s=String(v||'').trim(),m=s.match(/^([^\d-]*)(-?[\d\s]*[.,]?\d*)(.*)$/);if(!m||!m[2].trim())return null;
  const comma=m[2].includes(','),num=parseFloat(m[2].replace(/\s/g,'').replace(',','.')),dec=(m[2].split(/[.,]/)[1]||'').length;return{pre:m[1],num,dec,comma,post:m[3]}}
function fmtVal(P,k){if(!P)return'';let s=(P.num*k).toFixed(P.dec);if(P.comma)s=s.replace('.',',');return P.pre+s+P.post}
S.stats={n:14,name:{uz:'Raketa raqamlar',ru:'Ракета и цифры',en:'Rocket Stats'},use:{uz:'asosiy raqamlar',ru:'ключевые цифры',en:'key numbers'},dur:4.8,
  fields:[{k:'heading',t:'text'},{k:'stats',t:'list',min:1,max:3,of:[{k:'value',t:'text'},{k:'label',t:'text'}]}],
  sample:{heading:'Raqamlar',stats:[{value:'8,2 mln',label:'Ko\'rsatkich 1'},{value:'24%',label:'Ko\'rsatkich 2'},{value:'120+',label:'Ko\'rsatkich 3'}]},
  draw(t,f,m){bg(t);const lift=eio(seg(t,.8,2.6)),ry=lerp(760,-260,lift),shake=t<1?Math.sin(t*90)*3:0;
    x.fillStyle=C.panelB;x.fillRect(640,800,320,24);
    for(let k=0;k<14;k++){const p=seg(t,.3+k*.05,2+k*.05);if(!p||p>=1)continue;const dir=k%2?1:-1;
      x.fillStyle=`rgba(200,204,214,${.45*(1-p)})`;x.beginPath();x.arc(800+dir*p*(140+k*20),790-p*40,30+p*70,0,7);x.fill()}
    x.save();x.translate(800+shake,ry);
    if(t>.5){for(let q=0;q<3;q++){const fl=60+Math.sin(t*40+q)*14+lift*50;x.fillStyle=['#fff6d8',C.hi,'#e08a3a'][q];x.globalAlpha=.9-q*.2;
      x.beginPath();x.moveTo(-26+q*6,60);x.quadraticCurveTo(0,60+fl*(1.6-q*.3),26-q*6,60);x.fill()}x.globalAlpha=1}
    const g=x.createLinearGradient(-40,0,40,0);g.addColorStop(0,C.lo);g.addColorStop(.5,'#ffffff');g.addColorStop(1,C.lo);
    x.fillStyle=g;x.beginPath();x.moveTo(0,-140);x.bezierCurveTo(48,-90,46,10,38,60);x.lineTo(-38,60);x.bezierCurveTo(-46,10,-48,-90,0,-140);x.fill();
    x.fillStyle=C.dark;x.beginPath();x.arc(0,-40,16,0,7);x.fill();x.strokeStyle=C.gold;x.lineWidth=4;x.stroke();
    x.fillStyle=C.lo;[-1,1].forEach(s=>{x.beginPath();x.moveTo(38*s,10);x.lineTo(74*s,70);x.lineTo(36*s,58);x.fill()});x.restore();
    headDeco(t,2.2,f.heading,m);const st=(f.stats||[]).slice(0,3),n=st.length||1;
    st.forEach((o,i)=>{const a=2.5+i*.3,p=eo(seg(t,a,a+1.6)),cx=800+(i-(n-1)/2)*440,al=clamp(seg(t,a,a+.4));if(!al)return;
      panel(cx-180,300+(1-p)*50,360,300,{alpha:al});
      x.save();x.globalAlpha=al;x.fillStyle=goldFill(380,480);const P=parseVal(o.value),s=P?fmtVal(P,p):(o.value||'');
      let size=100*m;x.font=`700 ${size}px ${F.head}`;while(x.measureText(s).width>320&&size>40){size-=4;x.font=`700 ${size}px ${F.head}`}
      x.textAlign='center';x.fillText(s,cx,470+(1-p)*50);x.restore()})},
  slots(f){const st=(f.stats||[]).slice(0,3),n=st.length||1,out=[headSlot(f,2.4)];
    st.forEach((o,i)=>{const cx=800+(i-(n-1)/2)*440;out.push({key:`stats.${i}.label`,x:cx-160,y:500,w:320,h:80,role:'s',align:'c',valign:'t',at:2.5+i*.3+.9,color:'muted',text:o.label||''})});return out}};

/* 15. Balance Scale */
const BAL={px:800,py:340,arm:420,tilt:-.1};
S.balance={n:15,name:{uz:'Tarozi',ru:'Весы',en:'Balance Scale'},use:{uz:'oldin/keyin, afzallik/kamchilik',ru:'до/после, плюсы/минусы',en:'before/after, pros/cons'},dur:4.4,
  fields:[{k:'heading',t:'text'},{k:'sides',t:'list',min:2,max:2,of:[{k:'title',t:'text'},{k:'text',t:'area'}]}],
  sample:{heading:'Taqqoslash',sides:[{title:'Oldin',text:'Birinchi holat\nIkkinchi holat'},{title:'Keyin',text:'Yangi holat\nNatija'}]},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const {px,py,arm}=BAL;
    const tilt=t<2?0:t<3.3?.16*eo(seg(t,2,2.6))+Math.sin((t-2)*9)*.04*clamp(1-(t-2)/1.3):lerp(.16,BAL.tilt,eo(seg(t,3.3,4)))+Math.sin((t-3.3)*9)*.03*clamp(1-(t-3.3)/1.2);
    x.fillStyle=C.gold;x.fillRect(px-8,py,16,400);x.fillRect(px-110,py+392,220,18);
    x.save();x.translate(px,py);x.rotate(tilt);x.fillStyle=goldFill(-10,10);rr(-arm,-9,arm*2,18,9);x.fill();x.restore();
    x.fillStyle=C.hi;x.beginPath();x.arc(px,py,16,0,7);x.fill();
    [[-1,[1,1.3,1.6]],[1,[2.6,2.9,3.2]]].forEach(([s,drops])=>{const ex=px+Math.cos(tilt)*arm*s,ey=py+Math.sin(tilt)*arm*s;
      x.strokeStyle=C.lo;x.lineWidth=2;x.beginPath();x.moveTo(ex,ey);x.lineTo(ex-120,ey+150);x.moveTo(ex,ey);x.lineTo(ex+120,ey+150);x.stroke();
      x.fillStyle=goldFill(ey+140,ey+190);x.beginPath();x.ellipse(ex,ey+156,150,30,0,0,Math.PI);x.fill();
      drops.forEach((a,k)=>{const sg=seg(t,a,a+.6);if(!sg)return;const by=lerp(ey-500,ey+140-k*32,bounce(sg));panel(ex-80+k*8,by-28,160,28,{r:6})})})},
  slots(f){const sd=f.sides||[],out=[headSlot(f,.2)],{px,py,arm,tilt}=BAL;
    [-1,1].forEach((s,i)=>{const ex=px+Math.cos(tilt)*arm*s,ey=py+Math.sin(tilt)*arm*s,o=sd[i]||{};
      out.push({key:`sides.${i}.title`,x:ex-200,y:ey+200,w:400,h:64,role:'t',align:'c',valign:'m',at:4.1,text:o.title||''});
      out.push({key:`sides.${i}.text`,x:ex-200,y:ey+270,w:400,h:170,role:'s',align:'c',at:4.3,text:o.text||''})});return out}};

/* 16. Window Blinds */
S.blinds={n:16,name:{uz:'Jalyuzi',ru:'Жалюзи',en:'Window Blinds'},use:{uz:'bitta katta rasm va izoh',ru:'одно большое фото с подписью',en:'one big photo with caption'},dur:3,
  fields:[{k:'heading',t:'text'},{k:'text',t:'text'},{k:'img',t:'img'}],
  sample:{heading:'Sarlavha',text:'Rasm izohi shu yerda'},
  draw(t,f){x.fillStyle=C.bg3;x.fillRect(0,0,W,H);picture(0,0,W,H,1,IMG[f.img]);x.fillStyle='rgba(0,0,0,.18)';x.fillRect(0,0,W,H);
    const n=12,sh=H/n;for(let i=0;i<n;i++){const a=.5+i*.08,p=eio(seg(t,a,a+.8)),hh=sh*(1-p);if(hh<.5)continue;const yc=i*sh+sh/2;
      const g=x.createLinearGradient(0,yc-hh/2,0,yc+hh/2);g.addColorStop(0,C.panelA);g.addColorStop(.5,C.bg);g.addColorStop(1,C.bg3);x.fillStyle=g;x.fillRect(0,yc-hh/2,W,hh);
      x.fillStyle=rgba(C.gold,.6*(1-p));x.fillRect(0,yc+hh/2-2,W,2)}
    const cp=1-seg(t,.4,1.3);if(cp>0){x.strokeStyle=rgba(C.hi,cp);x.lineWidth=3;[400,1200].forEach(cx=>{x.beginPath();x.moveTo(cx,0);x.lineTo(cx,H);x.stroke()})}
    const q=eo(seg(t,2,2.7));if(q>0){x.save();x.globalAlpha=q;const g=x.createLinearGradient(0,580,0,H);g.addColorStop(0,'rgba(0,0,0,0)');g.addColorStop(1,rgba(C.bg3,.93));x.fillStyle=g;x.fillRect(0,560,W,340);
      x.fillStyle=C.gold;x.fillRect(120,700,200*q,4);x.restore()}},
  slots(f){return[{key:'heading',x:120,y:590,w:1360,h:100,role:'h',align:'l',valign:'b',at:2.3,text:f.heading||''},
    {key:'text',x:120,y:720,w:1360,h:110,role:'b',align:'l',at:2.6,text:f.text||''}]}};

/* 17. Spotlight Stage */
function spotX(n){return Array.from({length:n},(_,i)=>800+(i-(n-1)/2)*Math.min(400,1400/n))}
S.spotlight={n:17,name:{uz:'Sahna nuri',ru:'Прожектор',en:'Spotlight Stage'},use:{uz:'jamoa, mutaxassislar',ru:'команда, эксперты',en:'team, people, experts'},dur:4,
  fields:[{k:'heading',t:'text'},{k:'people',t:'list',min:1,max:4,of:[{k:'img',t:'img'},{k:'name',t:'text'},{k:'role',t:'text'}]}],
  sample:{heading:'Jamoa',people:[{name:'Ism Familiya',role:'Lavozim'},{name:'Ism Familiya',role:'Lavozim'},{name:'Ism Familiya',role:'Lavozim'}]},
  draw(t,f,m){x.fillStyle=C.bg3;x.fillRect(0,0,W,H);if(IMG.bg){x.save();x.globalAlpha=IMG.vis*.6;cover(IMG.bg,0,0,W,H);x.restore()}
    const fg=x.createLinearGradient(0,620,0,H);fg.addColorStop(0,C.bg2);fg.addColorStop(1,C.bg3);x.fillStyle=fg;x.fillRect(0,640,W,260);
    const pp=f.people||[],X=spotX(clamp(pp.length,1,4));
    X.forEach((cx,i)=>{const a=.5+i*.7,on=seg(t,a,a+.15),fl=on<1?(Math.sin(t*80)>0?.4:0):1,L=on*fl;
      if(L>0){x.save();x.globalCompositeOperation='lighter';const g=x.createLinearGradient(cx,0,cx,700);g.addColorStop(0,`rgba(255,240,200,${.26*L})`);g.addColorStop(1,`rgba(255,240,200,${.05*L})`);
        x.fillStyle=g;x.beginPath();x.moveTo(cx-30,0);x.lineTo(cx+30,0);x.lineTo(cx+220,700);x.lineTo(cx-220,700);x.fill();
        x.fillStyle=`rgba(255,236,190,${.2*L})`;x.beginPath();x.ellipse(cx,700,220,40,0,0,7);x.fill();x.restore()}
      x.fillStyle=C.panelA;rr(cx-26,-6,52,30,6);x.fill();x.fillStyle=L?'#fff6d8':C.panelB;x.fillRect(cx-18,20,36,6);
      const p=eo(seg(t,a+.2,a+.8));if(p>0){const o=pp[i]||{},Y=290+(1-p)*30;x.save();x.globalAlpha=p;panel(cx-150,Y,300,400,{});
        x.save();x.beginPath();x.arc(cx,Y+110,74,0,7);x.clip();if(IMG[o.img])cover(IMG[o.img],cx-74,Y+36,148,148);else{x.fillStyle=C.panelA;x.fillRect(cx-74,Y+36,148,148);
          x.fillStyle=goldFill(Y+80,Y+140);x.font=`700 56px ${F.head}`;x.textAlign='center';x.textBaseline='middle';x.fillText((o.name||'?').split(/\s+/).map(w=>w[0]||'').join('').slice(0,2).toUpperCase(),cx,Y+114)}
        x.restore();x.strokeStyle=C.gold;x.lineWidth=3;x.beginPath();x.arc(cx,Y+110,74,0,7);x.stroke();x.restore()}});
    headDeco(t,2.6,f.heading,m)},
  slots(f){const pp=f.people||[],X=spotX(clamp(pp.length,1,4)),out=[headSlot(f,2.8)];
    X.forEach((cx,i)=>{const o=pp[i]||{},a=.5+i*.7+.8;
      out.push({key:`people.${i}.name`,x:cx-140,y:500,w:280,h:90,role:'t',align:'c',valign:'m',at:a,text:o.name||''});
      out.push({key:`people.${i}.role`,x:cx-140,y:595,w:280,h:70,role:'s',align:'c',at:a+.15,color:'accent',text:o.role||''})});return out}};

/* 18. Growing Tree */
const TREE={gx:800,gy:820,trunk:420,br:[[-2.6,250,0],[-.55,250,0],[-2.2,230,1],[-.95,230,1],[-1.57,140,2]]};
function treeEnds(){const {gx,gy,trunk}=TREE,bases=[[gx,gy-250],[gx,gy-370],[gx,gy-trunk]];
  return TREE.br.map(([ang,len,b])=>{const [bx,by]=bases[b];return[bx+Math.cos(ang)*len,by+Math.sin(ang)*len,ang]})}
S.tree={n:18,name:{uz:'O\'suvchi daraxt',ru:'Растущее дерево',en:'Growing Tree'},use:{uz:'g\'oyalar, aql xaritasi',ru:'идеи, карта мыслей',en:'ideas, mind map'},dur:4.4,
  fields:[{k:'heading',t:'text'},{k:'items',t:'list',min:3,max:5,of:[{k:'title',t:'text'}]}],
  sample:{heading:'G\'oyalar',items:[{title:'G\'oya 1'},{title:'G\'oya 2'},{title:'G\'oya 3'},{title:'G\'oya 4'},{title:'Asosiy g\'oya'}]},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const {gx,gy,trunk}=TREE,n=clamp((f.items||[]).length,1,5);
    x.fillStyle=rgba(C.gold,.25);x.beginPath();x.ellipse(gx,gy,300,24,0,0,7);x.fill();
    const sd=seg(t,.2,.8);if(sd<1){x.fillStyle=C.gold;x.beginPath();x.ellipse(gx,lerp(120,gy-8,bounce(sd)),12,16,0,0,7);x.fill()}
    x.strokeStyle=goldFill(300,800);x.lineCap='round';const tr=eo(seg(t,.8,1.8));
    const br=(x0,y0,ang,len,p,w)=>{const x1=x0+Math.cos(ang)*len*p,y1=y0+Math.sin(ang)*len*p;x.lineWidth=w;x.beginPath();x.moveTo(x0,y0);x.quadraticCurveTo((x0+x1)/2+Math.sin(ang)*20,(y0+y1)/2,x1,y1);x.stroke()};
    br(gx,gy,-Math.PI/2,trunk,tr,26);const bases=[[gx,gy-250*tr],[gx,gy-370*tr],[gx,gy-trunk*tr]],ends=treeEnds();
    const order=n===5?[0,1,2,3,4]:[0,1,2,3,4].slice(0,n);
    order.forEach((bi,i)=>{const [ang,len,b]=TREE.br[bi],a=1.6+i*.3,p=eo(seg(t,a,a+.7));if(!p)return;br(bases[b][0],bases[b][1],ang,len,p,12);
      const [ex,ey]=ends[bi];for(let k=0;k<7;k++){const q=back(seg(t,a+.4+k*.05,a+.8+k*.05));if(q<=0)continue;x.save();x.fillStyle=k%2?C.lo:C.hi;
        x.translate(ex+Math.cos(k*1.7)*40,ey+Math.sin(k*1.7)*30);x.rotate(k);x.scale(q,q);x.beginPath();x.ellipse(0,0,22,11,0,0,7);x.fill();x.restore()}
      const q=eo(seg(t,a+.7,a+1.1));if(q>0){const [bx,by,bw]=bubble(bi,ex,ey);panel(bx-bw/2,by-34,bw,68,{r:34,alpha:q})}})},
  slots(f){const it=f.items||[],n=clamp(it.length,1,5),ends=treeEnds(),out=[headSlot(f,.2)];
    it.slice(0,n).forEach((o,i)=>{const [ex,ey]=ends[i],[bx,by,bw]=bubble(i,ex,ey);out.push({key:`items.${i}.title`,x:bx-bw/2+16,y:by-30,w:bw-32,h:60,role:'s',align:'c',valign:'m',at:1.6+i*.3+1,text:o.title||''})});return out}};
function bubble(bi,ex,ey){const ang=TREE.br[bi][0];const dx=bi===4?0:Math.cos(ang)>0?150:-150;return[ex+dx,ey-(bi===4?70:0),300]}

/* 19. Device Mockup */
S.device={n:19,name:{uz:'Qurilma',ru:'Устройство',en:'Device Mockup'},use:{uz:'ilova, sayt yoki video namoyishi',ru:'демо приложения или сайта',en:'app, website or video demo'},dur:3.8,
  fields:[{k:'heading',t:'text'},{k:'img',t:'img'},{k:'img2',t:'img'}],
  sample:{heading:'Loyiha'},
  draw(t,f,m){bg(t);headDeco(t,2.8,f.heading,m);const up=eo(seg(t,.2,1)),ly=lerp(1100,600,up),open=eio(seg(t,1,1.8)),on=seg(t,1.8,2.2);
    x.save();x.translate(700,ly);x.fillStyle=goldFill(0,30);rr(-440,0,880,30,10);x.fill();x.fillStyle=C.deep;x.fillRect(-80,4,160,8);
    const sh=430*open;x.fillStyle=C.dark;rr(-380,-sh,760,sh,14);x.fill();x.lineWidth=6;x.strokeStyle=C.gold;x.stroke();
    if(open>.6&&sh>40){x.save();rr(-360,-sh+20,720,sh-40,6);x.clip();x.fillStyle='#050912';x.fillRect(-360,-sh,720,sh);
      if(on){x.globalAlpha=on;if(IMG[f.img])cover(IMG[f.img],-360,-sh+20,720,sh-40);else picture(-360,-sh+20,720,sh-40,0)}x.restore()}x.restore();
    const pp=back(seg(t,2.2,3));if(pp>0){x.save();x.translate(lerp(1800,1300,pp),600);x.rotate(.08*(1-pp));x.fillStyle=C.dark;rr(-115,-400,230,420,28);x.fill();x.lineWidth=5;x.strokeStyle=C.gold;x.stroke();
      x.save();rr(-100,-380,200,380,18);x.clip();if(IMG[f.img2])cover(IMG[f.img2],-100,-380,200,380);else picture(-100,-380,200,380,3);x.restore();x.restore()}},
  slots(f){return[headSlot(f,3)]}};

/* 20. Quiz Cards */
S.quiz={n:20,name:{uz:'Savol kartalari',ru:'Карточки-викторина',en:'Quiz Cards'},use:{uz:'tinglovchilar uchun savollar',ru:'вопросы для аудитории',en:'audience quiz'},dur:5,
  fields:[{k:'heading',t:'text'},{k:'questions',t:'list',min:1,max:3,of:[{k:'q',t:'area'},{k:'a',t:'text'},{k:'b',t:'text'},{k:'c',t:'text'},{k:'ok',t:'choice',opts:['A','B','C']}]}],
  sample:{heading:'Savol',questions:[1,2,3].map((n,i)=>({q:`${n}-savol matni shu yerda?`,a:'Javob',b:'Javob',c:'Javob',ok:['A','B','C'][i]}))},
  draw(t,f,m){bg(t);headDeco(t,0,f.heading,m);const qs=(f.questions||[]).slice(0,3),n=qs.length||1;
    qs.forEach((o,i)=>{const cx=800+(i-(n-1)/2)*400,a=.3+i*.2,inn=back(seg(t,a,a+.6));if(!seg(t,a,a+.6))return;
      const fs=1.6+i*.8,flip=eio(seg(t,fs,fs+.6)),sx=Math.abs(Math.cos(flip*Math.PI)),front=flip>.5;
      x.save();x.translate(cx,lerp(1200,500,inn));x.scale(Math.max(sx,.02),1);panel(-170,-270,340,540,{});
      if(!front){x.save();x.strokeStyle=rgba(C.gold,.4);x.lineWidth=2;for(let k=0;k<6;k++){rr(-140+k*6,-240+k*6,280-k*12,480-k*12,10);x.stroke()}
        x.fillStyle=goldFill(-80,80);x.font=`700 200px ${F.head}`;x.textAlign='center';x.textBaseline='middle';x.fillText('?',0,10);x.restore()}
      else{const ok='ABC'.indexOf(o.ok||'A'),rv=seg(t,fs+1,fs+1.3);
        [0,1,2].forEach(k=>{const good=k===ok&&rv>0;rr(-145,-40+k*82,290,62,31);x.fillStyle=good?rgba(C.gold,.25+.6*rv):'rgba(255,255,255,.06)';x.fill();x.lineWidth=2;x.strokeStyle=good?C.hi:rgba(C.gold,.5);x.stroke()})}
      x.restore()})},
  slots(f){const qs=(f.questions||[]).slice(0,3),n=qs.length||1,out=[headSlot(f,.2)];
    qs.forEach((o,i)=>{const cx=800+(i-(n-1)/2)*400,at=1.6+i*.8+.6,Y=500;
      out.push({key:`questions.${i}.q`,x:cx-145,y:Y-245,w:290,h:180,role:'s',align:'c',valign:'m',at,text:o.q||''});
      ['a','b','c'].forEach((k,j)=>out.push({key:`questions.${i}.${k}`,x:cx-130,y:Y-40+j*82,w:260,h:62,role:'s',align:'c',valign:'m',at:at+.2+j*.1,text:`${'ABC'[j]})  ${o[k]||''}`}))});return out}};

/* 21. Glassmorphism */
function glassBg(t){x.fillStyle=C.bg3;x.fillRect(0,0,W,H);x.save();x.filter='blur(60px)';
  const cols=[C.gold,mix(C.bg,'#5a7bd8',.6),mix(C.gold,'#b0508a',.6),mix(C.bg,'#2fa3a0',.6),mix(C.gold,'#e08a3a',.4)];
  [[300,250,260],[1250,220,300],[900,700,320],[200,750,220],[1400,720,240]].forEach(([bx,by,r],i)=>{x.fillStyle=cols[i];x.globalAlpha=.8;x.beginPath();x.arc(bx+Math.sin(t*.6+i)*80,by+Math.cos(t*.5+i*2)*60,r,0,7);x.fill()});x.restore();
  if(IMG.bg){x.save();x.globalAlpha=IMG.vis;cover(IMG.bg,0,0,W,H);x.restore()}}
function glassCard(X,Y,w,h,t,alpha=1){x.save();x.globalAlpha=alpha;rr(X,Y,w,h,28);x.clip();x.filter='blur(24px)';glassBg(t);x.filter='none';
  x.fillStyle='rgba(255,255,255,.12)';x.fillRect(X,Y,w,h);const sh=x.createLinearGradient(X,Y,X+w,Y+h);sh.addColorStop(0,'rgba(255,255,255,.22)');sh.addColorStop(.5,'rgba(255,255,255,0)');x.fillStyle=sh;x.fillRect(X,Y,w,h);x.restore();
  x.save();x.globalAlpha=alpha;rr(X,Y,w,h,28);x.lineWidth=2;x.strokeStyle='rgba(255,255,255,.45)';x.stroke();x.restore()}
S.glass={n:21,name:{uz:'Shisha (glassmorphism)',ru:'Стекло (glassmorphism)',en:'Glassmorphism'},use:{uz:'matn va raqamlar shisha panellarda',ru:'текст и цифры на стеклянных панелях',en:'text and numbers on glass panels'},dur:2.8,
  fields:[{k:'heading',t:'text'},{k:'text',t:'area'},{k:'stats',t:'list',min:0,max:3,of:[{k:'value',t:'text'},{k:'label',t:'text'}]}],
  sample:{heading:'Sarlavha',text:'Asosiy matn shu yerda. Shisha panel ustida o\'qilishi oson.',stats:[{value:'8,2 mln',label:'Ko\'rsatkich'},{value:'24%',label:'Ko\'rsatkich'},{value:'120+',label:'Ko\'rsatkich'}]},
  draw(t,f){glassBg(t);const st=(f.stats||[]).slice(0,3),wide=!st.length;const a=back(seg(t,.3,1.1));
    glassCard(160,lerp(900,140,a),wide?1280:760,620,t,clamp(seg(t,.3,.7)));
    st.forEach((o,i)=>{const s0=1+i*.25,p=back(seg(t,s0,s0+.8));if(p<=0)return;glassCard(lerp(1700,1000,p),140+i*215,440,180,t)})},
  slots(f){const st=(f.stats||[]).slice(0,3),wide=!st.length,out=[{key:'heading',x:220,y:190,w:wide?1160:640,h:110,role:'h',align:'l',valign:'b',at:1,color:'#ffffff',text:f.heading||''},
    {key:'text',x:220,y:330,w:wide?1160:640,h:400,role:'b',align:'l',at:1.3,color:'#ffffff',text:f.text||''}];
    st.forEach((o,i)=>{out.push({key:`stats.${i}.value`,x:1040,y:165+i*215,w:360,h:80,role:'n',align:'l',valign:'m',at:1.6+i*.25,color:'#ffffff',text:o.value||''});
      out.push({key:`stats.${i}.label`,x:1040,y:248+i*215,w:360,h:50,role:'s',align:'l',at:1.75+i*.25,color:'#ffffff',text:o.label||''})});return out}};

/* 22. Kinetic Type: the words are part of the animation */
function kineticLayout(words,m){const out=[];let size=118*m;const maxW=1340;
  for(;size>40;size-=6){const lines=[[]];let w=0;x.font=`800 ${size}px ${F.head}`;const sp=size*.3;
    words.forEach(wd=>{const ww=x.measureText(wd).width;if(w&&w+sp+ww>maxW){lines.push([]);w=0}lines[lines.length-1].push([wd,ww]);w+=(w?sp:0)+ww});
    if(lines.length*size*1.18<=560){const top=450-lines.length*size*1.18/2+size*.85;
      lines.forEach((ln,li)=>{const tw=ln.reduce((s,[,ww])=>s+ww,0)+sp*(ln.length-1);let X=800-tw/2;ln.forEach(([wd,ww])=>{out.push({w:wd,x:X+ww/2,y:top+li*size*1.18,size});X+=ww+sp})});return out}}
  return out}
S.kinetic={n:22,name:{uz:'Kinetik matn',ru:'Кинетическая типографика',en:'Kinetic Type'},use:{uz:'shior yoki qisqa jumla',ru:'слоган или короткая фраза',en:'slogan or short sentence'},dur:5,textInVideo:true,
  fields:[{k:'sentence',t:'text'}],sample:{sentence:'Biz yangi g\'oyalarni taqdim etamiz'},
  draw(t,f,m){bg(t);const words=(f.sentence||'').trim().split(/\s+/).filter(Boolean).slice(0,10).map(w=>w.toUpperCase());if(!words.length)return;
    const L=kineticLayout(words,m),per=Math.min(.7,3/words.length),settle=eio(seg(t,per*words.length+.3,per*words.length+1.1)),kinds=['scale','left','rot','up','right'];
    L.forEach((o,i)=>{const a=i*per,inn=seg(t,a,a+.45);if(!inn)return;const kind=kinds[i%5],e=back(inn),big=Math.min(o.size*1.7,1500/Math.max(4,o.w.length)*1.6);
      let X=800,Y=480,S=big,R=0,A=1;if(kind==='scale'){S=lerp(big*3,big,eo(inn));A=inn}if(kind==='left')X=lerp(-700,800,e);if(kind==='right')X=lerp(2300,800,e);
      if(kind==='up')Y=lerp(-200,480,e);if(kind==='rot'){R=lerp(-1.6,0,e);S=lerp(10,big,eo(inn))}
      const next=i<L.length-1?seg(t,(i+1)*per,(i+1)*per+.25):0;X=lerp(X,o.x,settle);Y=lerp(Y,o.y,settle);S=lerp(S,o.size,settle);
      x.save();x.translate(X,Y);x.rotate(R);x.globalAlpha=A*(settle>0?1:1-next);x.font=`800 ${S}px ${F.head}`;x.textAlign='center';x.textBaseline='alphabetic';
      x.shadowColor='rgba(0,0,0,.7)';x.shadowBlur=14;x.shadowOffsetY=6;x.fillStyle=i%2?C.text:goldFill(-S*.8,0);x.fillText(o.w,0,S*.35-S*.35);x.restore()});
    const u=eo(seg(t,per*words.length+1.1,per*words.length+1.7));const lastY=L.length?L[L.length-1].y:600;x.fillStyle=C.gold;x.fillRect(800-300*u,lastY+40,600*u,6)},
  durFor(f){const n=((f.sentence||'').trim().split(/\s+/).filter(Boolean).length)||1;return Math.min(.7,3/n)*n+1.9},
  slots(){return[]}};

/* 23. Split-flap Display */
/* tiles flip through letters of the same alphabet as the letter they land on */
const FLAP_LAT=" ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!?.,-'",FLAP_CYR=" АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯЎҚҒҲ0123456789!?.,-";
S.flap={n:23,name:{uz:'Aeroport tablosi',ru:'Табло (split-flap)',en:'Split-flap Display'},use:{uz:'sarlavha, sana, shior',ru:'заголовок, дата, слоган',en:'title, date, slogan'},dur:4.2,textInVideo:true,
  fields:[{k:'lines',t:'list',min:1,max:3,of:[{k:'text',t:'text',max:13}]}],sample:{lines:[{text:'TAQDIMOT'},{text:'2026'},{text:'BOSHLANDI!'}]},
  draw(t,f){bg(t);const rows=(f.lines||[]).slice(0,3).map(o=>(o.text||'').toUpperCase().slice(0,13));if(!rows.length)return;
    const cols=13,cw=96,ch=150,gap=12,X0=800-(cols*cw+(cols-1)*gap)/2,Y0=450-rows.length*(ch+30)/2+15;
    x.save();rr(X0-30,Y0-30,cols*(cw+gap)-gap+60,rows.length*(ch+30)+30,20);x.fillStyle=C.bg3;x.fill();x.strokeStyle=C.lo;x.lineWidth=2;x.stroke();x.restore();
    rows.forEach((line,r)=>{const pad=Math.floor((cols-line.length)/2),full=' '.repeat(pad)+line+' '.repeat(cols-pad-line.length);
      const cyr=/[\u0400-\u04FF]/.test(line),base=cyr?FLAP_CYR:FLAP_LAT;
      [...full].forEach((target,c)=>{let ti=base.indexOf(target);const set=ti<0?base+target:base;if(ti<0)ti=set.length-1;
        const start=.3+r*.3+c*.05,steps=8+((c*7+r*3)%9),per=.07,k=Math.max(0,Math.min(steps,Math.floor((t-start)/per)));
        const cur=k>=steps?ti:(ti-steps+k+set.length*4)%set.length,prev=(cur-1+set.length)%set.length,flip=t<start?1:k>=steps?clamp((t-start-steps*per)/per):((t-start)/per)%1;
        flapCell(X0+c*(cw+gap),Y0+r*(ch+30),cw,ch,set[cur],set[prev],flip)})})},
  slots(){return[]}};
function flapCell(X,Y,w,h,ch,prev,flip){
  const draw=(c,half)=>{x.save();x.beginPath();half==='top'?x.rect(X,Y,w,h/2):x.rect(X,Y+h/2,w,h/2);x.clip();rr(X,Y,w,h,8);x.fillStyle=C.dark;x.fill();
    x.fillStyle=C.hi;x.font=`700 ${h*.66}px ${F.head}`;x.textAlign='center';x.textBaseline='middle';x.fillText(c,X+w/2,Y+h/2+h*.04);x.restore()};
  draw(ch,'top');draw(flip<1?prev:ch,'bottom');
  if(flip<1){x.save();x.translate(0,Y+h/2);if(flip<.5){x.scale(1,1-flip*2);x.translate(0,-(Y+h/2));draw(prev,'top')}else{x.scale(1,(flip-.5)*2);x.translate(0,-(Y+h/2));draw(ch,'bottom')}x.restore()}
  x.fillStyle='rgba(0,0,0,.8)';x.fillRect(X,Y+h/2-1.5,w,3)}

/* ================= text: sizes, wrapping, drawing ================= */
const ROLE={h:{px:64,font:'head',weight:700,spc:4},t:{px:32,font:'head',weight:700,spc:0},b:{px:26,font:'body',weight:400,spc:0},s:{px:21,font:'body',weight:400,spc:0},n:{px:66,font:'head',weight:700,spc:0}};
const LINE=1.22;
function wrap(text,font,size,weight,w,spc=0){x.save();x.font=`${weight} ${size}px ${font}`;const out=[];
  String(text).split('\n').forEach(par=>{const words=par.split(/\s+/).filter(Boolean);let line='';
    if(!words.length){out.push('');return}
    words.forEach(wd=>{const test=line?line+' '+wd:wd;if(line&&x.measureText(test).width+test.length*spc>w){out.push(line);line=wd}else line=test});out.push(line)});x.restore();return out}
/* the size a slot's text gets after shrinking to fit its box */
function fitSlot(s,mult){const R=ROLE[s.role]||ROLE.b,font=R.font==='head'?F.head:F.body;let size=(s.size||R.px)*mult;
  for(let i=0;i<30;i++){const lines=wrap(s.text,font,size,R.weight,s.w,R.spc);if(lines.length*size*LINE<=s.h+2||size<11)return{size,lines,font,weight:R.weight,spc:R.spc};size*=.93}
  return{size,lines:wrap(s.text,font,size,R.weight,s.w,R.spc),font,weight:R.weight,spc:R.spc}}
function slotColor(s){if(!s.color)return s.role==='h'?C.hi:s.role==='t'?C.hi:s.role==='n'?C.hi:C.text;if(s.color==='muted')return C.muted;if(s.color==='accent')return C.gold;return s.color}
const EFFECTS=['fade','rise','fly','wipe','split','zoom','typewriter','words'];
const SPEED={slow:.9,normal:.6,fast:.4};
/* draws a slot as it will look in PowerPoint, with an approximation of its entrance effect */
function drawSlot(s,t,mult,effect='fade',speed='normal'){if(!s.text)return;const d=SPEED[speed]||.6,p=seg(t,s.at,s.at+(effect==='typewriter'?Math.max(d,s.text.length*.035):effect==='words'?Math.max(d,s.text.split(/\s+/).length*.12):d));if(p<=0)return;
  const fit=fitSlot(s,mult),lh=fit.size*LINE,blockH=fit.lines.length*lh;let y0=s.y;if(s.valign==='m')y0=s.y+(s.h-blockH)/2;if(s.valign==='b')y0=s.y+s.h-blockH;
  x.save();x.translate(s.x+s.w/2,s.y+s.h/2);if(s.rot)x.rotate(s.rot*Math.PI/180);x.translate(-(s.x+s.w/2),-(s.y+s.h/2));
  let alpha=1,dy=0,sc=1;const e=eo(p);
  if(effect==='fade')alpha=p;if(effect==='rise'){alpha=p;dy=36*(1-e)}if(effect==='fly')dy=(H-s.y)*(1-e);if(effect==='zoom'){sc=lerp(.2,1,e);alpha=p}
  if(effect==='wipe'){x.beginPath();x.rect(s.x-10,s.y-20,(s.w+20)*e,s.h+40);x.clip()}
  if(effect==='split'){x.beginPath();x.rect(s.x+s.w/2-(s.w/2+10)*e,s.y-20,(s.w+20)*e,s.h+40);x.clip()}
  x.translate(s.x+s.w/2,s.y+s.h/2);x.scale(sc,sc);x.translate(-(s.x+s.w/2),-(s.y+s.h/2)+dy);x.globalAlpha=alpha;
  x.font=`${fit.weight} ${fit.size}px ${fit.font}`;if('letterSpacing' in x)x.letterSpacing=(fit.spc)+'px';x.fillStyle=slotColor(s);x.textBaseline='alphabetic';
  x.shadowColor=s.color&&s.color.startsWith('#2')?'transparent':'rgba(0,0,0,.6)';x.shadowBlur=10;x.shadowOffsetY=3;
  const total=fit.lines.join(' ').length;let shown=effect==='typewriter'?Math.floor(total*p):Infinity,wordsShown=effect==='words'?Math.ceil(fit.lines.join(' ').split(/\s+/).length*p):Infinity;
  fit.lines.forEach((ln,i)=>{let txt=ln;if(shown!==Infinity){txt=ln.slice(0,Math.max(0,shown));shown-=ln.length+1}
    if(wordsShown!==Infinity){const ws=ln.split(/\s+/);txt=ws.slice(0,Math.max(0,wordsShown)).join(' ');wordsShown-=ws.length}
    const lw=x.measureText(ln).width,X=s.align==='c'?s.x+(s.w-lw)/2:s.align==='r'?s.x+s.w-lw:s.x;x.textAlign='left';x.fillText(txt,X,y0+i*lh+fit.size*.92)});
  x.restore()}

/* ================= public API ================= */
const ORDER=['split','points','puzzle','gallery','domino','stats','balance','blinds','spotlight','tree','device','quiz','glass','kinetic','flap'];
function durOf(st,f){return st.durFor?st.durFor(f):st.dur}
/* the text boxes of a page, with its own values filled in */
function slotsOf(id,f){return (S[id].slots(f)||[]).filter(s=>String(s.text||'').trim())}
/* draws one frame: background animation, plus (optionally) the text as PowerPoint will show it */
function render(ctx,id,f,t,T,{text=true}={}){x=ctx;setTheme(T);const st=S[id],m=SIZES[T.size]||1;
  ctx.save();ctx.clearRect(0,0,W,H);st.draw(t,f,m);
  if(text)slotsOf(id,f).forEach(s=>drawSlot(s,t,m,T.effect,T.speed));ctx.restore()}
function fitFor(ctx,s,T){x=ctx;setTheme(T);return fitSlot(s,SIZES[T.size]||1)}
function colorFor(ctx,s,T){x=ctx;setTheme(T);return slotColor(s)}
window.IMXStyles={W,H,S,ORDER,THEMES,FONTS,SIZES,EFFECTS,SPEED,render,slotsOf,durOf,fitFor,colorFor,ROLE,LINE,mix};
})();
