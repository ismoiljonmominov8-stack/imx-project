/* IMX Morph templates (24-33).
   These pages are made of normal PowerPoint shapes. One IMX page becomes one or more slides;
   shapes with the same name on neighbouring slides are moved by PowerPoint's Morph transition.
   Each template returns slides as a list of elements (in 1600x900 design pixels) plus the
   animations that play inside the slide. The same description is used for the live preview
   (drawn on a canvas) and for the PPTX file. */
(function(){
'use strict';
const ST=window.IMXStyles,mix=ST.mix;
const clamp=(v,a=0,b=1)=>Math.max(a,Math.min(b,v));
const lerp=(a,b,k)=>a+(b-a)*k;
const eio=p=>p<.5?4*p*p*p:1-Math.pow(-2*p+2,3)/2, eo=p=>1-Math.pow(1-p,3);
const TILE=['#e8432f','#f08a24','#3cae4f','#13a39a','#2f6fe4','#8e44c9','#d6336c','#0ea5e9'];

/* ---------- element builders ---------- */
const box=(n,x,y,w,h,fill,o={})=>({n,t:o.r!=null?'round':'rect',x,y,w,h,fill,...o});
const oval=(n,x,y,w,h,fill,o={})=>({n,t:'oval',x,y,w,h,fill,...o});
const poly=(n,pts,fill,o={})=>{const xs=pts.map(p=>p[0]),ys=pts.map(p=>p[1]),x=Math.min(...xs),y=Math.min(...ys);
  return{n,t:'poly',x,y,w:Math.max(...xs)-x||1,h:Math.max(...ys)-y||1,pts:pts.map(([a,b])=>[a-x,b-y]),fill,...o}};
const text=(n,x,y,w,h,tx,sz,col,o={})=>({n,t:'text',x,y,w,h,tx:String(tx??''),sz,col,font:o.b?'head':'body',...o});
const pic=(n,img,x,y,w,h,o={})=>({n,t:'pic',img,x,y,w,h,...o});
function diamond(cx,cy,size,outer,inner,n='logo'){return[box(n,cx-size/2,cy-size/2,size,size,outer,{r:.12,rot:45}),
  poly(n+'_in',[[cx,cy-size*.36],[cx+size*.16,cy],[cx,cy+size*.36],[cx-size*.16,cy]],inner)]}
function logoEls(P,cx,cy,size){return P.logo?[pic('logo',P.logo,cx-size/2,cy-size/2,size,size,{r:.2})]:diamond(cx,cy,size,P.c.mint,P.c.inner)}

/* ---------- palettes from the deck theme ---------- */
function palette(T){const bg=T.bg,ac=T.accent,tx=T.text;
  return{
    dark:{bg,bg2:mix(bg,'#ffffff',.04),panel:mix(bg,ac,.1),mint:ac,hi:mix(ac,'#ffffff',.45),inner:mix(ac,'#ffffff',.85),
      text:tx,white:'#ffffff',muted:mix(tx,bg,.42),line:mix(bg,ac,.28),glass:mix(bg,ac,.22)},
    light:{frame:bg,bg:'#f4f6fb',card:'#ffffff',navy:mix(ac,'#000000',.05),ink:'#1b2340',muted:'#7a8399',line:'#d9dee9',soft:'#e9edf6',onNavy:mix(ac,'#ffffff',.7)},
    photo:{bg:mix(bg,'#000000',.45),text:tx,muted:mix(tx,bg,.35)}}}

/* ---------- generated placeholder pictures ---------- */
const GEN={};
function genImage(key,accent){const id=key+'|'+accent;if(GEN[id])return GEN[id];
  const c=document.createElement('canvas'),[kind,arg]=key.split(':'),x=c.getContext('2d');
  if(kind==='ribbon'){c.width=900;c.height=1200;const ph=(+arg||0)*1.3;
    x.fillStyle=mix(accent,'#000000',.92);x.fillRect(0,0,900,1200);x.filter='blur(5px)';
    for(let k=0;k<60;k++){const t=k/60;x.beginPath();for(let i=0;i<=120;i++){const yy=-.2+1.4*i/120,xx=(.5+.33*Math.sin(yy*5.2+ph+t*1.9)+.12*Math.sin(yy*11+ph*2+t*4)+(t-.5)*.25)*900;i?x.lineTo(xx,yy*1200):x.moveTo(xx,yy*1200)}
      const b=.35+.65*(.5+.5*Math.sin(t*7+ph));x.strokeStyle=mix(mix(accent,'#000000',.6),mix(accent,'#ffffff',.25),b);x.globalAlpha=.1+.18*t;x.lineWidth=14*(.4+t);x.stroke()}
    x.filter='none';x.globalAlpha=.25;x.strokeStyle=mix(accent,'#ffffff',.7);x.lineWidth=2;
    for(let k=0;k<14;k++){const t=k/14;x.beginPath();for(let i=0;i<=120;i++){const yy=-.2+1.4*i/120,xx=(.5+.33*Math.sin(yy*5.2+ph+t*1.9+.05)+.12*Math.sin(yy*11+ph*2+t*4)+(t-.5)*.25)*900;i?x.lineTo(xx,yy*1200):x.moveTo(xx,yy*1200)}x.stroke()}
    x.globalAlpha=1;const g=x.createRadialGradient(450,600,200,450,600,900);g.addColorStop(0,'rgba(0,0,0,0)');g.addColorStop(1,'rgba(0,0,0,.6)');x.fillStyle=g;x.fillRect(0,0,900,1200)}
  else{/* a neutral "your photo here" placeholder */
    c.width=1600;c.height=900;const warm=kind==='soft'&&(+arg%2);const g=x.createLinearGradient(0,0,1600,900);
    g.addColorStop(0,warm?'#6b4a33':mix(accent,'#203040',.55));g.addColorStop(1,warm?'#c9a27a':mix(accent,'#c8d6d0',.5));x.fillStyle=g;x.fillRect(0,0,1600,900);
    x.globalAlpha=.18;x.fillStyle='#ffffff';for(let k=0;k<14;k++){x.beginPath();x.arc((k*397)%1600,(k*233)%900,40+(k*37)%120,0,7);x.fill()}x.globalAlpha=.75;
    const s=110,cx=800,cy=kind==='photo'?270:450;x.strokeStyle='#ffffff';x.lineWidth=8;x.lineJoin='round';
    x.beginPath();x.roundRect(cx-s,cy-s*.75,s*2,s*1.5,18);x.stroke();x.beginPath();x.moveTo(cx-s*.8,cy+s*.5);x.lineTo(cx-s*.25,cy-s*.1);x.lineTo(cx+s*.1,cy+s*.25);x.lineTo(cx+s*.4,cy);x.lineTo(cx+s*.8,cy+s*.5);x.stroke();
    x.beginPath();x.arc(cx+s*.45,cy-s*.35,s*.15,0,7);x.stroke();x.globalAlpha=1}
  return GEN[id]=c}

/* ================= the templates ================= */
const M={};
const SAY=(uz,ru,en)=>({uz,ru,en});

/* 24 Glass Title Card */
M.glassTitle={n:24,group:'dark',name:SAY('Shisha sarlavha kartasi','Стеклянная титульная карточка','Glass Title Card'),use:SAY('kompaniya yoki mavzu nomi','название компании или темы','company or topic name'),
  fields:[{k:'title',t:'text'},{k:'subtitle',t:'text'},{k:'logo',t:'img'}],sample:{title:'Kompaniya nomi',subtitle:'KOMPANIYA PROFILI'},
  build(f,P){const c=P.c,L=f.logo;P.logo=L;
    const a={bg:'#ffffff',els:[box('plate',0,0,1600,900,'#ffffff'),...logoEls(P,800,450,46)]};
    const b={bg:c.bg,els:[box('g1',120,250,1360,400,c.mint,{a:.08,line:c.mint,lw:1}),box('g2',300,190,1000,520,c.mint,{a:.1,line:c.mint,lw:1}),
      box('plate',520,230,560,440,c.glass,{a:.55,line:c.hi,lw:1.25}),...logoEls(P,800,340,92),
      text('title',520,430,560,70,f.title,34,c.text,{b:1,al:'c',va:'m'}),text('sub',520,505,560,30,f.subtitle,11,c.muted,{al:'c',sp:3})],
      anims:[{n:'g1',k:'fade',d:0,u:.6},{n:'g2',k:'fade',d:.15,u:.6},{n:'title',k:'rise',d:.35,u:.6},{n:'sub',k:'fade',d:.6,u:.5}]};
    return[a,b]}};

/* 26 Section Divider */
M.divider={n:26,group:'dark',name:SAY('Bo\'lim ajratgichi','Разделитель раздела','Section Divider'),use:SAY('katta rasm va bo\'lim nomi','большое фото и название раздела','big picture and section name'),
  fields:[{k:'title',t:'text'},{k:'img',t:'img'},{k:'left',t:'area'},{k:'right',t:'area'}],sample:{title:'Mundarija',left:'Kompaniya nomi\nKompaniya profili',right:'Ism Familiya\nLavozim'},
  build(f,P){const c=P.c;return[{bg:c.bg,els:[pic('img',f.img||'ribbon:0',0,0,1000,900),box('panel',1000,0,600,900,c.bg2),
    text('h',90,330,820,140,f.title,54,'#ffffff',{b:1,va:'m'}),box('rule',1060,470,480,1.5,c.line),
    text('co',1060,420,280,44,f.left,12,c.text),text('who',1360,420,200,44,f.right,12,c.text),...logoEls(P,1000,830,70)],
    anims:[{n:'h',k:'rise',d:0,u:.7},{n:'co',k:'fade',d:.4,u:.5},{n:'who',k:'fade',d:.5,u:.5}]}]}};

/* 25 Accordion List */
M.accordion={n:25,focus:i=>i,group:'dark',name:SAY('Ochiluvchi ro\'yxat','Раскрывающийся список','Accordion List'),use:SAY('bandlar birma-bir ochiladi, rasm va raqam o\'zgaradi','пункты раскрываются по очереди','items open one at a time'),
  fields:[{k:'heading',t:'text'},{k:'items',t:'list',min:2,max:6,of:[{k:'title',t:'text'},{k:'text',t:'area'},{k:'value',t:'text'},{k:'label',t:'text'},{k:'img',t:'img'}]}],
  sample:{heading:'Mundarija',items:[{title:'Kompaniya haqida',text:'Biz kimmiz va qanday ishlaymiz — qisqa tanishtiruv.',value:'12 yil',label:'bozorda tajriba'},
    {title:'Mijoz vazifasi',text:'Daromad bozordan sekin o\'smoqda, qarorlar esa ma\'lumotga emas, his-tuyg\'uga asoslanadi.',value:'3 / 4',label:'mijozlar shu bosqichdan o\'tadi'},
    {title:'Bizning yondashuv',text:'3 haftalik diagnostika, bir yillik yo\'l xaritasi va uni birga joriy qiladigan jamoa.',value:'3 hafta',label:'to\'liq diagnostikaga'},
    {title:'Natija',text:'Aniq strategiya, o\'lchanadigan o\'sish va jamoaning yangi ko\'nikmalari.',value:'+23%',label:'o\'rtacha o\'sish'}]},
  build(f,P){const c=P.c,it=f.items||[],n=it.length,slides=[],rowH=n>4?58:70,openH=n>4?140:160;
    for(let k=0;k<n;k++){const els=[box('panel',0,0,1000,900,c.bg),text('h',90,70,600,60,f.heading,28,c.text)];let y=175;
      it.forEach((o,i)=>{const open=i===k;if(open)els.push(box('accent',90,y+4,4,openH-10,c.mint));
        els.push(text('n'+i,112,y+14,50,30,String(i+1).padStart(2,'0'),12,c.muted),text('t'+i,170,y+10,640,36,o.title,17,'#ffffff',{b:1}),text('x'+i,880,y+8,40,36,open?'×':'+',20,c.hi,{al:'r'}));
        if(open){els.push(text('desc',170,y+56,650,openH-70,o.text,13,c.muted,{lh:1.15}));y+=openH}else y+=rowH;
        els.push(box('l'+i,112,y-8,808,1.2,c.line))});
      els.push(text('pg',90,820,200,24,`${String(k+1).padStart(2,'0')} / ${String(n).padStart(2,'0')}`,11,c.muted),box('track',90,850,830,3,c.line),box('prog',90,850,830*(k+1)/n,3,c.mint),
        pic('img',it[k].img||('ribbon:'+(k%5)),1000,0,600,900),box('badge',1040,60,300,120,mix(c.bg,'#000000',.1),{a:.72,line:c.mint,lw:1}),
        text('bv',1065,72,260,60,it[k].value,30,'#ffffff',{b:1}),text('bl',1065,136,260,30,it[k].label,11,c.muted),...logoEls(P,1500,800,70));
      slides.push({bg:c.bg,els,anims:[{n:'desc',k:'fade',d:.1,u:.5},{n:'badge',k:'pulse',d:.3,u:.3}]})}
    return slides}};

/* 27 Contact / Closing Card */
M.contact={n:27,group:'dark',name:SAY('Aloqa kartasi','Контакты / финал','Contact / Closing Card'),use:SAY('yakuniy sahifa va aloqa','финальная страница с контактами','closing page with contacts'),
  fields:[{k:'chip',t:'text'},{k:'line1',t:'text'},{k:'line2',t:'text'},{k:'contact1',t:'area'},{k:'contact2',t:'area'},{k:'left',t:'area'},{k:'right',t:'area'},{k:'img',t:'img'}],
  sample:{chip:'ALOQA',line1:'O\'sishingizni muhokama qilamiz.',line2:'Birinchi uchrashuv — biz tomondan.',contact1:'info@kompaniya.uz\n+998 90 000 00 00',contact2:'kompaniya.uz\nToshkent',left:'Kompaniya nomi\nKompaniya profili',right:'Ism Familiya\nLavozim'},
  build(f,P){const c=P.c;return[{bg:c.bg,els:[pic('img',f.img||'ribbon:4',0,0,620,900),box('panel',620,0,980,900,c.bg),box('card',700,90,820,470,c.panel,{line:c.line,lw:1}),
    box('chip',750,135,140,34,null,{r:.5,line:c.mint,lw:1.5}),text('chipt',750,135,140,34,f.chip,10,c.hi,{al:'c',va:'m',sp:2}),
    text('h',750,190,740,70,f.line1,23,'#ffffff',{b:1,va:'b'}),text('h2',750,265,740,70,f.line2,23,c.hi,{b:1}),box('rule',750,350,720,1.2,c.line),
    text('c1',750,380,340,70,f.contact1,13,c.text),text('c2',1110,380,340,70,f.contact2,13,c.muted),box('card2',700,620,820,170,c.panel,{line:c.line,lw:1}),
    text('co',760,665,340,80,f.left,14,c.text),text('who',1130,665,340,80,f.right,14,c.text),...logoEls(P,620,820,90)],
    anims:[{n:'h',k:'rise',d:0,u:.6},{n:'h2',k:'rise',d:.2,u:.6},{n:'c1',k:'fade',d:.5,u:.5},{n:'c2',k:'fade',d:.6,u:.5}]}]}};

/* light frame shared by 28-31 */
function sheet(c){return box('sheet',40,40,1520,820,c.bg,{r:.03})}
function badgeNum(c,cx,cy,size,num,word){return[box('badge',cx-size/2,cy-size/2,size,size,c.navy,{r:.28}),
  text('bsub',cx-size/2,cy-size*.36,size,size*.16,word,Math.max(7,size*.07),c.onNavy,{al:'c',va:'m',sp:1}),
  text('bn',cx-size/2,cy-size*.25,size,size*.6,num,Math.max(1,size*.42),'#ffffff',{b:1,al:'c',va:'m'})]}

/* 28 Steps Hub */
M.stepsHub={n:28,focus:(i,n)=>n-1,group:'light',name:SAY('Qadamlar markazi','Центр шагов','Steps Hub'),use:SAY('3–8 qadam rangli plitkalarda','3–8 шагов цветными плитками','3–8 steps as colour tiles'),
  fields:[{k:'title',t:'text'},{k:'word',t:'text'},{k:'steps',t:'list',min:3,max:8,of:[{k:'name',t:'text'}]}],
  sample:{title:'Taqdimot yaratishning 6 qadami',word:'QADAM',steps:['G\'oya','Tuzilma','Dizayn','Vizual','Mashq','Chiqish'].map(name=>({name}))},
  build(f,P){const c=P.c,st=f.steps||[],n=st.length,num=String(n),slides=[];
    slides.push({bg:c.frame,els:[sheet(c),...badgeNum(c,800,450,40,num,f.word)]});
    let els=[sheet(c),...badgeNum(c,800,450,230,num,f.word)],an=[];
    for(let i=0;i<n;i++){const a=(-90+360*i/n)*Math.PI/180,x=800+330*Math.cos(a)-45,y=450+300*Math.sin(a)-40;
      els.push(box('tile'+i,x,y,90,80,'#ffffff',{r:.26,line:c.line,lw:1.5}),text('tl'+i,x,y,90,80,String(i+1).padStart(2,'0'),14,c.muted,{b:1,al:'c',va:'m'}));
      an.push({n:'tile'+i,k:'zoom',d:.15*i,u:.35},{n:'tl'+i,k:'fade',d:.15*i+.1,u:.3})}
    an.push({n:'badge',k:'pulse',d:.15*n+.1,u:.35});slides.push({bg:c.frame,els,anims:an});
    els=[sheet(c),...badgeNum(c,800,450,200,num,f.word)];
    for(let i=0;i<n;i++){const a=(-90+360*(i+.5)/n)*Math.PI/180,x=800+300*Math.cos(a)-60,y=450+280*Math.sin(a)-55;
      els.push(box('tile'+i,x,y,120,110,TILE[i],{r:.26,rot:i%2?12:-12}),text('tl'+i,x,y,120,110,String(i+1).padStart(2,'0'),20,'#ffffff',{b:1,al:'c',va:'m'}))}
    slides.push({bg:c.frame,els,dur:1.4});
    els=[sheet(c),text('title',0,110,1600,40,f.title,16,c.muted,{al:'c'}),...badgeNum(c,800,132,.1,num,f.word)];an=[];
    const w=Math.min(200,(1400-(n-1)*24)/n),x0=800-(n*w+(n-1)*24)/2;
    st.forEach((o,i)=>{const x=x0+i*(w+24),y=260+(i%2?60:0),h=w*.9;
      els.push(box('tile'+i,x,y,w,h,TILE[i],{r:.26}),text('tl'+i,x,y+h*.16,w,h*.28,String(i+1).padStart(2,'0'),12,'#ffffff',{b:1,al:'c',va:'m'}),
        text('tn'+i,x+8,y+h*.42,w-16,h*.46,o.name,Math.max(13,Math.min(20,w/9)),'#ffffff',{b:1,al:'c',va:'m'}),
        oval('ic'+i,x+w/2-26,y+h+40,52,52,'#ffffff',{line:c.line,lw:1.5}),oval('icd'+i,x+w/2-9,y+h+57,18,18,TILE[i]));
      an.push({n:'tn'+i,k:'rise',d:.1*i,u:.4},{n:'ic'+i,k:'zoom',d:.5+.1*i,u:.3},{n:'icd'+i,k:'zoom',d:.6+.1*i,u:.3},{n:'tile'+i,k:'pulse',d:1.4+.25*i,u:.25})});
    slides.push({bg:c.frame,els,anims:an,dur:1.4});return slides}};

/* 29 Step Photo Card */
M.stepPhoto={n:29,focus:i=>i,group:'light',name:SAY('Qadam rasm kartasi','Шаг с фото','Step Photo Card'),use:SAY('har bir qadam: rasm, matn, maslahat','каждый шаг: фото, текст, совет','each step: photo, text, tip'),
  fields:[{k:'word',t:'text'},{k:'tipWord',t:'text'},{k:'steps',t:'list',min:1,max:8,of:[{k:'title',t:'text'},{k:'text',t:'area'},{k:'tip',t:'area'},{k:'img',t:'img'}]}],
  sample:{word:'QADAM',tipWord:'MASLAHAT',steps:[{title:'Mashq',text:'Chiqishni ovoz chiqarib uch marta takrorlang: taymer bilan, kameraga va do\'stlar oldida.',tip:'Har bir slaydga bitta asosiy fikr — ortiqchasini olib tashlang.'},
    {title:'Chiqish',text:'Tinglovchilar bilan ko\'z aloqasini saqlang va asosiy fikrni oxirida yana takrorlang.',tip:'Oxirgi slayd — «rahmat» emas, keyingi qadam bo\'lsin.'}]},
  build(f,P){const c=P.c,st=f.steps||[],n=st.length;return st.map((o,k)=>{const els=[sheet(c),box('frame',130,150,560,600,'#ffffff',{r:.16,line:c.line,lw:1.5}),
      pic('photo',o.img||('soft:'+k),150,170,520,560,{r:.15}),oval('mic',620,140,70,70,'#ffffff',{line:c.line,lw:1.5}),
      text('micn',620,140,70,70,String(k+1).padStart(2,'0'),14,c.navy,{b:1,al:'c',va:'m'}),text('step',790,220,500,30,`${f.word} ${String(k+1).padStart(2,'0')}`,11,c.muted,{sp:2}),
      text('h',790,255,680,70,o.title,36,c.ink,{b:1}),text('body',790,340,620,130,o.text,15,c.muted,{lh:1.2}),box('tip',790,480,560,130,c.soft,{r:.12}),
      text('tiph',820,500,500,26,f.tipWord,10,c.navy,{b:1,sp:2}),text('tipt',820,530,500,70,o.tip,13,c.ink)];
    const dw=Math.min(48,(560-(n-1)*14)/n);
    for(let i=0;i<n;i++){const on=i===k;els.push(box('dot'+i,790+i*(dw+14),680,dw,48,on?TILE[i]:'#ffffff',{r:.3,line:on?null:c.line,lw:1.5}),
      text('dotn'+i,790+i*(dw+14),680,dw,48,String(i+1),12,on?'#ffffff':c.muted,{al:'c',va:'m'}))}
    return{bg:c.frame,els,anims:[{n:'body',k:'fade',d:.1,u:.5},{n:'tip',k:'rise',d:.3,u:.5},{n:'tiph',k:'fade',d:.4,u:.4},{n:'tipt',k:'fade',d:.45,u:.4}]}})}};

/* 30 Wheel */
function wheelEls(c,parts,word,cx,cy,R,{explode=0,pull=null,dim=false}={}){const n=parts.length,els=[],pts=[];
  for(let i=0;i<n;i++){const a=(-90+360*i/n)*Math.PI/180;pts.push([cx+R*Math.cos(a),cy+R*Math.sin(a)])}
  for(let i=0;i<n;i++){const A=pts[i],B=pts[(i+1)%n],mid=(-90+360*(i+.5)/n)*Math.PI/180,off=explode+(pull===i?36:0),dx=off*Math.cos(mid),dy=off*Math.sin(mid);
    const pp=[[cx+dx,cy+dy],[A[0]+dx,A[1]+dy],[B[0]+dx,B[1]+dy]];els.push(poly('seg'+i,pp,TILE[i],{a:dim&&pull!=null&&pull!==i?.35:1,rot:explode?(i%2?30:-30):0}));
    const mx=(pp[0][0]+pp[1][0]+pp[2][0])/3,my=(pp[0][1]+pp[1][1]+pp[2][1])/3;
    els.push(text('sn'+i,mx-R*.3,my-18,R*.6,36,explode?'':parts[i].name,Math.max(11,Math.min(16,R/19)),'#ffffff',{b:1,al:'c',va:'m'}))}
  const r=R*.29;els.push(oval('badge',cx-r,cy-r,2*r,2*r,c.navy,{line:'#ffffff',lw:6}),text('bsub',cx-r,cy-r*.62,2*r,r*.3,word,Math.max(7,r/9),c.onNavy,{al:'c',va:'m',sp:1}),
    text('bn',cx-r,cy-r*.42,2*r,r*1.1,String(n),r*.62,'#ffffff',{b:1,al:'c',va:'m'}));return els}
M.wheel={n:30,focus:i=>i+2,group:'light',name:SAY('G\'ildirak','Колесо','Wheel'),use:SAY('3–8 bo\'lakli doira, har biri batafsil','круг из 3–8 частей с пояснениями','3–8 part wheel with details'),
  fields:[{k:'title',t:'text'},{k:'word',t:'text'},{k:'parts',t:'list',min:3,max:8,of:[{k:'name',t:'text'},{k:'text',t:'area'}]}],
  sample:{title:'Taqdimot yaratishning 6 qadami',word:'QADAM',parts:[['G\'oya','Mavzuni tanlang va tinglovchiga qanday foyda berishini bir jumlada yozing.'],['Tuzilma','Kirish, asosiy qism va xulosani rejalang: har biriga 2–3 slayd.'],
    ['Dizayn','Bitta rang mavzusi va ikki shrift tanlang.'],['Vizual','Matnni rasm, sxema va raqamlar bilan almashtiring.'],['Mashq','Chiqishni ovoz chiqarib uch marta takrorlang.'],['Chiqish','Asosiy fikrni oxirida yana takrorlang.']].map(([name,text])=>({name,text}))},
  build(f,P){const c=P.c,pa=f.parts||[],T=text('title',0,90,1600,40,f.title,16,c.muted,{al:'c'}),slides=[];
    slides.push({bg:c.frame,els:[sheet(c),T,...wheelEls(c,pa,f.word,800,460,190,{explode:140})],dur:1.2});
    slides.push({bg:c.frame,els:[sheet(c),T,...wheelEls(c,pa,f.word,800,480,300)],dur:1.4,anims:[...pa.map((_,i)=>({n:'sn'+i,k:'zoom',d:.12*i,u:.35})),{n:'badge',k:'spin',d:.12*pa.length+.1,u:.8}]});
    pa.forEach((o,k)=>slides.push({bg:c.frame,dur:1.2,els:[sheet(c),T,...wheelEls(c,pa,f.word,1120,490,250,{pull:k,dim:true}),box('card',140,230,600,470,'#ffffff',{r:.07,line:c.line,lw:1.5}),
      box('cnum',190,280,90,90,TILE[k],{r:.28}),text('cnumt',190,280,90,90,String(k+1).padStart(2,'0'),26,'#ffffff',{b:1,al:'c',va:'m'}),
      text('ch',310,290,400,70,o.name,30,c.ink,{b:1,va:'m'}),text('cd',190,400,500,260,o.text,16,c.muted,{lh:1.25})],anims:[{n:'cd',k:'rise',d:.1,u:.5},{n:'cnum',k:'pulse',d:.4,u:.3}]}));
    return slides}};

/* 31 Flip Cards */
M.flipCards={n:31,focus:i=>i+1,group:'light',name:SAY('Ochiluvchi kartalar','Открывающиеся карточки','Flip Cards'),use:SAY('raqamli kartalar birma-bir ochiladi','нумерованные карточки открываются по очереди','numbered cards open one by one'),
  fields:[{k:'title',t:'text'},{k:'tag',t:'text'},{k:'fixWord',t:'text'},{k:'cards',t:'list',min:2,max:6,of:[{k:'label',t:'text'},{k:'title',t:'text'},{k:'text',t:'area'},{k:'fix',t:'area'}]}],
  sample:{title:'Taqdimotdagi 5 ta xato',tag:'XATO',fixWord:'QANDAY TUZATISH',cards:[['Ko\'p matn','Juda ko\'p matn','Slaydda matn ko\'p bo\'lsa, tinglovchi o\'qiydi, sizni esa eshitmaydi.','Bitta slayd — bitta fikr, 6 qatordan oshmasin.'],
    ['Mayda shrift','Mayda shrift','Orqa qatordagilar 18 pt dan kichik matnni o\'qiy olmaydi.','Matn kamida 24 pt, sarlavha 36 pt.'],['Xulosa yo\'q','Xulosa yo\'q','Tinglovchi «rahmat» bilan emas, aniq harakat bilan ketishi kerak.','Oxirgi slaydda — keyingi qadam.'],
    ['Ko\'p rang','Juda ko\'p rang','Har slaydda yangi rang tinglovchini chalg\'itadi.','Bitta asosiy rang va bitta urg\'u rangi.'],['Rasm sifati','Xira rasmlar','Past sifatli rasm butun taqdimotni arzon ko\'rsatadi.','Faqat aniq, katta o\'lchamli rasmlar.']].map(([label,title,text,fix])=>({label,title,text,fix}))},
  build(f,P){const c=P.c,cs=f.cards||[],n=cs.length,T=text('title',0,90,1600,40,f.title,16,c.muted,{al:'c'}),slides=[];
    const els0=[sheet(c),T];cs.forEach((o,i)=>{const x=800-n*57+i*115,y=380+(i%2?50:0);
      els0.push(box('b'+i,x,y,100,100,c.navy,{r:.28}),text('bn'+i,x,y+14,100,40,String(i+1).padStart(2,'0'),22,'#ffffff',{b:1,al:'c'}),text('bl'+i,x+6,y+58,88,36,o.label,9,c.onNavy,{al:'c'}))});
    els0.push(box('card',800,450,1,1,'#ffffff'));slides.push({bg:c.frame,els:els0,anims:cs.map((_,i)=>({n:'b'+i,k:'zoom',d:.1*i,u:.35}))});
    cs.forEach((o,k)=>{const els=[sheet(c),T,box('card',560,190,480,560,'#ffffff',{r:.07,line:c.line,lw:1.5}),box('tag',600,170,170,40,'#e8432f',{r:.5}),
      text('tagt',600,170,170,40,`${f.tag} ${String(k+1).padStart(2,'0')}`,11,'#ffffff',{b:1,al:'c',va:'m',sp:1}),text('ch',610,240,380,110,o.title,24,c.ink,{b:1,va:'b'}),
      text('cd',610,365,380,170,o.text,14,c.muted,{lh:1.2}),box('fix',600,560,400,150,c.soft,{r:.1}),text('fixh',625,578,360,24,f.fixWord,10,c.navy,{b:1,sp:2}),text('fixt',625,608,360,90,o.fix,13,c.ink)];
      cs.forEach((q,i)=>{let x,y,col,sz=90;if(i<k){x=520-(k-i)*110;y=430;col='#c3c9d6'}else if(i===k){x=760;y=440;col=c.navy;sz=1}else{x=1110+(i-k-1)*110;y=430+(i%2?40:0);col=c.navy}
        els.push(box('b'+i,x,y,sz,sz,col,{r:.28}),text('bn'+i,x,y+12,sz,40,i===k?'':String(i+1).padStart(2,'0'),20,'#ffffff',{b:1,al:'c'}),text('bl'+i,x+4,y+52,Math.max(sz-8,1),32,i===k?'':q.label,8,'#ffffff',{al:'c'}))});
      slides.push({bg:c.frame,els,anims:[{n:'cd',k:'fade',d:.1,u:.5},{n:'fix',k:'rise',d:.3,u:.5},{n:'fixh',k:'fade',d:.4,u:.4},{n:'fixt',k:'fade',d:.45,u:.4}]})});
    return slides}};

/* 32 Photo-in-Text Title and 33 Wave Photo: the photo is the user's own picture */
M.photoTitle={n:32,group:'photo',name:SAY('Rasmli sarlavha','Фото в буквах','Photo-in-Text Title'),use:SAY('sarlavha harflari rasm bilan to\'ldiriladi','буквы заголовка заполнены фото','title letters filled with your photo'),
  fields:[{k:'word',t:'text'},{k:'subtitle',t:'text'},{k:'img',t:'img'}],sample:{word:'SARLAVHA',subtitle:'SARLAVHA OSTIDAGI QISQA IZOH'},
  build(f,P){const c=P.c,img=f.img||'photo:0',len=Math.max(4,(f.word||'').length);
    return[{bg:c.bg,els:[pic('photo',img,0,0,1600,900,{a:.18}),box('shade',0,0,1600,900,c.bg,{a:.55}),
      text('title',20,300,1560,240,f.word,Math.min(128,Math.max(44,1300/len)),'#ffffff',{b:1,al:'c',va:'m',sp:4,fillImg:img}),
      text('sub',100,560,1400,40,f.subtitle,14,c.text,{al:'c',sp:3})],anims:[{n:'title',k:'zoom',d:0,u:.8},{n:'sub',k:'fade',d:.5,u:.6}]}]}};
M.wavePhoto={n:33,group:'photo',name:SAY('To\'lqinli rasm','Фото с волной','Wave Photo'),use:SAY('rasm, to\'lqin va ikki ustun matn','фото, волна и две колонки текста','photo, wave and two text columns'),
  fields:[{k:'title',t:'text'},{k:'col1',t:'area'},{k:'col2',t:'area'},{k:'img',t:'img'}],
  sample:{title:'SARLAVHA',col1:'Birinchi ustun: qisqa tavsif matni. Bu yerda asosiy fikrni ikki-uch qatorda yozing.',col2:'Ikkinchi ustun: qo\'shimcha ma\'lumot yoki muhim fakt shu yerda joylashadi.'},
  build(f,P){const c=P.c;return[{bg:c.bg,els:[pic('photo',f.img||'photo:0',0,0,1600,900),box('shade',0,0,1600,900,c.bg,{a:.01}),
    {n:'wave',t:'wave',x:0,y:470,w:1600,h:430,yl:470,ym:740,yr:470,fill:c.bg},text('title',100,40,1400,120,f.title,40,'#ffffff',{b:1,al:'c',va:'m',sp:4}),
    text('c1',110,750,560,120,f.col1,14,c.text,{lh:1.2}),text('c2',930,750,560,120,f.col2,14,c.text,{al:'r',lh:1.2})],
    anims:[{n:'c1',k:'rise',d:.1,u:.6},{n:'c2',k:'rise',d:.3,u:.6}]}]}};

const ORDER=['glassTitle','divider','accordion','contact','stepsHub','stepPhoto','wheel','flipCards','photoTitle','wavePhoto'];

/* ---------- building slides for a page ---------- */
const BASE=new Set(['sheet','panel','plate']);
function slidesFor(id,f,T){const P={c:palette(T)[M[id].group],logo:null};const out=M[id].build(f,P)||[];
  return out.map(s=>{const els=s.els.map(e=>({...e}));
    /* the deck / page background picture sits on the base layer, under cards and text */
    if(T.bgImgId&&T.images&&T.images[T.bgImgId]&&T.bgVis>0){let k=-1;els.forEach((e,i)=>{if(BASE.has(e.n))k=i});
      els.splice(k+1,0,{n:'bgimg',t:'pic',img:T.bgImgId,x:0,y:0,w:1600,h:900,a:T.bgVis,bg:true})}
    return{...s,els}})}

/* ---------- drawing (preview) ---------- */
function fitOf(img){const f=img&&img._fit||{};return{z:Math.max(1,(+f.z||100)/100),x:(f.x??50)/100,y:(f.y??50)/100,o:(f.o??100)/100}}
const PX=1/0.6;   /* 1 pt = 1/0.6 design pixels on a 1600-wide slide */
function imgFor(key,T){if(!key)return null;if(T.images&&T.images[key])return T.images[key];
  if(/^(ribbon|soft|photo):/.test(key))return genImage(key,T.accent);return null}
function fontOf(e,T){return e.font==='head'?`"${T.headFont||'Georgia'}", Georgia, serif`:`"${T.bodyFont||'Calibri'}", Calibri, Arial, sans-serif`}
function wrap(x,s,w){const out=[];String(s).split('\n').forEach(p=>{const ws=p.split(/\s+/).filter(Boolean);let line='';if(!ws.length){out.push('');return}
  ws.forEach(word=>{const t=line?line+' '+word:word;if(line&&x.measureText(t).width>w){out.push(line);line=word}else line=t});out.push(line)});return out}
function shapePath(x,e){const{w,h}=e;x.beginPath();
  if(e.t==='oval')x.ellipse(w/2,h/2,w/2,h/2,0,0,7);
  else if(e.t==='round')x.roundRect(0,0,w,h,Math.min(w,h)*(e.r||0));
  else if(e.t==='poly'){e.pts.forEach(([a,b],i)=>i?x.lineTo(a,b):x.moveTo(a,b));x.closePath()}
  else if(e.t==='wave'){const t=e.y;x.moveTo(0,e.yl-t);x.bezierCurveTo(w*.25,e.yl-t,w*.32,e.ym-t,w*.5,e.ym-t);x.bezierCurveTo(w*.68,e.ym-t,w*.75,e.yr-t,w,e.yr-t);x.lineTo(w,h);x.lineTo(0,h);x.closePath()}
  else x.rect(0,0,w,h)}
const _mc=document.createElement('canvas').getContext('2d');
const _ptCache=new Map();
function textPt(e,T){const R=ST.ratios(T),base=e.sz*(e.font==='head'?R.head:R.body);
  if(T.fit===false||T.fit==='no'||!e.tx)return base;
  const key=[e.tx,base,e.w,e.h,e.b,e.sp,e.lh,fontOf(e,T)].join('|');if(_ptCache.has(key))return _ptCache.get(key);
  if(_ptCache.size>4000)_ptCache.clear();const v=_textPt(e,T,base);_ptCache.set(key,v);return v}
function _textPt(e,T,pt){
  for(let i=0;i<30&&pt>6;i++){const px=pt*PX;_mc.font=`${e.b?700:400} ${px}px ${fontOf(e,T)}`;if('letterSpacing' in _mc)_mc.letterSpacing=((e.sp||0)*PX)+'px';
    const n=wrap(_mc,e.tx,e.w+2).length;if(n*px*1.2*(e.lh||1)<=Math.max(e.h,px*1.25)+2&&wrap(_mc,e.tx,1e9).every(l=>_mc.measureText(l).width<=e.w+2||l.split(/\s+/).length>1))return pt;pt*=.94}
  return pt}
function drawEl(x,e,T,alpha=1,mult=1){if(alpha<=0.003||e.w<.5&&e.h<.5)return;x.save();x.globalAlpha=alpha*(e.t==='pic'||e.t==='text'?1:(e.a??1));
  x.translate(e.x+e.w/2,e.y+e.h/2);if(e.rot)x.rotate(e.rot*Math.PI/180);if(e.sc&&e.sc!==1)x.scale(e.sc,e.sc);x.translate(-e.w/2,-e.h/2);
  if(e.t==='pic'){const im=imgFor(e.img,T);x.globalAlpha=alpha*(e.a??1);x.beginPath();x.roundRect(0,0,e.w,e.h,Math.min(e.w,e.h)*(e.r||0));x.clip();
    if(im){const f=fitOf(im),[W0,H0]=ST.dimsOf(im),k=Math.max(e.w/W0,e.h/H0)*f.z;if(!e.bg)x.globalAlpha*=f.o;x.drawImage(im,(e.w-W0*k)*f.x,(e.h-H0*k)*f.y,W0*k,H0*k)}else{x.fillStyle='#334';x.fillRect(0,0,e.w,e.h)}}
  else if(e.t==='text'){if(e.tx){const size=textPt(e,T)*PX;x.font=`${e.b?700:400} ${size}px ${fontOf(e,T)}`;if('letterSpacing' in x)x.letterSpacing=((e.sp||0)*PX)+'px';
    const lines=wrap(x,e.tx,e.w+2),lh=size*1.2*(e.lh||1),bh=lines.length*lh,y0=e.va==='m'?(e.h-bh)/2:e.va==='b'?e.h-bh:0;
    const paint=ctx=>{ctx.textBaseline='alphabetic';ctx.textAlign='left';ctx.direction='ltr';lines.forEach((ln,i)=>{const lw=ctx.measureText(ln).width,X=e.al==='c'?(e.w-lw)/2:e.al==='r'?e.w-lw:0;ctx.fillText(ln,X,y0+i*lh+size*.95)})};
    const im=e.fillImg&&imgFor(e.fillImg,T);
    if(im){const c2=document.createElement('canvas');c2.width=Math.ceil(e.w);c2.height=Math.ceil(e.h);const y=c2.getContext('2d');y.font=x.font;if('letterSpacing' in y)y.letterSpacing=x.letterSpacing;y.fillStyle='#000';paint(y);
      y.globalCompositeOperation='source-in';const [W0,H0]=ST.dimsOf(im),k=Math.max(e.w/W0,e.h/H0);y.drawImage(im,(e.w-W0*k)/2,(e.h-H0*k)/2,W0*k,H0*k);x.drawImage(c2,0,0)}
    else{x.fillStyle=e.col;paint(x)}}}
  else{shapePath(x,e);if(e.fill){x.fillStyle=e.fill;x.fill()}if(e.line){x.globalAlpha=alpha;x.strokeStyle=e.line;x.lineWidth=(e.lw||1)*PX;x.stroke()}}
  x.restore()}
/* in-slide animations: returns [alpha, extra] for an element at time t after the slide settles */
function animState(s,name,t){let a=1,dy=0,sc=1,rot=0;for(const an of s.anims||[]){if(an.n!==name)continue;const p=clamp((t-an.d)/an.u),e=eo(p);
  if(an.k==='fade'){a*=p}else if(an.k==='rise'){a*=p;dy+=40*(1-e)}else if(an.k==='zoom'){a*=p>0?1:0;sc*=lerp(.05,1,e)}
  else if(an.k==='pulse'){sc*=1+.14*Math.sin(Math.PI*clamp(p))}else if(an.k==='spin'){rot+=360*eio(p)}}return{a,dy,sc,rot}}
function animEnd(s){return Math.max(0,...(s.anims||[]).map(a=>a.d+a.u))}
/* blend two slides that share element names, as Morph does */
function lerpColor(a,b,k){if(!a||!b||a===b)return k<.5?a:b;return mix(a,b,k)}
function blendEl(A,B,k){const e={...B};for(const p of ['x','y','w','h','rot','a','lw','sz','r'])if(A[p]!=null&&B[p]!=null)e[p]=lerp(+A[p],+B[p],k);
  e.fill=lerpColor(A.fill,B.fill,k);e.line=lerpColor(A.line,B.line,k);e.col=lerpColor(A.col,B.col,k);
  if(A.t==='poly'&&B.t==='poly'&&A.pts.length===B.pts.length)e.pts=B.pts.map((p,i)=>[lerp(A.pts[i][0],p[0],k),lerp(A.pts[i][1],p[1],k)]);return e}
function drawSlide(x,s,T,t,mult){x.fillStyle=s.bg;x.fillRect(0,0,1600,900);
  for(const e of s.els){const st=animState(s,e.n,t);const ee={...e,y:e.y+st.dy,sc:(e.sc||1)*st.sc,rot:(e.rot||0)+st.rot};drawEl(x,ee,T,st.a,mult)}}
function drawTransition(x,A,B,T,k,mult){const kk=eio(k);x.fillStyle=lerpColor(A?A.bg:B.bg,B.bg,kk);x.fillRect(0,0,1600,900);
  const names=new Set(B.els.map(e=>e.n)),aMap=new Map((A?A.els:[]).map(e=>[e.n,e]));
  if(A)for(const e of A.els)if(!names.has(e.n))drawEl(x,e,T,1-kk,mult);
  for(const e of B.els){const a=aMap.get(e.n),hidden=(B.anims||[]).some(an=>an.n===e.n&&['fade','rise','zoom'].includes(an.k));
    if(a){if(a.t==='text'&&a.tx!==e.tx){const m=blendEl(a,e,kk);drawEl(x,{...m,tx:a.tx},T,1-kk,mult);if(!hidden)drawEl(x,m,T,kk,mult)}
      else drawEl(x,blendEl(a,e,kk),T,hidden?1-kk:1,mult)}
    else if(!hidden)drawEl(x,e,T,kk,mult)}}
/* a full playback timeline for one page */
const TRANS=1.25,HOLD=1.4;
function timeline(slides){let t=0;return slides.map(s=>{const d=s.dur||TRANS,seg={s,start:t,trans:d,anim:animEnd(s)};t+=d+seg.anim+HOLD;seg.end=t;return seg})}
function render(ctx,id,f,t,T){const slides=slidesFor(id,f,T),tl=timeline(slides),mult=ST.SIZES[T.size]||1;const total=tl.length?tl[tl.length-1].end:1;
  let i=tl.findIndex(g=>t<g.end);if(i<0)i=tl.length-1;const g=tl[i],lt=t-g.start;
  if(lt<g.trans)drawTransition(ctx,i?tl[i-1].s:null,g.s,T,lt/g.trans,mult);else drawSlide(ctx,g.s,T,lt-g.trans,mult);return total}
/* the moment of the page timeline where slide k is fully shown (used while editing item k) */
function timeAtSlide(id,f,T,item){const tl=timeline(slidesFor(id,f,T));if(!tl.length)return 0;const m=M[id].focus;
  const k=Math.max(0,Math.min(tl.length-1,m?m(item,tl.length):tl.length-1)),g=tl[k];return g.start+g.trans+g.anim+.5}
function durOf(id,f,T){const tl=timeline(slidesFor(id,f,T));return tl.length?tl[tl.length-1].end:1}

/* ---------- join the page-style list used by the editor and the exporter ---------- */
ST.THEMES.corporate={name:'Corporate Green',bg:'#0b2a24',accent:'#2fb784',text:'#e8f5ef'};
ST.THEMES.light={name:'Clean Light',bg:'#1d3a8a',accent:'#2f6fe4',text:'#eef2ff'};
const DEF={bg:'#16264a',accent:'#d9b25f',text:'#ece6d8'};
ORDER.forEach(id=>{ST.S[id]={...M[id],id,kind:'morph',dur:0,durFor:f=>durOf(id,f,DEF),slots:()=>[],draw(){}};ST.ORDER.push(id)});
const baseRender=ST.render;
ST.render=function(ctx,id,f,t,T,o){if(ST.S[id]&&ST.S[id].kind==='morph'){ctx.save();render(ctx,id,f,t,T);ctx.restore();return}return baseRender(ctx,id,f,t,T,o)};

window.IMXMorph={textPt,fitOf,timeAtSlide,M,ORDER,TILE,slidesFor,render,durOf,genImage,imgFor,palette,PX};
})();
