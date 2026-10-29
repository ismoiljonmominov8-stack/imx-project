/* IMX PPTX exporter.
   For every page: the background animation is rendered in the browser and encoded as an MP4
   (WebCodecs H.264, Chrome/Edge) or, if that is not available, an animated GIF. The text becomes
   real PowerPoint text boxes with entrance animations. Everything starts by itself in Slide Show. */
(function(){
'use strict';
const ST=window.IMXStyles,EMU=12192000/1600,FPS=25,HOLD=.6;
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const hex=c=>(c||'#ffffff').replace('#','').toUpperCase();
const E=v=>Math.round(v*EMU);
const NS='xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const LANG={uz:'uz-Latn-UZ',ru:'ru-RU',en:'en-US'};

/* ---------- rendering the background animation ---------- */
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function pickCodec(){
  if(!('VideoEncoder' in window))return null;
  for(const codec of ['avc1.640028','avc1.4d0028','avc1.42001f']){
    try{const c={codec,width:1280,height:720,bitrate:5e6,framerate:FPS,avc:{format:'avc'}};if((await VideoEncoder.isConfigSupported(c)).supported)return c}catch(_){}}
  return null}
/* videos in picture slots: the export works on its own copies so the live preview is not disturbed */
const isVid=el=>typeof HTMLVideoElement!=='undefined'&&el instanceof HTMLVideoElement;
const once=(el,ev)=>new Promise(r=>el.addEventListener(ev,r,{once:true}));
const dimsOf=el=>({w:el.videoWidth||el.naturalWidth||el.width,h:el.videoHeight||el.naturalHeight||el.height});
async function exportCopies(T){const made=new Map();
  for(const k in (T.images||{})){const el=T.images[k];if(!isVid(el))continue;
    if(!made.has(el)){const v=document.createElement('video');v.muted=true;v.playsInline=true;v.preload='auto';v.src=el.src;
      if(v.readyState<2)await Promise.race([once(v,'loadeddata'),sleep(10000)]);v._dur=await realDuration(v);v._fit=el._fit;made.set(el,v)}
    T.images[k]=made.get(el)}
  return [...made.values()]}
async function realDuration(v){if(v.readyState<1)await Promise.race([once(v,'loadedmetadata'),sleep(10000)]);
  if(!(isFinite(v.duration)&&v.duration>0)){const p=Promise.race([once(v,'durationchange'),sleep(6000)]);try{v.currentTime=1e7}catch(_){}await p;
    const q=Promise.race([once(v,'seeked'),sleep(3000)]);v.currentTime=0;await q}
  return isFinite(v.duration)&&v.duration>0?v.duration:0}
async function seekAll(vids,t){await Promise.all(vids.map(v=>{const d=v._dur||(isFinite(v.duration)&&v.duration)||1,tt=((t%d)+d)%d;if(Math.abs(v.currentTime-tt)<.002)return null;
  const p=Promise.race([once(v,'seeked'),sleep(2000)]);v.currentTime=tt;return p}))}
const VEXT={'video/mp4':'mp4','video/webm':'webm','video/quicktime':'mov','video/x-m4v':'m4v'};
/* which codec is inside an MP4 / MOV file (looks for the sample-entry name) */
async function codecOf(b){const look=async blob=>{const t=new Uint8Array(await blob.arrayBuffer());let s='';for(let i=0;i<t.length;i++)s+=String.fromCharCode(t[i]);return s};
  const head=await look(b.slice(0,Math.min(b.size,4e6))),tail=b.size>4e6?await look(b.slice(b.size-4e6)):'';const all=head+tail;
  return /hvc1|hev1/.test(all)?'hevc':/avc1/.test(all)?'h264':'other'}
/* PowerPoint plays H.264 MP4 smoothly on every computer. Anything else (WebM, HEVC / iPhone "High Efficiency",
   very large 4K or 60 fps files, very high bitrate) is converted to a 720p H.264 MP4 when the browser can encode. */
async function videoBytes(el,cfg){const b=await (await fetch(el.src)).blob(),ext=VEXT[b.type]||'mp4';
  let convert=ext!=='mp4';
  if(!convert){const codec=await codecOf(b),d=el._dur||el.duration||0,big=(el.videoWidth||0)*(el.videoHeight||0)>1920*1080,heavy=d>0&&b.size*8/d>16e6;
    convert=codec!=='h264'||big||heavy}
  if(convert&&cfg){try{return{data:await transcodeMP4(el,cfg),ext:'mp4'}}catch(e){console.warn('video convert failed',e)}}
  return{data:new Uint8Array(await b.arrayBuffer()),ext}}
async function transcodeMP4(src,cfg,maxDur=1800){
  const v=(await exportCopies({images:{k:src}}))[0],d0=dimsOf(v),k=Math.min(1,1280/d0.w,720/d0.h);
  const w=Math.max(2,Math.round(d0.w*k/2)*2),h=Math.max(2,Math.round(d0.h*k/2)*2);
  let c2={...cfg,width:w,height:h};try{if(!(await VideoEncoder.isConfigSupported(c2)).supported)c2={...cfg,codec:'avc1.42001f',width:w,height:h}}catch(_){}
  const c=frameCanvas(w,h),x=c.getContext('2d'),target=new Mp4Muxer.ArrayBufferTarget();
  const muxer=new Mp4Muxer.Muxer({target,video:{codec:'avc',width:w,height:h},fastStart:'in-memory'});
  let err=null;const enc=new VideoEncoder({output:(ch,meta)=>muxer.addVideoChunk(ch,meta),error:e=>{err=e}});enc.configure(c2);
  const end=Math.min(maxDur,v._dur||600);
  let last=-1,nf=0;const put=ts=>{x.drawImage(v,0,0,w,h);const fr=new VideoFrame(c,{timestamp:ts});enc.encode(fr,{keyFrame:nf++%50===0});fr.close();last=ts};
  if('requestVideoFrameCallback' in v){await seekAll([v],0);put(0);
    await new Promise(res=>{let seen=performance.now(),done=false;const stop=()=>{if(done)return;done=true;v.pause();clearInterval(watch);res()};
      const cb=(now,meta)=>{if(done)return;seen=performance.now();if(err||meta.mediaTime>=end-.001){stop();return}
        const ts=Math.round(meta.mediaTime*1e6);if(ts>last)put(ts);v.requestVideoFrameCallback(cb)};
      /* stop only when the video really ends or stops delivering frames for a while */
      const watch=setInterval(()=>{if(v.ended||performance.now()-seen>4000)stop()},500);
      v.requestVideoFrameCallback(cb);v.onended=stop;v.playbackRate=1;v.play().catch(stop)})}
  /* anything not captured yet (or no frame callbacks in this browser): step through the rest frame by frame */
  if(last/1e6<end-.2){const from=Math.max(0,Math.ceil((last/1e6+.001)*FPS)),n=Math.floor(end*FPS);
    for(let i=from;i<n;i++){if(err)throw err;await seekAll([v],i/FPS);const ts=Math.round(i*1e6/FPS);if(ts>last)put(ts);while(enc.encodeQueueSize>8)await sleep(4)}}
  await enc.flush();enc.close();muxer.finalize();if(err)throw err;return new Uint8Array(target.buffer)}
function frameCanvas(w,h){const c=document.createElement('canvas');c.width=w;c.height=h;return c}
async function posterOf(page,T,at=0,vids=[],vt=at){const c=frameCanvas(1280,720),x=c.getContext('2d');x.setTransform(.8,0,0,.8,0,0);
  if(vids.length)await seekAll(vids,vt);ST.render(x,page.style,page.f,at,T,{text:false});return new Uint8Array(await (await new Promise(r=>c.toBlob(r,'image/jpeg',.88))).arrayBuffer())}
/* intro: the page animation; loop: the finished page while the videos in it keep playing */
async function renderMP4(page,T,cfg,dur,progress,{vids=[],loop=0}={}){
  if(vids.length)await seekAll(vids,.3);
  const c=frameCanvas(1280,720),x=c.getContext('2d',{willReadFrequently:false});x.setTransform(.8,0,0,.8,0,0);
  const target=new Mp4Muxer.ArrayBufferTarget();
  const muxer=new Mp4Muxer.Muxer({target,video:{codec:'avc',width:1280,height:720},fastStart:'in-memory'});
  let err=null;const enc=new VideoEncoder({output:(chunk,meta)=>muxer.addVideoChunk(chunk,meta),error:e=>{err=e}});enc.configure(cfg);
  const n=Math.ceil((loop||dur+HOLD)*FPS);
  for(let i=0;i<n;i++){if(err)throw err;
    const st=loop?dur+HOLD:Math.min(i/FPS,dur+HOLD);
    x.setTransform(.8,0,0,.8,0,0);ST.render(x,page.style,page.f,st,T,{text:false});
    const fr=new VideoFrame(c,{timestamp:Math.round(i*1e6/FPS),duration:Math.round(1e6/FPS)});enc.encode(fr,{keyFrame:i%(FPS*2)===0});fr.close();
    while(enc.encodeQueueSize>8)await sleep(4);if(i%5===0){progress(i/n);await sleep(0)}}
  await enc.flush();enc.close();muxer.finalize();if(err)throw err;return new Uint8Array(target.buffer)}
async function renderGIF(page,T,dur,progress){
  const {GIFEncoder,quantize,applyPalette}=window.gifenc,w=960,h=540,fps=12,c=frameCanvas(w,h),x=c.getContext('2d',{willReadFrequently:true});
  const draw=t=>{x.setTransform(.6,0,0,.6,0,0);ST.render(x,page.style,page.f,t,T,{text:false});return x.getImageData(0,0,w,h).data};
  const palette=quantize(draw(dur+HOLD),256,{format:'rgb565'});const gif=GIFEncoder(),n=Math.ceil((dur+HOLD)*fps);
  for(let i=0;i<n;i++){const data=draw(Math.min(i/fps,dur+HOLD));gif.writeFrame(applyPalette(data,palette,'rgb565'),w,h,{palette:i?undefined:palette,delay:Math.round(1000/fps),repeat:-1});
    if(i%3===0){progress(i/n);await sleep(0)}}
  gif.finish();return gif.bytes()}

/* ---------- PowerPoint entrance effects ---------- */
function effectXML(effect,spid,delay,speed,first,id){
  const dur=Math.round((ST.SPEED[speed]||.6)*1000),T=()=>id.n++,outer=T();
  const tgt=`<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
  const set=`<p:set><p:cBhvr><p:cTn id="${T()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
  const fx=(filter,d=dur)=>`<p:animEffect transition="in" filter="${filter}"><p:cBhvr><p:cTn id="${T()}" dur="${d}"/>${tgt}</p:cBhvr></p:animEffect>`;
  const anim=(attr,from,to,d=dur)=>`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${T()}" dur="${d}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const M={
    fade:[10,0,()=>set+fx('fade')],
    rise:[42,0,()=>set+fx('fade')+anim('ppt_x','#ppt_x','#ppt_x')+anim('ppt_y','#ppt_y+.1','#ppt_y')],
    fly:[2,4,()=>set+anim('ppt_x','#ppt_x','#ppt_x')+anim('ppt_y','1+#ppt_h/2','#ppt_y')],
    wipe:[22,8,()=>set+fx('wipe(left)')],
    split:[16,42,()=>set+fx('barn(outVertical)')],
    zoom:[53,16,()=>set+anim('ppt_w','0','#ppt_w')+anim('ppt_h','0','#ppt_h')+fx('fade')],
    typewriter:[10,0,()=>set+fx('fade',120)],
    words:[10,0,()=>set+fx('fade',400)]};
  const [preset,sub,body]=M[effect]||M.fade;
  const iter=effect==='typewriter'?'<p:iterate type="lt"><p:tmPct val="25000"/></p:iterate>':effect==='words'?'<p:iterate type="wd"><p:tmPct val="30000"/></p:iterate>':'';
  return `<p:par><p:cTn id="${outer}" presetID="${preset}" presetClass="entr" presetSubtype="${sub}" fill="hold" grpId="0" nodeType="${first?'afterEffect':'withEffect'}"><p:stCondLst><p:cond delay="${Math.round(delay*1000)}"/></p:stCondLst>${iter}<p:childTnLst>${body()}</p:childTnLst></p:cTn></p:par>`}

/* a muted, looping video in a picture place (cover-fit, zoom and position from the picture settings) */
function videoPicXML(id,o,mrid,vrid,prid){const f=o.fit||{z:1,x:.5,y:.5};
  const tr=o.w/o.h,ir=o.vw/o.vh;let vw=1,vh=1;if(ir>tr)vw=tr/ir;else vh=ir/tr;vw/=f.z;vh/=f.z;
  const l=(1-vw)*f.x,r=(1-vw)*(1-f.x),t=(1-vh)*f.y,b=(1-vh)*(1-f.y);
  const g=o.geom||{t:'rect'},e={t:g.t==='poly'?'poly':g.t,r:g.r,w:o.w,h:o.h,pts:g.pts?g.pts.map(([a,c])=>[a-o.x,c-o.y]):null};
  return`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Video ${id}"></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="${vrid}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${mrid}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${prid}"/><a:srcRect l="${Math.round(l*1e5)}" t="${Math.round(t*1e5)}" r="${Math.round(r*1e5)}" b="${Math.round(b*1e5)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm${o.rot?` rot="${Math.round(((o.rot%360)+360)%360*60000)}"`:''}><a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${E(o.w)}" cy="${E(o.h)}"/></a:xfrm>${geomXML(e)}</p:spPr></p:pic>`}

/* ---------- one slide ---------- */
function textShape(id,s,fit,color,T,lang){
  const sz=Math.max(800,Math.round(fit.size*.6*100)),spc=fit.spc?Math.round(fit.spc*.6*100):0,algn={l:'l',c:'ctr',r:'r'}[s.align||'l'],anchor={t:'t',m:'ctr',b:'b'}[s.valign||'t'];
  const rot=s.rot?` rot="${Math.round(s.rot*60000)}"`:'',shadow=color.toUpperCase().startsWith('#2')?'':'<a:effectLst><a:outerShdw blurRad="63500" dist="19050" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="55000"/></a:srgbClr></a:outerShdw></a:effectLst>';
  const rpr=`<a:rPr lang="${lang}" sz="${sz}" b="${fit.weight>=700?1:0}"${spc?` spc="${spc}"`:''} dirty="0"><a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>${shadow}<a:latin typeface="${esc(fit.font.match(/"([^"]+)"/)[1])}"/><a:cs typeface="${esc(fit.font.match(/"([^"]+)"/)[1])}"/></a:rPr>`;
  const paras=String(s.text).split('\n').map(line=>`<a:p><a:pPr algn="${algn}"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>${line?`<a:r>${rpr}<a:t>${esc(line)}</a:t></a:r>`:''}<a:endParaRPr lang="${lang}" sz="${sz}" dirty="0"/></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc((s.key||'Text')+' '+id)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${E(s.x)}" y="${E(s.y)}"/><a:ext cx="${E(s.w)}" cy="${E(s.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp>`}
function slideXML(page,T,media,lang,measure){
  const id={n:3},shapes=[],effects=[];let spid=3;
  const vid=2,picRels=media.type==='mp4'
    ?`<p:pic><p:nvPicPr><p:cNvPr id="${vid}" name="Animation"></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rId3"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId2"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    :`<p:pic><p:nvPicPr><p:cNvPr id="${vid}" name="Animation"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  shapes.push(picRels);
  /* videos placed in picture places: real PowerPoint videos on top of the page animation */
  (media.over||[]).forEach((o,k)=>shapes.push(videoPicXML(100+k,o,`rIdV${k}m`,`rIdV${k}v`,`rIdV${k}p`)));
  const slots=ST.slotsOf(page.style,page.f);
  slots.forEach(s=>{const sid=spid++,fit=measure(s),color=measure.color(s);shapes.push(textShape(sid,s,fit,color,T,lang));effects.push([sid,s.at])});
  /* timing: the video plays from the start; each text box enters at its own moment */
  id.n=5;let inner='';
  if(media.type==='mp4'){const a=id.n++,b=id.n++;
    inner+=`<p:par><p:cTn id="${a}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${b}" dur="${Math.round(media.dur*1000)}" fill="hold"/><p:tgtEl><p:spTgt spid="${vid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`}
  let firstFx=media.type!=='mp4';
  effects.forEach(([sid,at],i)=>{inner+=effectXML(T.effect||'fade',sid,at,T.speed,media.type!=='mp4'&&!(media.over||[]).length&&i===0,id)});
  (media.over||[]).forEach((o,k)=>{const sp=100+k;inner+=effectXML('fade',sp,o.at,'fast',firstFx,id);firstFx=false;
    inner+=`<p:par><p:cTn id="${id.n++}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="${Math.round(o.at*1000)}"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${id.n++}" dur="${Math.round(o.dur*1000)}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`});
  let timing='';
  if(inner){const mediaNode=(media.type==='mp4'?`<p:video><p:cMediaNode vol="80000" mute="1"><p:cTn id="${id.n++}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${vid}"/></p:tgtEl></p:cMediaNode></p:video>`:'')
      +(media.over||[]).map((o,k)=>`<p:video><p:cMediaNode vol="0" mute="1"><p:cTn id="${id.n++}" repeatCount="indefinite" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${100+k}"/></p:tgtEl></p:cMediaNode></p:video>`).join('');
    timing=`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${mediaNode}</p:childTnLst></p:cTn></p:par></p:tnLst>${effects.length||(media.over||[]).length?`<p:bldLst>${(media.over||[]).map((o,k)=>`<p:bldP spid="${100+k}" grpId="0"/>`).join('')}${effects.map(([sid])=>`<p:bldP spid="${sid}" grpId="0"/>`).join('')}</p:bldLst>`:''}</p:timing>`}
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex(ST.mix(T.bg,'#000000',.6))}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition spd="med"><p:fade/></p:transition>${timing}</p:sld>`}
function slideRels(i,media){
  const L='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>';
  const body=media.type==='mp4'
    ?`<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/media${i}.mp4"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media${i}.mp4"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/poster${i}.jpeg"/>`
    :`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/anim${i}.gif"/>`;
  const loop=(media.over||[]).map((o,k)=>`<Relationship Id="rIdV${k}m" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/${o.file}"/><Relationship Id="rIdV${k}v" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/${o.file}"/><Relationship Id="rIdV${k}p" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${o.post}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${L}${body}${loop}</Relationships>`}

/* ---------- Morph pages (templates 24-33): native shapes, Morph between slides ---------- */
const MORPH_TR=d=>`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="${Math.round(d*1000)}"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
const fillXML=(c,a=1)=>c?`<a:solidFill><a:srgbClr val="${hex(c)}">${a<1?`<a:alpha val="${Math.round(clampA(a)*100000)}"/>`:''}</a:srgbClr></a:solidFill>`:'<a:noFill/>';
const clampA=a=>Math.max(0,Math.min(1,+a));
const lineXML=e=>e.line?`<a:ln w="${Math.round((e.lw||1)*12700)}"><a:solidFill><a:srgbClr val="${hex(e.line)}"/></a:solidFill></a:ln>`:'<a:ln><a:noFill/></a:ln>';
const xfrm=e=>`<a:xfrm${e.rot?` rot="${Math.round(((e.rot%360)+360)%360*60000)}"`:''}><a:off x="${E(e.x)}" y="${E(e.y)}"/><a:ext cx="${Math.max(1,E(e.w))}" cy="${Math.max(1,E(e.h))}"/></a:xfrm>`;
function geomXML(e){
  if(e.t==='oval')return'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
  if(e.t==='round')return`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.round(Math.min(.5,e.r||0)*100000)}"/></a:avLst></a:prstGeom>`;
  const W=Math.max(1,E(e.w)),H=Math.max(1,E(e.h)),pt=(x,y)=>`<a:pt x="${Math.round(E(x))}" y="${Math.round(E(y))}"/>`;
  let path='';
  if(e.t==='poly')path=e.pts.map(([x,y],i)=>i?`<a:lnTo>${pt(x,y)}</a:lnTo>`:`<a:moveTo>${pt(x,y)}</a:moveTo>`).join('')+'<a:close/>';
  else if(e.t==='wave'){const t=e.y,w=e.w;path=`<a:moveTo>${pt(0,e.yl-t)}</a:moveTo><a:cubicBezTo>${pt(w*.25,e.yl-t)}${pt(w*.32,e.ym-t)}${pt(w*.5,e.ym-t)}</a:cubicBezTo><a:cubicBezTo>${pt(w*.68,e.ym-t)}${pt(w*.75,e.yr-t)}${pt(w,e.yr-t)}</a:cubicBezTo><a:lnTo>${pt(w,e.h)}</a:lnTo><a:lnTo>${pt(0,e.h)}</a:lnTo><a:close/>`}
  else return'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
  return`<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${W}" h="${H}">${path}</a:path></a:pathLst></a:custGeom>`}
function morphShapeXML(e,id,T,lang,rel,mult){
  const nv=`<p:cNvPr id="${id}" name="!!${esc(e.n)}"/>`;
  if(e.t==='pic'){const im=rel(e.img,e.n==='logo');if(!im)return'';
    const src=window.IMXMorph.imgFor(e.img,T),f=window.IMXMorph.fitOf(src);
    /* visible part of the picture: cover-fit, then zoom, then shift toward the chosen point */
    const tr=e.w/e.h,ir=im.w/im.h;let vw=1,vh=1;if(ir>tr)vw=tr/ir;else vh=ir/tr;vw/=f.z;vh/=f.z;
    const l=(1-vw)*f.x,r=(1-vw)*(1-f.x),t=(1-vh)*f.y,b=(1-vh)*(1-f.y),alpha=(e.a??1)*(e.bg?1:f.o);
    const geo=e.r?`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.round(Math.min(.5,e.r)*100000)}"/></a:avLst></a:prstGeom>`:'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    if(im.video){VIDS.push({sp:id,dur:im.video.dur});
      return`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="!!${esc(e.n)}"></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="${im.video.link}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${im.video.media}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${im.rid}"/><a:srcRect l="${Math.round(l*100000)}" t="${Math.round(t*100000)}" r="${Math.round(r*100000)}" b="${Math.round(b*100000)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}${geo}</p:spPr></p:pic>`}
    return`<p:pic><p:nvPicPr>${nv}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${im.rid}">${alpha<1?`<a:alphaModFix amt="${Math.round(clampA(alpha)*100000)}"/>`:''}</a:blip><a:srcRect l="${Math.round(l*100000)}" t="${Math.round(t*100000)}" r="${Math.round(r*100000)}" b="${Math.round(b*100000)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}${geo}</p:spPr></p:pic>`}
  if(e.t==='text'){const sz=Math.max(100,Math.round(window.IMXMorph.textPt(e,T)*100)),algn={l:'l',c:'ctr',r:'r'}[e.al||'l'],anchor={t:'t',m:'ctr',b:'b'}[e.va||'t'];
    const face=esc(e.font==='head'?(T.headFont||'Georgia'):(T.bodyFont||'Calibri'));
    let fill=fillXML(e.col);if(e.fillImg){const im=rel(e.fillImg);if(im)fill=`<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${im.rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`}
    const rpr=`<a:rPr lang="${lang}" sz="${sz}" b="${e.b?1:0}"${e.sp?` spc="${Math.round(e.sp*100)}"`:''} dirty="0">${fill}<a:latin typeface="${face}"/><a:cs typeface="${face}"/></a:rPr>`;
    const paras=String(e.tx).split('\n').map(line=>`<a:p><a:pPr algn="${algn}"><a:lnSpc><a:spcPct val="${Math.round((e.lh||1)*100000)}"/></a:lnSpc></a:pPr>${line?`<a:r>${rpr}<a:t>${esc(line)}</a:t></a:r>`:''}<a:endParaRPr lang="${lang}" sz="${sz}" dirty="0"/></a:p>`).join('');
    return`<p:sp><p:nvSpPr>${nv}<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp>`}
  return`<p:sp><p:nvSpPr>${nv}<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(e)}${geomXML(e)}${fillXML(e.fill,e.a??1)}${lineXML(e)}</p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="${lang}" dirty="0"/></a:p></p:txBody></p:sp>`}
/* the animations that play inside one Morph slide, after the transition */
function morphTiming(anims,spidOf,vids=[]){const id={n:5},T=()=>++id.n,tgt=sp=>`<p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl>`;
  const setv=sp=>`<p:set><p:cBhvr><p:cTn id="${T()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt(sp)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
  const anim=(sp,at,a,b,d)=>`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${T()}" dur="${d}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>${at}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${a}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${b}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const fx=(sp,d)=>`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${T()}" dur="${d}"/>${tgt(sp)}</p:cBhvr></p:animEffect>`;
  let body='',first=true;const builds=new Set();
  /* videos in picture places start by themselves with the slide, muted, and loop */
  for(const v of vids){body+=`<p:par><p:cTn id="${T()}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="${first?'afterEffect':'withEffect'}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${T()}" dur="${Math.round((v.dur||5)*1000)}" fill="hold"/>${tgt(v.sp)}</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`;first=false}
  for(const a of anims||[]){const sp=spidOf(a.n);if(!sp)continue;const d=Math.max(50,Math.round(a.u*1000)),delay=Math.round(a.d*1000),node=first?'afterEffect':'withEffect',outer=T();
    if(['zoom','fade','rise'].includes(a.k)){if(builds.has(sp))continue;
      const preset={zoom:53,fade:10,rise:42}[a.k],inner=a.k==='zoom'?setv(sp)+anim(sp,'ppt_w','0','#ppt_w',d)+anim(sp,'ppt_h','0','#ppt_h',d)+fx(sp,d):a.k==='fade'?setv(sp)+fx(sp,d):setv(sp)+fx(sp,d)+anim(sp,'ppt_x','#ppt_x','#ppt_x',d)+anim(sp,'ppt_y','#ppt_y+.1','#ppt_y',d);
      body+=`<p:par><p:cTn id="${outer}" presetID="${preset}" presetClass="entr" presetSubtype="${a.k==='zoom'?16:0}" fill="hold" grpId="0" nodeType="${node}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;builds.add(sp)}
    else if(a.k==='pulse')body+=`<p:par><p:cTn id="${outer}" presetID="6" presetClass="emph" presetSubtype="0" autoRev="1" fill="hold" nodeType="${node}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="${T()}" dur="${d}" fill="hold"/>${tgt(sp)}</p:cBhvr><p:by x="114000" y="114000"/></p:animScale></p:childTnLst></p:cTn></p:par>`;
    else if(a.k==='spin')body+=`<p:par><p:cTn id="${outer}" presetID="8" presetClass="emph" presetSubtype="0" fill="hold" nodeType="${node}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst><p:animRot by="21600000"><p:cBhvr><p:cTn id="${T()}" dur="${d}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>`;
    else continue;first=false}
  if(!body)return'';
  const media=vids.map(v=>`<p:video><p:cMediaNode vol="0" mute="1"><p:cTn id="${T()}" repeatCount="indefinite" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>${tgt(v.sp)}</p:cMediaNode></p:video>`).join('');
  return`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${media}</p:childTnLst></p:cTn></p:par></p:tnLst>${builds.size?`<p:bldLst>${[...builds].map(b=>`<p:bldP spid="${b}" grpId="0" animBg="1"/>`).join('')}</p:bldLst>`:''}</p:timing>`}
let VIDS=[];
function morphSlideXML(s,T,lang,rel,mult){let id=2;const map={},shapes=[];VIDS=[];
  for(const e of s.els){if(e.t==='text'&&!String(e.tx||'').trim())continue;const xml=morphShapeXML(e,id,T,lang,rel,mult);if(!xml)continue;map[e.n]=id;shapes.push(xml);id++}
  return`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld ${NS}><p:cSld><p:bg><p:bgPr>${fillXML(s.bg)}<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${MORPH_TR(s.dur||1.25)}${morphTiming(s.anims,n=>map[n],VIDS)}</p:sld>`}
/* pictures used by Morph slides, saved once each */
async function encodeImage(src,png){const D=dimsOf(src),k=Math.min(1,1920/Math.max(D.w,D.h)),c=frameCanvas(Math.round(D.w*k),Math.round(D.h*k));
  c.getContext('2d').drawImage(src,0,0,c.width,c.height);const b=await new Promise(r=>c.toBlob(r,png?'image/png':'image/jpeg',.9));return{data:new Uint8Array(await b.arrayBuffer()),w:c.width,h:c.height}}

/* ---------- the whole file ---------- */
async function build({pages,themeFor,lang,template,title},onProgress=()=>{}){
  const MO=window.IMXMorph,needVideo=pages.some(p=>ST.S[p.style].kind!=='morph');
  const zip=await JSZip.loadAsync(template),cfg=needVideo?await pickCodec():null,mode=cfg?'mp4':'gif',vcfg=cfg||(await pickCodec());
  const meas=document.createElement('canvas').getContext('2d');const ln=LANG[lang]||'uz-Latn-UZ';
  const IMGS=new Map(),videoExts=new Set();let imgN=0,n=0;
  const VF=new Map();
  async function prepVideo(id,el){if(VF.has(id))return VF.get(id);
    const v=(await exportCopies({images:{k:el}}))[0];await seekAll([v],Math.min(.3,(v.duration||1)/2));
    const post=await encodeImage(v,false),vb=await videoBytes(el,vcfg),k=++imgN,file=`imx_vid${k}.${vb.ext}`,pf=`imx_vpost${k}.jpeg`;
    zip.file('ppt/media/'+file,vb.data);zip.file('ppt/media/'+pf,post.data);videoExts.add(vb.ext);
    const r={file,post:pf,vw:post.w,vh:post.h,dur:v._dur||600};VF.set(id,r);return r}
  const L7='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>';
  const relsDoc=body=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${L7}${body}</Relationships>`;
  for(let i=0;i<pages.length;i++){const p=pages[i],T=themeFor(p),st=ST.S[p.style];
    const prog=k=>onProgress({slide:i+1,total:pages.length,frac:(i+k)/pages.length,mode});prog(0);
    if(st.kind==='morph'){const slides=MO.slidesFor(p.style,p.f,T),mult=ST.SIZES[T.size]||1;
      /* collect the pictures first (encoding is async), then write the XML */
      const need=[];slides.forEach(s=>s.els.forEach(e=>{if(e.t==='pic')need.push([e.img,e.n==='logo']);if(e.fillImg)need.push([e.fillImg,false])}));
      for(const [key,png] of need){const kk=key+'|'+T.accent+(png?'|png':'');if(IMGS.has(kk))continue;const src=MO.imgFor(key,T);if(!src){IMGS.set(kk,null);continue}
        if(isVid(src)){const v=(await exportCopies({images:{k:src}}))[0];await seekAll([v],Math.min(.3,(v._dur||2)/2));
          const enc=await encodeImage(v,false),vb=await videoBytes(src,vcfg),post=`imx_vpost${++imgN}.jpeg`,file=`imx_vid${imgN}.${vb.ext}`;
          zip.file('ppt/media/'+post,enc.data);zip.file('ppt/media/'+file,vb.data);videoExts.add(vb.ext);IMGS.set(kk,{file:post,video:file,w:enc.w,h:enc.h,dur:v._dur||600});continue}
        const enc=await encodeImage(src,png),file=`imx_img${++imgN}.${png?'png':'jpeg'}`;zip.file('ppt/media/'+file,enc.data);IMGS.set(kk,{file,w:enc.w,h:enc.h})}
      for(const s of slides){n++;const rels=[],rid={};
        const rel=(key,png)=>{const im=IMGS.get(key+'|'+T.accent+(png?'|png':''));if(!im)return null;
          if(im.video&&!rid[im.video]){const a='rId'+(rels.length+2),b='rId'+(rels.length+3);rid[im.video]=[a,b];
            rels.push(`<Relationship Id="${a}" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/${im.video}"/>`,`<Relationship Id="${b}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/${im.video}"/>`)}
          if(!rid[im.file]){rid[im.file]='rId'+(rels.length+2);rels.push(`<Relationship Id="${rid[im.file]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${im.file}"/>`)}return{rid:rid[im.file],w:im.w,h:im.h,video:im.video?{media:rid[im.video][0],link:rid[im.video][1],dur:im.dur}:null}};
        zip.file(`ppt/slides/slide${n}.xml`,morphSlideXML(s,T,ln,rel,mult));zip.file(`ppt/slides/_rels/slide${n}.xml.rels`,relsDoc(rels.join('')))}
      prog(1);continue}
    n++;const dur=ST.durOf(st,p.f);let media;const vids=await exportCopies(T);
    if(mode==='mp4'){const data=await renderMP4(p,T,cfg,dur,k=>prog(vids.length?k*.6:k),{vids});zip.file(`ppt/media/media${n}.mp4`,data);zip.file(`ppt/media/poster${n}.jpeg`,await posterOf(p,T,0,vids));media={type:'mp4',dur:dur+HOLD};
        }
    else{if(vids.length)await seekAll(vids,1);zip.file(`ppt/media/anim${n}.gif`,await renderGIF(p,T,dur,prog));media={type:'gif'}}
    /* every picture place holding a video gets the real video, placed where the picture ends up */
    media.over=[];for(const m of ST.mediaOf(p.style,p.f)){const live=T.images[m.img];if(!isVid(live))continue;
      const vf=await prepVideo(m.img,live);if(vf)media.over.push({...m,...vf,fit:window.IMXMorph.fitOf(live)})}
    if(vids.length)await seekAll(vids,.3);
    const measure=s=>ST.fitFor(meas,s,T);measure.color=s=>ST.colorFor(meas,s,T);
    zip.file(`ppt/slides/slide${n}.xml`,slideXML(p,T,media,ln,measure));zip.file(`ppt/slides/_rels/slide${n}.xml.rels`,slideRels(n,media))}
  const count=n,idx=[...Array(count).keys()];
  /* register the slides */
  let pres=await zip.file('ppt/presentation.xml').async('string');
  pres=pres.replace('</p:sldMasterIdLst>','</p:sldMasterIdLst><p:sldIdLst>'+idx.map(i=>`<p:sldId id="${256+i}" r:id="rIdS${i+1}"/>`).join('')+'</p:sldIdLst>');
  zip.file('ppt/presentation.xml',pres);
  let rels=await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  rels=rels.replace('</Relationships>',idx.map(i=>`<Relationship Id="rIdS${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join('')+'</Relationships>');
  zip.file('ppt/_rels/presentation.xml.rels',rels);
  let ct=await zip.file('[Content_Types].xml').async('string');
  const def=(ext,type)=>{if(!ct.includes(`Extension="${ext}"`))ct=ct.replace('<Default ',`<Default Extension="${ext}" ContentType="${type}"/><Default `)};
  def('mp4','video/mp4');def('gif','image/gif');if(videoExts.has('webm'))def('webm','video/webm');if(videoExts.has('mov'))def('mov','video/quicktime');if(videoExts.has('m4v'))def('m4v','video/x-m4v');def('jpeg','image/jpeg');def('png','image/png');
  ct=ct.replace('</Types>',idx.map(i=>`<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')+'</Types>');
  zip.file('[Content_Types].xml',ct);
  let pp=await zip.file('ppt/presProps.xml').async('string');
  if(!pp.includes('showMediaCtrls'))pp=pp.replace('<p:extLst>','<p:showPr showNarration="1"><p:sldAll/><p:penClr><a:prstClr val="red"/></p:penClr><p:extLst><p:ext uri="{2FDB2607-1784-4EEB-B798-7EB5836EED8A}"><p14:showMediaCtrls xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext></p:extLst></p:showPr><p:extLst>');
  zip.file('ppt/presProps.xml',pp);
  let core=await zip.file('docProps/core.xml').async('string');core=core.replace(/<dc:title>[^<]*<\/dc:title>|<dc:title\/>/,'').replace('</cp:coreProperties>',`<dc:title>${esc(title||'IMX')}</dc:title></cp:coreProperties>`);zip.file('docProps/core.xml',core);
  onProgress({slide:pages.length,total:pages.length,frac:1,mode,packing:true});
  return {blob:await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'STORE'}),mode:needVideo?mode:'shapes'}}
window.IMXExport={build,pickCodec};
})();
