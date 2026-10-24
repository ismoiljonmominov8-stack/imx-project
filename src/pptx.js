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
function frameCanvas(w,h){const c=document.createElement('canvas');c.width=w;c.height=h;return c}
async function posterOf(page,T){const c=frameCanvas(1280,720),x=c.getContext('2d');x.setTransform(.8,0,0,.8,0,0);
  ST.render(x,page.style,page.f,0,T,{text:false});return new Uint8Array(await (await new Promise(r=>c.toBlob(r,'image/jpeg',.88))).arrayBuffer())}
async function renderMP4(page,T,cfg,dur,progress){
  const c=frameCanvas(1280,720),x=c.getContext('2d',{willReadFrequently:false});x.setTransform(.8,0,0,.8,0,0);
  const target=new Mp4Muxer.ArrayBufferTarget();
  const muxer=new Mp4Muxer.Muxer({target,video:{codec:'avc',width:1280,height:720},fastStart:'in-memory'});
  let err=null;const enc=new VideoEncoder({output:(chunk,meta)=>muxer.addVideoChunk(chunk,meta),error:e=>{err=e}});enc.configure(cfg);
  const n=Math.ceil((dur+HOLD)*FPS);
  for(let i=0;i<n;i++){if(err)throw err;
    x.setTransform(.8,0,0,.8,0,0);ST.render(x,page.style,page.f,Math.min(i/FPS,dur+HOLD),T,{text:false});
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
    ?`<p:pic><p:nvPicPr><p:cNvPr id="${vid}" name="Animation"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rId3"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId2"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    :`<p:pic><p:nvPicPr><p:cNvPr id="${vid}" name="Animation"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  shapes.push(picRels);
  const slots=ST.slotsOf(page.style,page.f);
  slots.forEach(s=>{const sid=spid++,fit=measure(s),color=measure.color(s);shapes.push(textShape(sid,s,fit,color,T,lang));effects.push([sid,s.at])});
  /* timing: the video plays from the start; each text box enters at its own moment */
  id.n=5;let inner='';
  if(media.type==='mp4'){const a=id.n++,b=id.n++;
    inner+=`<p:par><p:cTn id="${a}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${b}" dur="${Math.round(media.dur*1000)}" fill="hold"/><p:tgtEl><p:spTgt spid="${vid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`}
  effects.forEach(([sid,at],i)=>{inner+=effectXML(T.effect||'fade',sid,at,T.speed,media.type!=='mp4'&&i===0,id)});
  let timing='';
  if(inner){const mediaNode=media.type==='mp4'?`<p:video><p:cMediaNode vol="80000" mute="1"><p:cTn id="${id.n++}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${vid}"/></p:tgtEl></p:cMediaNode></p:video>`:'';
    timing=`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${mediaNode}</p:childTnLst></p:cTn></p:par></p:tnLst>${effects.length?`<p:bldLst>${effects.map(([sid])=>`<p:bldP spid="${sid}" grpId="0"/>`).join('')}</p:bldLst>`:''}</p:timing>`}
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex(ST.mix(T.bg,'#000000',.6))}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition spd="med"><p:fade/></p:transition>${timing}</p:sld>`}
function slideRels(i,media){
  const L='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>';
  const body=media.type==='mp4'
    ?`<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/media${i}.mp4"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media${i}.mp4"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/poster${i}.jpeg"/>`
    :`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/anim${i}.gif"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${L}${body}</Relationships>`}

/* ---------- the whole file ---------- */
async function build({pages,themeFor,lang,template,title},onProgress=()=>{}){
  const zip=await JSZip.loadAsync(template),cfg=await pickCodec(),mode=cfg?'mp4':'gif';
  const meas=document.createElement('canvas').getContext('2d');const ln=LANG[lang]||'uz-Latn-UZ';
  for(let i=0;i<pages.length;i++){const p=pages[i],T=themeFor(p),n=i+1,dur=ST.durOf(ST.S[p.style],p.f);
    const prog=k=>onProgress({slide:n,total:pages.length,frac:(i+k)/pages.length,mode});prog(0);
    let media;
    if(mode==='mp4'){const data=await renderMP4(p,T,cfg,dur,prog);zip.file(`ppt/media/media${n}.mp4`,data);zip.file(`ppt/media/poster${n}.jpeg`,await posterOf(p,T));media={type:'mp4',dur:dur+HOLD}}
    else{zip.file(`ppt/media/anim${n}.gif`,await renderGIF(p,T,dur,prog));media={type:'gif'}}
    const measure=s=>ST.fitFor(meas,s,T);measure.color=s=>ST.colorFor(meas,s,T);
    zip.file(`ppt/slides/slide${n}.xml`,slideXML(p,T,media,ln,measure));zip.file(`ppt/slides/_rels/slide${n}.xml.rels`,slideRels(n,media))}
  /* register the slides */
  let pres=await zip.file('ppt/presentation.xml').async('string');
  pres=pres.replace('</p:sldMasterIdLst>','</p:sldMasterIdLst><p:sldIdLst>'+pages.map((_,i)=>`<p:sldId id="${256+i}" r:id="rIdS${i+1}"/>`).join('')+'</p:sldIdLst>');
  zip.file('ppt/presentation.xml',pres);
  let rels=await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  rels=rels.replace('</Relationships>',pages.map((_,i)=>`<Relationship Id="rIdS${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join('')+'</Relationships>');
  zip.file('ppt/_rels/presentation.xml.rels',rels);
  let ct=await zip.file('[Content_Types].xml').async('string');
  const def=(ext,type)=>{if(!ct.includes(`Extension="${ext}"`))ct=ct.replace('<Default ',`<Default Extension="${ext}" ContentType="${type}"/><Default `)};
  def('mp4','video/mp4');def('gif','image/gif');def('jpeg','image/jpeg');
  ct=ct.replace('</Types>',pages.map((_,i)=>`<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')+'</Types>');
  zip.file('[Content_Types].xml',ct);
  let core=await zip.file('docProps/core.xml').async('string');core=core.replace(/<dc:title>[^<]*<\/dc:title>|<dc:title\/>/,'').replace('</cp:coreProperties>',`<dc:title>${esc(title||'IMX')}</dc:title></cp:coreProperties>`);zip.file('docProps/core.xml',core);
  onProgress({slide:pages.length,total:pages.length,frac:1,mode,packing:true});
  return {blob:await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'STORE'}),mode}}
window.IMXExport={build,pickCodec};
})();
