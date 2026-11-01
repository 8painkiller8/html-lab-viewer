'use strict';
const $=id=>document.getElementById(id), picker=$('file'), viewer=$('viewer');
const KEY='htmlLabViewer.v2';
const WIDTHS=[960,1120,1280,1440,1600,1920];
// Modes: 'fill' = page laid out at pageWidth and stretched to fill the box exactly
// (the page sees a window with the box's shape); 'fit' = whole page shrunk to fit;
// 'manual' = fixed zoom with scrolling.
let selected=null, mode='fill', zoom=1, pageWidth=1280, height=900, revision=0, settings=null;
let dirty=false, saving=false, pending=false, saveTimer;
let token='';
function message(text){$('status').textContent=text;}
function layout(){
 const area=$('viewport'), aw=area.clientWidth, ah=area.clientHeight, w=pageWidth;
 let h=height, scale;
 if(mode==='fill'){scale=Math.max(.01,aw/w);h=Math.max(200,Math.floor(ah/scale));}
 else scale=mode==='fit'?Math.min(aw/w,ah/h):zoom;
 const safe=Math.max(.01,scale);
 viewer.style.width=w+'px';viewer.style.height=h+'px';
 viewer.style.transform='scale('+safe+')';
 $('surface').style.width=w*safe+'px';$('surface').style.height=h*safe+'px';
 area.style.overflow=mode==='fill'?'hidden':'auto';
 $('zoom').textContent=Math.round(safe*100)+'%';
 $('fit').setAttribute('aria-pressed',String(mode==='fit'));
 $('fill').setAttribute('aria-pressed',String(mode==='fill'));
 $('width').value=String(w);
 return safe;
}
function payload(){return {version:2,file:selected,mode,zoom,pageWidth};}
function persist(){
 clearTimeout(saveTimer);
 if(!settings){message('Preview only: saving is available when this viewer runs inside PowerPoint.');return;}
 if(saving){pending=true;return;}
 const data=payload();
 if(JSON.stringify(data).length*2>1500000){message('This HTML is too large to embed here. It is loaded for this session only; choose a smaller file to save it.');return;}
 saving=true;pending=false;message('Saving HTML and view settings into the presentation…');
 try{settings.set(KEY,data);settings.saveAsync(result=>{
  saving=false;
  if(result.status==='succeeded'){message('Stored in the presentation. Save the PowerPoint file (⌘S) to keep this copy.');}
  else message('Could not save in PowerPoint: '+(result.error?.message||'Please try Save in presentation again.'));
  if(pending)persist();
 });}catch(e){saving=false;message('Could not save in PowerPoint: '+e.message);}
}
function changed(){dirty=true;clearTimeout(saveTimer);saveTimer=setTimeout(persist,400);}
function display(){
 height=900;growth=0;token=String(Math.random()).slice(2)+Date.now();
 $('filename').textContent=selected?.name||'No file selected';
 if(!selected){viewer.srcdoc='';layout();return;}
 // Measure the isolated page without granting it access to the parent or Office APIs.
 const doc=new DOMParser().parseFromString(selected.html,'text/html');
 const script=doc.createElement('script');
 script.textContent=`(()=>{let timer;function measure(){clearTimeout(timer);timer=setTimeout(()=>{parent.postMessage({type:'lab-height',token:${JSON.stringify(token)},height:Math.max(document.body?document.body.scrollHeight:0,document.documentElement.scrollHeight)},'*')},80)}addEventListener('load',()=>{measure();new ResizeObserver(measure).observe(document.body);if(document.fonts)document.fonts.ready.then(measure)});})();`;
 doc.head.appendChild(script);
 viewer.srcdoc='<!DOCTYPE html>\n'+doc.documentElement.outerHTML;
 layout();
}
// A page that sizes itself to the frame can report a slightly larger height each
// time the frame grows. Stop after a few small growth steps in a row so the frame
// can never keep enlarging (which shows as the page zooming out by itself).
let growth=0;
window.addEventListener('message',event=>{
 if(event.source!==viewer.contentWindow||event.data?.type!=='lab-height'||event.data.token!==token)return;
 const h=Number(event.data.height);
 if(!Number.isFinite(h)||h<900||h>30000||Math.abs(h-height)<=2)return;
 if(h>height&&h-height<80){if(++growth>3)return;}else growth=0;
 height=h;layout();
});
$('choose').addEventListener('click',()=>{picker.value='';picker.click();});
picker.addEventListener('change',async()=>{
 const file=picker.files[0];if(!file)return;const current=++revision;
 if(!/\.html?$/i.test(file.name)){message('Please select an .html or .htm file.');return;}
 if(file.size>25*1024*1024){message('Please choose a file smaller than 25 MB.');return;}
 try{const html=await file.text();if(current!==revision)return;
 selected={name:file.name,html};dirty=true;display();persist();
 }catch(e){message('Could not read the selected file. Please try again.');}
});
$('reset').addEventListener('click',()=>{++revision;selected=null;picker.value='';dirty=true;display();persist();});
$('save').addEventListener('click',persist);
function adjust(factor){
 if(mode==='fill'){
  // In Fill window, larger = a narrower page stretched over the same box
  const i=WIDTHS.indexOf(pageWidth), next=i+(factor>1?-1:1);
  if(next>=0&&next<WIDTHS.length){pageWidth=WIDTHS[next];layout();changed();}
  return;
 }
 zoom=Math.max(.1,Math.min(3,layout()*factor));mode='manual';layout();changed();
}
$('in').addEventListener('click',()=>adjust(1.1));$('out').addEventListener('click',()=>adjust(1/1.1));
$('fit').addEventListener('click',()=>{mode='fit';layout();changed();});
$('fill').addEventListener('click',()=>{mode='fill';layout();changed();});
$('width').addEventListener('change',e=>{pageWidth=Number(e.target.value)||1280;layout();changed();});
$('hide').addEventListener('click',()=>{document.body.classList.add('clean');layout();});
$('show').addEventListener('click',()=>{document.body.classList.remove('clean');layout();});
new ResizeObserver(layout).observe($('viewport'));
function initializeOffice(){
 if(!window.Office){message('Choose an HTML file to preview. Saving requires PowerPoint.');return;}
 Office.onReady(info=>{
  if(!info.host||!Office.context?.document?.settings)return;
  settings=Office.context.document.settings;
  if(dirty){persist();return;}
  try{const stored=settings.get(KEY);
   if(stored?.version===2){
    if(stored.file&&typeof stored.file.html==='string'&&typeof stored.file.name==='string')selected=stored.file;
    mode=['manual','fit','fill'].includes(stored.mode)?stored.mode:'fit';zoom=Math.max(.1,Math.min(3,Number(stored.zoom)||1));
    pageWidth=WIDTHS.includes(Number(stored.pageWidth))?Number(stored.pageWidth):1280;display();
    message(selected?'Restored the HTML saved in this presentation.':'Choose an HTML file to save with this presentation.');
   }else message('Choose an HTML file. A copy will be stored in this presentation; save PowerPoint afterward.');
  }catch(e){message('Could not restore saved HTML. Please choose the file again.');}
 });
}
window.addEventListener('load',initializeOffice);
display();
