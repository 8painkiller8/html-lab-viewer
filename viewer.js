'use strict';
const $=id=>document.getElementById(id), picker=$('file'), viewer=$('viewer');
const KEY='htmlLabViewer.v2';
let selected=null, mode='fit', zoom=1, height=900, revision=0, settings=null;
let dirty=false, saving=false, pending=false, saveTimer;
let token='';
function message(text){$('status').textContent=text;}
function layout(){
 const area=$('viewport');
 const scale=mode==='fit'?Math.min(area.clientWidth/1280,area.clientHeight/height):zoom;
 const safe=Math.max(.01,scale);
 viewer.style.width='1280px';viewer.style.height=height+'px';
 viewer.style.transform='scale('+safe+')';
 $('surface').style.width=1280*safe+'px';$('surface').style.height=height*safe+'px';
 $('zoom').textContent=Math.round(safe*100)+'%';
 $('fit').setAttribute('aria-pressed',String(mode==='fit'));
 return safe;
}
function payload(){return {version:2,file:selected,mode,zoom};}
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
 height=900;token=String(Math.random()).slice(2)+Date.now();
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
window.addEventListener('message',event=>{
 if(event.source!==viewer.contentWindow||event.data?.type!=='lab-height'||event.data.token!==token)return;
 const h=Number(event.data.height);
 if(Number.isFinite(h)&&h>=900&&h<=30000&&Math.abs(h-height)>2){height=h;layout();}
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
function adjust(factor){zoom=Math.max(.1,Math.min(3,layout()*factor));mode='manual';layout();changed();}
$('in').addEventListener('click',()=>adjust(1.1));$('out').addEventListener('click',()=>adjust(1/1.1));
$('fit').addEventListener('click',()=>{mode='fit';layout();changed();});
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
    mode=stored.mode==='manual'?'manual':'fit';zoom=Math.max(.1,Math.min(3,Number(stored.zoom)||1));display();
    message(selected?'Restored the HTML saved in this presentation.':'Choose an HTML file to save with this presentation.');
   }else message('Choose an HTML file. A copy will be stored in this presentation; save PowerPoint afterward.');
  }catch(e){message('Could not restore saved HTML. Please choose the file again.');}
 });
}
window.addEventListener('load',initializeOffice);
display();
