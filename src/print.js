// Freeze the visible dashboard before print CSS or the printer can reflow it.
const PRINT_PAPER={width:281*96/25.4,height:194*96/25.4};
function printFit(width,height){
 if(!(width>0&&height>0))throw new Error('Нет видимого баланса для печати.');
 const scale=Math.min(PRINT_PAPER.width/width,PRINT_PAPER.height/height);
 return {scale,left:(PRINT_PAPER.width-width*scale)/2,top:0};
}
function closePrintPreview(){
 document.getElementById('printSnapshot')?.remove();
 document.body.classList.remove('snapshot-open');
 $('printReport').focus();
}
function openPrintPreview(){
 closePrintPreview();
 const source=$('dashboard'),box=source.getBoundingClientRect(),fit=printFit(box.width,box.height);
 const snapshot=source.cloneNode(true),originals=[source,...source.querySelectorAll('*')],copies=[snapshot,...snapshot.querySelectorAll('*')];
 for(let i=0;i<originals.length;i++){
  const style=getComputedStyle(originals[i]),copy=copies[i];
  copy.removeAttribute('id');
  for(const prop of style)copy.style.setProperty(prop,style.getPropertyValue(prop),'important');
  copy.style.setProperty('animation','none','important');
  copy.style.setProperty('transition','none','important');
  copy.style.setProperty('container-type','normal','important');
  copy.style.setProperty('contain','none','important');
  copy.style.setProperty('pointer-events','none','important');
 }
 snapshot.classList.add('print-snapshot-stage');
 for(const [prop,value] of Object.entries({position:'absolute',left:fit.left+'px',top:fit.top+'px',margin:'0',width:box.width+'px',height:box.height+'px',transform:'scale('+fit.scale+')','transform-origin':'top left'}))snapshot.style.setProperty(prop,value,'important');
 const preview=document.createElement('section');preview.id='printSnapshot';preview.setAttribute('role','dialog');preview.setAttribute('aria-modal','true');preview.setAttribute('aria-label','Предпросмотр печати');
 const previewScale=Math.min(1,(window.innerWidth-32)/PRINT_PAPER.width,(window.innerHeight-92)/PRINT_PAPER.height);
 preview.style.setProperty('--preview-scale',previewScale);
 preview.style.setProperty('--preview-width',PRINT_PAPER.width*previewScale+'px');
 preview.style.setProperty('--preview-height',PRINT_PAPER.height*previewScale+'px');
 preview.innerHTML='<header class="print-preview-toolbar"><strong>Предпросмотр печати · A4 альбомный</strong><div><button class="btn primary" id="confirmSnapshotPrint">Печать / PDF</button><button class="btn" id="closeSnapshotPrint">Закрыть</button></div></header><div class="print-page-frame"><div class="print-paper"></div></div>';
 preview.querySelector('.print-paper').append(snapshot);document.body.append(preview);document.body.classList.add('snapshot-open');
 // Keep the same visible part of any scrolled table in the frozen copy.
 for(let i=0;i<originals.length;i++){copies[i].scrollLeft=originals[i].scrollLeft;copies[i].scrollTop=originals[i].scrollTop;}
 $('confirmSnapshotPrint').onclick=()=>window.print();$('closeSnapshotPrint').onclick=closePrintPreview;
 $('confirmSnapshotPrint').focus();
}
window.addEventListener?.('beforeprint',()=>{if(!document.getElementById('printSnapshot')&&book)openPrintPreview();});
document.addEventListener('keydown',e=>{
 if(document.getElementById('printSnapshot')){
  if(e.key==='Escape'){e.preventDefault();closePrintPreview();return;}
  if(e.key==='Tab'){e.preventDefault();const print=$('confirmSnapshotPrint'),close=$('closeSnapshotPrint');(document.activeElement===print?close:print).focus();return;}
 }
 if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='p'&&book){e.preventDefault();openPrintPreview();}
});
