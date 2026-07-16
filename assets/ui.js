'use strict';

const REQUIRED_PROGRESS_FIELDS = [
  'nama','tglLahir','tglPeriksa','riwayatKerjaId','riwayatSakitId','gejalaSubId','gejalaObjId',
  'tinggi','berat','lingkarPerut','tekananDarah','hb','rbc','got','gpt','ggtp','ldl','hdl',
  'trigliserida','gulaDarah','mataKanan','mataKiri','ekgId','rontgenTanggal','rontgenNo',
  'rontgenTemuanId','diagnosisId','keteranganId','tglDokumen','dokterNama','klinikNama'
];

function isCompletedField(id){
  const node=document.getElementById(id);
  if(!node) return false;
  if(node.type==='checkbox') return node.checked;
  return String(node.value||'').trim()!=='';
}

function updateFormProgress(){
  const completed=REQUIRED_PROGRESS_FIELDS.filter(isCompletedField).length;
  const total=REQUIRED_PROGRESS_FIELDS.length;
  const percent=Math.round((completed/total)*100);
  const progress=document.getElementById('formProgress');
  const percentNode=document.getElementById('progressPercent');
  const caption=document.getElementById('progressCaption');
  if(progress) progress.value=percent;
  if(percentNode) percentNode.textContent=`${percent}%`;
  if(caption){
    if(percent===100) caption.textContent='Semua data utama sudah terisi dan dokumen siap diperiksa.';
    else if(percent>=75) caption.textContent=`${completed} dari ${total} data utama terisi. Tinggal beberapa bagian.`;
    else if(percent>=40) caption.textContent=`${completed} dari ${total} data utama terisi. Lanjutkan pemeriksaan.`;
    else if(completed>0) caption.textContent=`${completed} dari ${total} data utama terisi.`;
    else caption.textContent='Belum ada data wajib yang terisi.';
  }
}

function setActivePageButton(targetId){
  document.querySelectorAll('.page-nav-btn').forEach(button=>{
    button.classList.toggle('active',button.dataset.pageTarget===targetId);
  });
}

document.querySelectorAll('.page-nav-btn').forEach(button=>{
  button.addEventListener('click',()=>{
    const target=document.getElementById(button.dataset.pageTarget);
    if(target){
      target.scrollIntoView({behavior:'smooth',block:'start'});
      setActivePageButton(button.dataset.pageTarget);
    }
  });
});

const previewPanel=document.querySelector('.preview-panel');
if(previewPanel && 'IntersectionObserver' in window){
  const observer=new IntersectionObserver(entries=>{
    const visible=entries.filter(entry=>entry.isIntersecting).sort((a,b)=>b.intersectionRatio-a.intersectionRatio)[0];
    if(visible) setActivePageButton(visible.target.id);
  },{root:previewPanel,threshold:[.2,.45,.7]});
  document.querySelectorAll('.page').forEach(page=>observer.observe(page));
}

document.addEventListener('input',event=>{
  if(event.target.matches('[data-field]')) updateFormProgress();
});
document.addEventListener('change',event=>{
  if(event.target.matches('[data-field]')) updateFormProgress();
});
['btnNew','btnSample','btnLoad','btnDelete'].forEach(id=>{
  const button=document.getElementById(id);
  if(button) button.addEventListener('click',()=>window.setTimeout(updateFormProgress,0));
});

updateFormProgress();
