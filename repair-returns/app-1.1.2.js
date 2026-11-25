'use strict';
const $=s=>document.querySelector(s);
const el=(tag,txt,cls)=>{const e=document.createElement(tag);if(txt!=null)e.textContent=txt;if(cls)e.className=cls;return e;};
const on=(s,name,fn)=>$(s).addEventListener(name,ev=>{Promise.resolve().then(()=>fn(ev)).catch(showError);});
const button=(text,action,cls)=>{const b=el('button',text,cls);b.type='button';b.onclick=()=>Promise.resolve().then(action).catch(showError);return b;};
let db,revision=0,data={current:Repair.empty(),history:[]},pending=Promise.resolve(),failed=null,ready=false,tab='editor',toastTimer,waiting=0;
const now=()=>new Date().toLocaleString('zh-TW',{hour12:false});
function toast(message){$('#toast').textContent=message;$('#toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('#toast').hidden=true,4000);}
function showError(err){$('#notice').hidden=false;$('#notice').textContent=err.message||String(err);}
function saveLabel(t){$('#saved').textContent=t;}
function lock(b){$('main').inert=b;$('nav').inert=b;}
function persistAutosave(){
  const value=Repair.copy(data);waiting++;saveLabel('儲存中…');
  pending=pending.then(async()=>{if(failed) return;try{revision=await Repair.write(db,value,revision);}catch(e){failed=e;showError(Error(e.message+'　可到「匯入與備份」匯出本頁資料。'));}}).finally(()=>{waiting--;saveLabel(failed?'尚未儲存':waiting?'儲存中…':'已存於本機');});
}
async function change(make,success,after,checkpoint=true){
  if(!ready||failed)throw Error('目前無法寫入，請先匯出本頁備份。');
  lock(true);
  try{
    await pending;if(failed)throw failed;
    const next=await make(Repair.copy(data));
    if(!next)return;
    Repair.validate(next);
    revision=await Repair.write(db,next,revision,checkpoint);
    data=next;saveLabel('已存於本機');render();
    if(after)after();if(success)toast(success);
  }finally{lock(false);}
}
function modal(title,content,actions=[]){
  const dlg=$('#dialog');if(dlg.open)dlg.close();
  $('#dialogTitle').textContent=title;
  $('#dialogContent').replaceChildren(typeof content==='string'?document.createTextNode(content):content);
  $('#dialogActions').replaceChildren(...actions.map(([name,fn,cls])=>button(name,()=>{dlg.close();return fn();},cls)));
  dlg.showModal();
}
const confirmAction=(title,text,label,fn)=>modal(title,text,[[label,fn,'primary'],['取消',()=>{}]]);
function go(name){
  tab=name;for(const id of ['editor','history','transfer'])$('#'+id).hidden=id!==name;
  document.querySelectorAll('[data-tab]').forEach(b=>{if(b.dataset.tab===name)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');});
  if(name==='history')historyView();else if(name==='transfer')inventory();
  window.scrollTo(0,0);
}
function progressText(items) {
 const count = label => items.filter(item => Repair.status(item) === label).length;
 return '已返回公司 '+count('已返回公司')+' / '+items.length+' 項　歸還工廠 '+count('歸還工廠')+' / '+items.length+' 項';
}
function progressView(items) {
 const block=el('div',null,'progress');
 block.append(el('strong',progressText(items)));
 const pending=items.filter(item=>Repair.status(item)==='返修中').length;
 const unknown=items.filter(item=>Repair.status(item)==='待確認').length;
 block.append(el('small','返修中 '+pending+' 項　·　待確認 '+unknown+' 項'));
 return block;
}
function render(){
  const d=data.current||Repair.empty();$('#heading').value=d.headerText;$('#total').value=d.totalAmount;
  $('#editInfo').hidden=!d.editingHistoryId;
  $('#editInfo').textContent=d.editingHistoryId?'編輯歷史紀錄　'+(d.editingHistoryTime||''):'';
  $('#saveRecord').textContent=d.editingHistoryId?'更新紀錄':'存為紀錄';$('#asNew').hidden=!d.editingHistoryId;
  $('#summary').replaceChildren(progressView(d.items));
  const rows=d.items.map((item,n)=>{
    const box=el('article',null,'row'),head=el('div',null,'rowhead');
    const state=Repair.status(item);
    const chip=button(state+(item.returnedDate?' · '+item.returnedDate:'')+' ▾',()=>chooseStatus(n),'chip '+({'已返回公司':'done','歸還工廠':'factory','待確認':'unknown'}[state]||''));
    chip.setAttribute('aria-label','品項 '+(n+1)+'：'+state+'，變更歸還狀態');
    head.append(el('h3','品項 '+(n+1)),chip);box.append(head);
    const grid=el('div',null,'rowgrid');
    for(const [key,label] of [['details','型號／品名／故障原因'],['qtyPrice','數量／金額']]){
      const lbl=el('label',label),t=el('textarea');t.rows=key==='details'?4:2;t.value=item[key];t.oninput=()=>{data.current.items[n][key]=t.value;persistAutosave();};lbl.append(t);grid.append(lbl);
    }
    box.append(grid);
    const ph=el('div',null,'rowbuttons photo');
    if(Word.picture(item.image)){const im=el('img');im.alt='品項 '+(n+1)+' 產品照片';im.src=item.image;ph.append(im);}
    const lab=el('label',item.image?'更換照片':'加入照片','btn'),file=el('input');file.type='file';file.accept='image/*';file.className='file';lab.append(file);ph.append(lab);
    file.onchange=async()=>{try{if(!file.files[0])return;const c=data.current,ref=c.items[n];const image=await imageData(file.files[0]);if(data.current!==c||data.current.items[n]!==ref)throw Error('明細已切換，請重新選擇照片');await change(x=>{x.current.items[n].image=image;return x;},'已儲存照片');}catch(e){showError(e);}};
    ph.append(button('移除品項',()=>confirmAction('移除品項','移除第 '+(n+1)+' 項？操作前的資料可在備份頁找回。','移除',()=>change(x=>{x.current.items.splice(n,1);return x;},'已移除品項')),'danger'));
    box.append(ph);return box;
  });$('#items').replaceChildren(...rows);historyView();inventory();
}
function chooseStatus(n){
 const item=data.current.items[n],div=el('div'),lab=el('label','歸還日期'),date=el('input');date.type='date';
 const today=new Date();const localDay=new Date(today.getTime()-today.getTimezoneOffset()*60000).toISOString().slice(0,10);
 const m=String(item.returnedDate||'').match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
 date.value=m?m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0'):(item.isReturned===true||item.returnDestination==='factory')?'':localDay;
 lab.append(date);div.append(lab,el('p','若舊單沒有記載日期，可清空日期後再儲存。'));
 const set=label=>change(x=>{
   const row=x.current.items[n];
   delete row.returnDestination;
   if(label==='歸還工廠') row.returnDestination='factory';
   row.isReturned=label==='待確認'?null:label==='已返回公司';
   row.returnedDate=(label==='已返回公司'||label==='歸還工廠')?date.value:'';
   return x;
 },'已更新歸還狀態');
 const options=['已返回公司','歸還工廠','返修中','待確認'];
 modal('品項 '+(n+1)+' · 歸還狀態',div,options.map(label=>[label,()=>set(label),label===Repair.status(item)?'soft':'']));
}
function historyView(){
 const q=$('#search').value.trim().toLowerCase(), list=data.history.filter(r=>[r.headerText,r.timestamp,r.sourceName,...r.items.map(i=>i.details)].join(' ').toLowerCase().includes(q)).sort((a,b)=>(b.timestampMs||0)-(a.timestampMs||0));
 $('#historyCount').textContent=data.history.length+' 份';const target=$('#historyList');target.replaceChildren();
 if(!list.length)target.append(el('p',q?'沒有符合的紀錄':'目前沒有歷史紀錄。可從明細頁儲存，或在匯入頁選取舊 Word 檔。','empty'));
 for(const r of list){
   const card=el('article',null,'record');card.append(el('h3',r.headerText||r.sourceName||'未命名明細'),el('p',[(r.timestamp||''),r.items.length+' 項',r.items.filter(i=>i.image).length+' 張照片',r.totalAmount].filter(Boolean).join(' · ')));
   card.append(progressView(r.items));
   if(r.sourceName)card.append(el('p',r.sourceName,'small'));
   const actions=el('div',null,'actions');
   actions.append(button('開啟明細',()=>openRecord(r)),button('刪除',()=>confirmAction('刪除歷史紀錄',r.headerText||r.sourceName||'這份紀錄','刪除紀錄',()=>change(x=>{x.history=x.history.filter(h=>h.docId!==r.docId);if(x.current.editingHistoryId===r.docId){x.current.editingHistoryId=null;x.current.editingHistoryTime='';}return x;},'已刪除紀錄')),'danger'));card.append(actions);target.append(card);
 }
}
function inventory(){const c=Repair.counts(data);$('#inventory').textContent='目前有 '+c.history+' 份歷史紀錄。草稿與歷史合計 '+c.items+' 個品項、'+c.images+' 張照片。';}
function archiveDraft(x){
 const d=x.current;if(Repair.isBlank(d))return;
 if(x.history.some(h=>Repair.equal(Repair.asReport(Repair.asDraft(h)),Repair.asReport(d))))return;
 x.history.push({...Repair.asReport(d),docId:Repair.id(),timestamp:'草稿留存 · '+now(),timestampMs:Date.now()});
}
function openRecord(r){return change(x=>{archiveDraft(x);x.current=Repair.asDraft(r);return x;},'已開啟明細 · '+progressText(r.items),()=>go('editor'));}
function saveRecord(force){
 return change(x=>{if(Repair.isBlank(x.current))throw Error('請先輸入表頭或加入品項');const d=x.current,docId=(!force&&d.editingHistoryId)||Repair.id(),index=x.history.findIndex(h=>h.docId===docId);
 const r={...(index>=0?x.history[index]:{}),...Repair.asReport(d),docId,timestamp:now(),timestampMs:Date.now()};
 if(index<0)x.history.push(r);else x.history[index]=r;d.editingHistoryId=docId;d.editingHistoryTime=r.timestamp;return x;},'已存入歷史紀錄');
}
async function imageData(file){const dataUrl=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=()=>rej(r.error);r.readAsDataURL(file);});
 if(Word.picture(dataUrl))return dataUrl;
 throw Error('此圖片格式無法預覽。請另存 JPEG 或 PNG 再選取。');}
function preparedFile(name,contents,type){
 const file=new File([contents],name,{type});const target=el('div');target.append(el('p','請將檔案儲存到「檔案」或 iCloud Drive。'));
 if(navigator.canShare?.({files:[file]}))target.append(button('分享／儲存到檔案',async()=>{try{await navigator.share({files:[file]});}catch(e){if(e.name!=='AbortError')throw e;}},'primary'));
 const link=el('a','下載檔案','btn');const url=URL.createObjectURL(file);link.href=url;link.download=name;target.append(link);
 modal(name,target);$('#dialog').addEventListener('close',()=>setTimeout(()=>URL.revokeObjectURL(url),60000),{once:true});
}
const stamp=()=>{const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')+'_'+[d.getHours(),d.getMinutes(),d.getSeconds()].map(v=>String(v).padStart(2,'0')).join('');};
function previewBackup(incoming){
 const c=Repair.counts(incoming);
 confirmAction('合併完整備份','檔案內含 '+c.history+' 份歷史紀錄、'+c.items+' 個品項與 '+c.images+' 張照片（含草稿）。目前明細會保留；內容不同的匯入草稿會另存為紀錄。','確認合併',()=>change(x=>Repair.merge(x,incoming),'備份已合併',()=>go('history')));
}
async function wordPreview(files,parsed=null) {
 const wrap=el('div'),results=parsed||[];lock(true);
 try{if(!parsed){toast('正在讀取 '+files.length+' 份 Word…');for(const f of files){try{results.push(await Word.parse(f));}catch(e){results.push({name:f.name,error:e.message});}}}}finally{lock(false);}
 const keys=new Set(data.history.map(h=>h.sourceWordKey||h.docId));const keep=[];
 for(const r of results){
   const row=el('div',null,'preview');row.append(el('strong',r.name));
   if(r.error)row.append(el('p','無法匯入 · '+r.error,'warn'));
   else{
     const d=r.report;
     const prior=data.history.find(h=>h.docId===r.key);
     let correction='';
     if(prior && r.legacyItems && !keys.has(r.key+'-grouped') && !prior.wordRowsGrouped && prior.items.length===r.legacyItems.length){
       if(Repair.equal(prior.items,r.legacyItems)){
         r.repairId=prior.docId;
         correction='修正原紀錄：'+prior.items.length+' 列合併為 ';
       }else{
         correction='原紀錄已有編輯，保留原紀錄並另存整理版：';
       }
       r.originalKey=r.key;
       r.key+='-grouped';
     }
     const dup=keys.has(r.key);keys.add(r.key);
     row.append(el('p',(dup?'已匯入，略過　':correction)+d.items.length+' 個品項 · '+d.items.filter(i=>i.image).length+' 張照片 · '+(d.totalAmount||'未載明金額')));
     if(d.headerText)row.append(el('p',d.headerText,'small'));
     if(r.warnings.length)row.append(el('p',r.warnings.join('；'),'warn'));
     if(!dup)keep.push(r);
   }wrap.append(row);
 }
 const actions=keep.length?[['匯入 '+keep.length+' 份紀錄',()=>change(x=>{
   for(const r of keep){
     if(r.repairId){
       const old=x.history.find(h=>h.docId===r.repairId);
       if(!old || !Repair.equal(old.items,r.legacyItems))throw Error('紀錄已變更，請重新選取檔案核對。');
       old.items=Repair.copy(r.report.items);old.wordRowsGrouped=true;
       if(x.current?.editingHistoryId===old.docId && Repair.equal(x.current.items,r.legacyItems))x.current.items=Repair.copy(old.items);
       continue;
     }
     x.history.push({...r.report,docId:r.key,sourceWordKey:r.key,sourceName:r.name,timestamp:'匯入 '+now(),timestampMs:Date.now()});}return x;
 },'已匯入 '+keep.length+' 份，原草稿已保留',()=>go('history')),'primary']]:[];
 actions.push(['關閉',()=>{}]);modal('匯入前核對',wrap,actions);
}
on('#closeDialog','click',()=>$('#dialog').close());
document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>go(b.dataset.tab));
for(const [sel,key] of [['#heading','headerText'],['#total','totalAmount']])on(sel,'input',()=>{data.current[key]=$(sel).value;persistAutosave();});
on('#saveRecord','click',()=>saveRecord(false));on('#asNew','click',()=>saveRecord(true));
on('#newDraft','click',()=>confirmAction('開新明細','目前有內容的草稿將保留在歷史紀錄。','開新明細',()=>change(x=>{archiveDraft(x);x.current=Repair.empty();return x;},'已開新明細')));
on('#add','click',()=>change(x=>{x.current.items.push({id:Repair.id(),image:'',details:'',qtyPrice:'',isReturned:false,returnedDate:''});return x;},'',()=>$('#items article:last-child textarea')?.focus(),false));
on('#search','input',historyView);
on('#wordFiles','change',async ev=>{const files=Array.from(ev.target.files);ev.target.value='';if(files.length)await wordPreview(files);});
on('#backupFile','change',async ev=>{const f=ev.target.files[0];ev.target.value='';if(f)previewBackup(await Repair.unpack(await f.text()));});
on('#recover','click',async()=>{const prev=await Repair.read(db,'previous');if(prev)previewBackup(prev.data);else toast('目前沒有操作前備份');});
on('#exportBackup','click',async()=>{const pack=await Repair.backup(Repair.copy(data));preparedFile('返修完整備份_'+stamp()+'.json',JSON.stringify(pack),'application/json');});
on('#exportWord','click',()=>preparedFile('返修明細_'+stamp()+'.doc','\ufeff'+Word.render(data.current),'application/msword'));
on('#pasteBtn','click',()=>{const div=el('div'),area=el('div');div.append(el('p','請從 Word 複製整張明細後，在下方長按並選「貼上」。'));area.contentEditable='true';area.setAttribute('aria-label','貼上 Word 表格');area.setAttribute('role','textbox');div.append(area);
 area.onpaste=async e=>{e.preventDefault();try{const html=e.clipboardData.getData('text/html');if(!html)throw Error('剪貼簿沒有 Word 表格，請改用選取檔案');const r=Word.fromHTML(html);await wordPreview([], [{...r,key:'paste-'+await Repair.digest(Repair.canonical(r.report)),name:'貼上的 Word 明細'}]);}catch(err){area.textContent=err.message;}};modal('貼上 Word 表格',div);area.focus();});
(async()=>{
 lock(true);
 try{db=await Repair.open();const value=await Repair.read(db);if(value){Repair.validate(value.data);revision=value.revision;data=value.data;}
 if(!data.current)data.current=Repair.empty();ready=true;render();saveLabel('已存於本機');lock(false);}
 catch(e){saveLabel('無法讀取');showError(Error('無法讀取這台裝置的資料：'+e.message+'。請確認可使用網站儲存空間後重新開啟。'));}
 if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js',{scope:'./'}).catch(()=>showError(Error('離線安裝尚未完成，請連網重新開啟。')));
})();
