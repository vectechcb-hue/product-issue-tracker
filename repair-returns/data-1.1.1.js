/* Storage and portable backups for repair reports. */
'use strict';
const Repair = (() => {
  const format = 'chenbom-repair-backup';
  const dbName = 'cbtrade-repair-returns-v1';
  const empty = () => ({headerText:'',items:[],totalAmount:'',editingHistoryId:null,editingHistoryTime:''});
  const id = () => crypto.randomUUID ? crypto.randomUUID() : Array.from(crypto.getRandomValues(new Uint8Array(16)),v=>v.toString(16).padStart(2,'0')).join('');
  const copy = v => JSON.parse(JSON.stringify(v));
  const sorted = v => Array.isArray(v) ? v.map(sorted) : v && typeof v === 'object' ? Object.fromEntries(Object.keys(v).sort().map(k=>[k,sorted(v[k])])) : v;
  const canonical = v => JSON.stringify(sorted(v));
  const equal = (a,b) => canonical(a)===canonical(b);
  async function digest(text) {
    const bytes = await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text));
    return Array.from(new Uint8Array(bytes),b=>b.toString(16).padStart(2,'0')).join('');
  }
  const counts = d => {
    const reports=[...(d.current?[d.current]:[]),...d.history];
    return {history:d.history.length,items:reports.reduce((n,r)=>n+r.items.length,0),images:reports.reduce((n,r)=>n+r.items.filter(x=>x.image).length,0)};
  };
  function validate(state) {
    if(!state || !Array.isArray(state.history) || !('current' in state)) throw Error('備份結構不完整');
    const ids=new Set();
    for(const h of state.history) {
      if(!h || typeof h.docId!=='string' || !h.docId || ids.has(h.docId)) throw Error('歷史紀錄識別碼缺漏或重複');
      ids.add(h.docId);
    }
    for(const r of [...(state.current===null?[]:[state.current]),...state.history]) {
      if(!r || typeof r.headerText!=='string' || typeof r.totalAmount!=='string' || !Array.isArray(r.items)) throw Error('表頭、金額或明細格式有誤');
      for(const i of r.items) if(!i || typeof i.id!=='string' || typeof i.details!=='string' || typeof i.qtyPrice!=='string' || typeof i.image!=='string' || ![true,false,null,undefined].includes(i.isReturned) || (i.returnedDate!=null && typeof i.returnedDate!=='string')) throw Error('明細欄位格式有誤');
    }
    return state;
  }
  async function backup(state) {
    validate(state);
    return {format,version:1,exportedAt:new Date().toISOString(),counts:counts(state),sha256:await digest(canonical(state)),data:state};
  }
  async function unpack(text) {
    const p=JSON.parse(text.replace(/^\uFEFF/,''));
    if(p.format!==format || p.version!==1) throw Error('無法辨識此備份版本');
    const d=validate(p.data);
    if(!equal(p.counts,counts(d)) || p.sha256!==await digest(canonical(d))) throw Error('備份的筆數或檢查碼不符');
    return d;
  }
  const body = r => { const {docId,...rest}=r; return rest; };
  function append(history, incoming) {
    const same=history.find(r=>r.docId===incoming.docId);
    if(!same) {history.push(copy(incoming));return incoming.docId;}
    if(equal(same,incoming)) return same.docId;
    const existing=history.find(r=>r.docId.startsWith(incoming.docId+'~') && equal(body(r),body(incoming)));
    if(existing) return existing.docId;
    const newId=incoming.docId+'~'+id();
    history.push({...copy(incoming),docId:newId});return newId;
  }
  const isBlank = r => !r || (!r.headerText && !r.items.length && !r.totalAmount && !r.editingHistoryId);
  const asDraft = r => {const {docId,timestamp,timestampMs,...fields}=r;return {...fields,editingHistoryId:docId,editingHistoryTime:timestamp||''};};
  const asReport = r => {const {editingHistoryId,editingHistoryTime,...v}=r;return v;};
  async function merge(local,incoming) {
    validate(incoming);
    const output=copy(local);
    const links=new Map(incoming.history.map(r=>[r.docId,append(output.history,r)]));
    if(incoming.current) {
      const c=copy(incoming.current);
      if(c.editingHistoryId && links.has(c.editingHistoryId)) c.editingHistoryId=links.get(c.editingHistoryId);
      if(isBlank(output.current)) output.current=c;
      else if(!isBlank(c) && !equal(output.current,c)) {
        const related=output.history.find(r=>r.docId===c.editingHistoryId && equal(asReport(asDraft(r)),asReport(c)));
        if(!related) {
          const key='draft-'+await digest(canonical(c));
          append(output.history,{...copy(c),docId:key,timestamp:'匯入的草稿',timestampMs:0});
        }
      }
    }
    return output;
  }
  function open() {
    return new Promise((resolve,reject)=>{
      const req=indexedDB.open(dbName,1);
      req.onupgradeneeded=()=>req.result.createObjectStore('reports');
      req.onsuccess=()=>resolve(req.result);
      req.onerror=()=>reject(req.error);
      req.onblocked=()=>reject(Error('請先關閉其他開啟的分頁'));
    });
  }
  function read(db,key='live') {
    return new Promise((resolve,reject)=>{
      const req=db.transaction('reports').objectStore('reports').get(key);
      req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
    });
  }
  function write(db,value,revision,checkpoint=false) {
    return new Promise((resolve,reject)=>{
      const tx=db.transaction('reports','readwrite'), s=tx.objectStore('reports');
      let error;
      tx.oncomplete=()=>resolve(revision+1);
      tx.onabort=()=>reject(error||tx.error||Error('儲存已中斷'));
      const request=s.get('live');
      request.onsuccess=()=>{
        const old=request.result;
        if((old?.revision||0)!==revision) {error=Error('另一個視窗已修改資料。請先匯出本頁備份，再重新開啟並合併。');tx.abort();return;}
        if(checkpoint && old) s.put(old,'previous');
        s.put({revision:revision+1,data:value},'live');
      };
    });
  }
  function status(item) {
    if (item.returnDestination==='factory') return '歸還工廠';
    return item.isReturned===true?'已返回公司':item.isReturned===false?'返修中':'待確認';
  }
  return {status,format,empty,id,copy,canonical,equal,digest,counts,validate,backup,unpack,merge,append,isBlank,asDraft,asReport,open,read,write};
})();
