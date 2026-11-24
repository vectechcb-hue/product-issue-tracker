/* 在舊系統的同一個專案、同一個登入環境，暫時用本元件作為 App。
   請先保留原 App.tsx 並等候舊版草稿儲存完畢。匯出完可還原原始程式。
   本元件只讀取原資料，不建立匿名帳號，也不寫入 Firestore。 */
import React, { useState } from 'react';
import { getApps, initializeApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signInWithCustomToken } from 'firebase/auth';
import { collection, doc, getDocFromServer, getDocsFromServer, getFirestore } from 'firebase/firestore';

declare const __firebase_config: string;
declare const __app_id: string;
declare const __initial_auth_token: string;
function canonical(v: any): any {
  if (Array.isArray(v)) return v.map(canonical);
  if (v && typeof v === 'object') return Object.fromEntries(Object.keys(v).sort().map(k => [k,canonical(v[k])]));
  return v;
}
export default function App() {
  const [status, info] = useState('請在仍可看到舊紀錄的原專案與帳號內開啟。');
  const [busy, setBusy] = useState(false);
  const [download, setDownload] = useState('');
  async function load() {
    setBusy(true);
    if (download) URL.revokeObjectURL(download);
    setDownload('');
    try {
      if (typeof __firebase_config === 'undefined') throw Error('請使用舊系統原本的執行環境。');
      const config = JSON.parse(__firebase_config);
      const app = getApps().find(a=>a.name==='[DEFAULT]') || initializeApp(config);
      if (app.options.projectId !== config.projectId) throw Error('Firebase 專案與原本設定不一致。');
      const auth = getAuth(app);
      await new Promise<void>((resolve,reject)=>{
        const stop=onAuthStateChanged(auth,()=>{stop();resolve();},reject);
      });
      if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) await signInWithCustomToken(auth,__initial_auth_token);
      const user = auth.currentUser;
      if (!user) throw Error('找不到舊版登入身分，請回到仍可看到正式紀錄的原裝置與帳號。');
      const db = getFirestore(app);
      const appId = typeof __app_id !== 'undefined' ? __app_id : 'default-app-id';
      const [draft, records] = await Promise.all([
        getDocFromServer(doc(db,'artifacts',appId,'users',user.uid,'current','state')),
        getDocsFromServer(collection(db,'artifacts',appId,'users',user.uid,'history'))
      ]);
      if (!draft.exists() && records.empty) throw Error('這個身分沒有草稿或歷史紀錄，請核對原帳號。');
      const current = draft.exists()?draft.data():null;
      const history = records.docs.map(r=>({...r.data(),docId:r.id}));
      const data = {current,history};
      const reports = [...(current?[current]:[]),...history] as {items:any[]}[];
      if (reports.some(r=>!Array.isArray(r.items))) throw Error('原資料包含未知結構，請保留原專案並另行核對。');
      const counts = {history:history.length,items:reports.reduce((n,r)=>n+r.items.length,0),images:reports.reduce((n,r)=>n+r.items.filter(i=>i.image).length,0)};
      const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(canonical(data))));
      const sha256 = Array.from(new Uint8Array(hash),v=>v.toString(16).padStart(2,'0')).join('');
      const result = {format:'chenbom-repair-backup',version:1,exportedAt:new Date().toISOString(),counts,sha256,data};
      setDownload(URL.createObjectURL(new Blob([JSON.stringify(result)],{type:'application/json'})));
      info('草稿 '+(current?'1':'0')+' 份、歷史紀錄 '+counts.history+' 份，共 '+counts.items+' 個品項、'+counts.images+' 張照片（含草稿）。請與舊版畫面核對後下載。');
    } catch (err) { info(err instanceof Error ? err.message : String(err)); } finally { setBusy(false); }
  }
  return <main style={{maxWidth:550,margin:'48px auto',padding:16,fontFamily:'system-ui'}}>
    <h1>舊系統完整備份</h1><p>讀取原本的草稿與全部歷史紀錄，保留照片、歸還日期與紀錄時間。請先暫停編輯。</p>
    <button onClick={load} disabled={busy}>{busy?'讀取中…':'從原系統讀取'}</button><p role="status">{status}</p>
    {download && <a href={download} download={'返修完整備份_'+new Date().toISOString().slice(0,10)+'.json'}>下載完整備份 JSON</a>}
  </main>;
}
