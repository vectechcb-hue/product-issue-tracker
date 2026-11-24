'use strict';
const Word = (() => {
  const esc = v => String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const line = v=>esc(v).replace(/\n/g,'<br>');
  const picture = v=>/^data:image\/(?:png|jpeg|jpg|gif|webp|bmp|avif);base64,/i.test(v||'');
  function textOf(el) {
    if(!el) return '';
    let out='';
    function walk(n) {
      if(n.nodeType===3) {out+=n.nodeValue;return;}
      if(n.nodeType!==1)return;
      if(n.tagName==='BR'){out+='\n';return;}
      const block=['P','DIV','LI'].includes(n.tagName);
      if(block && out && !out.endsWith('\n'))out+='\n';
      n.childNodes.forEach(walk);
      if(block && !out.endsWith('\n'))out+='\n';
    }
    walk(el);return out.replace(/\n{3,}/g,'\n\n').trim();
  }
  function fromHTML(html) {
    // A template is inert, including external image references from the clipboard.
    const box=document.createElement('template');box.innerHTML=html;
    const root=box.content, table=root.querySelector('table');
    if(!table) throw Error('找不到明細表格');
    const range=document.createRange();range.selectNodeContents(root);range.setEndBefore(table);
    const h=document.createElement('div');h.append(range.cloneContents());h.querySelectorAll('head,style,script,title,meta').forEach(n=>n.remove());
    const result={headerText:textOf(h),totalAmount:'',items:[]};
    const warnings=[];let matched=false,missing=0,unknown=0;
    const rows=[...table.rows];
    for(const row of rows) {
      const c=[...row.cells];
      if(c.length>=4 && /[产產][品物]照片|故障原因/.test(c[1].textContent+c[2].textContent) && /[数數]量|金[额額]/.test(c[3].textContent) && !matched){matched=true;continue;}
      if(c.length<4) {
        const total=c.map(textOf).join(' ').trim().match(/(?:合[計计](?:金[额額])?|總金額|总金额)\s*[:：]?\s*([\s\S]*)/);
        if(total)result.totalAmount=total[1].trim();
        continue;
      }
      if(!matched && !/^\d+[.)、]?\s*$/.test(textOf(c[0]))) continue;
      const img=c[1].querySelector('img');
      let image=img?.getAttribute('src')||'';
      if(image && !picture(image)) {missing++;image='';}
      const status=textOf(c[4]);let isReturned=null;
      if(/已返回|已歸還|已归还/.test(status)) isReturned=true;
      else if(/未返回|返修中|未歸還/.test(status))isReturned=false;
      if(isReturned===null)unknown++;
      result.items.push({id:String(result.items.length+1),details:textOf(c[2]),qtyPrice:textOf(c[3]),image,isReturned,returnedDate:textOf(c[5])});
    }
    if(!matched && !result.items.length)throw Error('無法辨識表格，請選用返修系統匯出的 Word 明細');
    if(!result.items.length && !result.totalAmount && !result.headerText)throw Error('表格內沒有可匯入的內容');
    if(missing)warnings.push(missing+' 張圖片只有連結或格式不支援，請從原件補上');
    if(unknown)warnings.push(unknown+' 項歸還狀態待確認');
    return {report:result,warnings};
  }
  const all=(el,name)=>Array.from(el.getElementsByTagNameNS('*',name));
  const children=(el,name)=>[...el.children].filter(e=>e.localName===name);
  function paraText(n) {
    if(n.localName==='t')return n.textContent;
    if(n.localName==='br'||n.localName==='cr')return '\n';
    if(n.localName==='tab')return '\t';
    return [...n.children].map(paraText).join('');
  }
  function xml(str) {
    const doc=new DOMParser().parseFromString(str,'application/xml');
    if(doc.querySelector('parsererror'))throw Error('Word 文件的 XML 不完整');
    return doc;
  }
  async function docx(buffer) {
    const zip=await JSZip.loadAsync(buffer);
    const docFile=zip.file('word/document.xml');if(!docFile)throw Error('找不到 Word 內文');
    const doc=xml(await docFile.async('string')),rels=zip.file('word/_rels/document.xml.rels');
    const relMap=new Map(rels?all(xml(await rels.async('string')),'Relationship').filter(n=>n.getAttribute('TargetMode')!=='External').map(n=>[n.getAttribute('Id'),n.getAttribute('Target')]):[]);
    const ct=zip.file('[Content_Types].xml'),types=ct?xml(await ct.async('string')):null;
    const mime=(name)=>types&&(all(types,'Override').find(t=>t.getAttribute('PartName')==='/'+name)||all(types,'Default').find(t=>t.getAttribute('Extension')?.toLowerCase()===name.split('.').pop().toLowerCase()))?.getAttribute('ContentType');
    const body=all(doc,'body')[0], table=body&&children(body,'tbl')[0];if(!table)throw Error('找不到 Word 明細表格');
    let html='';
    for(const node of body.children){if(node===table)break;if(node.localName==='p')html+='<p>'+line(paraText(node))+'</p>';}
    html+='<table>';let extraImages=0;
    for(const row of children(table,'tr')) {
      html+='<tr>';
      for(const cell of children(row,'tc')) {
        html+='<td>'+all(cell,'p').map(p=>'<p>'+line(paraText(p))+'</p>').join('');
        const imgs=[...all(cell,'blip'),...all(cell,'imagedata')];
        if(imgs.length>1)extraImages+=imgs.length-1;
        if(imgs.length) {
          const rel=Array.from(imgs[0].attributes).find(a=>['embed','id','link'].includes(a.localName));
          const target=rel&&relMap.get(rel.value);
          const name=target?new URL(target,'https://word.invalid/word/').pathname.slice(1):'';
          const file=zip.file(name);
          let src='missing';
          if(file)src='data:'+mime(name)+';base64,'+await file.async('base64');
          html+='<img src="'+esc(src)+'">';
        }
        html+='</td>';
      }
      html+='</tr>';
    }
    const result=fromHTML(html+'</table>');if(extraImages)result.warnings.push('同一欄位有多張照片，已讀取第一張；請保留 Word 原檔');return result;
  }
  async function parse(file) {
    const buf=await file.arrayBuffer(), a=new Uint8Array(buf);
    const fingerprint=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',buf)),b=>b.toString(16).padStart(2,'0')).join('');
    let result;
    if(a[0]===0x50 && a[1]===0x4b) result=await docx(buf);
    else {
      if(a[0]===0xd0 && a[1]===0xcf)throw Error('舊式二進位 .doc，請用 Word 另存 .docx 或複製表格貼上');
      const encoding=a[0]===255&&a[1]===254?'utf-16le':a[0]===254&&a[1]===255?'utf-16be':'utf-8';
      result=fromHTML(new TextDecoder(encoding).decode(a));
    }
    return {...result,key:'word-'+fingerprint,name:file.name};
  }
  function render(r) {
    return '<!doctype html><html><head><meta charset="utf-8"><title>返修品明細</title></head><body><p>'+line(r.headerText)+'</p><table border="1" cellspacing="0" cellpadding="6"><tr>'+['序號','產品照片','型號／品名／故障原因','數量／金額','歸還狀態','歸還日期'].map(v=>'<th>'+v+'</th>').join('')+'</tr>'+
    r.items.map((i,n)=>'<tr><td>'+(n+1)+'</td><td>'+(picture(i.image)?'<img width="160" src="'+esc(i.image)+'">':'')+'</td><td>'+line(i.details)+'</td><td>'+line(i.qtyPrice)+'</td><td>'+(i.isReturned===true?'已返回公司':i.isReturned===false?'返修中':'待確認')+'</td><td>'+line(i.returnedDate)+'</td></tr>').join('')+
    '<tr><td colspan="2"></td><td colspan="4">合計金額 '+line(r.totalAmount)+'</td></tr></table></body></html>';
  }
  return {parse,fromHTML,render,esc,picture};
})();
