/* Ekspor XLSX: template.xlsx asli disalin apa adanya; hanya sheet1.xml, workbook.xml & sharedStrings.xml diganti. */
const BL=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const tAjar=mo=>{const[y,m]=mo.split('-').map(Number);return m>=7?`${y} / ${y+1}`:`${y-1} / ${y}`};
const guru=(S,e)=>[S.p.nama,e.r].filter(Boolean).join(' & ');
const wkt=e=>e.a.replace(':','.')+' - '+e.b.replace(':','.');
function groupDays(S,mo){const l=S.e.filter(e=>e.d.startsWith(mo)).sort((a,b)=>a.d.localeCompare(b.d)||a.a.localeCompare(b.a)),g=[];
 l.forEach(e=>{const t=g[g.length-1];t&&t.d==e.d?t.rows.push(e):g.push({d:e.d,rows:[e]})});return g}
const crc=(()=>{const t=new Uint32Array(256).map((_,n)=>{for(let k=0;k<8;k++)n=n&1?0xEDB88320^n>>>1:n>>>1;return n>>>0});return u=>{let c=-1;for(const b of u)c=t[(c^b)&255]^c>>>8;return(c^-1)>>>0}})();
function zipReplace(buf,R){const v=new DataView(buf),u=new Uint8Array(buf),E=new TextEncoder(),w=[];let o=0,p=u.length-22;
 const push=a=>{w.push(a);o+=a.length};while(v.getUint32(p,true)!=0x06054b50)p--;
 const n=v.getUint16(p+10,true),C=[];let q=v.getUint32(p+16,true);
 for(let i=0;i<n;i++){const nl=v.getUint16(q+28,true),el=v.getUint16(q+30,true),cl=v.getUint16(q+32,true),lo=v.getUint32(q+42,true),nm=new TextDecoder().decode(u.subarray(q+46,q+46+nl));
  let m=v.getUint16(q+10,true),cr=v.getUint32(q+16,true),cs=v.getUint32(q+20,true),us=v.getUint32(q+24,true),d;
  if(R[nm]!=null){d=E.encode(R[nm]);m=0;cr=crc(d);cs=us=d.length}else{const s=lo+30+v.getUint16(lo+26,true)+v.getUint16(lo+28,true);d=u.subarray(s,s+cs)}
  const nb=E.encode(nm),h=new DataView(new ArrayBuffer(30));h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(8,m,true);h.setUint16(12,0x5821,true);h.setUint32(14,cr,true);h.setUint32(18,cs,true);h.setUint32(22,us,true);h.setUint16(26,nb.length,true);
  C.push({nb,m,cr,cs,us,o});push(new Uint8Array(h.buffer));push(nb);push(d);q+=46+nl+el+cl}
 const s=o;for(const c of C){const h=new DataView(new ArrayBuffer(46));h.setUint32(0,0x02014b50,true);h.setUint16(4,20,true);h.setUint16(6,20,true);h.setUint16(10,c.m,true);h.setUint16(14,0x5821,true);h.setUint32(16,c.cr,true);h.setUint32(20,c.cs,true);h.setUint32(24,c.us,true);h.setUint16(28,c.nb.length,true);h.setUint32(42,c.o,true);push(new Uint8Array(h.buffer));push(c.nb)}
 const e=new DataView(new ArrayBuffer(22));e.setUint32(0,0x06054b50,true);e.setUint16(8,C.length,true);e.setUint16(10,C.length,true);e.setUint32(12,o-s,true);e.setUint32(16,s,true);push(new Uint8Array(e.buffer));
 return new Blob(w,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
async function buildXlsx(S,mo,get){
 const[T,H,L,W]=await Promise.all([get('template.xlsx',1),get('sheet-head.xml'),get('sheet-tail.xml'),get('wb.xml')]);
 const[y,mi]=mo.split('-').map(Number),P=S.p,val={FLC:P.flc,WIL:P.wil,WILU:(P.wil||'').toUpperCase(),KOOR:P.koor,TA:tAjar(mo),BULAN:BL[mi-1],SHEET:BL[mi-1]+' '+y};
 const fill=s=>s.replace(/\{\{(\w+)\}\}/g,(_,k)=>esc(val[k]));
 const c=(r,s,v)=>v==null||v===''?`<c r="${r}" s="${s}"/>`:typeof v=='number'?`<c r="${r}" s="${s}"><v>${v}</v></c>`:`<c r="${r}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
 let r=9,x='';
 groupDays(S,mo).forEach((dy,i)=>{dy.rows.forEach((e,j)=>{const g=guru(S,e),ln=Math.max(1,Math.ceil(e.m.length/40),Math.ceil(e.k.length/14),Math.ceil((e.n||'').length/30),Math.ceil(g.length/20)),
   [Y,M,D]=dy.d.split('-').map(Number),ser=(Date.UTC(Y,M-1,D)-Date.UTC(1899,11,30))/864e5;
   x+=`<row r="${r}" ht="${Math.max(14.25,15*ln)}" customHeight="1">`+c('A'+r,28,j?null:i+1)+c('B'+r,j?35:29,j?null:ser)+c('C'+r,30,wkt(e))+c('D'+r,g.length>17?32:30,g)+c('E'+r,32,e.k)+c('F'+r,32,e.m)+c('G'+r,30,e.h===''||e.h==null?null:+e.h)+c('H'+r,37,e.t?'√':null)+c('I'+r,34,e.n)+'</row>';r++});
  x+=`<row r="${r}" ht="14.25" customHeight="1">`+[28,35,37,37,37,38,37,37,36].map((s,k)=>c('ABCDEFGHI'[k]+r,s)).join('')+'</row>';r++});
 const sst='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>';
 return zipReplace(T,{'xl/worksheets/sheet1.xml':fill(H)+x+L,'xl/workbook.xml':fill(W),'xl/sharedStrings.xml':sst})}
if(typeof module!='undefined')module.exports={buildXlsx,groupDays,tAjar,guru,wkt,BL,esc};
