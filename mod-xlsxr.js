/* mod-xlsxr.js — 엑셀(.xlsx) 읽기 (외부 라이브러리 없이 ZIP·XML 해석)
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.xlsxr로 내보냄.
   MODULE xlsxr V1.16 (2026-10-29) */
(function(M){
/* ---- .xlsx 읽기 (외부 라이브러리 없이: ZIP 목차 → 압축 풀기 → XML) ---- */
async function readZip(buf){
  const dv=new DataView(buf), u8=new Uint8Array(buf), dec=new TextDecoder();
  let e=-1;
  for(let i=buf.byteLength-22;i>=Math.max(0,buf.byteLength-65557);i--){ if(dv.getUint32(i,true)===0x06054b50){ e=i; break; } }
  if(e<0) throw new Error('엑셀(.xlsx) 파일을 읽을 수 없습니다');
  const n=dv.getUint16(e+10,true); let p=dv.getUint32(e+16,true);
  const out=new Map();
  for(let i=0;i<n&&dv.getUint32(p,true)===0x02014b50;i++){
    const method=dv.getUint16(p+10,true), csize=dv.getUint32(p+20,true), nlen=dv.getUint16(p+28,true), xlen=dv.getUint16(p+30,true), clen=dv.getUint16(p+32,true), lho=dv.getUint32(p+42,true);
    const name=dec.decode(u8.subarray(p+46,p+46+nlen));
    const start=lho+30+dv.getUint16(lho+26,true)+dv.getUint16(lho+28,true);
    out.set(name,{method,data:u8.subarray(start,start+csize)});
    p+=46+nlen+xlen+clen;
  }
  return out;
}
async function zipText(ents,name){
  const e=ents.get(name); if(!e) return null;
  let bytes=e.data;
  if(e.method===8){
    if(typeof DecompressionStream==='undefined') throw new Error('이 브라우저는 엑셀 읽기를 지원하지 않습니다. 브라우저를 업데이트하거나 내용을 복사해 붙여넣어 주세요');
    bytes=new Uint8Array(await new Response(new Blob([e.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
  } else if(e.method!==0) throw new Error('지원하지 않는 엑셀 압축 형식입니다');
  return new TextDecoder().decode(bytes);
}
async function xlsxRows(buf){
  const ents=await readZip(buf), parse=x=>new DOMParser().parseFromString(x,'application/xml');
  const ss=[], ssXml=await zipText(ents,'xl/sharedStrings.xml');
  if(ssXml) for(const si of parse(ssXml).getElementsByTagName('si')){ let t=''; for(const x of si.getElementsByTagName('t')) t+=x.textContent; ss.push(t); }
  const colIdx=ref=>{ let n=0; for(const ch of ref.replace(/[^A-Z]/g,'')) n=n*26+ch.charCodeAt(0)-64; return n-1; };
  const names=[...ents.keys()].filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort((a,b)=>parseInt(a.match(/\d+/))-parseInt(b.match(/\d+/)));
  const sheets=[];
  for(const name of names){
    const rows=[];
    for(const row of parse(await zipText(ents,name)).getElementsByTagName('row')){
      const arr=[];
      for(const c of row.getElementsByTagName('c')){
        const t=c.getAttribute('t'), r=c.getAttribute('r'), vEl=c.getElementsByTagName('v')[0];
        let v='';
        if(t==='s') v=vEl?(ss[+vEl.textContent]??''):'';
        else if(t==='inlineStr') v=[...c.getElementsByTagName('t')].map(x=>x.textContent).join('');
        else if(vEl){ v=vEl.textContent; if(t!=='str'&&t!=='e'&&v!==''&&!isNaN(v)) v=Number(v); }
        arr[r?colIdx(r):arr.length]=v;
      }
      rows.push(arr);
    }
    sheets.push(rows);
  }
  return sheets;
}
M.xlsxr={xlsxRows};
})(window.M=window.M||{});
