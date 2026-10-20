/* ================= 업무 공통 모듈 core.js =================
   업무허브(index.html) · 업무일지 연동 · 가계부(expenses.html)가 함께 쓰는 코드
   - 카드 엑셀/붙여넣기 읽기 → 가계부 + 업무일지 반영 + 원본 평생 보관 (importCore)
   - 분류(생활·업무·기타비용 · 세부 · 가게 묶음), 고정지출
   - 동기화: GitHub 비공개 Gist의 expenses.json + worklog.json을 한 번에 합치고 올림 (syncNow)
   각 화면은 uiSync(state) · uiChanged() 두 함수를 정의해 동기화 상태를 그립니다.
   CORE V1.04 (2026-10-20) */
const CARD_CATS={fuel:{label:'주유',ic:'⛽'},urea:{label:'요소수',ic:'💧'},toll:{label:'통행료',ic:'🛣️'}};
/* ---- 공통 ---- */
const $=id=>document.getElementById(id);
const pad=n=>String(n).padStart(2,'0');
const fmt=n=>Math.round(Number(n)||0).toLocaleString('ko-KR');
const escH=v=>String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const keyOf=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const parseKey=k=>{ const [y,m,d]=k.split('-').map(Number); return new Date(y,m-1,d); };
const WD=['일','월','화','수','목','금','토'];
const ls={ get(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } }, set(k,v){ try{ localStorage.setItem(k,v); return true; }catch(e){ return false; } } };
const rj=(k,d)=>{ try{ const v=JSON.parse(ls.get(k)||'null'); return v==null?d:v; }catch(e){ return d; } };
let tt; function toast(m){ const t=$('toast'); t.textContent=m; t.classList.add('on'); clearTimeout(tt); tt=setTimeout(()=>t.classList.remove('on'),2600); }
try{ navigator.storage&&navigator.storage.persist&&navigator.storage.persist(); }catch(e){}

/* ---- 카드 엑셀 읽기 · 업무 비용 분류 (업무일지와 같은 규칙) ---- */
/* ================= 💳 카드내역 가져오기 (삼성카드 이용내역 엑셀 · CSV · 붙여넣기) =================
   ① 파일(.xlsx/.csv) 또는 복사한 표를 읽고  ② 같은 승인번호의 승인·취소를 상계한 뒤
   ③ 가맹점명으로 주유 · 요소수 · 통행료를 나누고(나머지는 제외)  ④ 미리보기에서 확인 → 날짜별 비용에 반영
   반영한 결제는 그날 기록(cd)에 남겨 같은 파일을 다시 올려도 두 번 들어가지 않음.
   카드 합계와 다른 값을 손으로 입력해 둔 날은 미리보기에서 '직접 입력값 유지'를 고를 수 있음.          */
const BIKE_MAX=60000;                         // 주유소 6만원 이하: 같은 날 큰 주유(>6만)가 있으면 요소수(업무), 없으면 바이크 주유(생활)
const STATION_RE=/주유소|SK에너지|칼텍스|오일뱅크|S-?OIL|에쓰오일|알뜰주유/, LPG_RE=/충전소|LPG|엘피지/;   // 충전소 = LPG 차량 충전(생활)
function isBigFuel(m,v){ const s=String(m||'').toUpperCase(); return !LPG_RE.test(s)&&STATION_RE.test(s)&&v>BIKE_MAX; }
const K_CARDRULES='worklog.cardRules';         // 사용자가 미리보기에서 넣고/뺀 가맹점 기억 (이 기기)
function cardRules(){ try{ return JSON.parse(ls.get(K_CARDRULES)||'{}')||{}; }catch(e){ return {}; } }
function classifyCard(m,v,big){   // big: 같은 날 큰 주유가 있음
  const rule=cardRules()[m];
  if(rule!==undefined) return {c:CARD_CATS[rule]?rule:null,rule:true};
  const s=m.toUpperCase();
  if(/요소수|유록스|UREA|롯데정밀화학/.test(s)) return {c:'urea'};
  if(/도로공사|하이패스|HI-?PASS|통행료|고속도로|유료도로/.test(s)) return {c:'toll'};
  if(LPG_RE.test(s)) return {c:null};                                       // LPG 충전 → 업무일지 제외(가계부 생활비용)
  if(STATION_RE.test(s)) return v>BIKE_MAX?{c:'fuel'}:big?{c:'urea',guess:true}:{c:null};   // 소액: 큰 주유 같은 날만 요소수
  return {c:null};
}

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
/* ---- .csv · 붙여넣기 읽기 ---- */
function decodeText(buf){
  let t=new TextDecoder('utf-8').decode(buf);
  if(t.includes('�')){ try{ t=new TextDecoder('euc-kr').decode(buf); }catch(e){} }   // 카드사 CSV는 EUC-KR인 경우가 많음
  return t.replace(/^﻿/,'');
}
function textRows(t){
  return t.split(/\r?\n/).filter(l=>l.trim()).map(l=>{
    if(l.includes('\t')) return l.split('\t');
    const out=[]; let cur='', q=false;
    for(let i=0;i<l.length;i++){ const ch=l[i];
      if(q){ if(ch==='"'&&l[i+1]==='"'){ cur+='"'; i++; } else if(ch==='"') q=false; else cur+=ch; }
      else if(ch==='"') q=true; else if(ch===','){ out.push(cur); cur=''; } else cur+=ch; }
    out.push(cur);
    return out.length>1?out:l.trim().split(/\s{2,}/);
  });
}
/* ---- 행 → 결제 ---- */
function parseCardDate(x){
  if(typeof x==='number'&&x>20000&&x<80000){ const d=new Date(Date.UTC(1899,11,30)+Math.round(x)*864e5); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`; }
  const s=String(x??'').trim();
  const m=s.match(/(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/)||s.match(/^(20\d{2})(\d{2})(\d{2})$/);
  if(!m) return null;
  const mo=+m[2], da=+m[3]; if(mo<1||mo>12||da<1||da>31) return null;
  return `${m[1]}-${pad(mo)}-${pad(da)}`;
}
function parseCardTime(x){ const m=String(x??'').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/); return m?{hm:`${pad(+m[1])}:${m[2]}`,full:`${pad(+m[1])}:${m[2]}:${m[3]||'00'}`}:{hm:'',full:''}; }
function parseCardAmt(x){
  if(typeof x==='number') return Number.isFinite(x)?x:null;
  const s=String(x??'').replace(/[,원\s]/g,'');
  return /^-?\d+(\.\d+)?$/.test(s)?Number(s):null;
}
function cardHeaderMap(row){
  const m={};
  row.forEach((c,i)=>{ const s=String(c??'').replace(/\s/g,''); if(!s) return;
    if(m.ap===undefined&&/승인번호/.test(s)) m.ap=i;
    else if(m.cancel===undefined&&/취소/.test(s)) m.cancel=i;
    else if(m.time===undefined&&/(승인|이용|거래)?(시각|시간)/.test(s)&&!/일자/.test(s)) m.time=i;
    else if(m.date===undefined&&/승인일자|이용일자|이용일|거래일자|거래일|승인일/.test(s)) m.date=i;
    else if(m.m===undefined&&/가맹점/.test(s)) m.m=i;
    else if(m.v===undefined&&/금액/.test(s)&&!/합계/.test(s)) m.v=i;
  });
  return m;
}
const CARD_SKIP=/^(일시불|할부|본인|가족|-|국내|해외|전체취소|부분취소)$/;
function cardRowLoose(r){
  const cells=r.map(x=>typeof x==='number'?x:String(x??'').trim());
  const di=cells.findIndex(c=>parseCardDate(c)); if(di<0) return null;
  const ti=cells.findIndex(c=>/^\d{1,2}:\d{2}/.test(String(c)));
  const mi=cells.findIndex((c,i)=>i!==di&&i!==ti&&typeof c==='string'&&c&&parseCardAmt(c)===null&&!parseCardDate(c)&&!CARD_SKIP.test(c));
  if(mi<0) return null;
  const vi=cells.findIndex((c,i)=>i>mi&&parseCardAmt(c)!==null);
  if(vi<0) return null;
  const t=parseCardTime(ti>=0?cells[ti]:'');
  let v=parseCardAmt(cells[vi]); if(cells.some(c=>/취소/.test(String(c)))&&v>0) v=-v;
  return {k:parseCardDate(cells[di]),tm:t.hm,ts:t.full,m:cells[mi],v,ap:''};
}
function rowsToCardTx(rows){
  let map=null, hi=-1;
  for(let i=0;i<Math.min(rows.length,20);i++){ const m=cardHeaderMap(rows[i]||[]); if(m.date!==undefined&&m.m!==undefined&&m.v!==undefined){ map=m; hi=i; break; } }
  const tx=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[]; let t=null;
    if(map){
      const k=parseCardDate(r[map.date]), v=parseCardAmt(r[map.v]), m=String(r[map.m]??'').trim();
      if(k&&v!==null&&m){
        const tt=parseCardTime(map.time!==undefined?r[map.time]:'');
        t={k,tm:tt.hm,ts:tt.full,m,v,ap:map.ap!==undefined?String(r[map.ap]??'').trim():''};
        if(map.cancel!==undefined&&/취소/.test(String(r[map.cancel]??''))&&t.v>0) t.v=-t.v;
      }
    } else t=cardRowLoose(r);
    if(t) tx.push(t);
  }
  return tx;
}
/* ---- 형식 없는 붙여넣기: 줄·칸 모양과 상관없이 '날짜'를 기준으로 한 건씩 묶고, 그 안에서 금액·가맹점을 찾음 ----
   예) 앱 화면(가맹점 → 금액 → 날짜 줄), 한 줄 표, 날짜가 먼저 오는 목록, 연도 없는 날짜(10.15 · 10월 15일)도 OK */
const FZ_DATE=/(20\d{2})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})\s*일?|(?<![\d.,])(\d{1,2})\s*(?:[.\/]|월\s*)(\d{1,2})\s*일?(?![\d,]|\s*원)/;
const FZ_AMT=/(-?\d{1,3}(?:,\d{3})+|-?\d+)\s*원|(-?\d{1,3}(?:,\d{3})+)(?![\d*])/g;
const FZ_SKIP_AMT=/(할인|적립|포인트|잔액|한도|수수료|이자|캐시백|혜택|청구|결제예정|누적|합계|총)\s*[:：]?\s*$/;
const FZ_STOP=/^(본인|가족|일시불|할부|무이자|\d+개월|승인|정상|국내|해외|신용|체크|일반|결제|매출|복지할인|할인|적립|포인트|원|KRW|매입|확정|대기|전표|이용|이용일|이용일시|가맹점|가맹점명|금액|이용금액|승인금액|승인번호|카드|구분|상태|사용처|누적|잔액|합계|한도|[\d\-:.,*\/]+|\[.*\]|\(.*\)|\S*님|\S*\d{3,}\*+\S*|\*+\d+)$/i;
function fzDate(m){
  if(m[1]) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  const mo=+m[4], d=+m[5]; if(mo<1||mo>12||d<1||d>31) return '';
  const t=new Date(); let y=t.getFullYear(); if(new Date(y,mo-1,d)>new Date(t.getFullYear(),t.getMonth(),t.getDate()+1)) y--;   // 연도 없으면 미래가 안 되게
  return `${y}-${pad(mo)}-${pad(d)}`;
}
function fuzzyCardTx(text){
  const L=String(text||'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean), di=[];
  L.forEach((l,i)=>{ const m=l.match(FZ_DATE); if(m&&fzDate(m)) di.push(i); });
  if(!di.length) return [];
  const amtIn=l=>{ const out=[]; let m; FZ_AMT.lastIndex=0; const s=l.replace(FZ_DATE,' ').replace(/\d{1,2}:\d{2}(:\d{2})?/g,' ');
    while((m=FZ_AMT.exec(s))){ if(!FZ_SKIP_AMT.test(s.slice(Math.max(0,m.index-10),m.index))) out.push(+(m[1]||m[2]).replace(/,/g,'')); } return out; };
  const before=L.slice(0,di[0]).some(l=>amtIn(l).length);          // 금액이 날짜보다 먼저 나오면 '앞쪽 묶음' 형식
  const tx=[], cnt={};
  di.forEach((d,n)=>{
    const blk=before?L.slice(n?di[n-1]+1:0,d+1):L.slice(d,n+1<di.length?di[n+1]:L.length);
    const dm=L[d].match(FZ_DATE), k=fzDate(dm), tt=parseCardTime(L[d].replace(FZ_DATE,' ')+' '+blk.join(' ').replace(FZ_DATE,' '));
    let v=null; for(const l of (before?blk.slice().reverse():blk)){ const a=amtIn(l).filter(x=>x!==0); if(a.length){ v=a[0]; break; } }
    if(v===null) return;
    const cands=[];
    blk.forEach(l=>{ if(FZ_SKIP_AMT.test(l.replace(/[\d,\s원]+$/,''))&&!/[가-힣A-Za-z]{2,}.*\d/.test(l.replace(FZ_SKIP_AMT,''))) return;
      const c=l.replace(FZ_DATE,' ').replace(/\d{1,2}:\d{2}(:\d{2})?/g,' ').replace(/-?\d{1,3}(,\d{3})+\s*원?|-?\d+\s*원/g,' ')
        .split(/\t|\s{2,}|\|/).map(seg=>seg.split(/\s+/).filter(w=>w&&!FZ_STOP.test(w)&&!/^(할인|적립|포인트)/.test(w)).join(' ').trim()).filter(x=>/[가-힣A-Za-z]{2,}/.test(x));
      c.forEach(x=>cands.push({x,dl:l===L[d]})); });
    const nd=cands.filter(c=>!c.dl), pick=(nd.length?nd:cands), m=pick.length?(before?pick[pick.length-1]:pick[0]).x:'';
    if(!m) return;
    if(/취소/.test(blk.join(' '))&&v>0) v=-v;
    const ck=`${k}|${m}|${Math.abs(v)}|${v<0?'-':'+'}`; cnt[ck]=(cnt[ck]||0)+1;
    tx.push({k,tm:tt.hm,ts:tt.full||`~${Math.abs(v)}~${cnt[ck]}`,m,v,ap:''});   // 시각 없으면 금액·순번으로 구분 (같은 가게 같은 날 2번도 따로)
  });
  tx.filter(t=>t.v<0).forEach(c=>{ const o=tx.find(t=>t.v===-c.v&&t.m===c.m&&!t.x); if(o){ o.x=c.x=1; } });   // 취소: 다른 날짜의 같은 가게·같은 금액 승인과 상계
  return tx.filter(t=>!t.x);
}
function pasteTx(text){        // 표 형식으로 읽은 결과와 비교해 더 많이 읽힌 쪽 사용
  const rows=textRows(text), a=rowsToCardTx(rows), b=fuzzyCardTx(text);
  return b.length>a.length?{tx:b,rows:rows,fuzzy:true}:{tx:a,rows:rows,fuzzy:false};
}
/* 같은 결제(승인번호, 없으면 날짜+시각+가맹점)의 승인·취소를 더해서 정리 */
function netCardTx(tx){
  const g=new Map();
  tx.forEach(t=>{
    const id=t.ap?`${t.k}#${t.ap}`:`${t.k}#${t.ts}#${t.m}`;
    let x=g.get(id); if(!x){ x={id,k:t.k,tm:t.tm,ts:t.ts,m:t.m,v:0,pos:false,neg:false}; g.set(id,x); }
    x.v+=t.v; if(t.v>0) x.pos=true; if(t.v<0) x.neg=true;
  });
  const items=[], orphan=[]; let canceled=0;
  for(const x of g.values()){
    x.v=Math.round(x.v);
    if(x.v>0) items.push(x); else if(x.pos&&x.neg) canceled++; else if(x.neg) orphan.push({id:x.id,k:x.k});
  }
  items.sort((a,b)=>a.k===b.k?(a.ts<b.ts?1:a.ts>b.ts?-1:0):(a.k<b.k?1:-1));     // 최근 날짜·시각이 위
  return {items,canceled,orphan,raw:tx.length};
}

/* ================= 데이터 =================
   expenses.data.v1 = { 't:<결제id>':{d:'YYYY-MM-DD', t:'HH:MM', m:가맹점, v:금액, g?:분류(직접 지정), n?:메모, e?:1(금액 직접 수정), man?:1(직접 추가)},
                        'r:<가게>':{g:분류} (같은 가게 분류 규칙) }
   지운 결제는 동기화 기록(del)에 남아 같은 엑셀을 다시 올려도 되살아나지 않음                                  */
const EK='expenses.data.v1', EM='expenses.sync.meta', WK='worklog.data.v2', WM='worklog.sync.meta';
let E=rj(EK,{});
const metaOf=k=>Object.assign({mt:{},del:{},last:0,dirty:false,note:''},rj(k,{}));
function touch(keys){
  const m=metaOf(EM), now=Date.now();
  keys.forEach(k=>{ if(E[k]){ m.mt[k]=now; delete m.del[k]; } else { m.del[k]=now; delete m.mt[k]; } });
  m.dirty=true; touchN++; ls.set(EM,JSON.stringify(m));
  if(!ls.set(EK,JSON.stringify(E))) toast('⚠ 저장 실패: 저장 공간을 확인하세요');
  schedSync(1500);
}
/* ---- 분류 (같은 계열 묶기) ---- */
const GROUPS=[      // 위에서부터 먼저 맞는 분류 (구체적인 것 먼저)
  ['통신비','📱',/휴대폰알림|알림서비스|LGU|LG ?U\+|LG유플러스|유플러스|SKT|SK텔레콤|SK브로드밴드|\bKT\b|케이티|올레|헬로모바일|알뜰폰|세븐모바일|이야기모바일|통신요금|통신비|인터넷요금/i],
  ['보험','🛡️',/보험|화재해상|해상화재|손해보험|손보|생명|라이나|메리츠|현대해상|DB손|KB손|삼성화재|한화손|흥국|AIA|처브|캐롯|하나손/i],
  ['의료','💊',/병원|의원|약국|치과|한의원|안과|피부과|내과|외과|정형|소아과|이비인후|산부인과|비뇨|정신건강|재활|검진|요양|메디|클리닉/],
  ['주유·차량','⛽',/주유소|에너지|칼텍스|오일뱅크|S-?OIL|에쓰오일|충전소|요소수|유록스|세차|정비|타이어|카센터|카센타|공업사|자동차|오토|모터스|부품/i],
  ['통행료','🛣️',/도로공사|하이패스|통행료|고속도로|유료도로/],
  ['교통','🚌',/택시|카카오T|티머니|캐시비|버스|지하철|코레일|KTX|SRT|철도|주차|쏘카|그린카|따릉이|킥보드|항공/i],
  ['주거·공과금','🏠',/입주자대표|관리비|관리사무소|도시가스|가스공사|한국전력|한전|전기요금|수도|상하수도|월세|부동산/],
  ['공공·세금','🏛️',/시청|구청|군청|도청|주민센터|행정복지|세무서|국세|지방세|우체국|경찰|법원|등기소|과태료|범칙금|정부24|민원/],
  ['교육','📚',/학원|교육|학교|유치원|어린이집|교재|강의|인강|클래스101|교습/],
  ['편의점','🏪',/GS25|지에스\(?GS\)?25|씨유|\(CU\)|^CU|세븐일레븐|코리아세븐|이마트24|미니스톱|편의점/],
  ['카페·간식','☕',/커피|스타벅스|카페|투썸|이디야|빽다방|메가엠지씨|메가커피|컴포즈|할리스|폴바셋|커피빈|탐앤탐스|더벤티|매머드|파리바게|뚜레쥬르|베이커리|제과|브레드|배스킨|베스킨|던킨|크리스피|공차|설빙|빙수|디저트|케이크|마카롱|도넛|와플|아이스크림|요거|떡집|빵/],
  ['외식','🍚',/아웃백|빕스|애슐리|하누돈|한우|한돈|정육식당|고깃|고기|삼겹|갈비|등심|BBQ|비비큐|교촌|BHC|굽네|네네|처갓집|호식이|치킨|맘스터치|버거|맥도날드|롯데리아|KFC|써브웨이|서브웨이|피자|도미노|본죽|죽이야기|김밥|분식|떡볶이|찜닭|두찜|마라|양꼬치|쌀국수|라멘|우동|소바|돈까스|돈가스|돈카츠|규동|덮밥|백반|한정식|정식|밥상|밥집|국시|면옥|냉면|칼국수|수제비|국수|국밥|해장|설렁탕|곰탕|감자탕|순대|보쌈|족발|곱창|막창|대창|닭갈비|닭발|닭한마리|닭|오리|장어|해물|아구|아귀|대게|꽃게|조개|낙지|쭈꾸미|주꾸미|횟집|수산|초밥|스시|오마카세|이자카야|포차|호프|술집|한식|중식|일식|양식|뷔페|식당|반점|짜장|짬뽕|가든|숯불|화로|찌개|탕|찹쌀/i],
  ['마트·생활','🛒',/마트|하나로|농협|노브랜드|이마트|트레이더스|홈플러스|코스트코|GS더프레시|롯데슈퍼|식자재|정육점|청과|과일|야채|반찬|다이소|아성|올리브영|생활용품|세탁|철물|무인/],
  ['쇼핑','🛍️',/쿠팡|11번가|G마켓|지마켓|옥션|SSG|무신사|유니클로|탑텐|스파오|ZARA|나이키|아디다스|백화점|아울렛|신세계|롯데쇼핑|의류|신발|안경|문구|교보|알라딘|예스24|네이버페이|카카오페이/i],
  ['여가·문화','🎬',/CGV|롯데시네마|메가박스|영화|넷플릭스|NETFLIX|유튜브|YOUTUBE|티빙|웨이브|디즈니|멜론|스포티파이|노래|PC방|피씨방|볼링|당구|헬스|골프|스크린|사우나|찜질|호텔|모텔|펜션|숙박|여행|야놀자|여기어때|에어비앤비|테마파크/i],
  ['대출','🏦',/대출|캐피탈|카드론|원리금|주택금융|HF|이자상환/],
  ['경조사','🎁',/축의|조의|부의|경조|기부|헌금|후원/],
  ['금융·수수료','💳',/수수료|연회비|이자/],
];
const FOOD_TAIL=/(당|옥|뜰|집|관|각|정|가든|돈|하우스|키친|식탁|밥|면)$/;   // 한글 가게 이름 끝말로 본 음식점 (예: 월미당, 정뜰, 하누돈)
const GICON=Object.fromEntries([...GROUPS.map(g=>[g[0],g[1]]),['기타','📦']]);
const GNAMES=[...GROUPS.map(g=>g[0]),'기타'];
const WORK_G=new Set(['주유·차량','통행료']);   // 업무비로 보는 분류 (교통·주차는 생활비)
/* 큰 분류 3가지: 생활비용 · 업무비용 · 기타비용 (세부 분류에서 자동으로 정하고, 결제·가게별로 직접 옮길 수 있음) */
const CATS=[['life','생활비용','🏠'],['work','업무비용','🚚'],['etc','기타비용','📦']];
const CATN=Object.fromEntries(CATS.map(c=>[c[0],c[1]])), CATI=Object.fromEntries(CATS.map(c=>[c[0],c[2]]));
const CAT_DEF={'주유·차량':'work','통행료':'work','기타':'etc','금융·수수료':'etc','경조사':'etc'};   // 나머지(주거·대출·식비…)는 생활비용
const FIXED_G=new Set(['주거·공과금','대출','통신비','보험']);                                      // 고정비로 보는 세부 분류
let _bf=null,_bfE=null,_bfN=-1,touchN=0;
function bigFuelDay(d){          // 그날 6만원 넘는 주유가 있었나 (요소수 ↔ 바이크 주유 구분)
  if(_bfE!==E||_bfN!==touchN){ _bf=new Set(); Object.keys(E).forEach(k=>{ const r=E[k]; if(k.startsWith('t:')&&r&&isBigFuel(r.m,r.v)) _bf.add(r.d); }); _bfE=E; _bfN=touchN; }
  return _bf.has(d);
}
function catOf(r){
  if(r.c) return r.c;
  const rule=E['r:'+brandOf(r.m)]; if(rule&&rule.c) return rule.c;
  const g=groupOf(r);
  if(g==='주유·차량'){ const s=String(r.m||'').toUpperCase();
    if(LPG_RE.test(s)) return 'life';                                              // 충전소 = LPG 차량 충전
    if(STATION_RE.test(s)&&r.v<=BIKE_MAX&&!bigFuelDay(r.d)) return 'life'; }       // 6만원 이하 단독 주유 = 바이크
  return CAT_DEF[g]||'life';
}
/* V1.01: 예전에 업무일지 비용으로 들어간 충전소(LPG)·바이크 주유를 업무일지에서 빼기 (몇 번 실행돼도 같은 결과) */
function migrateWork(){
  const W=rj(WK,{}), wm=metaOf(WM), now=Date.now(); let n=0;
  Object.keys(W).forEach(k=>{ const rec=W[k]; if(!rec||!rec.cd||!rec.cd.length) return;
    const big=rec.cd.some(e=>e.c==='fuel'&&e.v>BIKE_MAX&&isBigFuel(e.m,e.v)), out=rec.cd.filter(e=>(e.c==='fuel'||e.c==='urea')&&classifyCard(e.m,e.v,big).c===null);
    if(!out.length) return;
    out.forEach(e=>{ const v=(+rec[e.c]||0)-e.v; if(v>0) rec[e.c]=v; else delete rec[e.c]; });
    rec.cd=rec.cd.filter(e=>!out.includes(e)); if(!rec.cd.length) delete rec.cd;
    wm.mt[k]=now; delete wm.del[k]; n++; });
  if(n){ wm.dirty=true; ls.set(WK,JSON.stringify(W)); ls.set(WM,JSON.stringify(wm)); }
  return n;
}                       // 업무비로 보는 분류
const BRANDS=[[/GS25|지에스\(?GS\)?25/,'GS25'],[/씨유|\(CU\)|^CU/,'CU'],[/세븐일레븐|코리아세븐/,'세븐일레븐'],[/이마트24/,'이마트24'],
  [/메가엠지씨|메가커피/,'메가커피'],[/스타벅스/,'스타벅스'],[/파리바게/,'파리바게뜨'],[/다이소/,'다이소'],[/노브랜드/,'노브랜드'],[/하나로/,'농협하나로'],[/도로공사|하이패스/,'하이패스']];
function brandOf(m){
  for(const [re,b] of BRANDS) if(re.test(m)) return b;
  const p=m.replace(/\(주\)|㈜|주식회사/g,'').trim().split(/\s+/);
  if(p.length>1&&/점$/.test(p[p.length-1])) p.pop();          // "윈플러스마트 양주점" → "윈플러스마트"
  return p.join(' ')||m;
}
function groupOf(r){
  if(r.g) return r.g;
  const rule=E['r:'+brandOf(r.m)]; if(rule&&rule.g) return rule.g;
  for(const [n,,re] of GROUPS) if(re.test(r.m)) return n;
  const last=brandOf(r.m).split(/\s+/).pop();
  if(/^[가-힣]{2,8}$/.test(last)&&FOOD_TAIL.test(last)) return '외식';
  const kk=E['k:'+brandOf(r.m)]; if(kk&&kk.g) return kk.g;               // 카카오 장소 검색으로 찾은 업종
  return '기타';
}
const ymKey=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}`;
function recTx(){                         // 고정지출(주거비·대출 등 매달 반복) → 시작 달부터 이번 달까지 자동 생성
  const out=[], nowYM=ymKey(new Date());
  Object.keys(E).filter(k=>k.startsWith('p:')).forEach(k=>{
    const p=E[k]; if(!p||!p.from) return;
    let [y,m]=p.from.split('-').map(Number); m--;
    const end=p.to&&p.to<nowYM?p.to:nowYM;
    for(let i=0;i<600&&`${y}-${pad(m+1)}`<=end;i++){
      const mo=`${y}-${pad(m+1)}`, dim=new Date(y,m+1,0).getDate();
      out.push({key:k+'|'+mo,rec:k,d:`${mo}-${pad(Math.min(p.day||1,dim))}`,m:p.m,v:p.v,g:p.g,c:p.c,n:p.n,man:1});
      if(++m>11){ m=0; y++; }
    }
  });
  return out;
}
const allTx=()=>Object.keys(E).filter(k=>k.startsWith('t:')).map(k=>Object.assign({key:k},E[k])).concat(recTx());
const sumV=a=>a.reduce((x,r)=>x+r.v,0);

/* ================= 엑셀 → 가계부 + 업무일지 ================= */
function applyWork(items){            // 업무일지(worklog.data.v2) 날짜별 주유·요소수·통행료와 카드 결제 목록(cd)
  const W=rj(WK,{}), wm=metaOf(WM), now=Date.now(); let days=0;
  const byDay={};
  const big=k=>items.some(x=>x.k===k&&isBigFuel(x.m,x.v))||((W[k]||{}).cd||[]).some(e=>e.c==='fuel'&&e.v>BIKE_MAX);
  items.forEach(it=>{ const c=classifyCard(it.m,it.v,big(it.k)).c; if(c) (byDay[it.k]=byDay[it.k]||[]).push(Object.assign({c},it)); });
  Object.keys(byDay).forEach(k=>{
    const rec=Object.assign({},W[k]||{}), old=rec.cd||[];
    const adds=byDay[k].filter(it=>!old.some(e=>e.id===it.id||(e.m===it.m&&e.v===it.v&&(!e.tm||!it.tm||e.tm===it.tm))));
    if(!adds.length) return;
    const cd=old.concat(adds.map(it=>{ const o={id:it.id,m:it.m,v:it.v,c:it.c}; if(it.tm) o.tm=it.tm; return o; }));
    Object.keys(CARD_CATS).forEach(c=>{ const s=cd.filter(e=>e.c===c).reduce((x,e)=>x+e.v,0); if(s) rec[c]=s; });
    rec.cd=cd;
    const o={}; ['s','odo','fuel','toll','urea','hol','r','cd'].forEach(f=>{ if(rec[f]!==undefined) o[f]=rec[f]; });   // 업무일지 저장 순서
    W[k]=o; wm.mt[k]=now; delete wm.del[k]; days++;
  });
  if(days){ wm.dirty=true; ls.set(WK,JSON.stringify(W)); ls.set(WM,JSON.stringify(wm)); }
  return days;
}
function saveOriginal(rows,src,keys){        // 올린 엑셀·붙여넣기 원본을 그대로 평생 보관 (같은 내용은 한 번만)
  if(!rows||!rows.length) return;
  const R=rows.slice(0,5000).map(r=>Array.from(r||[],c=>c==null?'':c)), js=JSON.stringify(R);
  let h=0; for(let i=0;i<js.length;i++) h=(h*31+js.charCodeAt(i))|0;
  const fk='f:'+(h>>>0).toString(36)+js.length.toString(36);
  if(!E[fk]){ E[fk]={n:src,at:new Date().toISOString(),rows:R}; keys.push(fk); }
}
function importCore(tx,src,rows){        // 엑셀·붙여넣기 → 가계부(전체 결제) + 업무일지(주유·요소수·통행료) + 원본 보관
  const n=netCardTx(tx), md=metaOf(EM), keys=[];
  saveOriginal(rows,src,keys);
  const AT=allTx(), seen=new Set(AT.map(r=>`${r.d}|${r.t||''}|${r.m}|${r.v}`)), seenAny=new Set(AT.map(r=>`${r.d}|${r.m}|${r.v}`)), seenNoT=new Set(AT.filter(r=>!r.t).map(r=>`${r.d}|${r.m}|${r.v}`));   // 시각 없는 붙여넣기(앱 화면)와 엑셀도 같은 결제로
  let add=0,dup=0,skip=0;
  n.items.forEach(it=>{
    const key='t:'+it.id, cur=E[key];
    if(md.del[key]){ skip++; return; }                                   // 직접 지운 결제는 되살리지 않음
    if(cur){ if(cur.v!==it.v&&!cur.e){ cur.v=it.v; keys.push(key); add++; } else dup++; return; }
    const dmv=`${it.k}|${it.m}|${it.v}`;
    if(seen.has(`${it.k}|${it.tm||''}|${it.m}|${it.v}`)||(!it.tm&&seenAny.has(dmv))||(it.tm&&seenNoT.has(dmv))){ dup++; return; }
    E[key]={d:it.k,m:it.m,v:it.v}; if(it.tm) E[key].t=it.tm; keys.push(key); add++;
  });
  n.orphan.forEach(o=>{ const key='t:'+o.id; if(E[key]){ delete E[key]; keys.push(key); } });
  if(keys.length) touch(keys);
  const days=applyWork(n.items);
  const ks=n.items.map(i=>i.k).sort();
  const res={src,add,dup,skip,cancel:n.canceled,days,from:ks[0]||'',to:ks[ks.length-1]||'',at:Date.now()};
  ls.set('expenses.lastImport',JSON.stringify(res));
  return res;
}
const importMsg=r=>`✅ 새 결제 ${r.add}건 · 업무일지 ${r.days}일 반영${r.dup?` · 중복 ${r.dup}건 제외`:''}${r.cancel?` · 취소 ${r.cancel}건 상계`:''}`;
async function readCardFile(f){             // .xlsx · .csv → {tx, rows}
    const buf=await f.arrayBuffer(), h=new Uint8Array(buf.slice(0,4)); let tx=[], rows=[];
    if(h[0]===0x50&&h[1]===0x4B){ for(const sh of await xlsxRows(buf)){ const t=rowsToCardTx(sh); if(t.length>tx.length){ tx=t; rows=sh; } } }
    else if(h[0]===0xD0&&h[1]===0xCF) throw new Error('예전 엑셀(.xls)입니다. .xlsx로 저장하거나 복사해 붙여넣어 주세요');
    else { rows=textRows(decodeText(buf)); tx=rowsToCardTx(rows); }
    if(!tx.length) throw new Error('결제 내역을 찾지 못했습니다');
    return {tx,rows};
}
/* ================= ☁️ 동기화 (업무일지와 같은 GitHub 비공개 Gist · 토큰 공유) =================
   한 번 읽어서 expenses.json(가계부)과 worklog.json(업무일지)을 함께 합치고, 바뀐 파일만 한 번에 올림.
   합치기 규칙은 업무일지와 같음: 항목별로 더 나중에 고친 쪽이 이김, 지운 항목은 양쪽에서 지워짐            */
const syncTok=()=>ls.get('muhansync_token')||'', syncGist=()=>ls.get('muhansync_gist')||'';
async function gh(path,opt){
  opt=opt||{};
  const headers={Authorization:'Bearer '+syncTok(),Accept:'application/vnd.github+json'};
  if(opt.body) headers['Content-Type']='application/json';
  let r;
  try{ r=await fetch('https://api.github.com'+path,{method:opt.method||'GET',headers,body:opt.body,cache:'no-store'}); }
  catch(e){ throw new Error('인터넷 연결 없음 · 연결되면 자동으로 다시 맞춥니다'); }
  if(r.status===401) throw new Error('토큰이 틀렸거나 만료됐습니다. 새 토큰으로 다시 연결하세요');
  if(r.status===403||r.status===404) throw new Error('토큰에 gist 권한이 없거나 동기화 Gist를 찾을 수 없습니다 ('+r.status+')');
  if(!r.ok) throw new Error('GitHub 응답 오류 '+r.status);
  return r.json();
}
function mergeDocs(L,R){
  const now=Date.now(), out={data:{},mt:{},del:{}};
  const keys=new Set([...Object.keys(L.data||{}),...Object.keys(R.data||{}),...Object.keys(L.del||{}),...Object.keys(R.del||{})]);
  keys.forEach(k=>{
    const lv=(L.data||{})[k], rv=(R.data||{})[k];
    const lm=lv?((L.mt||{})[k]||0):-1, rm=rv?((R.mt||{})[k]||0):-1;
    const dT=Math.max((L.del||{})[k]||0,(R.del||{})[k]||0);
    let val=null, mt=-1;
    if(lv&&rv){
      if(JSON.stringify(lv)===JSON.stringify(rv)){ val=lv; mt=Math.max(lm,rm); }
      else if(lm>rm){ val=lv; mt=lm; }
      else if(rm>lm){ val=rv; mt=rm; }
      else { val=Object.assign({},rv,lv); mt=now; }          // 시각이 같으면 항목별로 합침
    } else if(lv){ val=lv; mt=lm; } else if(rv){ val=rv; mt=rm; }
    if(val&&mt>=dT){ out.data[k]=val; if(mt>0) out.mt[k]=mt; }
    else if(dT>0) out.del[k]=dT;
  });
  return out;
}


const canon=d=>JSON.stringify(Object.keys(d||{}).sort().map(k=>[k,d[k]]));
const DOCS=[
  {file:'expenses.json',mk:EM,app:'expenses',get:()=>E,set:d=>{ E=d; ls.set(EK,JSON.stringify(E)); }},
  {file:'worklog.json',mk:WM,app:'worklog',get:()=>rj(WK,{}),set:d=>ls.set(WK,JSON.stringify(d))}
];
let syncBusy=false, syncAgain=false, syncErr='', syncT=null;
function schedSync(ms){ if(!syncTok()||!syncGist()) return; clearTimeout(syncT); syncT=setTimeout(syncNow,ms); }
async function readGistFile(f){
  if(!f) return {};
  let c=f.content;
  if(f.truncated&&f.raw_url){ const r=await fetch(f.raw_url,{cache:'no-store'}); c=await r.text(); }   // 1MB 넘는 파일
  try{ return JSON.parse(c||'{}')||{}; }catch(e){ return {}; }
}
async function syncNow(){
  if(!syncTok()||!syncGist()){ uiSync(); return; }
  if(syncBusy){ syncAgain=true; return; }
  syncBusy=true; uiSync('busy');
  try{
    const g=await gh('/gists/'+syncGist()), files={}; let changed=false;
    for(const D of DOCS){
      const R=await readGistFile(g.files&&g.files[D.file]); R.data=R.data||{};
      const meta=metaOf(D.mk), L={data:D.get(),mt:meta.mt,del:meta.del}, M=mergeDocs(L,R);
      if(canon(M.data)!==canon(L.data)){ D.set(M.data); changed=true; }
      if(canon(M.data)!==canon(R.data)||canon(M.mt)!==canon(R.mt||{})||canon(M.del)!==canon(R.del||{}))
        files[D.file]={content:JSON.stringify({schema:'muhan-sync/1',app:D.app,version:APP_VERSION,updated:new Date().toISOString(),device:/Android|iPhone|iPad|Mobile/i.test(navigator.userAgent)?'모바일':'PC',data:M.data,mt:M.mt,del:M.del})};
      Object.assign(meta,{mt:M.mt,del:M.del,last:Date.now(),dirty:false}); ls.set(D.mk,JSON.stringify(meta));
    }
    if(Object.keys(files).length) await gh('/gists/'+syncGist(),{method:'PATCH',body:JSON.stringify({files})});
    syncErr='';
    if(changed){ if(migrateWork()) schedSync(800); uiChanged(); }
  }catch(e){ syncErr=e.message; }
  finally{ syncBusy=false; uiSync(); if(syncAgain){ syncAgain=false; schedSync(300); } }
}
migrateWork();
/* V1.02: 다른 탭·뒤로가기로 돌아오면 저장된 기록을 다시 읽어 옛 기록으로 덮어쓰지 않게 */
function reloadE(){ const n=rj(EK,null); if(n){ E=n; if(typeof uiChanged==='function') uiChanged(); } }
window.addEventListener('storage',e=>{ if(e.key===EK||e.key===WK) reloadE(); });
window.addEventListener('pageshow',e=>{ if(e.persisted) reloadE(); });
