/* mod-card.js — 카드 내역 읽기(엑셀·CSV·붙여넣기) · 중복 상계 · 업무 비용 분류
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.card로 내보냄.
   MODULE card V1.20 (2026-10-30) */
(function(M){
const {$,CARD_CATS,ls,pad}=M.base;
const {xlsxRows}=M.xlsxr;
/* ================= 💳 카드내역 가져오기 (삼성카드 이용내역 엑셀 · CSV · 붙여넣기) =================
   ① 파일(.xlsx/.csv) 또는 복사한 표를 읽고  ② 같은 승인번호의 승인·취소를 상계한 뒤
   ③ 가맹점명으로 주유 · 요소수 · 통행료를 나누고(나머지는 제외)  ④ 미리보기에서 확인 → 날짜별 비용에 반영
   반영한 결제는 그날 기록(cd)에 남겨 같은 파일을 다시 올려도 두 번 들어가지 않음.
   카드 합계와 다른 값을 손으로 입력해 둔 날은 미리보기에서 '직접 입력값 유지'를 고를 수 있음.          */
const BIKE_MAX=60000;                         // 주유소 6만원 이하: 같은 날 큰 주유(>6만)가 있으면 요소수(업무), 없으면 바이크 주유(생활)
const STATION_RE=/주유소|SK에너지|칼텍스|오일뱅크|S-?OIL|에쓰오일|알뜰주유/;
const LPG_RE=/충전소|LPG|엘피지/;   // 충전소 = LPG 차량 충전(생활)
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
  if(typeof x==='number'&&x>20000&&x<80000){ const d=new Date(Date.UTC(1899,11,30)+Math.floor(x)*864e5); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`; }
  const s=String(x??'').trim();
  const m=s.match(/(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/)||s.match(/^(20\d{2})(\d{2})(\d{2})$/);
  if(!m) return null;
  const mo=+m[2], da=+m[3]; if(mo<1||mo>12||da<1||da>31) return null;
  return `${m[1]}-${pad(mo)}-${pad(da)}`;
}
function parseCardTime(x){
  if(typeof x==='number'){ const f=x%1; if(!(f>0)) return {hm:'',full:''}; const t=Math.round(f*86400), h=Math.floor(t/3600)%24, mi=Math.floor(t%3600/60), se=t%60; return {hm:`${pad(h)}:${pad(mi)}`,full:`${pad(h)}:${pad(mi)}:${pad(se)}`}; }   // 엑셀 시각(소수)
  const m=String(x??'').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/); return m?{hm:`${pad(+m[1])}:${m[2]}`,full:`${pad(+m[1])}:${m[2]}:${m[3]||'00'}`}:{hm:'',full:''}; }
function parseCardAmt(x){
  if(typeof x==='number') return Number.isFinite(x)?x:null;
  const s=String(x??'').replace(/[,원\s]/g,'');
  return /^-?\d+(\.\d+)?$/.test(s)?Number(s):null;
}
function cardHeaderMap(row){     // 머리글 이름이 카드사마다 달라도 찾음 (이용일·거래일·사용일 / 가맹점·이용처·상호 / 이용금액·결제금액…)
  const m={}, vs=[];
  row.forEach((c,i)=>{ const s=String(c??'').replace(/\s/g,''); if(!s||s.length>16||/^[\d,.\-:\/]+$/.test(s)) return;
    if(m.ap===undefined&&/승인번호|승인No/i.test(s)) m.ap=i;
    else if(m.time===undefined&&/(시각|시간)$/.test(s)&&!/일자|일시/.test(s)) m.time=i;
    else if(m.date===undefined&&(/(승인|이용|거래|사용|매출|결제)(일자|일시|일|날짜)$/.test(s)||/^(일자|날짜|일시)$/.test(s))) m.date=i;
    else if(m.m===undefined&&/가맹점|이용처|사용처|상호|이용하신곳|거래처|사용내역|이용내역/.test(s)) m.m=i;
    else if(/금액|이용액|결제액|승인액/.test(s)&&!/합계|할인|적립|수수료|잔액|포인트|외화|해외|통화|US|\$|달러|현지|이자|예정|할부잔|원금잔/.test(s)) vs.push([i,/승인금액|이용금액|결제금액|거래금액|매출금액|승인액|이용액/.test(s)?3:/청구/.test(s)?1:2]);
    else if(m.cancel===undefined&&/취소|상태/.test(s)) m.cancel=i;
  });
  if(vs.length){ vs.sort((a,b)=>b[1]-a[1]||a[0]-b[0]); m.v=vs[0][0]; }
  return m;
}
const normM=m=>String(m||'').replace(/\(주\)|㈜|주식회사|\s|[()\[\].,·\-_'"]/g,'').toUpperCase();   // 띄어쓰기·(주) 달라도 같은 가게
const CARD_SKIP=/^(일시불|할부|본인|가족|-|국내|해외|전체취소|부분취소)$/;
const CARD_SUM=/^(합\s*계|소\s*계|총\s*계|누\s*계|총\s*이용.*|이용\s*합계|.*합계|일시불\s*\+\s*할부.*|할부\s*\+\s*일시불.*)$/;
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
function findHeader(rows){        // 머리글 줄 찾기 (두 줄로 나뉜 머리글도 합쳐서)
  const ok=m=>m.date!==undefined&&m.m!==undefined&&m.v!==undefined;
  for(let i=0;i<Math.min(rows.length,40);i++){
    let m=cardHeaderMap(rows[i]||[]); if(ok(m)) return {map:m,hi:i};
    const a=rows[i]||[], b=rows[i+1]||[]; if(!a.length||!b.length) continue;
    m=cardHeaderMap(Array.from({length:Math.max(a.length,b.length)},(_,k)=>`${a[k]??''}${b[k]??''}`)); if(ok(m)) return {map:m,hi:i+1};
  }
  return {map:null,hi:-1};
}
function rowsToCardTx(rows){
  const {map,hi}=findHeader(rows), tx=[], cnt={};
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[]; let t=null;
    if(map){
      const k=parseCardDate(r[map.date]), v=parseCardAmt(r[map.v]), m=String(r[map.m]??'').trim();
      if(k&&v!==null&&v!==0&&m&&!CARD_SUM.test(m)){
        const tt=parseCardTime(map.time!==undefined?r[map.time]:r[map.date]);   // 시각 칸이 없으면 날짜 칸 안의 시각(2026-08-24 17:32)
        t={k,tm:tt.hm,ts:tt.full,m,v,ap:map.ap!==undefined?String(r[map.ap]??'').trim():''};
        if(map.cancel!==undefined&&/취소/.test(String(r[map.cancel]??''))&&t.v>0) t.v=-t.v;
      }
    } else t=cardRowLoose(r);
    if(t&&CARD_SUM.test(t.m)) t=null;
    if(t){ if(!t.ap&&!t.ts){ const ck=`${t.k}|${t.m}|${Math.abs(t.v)}|${t.v<0?'-':'+'}`; cnt[ck]=(cnt[ck]||0)+1; t.ts=`~${Math.abs(t.v)}~${cnt[ck]}`; }   // 시각·승인번호 없는 같은 결제 2번도 따로
      tx.push(t); }
  }
  tx.filter(t=>t.v<0&&!t.ap).forEach(c=>{ const o=tx.find(t=>t.v===-c.v&&!t.ap&&!t.pair&&normM(t.m)===normM(c.m)&&t.k<=c.k); if(o){ o.pair=1; c.k=o.k; c.ts=o.ts; c.m=o.m; } });   // 승인번호 없는 취소 → 같은 가게·같은 금액 승인과 상계
  tx.filter(t=>t.v<0&&t.ap).forEach(c=>{ const o=tx.find(t=>t.v>0&&t.ap===c.ap&&t.k!==c.k); if(o&&!tx.some(t=>t.v>0&&t.ap===c.ap&&t.k===c.k)){ c.k=o.k; c.ts=o.ts; c.tm=o.tm; } });   // 같은 승인번호인데 취소가 다음 날짜로 찍힌 경우
  tx.filter(t=>t.v<0&&t.ap).forEach(c=>{ if(tx.some(t=>t.v>0&&t.ap===c.ap)) return;          // 승인번호가 다른 취소 줄(신한 등) → 같은 가게·같은 금액 승인과 상계
    const o=tx.find(t=>t.v>0&&t.v===-c.v&&!t.pair&&t.ap!==c.ap&&normM(t.m)===normM(c.m)&&t.k<=c.k); if(o){ o.pair=1; c.ap=o.ap; c.k=o.k; c.ts=o.ts; c.m=o.m; c.tm=o.tm; } });
  if(tx.length>2){ const tot=tx.reduce((x,t)=>x+t.v,0); for(let i=tx.length-1;i>=0;i--) if(tx[i].v>0&&tx[i].v*2===tot){ tx.splice(i,1); break; } }   // 나머지 합과 같은 금액 = 합계 줄
  return tx;
}
/* 데이터처럼 보이는데(날짜+금액) 결제로 못 읽은 줄 → '확인 필요'로 알려줌 */
function missedRows(rows,tx){
  const {hi}=findHeader(rows); let n=0;
  for(let i=hi+1;i<rows.length;i++){ const r=rows[i]||[];
    const hasD=r.some(c=>parseCardDate(c)), hasV=r.some(c=>typeof c==='number'?Math.abs(c)>=100&&!(c>20000&&c<80000&&parseCardDate(c)):/\d{1,3}(,\d{3})+|\d+\s*원/.test(String(c??'')));
    if(hasD&&hasV&&!r.some(c=>CARD_SUM.test(String(c??'').trim()))) n++; }
  const m=Math.max(0,n-tx.length);
  return m;
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
    if(x.v>0) items.push(x); else if(x.pos&&x.neg) canceled++; else if(x.neg) orphan.push({id:x.id,k:x.k,m:x.m,v:-x.v,tm:x.tm});
  }
  items.sort((a,b)=>a.k===b.k?(a.ts<b.ts?1:a.ts>b.ts?-1:0):(a.k<b.k?1:-1));     // 최근 날짜·시각이 위
  return {items,canceled,orphan,raw:tx.length};
}
async function readCardFile(f){             // .xlsx · .csv → {tx, rows}
    const buf=await f.arrayBuffer(), h=new Uint8Array(buf.slice(0,4)); let tx=[], rows=[];
    let miss=0;
    if(h[0]===0x50&&h[1]===0x4B){        // 모든 시트를 읽어 합침 (국내·해외 시트 등), 시트끼리 똑같은 줄은 한 번만
      const seen=new Set();
      for(const sh of await xlsxRows(buf)){ const t=rowsToCardTx(sh); if(!t.length) continue; miss+=missedRows(sh,t);
        t.forEach(x=>{ const id=[x.k,x.ts,x.m,x.v,x.ap].join('|'); if(!seen.has(id)){ seen.add(id); tx.push(x); } }); rows=rows.concat(rows.length?[[]]:[],sh); } }
    else if(h[0]===0xD0&&h[1]===0xCF) throw new Error('예전 엑셀(.xls)입니다. .xlsx로 저장하거나 복사해 붙여넣어 주세요');
    else { const txt=decodeText(buf); rows=textRows(txt); tx=rowsToCardTx(rows); const fz=fuzzyCardTx(txt); if(fz.length>tx.length) tx=fz; else miss=missedRows(rows,tx); }
    if(!tx.length) throw new Error('결제 내역을 찾지 못했습니다 (날짜·가맹점·금액 칸이 있는지 확인)');
    return {tx,rows,miss};
}
M.card={BIKE_MAX,LPG_RE,STATION_RE,classifyCard,isBigFuel,netCardTx,normM,pasteTx,readCardFile};
})(window.M=window.M||{});
