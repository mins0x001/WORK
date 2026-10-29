/* mod-stats.js — 업무일지 기록 모델 · 근무 기간 · 집계 (순수 계산)
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.stats로 내보냄.
   MODULE stats V1.17 (2026-10-29) */
(function(M){
const {$,CARD_CATS,addDays,keyOf}=M.base;
const {holNames}=M.holiday;
const {routeInfo}=M.dealer;
/* ================= 업무일지 기록 모델 · 근무 기간 · 집계 =================
   data(날짜별 기록)를 매개변수로 받는 순수 계산만 담았습니다. 화면·저장소와 무관합니다.   */
const SHIFT={
  '1':  {label:'1회전',cls:'t1',color:'var(--t1)'},
  '2':  {label:'2회전',cls:'t2',color:'var(--t2)'},
  '3':  {label:'3회전',cls:'t3',color:'var(--t3)'},
  'off':{label:'휴무', cls:'off',color:'var(--off)'}
};
const NUMF=['odo','fuel','toll','urea'];
const ROUNDS=['1','2','3'];
const MAX_STOPS=40;
const RENAMED={'충남(의정부)':'충남','의충남':'충남','충남(장안)':'장충남','C동우':'동우'};      // 이름이 바뀐 대리점: 예전 이름으로 저장된 회전 기록·검색어를 새 이름으로

function normRounds(src){
  if(!src||typeof src!=='object') return null;
  const out={};
  ROUNDS.forEach(n=>{
    const t=src[n]; if(!t||!Array.isArray(t.st)) return;
    const st=[];
    t.st.forEach(x=>{
      if(typeof x==='string') x={n:x};
      if(!x||typeof x.n!=='string'||!x.n.trim()||st.length>=MAX_STOPS) return;
      const nm=x.n.trim().slice(0,20);
      const o={n:RENAMED[nm]||nm};
      if(typeof x.c==='string'&&x.c.trim()) o.c=x.c.trim().slice(0,30);
      st.push(o);
    });
    if(st.length) out[n]={st};
  });
  return Object.keys(out).length?out:null;
}
/* cd = 그날 반영한 카드 결제 [{id:'날짜#승인번호', tm:'17:32', m:'가맹점', v:금액, c:'fuel'|'urea'|'toll'}] */
function normCard(a){
  if(!Array.isArray(a)) return null;
  const out=[], seen=new Set();
  a.forEach(e=>{
    if(!e||typeof e.id!=='string'||!e.id||seen.has(e.id)||!CARD_CATS[e.c]) return;
    const v=Math.round(Number(e.v)); if(!Number.isFinite(v)||v<=0) return;
    seen.add(e.id);
    const o={id:e.id.slice(0,80),m:String(e.m||'').trim().slice(0,40),v,c:e.c};
    if(typeof e.tm==='string'&&/^\d{2}:\d{2}$/.test(e.tm)) o.tm=e.tm;
    out.push(o);
  });
  return out.length?out.slice(0,40):null;
}
function normalize(src){
  const out={};
  for(const k in src){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(k)) continue;
    let v=src[k];
    if(typeof v==='string') v={s:v};
    if(!v||typeof v!=='object') continue;
    const r={};
    if(SHIFT[v.s]) r.s=v.s;
    NUMF.forEach(f=>{ const n=Number(v[f]); if(Number.isFinite(n)&&n>0) r[f]=Math.round(n); });
    if(typeof v.hol==='string'&&v.hol.trim()) r.hol=v.hol.trim().slice(0,12);
    const rr=normRounds(v.r); if(rr) r.r=rr;
    const cd=normCard(v.cd); if(cd) r.cd=cd;
    if(Object.keys(r).length) out[k]=r;
  }
  return out;
}
/* ================= 근무 기간 ================= */
function periodOf(d){ let y=d.getFullYear(),m=d.getMonth(); if(d.getDate()>=25){ m++; if(m>11){m=0;y++;} } return {y,m}; }
function rangeOf(p){ return {start:new Date(p.y,p.m-1,25), end:new Date(p.y,p.m,24)}; }

/* ---- 회전 기록 읽기 ---- */
function roundsIn(data,k){
  const r=(data[k]||{}).r||{};
  return ROUNDS.filter(n=>r[n]&&r[n].st&&r[n].st.length).map(n=>({n,st:r[n].st,names:r[n].st.map(s=>s.n)}));
}
const dayRouteKmIn=(data,k)=>roundsIn(data,k).reduce((x,r)=>x+routeInfo(r.names).km,0);
const roundTextIn=(data,k,n)=>{ const t=((data[k]||{}).r||{})[n]; return t?t.st.map(s=>s.n).join(' → '):''; };
/* ---- 집계 엔진 ---- */
function statsFor(data,p){
  const {start,end}=rangeOf(p);
  const c={'1':0,'2':0,'3':0,off:0,none:0,days:0,fuel:0,toll:0,urea:0,fuelN:0,tollN:0,ureaN:0,hol:0,trips:0,stops:0,rkm:0,runk:0};
  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    const k=keyOf(d), r=data[k]||{}; c.days++;
    if(r.s) c[r.s]++; else c.none++;
    if(holNames(k,(data[k]||{}).hol).length) c.hol++;
    ['fuel','toll','urea'].forEach(f=>{ if(r[f]){ c[f]+=r[f]; c[f+'N']++; } });
    roundsIn(data,k).forEach(x=>{ const i=routeInfo(x.names); c.trips++; c.stops+=x.st.length; c.rkm+=i.km; c.runk+=i.unk; });
  }
  c.rkm=Math.round(c.rkm*10)/10;
  c.work=c['1']+c['2']+c['3']; c.cost=c.fuel+c.toll+c.urea; c.odo=odoFor(data,p);
  return c;
}
function odoFor(data,p){
  const {start,end}=rangeOf(p);
  const sK=keyOf(start), eK=keyOf(end), pK=keyOf(addDays(start,-1));
  let sv=data[sK]&&data[sK].odo, sSrc='start';
  if(!sv&&data[pK]&&data[pK].odo){ sv=data[pK].odo; sSrc='prev'; }
  const ev=data[eK]&&data[eK].odo;
  const o={start:sv||null,end:ev||null,sSrc,sK,eK,pK,dist:null,state:'none'};
  if(sv&&ev){ o.dist=ev-sv; o.state=o.dist<0?'neg':'ok'; }
  else if(sv) o.state='noEnd'; else if(ev) o.state='noStart';
  return o;
}
M.stats={MAX_STOPS,RENAMED,ROUNDS,SHIFT,dayRouteKmIn,normalize,odoFor,periodOf,rangeOf,roundTextIn,roundsIn,statsFor};
})(window.M=window.M||{});
