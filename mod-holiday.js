/* mod-holiday.js — 공휴일 · 음력 계산 (한국 공휴일 엔진)
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.holiday로 내보냄.
   MODULE holiday V1.00 (2026-10-26) */
(function(M){
const {addDays,keyOf,parseKey}=M.base;
/* ================= 공휴일 엔진 =================
   ① 양력 고정 공휴일  ② 음력 공휴일(설·추석·부처님오신날, 기기 내장 음력 달력으로 계산)
   ③ 선거일·임시공휴일  ④ 대체공휴일 규칙 자동 적용  ⑤ 사용자 지정 휴일                 */
const FIXED=[ // [월,일,이름,대체공휴일 적용,시행연도]
  [1,1,'신정',false],[3,1,'삼일절',true],[5,1,'노동절',true,2026],[5,5,'어린이날',true],
  [6,6,'현충일',false],[7,17,'제헌절',true,2026],[8,15,'광복절',true],[10,3,'개천절',true],
  [10,9,'한글날',true],[12,25,'성탄절',true]
];
const EXTRA={ // 선거일·임시공휴일 (확정분)
  '2024-04-10':'국회의원선거','2024-10-01':'임시공휴일','2025-01-27':'임시공휴일',
  '2025-06-03':'대통령선거','2026-06-03':'지방선거','2028-04-12':'국회의원선거'
};
const LUNAR_TABLE={ // 기기에서 음력 계산이 안 될 때 쓰는 예비표 [설날, 부처님오신날, 추석]
  2024:['2-10','5-15','9-17'],2025:['1-29','5-5','10-6'],2026:['2-17','5-24','9-25'],2027:['2-7','5-13','9-15'],
  2028:['1-27','5-2','10-3'],2029:['2-13','5-20','9-22'],2030:['2-3','5-9','9-12'],2031:['1-23','5-28','10-1'],
  2032:['2-11','5-16','9-19'],2033:['1-31','5-6','9-8'],2034:['2-19','5-25','9-27'],2035:['2-8','5-15','9-16']
};
let lunarFmt=null;
try{
  const f=new Intl.DateTimeFormat('ko-KR-u-ca-dangi',{month:'numeric',day:'numeric'});
  if(f.resolvedOptions().calendar==='dangi') lunarFmt=f;
}catch(e){}
function lunarDates(y){
  if(lunarFmt){
    const r={};
    for(let d=new Date(y,0,1); d.getFullYear()===y; d.setDate(d.getDate()+1)){
      const p=lunarFmt.formatToParts(d);
      const m=(p.find(x=>x.type==='month')||{}).value, dd=(p.find(x=>x.type==='day')||{}).value;
      const k=m+'/'+dd;
      if(k==='1/1') r.seol=new Date(d); else if(k==='4/8') r.buddha=new Date(d); else if(k==='8/15') r.chuseok=new Date(d);
    }
    if(r.seol&&r.buddha&&r.chuseok) return r;
  }
  const t=LUNAR_TABLE[y]; if(!t) return null;
  const mk=s=>{ const [m,d]=s.split('-').map(Number); return new Date(y,m-1,d); };
  return {seol:mk(t[0]),buddha:mk(t[1]),chuseok:mk(t[2])};
}
const holCache={};
function holidaysOf(y){
  if(holCache[y]) return holCache[y];
  const base={};   // key → [{name,type}]  type: fix | sub | grp
  const add=(d,name,type,grp)=>{ const k=keyOf(d); (base[k]=base[k]||[]).push({name,type,grp}); };
  FIXED.forEach(([m,d,n,sub,from])=>{ if(!from||y>=from) add(new Date(y,m-1,d),n,sub?'sub':'fix'); });
  const L=lunarDates(y);
  if(L){
    [['seol','설날'],['chuseok','추석']].forEach(([g,n])=>{
      add(addDays(L[g],-1),n+' 연휴','grp',g); add(L[g],n,'grp',g); add(addDays(L[g],1),n+' 연휴','grp',g);
    });
    add(L.buddha,'부처님오신날','sub');
  }
  for(const k in EXTRA) if(k.startsWith(y+'-')) (base[k]=base[k]||[]).push({name:EXTRA[k],type:'fix'});

  /* 대체공휴일 규칙
     - 설·추석 연휴: 일요일 또는 다른 공휴일과 겹치면 연휴 다음 첫 평일
     - 삼일절·노동절·어린이날·제헌절·광복절·개천절·한글날·성탄절·부처님오신날:
       토·일요일 또는 다른 공휴일과 겹치면 다음 첫 평일
     - 신정·현충일·선거일: 대체공휴일 없음                                  */
  const out={}; for(const k in base) out[k]=base[k].map(x=>x.name);
  const taken=new Set(Object.keys(base));
  const jobs=[];
  Object.keys(base).sort().forEach(k=>{
    const items=base[k], d=parseKey(k), dow=d.getDay();
    const elig=items.filter(x=>x.type==='sub').length;
    const grps=[...new Set(items.filter(x=>x.type==='grp').map(x=>x.grp))];
    let n;
    if(dow===0) n=elig+grps.length;
    else if(dow===6) n=elig;
    else n=Math.min(items.length-1, elig+grps.length);
    if(n<=0) return;
    let after=d;
    if(grps.length && L){ grps.forEach(g=>{ const last=addDays(L[g],1); if(last>after) after=last; }); }
    for(let i=0;i<n;i++) jobs.push(after);
  });
  jobs.sort((a,b)=>a-b).forEach(after=>{
    let d=addDays(after,1);
    while(d.getDay()===0||d.getDay()===6||taken.has(keyOf(d))) d=addDays(d,1);
    taken.add(keyOf(d)); (out[keyOf(d)]=out[keyOf(d)]||[]).push('대체공휴일');
  });
  return holCache[y]=out;
}
function holNames(k,userName){            // userName: 사용자가 직접 적은 그날 휴일 이름 (있으면 덧붙임)
  const names=[...(holidaysOf(+k.slice(0,4))[k]||[])];
  if(userName) names.push(userName);
  return names;
}
M.holiday={holNames,holidaysOf};
})(window.M=window.M||{});
