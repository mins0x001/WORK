/* mod-fare.js — 운송비 명세서 읽기(엑셀·CSV·붙여넣기·OCR 글자) · 검산 · 평균·추세 분석
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.fare로 내보냄.
   MODULE fare V1.16 (2026-10-29) */
(function(M){
const {$,bus,ls,metaOf,rj}=M.base;
const {xlsxRows}=M.xlsxr;
/* 운송비 명세서: 읽기(엑셀·CSV·붙여넣기·사진 글자) → 검산 → 평균·추세·적정 운임 분석. 화면과 무관한 순수 계산 + 저장소.
   명세서 한 장 = {id,ym,name,veh,work,co,days,kmFrom,kmTo,kmDelta,pay:[{n,v,k}],sub,vat,supply,payTotal,ded:[{n,v}],dedTotal,net,src,at}
   (표시된 값은 그대로 저장하고, 검산·계산값은 fareDerive()가 읽을 때마다 새로 만든다 → 원본은 절대 고쳐지지 않음) */
const FARE_KEY='fare.data.v1';
const FARE_META='fare.sync.meta';
const FARE_SET='fare.settings.v1';
const FARE_ROLES=[   // 이름표 → 역할 (위에서부터 먼저 맞는 것). 띄어쓰기·괄호·기호는 무시하고 비교
  ['net',/차인|실수령|실지급|입금액|수령액/],
  ['dedTotal',/공제(액)?(계|합계|총액|총계)$|^총?공제/],
  ['payTotal',/지급(액)?(계|합계|총액|총계)$|총지급|^지급액$/],
  ['supply',/세금계산서|발행금액|공급가액|공급가/],
  ['vat',/부가세|부가가치세|^vat$/i],
  ['sub',/^(소계|합계|계|합)$/],
  ['ded',/공제|차감|환수/]
];
const FARE_KIND=[   // 지급 항목 종류 (분석용 묶음)
  ['base',/당월지급|운임|운송료|기본|배송료|지급액$/],
  ['hipass',/하이패스|통행|톨/],
  ['allow',/수당|직계약|인센/],
  ['fuel',/유류|유가|연료|경유/]
];
const FARE_KIND_NAME={base:'기본 운임',hipass:'하이패스 지원',allow:'수당',fuel:'유류변동',etc:'기타'};
const frNorm=s=>String(s==null?'':s).replace(/[\s()（）\[\]{}<>·ㆍ.:：,*\-_/|]+/g,'');
const frNum=s=>{ const t=String(s).replace(/[,\s원₩]/g,''); return /^-?\d+(\.\d+)?$/.test(t)?Number(t):null; };
function fareRole(label){
  const n=frNorm(label);
  if(!n) return '';
  for(const [r,re] of FARE_ROLES) if(re.test(n)) return r;
  return 'pay';
}
function fareKind(label){ const n=frNorm(label); for(const [k,re] of FARE_KIND) if(re.test(n)) return k; return 'etc'; }
const FARE_NOISE=/^(지급내역|공제내역|지급액|공제액|내역|구분|항목|금액|운행거리|비고|증가|운송비|운송비명세서|명세서|명세|km)$|명세서|귀하|감사|제이엠|주식회사/;
const FARE_MONTH_RE=/(20\d{2})\s*(?:년|[.\-/])\s*(\d{1,2})\s*(?:월|(?=\D|$))/;
function fareYm(s){
  const m=String(s).match(FARE_MONTH_RE); if(m&&+m[2]>=1&&+m[2]<=12) return {ym:m[1]+'-'+String(m[2]).padStart(2,'0'),at:m.index,len:m[0].length};
  const k=String(s).match(/(?:^|\D)(\d{1,2})\s*월\s*(?:분\s*)?(?:운송비|명세|정산|지급)/);
  if(k&&+k[1]>=1&&+k[1]<=12){ const n=new Date(), y=(+k[1]>n.getMonth()+2)?n.getFullYear()-1:n.getFullYear(); return {ym:y+'-'+String(k[1]).padStart(2,'0'),at:k.index,len:k[0].length,guess:true}; }
  return null;
}
/* ---- 입력 → 글자 (엑셀 칸은 탭으로 이어 붙임) ---- */
function fareText(input){
  if(Array.isArray(input)) return input.map(r=>(Array.isArray(r)?r:[r]).map(c=>c==null?'':String(c).trim()).filter(c=>c!=='').join('\t')).filter(Boolean).join('\n');
  return String(input==null?'':input).replace(/\r\n?/g,'\n').replace(/[​﻿]/g,'');
}
/* ---- 한 줄 → [{label, v}] : '이름 금액 이름 금액' 어떤 배치든 순서대로 짝지음 ---- */
function frPairs(line,st){
  const out=[]; let rest=line;
  rest=rest.replace(/(\d[\d,]*)\s*km?\s*[-~–—→>]+\s*(\d[\d,]*)\s*km/gi,' ').replace(/(\d[\d,]*)\s*km/gi,' ')       // 거리
    .replace(/(20\d{2})\s*(?:년|[.\-/])\s*\d{1,2}\s*월?/g,' ').replace(/\d+\s*(?:일|월|년|대|건|회|%)/g,' ')           // 날짜·일수
    .replace(/[가-힣]{2}\s?\d{2,3}\s?[가-힣]\s?\d{4}/g,' ');                                                     // 차량번호
  const toks=rest.split(/(-?\d[\d,]*(?:\.\d+)?)/);
  let lab='';
  for(let i=0;i<toks.length;i++){
    if(i%2===0){ const t=toks[i].replace(/[\t:：|]+/g,' ').trim(); if(t) lab=(lab?lab+' ':'')+t; }
    else{ const v=frNum(toks[i]); if(v==null) continue; if(!/,/.test(toks[i])&&Math.abs(v)<100&&!lab) continue; out.push({label:lab,v}); lab=''; }
  }
  if(lab) out.push({label:lab,v:null});
  return out;
}
/* ---- 글자 → 항목 목록 (이름만 있는 줄 → 금액만 있는 줄이 이어지는 '세로 배치'도 순서대로 짝지음) ---- */
function frEntries(text){
  const ents=[], pend=[];
  for(const line of text.split('\n')){
    for(const p of frPairs(line)){
      if(p.v==null){ const n=frNorm(p.label); if(n.length>=1&&n.length<=24&&!FARE_NOISE.test(n)) pend.push(p.label); continue; }
      if(p.label){ ents.push({label:p.label,v:p.v}); continue; }
      if(pend.length) ents.push({label:pend.shift(),v:p.v}); else ents.push({label:'',v:p.v});
    }
  }
  return ents;
}
/* ---- 항목 목록 → 명세서 한 장 ---- */
function frBuild(text,ymHint){
  const st={ym:ymHint||'',name:'',veh:'',work:'',co:'',days:null,kmFrom:null,kmTo:null,kmDelta:null,pay:[],sub:null,vat:null,supply:null,payTotal:null,ded:[],dedTotal:null,net:null};
  st.raw=text.slice(0,4000);
  const y=fareYm(text); if(!st.ym&&y) st.ym=y.ym;
  let m;
  if((m=text.match(/성\s*명\s*[:：]?[ \t]*([^\n\t]*)/))){ const nm=m[1].split(/\s{2,}/)[0].replace(/(운행|차량|업무).*/,'').replace(/\s/g,'').match(/^[가-힣]{2,4}/); if(nm) st.name=nm[0]; }
  if((m=text.match(/([가-힣]{2}\s?\d{2,3}\s?[가-힣]\s?\d{4})/))) st.veh=m[1].replace(/\s/g,'');
  if((m=text.match(/업무\s*구분\s*[:：]?\s*([^\n\t]+)/))) st.work=m[1].replace(/\s{2,}.*/,'').trim();
  if((m=text.match(/(?:주식회사|\(주\)|㈜)\s*([가-힣A-Za-z0-9]+(?:\s?[가-힣A-Za-z0-9]+)?)/))) st.co=('주식회사 '+m[1]).trim();
  if((m=text.match(/운행\s*일수\s*[:：]?\s*(\d{1,2})/))) st.days=+m[1];
  if((m=text.match(/(\d[\d,]*)\s*km?\s*[-~–—→>]+\s*(\d[\d,]*)\s*km/i))){ st.kmFrom=frNum(m[1]); st.kmTo=frNum(m[2]); }
  if((m=text.match(/(\d[\d,]*)\s*km\s*(?:증가|운행|주행)/i))) st.kmDelta=frNum(m[1]);
  const body=text.split('\n').filter(l=>(!/성\s*명|차량\s*번호|업무\s*구분|운행\s*일수|귀하의|감사|주식회사/.test(l)||/\d{1,3},\d{3}/.test(l))&&!(/km/i.test(l)&&!/\d{1,3},\d{3}/.test(l))).join('\n');
  const ents=frEntries(body);
  let seenSub=false;
  ents.forEach((e,i)=>{
    const r=e.label?fareRole(e.label):'';
    if(r==='net') st.net=e.v;
    else if(r==='dedTotal') st.dedTotal=e.v;
    else if(r==='payTotal') st.payTotal=e.v;
    else if(r==='supply') st.supply=e.v;
    else if(r==='vat') st.vat=e.v;
    else if(r==='sub'){ if(st.sub==null) st.sub=e.v; seenSub=true; }
    else if(r==='ded') st.ded.push({n:e.label.replace(/\s+/g,' ').trim(),v:e.v});
    else if(r==='pay') st.pay.push({n:e.label.replace(/\s+/g,' ').trim(),v:e.v,k:fareKind(e.label)});
    else if(!e.label){ if(!seenSub&&st.sub==null) st.pay.push({n:'(이름 없음)',v:e.v,k:'etc'}); }
  });
  // 이름 없는 금액 정리: 지급 항목 뒤 합계 자리에 놓인 값은 합계로 인정하지 않고 검산에서 알려줌
  const noName=st.pay.filter(p=>p.n==='(이름 없음)');
  if(noName.length&&st.sub==null){ const s=st.pay.filter(p=>p.n!=='(이름 없음)').reduce((a,p)=>a+p.v,0); const hit=noName.find(p=>p.v===s); if(hit){ st.sub=hit.v; st.pay=st.pay.filter(p=>p!==hit); } }
  return st;
}
const frHas=st=>st.pay.length+st.ded.length>0||st.net!=null||st.payTotal!=null;
/* ---- 가로 표 (한 줄 = 한 달, 칸 이름이 맨 위 줄) 읽기 ---- */
function frTable(rows){
  let hi=-1, cols=null;
  for(let i=0;i<Math.min(rows.length,12);i++){
    const r=rows[i]||[], map=r.map(c=>typeof c==='string'?fareRole(c):'');
    const hit=map.filter(x=>x&&x!=='pay').length, hasMonth=r.some(c=>/월|연월|기간|년월/.test(String(c)));
    if(hit>=4&&(hasMonth||r.some(c=>/지급|운임/.test(String(c))))){ hi=i; cols=r.map(c=>String(c==null?'':c)); break; }
  }
  if(hi<0) return null;
  const out=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[]; let ym='';
    for(const c of r){ if(typeof c==='number') continue; const s=String(c); const y=fareYm(s+' '); if(y){ ym=y.ym; break; } const m=s.match(/^(20\d{2})(\d{2})$/); if(m&&+m[2]>=1&&+m[2]<=12){ ym=m[1]+'-'+m[2]; break; } }
    if(!ym){ r.forEach((c,j)=>{ if(!ym&&typeof c==='number'&&c>=40000&&c<=60000&&/연월|년월|기간|일자|날짜|월분/.test(cols[j]||'')){ const d=new Date(Math.round((c-25569)*86400000)); ym=d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0'); } }); }
    if(!ym) continue;
    const st=frBuild('',ym); st.raw=r.map(c=>c==null?'':c).join(' | ').slice(0,1000);
    cols.forEach((h,j)=>{
      const v=typeof r[j]==='number'?r[j]:frNum(r[j]); if(v==null||!h) return;
      const role=fareRole(h);
      if(/일수/.test(h)){ st.days=v; return; }
      if(/시작.*km|전월.*km|이전/i.test(h)){ st.kmFrom=v; return; }
      if(/종료.*km|당월.*km|현재/i.test(h)){ st.kmTo=v; return; }
      if(role==='net') st.net=v; else if(role==='dedTotal') st.dedTotal=v; else if(role==='payTotal') st.payTotal=v;
      else if(role==='supply') st.supply=v; else if(role==='vat') st.vat=v; else if(role==='sub') st.sub=v;
      else if(role==='ded') st.ded.push({n:String(h).trim(),v}); else if(role==='pay'&&!/^(연월|년월|월|기간|번호|no|차량)/i.test(String(h).trim())) st.pay.push({n:String(h).trim(),v,k:fareKind(h)});
    });
    if(frHas(st)) out.push(st);
  }
  return out.length?out:null;
}
/* ---- 공개: 입력(엑셀 칸들·글자) → 명세서 목록 ---- */
function fareParse(input,srcName){
  const sheets=Array.isArray(input)&&input.length&&Array.isArray(input[0])&&Array.isArray(input[0][0])?input:null;   // 엑셀 시트 여러 개
  const out=[], warn=[];
  const parts=sheets||[input];
  for(const part of parts){
    if(Array.isArray(part)){ const t=frTable(part); if(t){ t.forEach(s=>{ s.src=srcName||''; out.push(s); }); continue; } }
    const text=fareText(part), lines=text.split('\n'), starts=[];
    lines.forEach((l,i)=>{ const y=fareYm(l); if(y&&(/명세서|운송비|정산/.test(l)||/^\s*20\d\d\s*(년|[.\-])\s*\d{1,2}\s*월?\s*$/.test(l))&&(!starts.length||starts[starts.length-1].ym!==y.ym)) starts.push({i,ym:y.ym}); });
    if(!starts.length){ const st=frBuild(text,''); if(frHas(st)){ st.src=srcName||''; out.push(st); } continue; }
    // 첫 월 머리글 앞의 줄(성명·차량번호 등)도 첫 명세서에 포함
    starts.forEach((s,k)=>{ const a=k===0?0:s.i, b=k+1<starts.length?starts[k+1].i:lines.length; const st=frBuild(lines.slice(a,b).join('\n'),s.ym); if(frHas(st)){ st.src=srcName||''; out.push(st); } });
  }
  out.forEach(s=>{ s.id=s.ym+'|'+(s.veh||''); });
  return out;
}
/* ---- 검산 ---- */
const FARE_TOL=1;                                   // 원 단위 반올림 허용
const frSum=a=>a.reduce((x,p)=>x+(+p.v||0),0);
function fareDerive(s){
  const d={paySum:frSum(s.pay),dedSum:frSum(s.ded)};
  d.sub=s.sub!=null?s.sub:(s.supply!=null?s.supply:d.paySum);
  d.vatExp=Math.round(d.sub*0.1);
  d.vat=s.vat!=null?s.vat:d.vatExp;
  d.payTotal=s.payTotal!=null?s.payTotal:d.sub+d.vat;
  d.dedTotal=s.dedTotal!=null?s.dedTotal:d.dedSum;
  d.net=s.net!=null?s.net:d.payTotal-d.dedTotal;
  d.km=s.kmDelta!=null?s.kmDelta:(s.kmFrom!=null&&s.kmTo!=null?s.kmTo-s.kmFrom:null);
  const by={}; s.pay.forEach(p=>{ by[p.k||'etc']=(by[p.k||'etc']||0)+p.v; }); d.by=by;
  d.base=by.base||0; d.fuel=by.fuel||0;
  d.perDay=s.days?d.sub/s.days:null;                // 공급가액 일평균
  d.basePerDay=s.days&&d.base?d.base/s.days:null;   // 기본 운임 일평균
  d.netPerDay=s.days?d.net/s.days:null;
  d.perKm=d.km?d.sub/d.km:null;
  d.fuelPerKm=d.km&&d.fuel?d.fuel/d.km:null;
  d.dedRate=d.payTotal?d.dedTotal/d.payTotal:null;
  return d;
}
function fareCheck(s){
  const d=fareDerive(s), L=[], won=n=>Math.abs(n).toLocaleString('ko-KR')+'원';
  const add=(lv,t,msg,diff)=>L.push({lv,t,msg,diff:diff||0});
  if(s.pay.length&&s.sub!=null){ const df=s.sub-d.paySum; if(Math.abs(df)>FARE_TOL) add('bad','지급 합계','지급 항목을 더하면 '+won(d.paySum)+' 인데 계는 '+won(s.sub)+' → '+(df>0?'항목이 빠졌을 수 있음 ':'계가 '+won(df)+' 적음 ')+(df>0?'('+won(df)+')':''),df); else add('ok','지급 합계','항목 합계 = 계'); }
  else if(s.sub==null) add('warn','지급 합계','계 금액이 없어 항목 합('+won(d.paySum)+')으로 계산했습니다');
  if(s.supply!=null&&s.sub!=null){ const df=s.supply-s.sub; if(Math.abs(df)>FARE_TOL) add('bad','세금계산서','발행금액 '+won(s.supply)+' ≠ 계 '+won(s.sub)+' (차이 '+won(df)+')',df); else add('ok','세금계산서','발행금액 = 계'); }
  if(s.vat!=null){ const df=s.vat-d.vatExp; if(Math.abs(df)>FARE_TOL) add(df<0?'bad':'warn','부가세','10%는 '+won(d.vatExp)+' 인데 '+won(s.vat)+' → '+(df<0?'적게':'많이')+' 계산됨 ('+won(df)+')',df); else add('ok','부가세','계의 10%와 일치'); }
  else add('warn','부가세','부가세 금액이 없어 10%('+won(d.vatExp)+')로 계산했습니다');
  if(s.payTotal!=null){ const df=s.payTotal-(d.sub+d.vat); if(Math.abs(df)>FARE_TOL) add(df<0?'bad':'warn','지급액 계','계+부가세는 '+won(d.sub+d.vat)+' 인데 지급액 계는 '+won(s.payTotal)+' ('+(df<0?'적음 ':'많음 ')+won(df)+')',df); else add('ok','지급액 계','계 + 부가세 = 지급액 계'); }
  if(s.ded.length&&s.dedTotal!=null){
    const df=d.dedSum-s.dedTotal;
    if(Math.abs(df)>FARE_TOL){ const hit=s.ded.find(p=>Math.abs(p.v-df)<=FARE_TOL); const sub2=frSubsetEq(s.ded,s.dedTotal);
      add('warn','공제 합계','공제 항목 합 '+won(d.dedSum)+' ≠ 공제액 계 '+won(s.dedTotal)+' (차이 '+won(df)+')'+(hit?' → 「'+hit.n+'」 '+won(hit.v)+'이 계에서 빠진 것으로 보임':sub2?' → '+sub2.map(p=>'「'+p.n+'」').join('·')+'만 합쳐야 계와 같음':'')+' · 안내용 표기인지 확인',df); }
    else add('ok','공제 합계','공제 항목 합 = 공제액 계');
  }
  if(s.net!=null){ const df=s.net-(d.payTotal-d.dedTotal); if(Math.abs(df)>FARE_TOL) add(df<0?'bad':'warn','차인지급액','지급액 계 − 공제액 계는 '+won(d.payTotal-d.dedTotal)+' 인데 차인지급액은 '+won(s.net)+' ('+(df<0?'적음 ':'많음 ')+won(df)+')',df); else add('ok','차인지급액','지급액 계 − 공제액 계와 일치'); }
  if(s.kmFrom!=null&&s.kmTo!=null){ if(s.kmDelta!=null&&s.kmTo-s.kmFrom!==s.kmDelta) add('bad','운행거리','끝−시작='+(s.kmTo-s.kmFrom)+'km 인데 증가분은 '+s.kmDelta+'km'); else if(s.kmTo<s.kmFrom) add('bad','운행거리','끝 거리가 시작보다 작음'); else add('ok','운행거리',(s.kmTo-s.kmFrom)+'km'); }
  const miss=[]; if(s.days==null) miss.push('운행일수'); if(d.km==null) miss.push('운행거리'); if(s.net==null) miss.push('차인지급액'); if(s.payTotal==null) miss.push('지급액 계'); if(!s.ym) miss.push('연월'); if(!s.pay.length) miss.push('지급 항목');
  if(miss.length) add('warn','누락',miss.join('·')+' 값을 찾지 못했습니다 (원본을 확인하세요)');
  return L;
}
function frSubsetEq(items,target){      // 항목 중 일부를 더해서 target이 되는 조합 (항목이 적을 때만)
  const n=items.length; if(n>12) return null;
  for(let m=1;m<(1<<n)-1;m++){ let s=0; const pick=[]; for(let i=0;i<n;i++) if(m&(1<<i)){ s+=items[i].v; pick.push(items[i]); } if(Math.abs(s-target)<=FARE_TOL) return pick; }
  return null;
}
/* ---- 여러 달 비교: 누락 월 · 거리 연속 · 평균 · 적정 운임 ---- */
const frAvg=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;
const frNextYm=ym=>{ const [y,m]=ym.split('-').map(Number); return m===12?(y+1)+'-01':y+'-'+String(m+1).padStart(2,'0'); };
function fareMonths(list){ const yms=[...new Set(list.map(s=>s.ym).filter(Boolean))].sort(); return yms; }
function fareGaps(list){
  const yms=fareMonths(list), out=[]; if(yms.length<2) return out;
  let c=yms[0]; while(c<yms[yms.length-1]){ c=frNextYm(c); if(!yms.includes(c)&&c<yms[yms.length-1]) out.push(c); }
  return out;
}
function fareAnalyze(list,opt){
  opt=opt||{};
  const L=list.filter(s=>s.ym).slice().sort((a,b)=>a.ym<b.ym?-1:a.ym>b.ym?1:0);
  const rows=L.map(s=>Object.assign({s},{d:fareDerive(s),issues:fareCheck(s)}));
  const refDaily=+opt.refDaily||0, tol=opt.tol==null?0.03:opt.tol;
  rows.forEach((r,i)=>{
    const prev=rows[i-1]; r.flags=[];
    if(prev){
      r.dSub=r.d.sub-prev.d.sub; r.dNet=r.d.net-prev.d.net; r.dBase=r.d.base-prev.d.base;
      if(r.s.kmFrom!=null&&prev.s.kmTo!=null&&frNextYm(prev.s.ym)===r.s.ym&&r.s.kmFrom!==prev.s.kmTo) r.flags.push({lv:'warn',t:'거리 불연속',msg:'지난달 끝 '+prev.s.kmTo+'km ≠ 이번달 시작 '+r.s.kmFrom+'km → 중간 거리가 누락됐을 수 있음'});
    }
    const others=rows.filter((x,j)=>j!==i&&x.d.basePerDay!=null);
    const avgBase=frAvg(others.map(x=>x.d.basePerDay));
    r.refBasis=refDaily?{daily:refDaily,label:'내 기준 일당'}:(avgBase!=null?{daily:avgBase,label:'다른 달 평균 일당('+others.length+'개월)'}:null);
    if(r.refBasis&&r.s.days&&r.d.base){
      r.expBase=Math.round(r.refBasis.daily*r.s.days);
      r.gap=r.d.base-r.expBase;
      r.gapPct=r.expBase?r.gap/r.expBase:0;
      if(r.gapPct<-tol) r.flags.push({lv:'bad',t:'적게 계산?',msg:'기본 운임 '+Math.abs(r.gap).toLocaleString('ko-KR')+'원('+(r.gapPct*100).toFixed(1)+'%) 적음 — '+r.refBasis.label+' × '+r.s.days+'일 = '+r.expBase.toLocaleString('ko-KR')+'원 기준'});
      else if(r.gapPct>tol) r.flags.push({lv:'ok',t:'평소보다 많음',msg:'기준보다 '+r.gap.toLocaleString('ko-KR')+'원(+'+(r.gapPct*100).toFixed(1)+'%) 많음'});
    }
    if(r.issues.some(x=>x.lv==='bad')) r.flags.unshift({lv:'bad',t:'검산 불일치',msg:r.issues.filter(x=>x.lv==='bad').map(x=>x.t).join('·')});
  });
  const pick=(f,a)=>a.map(f).filter(v=>v!=null&&isFinite(v));
  const sumA=a=>({avg:frAvg(a),min:a.length?Math.min(...a):null,max:a.length?Math.max(...a):null,n:a.length});
  const rec3=rows.slice(-3);
  const agg=fn=>({all:sumA(pick(fn,rows)),r3:sumA(pick(fn,rec3))});
  const tot={sub:rows.reduce((x,r)=>x+r.d.sub,0),net:rows.reduce((x,r)=>x+r.d.net,0),days:rows.reduce((x,r)=>x+(r.s.days||0),0),km:rows.reduce((x,r)=>x+(r.d.km||0),0),ded:rows.reduce((x,r)=>x+r.d.dedTotal,0)};
  return {rows,gaps:fareGaps(rows.map(r=>r.s)),tot,
    avg:{sub:agg(r=>r.d.sub),net:agg(r=>r.d.net),perDay:agg(r=>r.d.perDay),basePerDay:agg(r=>r.d.basePerDay),netPerDay:agg(r=>r.d.netPerDay),perKm:agg(r=>r.d.perKm),fuelPerKm:agg(r=>r.d.fuelPerKm),fuel:agg(r=>r.d.fuel),days:agg(r=>r.s.days),km:agg(r=>r.d.km),dedRate:agg(r=>r.d.dedRate)}};
}
/* ---- 저장소 (같은 달·차량이면 덮어쓰기, 지우면 동기화에 삭제로 기록) ---- */
const F=rj(FARE_KEY,{});
function farePut(list){
  const m=metaOf(FARE_META), now=Date.now(); let add=0, upd=0;
  list.forEach(s=>{ if(!s.ym) return; const id=s.ym+'|'+(s.veh||''); s.id=id; s.at=now; if(F[id]) upd++; else add++; F[id]=s; m.mt[id]=now; delete m.del[id]; });
  m.dirty=true; ls.set(FARE_META,JSON.stringify(m)); ls.set(FARE_KEY,JSON.stringify(F)); bus.emit('fare:changed'); return {add,upd};
}
function fareDel(ids){
  const m=metaOf(FARE_META), now=Date.now();
  ids.forEach(id=>{ delete F[id]; m.del[id]=now; delete m.mt[id]; });
  m.dirty=true; ls.set(FARE_META,JSON.stringify(m)); ls.set(FARE_KEY,JSON.stringify(F)); bus.emit('fare:changed');
}
function fareReplace(n){ Object.keys(F).forEach(k=>delete F[k]); Object.assign(F,n||{}); }
const fareAll=()=>Object.values(F);
const fareSettings=()=>Object.assign({refDaily:0,tol:3},rj(FARE_SET,{}));
function fareSaveSettings(o){ ls.set(FARE_SET,JSON.stringify(Object.assign(fareSettings(),o))); }
/* ---- 다른 모듈(동기화)과 연결 ---- */
function fareLink(sync){
  if(!sync) return;
  sync.register({file:'fare.json',app:'fare',mk:FARE_META,get:()=>F,set:d=>{ fareReplace(d); ls.set(FARE_KEY,JSON.stringify(F)); bus.emit('fare:changed'); }});
}
/* ---- 파일 읽기: .xlsx → 시트별 칸, 그 밖 → 글자 ---- */
async function fareReadFile(f,xlsxRows){
  const buf=await f.arrayBuffer(), h=new Uint8Array(buf.slice(0,4));
  if(h[0]===0x50&&h[1]===0x4B) return fareParse(await xlsxRows(buf),f.name);
  if(h[0]===0xD0&&h[1]===0xCF) throw new Error('예전 엑셀(.xls)입니다. .xlsx로 저장하거나 내용을 복사해 붙여넣어 주세요');
  if(h[0]===0x25&&h[1]===0x50) throw new Error('PDF는 글자를 복사해 붙여넣거나 엑셀로 저장해 올려 주세요');
  if(h[0]===0xFF&&h[1]===0xD8||h[0]===0x89&&h[1]===0x50) throw new Error('사진은 글자를 읽을 수 없습니다. 휴대폰의 글자 인식(문자 복사)으로 복사해 붙여넣어 주세요');
  let txt=new TextDecoder('utf-8').decode(buf);
  if(txt.includes('�')) { try{ txt=new TextDecoder('euc-kr').decode(buf); }catch(e){} }
  const rows=txt.split('\n').map(l=>l.split(/\t|,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(c=>c.replace(/^"|"$/g,'').trim()).map(c=>{ const n=frNum(c); return n!=null&&/^[\d,.\-]+$/.test(c)?n:c; }));
  const t=frTable(rows); if(t){ t.forEach(s=>{ s.src=f.name; s.id=s.ym+'|'+(s.veh||''); }); return t; }
  return fareParse(txt,f.name);
}
M.fare={F,FARE_KEY,FARE_KIND_NAME,FARE_META,fareAll,fareAnalyze,fareCheck,fareDel,fareDerive,fareKind,fareLink,fareParse,farePut,fareReadFile,fareReplace,fareSaveSettings,fareSettings,frHas,frNum};
})(window.M=window.M||{});
