/* mod-ledger.js — 가계부 기록 · 분류 · 고정지출 · 카드 반영(importCore)
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.ledger로 내보냄.
   MODULE ledger V1.15 (2026-10-29) */
(function(M){
const {$,CARD_CATS,bus,ls,metaOf,pad,rj,toast}=M.base;
const {BIKE_MAX,LPG_RE,STATION_RE,classifyCard,isBigFuel,netCardTx,normM}=M.card;
const sync=M.sync;                                    // 동기화 모듈은 있으면 연결(없어도 동작)
/* ================= 데이터 =================
   expenses.data.v1 = { 't:<결제id>':{d:'YYYY-MM-DD', t:'HH:MM', m:가맹점, v:금액, g?:분류(직접 지정), n?:메모, e?:1(금액 직접 수정), man?:1(직접 추가)},
                        'r:<가게>':{g:분류} (같은 가게 분류 규칙) }
   지운 결제는 동기화 기록(del)에 남아 같은 엑셀을 다시 올려도 되살아나지 않음                                  */
const EK='expenses.data.v1', EM='expenses.sync.meta', WK='worklog.data.v2', WM='worklog.sync.meta';
let touchN=0;                                       // 기록이 바뀔 때마다 늘림 (계산 결과 다시 만들기용)
const E=rj(EK,{});                                   // 같은 객체를 계속 씀(내용만 바꿈) → 다른 파일이 가져가도 최신
function replaceE(n){ Object.keys(E).forEach(k=>delete E[k]); Object.assign(E,n||{}); touchN++; }
function touch(keys){
  const m=metaOf(EM), now=Date.now();
  keys.forEach(k=>{ if(E[k]){ m.mt[k]=now; delete m.del[k]; } else { m.del[k]=now; delete m.mt[k]; } });
  m.dirty=true; touchN++; ls.set(EM,JSON.stringify(m));
  if(!ls.set(EK,JSON.stringify(E))) toast('⚠ 저장 실패: 저장 공간을 확인하세요');
  if(sync) sync.schedule(1500);
}
/* ---- 분류 (같은 계열 묶기) ---- */
const MERGED='식비';                                                              // 외식 + 편의점 + 카페·간식 + 마트·생활을 하나로 합친 세부 분류
const SHOP_NAME='쇼핑·생필품';                                                      // 쇼핑 + 생필품(다이소·올리브영 등 생활용품점)을 하나로 합친 세부 분류
const FIXED_NAME='고정비(주거·공과금·통신비)';                                      // 주거·공과금 + 통신비를 하나로 합친 세부 분류
const GROUP_ALIAS={'쇼핑':SHOP_NAME,'외식':MERGED,'편의점·마트·카페':MERGED,'편의점':MERGED,'카페·간식':MERGED,'마트·생활':MERGED,'통신비':FIXED_NAME,'주거·공과금':FIXED_NAME};         // 예전에 저장된 분류 이름도 합친 이름으로 읽음
const normGroup=g=>GROUP_ALIAS[g]||g;
const GROUPS=[      // 위에서부터 먼저 맞는 분류 (구체적인 것 먼저)
  [FIXED_NAME,'🏠',/휴대폰알림|알림서비스|LGU|LG ?U\+|LG유플러스|유플러스|SKT|SK텔레콤|SK브로드밴드|\bKT\b|케이티|올레|헬로모바일|알뜰폰|세븐모바일|이야기모바일|통신요금|통신비|인터넷요금/i],
  ['보험','🛡️',/보험|화재해상|해상화재|손해보험|손보|생명|라이나|메리츠|현대해상|DB손|KB손|삼성화재|한화손|흥국|AIA|처브|캐롯|하나손/i],
  ['의료','💊',/병원|의원|약국|치과|한의원|안과|피부과|내과|외과|정형|소아과|이비인후|산부인과|비뇨|정신건강|재활|검진|요양|메디|클리닉/],
  ['주유·차량','⛽',/주유소|에너지|칼텍스|오일뱅크|S-?OIL|에쓰오일|충전소|요소수|유록스|세차|정비|타이어|카센터|카센타|공업사|자동차|오토|모터스|부품/i],
  ['통행료','🛣️',/도로공사|하이패스|통행료|고속도로|유료도로/],
  ['교통','🚌',/택시|카카오T|티머니|캐시비|버스|지하철|코레일|KTX|SRT|철도|주차|쏘카|그린카|따릉이|킥보드|항공/i],
  [FIXED_NAME,'🏠',/입주자대표|관리비|관리사무소|도시가스|가스공사|한국전력|한전|전기요금|수도|상하수도|월세|부동산/],
  ['공공·세금','🏛️',/시청|구청|군청|도청|주민센터|행정복지|세무서|국세|지방세|우체국|경찰|법원|등기소|과태료|범칙금|정부24|민원/],
  ['교육','📚',/학원|교육|학교|유치원|어린이집|교재|강의|인강|클래스101|교습/],
  [MERGED,'🍚',/GS25|지에스\(?GS\)?25|씨유|\(CU\)|^CU|세븐일레븐|코리아세븐|이마트24|미니스톱|편의점/],
  [MERGED,'🍚',/커피|스타벅스|카페|투썸|이디야|빽다방|메가엠지씨|메가커피|컴포즈|할리스|폴바셋|커피빈|탐앤탐스|더벤티|매머드|파리바게|뚜레쥬르|베이커리|제과|브레드|배스킨|베스킨|던킨|크리스피|공차|설빙|빙수|디저트|케이크|마카롱|도넛|와플|아이스크림|요거|떡집|빵/],
  [MERGED,'🍚',/아웃백|빕스|애슐리|하누돈|한우|한돈|정육식당|고깃|고기|삼겹|갈비|등심|BBQ|비비큐|교촌|BHC|굽네|네네|처갓집|호식이|치킨|맘스터치|버거|맥도날드|롯데리아|KFC|써브웨이|서브웨이|피자|도미노|본죽|죽이야기|김밥|분식|떡볶이|찜닭|두찜|마라|양꼬치|쌀국수|라멘|우동|소바|돈까스|돈가스|돈카츠|규동|덮밥|백반|한정식|정식|밥상|밥집|국시|면옥|냉면|칼국수|수제비|국수|국밥|해장|설렁탕|곰탕|감자탕|순대|보쌈|족발|곱창|막창|대창|닭갈비|닭발|닭한마리|닭|오리|장어|해물|아구|아귀|대게|꽃게|조개|낙지|쭈꾸미|주꾸미|횟집|수산|초밥|스시|오마카세|이자카야|포차|호프|술집|한식|중식|일식|양식|뷔페|식당|반점|짜장|짬뽕|가든|숯불|화로|찌개|탕|찹쌀/i],
  [MERGED,'🍚',/마트|하나로|농협|노브랜드|이마트|트레이더스|홈플러스|코스트코|GS더프레시|롯데슈퍼|식자재|정육점|청과|과일|야채|반찬/],
  [SHOP_NAME,'🛍️',/다이소|아성|올리브영|생활용품|세탁|철물|무인|쿠팡|11번가|G마켓|지마켓|옥션|SSG|무신사|유니클로|탑텐|스파오|ZARA|나이키|아디다스|백화점|아울렛|신세계|롯데쇼핑|의류|신발|안경|문구|교보|알라딘|예스24|네이버페이|카카오페이/i],
  ['여가·문화','🎬',/CGV|롯데시네마|메가박스|영화|넷플릭스|NETFLIX|유튜브|YOUTUBE|티빙|웨이브|디즈니|멜론|스포티파이|노래|PC방|피씨방|볼링|당구|헬스|골프|스크린|사우나|찜질|호텔|모텔|펜션|숙박|여행|야놀자|여기어때|에어비앤비|테마파크/i],
  ['대출','🏦',/대출|캐피탈|카드론|원리금|주택금융|HF|이자상환/],
  ['경조사','🎁',/축의|조의|부의|경조|기부|헌금|후원/],
  ['금융·수수료','💳',/수수료|연회비|이자/],
];
const FOOD_TAIL=/(당|옥|뜰|집|관|각|정|가든|돈|하우스|키친|식탁|밥|면)$/;   // 한글 가게 이름 끝말로 본 음식점 (예: 월미당, 정뜰, 하누돈)
const GICON=Object.fromEntries([...GROUPS.map(g=>[g[0],g[1]]),['기타','📦']]);
const GNAMES=[...new Set(GROUPS.map(g=>g[0]))].concat('기타');
const WORK_G=new Set(['주유·차량','통행료']);   // 업무비로 보는 분류 (교통·주차는 생활비)
/* 큰 분류 3가지: 생활비용 · 업무비용 · 기타비용 (세부 분류에서 자동으로 정하고, 결제·가게별로 직접 옮길 수 있음) */
const CATS=[['life','생활비용','🏠'],['work','업무비용','🚚'],['etc','기타비용','📦']];
const CATN=Object.fromEntries(CATS.map(c=>[c[0],c[1]]));
const CATI=Object.fromEntries(CATS.map(c=>[c[0],c[2]]));
const CAT_DEF={'주유·차량':'work','통행료':'work','기타':'etc','금융·수수료':'etc','경조사':'etc'};   // 나머지(주거·대출·식비…)는 생활비용
const FIXED_G=new Set([FIXED_NAME,'대출','보험']);                                      // 고정비로 보는 세부 분류
let _bf=null;
let _bfE=null;
let _bfN=-1;
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
  if(r.g) return normGroup(r.g);
  const rule=E['r:'+brandOf(r.m)]; if(rule&&rule.g) return normGroup(rule.g);
  for(const [n,,re] of GROUPS) if(re.test(r.m)) return n;
  const last=brandOf(r.m).split(/\s+/).pop();
  if(/^[가-힣]{2,8}$/.test(last)&&FOOD_TAIL.test(last)) return MERGED;
  const kk=E['k:'+brandOf(r.m)]; if(kk&&kk.g) return normGroup(kk.g);               // 카카오 장소 검색으로 찾은 업종
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
/* 고른 세부(통신비·주유 등)의 결제만 남김 — 비우면 전체 */
const pickGroup=(list,grp)=>grp?list.filter(r=>groupOf(r)===grp):list;
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
    const adds=byDay[k].filter(it=>!old.some(e=>e.id===it.id||(normM(e.m)===normM(it.m)&&e.v===it.v&&(!e.tm||!it.tm||e.tm===it.tm))));
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
function unapplyWork(gone){           // 취소된 결제를 업무일지의 카드 결제 목록(cd)에서도 빼고 주유·요소수·통행료 합계를 다시 계산
  if(!gone.length) return 0;
  const W=rj(WK,{}), wm=metaOf(WM), now=Date.now(); let days=0;
  gone.forEach(g=>{
    const rec=W[g.k]; if(!rec||!rec.cd) return;
    const i=rec.cd.findIndex(e=>e.id===g.id||(normM(e.m)===normM(g.m)&&e.v===g.v)); if(i<0) return;
    const cd=rec.cd.slice(); cd.splice(i,1);
    const o=Object.assign({},rec); delete o.cd; Object.keys(CARD_CATS).forEach(c=>{ delete o[c]; });
    Object.keys(CARD_CATS).forEach(c=>{ const sm=cd.filter(e=>e.c===c).reduce((x,e)=>x+e.v,0); if(sm) o[c]=sm; });
    if(cd.length) o.cd=cd;
    if(Object.keys(o).length) W[g.k]=o; else delete W[g.k];
    wm.mt[g.k]=now; delete wm.del[g.k]; days++;
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
function dropOriginals(keys){        // 올린 원본(f:) 목록만 지움 — 지출(t:)·업무일지 기록은 건드리지 않음 (지운 표시는 동기화로 다른 기기에도 전달)
  const ks=keys.filter(k=>k.startsWith('f:')&&E[k]);
  ks.forEach(k=>{ delete E[k]; });
  if(ks.length) touch(ks);
  return ks.length;
}
function importCore(tx,src,rows,miss){        // 엑셀·붙여넣기 → 가계부(전체 결제) + 업무일지(주유·요소수·통행료) + 원본 보관
  const n=netCardTx(tx), md=metaOf(EM), keys=[];
  saveOriginal(rows,src,keys);
  /* 중복 확인: 같은 날·같은 가게(띄어쓰기·(주) 무시)·같은 금액인 기존 결제를 '한 건씩' 짝지음
     → 시각이 같거나 한쪽에 시각이 없으면 같은 결제, 둘 다 시각이 있는데 다르면 다른 결제(같은 날 2번 결제도 빠지지 않음) */
  const ids=new Set(n.items.map(it=>'t:'+it.id)), pool={}, pk=r=>`${r.d}|${normM(r.m)}|${r.v}`;
  Object.keys(E).forEach(k=>{ const r=E[k]; if(k.startsWith('t:')&&r&&r.d&&!ids.has(k)) (pool[pk(r)]=pool[pk(r)]||[]).push(k); });
  const claim=it=>{ const L=pool[pk({d:it.k,m:it.m,v:it.v})]; if(!L||!L.length) return null;
    let i=L.findIndex(k=>E[k].t&&it.tm&&E[k].t===it.tm); if(i<0) i=L.findIndex(k=>!E[k].t||!it.tm); if(i<0) return null;
    return L.splice(i,1)[0]; };
  let add=0,dup=0,skip=0;
  n.items.forEach(it=>{
    const key='t:'+it.id, cur=E[key];
    if(md.del[key]){ skip++; return; }                                   // 직접 지운 결제는 되살리지 않음
    if(cur){ if(cur.v!==it.v&&!cur.e){ cur.v=it.v; keys.push(key); add++; } else dup++; return; }
    const same=claim(it);
    if(same){ dup++; if(it.tm&&!E[same].t){ E[same].t=it.tm; keys.push(same); } return; }      // 시각 없던 기록에 시각 채움
    E[key]={d:it.k,m:it.m,v:it.v}; if(it.tm) E[key].t=it.tm; keys.push(key); add++;
  });
  /* 원래 승인이 이번 파일에 없는 취소: 같은 승인번호 → 같은 가게·같은 금액의 이전 결제(가계부에 이미 있는 것) 순으로 찾아 상계 */
  const gone=[]; let lone=0;
  n.orphan.forEach(o=>{
    let key='t:'+o.id;
    if(!E[key]){ key=''; let best='';
      Object.keys(E).forEach(k=>{ const r=E[k]; if(k.startsWith('t:')&&r&&r.d&&r.d<=o.k&&r.v===o.v&&normM(r.m)===normM(o.m)&&(!best||r.d>E[best].d)) best=k; });
      key=best; }
    if(key&&E[key]){ gone.push({id:key.slice(2),k:E[key].d,m:E[key].m,v:E[key].v}); delete E[key]; keys.push(key); } else lone++;
  });
  const unDays=unapplyWork(gone);
  if(keys.length) touch(keys);
  const days=applyWork(n.items)+unDays;
  const ks=n.items.map(i=>i.k).sort();
  const res={src,read:n.items.length,add,dup,skip,cancel:n.canceled+gone.length,lone,miss:miss||0,days,from:ks[0]||'',to:ks[ks.length-1]||'',at:Date.now()};
  ls.set('expenses.lastImport',JSON.stringify(res));
  return res;
}
const importMsg=r=>`✅ 읽음 ${r.read}건 → 새 결제 ${r.add}건 · 업무일지 ${r.days}일${r.dup?` · 이미 있음 ${r.dup}건`:''}${r.skip?` · 지운 결제 ${r.skip}건`:''}${r.cancel?` · 취소 ${r.cancel}건 상계`:''}${r.lone?` · 원래 결제가 없는 취소 ${r.lone}건은 제외`:''}${r.miss?` · ⚠ 못 읽은 줄 ${r.miss}개`:''}`;
/* ---- 다른 모듈과의 연결 (동기화 모듈이 있으면 기록 종류를 등록) ---- */
function reloadE(){ const n=rj(EK,null); if(n){ replaceE(n); bus.emit('data:changed'); } }      // 다른 탭·뒤로가기로 돌아오면 저장된 기록을 다시 읽어 옛 기록으로 덮어쓰지 않게
window.addEventListener('storage',e=>{ if(e.key===EK||e.key===WK) reloadE(); });
window.addEventListener('pageshow',e=>{ if(e.persisted) reloadE(); });
if(sync){
  sync.register({file:'expenses.json',app:'expenses',mk:EM,get:()=>E,set:d=>{ replaceE(d); ls.set(EK,JSON.stringify(E)); }});
  sync.register({file:'worklog.json',app:'worklog',mk:WM,get:()=>rj(WK,{}),set:d=>{ ls.set(WK,JSON.stringify(d)); }});
  bus.on('sync:merged',()=>{ if(migrateWork()) sync.schedule(800); });                    // 합친 뒤 업무일지의 충전소·바이크 비용 정리
}
migrateWork();
M.ledger={CATN,CATS,CAT_DEF,E,EK,EM,FIXED_G,FIXED_NAME,GICON,GNAMES,MERGED,SHOP_NAME,WK,allTx,brandOf,catOf,dropOriginals,groupOf,importCore,importMsg,pickGroup,replaceE,sumV,touch,ymKey,migrateWork};
})(window.M=window.M||{});
