/* mod-export.js — 업무일지 엑셀·CSV 내용 만들기
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.export로 내보냄.
   MODULE export V1.16 (2026-10-29) */
(function(M){
const {$,CARD_CATS,WD,XL,keyOf,pad,parseKey}=M.base;
const {holNames}=M.holiday;
const {BYNAME,DISTANCE_EDGES,ROWS,TIER_LABELS,routeInfo}=M.dealer;
const {SHIFT,dayRouteKmIn,periodOf,rangeOf,roundTextIn,roundsIn,statsFor}=M.stats;
/* ================= 업무일지 내보내기 (엑셀 · CSV 행 만들기) =================
   data를 매개변수로 받아 파일 내용만 만듭니다. 파일 전달(저장·공유)은 화면 담당.   */
function buildWorkbookFor(data){
  const keys=Object.keys(data).sort();
  const Hd=s=>({v:s,s:1});
  // ① 월별요약: 기록이 있는 근무 기간만
  const seen=new Set(), periods=[];
  keys.forEach(k=>{ const p=periodOf(parseKey(k)), id=p.y*12+p.m; if(!seen.has(id)){ seen.add(id); periods.push(p); } });
  periods.sort((a,b)=>(a.y*12+a.m)-(b.y*12+b.m));
  const sum=[[Hd('근무기간'),Hd('시작일'),Hd('종료일'),Hd('근무일'),Hd('1회전'),Hd('2회전'),Hd('3회전'),Hd('휴무'),Hd('운행 km'),Hd('주유(원)'),Hd('통행료(원)'),Hd('요소수(원)'),Hd('비용 합계(원)'),Hd('저장 회전'),Hd('방문 대리점'),Hd('회전 경로 km')]];
  const tot=periods.map(p=>statsFor(data,p));
  periods.forEach((p,i)=>{
    const c=tot[i], r=rangeOf(p);
    sum.push([`${p.y}-${pad(p.m+1)}`,{d:keyOf(r.start)},{d:keyOf(r.end)},{v:c.work,s:4},{v:c['1'],s:4},{v:c['2'],s:4},{v:c['3'],s:4},{v:c.off,s:4},
      c.odo.state==='ok'?{v:c.odo.dist,s:4}:'',{v:c.fuel,s:4},{v:c.toll,s:4},{v:c.urea,s:4},{v:c.cost,s:4},{v:c.trips,s:4},{v:c.stops,s:4},{v:c.rkm,s:4}]);
  });
  if(periods.length){
    const n=periods.length+1, S=(col,val)=>({f:`SUM(${col}2:${col}${n})`,v:val,s:4});
    const t=f=>tot.reduce((x,c)=>x+f(c),0);
    sum.push([{v:'합계',s:7},'','',S('D',t(c=>c.work)),S('E',t(c=>c['1'])),S('F',t(c=>c['2'])),S('G',t(c=>c['3'])),S('H',t(c=>c.off)),
      S('I',t(c=>c.odo.state==='ok'?c.odo.dist:0)),S('J',t(c=>c.fuel)),S('K',t(c=>c.toll)),S('L',t(c=>c.urea)),S('M',t(c=>c.cost)),
      S('N',t(c=>c.trips)),S('O',t(c=>c.stops)),S('P',t(c=>c.rkm))]);
  }
  // ② 일별기록
  const day=[[Hd('날짜'),Hd('요일'),Hd('근무기간'),Hd('공휴일'),Hd('근무'),Hd('계기판(km)'),Hd('주유(원)'),Hd('통행료(원)'),Hd('요소수(원)'),Hd('1회전'),Hd('2회전'),Hd('3회전'),Hd('회전 경로 km')]];
  keys.forEach(k=>{
    const d=parseKey(k), r=data[k], p=periodOf(d);
    day.push([{d:k},WD[d.getDay()],`${p.y}-${pad(p.m+1)}`,holNames(k,(data[k]||{}).hol).join('/'),r.s?SHIFT[r.s].label:'',
      r.odo?{v:r.odo,s:4}:'',r.fuel?{v:r.fuel,s:4}:'',r.toll?{v:r.toll,s:4}:'',r.urea?{v:r.urea,s:4}:'',
      roundTextIn(data,k,'1'),roundTextIn(data,k,'2'),roundTextIn(data,k,'3'),r.r?{v:dayRouteKmIn(data,k),s:4}:'']);
  });
  // ③ 회전기록: 방문 대리점 한 곳당 한 줄
  const trip=[[Hd('날짜'),Hd('요일'),Hd('근무기간'),Hd('회전'),Hd('순서'),Hd('대리점'),Hd('코드'),Hd('전화'),Hd('주소'),Hd('회전 경로 km'),Hd('확인 필요 구간')]];
  keys.forEach(k=>roundsIn(data,k).forEach(x=>{
    const d=parseKey(k), p=periodOf(d), info=routeInfo(x.names);
    x.st.forEach((s,i)=>{
      const row=BYNAME.get(s.n);
      trip.push([{d:k},WD[d.getDay()],`${p.y}-${pad(p.m+1)}`,`${x.n}회전`,{v:i+1,s:4},s.n,row?row.codes.join(', '):(s.c||''),row?row.phone:'',row?(row.address||''):'',
        i===0?{v:info.km,s:4}:'',i===0&&info.unk?{v:info.unk,s:4}:'']);
    });
  }));
  // ④ 대리점목록 · ⑤ 거리표
  const list=[[Hd('권역'),Hd('대리점명'),Hd('코드1'),Hd('코드2'),Hd('전화'),Hd('주소'),Hd('포천 기준 km')]];
  ROWS.forEach(r=>list.push([TIER_LABELS[r.tier].replace(/^\S+\s/,''),r.name,r.codes[0]||'',r.codes[1]||'',r.phone,r.address||'',r.dist?{v:r.dist}:'']));
  const edge=[[Hd('출발'),Hd('도착'),Hd('km'),Hd('비고')]];
  DISTANCE_EDGES.forEach(e=>edge.push([e.from,e.to,{v:e.km,s:4},e.note||'']));
  // ⑥ 카드내역: 일지에 반영한 카드 결제
  const card=[[Hd('날짜'),Hd('요일'),Hd('시각'),Hd('가맹점'),Hd('분류'),Hd('금액(원)')]];
  keys.forEach(k=>((data[k]||{}).cd||[]).forEach(e=>{ const d=parseKey(k); card.push([{d:k},WD[d.getDay()],e.tm||'',e.m,CARD_CATS[e.c].label,{v:e.v,s:4}]); }));
  return XL.build([
    {name:'월별요약',cols:[10,12,12,8,7,7,7,7,10,12,12,12,14,10,12,13],freeze:true,rows:sum},
    {name:'일별기록',cols:[12,6,10,16,8,12,11,11,11,30,30,30,13],freeze:true,rows:day},
    {name:'회전기록',cols:[12,6,10,8,6,14,14,16,36,13,14],freeze:true,rows:trip},
    {name:'대리점목록',cols:[16,14,9,9,16,36,13],freeze:true,rows:list},
    {name:'거리표',cols:[12,12,8,12],freeze:true,rows:edge},
    {name:'카드내역',cols:[12,6,8,30,9,12],freeze:true,rows:card}
  ]);
}
function csvRowsFor(data){
  const keys=Object.keys(data).sort();
  const rows=[['날짜','요일','근무기간','공휴일','근무','계기판(km)','주유(원)','통행료(원)','요소수(원)','1회전','2회전','3회전','회전 경로(km)']];
  keys.forEach(k=>{
    const d=parseKey(k), r=data[k], p=periodOf(d);
    rows.push([k,WD[d.getDay()],`${p.y}-${pad(p.m+1)}`,holNames(k,(data[k]||{}).hol).join('/'),r.s?SHIFT[r.s].label:'',r.odo||'',r.fuel||'',r.toll||'',r.urea||'',
      roundTextIn(data,k,'1'),roundTextIn(data,k,'2'),roundTextIn(data,k,'3'),r.r?dayRouteKmIn(data,k):'']);
  });
  return rows;
}
M.export={buildWorkbookFor,csvRowsFor};
})(window.M=window.M||{});
