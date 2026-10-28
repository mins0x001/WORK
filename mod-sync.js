/* mod-sync.js — GitHub 비공개 Gist 동기화 엔진
   업무허브 · 업무일지 · 가계부가 함께 쓰는 독립 모듈. 공개 API만 M.sync로 내보냄.
   MODULE sync V1.14 (2026-10-28) */
(function(M){
const {$,bus,ls,metaOf}=M.base;
/* GitHub 비공개 Gist 동기화 엔진 — 기록 종류(문서)를 register()로 등록하면 한 번 읽고, 합치고, 바뀐 파일만 한 번에 올림.
   합치기 규칙: 항목별로 더 나중에 고친 쪽이 이김, 지운 항목은 양쪽에서 지워짐. 화면은 bus 'sync:state' · 'data:changed' 로만 알림 받음. */
const SY_TOK='muhansync_token', SY_GID='muhansync_gist', SY_DESC='muhan-sync (무한매수 기록 동기화 · 지우지 마세요)';
const docs=[];
let cfgVersion='';
let syBusy=false;
let syAgain=false;
let syErr='';
let syTimer=null;
const syTok=()=>ls.get(SY_TOK)||'';
const syGist=()=>ls.get(SY_GID)||'';
const isOn=()=>!!(syTok()&&syGist());
const device=()=>/Android|iPhone|iPad|Mobile/i.test(navigator.userAgent)?'모바일':'PC';
function register(d){ if(!docs.some(x=>x.file===d.file)) docs.push(d); }     // d:{file, app, mk(기록 키), get(), set(data)→알림문구?, norm?(data)}
function configure(o){ if(o.version) cfgVersion=o.version; }
async function gh(path,opt){
  opt=opt||{};
  const headers={Authorization:'Bearer '+syTok(),Accept:'application/vnd.github+json'};
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
async function readGistFile(f){
  if(!f) return {};
  let c=f.content;
  if(f.truncated&&f.raw_url){ const r=await fetch(f.raw_url,{cache:'no-store'}); c=await r.text(); }   // 1MB 넘는 파일
  try{ return JSON.parse(c||'{}')||{}; }catch(e){ return {}; }
}
function schedule(ms){ if(!isOn()) return; clearTimeout(syTimer); syTimer=setTimeout(()=>now(),ms); }
async function now(why){
  if(!isOn()){ bus.emit('sync:state'); return; }
  if(syBusy){ syAgain=true; return; }
  syBusy=true; bus.emit('sync:state','busy');
  let changed=false;
  try{
    const g=await gh('/gists/'+syGist()), files={};
    for(const D of docs){
      const R=await readGistFile(g.files&&g.files[D.file]); R.data=R.data||{}; if(D.norm) R.data=D.norm(R.data);
      const meta=metaOf(D.mk), L={data:D.get(),mt:meta.mt,del:meta.del}, M=mergeDocs(L,R); let note='';
      if(canon(M.data)!==canon(L.data)){ note=D.set(M.data)||'다른 기기 변경 반영'; changed=true; }
      if(canon(M.data)!==canon(R.data)||canon(M.mt)!==canon(R.mt||{})||canon(M.del)!==canon(R.del||{})){
        files[D.file]={content:JSON.stringify({schema:'muhan-sync/1',app:D.app,version:cfgVersion,updated:new Date().toISOString(),device:device(),data:M.data,mt:M.mt,del:M.del})};
        if(!note) note='이 기기 변경을 올렸습니다'; }
      Object.assign(meta,{mt:M.mt,del:M.del,last:Date.now(),dirty:false,note:note||meta.note||''}); ls.set(D.mk,JSON.stringify(meta));
    }
    if(Object.keys(files).length) await gh('/gists/'+syGist(),{method:'PATCH',body:JSON.stringify({files})});
    syErr='';
    if(changed){ bus.emit('sync:merged',why); bus.emit('data:changed',why); }
  }catch(e){ syErr=e.message; }
  finally{ syBusy=false; bus.emit('sync:state'); if(syAgain){ syAgain=false; schedule(300); } }
}
async function connect(tok){      // 토큰으로 동기화 Gist를 찾거나 새로 만들고 첫 동기화
  ls.set(SY_TOK,tok);
  try{
    let id='';
    for(let page=1; page<=3&&!id; page++){
      const list=await gh(`/gists?per_page=100&page=${page}`);
      const hit=list.find(x=>x.description===SY_DESC); if(hit) id=hit.id;
      if(list.length<100) break;
    }
    if(!id){
      const f={}; docs.slice(0,1).forEach(D=>{ f[D.file]={content:JSON.stringify({schema:'muhan-sync/1',data:{},mt:{},del:{}})}; });
      id=(await gh('/gists',{method:'POST',body:JSON.stringify({description:SY_DESC,public:false,files:f})})).id;
    }
    ls.set(SY_GID,id); syErr='';
    await now('connect');
  }catch(e){ ls.del(SY_TOK); ls.del(SY_GID); throw e; }
}
function disconnect(){ ls.del(SY_TOK); ls.del(SY_GID); syErr=''; bus.emit('sync:state'); }
const err=()=>syErr;
M.sync={register,configure,schedule,now,connect,disconnect,isOn,device,err,mergeDocs};
})(window.M=window.M||{});
