const $=s=>document.querySelector(s);
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let token='', current=null;
const names={baseline:'기본 초안',coerce:'범위·단위 오류',policy_gap:'정책 갭 +60bp'};
function message(text,error=false){$('#message').textContent=text;$('#message').className=error?'error':'';}
async function api(path,body){
 const options=body===undefined?{}:{method:'POST',headers:{'Content-Type':'application/json','X-Demo-Token':token},body:JSON.stringify(body)};
 const response=await fetch(path,options); const data=await response.json();
 if(!response.ok)throw Error(data.error||'요청을 처리하지 못했습니다.'); return data;
}
async function action(fn){
 const buttons=[...document.querySelectorAll('button')];buttons.forEach(b=>b.disabled=true);
 message('로컬 Python에서 처리하고 있습니다…');
 try{await fn();}catch(error){message(error.message,true);}finally{buttons.forEach(b=>b.disabled=false);}
}
function score(v,decimals=1){return typeof v==='number'?(v>0?'+':'')+v.toFixed(decimals):'—';}
function render(result){
 current=result;
 if(!result.state){message(result.run.error||'완료되지 않은 실행입니다.',true);return;}
 $('#empty').classList.add('hidden');$('#result').classList.remove('hidden');
 $('#download').classList.remove('hidden');$('#download').href='/download?id='+result.run.id;
 $('#news-count').textContent=result.news_count+'건';$('#dur').textContent=score(result.state.dur.total,2);
 $('#verdict').textContent=result.state.dur.verdict+' · 사용자 확정 전';
 $('#missing').textContent=result.state.rows.filter(r=>r.missing_basis).length+'행';
 $('#locked').textContent=result.state.edited_n+'칸';
 $('#news-list').innerHTML=result.input.items.map(n=>`<div class="news"><small>${esc(n.id)} <span class="tag">${esc(n.kind)}</span> ${esc(n.factor)} / ${esc(n.region)}</small><h3>${esc(n.title)}</h3><p>${esc(n.raw)}</p><small>분류 예시: ${esc(n.direction)} / 강도 ${score(n.intensity)}</small></div>`).join('');
 $('#rows').innerHTML=result.state.rows.map(r=>`<tr><td>${esc(r.label)}</td><td class="score">${score(r.dur)}</td><td class="score">${score(r.cur)}</td><td class="note">${r.missing_basis?'<span class="tag warn">근거 수치 없음</span>':r.locked.length?'<span class="tag good">사람 수정 보호</span>':'<span class="tag">검토 전 초안</span>'}\n${esc(r.note)}</td></tr>`).join('');
 $('#warnings').innerHTML=result.apply.warns.map(w=>'<li>'+esc(w)+'</li>').join('')||'<li>형식 보정 경고 없음. 사실 정확성 검증 결과를 뜻하지 않습니다.</li>';
 $('#run-meta').textContent=result.run.created+' / 최초 실행 '+result.run.elapsed+'초';
 $('#trace').innerHTML=result.trace.map(t=>`<div class="trace-entry"><span>${esc(t.time.slice(0,19))} UTC</span><strong>${esc(t.stage)}</strong><span>${esc(JSON.stringify(Object.fromEntries(Object.entries(t).filter(([k])=>!['time','stage'].includes(k)))))}</span></div>`).join('');
}
async function history(){
 const runs=await api('/api/history');
 $('#history-list').innerHTML=runs.length?runs.map(r=>`<div class="history-row"><span>${esc(r.created)}</span><strong>${esc(names[r.scenario]||r.scenario)}</strong><span class="status">${esc(r.status)}</span><span>${r.elapsed??'—'}초</span><button data-run="${esc(r.id)}">결과 보기</button></div>`).join(''):'<p>아직 실행 이력이 없습니다. 위에서 첫 시연을 실행하세요.</p>';
 document.querySelectorAll('[data-run]').forEach(b=>b.addEventListener('click',()=>action(async()=>{render(await api('/api/state?id='+b.dataset.run));message('저장된 실행 결과를 불러왔습니다.');})));
 return runs;
}
$('#run').addEventListener('click',()=>action(async()=>{render(await api('/api/run',{scenario:$('#scenario').value}));await history();message('초안을 생성했습니다. 정책·TP의 근거 누락과 사람 검토가 필요한 항목을 확인하세요.');}));
$('#review').addEventListener('click',()=>action(async()=>{
 if(!current?.state)throw Error('새 시연을 먼저 실행하세요.');
 const before=current.state.rows[0].dur;
 const value=Number($('#score').value);
 render(await api('/api/review',{id:current.run.id,score:value}));
 message(before===value?'같은 값을 저장했습니다. 초안과 다른 값으로 수정해야 변경 감지를 시연할 수 있습니다.':'검토 점수를 저장했습니다. 재실행 버튼으로 보존 여부를 확인하세요.');
}));
$('#reapply').addEventListener('click',()=>action(async()=>{
 if(!current?.state)throw Error('새 시연을 먼저 실행하세요.');
 render(await api('/api/reapply',{id:current.run.id}));
 message('재실행 완료. 현재 글로벌 펀더 점수 '+score(current.state.rows[0].dur)+' / 수정 보호 '+current.state.edited_n+'칸 / 뉴스 '+current.news_count+'건.');
}));
$('#refresh').addEventListener('click',()=>action(async()=>{await history();message('실행 이력을 갱신했습니다.');}));
document.querySelectorAll('[data-tab]').forEach(button=>button.addEventListener('click',()=>{
 document.querySelectorAll('.page').forEach(p=>p.classList.toggle('hidden',p.id!==button.dataset.tab));
 document.querySelectorAll('[data-tab]').forEach(b=>{b.classList.toggle('active',b===button);b.setAttribute('aria-selected',String(b===button));});
 message(''); window.scrollTo(0,0);
}));
function renderEvidence(data){
 $('#evidence-total').textContent=`검증 ${data.passed??0} / ${data.total??0} 통과`;
 $('#evidence-time').textContent=data.generated_at?`실행 ${data.generated_at} · 전체 ${data.seconds}초`:data.notice;
 $('#evidence-scope').textContent=data.scope||'';
 $('#checks').innerHTML=(data.cases||[]).map(c=>`<tr><td>${esc(c.name)}</td><td>${esc(JSON.stringify(c.expected))}</td><td>${esc(JSON.stringify(c.actual))}</td><td class="${c.passed?'pass':'fail'}">${c.passed?'PASS':'FAIL'}</td></tr>`).join('');
}
async function init(){
 try{
 token=(await api('/api/config')).token;
 const results=await Promise.all([api('/api/evidence'),history()]);
 renderEvidence(results[0]);
 const latest=results[1].find(r=>r.status==='ok');if(latest)render(await api('/api/state?id='+latest.id));
 }catch(error){message('초기화 실패: '+error.message+' · 로컬 서버 실행 상태를 확인하세요.',true);}
}
init();
