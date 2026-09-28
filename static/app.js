const $ = selector => document.querySelector(selector);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let token = '', current = null, busy = false, filterKind = 'all';
const names = {baseline:'기본 초안', coerce:'범위·단위 오류', policy_gap:'정책 갭 +60bp'};
const descriptions = {
  baseline:'수치 근거가 없는 정책 점수가 어떻게 처리되는지 확인합니다.',
  coerce:'0.7과 -3으로 입력된 점수가 허용된 단위와 범위로 보정됩니다.',
  policy_gap:'시장 5.1%와 개인 경로 4.5%의 차이인 +60bp를 적용합니다.'
};
function message(text, kind = '') {
  $('#message').textContent = text;
  $('#message').className = kind;
  $('#message').setAttribute('role', kind === 'error' ? 'alert' : 'status');
}
function controls() {
  $('#run').disabled = busy || !token;
  $('#review').disabled = $('#reapply').disabled = busy || !token || !current;
  $('#refresh').disabled = busy || !token;
  $('#scenario').disabled = busy;
  document.querySelectorAll('[data-run]').forEach(button => button.disabled = busy);
  $('#run').textContent = busy ? '처리 중…' : '새 시연 실행';
  $('#demo').setAttribute('aria-busy', String(busy));
}
async function api(path, body) {
  const options = body === undefined ? {} : {method:'POST', headers:{'Content-Type':'application/json', 'X-Demo-Token':token}, body:JSON.stringify(body)};
  const response = await fetch(path, options);
  const data = await response.json();
  if (!response.ok) throw Error(data.error || '요청을 처리하지 못했습니다.');
  return data;
}
async function action(fn) {
  if (busy) return;
  busy = true; controls(); message('검증하고 기록하는 중입니다. 잠시 기다려 주세요.', 'busy');
  try { await fn(); }
  catch (error) { message(error.message + ' 다시 시도하거나 로컬 서버의 실행 상태를 확인하세요.', 'error'); }
  finally { busy = false; controls(); }
}
function score(value, decimals = 1) {
  return typeof value === 'number' && Number.isFinite(value) ? (value > 0 ? '+' : '') + value.toFixed(decimals) : '—';
}
function scoreClass(value) { return value > 0 ? 'positive' : value < 0 ? 'negative' : ''; }
function renderNews() {
  if (!current) return;
  const query = $('#news-search').value.trim().toLocaleLowerCase();
  const all = current.input.items;
  const visible = all.filter(n => (filterKind === 'all' || n.kind === filterKind) &&
    [n.title,n.raw,n.content,n.keyword,n.factor,n.region,n.source].join(' ').toLocaleLowerCase().includes(query));
  $('#filter-count').textContent = `${visible.length} / ${all.length}건 표시 · 필터는 점수 집계에 영향을 주지 않습니다.`;
  $('#news-list').innerHTML = visible.map(n => `<article class="news"><div class="news-heading"><span class="tag">${esc(n.kind)}</span><span>${esc(n.factor)} / ${esc(n.region)}</span><span>${esc(n.id)}</span></div><h3>${esc(n.title)}</h3><p>${esc(n.raw)}</p><small>${esc(n.source)} · ${esc(n.date)}</small><span class="direction">분류 예시: ${esc(n.direction)} / 강도 ${score(n.intensity)}</span></article>`).join('') || '<p class="no-results">일치하는 원문이 없습니다.<br>검색어를 지우거나 다른 유형을 선택하세요.</p>';
}
function render(result) {
  if (!result.state) throw Error('선택한 실행이 완료되지 않아 결과를 불러올 수 없습니다. 새 시연을 실행하세요.');
  current = result;
  $('#empty').classList.add('hidden'); $('#result').classList.remove('hidden');
  $('#download').classList.remove('hidden'); $('#download').href = '/download?id=' + encodeURIComponent(result.run.id);
  $('#active-run').textContent = '현재 결과: ' + (names[result.run.scenario] || result.run.scenario);
  $('#sample-date').textContent = '가상 자료 기준일 ' + result.input.date;
  $('#news-count').textContent = result.news_count + '건';
  $('#dur').textContent = score(result.state.dur.total, 2);
  $('#verdict').textContent = result.state.dur.verdict + ' · 확정 전';
  $('#missing').textContent = result.state.rows.filter(row => row.missing_basis).length + '행';
  $('#locked').textContent = result.state.edited_n + '칸';
  renderNews();
  $('#rows').innerHTML = result.state.rows.map(row => `<tr class="${row.row === 8 ? 'editable-row' : ''}"><td class="label">${esc(row.label)}</td><td class="score ${scoreClass(row.dur)}">${score(row.dur)}</td><td class="score ${scoreClass(row.cur)}">${score(row.cur)}</td><td class="note">${row.missing_basis ? '<span class="tag warn">근거 수치 없음</span>' : row.locked.length ? '<span class="tag good">사람 수정 보호</span>' : '<span class="tag">검토 전 초안</span>'}<p>${esc(row.note)}</p></td></tr>`).join('');
  const first = result.state.rows.find(row => row.row === 8);
  const stages = result.trace.map(entry => entry.stage);
  const lastReview = stages.lastIndexOf('human_review');
  const pending = lastReview > stages.lastIndexOf('draft_apply');
  $('#original-score').textContent = score(first.proposed_dur);
  $('#current-score').textContent = score(first.dur);
  $('#review-status').textContent = pending ? '저장됨 · 재실행 대기' : first.locked.includes('dur') ? '수정값 보존 확인' : '자동 초안';
  $('#review-status').className = 'badge ' + (pending ? 'warn' : first.locked.includes('dur') ? 'good' : 'subtle');
  const steps = ['#step-source','#step-review','#step-preserve'];
  steps.forEach((selector, index) => {
    const done = index === 0 || (index === 1 && lastReview >= 0) || (index === 2 && !pending && first.locked.includes('dur'));
    $(selector).classList.toggle('done', done);
  });
  $('#warnings').innerHTML = result.apply.warns.map(warning => '<li>' + esc(warning) + '</li>').join('') || '<li>형식 보정 경고가 없습니다. 사실 정확성을 보장하는 결과는 아닙니다.</li>';
  $('#warning-count').textContent = result.apply.warns.length + '건';
  $('#run-meta').textContent = result.run.created + ' / 최초 실행 ' + result.run.elapsed + '초';
  $('#trace').innerHTML = result.trace.map(entry => `<div class="trace-entry"><span>${esc(entry.time.slice(0,19))} UTC</span><strong>${esc(entry.stage)}</strong><span>${esc(JSON.stringify(Object.fromEntries(Object.entries(entry).filter(([key]) => !['time','stage'].includes(key)))))}</span></div>`).join('');
  document.querySelectorAll('[data-history]').forEach(row => row.classList.toggle('is-current', row.dataset.history === result.run.id));
  controls();
}
async function history() {
  const runs = await api('/api/history');
  const statuses = {ok:'완료', failed:'실패', running:'진행 중'};
  $('#history-list').innerHTML = runs.length ? runs.map(run => `<div class="history-row ${current?.run.id === run.id ? 'is-current' : ''}" data-history="${esc(run.id)}"><span>${esc(run.created)}</span><strong>${esc(names[run.scenario] || run.scenario)}</strong><span class="status ${run.status === 'ok' ? 'ok' : run.status === 'failed' ? 'failed' : ''}">${esc(statuses[run.status] || run.status)}</span><span>${run.elapsed == null ? '—' : esc(run.elapsed) + '초'}</span><button data-run="${esc(run.id)}" aria-label="${esc(run.created)} ${esc(names[run.scenario] || run.scenario)} 결과 보기">결과 보기</button></div>`).join('') : '<p class="history-empty">아직 실행 이력이 없습니다. 새 시연을 실행하면 이곳에 기록됩니다.</p>';
  document.querySelectorAll('[data-run]').forEach(button => button.addEventListener('click', () => action(async () => {
    render(await api('/api/state?id=' + encodeURIComponent(button.dataset.run)));
    message('저장된 결과를 불러왔습니다. 현재 결과의 시나리오와 자료 기준일을 확인하세요.');
    $('#result').scrollIntoView({block:'start'});
  })));
  controls();
  return runs;
}
$('#scenario').addEventListener('change', () => $('#scenario-help').textContent = descriptions[$('#scenario').value]);
$('#news-search').addEventListener('input', renderNews);
document.querySelectorAll('[data-kind]').forEach(button => button.addEventListener('click', () => {
  filterKind = button.dataset.kind;
  document.querySelectorAll('[data-kind]').forEach(item => item.setAttribute('aria-pressed', String(item === button)));
  renderNews();
}));
$('#run').addEventListener('click', () => action(async () => {
  render(await api('/api/run', {scenario:$('#scenario').value}));
  await history();
  message('초안을 만들었습니다. 원문과 근거 누락 항목을 확인한 뒤 검토 점수를 바꿔보세요.');
}));
$('#review').addEventListener('click', () => action(async () => {
  if (!current) throw Error('새 시연을 먼저 실행하세요.');
  const before = current.state.rows.find(row => row.row === 8).dur;
  const value = Number($('#score').value);
  render(await api('/api/review', {id:current.run.id, score:value}));
  message(before === value ? '같은 값을 저장했습니다. 변경 감지를 확인하려면 자동 초안과 다른 점수를 저장하세요.' : '검토 점수를 저장했습니다. 재실행·보존 확인을 눌러 수정값이 남는지 확인하세요.');
}));
$('#reapply').addEventListener('click', () => action(async () => {
  if (!current) throw Error('새 시연을 먼저 실행하세요.');
  render(await api('/api/reapply', {id:current.run.id}));
  message(`재실행 완료. 현재 점수 ${score(current.state.rows.find(row => row.row === 8).dur)} / 수정 보호 ${current.state.edited_n}칸 / 뉴스 ${current.news_count}건.`);
}));
$('#refresh').addEventListener('click', () => action(async () => { await history(); message('실행 이력을 갱신했습니다.'); }));
const tabs = [...document.querySelectorAll('[data-tab]')];
function selectTab(button, focus = false) {
  document.querySelectorAll('.page').forEach(page => page.classList.toggle('hidden', page.id !== button.dataset.tab));
  tabs.forEach(tab => { const selected = tab === button; tab.classList.toggle('active', selected); tab.setAttribute('aria-selected', String(selected)); tab.tabIndex = selected ? 0 : -1; });
  if (focus) button.focus();
}
tabs.forEach((button, index) => {
  button.addEventListener('click', () => selectTab(button));
  button.addEventListener('keydown', event => {
    const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length : event.key === 'ArrowLeft' ? (index + tabs.length - 1) % tabs.length : event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : null;
    if (next !== null) { event.preventDefault(); selectTab(tabs[next], true); }
  });
});
function renderEvidence(data) {
  $('#evidence-total').textContent = data.total == null ? '검증 기록이 없습니다' : `검증 ${data.passed} / ${data.total} 통과`;
  $('#evidence-time').textContent = data.generated_at ? `실행 ${data.generated_at} · 전체 ${data.seconds}초` : data.notice || '';
  $('#evidence-scope').textContent = data.scope || '';
  $('#checks').innerHTML = (data.cases || []).map(item => `<tr><td>${esc(item.name)}</td><td>${esc(JSON.stringify(item.expected))}</td><td>${esc(JSON.stringify(item.actual))}</td><td class="${item.passed ? 'pass' : 'fail'}">${item.passed ? '통과' : '실패'}</td></tr>`).join('');
}
async function init() {
  try { token = (await api('/api/config')).token; }
  catch { message('로컬 서버에 연결하지 못했습니다. 서버 실행 후 페이지를 새로고침하세요.', 'error'); return; }
  controls();
  const results = await Promise.allSettled([api('/api/evidence'), history()]);
  if (results[0].status === 'fulfilled') renderEvidence(results[0].value);
  else { $('#evidence-total').textContent = '검증 기록을 불러오지 못했습니다'; $('#evidence-scope').textContent = '서버의 evidence 파일을 확인한 후 새로고침하세요.'; }
  if (results[1].status === 'fulfilled') {
    const latest = results[1].value.find(run => run.status === 'ok');
    if (latest) {
      try { render(await api('/api/state?id=' + encodeURIComponent(latest.id))); }
      catch { message('이전 결과 파일을 불러오지 못했습니다. 새 시연을 실행해 주세요.', 'error'); }
    }
  } else { $('#history-list').innerHTML = '<p class="history-empty">이력을 불러오지 못했습니다. 이력 새로고침을 눌러 다시 시도하세요.</p>'; }
  controls();
}
init();
