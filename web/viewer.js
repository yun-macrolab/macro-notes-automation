// 내 기록 보기 — 기록 워크북을 고르고, 저장될 때마다 다시 읽어 화면에 그린다.
// 파일 내용은 브라우저 작업 스레드(web/worker.mjs)로만 넘기고 서버나 저장소로 보내지 않는다.
// Chrome·Edge는 파일 핸들을 기억해 3초마다 수정 시각을 확인하고, 다른 브라우저는 파일을 다시 고를 때 읽는다.
"use strict";
const $ = selector => document.querySelector(selector);
const POLL_MS = 3000;
const DB = { name: "macro-notes-viewer", store: "handles", key: "last" };
const canWatch = typeof window.showOpenFilePicker === "function";

let worker = null;
let engineReady = false;
let engineFailed = false;
let seq = 0;
const pending = new Map();
let handle = null;        // 지금 보고 있는 파일 핸들(Chrome·Edge)
let remembered = null;    // 지난 방문에 연 파일 핸들
let lastFile = null;      // 마지막으로 읽은 File(이름·수정 시각)
let seenStamp = null;     // 마지막으로 읽기에 성공한 파일의 수정 시각·크기
let failedStamp = null;   // 읽기에 실패한 파일의 수정 시각·크기(같은 파일을 되풀이해 읽지 않게)
let data = null;
let reading = false;
let timer = null;
let filterKind = "all";

function message(text, kind = "") {
  const box = $("#message");
  box.textContent = text;
  box.className = kind;
  box.setAttribute("role", kind === "error" ? "alert" : "status");
}

function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "class") node.className = value;
    else if (key === "text") node.textContent = value;
    else node.setAttribute(key, value);
  }
  children.forEach(child => { if (child !== null && child !== undefined && child !== false) node.append(child); });
  return node;
}

const score = (value, decimals = 1) =>
  typeof value === "number" && Number.isFinite(value) ? (value > 0 ? "+" : "") + value.toFixed(decimals) : "—";
const scoreClass = value => (value > 0 ? "positive" : value < 0 ? "negative" : "");
const clock = ms => new Date(ms).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "medium" });
const stamp = file => `${file.lastModified}:${file.size}`;
function addDays(iso, days) {
  const day = new Date(iso + "T00:00:00Z");
  day.setUTCDate(day.getUTCDate() + days);
  return day.toISOString().slice(0, 10);
}
function mondayOf(iso) {
  const day = new Date(iso + "T00:00:00Z");
  if (Number.isNaN(day.getTime())) return null;
  return addDays(iso, -((day.getUTCDay() + 6) % 7));
}

/* ---------- 작업 스레드 ---------- */
function engineFailure() {
  engineFailed = true;
  message("기록을 읽는 엔진을 시작하지 못했습니다. 최신 Chrome·Edge·Safari·Firefox에서 다시 열어 주세요.", "error");
  pending.forEach(({ resolve }) => resolve({ ok: false, text: JSON.stringify({ error: "기록을 읽는 엔진을 시작하지 못했습니다." }) }));
  pending.clear();
}
try {
  worker = new Worker(new URL("web/worker.mjs", location.href), { type: "module" });
  worker.onerror = engineFailure;
  worker.onmessage = ({ data: reply }) => {
    if (reply.type === "ready") {
      engineReady = true;
      if (!data && !reading) message(canWatch
        ? "준비됐습니다. ‘기록 파일 열기’를 누르세요. 엑셀에서 저장할 때마다 자동으로 다시 읽습니다."
        : "준비됐습니다. ‘기록 파일 열기’를 누르세요. 이 브라우저에서는 저장한 뒤 파일을 다시 골라야 바뀐 내용이 보입니다.");
      return;
    }
    if (reply.type === "failed") {
      console.error(reply.error);
      engineFailure();
      return;
    }
    const entry = pending.get(reply.id);
    if (entry) {
      pending.delete(reply.id);
      entry.resolve(reply);
    }
  };
} catch (error) {
  console.error(error);
  engineFailure();
}

function parse(bytes, name) {
  if (engineFailed || !worker) {
    return Promise.resolve({ ok: false, text: JSON.stringify({ error: "기록을 읽는 엔진을 시작하지 못했습니다." }) });
  }
  return new Promise(resolve => {
    const id = ++seq;
    pending.set(id, { resolve });
    worker.postMessage({ id, bytes, name }, [bytes]);
  });
}

/* ---------- 파일 읽기 ---------- */
async function readFile(file, { quiet = false } = {}) {
  if (reading) return false;
  reading = true;
  buttons();
  if (!quiet) {
    message(engineReady ? "기록 파일을 읽고 있습니다." : "기록을 읽는 엔진을 준비하고 있습니다. 처음에는 약 14MB를 내려받아 10초 안팎 걸립니다.", "busy");
  }
  try {
    const reply = await parse(await file.arrayBuffer(), file.name);
    const body = JSON.parse(reply.text);
    if (!reply.ok) {
      failedStamp = stamp(file);
      message(body.error, "error");
      return false;
    }
    data = body;
    lastFile = file;
    seenStamp = stamp(file);
    failedStamp = null;
    render();
    message(quiet ? `저장된 변경을 반영했습니다(${clock(Date.now())}).` : "기록을 불러왔습니다.");
    return true;
  } catch (error) {
    console.error(error);
    failedStamp = stamp(file);
    message("파일을 읽지 못했습니다. 다시 열어 주세요.", "error");
    return false;
  } finally {
    reading = false;
    buttons();
  }
}

async function useHandle(next) {
  stopWatching();
  handle = next;
  remembered = next;
  seenStamp = failedStamp = null;
  try {
    await idb("readwrite", store => store.put(next, DB.key));
  } catch {
    // 핸들을 기억하지 못해도 지금 보는 데는 지장이 없다.
  }
  try {
    await readFile(await next.getFile());
  } catch (error) {
    console.error(error);
    message("파일을 열지 못했습니다. ‘기록 파일 열기’로 다시 골라 주세요.", "error");
    handle = null;
  }
  if (handle) timer = setInterval(check, POLL_MS);
  buttons();
}

async function check() {
  if (!handle || reading || document.hidden) return;
  let file;
  try {
    file = await handle.getFile();
  } catch (error) {
    stopWatching();
    handle = null;
    message(error.name === "NotFoundError"
      ? "파일을 찾을 수 없습니다. 옮기거나 이름을 바꿨다면 다시 열어 주세요."
      : "파일을 읽을 권한이 없어졌습니다. ‘지난 파일 다시 열기’를 누르세요.", "error");
    buttons();
    return;
  }
  const current = stamp(file);
  if (current !== seenStamp && current !== failedStamp) await readFile(file, { quiet: true });
}

function stopWatching() {
  clearInterval(timer);
  timer = null;
}

async function openFile() {
  if (!canWatch) {
    $("#picker").click();
    return;
  }
  let picked;
  try {
    [picked] = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: "엑셀 기록 파일", accept: {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
        "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"],
      } }],
    });
  } catch (error) {
    if (error.name !== "AbortError") message("파일을 열지 못했습니다.", "error");
    return;
  }
  await useHandle(picked);
}

async function reopen() {
  if (!remembered) return;
  try {
    if ((await remembered.requestPermission({ mode: "read" })) !== "granted") {
      message("파일을 읽을 권한을 받지 못했습니다. 다시 눌러 ‘허용’을 고르거나 파일을 새로 여세요.", "error");
      return;
    }
  } catch {
    message("지난 파일을 다시 열지 못했습니다. ‘기록 파일 열기’로 다시 골라 주세요.", "error");
    return;
  }
  await useHandle(remembered);
}

function closeFile() {
  stopWatching();
  handle = remembered = lastFile = data = null;
  seenStamp = failedStamp = null;
  idb("readwrite", store => store.delete(DB.key)).catch(() => {});
  $("#result").classList.add("hidden");
  $("#empty").classList.remove("hidden");
  message("파일을 닫았습니다. 화면에 띄웠던 내용과 기억해 둔 파일 위치도 지웠습니다.");
  buttons();
}

/* ---------- 지난 파일 기억(IndexedDB에는 파일 위치만, 내용은 저장하지 않는다) ---------- */
function idb(mode, action) {
  return new Promise((resolve, reject) => {
    const open = indexedDB.open(DB.name, 1);
    open.onupgradeneeded = () => open.result.createObjectStore(DB.store);
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = db.transaction(DB.store, mode);
      const request = action(tx.objectStore(DB.store));
      tx.oncomplete = () => { db.close(); resolve(request.result); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    };
  });
}

async function restore() {
  if (!canWatch) return;
  try {
    remembered = (await idb("readonly", store => store.get(DB.key))) || null;
  } catch {
    remembered = null;
  }
  buttons();
  if (!remembered) return;
  try {
    if ((await remembered.queryPermission({ mode: "read" })) === "granted") await useHandle(remembered);
  } catch {
    // 권한을 물어보려면 사용자의 클릭이 필요하다. ‘지난 파일 다시 열기’ 버튼으로 받는다.
  }
}

/* ---------- 화면 ---------- */
function buttons() {
  $("#open").disabled = reading;
  $("#reload").disabled = reading;
  $("#reload").classList.toggle("hidden", !handle);
  $("#close").classList.toggle("hidden", !data && !remembered);
  $("#reopen").classList.toggle("hidden", !remembered || !!handle);
  if (remembered) $("#reopen").textContent = "지난 파일 다시 열기 · " + remembered.name;
  $("#file-name").textContent = lastFile ? lastFile.name : "아직 열지 않았습니다";
  const meta = $("#file-meta");
  meta.replaceChildren();
  if (lastFile) {
    if (handle) meta.append(el("span", { class: "live", text: "실시간 연결" }));
    meta.append(`마지막 저장 ${clock(lastFile.lastModified)}`);
    if (!handle) meta.append(" · 저장한 뒤 다시 열면 바뀐 내용이 보입니다");
  }
}

function sourceNode(text) {
  const match = /https?:\/\/[^\s<>"']+/.exec(text || "");
  const label = (text || "").replace(match ? match[0] : "", "").trim();
  const box = el("small", { text: label || (match ? "" : "출처 없음") });
  if (match) {
    let url;
    try { url = new URL(match[0]); } catch { url = null; }
    if (url) {
      box.append(el("br"), el("a", { class: "link", href: url.href, target: "_blank", rel: "noopener noreferrer", text: url.hostname + url.pathname }));
    } else {
      box.append(" " + match[0]);
    }
  }
  return box;
}

function weekOptions() {
  const select = $("#week-filter");
  const previous = select.value;
  const weeks = [...new Set(data.news.map(n => mondayOf(n.date)).filter(Boolean))].sort().reverse();
  const options = [];
  if (data.week.start) options.push(["current", `이번 주 · ${data.week.start} ~ ${data.week.end.slice(5)}`]);
  options.push(["all", "전체 기간"]);
  weeks.forEach(monday => options.push([monday, `${monday} 주`]));
  select.replaceChildren(...options.map(([value, label]) => el("option", { value, text: label })));
  select.value = options.some(([value]) => value === previous) ? previous : options[0][0];
}

function inPeriod(news, period) {
  if (period === "all") return true;
  const start = period === "current" ? data.week.start : period;
  const end = period === "current" ? data.week.end : addDays(period, 6);
  return news.date >= start && news.date <= end;
}

function renderNews() {
  if (!data) return;
  const query = $("#news-search").value.trim().toLocaleLowerCase();
  const period = $("#week-filter").value;
  const inRange = data.news.filter(n => inPeriod(n, period));
  const visible = inRange.filter(n => (filterKind === "all" || n.kind === filterKind) &&
    [n.title, n.content, n.keyword, n.factor, n.region, n.source].join(" ").toLocaleLowerCase().includes(query));
  $("#filter-count").textContent = `${visible.length} / ${inRange.length}건 · 강도 +는 금리 상방(점수와 부호가 반대)`;
  const list = $("#news-list");
  list.replaceChildren(...visible.map(n => el("article", { class: "news" },
    el("div", { class: "news-heading" },
      el("span", { class: "tag", text: n.kind || "유형 없음" }),
      el("span", { text: [n.factor, n.region].filter(Boolean).join(" / ") }),
      el("span", { text: n.date || "날짜 없음" })),
    el("h3", { text: n.title || "(제목 없음)" }),
    n.content ? el("p", { text: n.content }) : null,
    sourceNode(n.source),
    el("span", { class: "direction", text: `방향 ${n.direction || "—"} · 강도 ${score(n.intensity)}${n.keyword ? " · " + n.keyword : ""}` }))));
  if (!visible.length) list.append(el("p", { class: "no-results", text: inRange.length ? "일치하는 원문이 없습니다. 검색어를 지우거나 다른 유형을 고르세요." : "이 기간에 기록된 뉴스가 없습니다." }));
}

function scoreCell(value, previous, tag, locked) {
  return el("td", { class: "score " + scoreClass(value) },
    score(value),
    el("span", { class: "sub", text: `(전주 ${score(previous)})` }),
    tag || locked ? el("span", { class: "chips" },
      tag ? el("span", { class: "tag", text: tag }) : null,
      locked ? el("span", { class: "tag good", text: "직접 수정" }) : null) : null);
}

function definitionList(target, pairs) {
  target.replaceChildren(...pairs.flatMap(([label, value]) => [el("dt", { text: label }), el("dd", { text: value || "—" })]));
}

function render() {
  const { state, week } = data;
  $("#empty").classList.add("hidden");
  $("#result").classList.remove("hidden");
  $("#week-label").textContent = week.start ? `기준 주간 ${week.start} ~ ${week.end}` : "기준 주간 정보 없음(Analysis I1 칸)";
  $("#confirm-label").textContent = state.confirmed ? "채점 확정" : "확정 전 초안";
  $("#state-badge").textContent = state.confirmed ? "채점 확정" : "확정 전";
  $("#state-badge").className = "badge " + (state.confirmed ? "good" : "subtle");
  $("#dur-total").textContent = score(state.dur.total, 2);
  $("#dur-verdict").textContent = state.dur.verdict;
  $("#cur-total").textContent = score(state.cur.total, 2);
  $("#cur-verdict").textContent = state.cur.verdict;
  $("#news-week").textContent = week.start ? data.news.filter(n => inPeriod(n, "current")).length + "건" : "—";
  $("#news-total").textContent = `전체 ${data.news.length}건`;
  $("#edited").textContent = state.edited_n + "칸";
  $("#conclusion").textContent = data.conclusion || "아직 적지 않았습니다.";
  weekOptions();
  renderNews();
  $("#rows").replaceChildren(...state.rows.map(row => el("tr", {},
    el("td", { class: "label", text: row.label }),
    scoreCell(row.dur, row.dur_prev, row.dur_tag, row.locked.includes("dur")),
    scoreCell(row.cur, row.cur_prev, row.cur_tag, row.locked.includes("cur")),
    el("td", { class: "note" },
      row.note ? el("p", { text: row.note }) : el("span", { class: "tag warn", text: "근거 미입력" }),
      row.digest ? el("p", { class: "ref", text: "뉴스 요약: " + row.digest }) : null,
      row.news_dur !== null || row.news_cur !== null
        ? el("p", { class: "ref", text: `뉴스 참고 듀 ${score(row.news_dur)} · 커브 ${score(row.news_cur)}` }) : null,
      row.house ? el("p", { class: "ref", text: "하우스: " + row.house }) : null))));
  $("#divergences").replaceChildren(...(data.divergences.length
    ? data.divergences.map(text => el("li", { text }))
    : [el("li", { text: "내 점수와 뉴스 참고 점수가 크게 다른 항목이 없습니다." })]));
  $("#divergence-count").textContent = data.divergences.length + "건";
  definitionList($("#memo"), data.memo.map(m => [m.label || "메모", m.value]));
  const f = state.flags;
  definitionList($("#flags"), [
    ["국면", f.regime], ["조기경보 ρ60", f.rho60 === null ? "" : String(f.rho60)],
    ["전환 경계(A→B)", f.transition], ["이벤트", f.event_on ? "ON" : "OFF"], ["TP 돌파", f.tp_break],
    ["미 근원 CPI 3% 미만", f.core_cpi_lt3 ? "Y" : "N"],
    ...(f.relax ? [["판정 완화", "국면 A′ + 근원 CPI 3% 미만 → U/W 한 단계 완화"]] : []),
  ]);
  buttons();
}

/* ---------- 이벤트 ---------- */
$("#open").addEventListener("click", openFile);
$("#reopen").addEventListener("click", reopen);
$("#reload").addEventListener("click", async () => {
  if (!handle) return;
  try {
    await readFile(await handle.getFile());
  } catch {
    message("파일을 다시 읽지 못했습니다. ‘기록 파일 열기’로 다시 골라 주세요.", "error");
  }
});
$("#close").addEventListener("click", closeFile);
$("#picker").addEventListener("change", async event => {
  const file = event.target.files[0];
  event.target.value = "";
  if (!file) return;
  stopWatching();
  handle = null;
  seenStamp = failedStamp = null;
  await readFile(file);
});
$("#week-filter").addEventListener("change", renderNews);
$("#news-search").addEventListener("input", renderNews);
document.querySelectorAll("[data-kind]").forEach(button => button.addEventListener("click", () => {
  filterKind = button.dataset.kind;
  document.querySelectorAll("[data-kind]").forEach(item => item.setAttribute("aria-pressed", String(item === button)));
  renderNews();
}));
document.addEventListener("visibilitychange", () => { if (!document.hidden) check(); });
document.addEventListener("dragover", event => {
  if (![...event.dataTransfer.types].includes("Files")) return;
  event.preventDefault();
  document.body.classList.add("dragging");
});
document.addEventListener("dragleave", event => {
  if (!event.relatedTarget) document.body.classList.remove("dragging");
});
document.addEventListener("drop", async event => {
  if (![...event.dataTransfer.types].includes("Files")) return;
  event.preventDefault();
  document.body.classList.remove("dragging");
  const item = [...event.dataTransfer.items].find(entry => entry.kind === "file");
  const pendingHandle = item && typeof item.getAsFileSystemHandle === "function" ? item.getAsFileSystemHandle() : null;
  const file = event.dataTransfer.files[0];
  const dropped = pendingHandle ? await pendingHandle.catch(() => null) : null;
  if (dropped && dropped.kind === "file") {
    await useHandle(dropped);
  } else if (file) {
    stopWatching();
    handle = null;
    seenStamp = failedStamp = null;
    await readFile(file);
  }
});

message("기록을 읽는 엔진을 준비하고 있습니다. 처음 방문하면 약 14MB를 내려받아 10초 안팎 걸립니다.", "busy");
buttons();
restore();
