// 매크로 기록 보기 — 공개한 주간 기록(records/)을 보여 주고, 내 PC의 기록 워크북도 열어 본다.
// 내 파일 내용은 브라우저 작업 스레드(web/worker.mjs)로만 넘기고 서버나 저장소로 보내지 않는다.
// Chrome·Edge는 파일 핸들을 기억해 3초마다 수정 시각을 확인하고, 다른 브라우저는 파일을 다시 고를 때 읽는다.
// 기록을 읽는 엔진(약 14MB)은 내 파일을 쓸 때만 내려받는다. 공개 기록은 JSON만 받는다.
// '공개본 만들기'는 화면에 띄운 내 기록에서 정해 둔 칸만 뽑아 검사한 뒤 JSON 파일로 내려받는다(올리는 것은 사람이 한다).
// '여러 주 한 번에'는 고른 주간 파일마다 같은 검사를 하고, 통과한 주만 zip 하나로 묶어 내려받는다.
"use strict";
const $ = selector => document.querySelector(selector);
const POLL_MS = 3000;
const DB = { name: "macro-notes-viewer", store: "handles", key: "last" };
const WORDS_KEY = "macro-notes-viewer.check-words";
const PUBLIC_SCHEMA = "macro-notes-public/1";
const RECORD_FILE = /^\d{4}-\d{2}-\d{2}\.json$/;
const BULK_MAX = 300;     // 폴더째 고를 때 읽을 엑셀 파일 수 상한
const SKIP_REASONS = { copy: "복사본·백업 파일", v1: "예전 형식(v1)", no_week: "주간 날짜 없음(I1 칸)", not_record: "기록 워크북 아님", broken: "읽지 못함", blocked: "검사에 걸림" };
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
let localData = null;     // 내 파일에서 읽은 기록
let publicList = [];      // records/index.json의 목록(최신 주가 먼저)
let publicData = null;    // 고른 공개 기록
let publicFile = null;
let publicFailed = false;
let publicSeq = 0;
let tab = "local";        // 화면에 그리는 쪽: "public" | "local"
let data = null;          // 지금 화면에 그린 기록(publicData 또는 localData)
let reading = false;
let exporting = false;
let bulkRunning = false;
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

/* ---------- 작업 스레드(내 파일을 쓸 때만 시작) ---------- */
const ENGINE_FAILED = JSON.stringify({ error: "기록을 읽는 엔진을 시작하지 못했습니다." });

function engineFailure() {
  engineFailed = true;
  message("기록을 읽는 엔진을 시작하지 못했습니다. 최신 Chrome·Edge·Safari·Firefox에서 다시 열어 주세요.", "error");
  pending.forEach(({ resolve }) => resolve({ ok: false, text: ENGINE_FAILED }));
  pending.clear();
}

// 내 파일 탭에서 지금 할 일을 알려 준다.
function localStatus() {
  if (engineFailed) {
    message("기록을 읽는 엔진을 시작하지 못했습니다. 최신 Chrome·Edge·Safari·Firefox에서 다시 열어 주세요.", "error");
  } else if (!engineReady) {
    message("기록을 읽는 엔진을 준비하고 있습니다. 처음 한 번은 약 14MB를 내려받아 10초 안팎 걸립니다.", "busy");
  } else if (!localData) {
    message(canWatch
      ? "준비됐습니다. ‘기록 파일 열기’를 누르세요. 엑셀에서 저장할 때마다 자동으로 다시 읽습니다."
      : "준비됐습니다. ‘기록 파일 열기’를 누르세요. 이 브라우저에서는 저장한 뒤 파일을 다시 골라야 바뀐 내용이 보입니다.");
  } else {
    message("");
  }
}

function startEngine() {
  if (worker || engineFailed) return;
  try {
    worker = new Worker(new URL("web/worker.mjs", location.href), { type: "module" });
  } catch (error) {
    console.error(error);
    engineFailure();
    return;
  }
  worker.onerror = engineFailure;
  worker.onmessage = ({ data: reply }) => {
    if (reply.type === "ready") {
      engineReady = true;
      if (tab === "local" && !reading) localStatus();
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
}

function ask(request, transfer = []) {
  startEngine();
  if (engineFailed || !worker) return Promise.resolve({ ok: false, text: ENGINE_FAILED });
  return new Promise(resolve => {
    const id = ++seq;
    pending.set(id, { resolve });
    worker.postMessage({ id, ...request }, transfer);
  });
}

/* ---------- 내 파일 읽기 ---------- */
async function readFile(file, { quiet = false } = {}) {
  if (reading) return false;
  reading = true;
  buttons();
  if (!quiet) {
    message(engineReady ? "기록 파일을 읽고 있습니다." : "기록을 읽는 엔진을 준비하고 있습니다. 처음 한 번은 약 14MB를 내려받아 10초 안팎 걸립니다.", "busy");
  }
  // 공개 기록 탭을 보는 동안 뒤에서 다시 읽은 결과는 알리지 않는다(내 파일 탭에서 보인다).
  const tell = (text, kind) => { if (!quiet || tab === "local") message(text, kind); };
  try {
    const bytes = await file.arrayBuffer();
    const reply = await ask({ type: "read", bytes, name: file.name }, [bytes]);
    const body = JSON.parse(reply.text);
    if (!reply.ok) {
      failedStamp = stamp(file);
      tell(body.error, "error");
      return false;
    }
    localData = body;
    lastFile = file;
    seenStamp = stamp(file);
    failedStamp = null;
    exportScope();
    if (tab === "local") show();
    tell(quiet ? `저장된 변경을 반영했습니다(${clock(Date.now())}).` : "기록을 불러왔습니다.");
    return true;
  } catch (error) {
    console.error(error);
    failedStamp = stamp(file);
    tell("파일을 읽지 못했습니다. 다시 열어 주세요.", "error");
    return false;
  } finally {
    reading = false;
    buttons();
  }
}

// 다른 파일로 바꿀 때: 지켜보던 핸들과 지난 공개본 결과를 내려놓는다.
function forgetHandle() {
  stopWatching();
  handle = null;
  seenStamp = failedStamp = null;
  $("#export-result").replaceChildren();
}

async function useHandle(next) {
  forgetHandle();
  handle = next;
  remembered = next;
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

async function closeFile() {
  forgetHandle();
  remembered = lastFile = localData = null;
  show();
  // 기억해 둔 위치를 다 지운 뒤에 알린다(바로 새로고침해도 다시 열리지 않게).
  await idb("readwrite", store => store.delete(DB.key)).catch(() => {});
  message("파일을 닫았습니다. 화면에 띄웠던 내용과 기억해 둔 파일 위치도 지웠습니다.");
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

// 지난 방문에 연 파일이 있으면 기억해 두고, 다시 읽을 권한이 이미 있는지 알려 준다.
async function rememberedHandle() {
  if (!canWatch) return false;
  try {
    remembered = (await idb("readonly", store => store.get(DB.key))) || null;
  } catch {
    remembered = null;
  }
  if (!remembered) return false;
  try {
    return (await remembered.queryPermission({ mode: "read" })) === "granted";
  } catch {
    // 권한을 물어보려면 사용자의 클릭이 필요하다. ‘지난 파일 다시 열기’ 버튼으로 받는다.
    return false;
  }
}

/* ---------- 공개 기록(records/) ---------- */
async function loadIndex() {
  try {
    const response = await fetch("records/index.json", { cache: "no-cache" });
    if (!response.ok) return [];
    const body = await response.json();
    return Array.isArray(body.records) ? body.records.filter(entry => entry && RECORD_FILE.test(entry.file)) : [];
  } catch (error) {
    console.warn(error);
    return [];
  }
}

function publicOptions() {
  $("#public-week").replaceChildren(...publicList.map((entry, index) => el("option", {
    value: entry.file,
    text: `${entry.start} ~ ${String(entry.end).slice(5)}${index === 0 ? " · 최신" : ""}${entry.confirmed === false ? " · 확정 전" : ""}`,
  })));
}

async function showPublic(file, { remember = false } = {}) {
  const mine = ++publicSeq;
  $("#public-week").value = file;
  if (tab === "public") message("공개 기록을 불러오고 있습니다.", "busy");
  try {
    const response = await fetch("records/" + file, { cache: "no-cache" });
    if (!response.ok) throw new Error(`records/${file}: HTTP ${response.status}`);
    const record = await response.json();
    if (record.schema !== PUBLIC_SCHEMA) throw new Error(`records/${file}: schema ${record.schema}`);
    if (mine !== publicSeq) return;
    publicData = record;
    publicFile = file;
    publicFailed = false;
    if (remember) history.replaceState(null, "", "#" + record.week.start);
    if (tab === "public") {
      message("");
      show();
    }
  } catch (error) {
    if (mine !== publicSeq) return;
    console.warn(error);
    publicFailed = !publicData;
    if (publicFile) $("#public-week").value = publicFile;
    if (tab === "public") {
      message("공개 기록을 불러오지 못했습니다. 잠시 뒤 다시 열어 주세요.", "error");
      show();
    }
  }
}

function publicMeta() {
  const meta = $("#public-meta");
  meta.replaceChildren();
  if (!publicData) return;
  const made = Date.parse(publicData.generated_at);
  if (!Number.isNaN(made)) meta.append(`${clock(made)}에 만든 공개본 · `);
  meta.append(el("a", { href: "records/" + publicFile, target: "_blank", rel: "noopener", text: "JSON 원본 ↗" }));
}

/* ---------- 탭 ---------- */
function showTab(name, focus = false) {
  tab = name;
  document.querySelectorAll("[data-tab]").forEach(button => {
    const selected = button.dataset.tab === name;
    button.classList.toggle("active", selected);
    button.setAttribute("aria-selected", String(selected));
    button.tabIndex = selected ? 0 : -1;
    if (selected && focus) button.focus();
  });
  $("#view").classList.remove("hidden");
  if (publicList.length) $("#view").setAttribute("aria-labelledby", "tab-" + name);
  $("#public-bar").classList.toggle("hidden", name !== "public");
  $("#local-tools").classList.toggle("hidden", name !== "local");
  if (name === "local") {
    startEngine();
    if (!reading) localStatus();
  } else {
    message(publicFailed ? "공개 기록을 불러오지 못했습니다. 잠시 뒤 다시 열어 주세요." : "", publicFailed ? "error" : "");
  }
  show();
}

function show() {
  data = tab === "public" ? publicData : localData;
  $("#empty").classList.toggle("hidden", tab !== "local" || !!data);
  $("#public-empty").classList.toggle("hidden", tab !== "public" || !!data);
  $("#public-empty-title").textContent = publicFailed ? "공개 기록을 불러오지 못했습니다." : "공개 기록을 불러오고 있습니다.";
  if (data) render();
  else $("#result").classList.add("hidden");
  if (tab === "public") publicMeta();
  buttons();
}

/* ---------- 공개본 만들기 ---------- */
function loadWords() {
  try {
    return localStorage.getItem(WORDS_KEY) || "";
  } catch {
    return "";
  }
}

function saveWords() {
  try {
    localStorage.setItem(WORDS_KEY, $("#check-words").value);
  } catch {
    // 저장하지 못해도 이번 검사에는 쓴다.
  }
}

function exportScope() {
  const week = localData && localData.week;
  $("#export-scope").textContent = week && week.start
    ? `화면에 띄운 기록 중 기준 주간 ${week.start} ~ ${week.end}의 내용만 ${week.start}.json 파일로 내려받습니다.`
    : "Analysis I1 칸에 주간 날짜가 없어 공개본을 만들 수 없습니다.";
}

function toggleExport(open = $("#export-panel").classList.contains("hidden")) {
  $("#export-panel").classList.toggle("hidden", !open);
  $("#export").setAttribute("aria-expanded", String(open));
  if (open) exportScope();
}

const checkWords = () => $("#check-words").value.split("\n").map(word => word.trim()).filter(Boolean);
const confirmText = value => (value === true ? "채점 확정" : value === false ? "확정 전" : "확정 표시 없음");

function download(name, content, type = "application/json") {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const link = el("a", { href: url, download: name, class: "hidden" });
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function blockedView(body) {
  const found = body.findings || [];
  return el("div", { class: "export-blocked", role: "alert" },
    el("strong", { text: body.error }),
    found.length ? el("ul", {}, ...found.slice(0, 20).map(f => el("li", { text: `${f.place} — ${f.kind} ‘${f.text}’` }))) : null,
    found.length > 20 ? el("p", { text: `외 ${found.length - 20}곳` }) : null);
}

function doneView(record) {
  const name = record.week.start + ".json";
  return el("div", { class: "export-done" },
    el("strong", { text: `${name}을 내려받았습니다.` }),
    el("p", { text: `뉴스 ${record.news.length}건 · 점수 ${record.state.rows.length}항목 · 메모 ${record.memo.length}칸이 실렸습니다.` }),
    el("ol", {},
      el("li", {}, "저장소의 records 폴더 올리기 화면을 엽니다. ",
        el("a", { href: $("#export-panel").dataset.upload, target: "_blank", rel: "noopener", text: "records에 올리기 ↗" })),
      el("li", { text: `내려받은 파일을 고치지 말고 끌어다 놓은 뒤 ‘Commit changes’를 누릅니다. 파일 이름이 ${name}인지 확인하세요(같은 이름이 이미 있으면 브라우저가 (1)을 붙입니다).` }),
      el("li", { text: "1~2분 뒤 배포가 끝나면 ‘공개 기록’ 탭에 보입니다." })));
}

async function runExport() {
  if (!localData || exporting) return;
  exporting = true;
  buttons();
  const result = $("#export-result");
  result.replaceChildren(el("p", { class: "help", text: "검사하고 있습니다." }));
  try {
    const reply = await ask({ type: "public", text: JSON.stringify(localData), words: checkWords() });
    const body = JSON.parse(reply.text);
    if (!reply.ok) {
      result.replaceChildren(blockedView(body));
      return;
    }
    download(body.week.start + ".json", reply.text);
    result.replaceChildren(doneView(body));
  } catch (error) {
    console.error(error);
    result.replaceChildren(el("div", { class: "export-blocked", role: "alert", text: "공개본을 만들지 못했습니다. 다시 눌러 주세요." }));
  } finally {
    exporting = false;
    buttons();
  }
}

/* ---------- 여러 주 한 번에 ---------- */
// 엑셀 기록 파일만 고른다(엑셀이 열어 둔 동안 생기는 ‘~$’ 잠금 파일은 뺀다).
// label은 고른 폴더 아래 경로다. 복사본·백업 판정(public_record.is_copy)에 고른 폴더 자체의 이름은 넣지 않는다.
const isBook = name => /\.(xlsx|xlsm)$/i.test(name) && !name.startsWith("~$");
const bulkItems = files => [...files].filter(file => isBook(file.name))
  .map(file => ({ file, label: file.webkitRelativePath ? file.webkitRelativePath.split("/").slice(1).join("/") : file.name }));

// 끌어다 놓은 폴더(Chrome·Edge)를 하위 폴더까지 훑는다(prefix는 그 폴더 아래 경로).
async function directoryItems(directory, prefix, out = []) {
  for await (const entry of directory.values()) {
    if (out.length >= BULK_MAX) break;
    if (entry.kind === "directory") await directoryItems(entry, prefix + entry.name + "/", out);
    else if (isBook(entry.name)) out.push({ file: await entry.getFile(), label: prefix + entry.name });
  }
  return out;
}

// 같은 주 파일이 여럿이면 채점 확정한 파일, 그다음 최근에 저장한 파일을 쓴다.
function preferred(a, b) {
  const confirmed = row => row.record.state.confirmed === true;
  if (confirmed(a) !== confirmed(b)) return confirmed(a);
  return a.file.lastModified > b.file.lastModified;
}

async function runBulk(items) {
  if (bulkRunning) return;
  const result = $("#bulk-result");
  if (!items.length) {
    result.replaceChildren(el("div", { class: "export-blocked", role: "alert", text: "고른 것 중에 엑셀 기록 파일(.xlsx·.xlsm)이 없습니다." }));
    return;
  }
  bulkRunning = true;
  buttons();
  const progress = el("p", { class: "help" });
  result.replaceChildren(progress);
  const words = checkWords();
  const rows = [];
  try {
    for (const [index, item] of items.slice(0, BULK_MAX).entries()) {
      progress.textContent = `${index + 1} / ${Math.min(items.length, BULK_MAX)} 확인하는 중 · ${item.label}`
        + (engineReady ? "" : " (처음 한 번은 엔진을 준비하느라 10초 안팎 걸립니다)");
      let reply;
      try {
        const bytes = await item.file.arrayBuffer();
        reply = await ask({ type: "bulk", bytes, name: item.file.name, label: item.label, words }, [bytes]);
      } catch (error) {
        console.error(error);
        reply = { ok: false, text: JSON.stringify({ reason: "broken", error: "파일을 읽지 못했습니다." }) };
      }
      const body = JSON.parse(reply.text);
      rows.push(reply.ok ? { ...item, ok: true, text: reply.text, record: body } : { ...item, ok: false, ...body });
    }
    const chosen = new Map();
    rows.filter(row => row.ok).forEach(row => {
      const held = chosen.get(row.record.week.start);
      if (!held || preferred(row, held)) chosen.set(row.record.week.start, row);
    });
    const picked = [...chosen.values()].sort((a, b) => b.record.week.start.localeCompare(a.record.week.start));
    let zipName = null;
    if (picked.length) {
      progress.textContent = `공개본 ${picked.length}주를 zip으로 묶는 중`;
      const reply = await ask({ type: "zip", text: JSON.stringify(picked.map(row => ({ name: row.record.week.start + ".json", file: row.text }))) });
      if (!reply.ok) throw new Error("zip: " + reply.text);
      const [first, last] = [picked[picked.length - 1].record.week.start, picked[0].record.week.start];
      zipName = first === last ? `공개본_${first}.zip` : `공개본_${first}_${last}.zip`;
      download(zipName, reply.bytes, "application/zip");
    }
    result.replaceChildren(bulkView(rows, picked, zipName));
  } catch (error) {
    console.error(error);
    result.replaceChildren(el("div", { class: "export-blocked", role: "alert", text: "공개본을 만드는 중 문제가 생겼습니다. 다시 골라 주세요." }));
  } finally {
    bulkRunning = false;
    buttons();
  }
}

function bulkView(rows, picked, zipName) {
  const used = new Set(picked);
  const skipped = rows.filter(row => !row.ok).sort((a, b) => a.label.localeCompare(b.label));
  const duplicates = rows.filter(row => row.ok && !used.has(row));
  const head = zipName
    ? el("div", { class: "export-done" },
      el("strong", { text: `공개본 ${picked.length}주를 ${zipName}로 내려받았습니다.` }),
      el("p", { text: `${picked[picked.length - 1].record.week.start} ~ ${picked[0].record.week.start} 주`
        + (skipped.length ? ` · 건너뛴 파일 ${skipped.length}개` : "") + (duplicates.length ? ` · 같은 주라 뺀 파일 ${duplicates.length}개` : "") }),
      el("ol", {},
        el("li", { text: "zip을 풀고, 안의 .json 파일만 모두 고릅니다(폴더째 올리면 records 아래에 폴더가 생겨 배포 검사에서 거절됩니다)." }),
        el("li", {}, el("a", { href: $("#export-panel").dataset.upload, target: "_blank", rel: "noopener", text: "records에 올리기 ↗" }),
          "를 열어 끌어다 놓고 ‘Commit changes’를 누릅니다. 이미 있는 주는 새 내용으로 바뀝니다."),
        el("li", { text: "1~2분 뒤 배포가 끝나면 ‘공개 기록’ 탭에서 주간을 골라 볼 수 있습니다." })))
    : el("div", { class: "export-blocked", role: "alert", text: "공개본을 만들 수 있는 주간 파일이 없었습니다. 아래 표의 이유를 확인하세요." });
  const table = el("div", { class: "table-scroll", role: "region", "aria-label": "파일별 결과", tabindex: "0" },
    el("table", { class: "bulk-table" },
      el("thead", {}, el("tr", {}, ...["주간", "파일", "뉴스", "듀레이션", "커브", "결과"].map(text => el("th", { scope: "col", text })))),
      el("tbody", {}, ...[...picked, ...duplicates, ...skipped].map(row => bulkRow(row, used)))));
  const blocked = skipped.filter(row => row.reason === "blocked");
  const findings = blocked.length ? el("details", { class: "bulk-findings" },
    el("summary", { text: `검사에 걸린 위치 · 파일 ${blocked.length}개` }),
    ...blocked.map(row => el("div", {},
      el("strong", { text: row.label }),
      el("ul", {}, ...(row.findings || []).slice(0, 10).map(f => el("li", { text: `${f.place} — ${f.kind} ‘${f.text}’` })))))) : null;
  return el("div", { class: "bulk-view" }, head, findings, table);
}

function bulkRow(row, used) {
  const record = row.record;
  const week = record ? record.week : row.week;
  const state = record && record.state;
  const status = !row.ok ? `건너뜀 · ${SKIP_REASONS[row.reason] || "읽지 못함"}`
    : used.has(row) ? `포함 · ${confirmText(state.confirmed)}` : "같은 주의 다른 파일을 씀";
  return el("tr", { class: row.ok && used.has(row) ? "" : "muted-row" },
    el("td", { text: week ? `${week.start} ~ ${week.end.slice(5)}` : "—" }),
    el("td", { class: "file", text: row.label }),
    el("td", { text: record ? record.news.length + "건" : "—" }),
    el("td", { text: state ? `${score(state.dur.total, 2)} ${state.dur.verdict}` : "—" }),
    el("td", { text: state ? `${score(state.cur.total, 2)} ${state.cur.verdict}` : "—" }),
    el("td", { text: status }));
}

/* ---------- 화면 ---------- */
function buttons() {
  $("#open").disabled = reading;
  $("#reload").disabled = reading;
  $("#reload").classList.toggle("hidden", !handle);
  $("#export-one").classList.toggle("hidden", !localData);
  $("#export-run").disabled = exporting || bulkRunning || !localData;
  $("#bulk-pick").disabled = $("#bulk-folder-pick").disabled = bulkRunning;
  $("#close").classList.toggle("hidden", !localData && !remembered);
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
  if (data.week.start) options.push(["current", `기준 주간 · ${data.week.start} ~ ${data.week.end.slice(5)}`]);
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
  // 공개본에는 그 주 뉴스만 있어 기간을 고르지 않는다.
  const period = tab === "public" ? "all" : $("#week-filter").value;
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
  const isPublic = tab === "public";
  $("#result").classList.remove("hidden");
  $("#week-label").textContent = week.start ? `기준 주간 ${week.start} ~ ${week.end}` : "기준 주간 정보 없음(Analysis I1 칸)";
  // confirmed가 null이면 채점 확정 칸이 생기기 전(자동 초안 도입 전) 파일이다.
  $("#confirm-label").textContent = state.confirmed === true ? "채점 확정"
    : state.confirmed === false ? "확정 전 초안" : "확정 표시 없음(자동 초안 도입 전 파일)";
  $("#state-badge").textContent = confirmText(state.confirmed);
  $("#state-badge").className = "badge " + (state.confirmed === true ? "good" : "subtle");
  $("#dur-total").textContent = score(state.dur.total, 2);
  $("#dur-verdict").textContent = state.dur.verdict;
  $("#cur-total").textContent = score(state.cur.total, 2);
  $("#cur-verdict").textContent = state.cur.verdict;
  $("#news-week").textContent = week.start ? data.news.filter(n => inPeriod(n, "current")).length + "건" : "—";
  $("#news-total").textContent = isPublic ? "공개본에 실린 뉴스" : `전체 ${data.news.length}건`;
  $("#edited").textContent = state.confirmed === null ? "—" : state.edited_n + "칸";
  $("#edited-note").textContent = state.confirmed === null ? "자동 초안 도입 전 기록" : "자동 초안이 다시 덮지 않는 칸";
  $("#conclusion").textContent = data.conclusion || "아직 적지 않았습니다.";
  $("#week-filter-box").classList.toggle("hidden", isPublic);
  if (!isPublic) weekOptions();
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
}

/* ---------- 이벤트 ---------- */
document.querySelectorAll("[data-tab]").forEach((button, index, tabs) => {
  button.addEventListener("click", () => { if (tab !== button.dataset.tab) showTab(button.dataset.tab); });
  button.addEventListener("keydown", event => {
    const next = event.key === "ArrowRight" ? (index + 1) % tabs.length
      : event.key === "ArrowLeft" ? (index + tabs.length - 1) % tabs.length
      : event.key === "Home" ? 0 : event.key === "End" ? tabs.length - 1 : null;
    if (next !== null) {
      event.preventDefault();
      showTab(tabs[next].dataset.tab, true);
    }
  });
});
$("#public-week").addEventListener("change", event => showPublic(event.target.value, { remember: true }));
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
$("#export").addEventListener("click", () => toggleExport());
$("#export-run").addEventListener("click", runExport);
$("#check-words").addEventListener("input", saveWords);
$("#bulk-pick").addEventListener("click", () => $("#bulk-files").click());
$("#bulk-folder-pick").addEventListener("click", () => $("#bulk-folder").click());
["#bulk-files", "#bulk-folder"].forEach(id => $(id).addEventListener("change", async event => {
  const files = [...event.target.files];
  event.target.value = "";
  if (files.length) await runBulk(bulkItems(files));
}));
$("#picker").addEventListener("change", async event => {
  const file = event.target.files[0];
  event.target.value = "";
  if (!file) return;
  forgetHandle();
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
  // 핸들은 이벤트 안에서 바로 받아야 한다(기다린 뒤에는 끌어다 놓은 항목에 접근할 수 없다).
  const pendingHandles = [...event.dataTransfer.items].filter(entry => entry.kind === "file")
    .map(entry => (typeof entry.getAsFileSystemHandle === "function" ? entry.getAsFileSystemHandle().catch(() => null) : null));
  const files = [...event.dataTransfer.files];
  if (tab !== "local") showTab("local");
  const dropped = await Promise.all(pendingHandles);
  const folders = dropped.filter(entry => entry && entry.kind === "directory");
  // 파일 여러 개나 폴더를 끌어다 놓으면 '여러 주 한 번에'
  if (files.length > 1 || folders.length) {
    toggleExport(true);
    let items = bulkItems(files);
    for (const folder of folders) items = items.concat(await directoryItems(folder, ""));
    await runBulk(items);
  } else if (dropped[0] && dropped[0].kind === "file") {
    await useHandle(dropped[0]);
  } else if (files[0]) {
    forgetHandle();
    await readFile(files[0]);
  }
});

/* ---------- 시작: 공개 기록이 있으면 그 탭을, 없으면 내 파일 화면을 연다 ---------- */
async function init() {
  $("#check-words").value = loadWords();
  const [list, granted] = await Promise.all([loadIndex(), rememberedHandle()]);
  publicList = list;
  $("#source-tabs").classList.toggle("hidden", !publicList.length);
  if (!publicList.length) {
    showTab("local");
    if (granted) await useHandle(remembered);
    return;
  }
  publicOptions();
  $("#view").setAttribute("role", "tabpanel");
  const linked = publicList.find(entry => "#" + entry.start === location.hash);
  if (granted && !linked) {
    // 지난번에 연 내 파일을 다시 읽을 권한이 이미 있다(이 PC의 주인). 내 파일부터 보여 준다.
    showTab("local");
    showPublic(publicList[0].file);
    await useHandle(remembered);
  } else {
    showTab("public");
    await showPublic((linked || publicList[0]).file);
  }
}

buttons();
init().catch(error => {
  // 공개 기록을 못 읽어도 내 파일 화면은 쓸 수 있게 한다.
  console.error(error);
  showTab("local");
});
