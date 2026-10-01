// 정적 사이트 연결부 — 화면(app.js)이 부르는 /api/* 요청과 Excel 내려받기를
// 브라우저 작업 스레드(web/worker.mjs)의 시연 엔진으로 넘긴다. app.js보다 먼저 실행해야 한다.
(() => {
  "use strict";
  const routes = /^\/api\/(config|history|evidence|state|run|review|reapply)$/;
  const failure = JSON.stringify({ error: "브라우저에서 시연 엔진을 시작하지 못했습니다." });
  const pending = new Map();
  let seq = 0;
  let status = "loading";
  let notice = "브라우저에서 시연 엔진(파이썬)을 준비하고 있습니다. 처음 방문하면 약 10MB를 내려받아 10초 안팎 걸립니다.";
  let noticeKind = "busy";

  function say(text, kind) {
    notice = text;
    noticeKind = kind;
    const box = document.getElementById("message");
    if (!box) return;
    box.textContent = text;
    box.className = kind;
    box.setAttribute("role", kind === "error" ? "alert" : "status");
  }
  document.addEventListener("DOMContentLoaded", () => say(notice, noticeKind));

  function failAll() {
    status = "failed";
    say("브라우저에서 시연 엔진을 시작하지 못했습니다. 최신 Chrome·Edge·Safari·Firefox에서 다시 열어 주세요.", "error");
    pending.forEach(({ resolve }) => resolve({ status: 500, text: failure, error: "failed" }));
    pending.clear();
  }

  let worker = null;
  try {
    worker = new Worker(new URL("worker.mjs", document.currentScript.src), { type: "module" });
    worker.onerror = failAll;
    worker.onmessage = ({ data }) => {
      if (data.type === "ready") {
        status = "ready";
        say(data.persistent
          ? "준비됐습니다. ‘새 시연 실행’을 누르세요. 실행 기록은 이 브라우저에만 저장됩니다."
          : "준비됐습니다. 이 브라우저에서는 새로고침하면 실행 기록이 사라집니다.", "");
        return;
      }
      if (data.type === "failed") {
        console.error(data.error);
        failAll();
        return;
      }
      const entry = pending.get(data.id);
      if (entry) {
        pending.delete(data.id);
        entry.resolve(data);
      }
    };
  } catch (error) {
    console.error(error);
    failAll();
  }

  function call(payload) {
    if (!worker || status === "failed") return Promise.resolve({ status: 500, text: failure, error: "failed" });
    return new Promise(resolve => {
      const id = ++seq;
      pending.set(id, { resolve });
      worker.postMessage({ id, ...payload });
    });
  }

  const nativeFetch = window.fetch.bind(window);
  window.fetch = async (input, init = {}) => {
    const url = new URL(typeof input === "string" ? input : input.url, location.href);
    if (url.origin !== location.origin || !routes.test(url.pathname)) return nativeFetch(input, init);
    const method = (init.method || "GET").toUpperCase();
    const reply = await call({ type: "api", method, url: url.pathname + url.search, body: init.body ?? null });
    return new Response(reply.text, { status: reply.status, headers: { "Content-Type": "application/json; charset=utf-8" } });
  };

  // 화면의 'Excel 받기' 링크는 로컬 서버 주소(/download?id=…)를 가리킨다. 여기서는 작업 스레드가 만든 파일을 내려준다.
  document.addEventListener("click", async event => {
    const link = event.target instanceof Element ? event.target.closest("a#download") : null;
    if (!link) return;
    event.preventDefault();
    const runId = new URL(link.href, location.href).searchParams.get("id");
    const reply = await call({ type: "workbook", runId });
    if (!reply.bytes) {
      say(reply.error === "failed" ? "브라우저에서 시연 엔진을 시작하지 못했습니다." : reply.error, "error");
      return;
    }
    const blob = new Blob([reply.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const href = URL.createObjectURL(blob);
    const anchor = Object.assign(document.createElement("a"), { href, download: "macro-notes-demo.xlsx" });
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(href), 10000);
  }, true);
})();
