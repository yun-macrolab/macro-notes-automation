// 시연 엔진을 브라우저 작업 스레드에서 돌린다 — 실행하는 1~3초 동안 화면이 멈추지 않게.
// Pyodide(브라우저용 CPython)로 core.py·engine/·web/web_api.py를 로컬 시연과 같은 파일 그대로 실행한다.
// 실행 기록 폴더는 IndexedDB에 연결해 새로고침 뒤에도 이 브라우저에만 남긴다(서버로 보내지 않는다).
import { loadPyodide } from "../pyodide/pyodide.mjs";

const root = new URL("../", import.meta.url);
const RUNTIME = "/app/runtime";
let py;
let api;
let persistent = false;

const syncfs = populate => new Promise((resolve, reject) =>
  py.FS.syncfs(populate, error => (error ? reject(error) : resolve())));

async function fetchBytes(path) {
  const response = await fetch(new URL(path, root), { cache: "no-cache" });
  if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function boot() {
  py = await loadPyodide({ indexURL: new URL("pyodide/", root).href });
  const manifest = JSON.parse(new TextDecoder().decode(await fetchBytes("py/manifest.json")));
  const [wheels, files] = await Promise.all([
    Promise.all(manifest.wheels.map(name => fetchBytes("wheels/" + name))),
    Promise.all(manifest.files.map(name => fetchBytes("py/" + name))),
  ]);
  wheels.forEach(bytes => py.unpackArchive(bytes, "wheel"));
  manifest.files.forEach((name, index) => {
    const target = "/app/" + name;
    py.FS.mkdirTree(target.slice(0, target.lastIndexOf("/")));
    py.FS.writeFile(target, files[index]);
  });
  py.FS.mkdirTree(RUNTIME);
  try {
    py.FS.mount(py.FS.filesystems.IDBFS, {}, RUNTIME);
    await syncfs(true);
    persistent = true;
  } catch (error) {
    console.warn("IndexedDB를 쓸 수 없어 이번 방문 동안만 실행 기록을 남깁니다.", error);
  }
  py.runPython("import sys; sys.path[:0] = ['/app', '/app/web']; import web_api");
  api = py.globals.get("web_api");
  api.init(RUNTIME);
  if (persistent) await syncfs(false);
}

const ready = boot();
ready.then(
  () => postMessage({ type: "ready", persistent }),
  error => postMessage({ type: "failed", error: String(error?.message || error) }),
);

async function handle(message) {
  try {
    await ready;
  } catch {
    postMessage({ id: message.id, status: 500, text: JSON.stringify({ error: "시연 엔진을 시작하지 못했습니다." }) });
    return;
  }
  if (message.type === "api") {
    const result = api.request(message.method, message.url, message.body ?? null);
    const [status, text] = result.toJs();
    result.destroy();
    if (persistent && message.method === "POST" && status === 200) {
      try {
        await syncfs(false);
      } catch (error) {
        console.warn("실행 기록을 브라우저 저장소에 쓰지 못했습니다.", error);
      }
    }
    postMessage({ id: message.id, status, text });
  } else if (message.type === "workbook") {
    try {
      const proxy = api.workbook(message.runId);
      const bytes = proxy.toJs().slice();
      proxy.destroy();
      postMessage({ id: message.id, bytes: bytes.buffer }, [bytes.buffer]);
    } catch {
      postMessage({ id: message.id, error: "Excel 파일을 만들지 못했습니다. 실행 기록을 다시 확인하세요." });
    }
  }
}

// 요청은 받은 순서대로 하나씩 처리한다(로컬 서버의 잠금 하나와 같은 역할).
let queue = Promise.resolve();
onmessage = event => {
  queue = queue.then(() => handle(event.data)).catch(error => console.error(error));
};
