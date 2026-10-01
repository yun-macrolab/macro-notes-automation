// 내 기록 보기 — 고른 기록 워크북을 브라우저 작업 스레드에서 읽는다(읽는 동안 화면이 멈추지 않게).
// Pyodide(브라우저용 CPython)가 engine/의 읽기 함수와 web/workbook_view.py를 그대로 실행한다.
// 파일은 이 스레드의 메모리에 잠시 올려 읽고 바로 지운다. 저장하거나 어디로도 보내지 않는다.
import { loadPyodide } from "../pyodide/pyodide.mjs";

const root = new URL("../", import.meta.url);
let py;
let view;

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
  py.runPython("import sys; sys.path[:0] = ['/app/engine', '/app/web']; import workbook_view");
  view = py.globals.get("workbook_view");
}

const ready = boot();
ready.then(
  () => postMessage({ type: "ready" }),
  error => postMessage({ type: "failed", error: String(error?.message || error) }),
);

async function handle(message) {
  try {
    await ready;
  } catch {
    postMessage({ id: message.id, ok: false, text: JSON.stringify({ error: "기록을 읽는 엔진을 시작하지 못했습니다." }) });
    return;
  }
  const path = "/tmp/record" + (/\.xlsm$/i.test(message.name || "") ? ".xlsm" : ".xlsx");
  py.FS.writeFile(path, new Uint8Array(message.bytes));
  try {
    const result = view.read_json(path);
    const [ok, text] = result.toJs();
    result.destroy();
    postMessage({ id: message.id, ok, text });
  } finally {
    py.FS.unlink(path);
  }
}

// 요청은 받은 순서대로 하나씩 처리한다.
let queue = Promise.resolve();
onmessage = event => {
  queue = queue.then(() => handle(event.data)).catch(error => {
    console.error(error);
    postMessage({ id: event.data.id, ok: false, text: JSON.stringify({ error: "기록 파일을 읽는 중 문제가 생겼습니다." }) });
  });
};
