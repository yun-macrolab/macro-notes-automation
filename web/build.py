"""정적 사이트 조립 — GitHub Pages에 올릴 폴더를 만든다(표준 라이브러리만 사용).

로컬 시연(app.py + static/)은 그대로 두고, 같은 화면과 같은 엔진 파일을 브라우저 안에서 돌리는 사본을 만든다.
화면 파일은 정적 사이트에 맞지 않는 경로와 '로컬 서버' 문구만 바꾸고, engine/·core.py는 바이트 그대로 싣는다.

  python -X utf8 web/build.py --pyodide <pyodide npm 패키지의 package 폴더> --wheels <휠 폴더> --out _site
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys

BASE = Path(__file__).resolve().parent.parent
README_URL = "https://github.com/yun-macrolab/macro-notes-automation#readme"
PYODIDE_FILES = ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"]
WHEELS = ["et_xmlfile-2.0.0-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl"]

# 원문이 바뀌어 한 곳이라도 정확히 한 번 찾지 못하면 조립을 멈춘다(조용히 어긋나지 않게).
HTML_EDITS = [
    ('<link rel="stylesheet" href="/style.css"><script src="/app.js" defer></script>',
     '<link rel="stylesheet" href="style.css"><script src="web/bridge.js"></script><script src="app.js" defer></script>'),
    ('<a class="brand" href="/"', '<a class="brand" href="./"'),
    ('href="/readme.md"', f'href="{README_URL}"'),
    ("이 브라우저가 연결된 로컬 서버의 실행 이력입니다.", "이 브라우저에만 저장된 실행 이력입니다."),
]
JS_EDITS = [
    ("' 다시 시도하거나 로컬 서버의 실행 상태를 확인하세요.'", "' 다시 시도하거나 페이지를 새로고침하세요.'"),
    ("'로컬 서버에 연결하지 못했습니다. 서버 실행 후 페이지를 새로고침하세요.'",
     "'브라우저에서 시연 엔진을 시작하지 못했습니다. 페이지를 새로고침하세요.'"),
    ("'서버의 evidence 파일을 확인한 후 새로고침하세요.'", "'페이지를 새로고침한 뒤 다시 확인하세요.'"),
]


def edited(path, edits):
    text = path.read_text(encoding="utf-8")
    for old, new in edits:
        if text.count(old) != 1:
            sys.exit(f"{path.relative_to(BASE)}: 바꿀 문구를 정확히 한 번 찾지 못했습니다 — {old[:50]}")
        text = text.replace(old, new)
    return text


def copy(src, dst):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pyodide", required=True, type=Path)
    parser.add_argument("--wheels", required=True, type=Path)
    parser.add_argument("--out", default=BASE / "_site", type=Path)
    args = parser.parse_args()
    out = args.out.resolve()
    if out == BASE or BASE in out.parents and out.name != "_site":
        sys.exit("--out은 저장소 밖 폴더이거나 저장소 안의 _site여야 합니다.")
    if out.exists():
        if any(out.iterdir()) and not (out / "py/manifest.json").exists():
            sys.exit(f"{out}: 이전 조립 결과가 아닌 폴더는 지우지 않습니다.")
        shutil.rmtree(out)
    out.mkdir(parents=True)

    (out / "index.html").write_text(edited(BASE / "static/index.html", HTML_EDITS), encoding="utf-8")
    (out / "app.js").write_text(edited(BASE / "static/app.js", JS_EDITS), encoding="utf-8")
    copy(BASE / "static/style.css", out / "style.css")
    for name in ("bridge.js", "worker.mjs"):
        copy(BASE / "web" / name, out / "web" / name)

    for name in PYODIDE_FILES:
        copy(args.pyodide / name, out / "pyodide" / name)
    for name in WHEELS:
        copy(args.wheels / name, out / "wheels" / name)

    files = ["core.py", "data/sample.json", "evidence/results.json", "web/web_api.py"]
    files += sorted(f"engine/{p.name}" for p in (BASE / "engine").glob("*.py"))
    for name in files:
        copy(BASE / name, out / "py" / name)
    manifest = {
        "files": files,
        "wheels": WHEELS,
        "commit": os.environ.get("GITHUB_SHA", "local"),
        # 줄바꿈 차이로 해시가 OS마다 달라지지 않게 LF로 맞춰 계산한다(확인용 기록).
        "engine_sha256_lf": {
            Path(name).name: hashlib.sha256((BASE / name).read_bytes().replace(b"\r\n", b"\n")).hexdigest()
            for name in files if name.startswith("engine/")
        },
    }
    (out / "py/manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / ".nojekyll").write_text("")
    size = sum(p.stat().st_size for p in out.rglob("*") if p.is_file())
    print(f"정적 사이트 조립: {out} · 파일 {sum(1 for p in out.rglob('*') if p.is_file())}개 · {size / 1e6:.1f}MB")


if __name__ == "__main__":
    main()
