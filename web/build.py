"""정적 사이트 조립 — GitHub Pages에 올릴 '내 기록 보기' 폴더를 만든다(표준 라이브러리만 사용).

사이트에는 데이터가 없다. 방문자가 자기 PC의 기록 워크북을 고르면 브라우저 안에서 읽는다.
화면 디자인은 시연 화면(static/style.css)을 그대로 쓰고, engine/은 바이트 그대로 싣는다.

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
PYODIDE_FILES = ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"]
WHEELS = ["et_xmlfile-2.0.0-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl"]
PAGES = {"index.html": "web/index.html", "viewer.js": "web/viewer.js", "viewer.css": "web/viewer.css",
         "style.css": "static/style.css", "web/worker.mjs": "web/worker.mjs"}


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

    for target, source in PAGES.items():
        copy(BASE / source, out / target)
    for name in PYODIDE_FILES:
        copy(args.pyodide / name, out / "pyodide" / name)
    for name in WHEELS:
        copy(args.wheels / name, out / "wheels" / name)

    files = sorted(f"engine/{p.name}" for p in (BASE / "engine").glob("*.py")) + ["web/workbook_view.py"]
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
