"""정적 사이트 조립 — GitHub Pages에 올릴 '매크로 기록 보기' 폴더를 만든다(표준 라이브러리만 사용).

사이트에 싣는 기록은 records/의 공개본(JSON)뿐이다. 정해 둔 칸·형식인지, 개인정보로 보이는 말이 없는지
다시 검사하고(web/public_record.validate) 하나라도 걸리면 조립을 멈춘다(배포하지 않는다).
방문자가 고른 자기 PC의 기록 워크북은 브라우저 안에서만 읽는다.
화면 디자인은 시연 화면(static/style.css)을 그대로 쓰고, engine/은 바이트 그대로 싣는다.

  python -X utf8 web/build.py --pyodide <pyodide npm 패키지의 package 폴더> --wheels <휠 폴더> --out _site
  검사어: 환경 변수 PUBLIC_CHECK_WORDS(한 줄에 하나). 배포에서는 저장소 비밀값으로 넣는다.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys

import public_record

BASE = Path(__file__).resolve().parent.parent
PYODIDE_FILES = ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"]
WHEELS = ["et_xmlfile-2.0.0-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl"]
PAGES = {"index.html": "web/index.html", "viewer.js": "web/viewer.js", "viewer.css": "web/viewer.css",
         "style.css": "static/style.css", "web/worker.mjs": "web/worker.mjs"}
INDEX_SCHEMA = "macro-notes-public-index/1"
RECORD_NAME = re.compile(r"\d{4}-\d{2}-\d{2}\.json")
MAX_RECORD_BYTES = 2_000_000


def copy(src, dst):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)


def load_records(source, words=()):
    """records/의 공개본을 읽어 검사한다 → [(파일 이름, 공개본)]. 문제가 있으면 ValueError.
    오류 문구는 공개 로그에 남으므로 파일 이름과 위치만 싣는다(걸린 말은 싣지 않는다)."""
    records = []
    for path in sorted(source.iterdir()) if source.is_dir() else []:
        where = f"records/{path.name}"
        if path.name == "README.md":
            continue
        if path.suffix.lower() != ".json" or not path.is_file():
            raise ValueError(f"{where}: records/에는 공개본(.json)과 README.md만 둡니다. "
                             "원본 엑셀 등을 잘못 올렸다면 저장소에서 바로 지우세요(공개 저장소입니다)")
        if not RECORD_NAME.fullmatch(path.name):
            raise ValueError(f"{where}: 파일 이름은 그 주 월요일 날짜여야 합니다(예: 2026-09-21.json)")
        if path.stat().st_size > MAX_RECORD_BYTES:
            raise ValueError(f"{where}: 파일이 너무 큽니다(2MB 넘음)")
        try:
            record = json.loads(path.read_text(encoding="utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError(f"{where}: JSON 파일로 읽지 못했습니다") from None
        try:
            public_record.validate(record, words)
        except ValueError as error:
            raise ValueError(f"{where}: {error}") from None
        if path.name != record["week"]["start"] + ".json":
            raise ValueError(f"{where}: 파일 이름이 기록의 주간 시작일({record['week']['start']})과 다릅니다")
        records.append((path.name, record))
    return records


def write_records(target, records):
    """검사를 통과한 공개본과 목록(index.json, 최신 주가 먼저)을 사이트 폴더에 쓴다 → 목록."""
    target.mkdir(parents=True, exist_ok=True)
    entries = []
    for name, record in records:
        (target / name).write_text(json.dumps(record, ensure_ascii=False, allow_nan=False), encoding="utf-8")
        state = record["state"]
        entries.append({"file": name, "start": record["week"]["start"], "end": record["week"]["end"],
                        "generated_at": record["generated_at"], "confirmed": state["confirmed"],
                        "dur": state["dur"], "cur": state["cur"], "news": len(record["news"])})
    entries.sort(key=lambda entry: entry["start"], reverse=True)
    index = {"schema": INDEX_SCHEMA, "records": entries}
    (target / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    return entries


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pyodide", required=True, type=Path)
    parser.add_argument("--wheels", required=True, type=Path)
    parser.add_argument("--records", default=BASE / "records", type=Path)
    parser.add_argument("--out", default=BASE / "_site", type=Path)
    args = parser.parse_args()
    out = args.out.resolve()
    if out == BASE or BASE in out.parents and out.name != "_site":
        sys.exit("--out은 저장소 밖 폴더이거나 저장소 안의 _site여야 합니다.")
    # 공개본 검사를 먼저 한다. 하나라도 걸리면 아무것도 만들지 않는다.
    try:
        records = load_records(args.records, os.environ.get("PUBLIC_CHECK_WORDS", "").splitlines())
    except ValueError as error:
        sys.exit(f"공개 기록 검사 실패 — {error}")
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

    files = sorted(f"engine/{p.name}" for p in (BASE / "engine").glob("*.py")) + [
        "web/workbook_view.py", "web/public_record.py"]
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
    entries = write_records(out / "records", records)
    (out / ".nojekyll").write_text("")
    size = sum(p.stat().st_size for p in out.rglob("*") if p.is_file())
    weeks = ", ".join(entry["start"] for entry in entries) or "없음"
    print(f"정적 사이트 조립: {out} · 파일 {sum(1 for p in out.rglob('*') if p.is_file())}개 · {size / 1e6:.1f}MB"
          f" · 공개 기록 {len(entries)}주({weeks})")


if __name__ == "__main__":
    main()
