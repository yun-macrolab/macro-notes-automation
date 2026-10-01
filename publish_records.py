#!/usr/bin/env python3
"""공개 기록 자동 올리기 — PC의 주간 기록 워크북 폴더를 훑어 공개본을 만들고, 바뀐 주만 저장소 records/에 올린다.

브라우저의 '공개본 만들기 → 여러 주 한 번에'와 같은 코드·같은 검사를 쓴다(web/workbook_view.public_from_workbook).
  - 검사에 걸린 주, 예전 형식(v1), 주간 날짜(I1)가 없는 파일은 건너뛴다. 같은 주 파일이 여럿이면 채점 확정 → 최근 저장 순.
  - 이미 올라간 주와 내용이 같으면(만든 시각만 다르면) 다시 올리지 않는다. 저장소의 공개본을 지우지는 않는다.
  - 바뀐 주는 GitHub API로 커밋 하나에 묶어 올린다. 그러면 배포(pages.yml)가 한 번 더 검사한 뒤 사이트에 싣는다.

  python -X utf8 publish_records.py             처음엔 폴더·토큰을 묻고 기억한다. 바뀔 주를 보여 주고 올릴지 묻는다
  python -X utf8 publish_records.py --dry-run   무엇이 바뀌는지만 본다(토큰 없어도 된다)
  python -X utf8 publish_records.py --yes       묻지 않고 올린다(예약 작업용 — Publish-Records.bat -Schedule)
  python -X utf8 publish_records.py --out 폴더  올리지 않고 공개본 파일만 그 폴더에 쓴다

설정은 이 PC 사용자 폴더의 .macro-notes/에 둔다(저장소에 들어가지 않는다).
  publish.json      기록 폴더·저장소
  token.txt         GitHub 토큰(환경 변수 MACRO_NOTES_TOKEN이 있으면 그것을 쓴다). 이 저장소 Contents 읽기·쓰기 권한이면 된다
  check-words.txt   검사어(선택, 한 줄에 하나)
  publish.log       실행 기록(토큰과 걸린 말은 남기지 않는다)
"""
import argparse
import base64
import datetime as dt
import getpass
import json
import os
from pathlib import Path
import sys
import urllib.error
import urllib.request

BASE = Path(__file__).resolve().parent
sys.path[:0] = [str(BASE / "engine"), str(BASE / "web")]

import workbook_view  # noqa: E402

DEFAULT_REPO = "yun-macrolab/macro-notes-automation"
BRANCH = "main"
API = "https://api.github.com"
HOME = Path.home() / ".macro-notes"
URLOPEN = urllib.request.urlopen       # 검사에서 가짜 GitHub로 바꾼다
REASONS = {"v1": "예전 형식(v1)", "no_week": "주간 날짜 없음(I1 칸)", "not_record": "기록 워크북 아님",
           "broken": "읽지 못함", "blocked": "검사에 걸림"}


class PublishError(Exception):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


# ---------- 워크북 → 공개본 ----------
def workbooks(folder):
    """폴더(하위 폴더 포함)의 엑셀 파일. 엑셀이 열어 둔 동안 생기는 '~$' 잠금 파일은 뺀다."""
    return sorted(p for p in folder.rglob("*")
                  if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))


def collect(folder, words=()):
    """워크북마다 공개본을 만든다 → (파일별 결과, {"YYYY-MM-DD.json": 쓸 결과})."""
    rows = []
    for path in workbooks(folder):
        ok, text = workbook_view.public_from_workbook(str(path), json.dumps(list(words), ensure_ascii=False))
        body = json.loads(text)
        row = {"label": str(path.relative_to(folder)), "ok": ok, "mtime": path.stat().st_mtime}
        if ok:
            row.update(text=text, record=body)
        else:
            row.update(reason=body.get("reason", "broken"), findings=body.get("findings", []), week=body.get("week"))
        rows.append(row)
    picked = {}
    for row in rows:
        if row["ok"]:
            name = row["record"]["week"]["start"] + ".json"
            if name not in picked or preferred(row, picked[name]):
                picked[name] = row
    return rows, picked


def preferred(a, b):
    """같은 주 파일이 여럿이면 채점 확정한 파일, 그다음 최근에 저장한 파일."""
    confirmed = lambda row: row["record"]["state"]["confirmed"] is True  # noqa: E731
    if confirmed(a) != confirmed(b):
        return confirmed(a)
    return a["mtime"] > b["mtime"]


def same(old_text, new_text):
    """만든 시각(generated_at)만 다르면 같은 공개본으로 본다."""
    try:
        old = json.loads(old_text)
    except ValueError:
        return False
    new = json.loads(new_text)
    old.pop("generated_at", None)
    new.pop("generated_at", None)
    return old == new


# ---------- GitHub ----------
class GitHub:
    """저장소의 records/를 읽고, 바뀐 공개본을 커밋 하나로 올린다(Git Data API)."""

    def __init__(self, repo, token=None, branch=BRANCH):
        self.repo, self.token, self.branch = repo, token, branch

    def call(self, method, path, body=None):
        request = urllib.request.Request(f"{API}/repos/{self.repo}{path}", method=method,
                                         data=None if body is None else json.dumps(body).encode("utf-8"))
        request.add_header("Accept", "application/vnd.github+json")
        request.add_header("X-GitHub-Api-Version", "2022-11-28")
        request.add_header("User-Agent", "macro-notes-publish")
        if body is not None:
            request.add_header("Content-Type", "application/json")
        if self.token:
            request.add_header("Authorization", f"Bearer {self.token}")
        try:
            with URLOPEN(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8") or "null")
        except urllib.error.HTTPError as error:
            try:
                detail = json.loads(error.read().decode("utf-8")).get("message", "")
            except (ValueError, OSError, AttributeError):
                detail = ""
            hint = " — 토큰에 이 저장소의 Contents 읽기·쓰기 권한이 있는지 확인하세요" if error.code in (401, 403, 404) else ""
            raise PublishError(f"GitHub {method} {path.split('?')[0]}: HTTP {error.code} {detail}{hint}", error.code) from None
        except urllib.error.URLError as error:
            raise PublishError(f"GitHub에 연결하지 못했습니다: {error.reason}") from None

    def snapshot(self):
        """지금 main의 커밋·트리와 records/의 공개본 {이름: blob sha}."""
        head = self.call("GET", f"/git/ref/heads/{self.branch}")["object"]["sha"]
        tree = self.call("GET", f"/git/commits/{head}")["tree"]["sha"]
        listing = self.call("GET", f"/git/trees/{tree}?recursive=1")["tree"]
        blobs = {entry["path"][len("records/"):]: entry["sha"] for entry in listing
                 if entry["type"] == "blob" and entry["path"].startswith("records/") and entry["path"].endswith(".json")
                 and "/" not in entry["path"][len("records/"):]}
        return head, tree, blobs

    def blob_text(self, sha):
        return base64.b64decode(self.call("GET", f"/git/blobs/{sha}")["content"]).decode("utf-8")

    def commit(self, head, tree, files, message):
        """files {이름: 공개본 글}를 records/에 쓰는 커밋 하나를 만들고 main을 그 커밋으로 옮긴다(빨리 감기만)."""
        new_tree = self.call("POST", "/git/trees", {"base_tree": tree, "tree": [
            {"path": "records/" + name, "mode": "100644", "type": "blob", "content": text}
            for name, text in sorted(files.items())]})["sha"]
        commit = self.call("POST", "/git/commits", {"message": message, "tree": new_tree, "parents": [head]})["sha"]
        self.call("PATCH", f"/git/refs/heads/{self.branch}", {"sha": commit, "force": False})
        return commit


def plan(github, picked):
    """올라간 공개본과 비교 → (head, tree, {이름: 상태}, {이름: 올릴 글})."""
    head, tree, blobs = github.snapshot()
    states, upload = {}, {}
    for name, row in sorted(picked.items()):
        if name not in blobs:
            states[name], upload[name] = "새로 올림", row["text"]
        elif same(github.blob_text(blobs[name]), row["text"]):
            states[name] = "그대로(이미 올라가 있음)"
        else:
            states[name], upload[name] = "바뀜 → 다시 올림", row["text"]
    return head, tree, states, upload


def commit_message(states, upload):
    new = [name[:-5] for name in upload if states[name] == "새로 올림"]
    changed = [name[:-5] for name in upload if states[name] != "새로 올림"]
    lines = [f"records: 공개 기록 {len(upload)}주 올림", ""]
    if new:
        lines.append("새로: " + ", ".join(new))
    if changed:
        lines.append("바뀜: " + ", ".join(changed))
    lines += ["", "publish_records.py로 올림(PC의 주간 기록 워크북 → 공개본, 같은 검사)."]
    return "\n".join(lines)


# ---------- 설정·출력 ----------
def read_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def read_words():
    words = os.environ.get("PUBLIC_CHECK_WORDS", "").splitlines()
    try:
        words += (HOME / "check-words.txt").read_text(encoding="utf-8-sig").splitlines()
    except OSError:
        pass
    return [word.strip() for word in words if word.strip()]


def read_token():
    token = os.environ.get("MACRO_NOTES_TOKEN", "").strip()
    if not token:
        try:
            token = (HOME / "token.txt").read_text(encoding="utf-8-sig").strip()
        except OSError:
            token = ""
    return token or None


def save_token(token):
    HOME.mkdir(parents=True, exist_ok=True)
    path = HOME / "token.txt"
    path.write_text(token + "\n", encoding="utf-8")
    try:
        path.chmod(0o600)
    except OSError:
        pass
    return path


def log(line):
    try:
        HOME.mkdir(parents=True, exist_ok=True)
        with open(HOME / "publish.log", "a", encoding="utf-8") as handle:
            handle.write(f"{dt.datetime.now().isoformat(timespec='seconds')} {line}\n")
    except OSError:
        pass


def score(value):
    return "—" if not isinstance(value, (int, float)) else f"{value:+.2f}"


def describe(row, picked, states):
    record = row.get("record")
    week = record["week"] if record else row.get("week")
    when = f"{week['start']} ~ {week['end'][5:]}" if week else "주간 —"
    if not row["ok"]:
        return f"  {when}  {row['label']}  → 건너뜀 · {REASONS.get(row['reason'], '읽지 못함')}"
    name = record["week"]["start"] + ".json"
    state = record["state"]
    confirmed = {True: "채점 확정", False: "확정 전"}.get(state["confirmed"], "확정 표시 없음")
    result = states.get(name, "올릴 공개본") if picked.get(name) is row else "같은 주의 다른 파일을 씀"
    return (f"  {when}  {row['label']}  뉴스 {len(record['news'])}건 · 듀 {score(state['dur']['total'])} "
            f"{state['dur']['verdict']} · 커브 {score(state['cur']['total'])} {state['cur']['verdict']} · {confirmed}"
            f"  → {result}")


def report(rows, picked, states):
    """만든 주(최신 주가 위) → 건너뛴 파일 순. 검사에 걸린 파일은 걸린 위치도 보여 준다(이 PC 화면에만)."""
    week_of = lambda row: row["record"]["week"]["start"] if row["ok"] else ""  # noqa: E731
    for row in sorted(sorted(rows, key=lambda row: row["label"]), key=lambda row: (row["ok"], week_of(row)), reverse=True):
        print(describe(row, picked, states))
        if not row["ok"] and row["reason"] == "blocked":
            for finding in row["findings"][:10]:
                print(f"      · {finding['place']} — {finding['kind']} ‘{finding['text']}’")


# ---------- 실행 ----------
def main(argv=None):
    parser = argparse.ArgumentParser(description="PC의 주간 기록 워크북 → 공개본 → 저장소 records/(바뀐 주만)")
    parser.add_argument("--folder", type=Path, help="주간 기록 워크북 폴더(한 번 주면 기억한다)")
    parser.add_argument("--repo", help=f"올릴 저장소(기본 {DEFAULT_REPO})")
    parser.add_argument("--dry-run", action="store_true", help="무엇이 바뀌는지만 보기")
    parser.add_argument("--yes", action="store_true", help="묻지 않고 올리기(예약 작업용)")
    parser.add_argument("--out", type=Path, help="올리지 않고 공개본 파일만 이 폴더에 쓰기")
    args = parser.parse_args(argv)
    interactive = not args.yes and sys.stdin.isatty()
    try:
        return run(args, interactive)
    except PublishError as error:
        print(f"올리지 못했습니다: {error}", file=sys.stderr)
        log(f"실패: {error}")
        return 1


def run(args, interactive):
    settings = read_json(HOME / "publish.json")
    folder = args.folder or (Path(settings["folder"]) if settings.get("folder") else None)
    if folder is None:
        if not interactive:
            raise PublishError("기록 폴더를 모릅니다. 한 번은 직접 실행해(또는 --folder) 폴더를 알려 주세요.")
        folder = Path(input("주간 기록 워크북이 든 폴더 경로를 붙여 넣으세요: ").strip().strip('"'))
    folder = folder.expanduser().absolute()      # 예약 작업이 다른 위치에서 돌아도 같은 폴더를 보게
    if not folder.is_dir():
        raise PublishError(f"폴더가 없습니다: {folder}")
    repo = args.repo or settings.get("repo") or DEFAULT_REPO
    if settings.get("folder") != str(folder) or settings.get("repo") != repo:
        HOME.mkdir(parents=True, exist_ok=True)
        (HOME / "publish.json").write_text(json.dumps({"folder": str(folder), "repo": repo}, ensure_ascii=False, indent=2),
                                           encoding="utf-8")

    print(f"기록 폴더: {folder}")
    rows, picked = collect(folder, read_words())
    if not rows:
        print("엑셀 기록 파일(.xlsx·.xlsm)이 없습니다.")
        return 0
    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
        for name, row in picked.items():
            (args.out / name).write_text(row["text"], encoding="utf-8")
        report(rows, picked, {name: f"{args.out / name}에 씀" for name in picked})
        return 0

    github = GitHub(repo, read_token())
    head, tree, states, upload = plan(github, picked)
    report(rows, picked, states)
    skipped = sum(1 for row in rows if not row["ok"])
    print(f"\n공개본 {len(picked)}주 · 올릴 것 {len(upload)}주 · 건너뛴 파일 {skipped}개")
    if not upload:
        print("바뀐 주가 없어 올릴 것이 없습니다.")
        log(f"바뀐 주 없음 · 공개본 {len(picked)}주 · 건너뜀 {skipped}")
        return 0
    if args.dry_run:
        print("(미리 보기) 올리지 않았습니다.")
        return 0
    if not github.token:
        if not interactive:
            raise PublishError("GitHub 토큰이 없습니다. 한 번은 직접 실행해 토큰을 넣어 주세요.")
        print("\nGitHub 토큰이 필요합니다(이 저장소 Contents 읽기·쓰기 권한의 fine-grained 토큰). README의 '자동으로 올리기' 참고.")
        github.token = getpass.getpass("토큰을 붙여 넣고 Enter(화면에 보이지 않습니다): ").strip()
        if not github.token:
            raise PublishError("토큰을 받지 못했습니다.")
        print(f"토큰을 이 PC에 저장했습니다: {save_token(github.token)} (예약 작업이 씁니다. 지우면 다시 묻습니다)")
    if interactive and input(f"{len(upload)}주를 올릴까요? [y/N] ").strip().lower() not in ("y", "yes", "ㅛ"):
        print("올리지 않았습니다.")
        return 0
    try:
        sha = github.commit(head, tree, upload, commit_message(states, upload))
    except PublishError as error:
        if error.status != 422:          # 그사이 main이 바뀌었으면 한 번만 다시 비교해 올린다
            raise
        head, tree, states, upload = plan(github, picked)
        if not upload:
            print("그사이 같은 내용이 올라가 있어 올릴 것이 없습니다.")
            return 0
        sha = github.commit(head, tree, upload, commit_message(states, upload))
    print(f"올렸습니다: 커밋 {sha[:7]} · 1~2분 뒤 배포가 끝나면 사이트 '공개 기록' 탭에 보입니다.")
    log(f"올림 {len(upload)}주({', '.join(name[:-5] for name in sorted(upload))}) · 커밋 {sha[:7]} · 건너뜀 {skipped}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
