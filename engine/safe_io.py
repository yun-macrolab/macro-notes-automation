#!/usr/bin/env python3
"""안전한 워크북 저장 유틸리티 (손상 방지).

문제: 동기화 폴더(OneDrive 등)에 openpyxl로 직접 wb.save()를 하면,
저장이 수 초에 걸쳐 진행되는 동안 동기화 클라이언트가 '반쯤 쓰인' 파일을
잡아가면서 파일이 잘려 손상될 수 있다(중앙 디렉터리/sharedStrings 유실).

해결:
  1) backup_valid(): 저장 직전, 기존 정상 파일을 _recovery에 스냅샷(최근 N개 유지)
  2) safe_save(): 같은 폴더의 임시파일에 저장 → 무결성 검증 → os.replace로 원자적 교체
     - 검증 실패 시 예외를 던지고 원본은 건드리지 않는다(손상 전파 차단).
     - 동기화 클라이언트는 '완성된 파일'만 보게 된다(부분쓰기 노출 X).
"""
import os
import sys
import glob
import shutil
import zipfile
import datetime
import openpyxl


def is_valid_xlsx(path, expect_sheets=None):
    """zip 무결성 + openpyxl 로드 + (선택)시트 존재까지 확인."""
    try:
        if not zipfile.is_zipfile(path):
            return False, "유효한 zip 아님(중앙 디렉터리 유실 가능)"
        wb = openpyxl.load_workbook(path, read_only=True)
        names = set(wb.sheetnames)
        wb.close()
        if expect_sheets:
            missing = [s for s in expect_sheets if s not in names]
            if missing:
                return False, f"시트 누락: {missing}"
        return True, "ok"
    except Exception as e:
        return False, f"로드 실패: {e!r}"


def backup_valid(path, keep=10):
    """기존 파일이 '정상'일 때만 _recovery 폴더에 타임스탬프 스냅샷.
    손상본은 백업하지 않는다(정상 복구점만 보관). 최근 keep개만 유지."""
    if not os.path.exists(path):
        return None
    ok, _ = is_valid_xlsx(path)
    if not ok:
        return None
    folder = os.path.dirname(os.path.abspath(path))
    rec = os.path.join(folder, "_recovery")
    os.makedirs(rec, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(rec, f"{base}.{ts}.bak.xlsx")
    shutil.copy2(path, dst)
    # 오래된 스냅샷 정리
    snaps = sorted(glob.glob(os.path.join(rec, f"{base}.*.bak.xlsx")),
                   key=os.path.getmtime)
    for old in snaps[:-keep]:
        try:
            os.remove(old)
        except OSError:
            pass
    return dst


def safe_save(wb, path, expect_sheets=None, do_backup=True, keep=10):
    """임시파일 저장 → 검증 → 원자적 교체. 실패 시 원본 보존 + 예외."""
    if do_backup:
        backup_valid(path, keep=keep)
    folder = os.path.dirname(os.path.abspath(path))
    os.makedirs(folder, exist_ok=True)
    base = os.path.basename(path)
    # 임시파일은 반드시 .xlsx 확장자로 끝나야 함(openpyxl이 확장자로 형식 검사).
    # 점(.) 접두로 숨김 처리 + PID로 충돌 방지.
    tmp = os.path.join(folder, f".~{base}.tmp{os.getpid()}.xlsx")
    try:
        wb.save(tmp)
        ok, msg = is_valid_xlsx(tmp, expect_sheets=expect_sheets)
        if not ok:
            raise RuntimeError(f"저장본 무결성 검증 실패: {msg}")
        try:
            os.replace(tmp, path)   # 같은 FS 내 원자적 교체
        except OSError:
            # 일부 동기화 마운트(FUSE 등)는 기존 파일 위에 rename을 허용하지 않아
            # PermissionError/OSError가 날 수 있다. 이 경우 복사+제거로 대체한다.
            # (tmp는 이미 위에서 무결성 검증을 통과했으므로 안전)
            shutil.copyfile(tmp, path)
            ok2, msg2 = is_valid_xlsx(path, expect_sheets=expect_sheets)
            if not ok2:
                raise RuntimeError(f"교체 후 검증 실패: {msg2}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return path


if __name__ == "__main__":
    # CLI: 파일 무결성 점검   python safe_io.py check <file> [Sheet1 Sheet2 ...]
    if len(sys.argv) >= 3 and sys.argv[1] == "check":
        ok, msg = is_valid_xlsx(sys.argv[2], expect_sheets=sys.argv[3:] or None)
        print(("OK   " if ok else "FAIL ") + sys.argv[2] + " :: " + msg)
        sys.exit(0 if ok else 1)
    print(__doc__)
