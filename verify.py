"""Reproducible behavioral checks on synthetic data; writes measured evidence."""
import contextlib
import copy
import datetime
import hashlib
import io
import json
import tempfile
import threading
import time
import urllib.error
import urllib.request

from core import BASE, DemoStore, sample, validate_item, json_write
from app import create_server

def main():
    results=[]
    def check(name, expected, actual):
        passed = actual == expected
        results.append({"name":name,"expected":expected,"actual":actual,"passed":passed})

    start = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix="macro_demo_verify_") as tmp:
        store=DemoStore(tmp)
        base=store.run()
        run_id=base["run"]["id"]
        rows={r["row"]:r for r in base["state"]["rows"]}
        check("정상 예시 8건 집계",8,base["news_count"])
        check("의견은 뉴스 참고 점수에서 제외",-1.0,rows[8]["news_dur"])
        check("정책 갭 없으면 제안한 -1.5를 버림",0.0,rows[10]["dur"])
        check("TP 수치 없으면 제안한 -1을 버림",0.0,rows[12]["dur"])
        check("결측 정책·TP를 별도로 표시",3,sum(r["missing_basis"] for r in rows.values()))
        check("채점 확정은 자동 실행으로 바뀌지 않음",False,base["state"]["confirmed"])
        store.review(run_id,1.5)
        after=store.reapply(run_id)
        check("사용자가 고친 G8은 재실행 후 보존",1.5,after["state"]["rows"][0]["dur"])
        check("변경 칸 잠금 기록",True,"dur" in after["state"]["rows"][0]["locked"])
        check("동일 입력 묶음 재실행의 중복 추가 방지",8,after["news_count"])
        check("검토와 재실행 이력 기록",True,{"human_review","draft_apply"}.issubset({r["stage"] for r in after["trace"]}))
        coerce=store.run("coerce")
        check("범위·단위 보정",[0.5,-2.0],[coerce["state"]["rows"][0][k] for k in ("dur","cur")])
        gap=store.run("policy_gap")
        check("정책 갭 +60bp 재계산",[2.0,-2.0],[gap["state"]["rows"][2][k] for k in ("dur","cur")])
        rejected=0
        for value in (True,float("nan"),float("inf"),3,0.7):
            try:
                store.review(run_id,value)
            except ValueError:
                rejected+=1
        check("잘못된 사용자 점수 5종 거절",5,rejected)
        invalid=copy.deepcopy(sample()["items"][0]); invalid["intensity"]=-1
        try:
            validate_item(invalid); rejected=False
        except ValueError:
            rejected=True
        check("방향성·강도 부호 불일치 거절",True,rejected)
        try:
            store.directory("../../CLAUDE.md"); rejected=False
        except ValueError:
            rejected=True
        check("허용되지 않은 파일 경로 접근 거절",True,rejected)
        check("DB 이력 재시작 후 유지",3,len(DemoStore(tmp).history()))
        server=create_server(0,tmp)
        thread=threading.Thread(target=server.serve_forever,daemon=True); thread.start()
        url=f"http://127.0.0.1:{server.server_port}"
        try:
            token=json.load(urllib.request.urlopen(url+"/api/config"))["token"]
            def post(path,payload,auth=True,origin=None):
                headers={"Content-Type":"application/json"}
                if auth: headers["X-Demo-Token"]=token
                if origin: headers["Origin"]=origin
                request=urllib.request.Request(url+path,json.dumps(payload).encode(),headers)
                try:
                    with urllib.request.urlopen(request) as response:
                        return response.status,json.load(response)
                except urllib.error.HTTPError as error:
                    return error.code,json.load(error)
            check("토큰 없는 변경 요청 차단",403,post("/api/reapply",{"id":run_id},False)[0])
            check("다른 Origin 변경 요청 차단",403,post("/api/reapply",{"id":run_id},origin="https://example.com")[0])
            check("정상 API 요청",200,post("/api/reapply",{"id":run_id})[0])
            check("API 점수 입력 검증",400,post("/api/review",{"id":run_id,"score":3})[0])
            with urllib.request.urlopen(url+"/download?id="+run_id) as response:
                check("워크북 다운로드는 유효 ZIP 서명",True,response.read(2)==b"PK")
        finally:
            server.shutdown(); server.server_close(); thread.join()
    hashes={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted((BASE/"engine").glob("*.py"))}
    output={"generated_at":datetime.datetime.now().isoformat(timespec="seconds"),
            "seconds":round(time.perf_counter()-start,3),"passed":sum(x["passed"] for x in results),
            "total":len(results),"cases":results,"engine_sha256":hashes,
            "scope":"가상 고정 입력의 결정론 규칙·파일 보존·로컬 API 검증. LLM 정확도·공격 전반·투자성과·시간절감은 측정하지 않음."}
    json_write(BASE/"evidence/results.json",output)
    lines=["# 데모 검증 결과", "", f"실행 시각: {output['generated_at']}",f"통과: {output['passed']}/{output['total']}","",output["scope"],"", "| 검증 | 기대 결과 | 실제 결과 | 판정 |","|---|---|---|---|"]
    for case in results:
        lines.append(f"| {case['name']} | {case['expected']} | {case['actual']} | {'PASS' if case['passed'] else 'FAIL'} |")
    (BASE/"evidence/검증결과.md").write_text("\n".join(lines)+"\n",encoding="utf-8")
    print(f"PASS {output['passed']}/{output['total']} ({output['seconds']} sec)")
    return 0 if output['passed']==output['total'] else 1

if __name__ == "__main__":
    raise SystemExit(main())
