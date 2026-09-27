"""학습자의 미완성 함수 검사용. 제품 검증 verify.py와 별개입니다."""
from exercise01 import is_consistent

cases=[("금리상방",1,True),("금리상방",-1,False),("금리상방",0,False),
       ("금리하방",-0.5,True),("금리하방",1,False),("금리하방",0,False),
       ("중립",0,True),("중립",0.5,False),("모름",1,False)]
if __name__ == "__main__":
    passed=0
    for direction,value,expected in cases:
        try:
            actual=is_consistent(direction,value)
            ok=actual is expected
            print(f"{'PASS' if ok else 'FAIL'} {direction}, {value}: 기대={expected}, 실제={actual}")
            passed+=ok
        except NotImplementedError:
            print("아직 작성 전입니다. exercise01.py의 함수를 먼저 채우세요.")
            break
    print(f"학습 과제: {passed}/{len(cases)}")
    raise SystemExit(0 if passed==len(cases) else 1)
