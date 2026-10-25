"""
VBA 매크로 결과(통합결과·검증요약 시트)를 검증기준 파일과 대조한다.
원자료에서 pandas로 독립적으로 다시 추출한 행과도 전표 단위로 비교한다.

사용법 (프로젝트 폴더에서):
    python verify/verify_vba_result.py                       # results/VBA_취합결과.xlsx 검증
    python verify/verify_vba_result.py <VBA 결과 xlsx> <A자료 폴더> <검증기준 xlsx>
"""
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent.parent
if len(sys.argv) >= 4:
    result_path, src_dir, criteria_path = map(Path, sys.argv[1:4])
else:
    result_path = HERE / "results" / "VBA_취합결과.xlsx"
    src_dir, criteria_path = HERE / "data", HERE / "answer" / "검증기준.xlsx"
CODE = "111300"

out = pd.read_excel(result_path, sheet_name="통합결과", dtype={"계정코드": str})
summ = pd.read_excel(result_path, sheet_name="검증요약")
crit = pd.read_excel(criteria_path)

ok = True
print(f"통합결과 {len(out)}행, 계정코드 종류: {sorted(out['계정코드'].str.strip().unique())}")
ok &= set(out["계정코드"].str.strip()) == {CODE}

# 1) 검증기준(행 수·차변·대변 합계)과 VBA 검증요약 비교
s = summ[summ["파일명"] != "합계"].set_index("파일명")
for _, c in crit.iterrows():
    f = c["파일명"]
    got = (int(s.loc[f, "추출 행 수"]), int(s.loc[f, "차변 합계"]), int(s.loc[f, "대변 합계"]))
    exp = (int(c["정기예금 행 수"]), int(c["정기예금 차변합계"]), int(c["정기예금 대변합계"]))
    flag = "PASS" if got == exp else "FAIL"
    ok &= got == exp
    print(f"  {f:<28} 행 {got[0]:>4}/{exp[0]:<4} 차변 {got[1]:>13,} 대변 {got[2]:>12,}  {flag}")

# 2) 통합결과 시트 자체에서 다시 합산해도 같은지 확인
g = out.groupby("원본파일").agg(n=("전표번호", "size"), dr=("차변", "sum"), cr=("대변", "sum"))
for _, c in crit.iterrows():
    r = g.loc[c["파일명"]]
    ok &= (r.n, r.dr, r.cr) == (c["정기예금 행 수"], c["정기예금 차변합계"], c["정기예금 대변합계"])

# 3) 원자료에서 독립 추출한 행과 전표번호 단위 비교
ref = pd.concat(pd.read_excel(p, dtype={"계정코드": str}).assign(원본파일=p.name)
                for p in sorted(src_dir.glob("A_*.xlsx")))
ref = ref[ref["계정코드"].str.strip() == CODE]
k_ref = set(zip(ref["원본파일"], ref["전표번호"].astype(int), ref["차변"], ref["대변"]))
k_out = set(zip(out["원본파일"], out["전표번호"].astype(int), out["차변"], out["대변"]))
same = k_ref == k_out
ok &= same
print(f"전표 단위 대조: 원자료 재추출 {len(k_ref)}건 / VBA {len(k_out)}건 -> {'PASS' if same else 'FAIL'}")
print("최종:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
