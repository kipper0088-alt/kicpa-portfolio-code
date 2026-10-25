"""
[복원본] 유형자산 취득리스트(144건) - 감가상각대장(131건) 대사 · 1차 실습

당시 작성한 원본 코드는 남아 있지 않습니다. 실습 자료(README·입력 엑셀·검증정답)의
목적과 처리 기준을 바탕으로 다시 작성했고, 검증정답과 대조해 같은 결과를 확인했습니다.

- 자산번호를 기준으로 두 자료를 합친다.
- 감가상각대장에 없는 취득 건을 예외목록으로 추출한다.
- 결과를 검증정답과 대조한다.

사용법 (프로젝트 폴더에서):
    python src/recon_basic.py
    python src/recon_basic.py --data data/basic --answer answer/검증정답_1차.xlsx --out results/대사결과_1차.xlsx
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent.parent
KEY = "자산번호"


def load(folder: Path):
    acq = pd.read_excel(folder / "취득리스트_144건.xlsx", dtype={KEY: str})
    dep = pd.read_excel(folder / "감가상각대장_131건.xlsx", dtype={KEY: str})
    for df in (acq, dep):
        df[KEY] = df[KEY].str.strip()
    return acq, dep


def reconcile(acq: pd.DataFrame, dep: pd.DataFrame) -> pd.DataFrame:
    merged = acq.merge(dep[[KEY]], on=KEY, how="outer", indicator=True)
    missing = merged[merged["_merge"] == "left_only"].drop(columns="_merge")
    missing = missing.copy()
    missing["판정"] = "감가상각대장 미등재"
    # 감가상각대장에만 있는 자산도 함께 확인한다 (1차 자료에서는 0건이어야 정상).
    only_dep = merged[merged["_merge"] == "right_only"]
    if len(only_dep):
        print(f"[참고] 감가상각대장에만 있는 자산 {len(only_dep)}건")
    return missing.sort_values(KEY).reset_index(drop=True)


def verify(result: pd.DataFrame, answer_path: Path) -> bool:
    ans = pd.read_excel(answer_path, dtype={KEY: str})
    got, exp = set(result[KEY]), set(ans[KEY].str.strip())
    ok = got == exp
    print(f"검증: 결과 {len(got)}건 / 정답 {len(exp)}건 / 일치 {len(got & exp)}건 -> {'PASS' if ok else 'FAIL'}")
    if not ok:
        print("  정답에만 있음:", sorted(exp - got))
        print("  결과에만 있음:", sorted(got - exp))
    return ok


def main():
    ap = argparse.ArgumentParser(description="유형자산 대사 1차 (복원본)")
    ap.add_argument("--data", type=Path, default=HERE / "data" / "basic")
    ap.add_argument("--answer", type=Path, default=HERE / "answer" / "검증정답_1차.xlsx")
    ap.add_argument("--out", type=Path, default=HERE / "results" / "대사결과_1차.xlsx")
    a = ap.parse_args()
    acq, dep = load(a.data)
    print(f"취득리스트 {len(acq)}건, 감가상각대장 {len(dep)}건, 차이 {len(acq) - len(dep)}건")
    result = reconcile(acq, dep)
    print(result[[KEY, "자산명/적요", "취득금액"]].to_string(index=False))
    a.out.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(a.out, index=False)
    ok = verify(result, a.answer)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
