"""
[복원본] 8개 사업부 취득리스트 - 감가상각대장 대사 · 확장 실습 (2026.09)

당시 작성한 원본 코드(asset_reconcile_batch.py)는 남아 있지 않습니다. 실습 자료의
목적과 처리 기준을 바탕으로 다시 작성했고, 검증정답과 대조해 같은 결과를 확인했습니다.

- 사업부별 파일을 한 번에 읽어 합친다.
- 사업부코드 + 자산번호를 키로 대사한다.
- 존재 여부와 함께 취득금액, 취득일, 내용연수를 비교한다.
- 예외만 하나의 결과 파일(예외목록 + 사업부별요약)로 정리한다.
- 결과를 검증정답과 대조한다.

사용법 (프로젝트 폴더에서):
    python src/recon_multi.py
    python src/recon_multi.py --data data/multi --answer answer/검증정답_8개사업부.xlsx --out results/대사결과_8개사업부.xlsx
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent.parent
KEYS = ["사업부코드", "자산번호"]
COMPARE = [  # (비교 컬럼, 예외유형)
    ("취득금액", "취득금액 불일치"),
    ("취득일", "취득일 불일치"),
    ("내용연수", "내용연수 불일치"),
]
SUMMARY_COLS = {
    "감가상각대장 미등재": "미등재",
    "취득금액 불일치": "금액불일치",
    "취득일 불일치": "날짜불일치",
    "내용연수 불일치": "내용연수불일치",
    "취득리스트 미등재": "대장에만존재",
}


def read_folder(folder: Path) -> pd.DataFrame:
    files = sorted(folder.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError(f"{folder} 에 xlsx 파일이 없습니다")
    frames = []
    for f in files:
        df = pd.read_excel(f, dtype={"사업부코드": str, "자산번호": str})
        df["원본파일"] = f.name
        frames.append(df)
    out = pd.concat(frames, ignore_index=True)
    for k in KEYS:
        out[k] = out[k].str.strip()
    out["취득일"] = pd.to_datetime(out["취득일"]).dt.normalize()
    return out


def reconcile(acq: pd.DataFrame, dep: pd.DataFrame) -> pd.DataFrame:
    for name, df in (("취득리스트", acq), ("감가상각대장", dep)):
        dup = df.duplicated(KEYS, keep=False)
        if dup.any():
            raise ValueError(f"{name}에 중복 키 {dup.sum()}건이 있습니다")

    cols = ["취득금액", "취득일", "내용연수"]
    m = acq[KEYS + cols].merge(
        dep[KEYS + cols], on=KEYS, how="outer", suffixes=("_취득", "_대장"), indicator=True
    )
    rows = []
    for _, r in m.iterrows():
        if r["_merge"] == "left_only":
            types = ["감가상각대장 미등재"]
        elif r["_merge"] == "right_only":
            types = ["취득리스트 미등재"]
        else:
            types = [t for c, t in COMPARE if r[f"{c}_취득"] != r[f"{c}_대장"]]
        for t in types:  # 한 자산에 예외가 여러 개면 각각 한 줄로 남긴다
            rows.append({
                "사업부코드": r["사업부코드"],
                "자산번호": r["자산번호"],
                "예외유형": t,
                "취득리스트_금액": r["취득금액_취득"],
                "대장_금액": r["취득금액_대장"],
                "취득리스트_취득일": r["취득일_취득"],
                "대장_취득일": r["취득일_대장"],
                "취득리스트_내용연수": r["내용연수_취득"],
                "대장_내용연수": r["내용연수_대장"],
            })
    return pd.DataFrame(rows).sort_values(KEYS + ["예외유형"]).reset_index(drop=True)


def summarize(acq, dep, exc) -> pd.DataFrame:
    s = acq.groupby("사업부코드").agg(사업부명=("사업부명", "first"), 취득리스트건수=("자산번호", "size"))
    s["감가상각대장건수"] = dep.groupby("사업부코드").size()
    counts = exc.pivot_table(index="사업부코드", columns="예외유형", values="자산번호",
                             aggfunc="count", fill_value=0)
    for t, col in SUMMARY_COLS.items():
        s[col] = counts[t] if t in counts else 0
    return s.reset_index()


def verify(exc: pd.DataFrame, summary: pd.DataFrame, answer_path: Path) -> bool:
    ans = pd.read_excel(answer_path, sheet_name=None, dtype={"사업부코드": str, "자산번호": str})
    exp = set(map(tuple, ans["예외정답"][KEYS + ["예외유형"]].values))
    got = set(map(tuple, exc[KEYS + ["예외유형"]].values))
    ok_rows = got == exp
    print(f"예외목록: 결과 {len(got)}건 / 정답 {len(exp)}건 / 일치 {len(got & exp)}건 -> {'PASS' if ok_rows else 'FAIL'}")
    if not ok_rows:
        print("  정답에만:", sorted(exp - got)[:20])
        print("  결과에만:", sorted(got - exp)[:20])

    a = ans["사업부별요약"].set_index("사업부코드").sort_index()
    b = summary.set_index("사업부코드").sort_index()[a.columns]
    ok_sum = a.astype(str).equals(b.astype(str))
    print(f"사업부별요약 {len(a)}개 사업부 × {len(a.columns)}개 항목 -> {'PASS' if ok_sum else 'FAIL'}")
    if not ok_sum:
        print(a.compare(b))

    # 금액·날짜·내용연수 값까지 정답과 같은지 확인
    key = KEYS + ["예외유형"]
    av = ans["예외정답"].set_index(key).sort_index()
    bv = exc.set_index(key).sort_index()[av.columns]
    diff = blank_filled = 0
    for c in av.columns:
        conv = pd.to_datetime if "일" in c else pd.to_numeric
        x, y = conv(av[c]), conv(bv[c])
        diff += int((x.notna() & (x != y)).sum())          # 정답에 값이 있는데 다름
        blank_filled += int((x.isna() & y.notna()).sum())  # 정답은 빈칸, 결과는 값 있음
    ok_val = diff == 0
    print(f"예외별 비교값(금액·취득일·내용연수) -> {'PASS' if ok_val else 'FAIL'} (불일치 {diff}칸)")
    if blank_filled:
        print(f"  [참고] 정답 파일에서 비어 있지만 결과에는 원자료 값이 채워진 칸 {blank_filled}개 (비교 제외)")
    return ok_rows and ok_sum and ok_val


def main():
    ap = argparse.ArgumentParser(description="유형자산 대사 8개 사업부 (복원본)")
    ap.add_argument("--data", type=Path, default=HERE / "data" / "multi")
    ap.add_argument("--answer", type=Path, default=HERE / "answer" / "검증정답_8개사업부.xlsx")
    ap.add_argument("--out", type=Path, default=HERE / "results" / "대사결과_8개사업부.xlsx")
    a = ap.parse_args()
    acq = read_folder(a.data / "취득리스트")
    dep = read_folder(a.data / "감가상각대장")
    print(f"취득리스트 {len(acq)}건 ({acq['원본파일'].nunique()}개 파일), "
          f"감가상각대장 {len(dep)}건 ({dep['원본파일'].nunique()}개 파일)")
    exc = reconcile(acq, dep)
    summary = summarize(acq, dep, exc)
    print(exc["예외유형"].value_counts().to_string())
    a.out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(a.out) as w:
        exc.to_excel(w, sheet_name="예외목록", index=False)
        summary.to_excel(w, sheet_name="사업부별요약", index=False)
    ok = verify(exc, summary, a.answer)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
