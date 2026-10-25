"""[재현본] 회계 주석 항목명 표준화·대사 (2026-07 실습을 재현).

원본 multi_note_crosscheck_v4.py 파일이 아닙니다. 당시 확인된 기능을
합성 데이터와 Python 표준 라이브러리로 재구현했습니다.

실행 예시:
    python note_mapping_demo.py
    python note_mapping_demo.py --no-alias
    python note_mapping_demo.py --approve 당기증갸=증가
    python note_mapping_demo.py --self-test
"""

from __future__ import annotations

import argparse
import difflib
import json
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


# 당시 확인된 증가·감소 계열 동의어와 데모 구동에 필요한 기초·기말·상각 표현.
# 기초·기말·상각 별칭은 이 복원본에서 편의상 추가한 설정입니다.
ALIASES: dict[str, tuple[str, ...]] = {
    "기초잔액": ("기초잔액", "기초"),
    "증가": ("증가", "취득", "당기증가", "증가액"),
    "감소": ("감소", "처분", "당기감소", "감소액"),
    "상각": ("상각", "감가상각", "당기상각"),
    "기말잔액": ("기말잔액", "기말"),
}
REQUIRED = ("기초잔액", "증가", "감소", "기말잔액")
OPTIONAL = ("상각",)


def normalize(label: str) -> str:
    """공백·구두점을 제거해 표기 형태만 달라진 문자열을 비교합니다."""
    normalized = unicodedata.normalize("NFKC", str(label)).lower()
    return "".join(ch for ch in normalized if ch.isalnum())


def lookup_table() -> dict[str, tuple[str, str]]:
    result: dict[str, tuple[str, str]] = {}
    for canonical, names in ALIASES.items():
        for name in names:
            key = normalize(name)
            if key in result and result[key][0] != canonical:
                raise ValueError(f"중복된 의미의 별칭: {name}")
            result[key] = (canonical, name)
    return result


def parse_approvals(raw_approvals: list[str]) -> dict[str, str]:
    """사용자가 확인·승인한 표현만 수동 매핑합니다."""
    approvals: dict[str, str] = {}
    for item in raw_approvals:
        if "=" not in item:
            raise ValueError("--approve는 '원래표현=기준항목' 형식이어야 합니다.")
        raw_label, canonical = (part.strip() for part in item.split("=", 1))
        if not raw_label or canonical not in ALIASES:
            raise ValueError(f"승인 규칙이 올바르지 않습니다: {item}")
        approvals[normalize(raw_label)] = canonical
    return approvals


def resolve_label(
    label: str, *, use_aliases: bool, approvals: dict[str, str]
) -> dict[str, Any]:
    """미등록·오타를 자동 확정하지 않고 사람에게 후보만 제공합니다."""
    key = normalize(label)
    canonical_keys = {normalize(name): name for name in ALIASES}
    known = lookup_table()

    if key in approvals:
        return {"standard": approvals[key], "method": "사용자 승인", "candidates": []}
    if key in canonical_keys:
        return {"standard": canonical_keys[key], "method": "정확 일치", "candidates": []}
    if use_aliases and key in known:
        return {"standard": known[key][0], "method": "동의어 매핑", "candidates": []}

    if key in known:
        # 기준 표현 자체는 정확하지만 동의어 규칙을 꺼 둔 상황.
        candidates = [{"표현": known[key][1], "기준항목": known[key][0]}]
        reason = "동의어 규칙 미적용"
    else:
        # 후보일 뿐 자동 변환하지 않습니다. 당시 대화에서 확인된 cutoff=0.5.
        nearby = difflib.get_close_matches(key, list(known), n=3, cutoff=0.5)
        candidates = [{"표현": known[name][1], "기준항목": known[name][0]} for name in nearby]
        reason = "미등록 표현 또는 오타"

    return {
        "standard": None,
        "method": "검토 필요",
        "reason": reason,
        "candidates": candidates,
    }


def number(value: Any) -> Decimal:
    try:
        result = Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(f"금액을 읽을 수 없습니다: {value!r}") from exc
    if not result.is_finite():
        raise ValueError(f"유효하지 않은 금액: {value!r}")
    return result


def json_number(value: Decimal) -> int | float:
    return int(value) if value == value.to_integral_value() else float(value)


def evaluate_note(
    note: dict[str, Any], *, use_aliases: bool = True, approvals: dict[str, str] | None = None
) -> dict[str, Any]:
    approvals = approvals or {}
    observed: dict[str, list[Decimal]] = {}
    mappings: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []

    for row in note["rows"]:
        label = str(row["label"])
        amount = number(row["amount"])
        resolution = resolve_label(label, use_aliases=use_aliases, approvals=approvals)
        entry = {"원본표현": label, "금액": json_number(amount), "기준항목": resolution["standard"], "처리": resolution["method"]}
        if resolution["standard"] is None:
            entry["사유"] = resolution["reason"]
            entry["추천후보"] = resolution["candidates"]
            unresolved.append(entry)
        else:
            observed.setdefault(resolution["standard"], []).append(amount)
        mappings.append(entry)

    missing = [key for key in REQUIRED if key not in observed]
    duplicates = [key for key, values in observed.items() if len(values) > 1]
    result: dict[str, Any] = {
        "id": note["id"],
        "주석": note["name"],
        "재무상태표금액": note["balance_sheet"],
        "원본항목_매핑과정": mappings,
        "미등록_검토대상": unresolved,
        "필수항목누락": missing,
        "중복항목": duplicates,
        "계산기말잔액": None,
        "주석기말잔액": None,
        "재무상태표차이": None,
    }
    if unresolved or missing or duplicates:
        result["결과"] = "검토 필요"
        return result

    values = {key: observed[key][0] for key in observed}
    computed = values["기초잔액"] + values["증가"] - values["감소"] - values.get("상각", Decimal(0))
    reported = values["기말잔액"]
    statement = number(note["balance_sheet"])
    result["계산기말잔액"] = json_number(computed)
    result["주석기말잔액"] = json_number(reported)
    result["재무상태표차이"] = json_number(computed - statement)
    result["결과"] = "일치" if computed == reported == statement else "금액 불일치"
    return result


def run_demo(
    notes: list[dict[str, Any]], *, use_aliases: bool = True, approvals: dict[str, str] | None = None
) -> dict[str, Any]:
    items = [evaluate_note(note, use_aliases=use_aliases, approvals=approvals) for note in notes]
    counts = Counter(item["결과"] for item in items)
    return {
        "실습설명": "복수 주석 항목명 표준화 복원 데모 - 합성 데이터",
        "동의어매핑적용": use_aliases,
        "수동승인규칙": approvals or {},
        "요약": {"일치": counts["일치"], "검토 필요": counts["검토 필요"], "금액 불일치": counts["금액 불일치"]},
        "주석별결과": items,
    }


def self_test(notes: list[dict[str, Any]]) -> None:
    standard = run_demo(notes)
    by_id = {item["id"]: item for item in standard["주석별결과"]}
    assert standard["요약"] == {"일치": 3, "검토 필요": 1, "금액 불일치": 0}
    assert by_id["ppe"]["계산기말잔액"] == 1150
    assert by_id["intangible"]["계산기말잔액"] == 540
    assert by_id["provision"]["계산기말잔액"] == 850
    assert by_id["typo"]["결과"] == "검토 필요"
    assert any(candidate["표현"] == "당기증가" for issue in by_id["typo"]["미등록_검토대상"] for candidate in issue["추천후보"])

    strict = run_demo(notes, use_aliases=False)
    assert strict["요약"]["검토 필요"] >= 3

    approved = run_demo(notes, approvals=parse_approvals(["당기증갸=증가"]))
    assert approved["요약"] == {"일치": 4, "검토 필요": 0, "금액 불일치": 0}
    print("SELF-TEST PASS: 동의어 매핑 / 미등록 후보 / 사용자 승인 / 재계산")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path(__file__).with_name("sample_notes.json"))
    parser.add_argument("--output", type=Path, help="결과 JSON 저장 경로(생략하면 화면 출력)")
    parser.add_argument("--no-alias", action="store_true", help="동의어 매핑 규칙을 사용하지 않음")
    parser.add_argument("--approve", action="append", default=[], help="검토 후 승인할 표현(예: 당기증갸=증가)")
    parser.add_argument("--self-test", action="store_true", help="시나리오 검증")
    args = parser.parse_args()

    notes = json.loads(args.input.read_text(encoding="utf-8"))["notes"]
    if args.self_test:
        self_test(notes)
        return
    report = run_demo(notes, use_aliases=not args.no_alias, approvals=parse_approvals(args.approve))
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.write_text(payload, encoding="utf-8")
        print(f"결과 저장: {args.output}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
