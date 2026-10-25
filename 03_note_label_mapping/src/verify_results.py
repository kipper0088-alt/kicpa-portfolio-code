"""
[재현본] 주석 항목명 표준화 — 결과 검증

세 가지 상태(매핑 전 / 동의어 매핑 / 오타 승인 후)를 다시 실행해
results 폴더의 기대 결과 JSON과 같은지 확인하고, 자체 테스트도 실행한다.

사용법 (프로젝트 폴더에서):
    python src/verify_results.py
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import note_mapping_demo as m  # noqa: E402

ROOT = HERE.parent
notes = json.loads((ROOT / "data" / "sample_notes.json").read_text(encoding="utf-8"))["notes"]
cases = [
    ("result_before.json", dict(use_aliases=False)),
    ("result_after_mapping.json", dict()),
    ("result_after_approval.json", dict(approvals=m.parse_approvals(["당기증갸=증가"]))),
]
ok = True
for name, kw in cases:
    got = m.run_demo(notes, **kw)
    exp = json.loads((ROOT / "results" / name).read_text(encoding="utf-8"))
    same = got == exp
    ok &= same
    print(f"{name:<28} 요약 {got['요약']}  -> {'PASS' if same else 'FAIL'}")
m.self_test(notes)
print("최종:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
