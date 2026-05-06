#!/usr/bin/env python3
"""
근로소득 간이세액표 엑셀 → JSON 변환기

입력:
  data/근로소득 간이세액표_2026.03.01.xlsx
    - 시트 '근로소득간이세액표': 룩업 표 + 1,000만원 초과 공식

출력:
  data/근로소득_간이세액표_2026-03.json
"""
from __future__ import annotations
import json
import re
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SRC = DATA / "근로소득 간이세액표_2026.03.01.xlsx"
OUT = DATA / "근로소득_간이세액표_2026-03.json"


def to_int(v):
    if v in (None, "", "-"):
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return 0
    return int(float(s))


def parse_amount(text: str) -> int:
    """'1,397,000원' → 1397000"""
    m = re.search(r"([\d,]+)\s*원", text)
    return int(m.group(1).replace(",", "")) if m else 0


def parse_thousand(text: str) -> int:
    """'14,000천원' → 14000"""
    m = re.search(r"([\d,]+)\s*천원", text)
    return int(m.group(1).replace(",", "")) if m else 0


def parse_main_rate(text: str) -> float:
    """공식 텍스트에서 98%를 제외한 적용세율(35%/38%/40%/42%/45%) 추출"""
    rates = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s*%", text)]
    rates = [r for r in rates if abs(r - 98) > 0.001]
    return rates[-1] / 100 if rates else 0.0


def has_98(text: str) -> bool:
    return "98%" in text or "0.98" in text


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["근로소득간이세액표"]

    brackets = []
    base_at_10m = None  # 10,000천원 정확 기준 부양가족 1~11인 세액
    formulas = []
    pending_floor = None  # 직전 '초과' 행에서 추출한 from(천원)
    pending_data = None

    for row in ws.iter_rows(min_row=5, values_only=True):
        a = row[0]
        b = row[1]

        # 일반 룩업 (이상/미만 둘 다 숫자)
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            taxes = [to_int(v) for v in row[2:13]]
            brackets.append([int(a), int(b), *taxes])
            continue

        # 10,000천원 정확 기준값 행
        if isinstance(a, str) and a.strip().startswith("10,000천원") and "초과" not in a:
            base_at_10m = [to_int(v) for v in row[2:13]]
            continue

        # 'X천원 초과' (공식 행)
        if isinstance(a, str) and "초과" in a:
            frm = parse_thousand(a)
            formula_text = row[2] if isinstance(row[2], str) else ""
            # 공식 텍스트의 모든 '원' 금액을 합산 → 가산상수
            amounts = [int(x.replace(",", "")) for x in re.findall(r"([\d,]+)\s*원", formula_text)]
            pending_data = {
                "from": frm,
                "to": None,  # 다음 '이하' 행에서 채움
                "addConst": sum(amounts),
                "rate": parse_main_rate(formula_text),
                "factor": 0.98 if has_98(formula_text) else 1.0,
            }
            formulas.append(pending_data)
            continue

        # 'X천원 이하' → 직전 공식의 to 채우기
        if isinstance(a, str) and "이하" in a and pending_data is not None:
            pending_data["to"] = parse_thousand(a)
            pending_data = None
            continue

    # 10,000천원 본인 1인 기준이 없으면 본 표 마지막 가까운 값에서 보간
    if base_at_10m is None:
        raise RuntimeError("10,000천원 기준값을 찾지 못했습니다.")

    out = {
        "effectiveFrom": "2026-03-01",
        "source": "data/근로소득 간이세액표_2026.03.01.xlsx",
        "unit": {"salary": "원", "tax": "원", "bracketSalary": "천원"},
        "dependentColumns": list(range(1, 12)),  # 1~11명
        "brackets": brackets,
        "baseAt10M": base_at_10m,
        "highIncomeFormula": formulas,
        "note": (
            "부양가족 11명 초과 시: tax_n = tax_11 - (tax_10 - tax_11) × (n - 11). "
            "1,000만원 초과 시: tax = baseAt10M[deps-1] + addConst + (초과금액(원) × factor × rate)."
        ),
    }

    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"✓ 생성: {OUT}")
    print(f"  - 룩업 행: {len(brackets)}")
    print(f"  - 1,000만원 기준값(11인): {base_at_10m}")
    print(f"  - 고소득 공식: {len(formulas)}개")
    for f in formulas:
        print(f"    · {f['from']:>5} ~ {f['to']!s:>5} 천원 | addConst={f['addConst']:>10} | rate={f['rate']:.2f} × factor={f['factor']}")


if __name__ == "__main__":
    main()
