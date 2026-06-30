# -*- coding: utf-8 -*-
"""加藥質量補償 — v2 驗證 (讀 Sheets _加藥規則, 取代 hardcode 典型值)。

跟 verify_with_dosing.py 結構一樣, 只差 chemical_calc → dosing_rules_loader。
"""
import json
from pathlib import Path

import step2_extract_v2 as s2
import tank_chemistry as tc
import dosing_rules_loader as drl  # v2: 讀 Sheets, 取代 chemical_calc

BASE = Path(__file__).parent
PDF_DIR = BASE / "參考" / "需審查之文件"
BASELINE_JSON = BASE / "baseline_findings.json"
OUTPUT = BASE / "verify_dosing_v2_result.json"


def _to_float(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def check_unit_with_dosing_v2(unit, rules):
    """跟 verify_with_dosing 的 check_unit_with_dosing 一樣, 只是 chemical_calc → drl。"""
    code = unit.get("raw_code") or unit.get("code_id") or "?"
    influent = unit.get("influent") or {}
    effluent = unit.get("effluent") or {}
    if not influent or not effluent:
        return []

    rule = tc.get_rule_for_unit(unit, rules)
    if not rule:
        return []

    std_tank = rule["標準槽體"]
    tol = rule.get("容忍度") or 5.0
    severity = rule.get("嚴重度") or "不合理"
    desc_text = rule.get("學理說明") or ""

    all_items = set()
    for stream in list(influent.values()) + list(effluent.values()):
        if isinstance(stream, dict):
            all_items.update(stream.keys())

    self_split = len(effluent) >= 2

    findings = []
    for item in sorted(all_items):
        cls = tc.classify_item(item, rule)
        if cls != "不應變動":
            continue

        in_mass = 0.0
        out_mass = 0.0
        in_q_sum = 0.0
        out_q_sum = 0.0
        in_conc_x_q = 0.0
        out_conc_x_q = 0.0
        in_has = False
        out_has = False
        unit_stream_q = unit.get("stream_q") or {}
        for stream_code, stream in influent.items():
            if not isinstance(stream, dict):
                continue
            v = stream.get(item)
            sq_info = unit_stream_q.get(stream_code) or {}
            q_from_sq = _to_float(sq_info.get("q_cmd")) if isinstance(sq_info, dict) else None
            if isinstance(v, dict):
                m = _to_float(v.get("質量"))
                c = _to_float(v.get("濃度"))
                q = q_from_sq or _to_float(v.get("Q") or v.get("q") or v.get("q_cmd"))
                if m is not None:
                    in_mass += m
                    in_has = True
                if c is not None and q is not None and q > 0:
                    in_conc_x_q += c * q
                    in_q_sum += q
        for stream_code, stream in effluent.items():
            if not isinstance(stream, dict):
                continue
            v = stream.get(item)
            sq_info = unit_stream_q.get(stream_code) or {}
            q_from_sq = _to_float(sq_info.get("q_cmd")) if isinstance(sq_info, dict) else None
            if isinstance(v, dict):
                m = _to_float(v.get("質量"))
                c = _to_float(v.get("濃度"))
                q = q_from_sq or _to_float(v.get("Q") or v.get("q") or v.get("q_cmd"))
                if m is not None:
                    out_mass += m
                    out_has = True
                if c is not None and q is not None and q > 0:
                    out_conc_x_q += c * q
                    out_q_sum += q

        if not (in_has and out_has and in_mass > 0):
            continue

        # ── v2: 從 Sheets 讀補償 ──
        q_for_chem = in_q_sum if in_q_sum > 0 else None
        chemical_added = drl.compute_chemical_mass(unit, item, q_for_chem)
        in_mass_with_chem = in_mass + chemical_added

        diff_pct = abs(out_mass - in_mass_with_chem) / in_mass_with_chem * 100
        effective_tol = max(tol, 0.5)
        if diff_pct <= effective_tol:
            continue

        direction = "減少" if out_mass < in_mass_with_chem else "增加"
        chem_desc = drl.describe_chemical_contribution(unit, item, q_for_chem) if chemical_added > 0 else ""

        findings.append({
            "嚴重度": severity,
            "類型": "質量平衡",
            "單元": code,
            "標準槽體": std_tank,
            "對照項目": item,
            "描述": (
                f"{item} 質量 進 {in_mass:.3f}"
                f"{' + 加藥估算 ' + f'{chemical_added:.3f}' if chemical_added > 0 else ''}"
                f" → 出 {out_mass:.3f} kg/d ({direction} {diff_pct:.1f}%, 容忍 {tol}%)。{chem_desc}"
            ),
            "依據": f"_槽體學理 + _加藥規則 [v2 加藥補償模式]",
            "_chemical_added": chemical_added,
        })
    return findings


def main():
    rules = tc.load_rules()
    print(f"[1/4] 規則庫條目: {len(rules)}")
    print(f"[2/4] 載入 baseline...")
    baseline = json.loads(BASELINE_JSON.read_text(encoding="utf-8"))

    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    print(f"[3/4] 跑 {len(pdfs)} 份 PDF (v2: 從 Sheets 讀補償)")
    print()

    all_v2 = {}
    diff_summary = []
    for pdf in pdfs:
        case_name = pdf.stem
        d = s2.extract_application(str(pdf), verbose=False)
        units = d.get("units") or {}

        case_v2 = {}
        case_total = 0
        for code, unit in units.items():
            unit_copy = dict(unit)
            unit_copy.setdefault("code_id", code)
            findings = check_unit_with_dosing_v2(unit_copy, rules)
            if findings:
                case_v2[code] = findings
                case_total += len(findings)
        all_v2[case_name] = case_v2

        baseline_case = baseline.get(case_name) or {}
        baseline_count = sum(len(v) for v in baseline_case.values())

        disappeared = []
        for unit_code, b_finds in baseline_case.items():
            d_finds = case_v2.get(unit_code) or []
            b_items = {f["對照項目"] for f in b_finds}
            d_items = {f["對照項目"] for f in d_finds}
            for vanished in b_items - d_items:
                disappeared.append({"unit": unit_code, "item": vanished})

        new_findings = []
        for unit_code, d_finds in case_v2.items():
            b_finds = baseline_case.get(unit_code) or []
            b_items = {f["對照項目"] for f in b_finds}
            d_items = {f["對照項目"] for f in d_finds}
            for new in d_items - b_items:
                new_findings.append({"unit": unit_code, "item": new})

        diff_summary.append({
            "case": case_name,
            "baseline_findings": baseline_count,
            "after_v2_findings": case_total,
            "delta": case_total - baseline_count,
            "disappeared": disappeared,
            "new": new_findings,
        })

    OUTPUT.write_text(
        json.dumps({"with_v2_dosing": all_v2, "diff_summary": diff_summary},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"[4/4] 結果寫入 {OUTPUT.name}")
    print()
    print("=" * 90)
    print(f"{'案件':45} {'baseline':>10} {'v2 後':>10} {'差':>8} {'消失':>6} {'新增':>6}")
    print("=" * 90)
    total_b = total_a = total_d = total_n = 0
    for s in diff_summary:
        d = s['delta']
        delta_str = f"{d:+d}" if d != 0 else "0"
        print(f"{s['case'][:45]:45} {s['baseline_findings']:>10} {s['after_v2_findings']:>10} "
              f"{delta_str:>8} {len(s['disappeared']):>6} {len(s['new']):>6}")
        total_b += s['baseline_findings']
        total_a += s['after_v2_findings']
        total_d += len(s['disappeared'])
        total_n += len(s['new'])
    print("=" * 90)
    delta_t_str = f"{total_a - total_b:+d}" if total_a != total_b else "0"
    print(f"{'總計':45} {total_b:>10} {total_a:>10} {delta_t_str:>8} {total_d:>6} {total_n:>6}")
    print()

    print("=" * 90)
    print("📉 v2 加藥補償後「消失」的 finding:")
    print("=" * 90)
    for s in diff_summary:
        if s['disappeared']:
            print(f"\n[{s['case']}]")
            for d in s['disappeared']:
                print(f"  - {d['unit']} / {d['item']}")
    print()
    print("=" * 90)
    print("📈 v2 加藥補償後「新增」的 finding:")
    print("=" * 90)
    for s in diff_summary:
        if s['new']:
            print(f"\n[{s['case']}]")
            for d in s['new']:
                print(f"  - {d['unit']} / {d['item']}")


if __name__ == "__main__":
    main()
