# -*- coding: utf-8 -*-
"""加藥質量補償 — 離線驗證工具 (純新增, 不動 tank_chemistry.py)。

目的:
    模擬「若 tank_chemistry.check_unit() 加上 chemical_calc 補償, finding 會怎麼變」
    用於 review 階段, 看出 patch 影響範圍, 確認方向後才動 code。

做法:
    1. 對所有 PDF 跑 step2_extract_v2
    2. 對每個有水質的單元, 用「複製粘貼版 check_unit + 加藥補償」算 finding
    3. 對比 baseline (tank_chemistry 原版) 跟「補償版」差異
    4. 輸出對比表

不修改任何既有檔案。
"""
import json
from pathlib import Path

import step2_extract_v2 as s2
import tank_chemistry as tc
import chemical_calc

BASE = Path(__file__).parent
PDF_DIR = BASE / "參考" / "需審查之文件"
BASELINE_JSON = BASE / "baseline_findings.json"
OUTPUT = BASE / "verify_dosing_result.json"


def _to_float(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def check_unit_with_dosing(unit, rules):
    """tank_chemistry.check_unit() 的「加藥補償版」。

    100% 複製原版邏輯, 只在 diff_pct 計算前注入 chemical_added。
    結果為「假設 patch 後」的 finding list。
    """
    code = unit.get("raw_code") or unit.get("code_id") or "?"
    influent = unit.get("influent") or {}
    effluent = unit.get("effluent") or {}
    if not influent or not effluent:
        return []

    # 用跟 tank_chemistry.check_unit 一樣的 rule 查表方式 (有同義字解析)
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

        # ── 新增: 加藥質量補償 ──
        q_for_chem = in_q_sum if in_q_sum > 0 else None
        chemical_added = chemical_calc.compute_chemical_mass(unit, item, q_for_chem)
        in_mass_with_chem = in_mass + chemical_added

        # 用補償後質量算 diff_pct
        diff_pct = abs(out_mass - in_mass_with_chem) / in_mass_with_chem * 100
        effective_tol = max(tol, 0.5)
        if diff_pct <= effective_tol:
            continue

        # 後續邏輯不變 (流向提示、加權平均、重金屬 hint)
        topology_hint = ""
        downgrade = False
        if self_split and in_q_sum > 0 and out_q_sum > 0:
            in_avg_c = in_conc_x_q / in_q_sum
            out_avg_c = out_conc_x_q / out_q_sum
            if in_avg_c > 0:
                conc_diff_pct = abs(out_avg_c - in_avg_c) / in_avg_c * 100
                if conc_diff_pct <= max(tol, 5.0):
                    topology_hint = " ⚠️ 流向提示 (略)"
                    downgrade = True

        direction = "減少" if out_mass < in_mass_with_chem else "增加"
        eff_sev = "待確認" if downgrade else severity

        # 加藥描述
        chem_desc = chemical_calc.describe_chemical_contribution(unit, item, q_for_chem) if chemical_added > 0 else ""

        findings.append({
            "嚴重度": eff_sev,
            "類型": "質量平衡",
            "單元": code,
            "標準槽體": std_tank,
            "對照項目": item,
            "描述": (
                f"{item} 質量 進 {in_mass:.3f}"
                f"{' + 加藥估算 ' + f'{chemical_added:.3f}' if chemical_added > 0 else ''}"
                f" → 出 {out_mass:.3f} kg/d "
                f"({direction} {diff_pct:.1f}%, 容忍 {tol}%)。{chem_desc}"
            ),
            "依據": f"_槽體學理 規則: {std_tank} ({rule['加藥類型']}) [加藥補償模式]",
            "_chemical_added": chemical_added,  # debug 用
        })
    return findings


def main():
    rules = tc.load_rules()
    print(f"[1/4] 規則庫條目: {len(rules)}")
    print(f"[2/4] 載入 baseline...")
    if not BASELINE_JSON.exists():
        print(f"[X] 找不到 {BASELINE_JSON.name}, 請先跑 _run_baseline.py")
        return
    baseline = json.loads(BASELINE_JSON.read_text(encoding="utf-8"))

    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    print(f"[3/4] 跑 {len(pdfs)} 份 PDF, 計算「假設加藥補償」後 finding...")
    print()

    all_with_dosing = {}
    diff_summary = []

    for pdf in pdfs:
        case_name = pdf.stem
        d = s2.extract_application(str(pdf), verbose=False)
        units = d.get("units") or {}

        case_with_dosing = {}
        case_findings_after = 0
        for code, unit in units.items():
            unit_copy = dict(unit)
            unit_copy.setdefault("code_id", code)
            findings = check_unit_with_dosing(unit_copy, rules)
            if findings:
                case_with_dosing[code] = findings
                case_findings_after += len(findings)
        all_with_dosing[case_name] = case_with_dosing

        # 對比 baseline
        baseline_case = baseline.get(case_name) or {}
        baseline_count = sum(len(v) for v in baseline_case.values())

        # 詳細 diff: 哪些 finding 消失了
        disappeared = []
        for unit_code, b_finds in baseline_case.items():
            d_finds = case_with_dosing.get(unit_code) or []
            b_items = {f["對照項目"] for f in b_finds}
            d_items = {f["對照項目"] for f in d_finds}
            for vanished in b_items - d_items:
                disappeared.append({"unit": unit_code, "item": vanished})

        # 哪些新冒出
        new_findings = []
        for unit_code, d_finds in case_with_dosing.items():
            b_finds = baseline_case.get(unit_code) or []
            b_items = {f["對照項目"] for f in b_finds}
            d_items = {f["對照項目"] for f in d_finds}
            for new in d_items - b_items:
                new_findings.append({"unit": unit_code, "item": new})

        diff_summary.append({
            "case": case_name,
            "baseline_findings": baseline_count,
            "after_dosing_findings": case_findings_after,
            "delta": case_findings_after - baseline_count,
            "disappeared": disappeared,
            "new": new_findings,
        })

    # 寫詳細結果
    OUTPUT.write_text(
        json.dumps({"with_dosing": all_with_dosing, "diff_summary": diff_summary},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 印對比表
    print(f"[4/4] 結果寫入 {OUTPUT.name}")
    print()
    print("=" * 90)
    print(f"{'案件':40} {'baseline':>10} {'加藥補償後':>12} {'差':>8} {'消失':>6} {'新增':>6}")
    print("=" * 90)
    total_b, total_a, total_d, total_n = 0, 0, 0, 0
    for s in diff_summary:
        name = s['case'][:40]
        d = s['delta']
        delta_str = f"{d:+d}" if d != 0 else "0"
        print(f"{name:40} {s['baseline_findings']:>10} {s['after_dosing_findings']:>12} "
              f"{delta_str:>8} {len(s['disappeared']):>6} {len(s['new']):>6}")
        total_b += s['baseline_findings']
        total_a += s['after_dosing_findings']
        total_d += len(s['disappeared'])
        total_n += len(s['new'])
    print("=" * 90)
    delta_total = total_a - total_b
    delta_total_str = f"{delta_total:+d}" if delta_total != 0 else "0"
    print(f"{'總計':40} {total_b:>10} {total_a:>12} "
          f"{delta_total_str:>8} {total_d:>6} {total_n:>6}")
    print()

    # 詳細列出消失 + 新增
    print("=" * 90)
    print("📉 加藥補償後「消失」的 finding (代表這些 baseline finding 其實是加藥造成的, 不算錯):")
    print("=" * 90)
    for s in diff_summary:
        if s['disappeared']:
            print(f"\n[{s['case']}]")
            for d in s['disappeared']:
                print(f"  - {d['unit']} / {d['item']}")
    print()
    print("=" * 90)
    print("📈 加藥補償後「新增」的 finding (理論上不應該有, 若有代表轉換係數過大):")
    print("=" * 90)
    for s in diff_summary:
        if s['new']:
            print(f"\n[{s['case']}]")
            for d in s['new']:
                print(f"  - {d['unit']} / {d['item']}")


if __name__ == "__main__":
    main()
