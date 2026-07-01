# -*- coding: utf-8 -*-
"""Baseline 收集器 (v2) — 支援 --quick / --pdf 參數。

對本機 PDF 用 step2_extract_v2 抽取 → 跑 tank_chemistry.check_unit() →
存 baseline_findings.json (預設寫到 data/outputs/)。

用法:
    # 全跑 (慢, 30秒~2分鐘)
    python -X utf8 tests/_run_baseline.py

    # 只跑 1 份指定 PDF (最快, 5-15秒)
    python -X utf8 tests/_run_baseline.py --pdf "線上資料(馥廷)(1150617).pdf"

    # 只跑最小 3 案快速樣本 (~20秒)
    python -X utf8 tests/_run_baseline.py --quick

    # 指定輸出檔
    python -X utf8 tests/_run_baseline.py --output my_baseline.json
"""
import argparse
import json
import sys
from pathlib import Path

# 因為此檔在 tests/, 需要把 repo root 加進 sys.path
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import step2_extract_v2 as s2  # noqa: E402
import tank_chemistry as tc  # noqa: E402

PDF_DIR = REPO_ROOT / "參考" / "需審查之文件"
DEFAULT_OUTPUT = REPO_ROOT / "data" / "outputs" / "baseline_findings.json"

# --quick 模式的樣本 (代表電鍍/PCB/紙板 3 大業別)
QUICK_SAMPLE = [
    "申請文件(秋棠)(1150519).pdf",  # 電鍍 46 單元, 5 findings
    "01 線上資料(邑昇)(1150529).pdf",  # PCB 34 單元, 9 findings
    "線上資料(永豐餘)(1150605).pdf",  # 紙板 34 單元, 0 findings
]


def parse_args():
    p = argparse.ArgumentParser(description="Baseline 檢查跑 tank_chemistry.")
    p.add_argument("--pdf", help="只跑單一 PDF (檔名 or 路徑)")
    p.add_argument("--quick", action="store_true",
                   help="只跑 QUICK_SAMPLE (秋棠/邑昇/永豐餘 3 案)")
    p.add_argument("--output", default=str(DEFAULT_OUTPUT),
                   help=f"輸出 JSON 路徑 (預設 {DEFAULT_OUTPUT.relative_to(REPO_ROOT)})")
    p.add_argument("--pdf-dir", default=str(PDF_DIR),
                   help="PDF 資料夾")
    return p.parse_args()


def resolve_pdfs(args):
    """依 args 決定要跑哪些 PDF。"""
    pdf_dir = Path(args.pdf_dir)
    if args.pdf:
        p = Path(args.pdf)
        if not p.is_absolute():
            # 相對名: 先找 pdf_dir/name 再找 name
            candidates = [pdf_dir / p.name, Path(args.pdf)]
            for c in candidates:
                if c.exists():
                    return [c]
            print(f"[X] 找不到 {args.pdf}")
            sys.exit(1)
        return [p]
    if args.quick:
        found = []
        for name in QUICK_SAMPLE:
            candidate = pdf_dir / name
            if candidate.exists():
                found.append(candidate)
            else:
                print(f"[warn] --quick 樣本找不到: {name}")
        return found
    # 全部
    return sorted(pdf_dir.glob("*.pdf"))


def main():
    args = parse_args()
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rules = tc.load_rules()
    print(f"[1/3] 規則庫條目: {len(rules)}")

    pdfs = resolve_pdfs(args)
    print(f"[2/3] 要跑 {len(pdfs)} 份 PDF:")
    for p in pdfs:
        print(f"  - {p.name}")
    print()

    all_baselines = {}
    summary = []
    for pdf in pdfs:
        case_name = pdf.stem
        try:
            d = s2.extract_application(str(pdf), verbose=False)
        except Exception as e:
            print(f"[X] {case_name} step2 失敗: {e}")
            continue

        units = d.get("units") or {}
        case_findings = {}
        total_units_with_wq = 0
        total_findings = 0
        for code, unit in units.items():
            inf = unit.get("influent") or {}
            eff = unit.get("effluent") or {}
            if not inf or not eff:
                continue
            total_units_with_wq += 1
            unit_copy = dict(unit)
            unit_copy.setdefault("code_id", code)
            findings = tc.check_unit(unit_copy, rules)
            if findings:
                case_findings[code] = findings
                total_findings += len(findings)

        all_baselines[case_name] = case_findings
        summary.append({
            "case": case_name,
            "total_units": len(units),
            "units_with_water_quality": total_units_with_wq,
            "units_with_findings": len(case_findings),
            "total_findings": total_findings,
        })
        print(f"  ✓ {case_name}: {len(units)} 單元 / "
              f"{total_units_with_wq} 有水質 / {total_findings} findings")

    output_path.write_text(
        json.dumps(all_baselines, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\n[3/3] Baseline 寫入 {output_path.relative_to(REPO_ROOT)}")
    print()
    print("=" * 80)
    print(f"{'案件':45} {'總單元':>8} {'有水質':>8} {'總findings':>12}")
    print("=" * 80)
    for s in summary:
        name = s['case'][:45]
        print(f"{name:45} {s['total_units']:>8} "
              f"{s['units_with_water_quality']:>8} {s['total_findings']:>12}")
    print("=" * 80)
    print()
    print("✅ baseline 已存. Step 3 改完 tank_chemistry 後重跑此檔對比差異.")


if __name__ == "__main__":
    main()
