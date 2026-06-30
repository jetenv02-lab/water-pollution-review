# -*- coding: utf-8 -*-
"""對比「五、推導應檢測水質」 vs 「PDF 實測水質項目」, 找漏測 finding。

策略:
    1. 跑 extract_production_scale 拿每案推導應檢測項目
    2. 跑 step2_extract_v2 拿每案 raw_water + 各單元 effluent 出現過的水質項目
    3. 推導項 - 實測項 = 漏測項 → 出 finding「💡 提醒: 廠商原料含 X, 但水質表沒測」

純新增, 不動任何既有 code。
"""
import json
from pathlib import Path

import step2_extract_v2 as s2
import extract_production_scale as eps

BASE = Path(__file__).parent
PDF_DIR = BASE / "參考" / "需審查之文件"
OUTPUT = BASE / "missing_tested_findings.json"

# 同義詞: 推導項 → PDF 上可能出現的寫法
SYNONYMS = {
    "pH": ["pH值", "pH", "PH", "ph"],
    "懸浮固體（mg/L）": ["懸浮固體（mg/L）", "懸浮固體(mg/L)", "SS", "懸浮固體"],
    "化學需氧量（mg/L）": ["化學需氧量（mg/L）", "化學需氧量(mg/L)", "COD", "化學需氧量"],
    "生化需氧量（mg/L）": ["生化需氧量（mg/L）", "生化需氧量(mg/L)", "BOD", "生化需氧量"],
    "氯離子": ["氯離子", "氯鹽", "氯化物", "Cl"],
    "硫酸根": ["硫酸根", "硫酸鹽", "SO4"],
    "硝酸鹽氮": ["硝酸鹽氮", "硝酸鹽", "NO3-N"],
    "氟鹽": ["氟鹽", "氟化物", "F"],
    "氰化物": ["氰化物", "CN"],
    "總磷": ["總磷", "TP"],
    "氨氮": ["氨氮", "氨氮（mg/L）", "NH3-N"],
    "油脂": ["油脂", "油脂（mg/L）", "油及脂"],
    "真色色度": ["真色色度", "色度"],
    "鋁": ["鋁", "Al"],
    "鈣": ["鈣", "Ca"],
    "鉀": ["鉀", "K"],
    "銅": ["銅", "Cu"],
    "鎳": ["鎳", "Ni"],
    "鋅": ["鋅", "Zn"],
    "鉛": ["鉛", "Pb"],
    "鎘": ["鎘", "Cd"],
    "錫": ["錫", "Sn"],
    "鈷": ["鈷", "Co"],
    "銀": ["銀", "Ag"],
}


def collect_tested_items(extract_data):
    """從 step2 抽取結果蒐集所有實測水質項目。"""
    tested = set()
    raw = extract_data.get("raw_water") or {}
    # raw_water 結構: {WM01: {水質項目: value}}
    if isinstance(raw, dict):
        for wm_data in raw.values():
            if isinstance(wm_data, dict):
                tested.update(wm_data.keys())

    # 各單元 influent / effluent 出現過的水質項目
    units = extract_data.get("units") or {}
    for unit in units.values():
        for streams in (unit.get("influent") or {}, unit.get("effluent") or {}):
            for s in streams.values():
                if isinstance(s, dict):
                    tested.update(s.keys())
    return tested


def is_tested(derived_item, tested_items):
    """判斷推導項是否已實測 (用同義詞)。"""
    candidates = SYNONYMS.get(derived_item, [derived_item])
    for cand in candidates:
        for t in tested_items:
            if cand == t or cand in t or t in cand:
                return True
    return False


def main():
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    all_findings = {}

    print(f"=== 對比「推導應檢測」 vs 「PDF 實測」 ({len(pdfs)} 案) ===\n")
    for pdf in pdfs:
        case_name = pdf.stem

        # 推導應檢測
        try:
            scale = eps.extract_section_5(str(pdf))
        except Exception as e:
            print(f"[X] {case_name} 抽五、失敗: {e}")
            continue
        derived = set(scale.get("推導應檢測水質") or [])

        # 實測
        try:
            ed = s2.extract_application(str(pdf), verbose=False)
        except Exception as e:
            print(f"[X] {case_name} step2 失敗: {e}")
            continue
        tested = collect_tested_items(ed)

        # 對比
        missing = []
        for item in sorted(derived):
            if not is_tested(item, tested):
                missing.append(item)

        all_findings[case_name] = {
            "業別": scale.get("業別"),
            "推導應檢測": sorted(derived),
            "missing": missing,
            "tested_count": len(tested),
        }

        # 印報告
        print(f"📋 {case_name}")
        print(f"  業別: {scale.get('業別')}")
        print(f"  推導應檢測: {len(derived)} 項 → {sorted(derived)}")
        print(f"  實測項目: {len(tested)} 項 (混合 raw + 各單元)")
        if missing:
            print(f"  ❌ 漏測 ({len(missing)} 項):")
            for m in missing:
                print(f"     - {m}")
        else:
            print(f"  ✅ 推導項目都有測")
        print()

    OUTPUT.write_text(json.dumps(all_findings, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n✅ 結果寫入 {OUTPUT.name}")

    # 漏測排行
    print("\n" + "=" * 70)
    print("📊 各案漏測項目排行")
    print("=" * 70)
    for k, v in sorted(all_findings.items(), key=lambda x: -len(x[1]["missing"])):
        print(f"  {k[:40]}: 漏測 {len(v['missing'])} 項")


if __name__ == "__main__":
    main()
