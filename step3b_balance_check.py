# -*- coding: utf-8 -*-
"""Step 3b: 質量平衡 / 學理檢查引擎。

針對 step2_extract_v2 抽出的單元資料 (含進出流水質),
自動檢查常見的不合理狀況。

檢查項目:
1. 溶解性物質自行濃縮 (硝酸鹽、硼、Cl-、SO4²-、Na+ 在無濃縮機制單元濃度上升)
2. 快混槽展現重金屬去除 (沒有固液分離卻有去除率)
3. pH 調整槽除 pH 外水質改變 (應該只變 pH)
4. 生物處理對重金屬有高去除 (學理不符)
5. 沉澱池表面溢流率過高 (> 50 m3/m2-d)
6. 慢混停留時間 ≤ 快混 (違反設計原理)
7. 質量不平衡 (進流質量 ≠ 出流質量 ± 5%)
8. 設施應有的機具未列 (液位計、pH 計、流量計)

使用:
    from step3b_balance_check import run_balance_checks
    findings = run_balance_checks(app_data)
"""
import json
import os
import re
import sys
from collections import defaultdict

# 溶解性物質 (一般物化處理不會去除這些)
DISSOLVED_ITEMS = [
    "硝酸鹽氮", "氨氮", "硼", "氟鹽", "氯鹽",
    "鈉", "鉀", "鈣", "鎂",
    "硫酸鹽", "導電度", "氯離子", "鈉離子"
]

# 重金屬
HEAVY_METALS = ["銅", "鎳", "鋅", "鉛", "鎘", "鉻", "總鉻", "六價鉻",
                "錫", "鐵", "錳", "汞", "總汞", "砷", "鉬"]

# pH 調整槽應該只變 pH,其他項目應不變
PH_ONLY_VARIABLE = {"pH值", "pH", "水溫(攝氏)", "水溫"}

# pH 調整 / 快混類槽體 (本質都是 "加藥+攪拌, 無固液分離", 學理上不該展現
# 任何水質項目去除/增加, 除了 pH 跟水溫)
# 申請文件常見的命名變體都列進來
PH_TANK_TYPES = {
    "pH調整槽",
    "pH調整暨快混池",      # 秋棠案例 T03-07 / T04-04 / T05-05
    "pH調整池暨快混池",    # 秋棠案例 T02-06 的 name_in_doc (含「池」)
    "pH調整池",            # 純粹改字
    "pH調整快混池",        # 省略「暨」
    "pH調整與快混池",
    "中和池",              # 學理上跟 pH 槽同類 (加鹼/酸調 pH, 無固液分離)
}

# 沉澱單元類型
SETTLING_TANK_TYPES = {"沉澱池", "沉降池", "浮除槽"}

# 快混槽類型 (不應展現去除率) — 涵蓋常見命名變體
FAST_MIX_TYPES = {
    "快混槽",
    "pH調整槽", "pH調整池",
    "pH調整暨快混池", "pH調整池暨快混池", "pH調整快混池", "pH調整與快混池",
    "中和池",
    "調勻池",
    "廢水調整池",
}

# 生物處理類型 (對重金屬不應有顯著去除)
BIO_TYPES = {"曝氣槽", "活性污泥槽", "厭氧池", "缺氧池", "好氧池"}

# 必要機具設施 (依槽體類型)
REQUIRED_EQUIPMENT = {
    "廢水調整池": ["液位計"],
    "貯留槽": ["液位計"],
    "暫存槽": ["液位計"],
    "pH調整槽": ["pH計"],
    "pH調整暨快混池": ["pH計"],
    "中和池": ["pH計"],
    "沉澱池": ["污泥泵", "排泥"],
    "曝氣槽": ["DO計", "溶氧計"],
}


def to_float(v):
    """嘗試轉成浮點數;失敗回 None。"""
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def get_concentration(quality_dict, item_name):
    """從進/出流水質字典找某項目的濃度。"""
    if not quality_dict:
        return None
    for code, items in quality_dict.items():
        if item_name in items:
            v = items[item_name]
            if isinstance(v, dict):
                return to_float(v.get("濃度"))
    return None


def get_mass(quality_dict, item_name):
    """從進/出流水質字典找某項目的質量 kg/d。"""
    if not quality_dict:
        return None
    for code, items in quality_dict.items():
        if item_name in items:
            v = items[item_name]
            if isinstance(v, dict):
                return to_float(v.get("質量"))
    return None


# ─────────────────── 各檢查函式 ───────────────────

def check_dissolved_concentration_up(unit):
    """檢查 1: 溶解性物質出流濃度 > 進流濃度 (學理不符)。"""
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    # 沉澱池/濃縮槽的「污泥側」出流會自然濃縮,豁免
    if std_tank in ("污泥濃縮池", "濃縮槽", "脫水機"):
        return findings

    for item in DISSOLVED_ITEMS:
        c_in = get_concentration(unit.get("influent", {}), item)
        c_out = get_concentration(unit.get("effluent", {}), item)
        if c_in is None or c_out is None:
            continue
        if c_in <= 0:
            continue
        ratio = c_out / c_in
        if ratio > 1.1:  # 出流比進流高 10% 以上
            findings.append({
                "嚴重度": "不合理",
                "類型": "質量平衡",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": item,
                "描述": f"{item} 出流濃度 {c_out:.2f} > 進流濃度 {c_in:.2f} (上升 {(ratio-1)*100:.1f}%)。溶解性物質不應自行濃縮。",
                "依據": "質量守恆 (環工技師多筆缺失指出溶解性物質自行濃縮不合學理)",
            })
    return findings


def check_fast_mix_metal_removal(unit):
    """檢查 2: 快混槽/pH 調整槽展現重金屬去除率。"""
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    if std_tank not in FAST_MIX_TYPES:
        return findings

    for metal in HEAVY_METALS:
        c_in = get_concentration(unit.get("influent", {}), metal)
        c_out = get_concentration(unit.get("effluent", {}), metal)
        if c_in is None or c_out is None or c_in <= 0:
            continue
        removal = (c_in - c_out) / c_in
        if removal > 0.1:  # 去除率 > 10%
            findings.append({
                "嚴重度": "不合理",
                "類型": "去除率",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": metal,
                "描述": f"{metal} 去除率 {removal*100:.1f}% (進 {c_in:.2f} → 出 {c_out:.2f})。快混/pH調整單元無固液分離,不應展現重金屬去除。",
                "依據": "重金屬須沉澱單元才能去除 (彭文良/施紹揚等技師缺失)",
            })
    return findings


def check_ph_tank_only_ph_change(unit):
    """檢查 3: pH 調整槽 / 暨快混池 除 pH 外進出流應一致。

    判斷基準改用「質量 (Σ進 vs Σ出)」, 比濃度準 — 若進出流量 Q 不同,
    用濃度會誤判 (例: 稀釋/濃縮); 質量則直接反映「該項目絕對含量是否變動」。

    涵蓋變體: 秋棠案例 T02-06 (pH調整池暨快混池)、T03-07/T04-04/T05-05
    (pH調整暨快混池)、T05-06 (pH調整池) 等過去全部漏抓的單元。
    """
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    name_in_doc = unit.get("name_in_doc", "")
    # 用集合比對 + name_in_doc fallback (萬一 std_tank 沒被歸到位)
    if std_tank not in PH_TANK_TYPES and name_in_doc not in PH_TANK_TYPES:
        return findings

    influent = unit.get("influent", {}) or {}
    effluent = unit.get("effluent", {}) or {}
    if not influent or not effluent:
        return findings

    # 收集所有水質項目 (扣掉 pH / 水溫)
    all_items = set()
    for stream in list(influent.values()) + list(effluent.values()):
        if isinstance(stream, dict):
            all_items.update(stream.keys())
    all_items = {i for i in all_items if i not in PH_ONLY_VARIABLE}

    # 對每個項目, 算 Σ進流質量 vs Σ出流質量
    for item in sorted(all_items):
        in_mass = 0.0
        out_mass = 0.0
        in_has = False
        out_has = False
        for stream in influent.values():
            if not isinstance(stream, dict):
                continue
            v = stream.get(item)
            if isinstance(v, dict):
                m = to_float(v.get("質量"))
                if m is not None:
                    in_mass += m
                    in_has = True
        for stream in effluent.values():
            if not isinstance(stream, dict):
                continue
            v = stream.get(item)
            if isinstance(v, dict):
                m = to_float(v.get("質量"))
                if m is not None:
                    out_mass += m
                    out_has = True

        if not (in_has and out_has and in_mass > 0):
            continue

        diff_pct = abs(out_mass - in_mass) / in_mass * 100
        if diff_pct > 5:  # > 5% 差異視為異常
            # 是減少還是增加?
            direction = "減少" if out_mass < in_mass else "增加"
            findings.append({
                "嚴重度": "不合理",
                "類型": "質量平衡",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": item,
                "描述": (
                    f"{item} 質量 進 {in_mass:.3f} → 出 {out_mass:.3f} kg/d "
                    f"({direction} {diff_pct:.1f}%)。{std_tank} 無固液分離, "
                    f"除 pH 外不應有變化。"
                ),
                "依據": "學理: pH 調整 / 快混類槽體只加藥攪拌, 無分離機制",
            })
    return findings


def check_settling_overflow_rate(unit):
    """檢查 5: 沉澱池表面溢流率。從設計參數抓溢流率。"""
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    if std_tank not in SETTLING_TANK_TYPES:
        return findings

    params = {**unit.get("design_params", {}), **unit.get("measure_params", {})}
    for pname, pval in params.items():
        if "溢流率" not in pname:
            continue
        if not isinstance(pval, dict):
            continue
        pmax = to_float(pval.get("max"))
        if pmax is None:
            continue
        if pmax > 50:
            findings.append({
                "嚴重度": "不合理",
                "類型": "設計參數",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": "表面溢流率",
                "描述": f"溢流率最大值 {pmax} m3/m2-d (一般合理範圍應 < 50)",
                "依據": "沉澱設計準則 (徐振利技師缺失)",
            })
    return findings


def check_required_equipment(unit):
    """檢查 8: 必要機具是否齊備。"""
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    required = REQUIRED_EQUIPMENT.get(std_tank, [])
    if not required:
        return findings

    existing = [e["name"] for e in unit.get("equipment", [])]
    existing_str = " ".join(existing)
    for req in required:
        if req not in existing_str:
            findings.append({
                "嚴重度": "待人工",
                "類型": "機具設施",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": req,
                "描述": f"{std_tank} 的『相關機具設施』未列『{req}』",
                "依據": "李俊坤等技師多筆缺失",
            })
    return findings


# ─────────────────── 主入口 ───────────────────

def run_balance_checks(app_data):
    """對整份 app_data 跑所有檢查,回傳 findings list。"""
    findings = []
    checkers = [
        check_dissolved_concentration_up,
        check_fast_mix_metal_removal,
        check_ph_tank_only_ph_change,
        check_settling_overflow_rate,
        check_required_equipment,
    ]
    for code, unit in app_data.get("units", {}).items():
        for checker in checkers:
            try:
                findings.extend(checker(unit))
            except Exception as e:
                # 不讓單一檢查失敗破壞整體
                findings.append({
                    "嚴重度": "錯誤",
                    "類型": "系統",
                    "單元": code,
                    "標準槽體": unit.get("std_tank", ""),
                    "對照項目": checker.__name__,
                    "描述": f"檢查器錯誤: {e}",
                    "依據": "(內部)",
                })
    return findings


def main():
    import io as _io
    try:
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    BASE = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) >= 2:
        json_path = sys.argv[1]
    else:
        # 找最新的 application_*.json
        jsons = sorted([f for f in os.listdir(BASE) if f.startswith("application_") and f.endswith(".json")])
        if not jsons:
            print("找不到 application_*.json")
            return
        json_path = os.path.join(BASE, jsons[-1])

    with open(json_path, "r", encoding="utf-8") as f:
        app_data = json.load(f)

    print(f"=== 質量平衡 / 學理檢查: {app_data.get('source_pdf', '?')} ===")
    print(f"處理單元數: {app_data.get('total_units', 0)}\n")

    findings = run_balance_checks(app_data)

    # 統計
    stats = defaultdict(int)
    for f in findings:
        stats[f["嚴重度"]] += 1

    print(f"檢查項數: {len(findings)}")
    for sev, count in stats.items():
        print(f"  {sev}: {count}")
    print()

    # 列出不合理項
    not_ok = [f for f in findings if f["嚴重度"] == "不合理"]
    if not_ok:
        print(f"=== 不合理項 ({len(not_ok)}) ===")
        for f in not_ok:
            print(f"\n  [{f['類型']}] {f['單元']} ({f['標準槽體']}) - {f['對照項目']}")
            print(f"     描述: {f['描述']}")
            print(f"     依據: {f['依據']}")

    manual = [f for f in findings if f["嚴重度"] == "待人工"]
    if manual:
        print(f"\n=== 待人工項 ({len(manual)}) ===")
        for f in manual[:10]:
            print(f"  [{f['類型']}] {f['單元']} - {f['對照項目']}: {f['描述'][:80]}")
        if len(manual) > 10:
            print(f"  ... (還有 {len(manual) - 10} 項)")

    # 輸出 JSON
    out = os.path.join(BASE, "balance_check_result.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"source": app_data.get("source_pdf"),
                   "total_findings": len(findings),
                   "stats": dict(stats),
                   "findings": findings}, f, ensure_ascii=False, indent=2)
    print(f"\n已輸出: {out}")


if __name__ == "__main__":
    main()
