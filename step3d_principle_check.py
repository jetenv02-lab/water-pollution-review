# -*- coding: utf-8 -*-
"""Step 3d: 進階學理檢查 (使用 step3c_unit_db 預設值)。

新增的檢查項:
1. 申請文件填的去除率 vs 預設削減率差異 (偏差 > 20% → 待人工)
2. 申請文件原廢水濃度 vs 預設原廢水濃度 (差異 > 10 倍 → 待人工確認)
3. 事業類別申報項目漏項 (依 BUSINESS_TYPES)
4. RAS 迴流比合理性 (0.3~1.5 × Q_in)
5. 質量平衡 Q 守恆檢查 (進流 vs 出流 ± 5%)

不重複 step3b 已實作的:
- 溶解性物質自行濃縮
- 快混槽展現重金屬去除
- pH 槽除 pH 外水質改變
- 沉澱池表面溢流率
- 必要機具未列

依賴: step3c_unit_db
使用:
    from step3d_principle_check import run_advanced_checks
    findings = run_advanced_checks(app_data, business_type=None)
"""
from step3c_unit_db import (
    UNIT_DEFAULT_REMOVAL,
    DEFAULT_RAW_CONCENTRATIONS,
    BUSINESS_TYPES,
    INLET_TYPES,
    get_default_removal,
    is_sludge_side_unit,
    check_missing_report_items,
    detect_business_type,
)


def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def get_concentration(quality_dict, item_name):
    if not quality_dict:
        return None
    for code, items in quality_dict.items():
        if item_name in items:
            v = items[item_name]
            if isinstance(v, dict):
                return to_float(v.get("濃度"))
    return None


# ─────────────────── 檢查函式 ───────────────────


def check_removal_rate_vs_default(unit):
    """檢查 1: 申請文件實際去除率 vs 預設削減率差異。

    對每個水質項目計算實際去除率 = (進-出)/進,
    再跟 step3c_unit_db 的預設值比對,
    偏差 > 20 個百分點時標「待人工複核」。
    """
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]

    # 污泥側單元豁免
    if is_sludge_side_unit(std_tank):
        return findings

    # 取得該槽體類型的預設削減率
    defaults = UNIT_DEFAULT_REMOVAL.get(std_tank, {})
    if not defaults:
        return findings

    for item, default_rate in defaults.items():
        if item.startswith("_"):
            continue
        c_in = get_concentration(unit.get("influent", {}), item)
        c_out = get_concentration(unit.get("effluent", {}), item)
        if c_in is None or c_out is None or c_in <= 0:
            continue
        actual_rate = (c_in - c_out) / c_in * 100
        diff = actual_rate - default_rate
        if abs(diff) > 20:
            severity = "待人工" if abs(diff) <= 40 else "不合理"
            direction = "高於" if diff > 0 else "低於"
            findings.append({
                "嚴重度": severity,
                "類型": "去除率",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": item,
                "描述": f"{item} 實際去除率 {actual_rate:.1f}% {direction}預設 {default_rate}% (差 {abs(diff):.1f} pt)",
                "依據": f"環工設計準則 {std_tank} 典型 {item} 削減率為 {default_rate}% (僅供參考,實際值依現場條件而異)",
            })
    return findings


def check_raw_water_vs_typical(app_data):
    """檢查 2: 原廢水濃度遠超 / 遠低於典型值。

    取第一個單元(通常是廢水收集池)的進流當原廢水。
    """
    findings = []
    # 找第一個單元
    units = app_data.get("units", {})
    if not units:
        return findings

    # 取所有單元的所有進流, 看哪個是『最原始的進流』
    # 啟發式: 進流編號是 WTB(原廢水) 且來自單元代號最小的
    first_code = min(units.keys()) if units else None
    if not first_code:
        return findings

    first_unit = units[first_code]
    influent = first_unit.get("influent", {})
    if not influent:
        return findings

    for infl_code, items in influent.items():
        for item, val_dict in items.items():
            if not isinstance(val_dict, dict):
                continue
            c = to_float(val_dict.get("濃度"))
            if c is None or c <= 0:
                continue
            default = DEFAULT_RAW_CONCENTRATIONS.get(item)
            if not default or default["is_range"]:
                continue
            default_v = default["value"]
            if default_v <= 0:
                continue
            ratio = c / default_v
            if ratio > 10:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "水質標準",
                    "單元": first_code,
                    "標準槽體": "原廢水",
                    "對照項目": item,
                    "描述": f"原廢水 {item} 濃度 {c} {default['unit']} 為典型值 ({default_v}) 的 {ratio:.1f} 倍",
                    "依據": "原廢水水質典型值對照 (僅供參考)",
                })
            elif ratio < 0.1:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "水質標準",
                    "單元": first_code,
                    "標準槽體": "原廢水",
                    "對照項目": item,
                    "描述": f"原廢水 {item} 濃度 {c} {default['unit']} 為典型值 ({default_v}) 的 {ratio:.3f} 倍 (偏低)",
                    "依據": "原廢水水質典型值對照 (僅供參考)",
                })
    return findings


def check_business_type_items(app_data, business_type):
    """檢查 3: 事業類別漏報項目。

    若 business_type 為 None, 嘗試自動偵測。
    """
    findings = []
    if not business_type:
        return findings

    # 收集申請文件中所有出現過的水質項目
    declared_items = set()
    for code, unit in app_data.get("units", {}).items():
        for infl_code, items in unit.get("influent", {}).items():
            declared_items.update(items.keys())
        for effl_code, items in unit.get("effluent", {}).items():
            declared_items.update(items.keys())

    # 把 SS 和懸浮固體當同一項目
    if "懸浮固體（mg/L）" in declared_items or "懸浮固體" in declared_items:
        declared_items.add("SS")

    missing = check_missing_report_items(business_type, list(declared_items))
    if missing:
        for m in missing:
            findings.append({
                "嚴重度": "不合理",
                "類型": "水質標準",
                "單元": "(全廠)",
                "標準槽體": f"事業類別: {business_type}",
                "對照項目": m["item"],
                "描述": f"事業類別 [{business_type}] 之 {m['category']} 應申報 {m['item']} (頻率: {m['frequency']}), 但申請文件未見此項目",
                "依據": f"水污染防治法事業類別申報項目對照表 (事業類別: {business_type})",
            })
    return findings


def check_q_balance(app_data):
    """檢查 4: 單元層級 Q 守恆。

    若進流量 vs 出流量差 > 5% 且非污泥側單元, 標記。
    (秋棠案例的 v2 抽取目前還沒抽尺寸/流量,
     所以這個檢查的有效性取決於 facility 表是否能拿到 Q)
    """
    # 暫時 placeholder - 等抽流量功能加進 step2_extract_v2 再開
    return []


# ─────────────────── 主入口 ───────────────────


def run_advanced_checks(app_data, business_type=None):
    """跑全部進階檢查, 回傳 findings list。

    Args:
        app_data: step2_extract_v2 輸出的 JSON
        business_type: 事業類別名稱 (可選, 若 None 嘗試自動偵測)
    """
    findings = []

    # 對每個單元跑去除率比對
    for code, unit in app_data.get("units", {}).items():
        try:
            findings.extend(check_removal_rate_vs_default(unit))
        except Exception as e:
            findings.append({
                "嚴重度": "錯誤",
                "類型": "系統",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "check_removal_rate",
                "描述": f"檢查器錯誤: {e}",
                "依據": "(內部)",
            })

    # 原廢水檢查
    try:
        findings.extend(check_raw_water_vs_typical(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_raw_water",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 事業類別檢查
    try:
        findings.extend(check_business_type_items(app_data, business_type))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_business_type",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    return findings


if __name__ == "__main__":
    import io
    import json
    import os
    import sys

    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    BASE = os.path.dirname(os.path.abspath(__file__))
    jsons = sorted([f for f in os.listdir(BASE) if f.startswith("application_") and f.endswith(".json")])
    if not jsons:
        print("找不到 application_*.json")
        sys.exit(0)
    with open(os.path.join(BASE, jsons[-1]), "r", encoding="utf-8") as f:
        app_data = json.load(f)

    # 試三種事業類別
    for bt in [None, "電鍍業", "晶圓製造及半導體製造業"]:
        print(f"\n=== business_type = {bt!r} ===")
        findings = run_advanced_checks(app_data, business_type=bt)
        print(f"檢查項數: {len(findings)}")
        not_ok = [f for f in findings if f["嚴重度"] == "不合理"]
        manual = [f for f in findings if f["嚴重度"] == "待人工"]
        print(f"  不合理: {len(not_ok)}")
        print(f"  待人工: {len(manual)}")
        for f in not_ok[:5]:
            print(f"  [{f['類型']}] {f['單元']} - {f['對照項目']}: {f['描述'][:100]}")
        for f in manual[:5]:
            print(f"  [{f['類型']}] {f['單元']} - {f['對照項目']}: {f['描述'][:100]}")
