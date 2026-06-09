# -*- coding: utf-8 -*-
"""放流水標準檢查模組。

從 discharge_standards.csv 讀環境部公告的各業別放流水標準,
檢查申請文件的「放流口」水質是否符合該業別的限值。

放流口識別:
    - 單元 std_tank == "放流池"
    - 或 出流編號是 "Dxx" (放流口代號)
    - 或 name_in_doc 含「放流」

業別識別:
    - 使用者在 streamlit selectbox 選的 business_type
    - 若沒選, 預設套「一般業」

注意:
    這些標準依環境部「水污染防治法施行細則」公告, 細節依「業別放流水
    標準」表附件而定。本表只列常見項目, 若有特殊業別或新法規, 請編輯
    discharge_standards.csv (Faye/同事可雲端協作)。
"""
import os
import csv

CSV_PATH = os.path.join(os.path.dirname(__file__), "discharge_standards.csv")


def _to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def load_standards(csv_path=None):
    """讀 discharge_standards.csv, 回傳 dict[業別 → dict[水質項目 → (限值, 單位)]]。

    pH 特殊處理: 表中 "pH" = 上限, "pH_min" = 下限
    """
    if csv_path is None:
        csv_path = CSV_PATH
    if not os.path.exists(csv_path):
        return {}

    standards = {}
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            business = row.get("業別", "").strip()
            if not business or business.startswith("#"):
                continue
            item = row.get("水質項目", "").strip()
            if not item:
                continue
            max_val = _to_float(row.get("最大限值"))
            if max_val is None:
                continue
            unit = row.get("單位", "").strip()
            note = row.get("備註", "").strip()

            standards.setdefault(business, {})[item] = {
                "max": max_val,
                "unit": unit,
                "note": note,
            }
    return standards


def get_business_standard(business_type, standards=None):
    """取得某業別的放流水標準, 若沒對應就退用「一般」。"""
    if standards is None:
        standards = load_standards()
    if business_type and business_type in standards:
        return standards[business_type]
    # 退用「一般」
    return standards.get("一般", {})


def find_discharge_units(app_data):
    """找出申請文件中的「放流口」單元。

    判定: std_tank=="放流池" 或 name_in_doc 含「放流」字眼。
    """
    discharge = []
    for code, unit in app_data.get("units", {}).items():
        std = (unit.get("std_tank") or "").strip()
        name = unit.get("name_in_doc", "")
        if std == "放流池" or "放流" in name:
            discharge.append(code)
    return discharge


def check_discharge_standard(app_data, business_type=None):
    """檢查放流口水質是否符合該業別放流水標準。

    對每個放流口單元的「出流水質」每個項目, 比對該業別的最大限值。
    超標 → 🔴 違規 (法規確定)。

    Returns: list of findings
    """
    findings = []
    standards = load_standards()
    if not standards:
        return findings

    std_map = get_business_standard(business_type, standards)
    if not std_map:
        return findings

    pH_max = std_map.get("pH", {}).get("max")
    pH_min = std_map.get("pH_min", {}).get("max")

    discharge_codes = find_discharge_units(app_data)
    if not discharge_codes:
        return findings

    business_label = business_type if business_type and business_type in standards else "一般"

    for code in discharge_codes:
        unit = app_data["units"][code]
        # 看出流, 若沒出流就看進流 (放流口可能只標進流)
        streams = unit.get("effluent") or unit.get("influent") or {}
        for stream_code, items in streams.items():
            if not isinstance(items, dict):
                continue
            for item_name, val_dict in items.items():
                if not isinstance(val_dict, dict):
                    continue
                # 1. pH 特殊處理
                if "pH" in item_name and pH_max:
                    rng = val_dict.get("範圍") or val_dict.get("濃度")
                    if not rng:
                        continue
                    # 解析 "6~9" 之類
                    import re
                    m = re.match(r"\s*(\d+(?:\.\d+)?)\s*[~～]\s*(\d+(?:\.\d+)?)", str(rng))
                    if m:
                        lo, hi = float(m.group(1)), float(m.group(2))
                        if hi > pH_max:
                            findings.append({
                                "嚴重度": "不合理",
                                "類型": "水質標準",
                                "單元": code,
                                "標準槽體": "放流池",
                                "對照項目": f"pH ({stream_code})",
                                "描述": (
                                    f"放流口 {stream_code} pH 上限 {hi} > {business_label}業放流水標準 {pH_max} "
                                    f"(法規限值)"
                                ),
                                "依據": f"水污染防治法施行細則 - {business_label}業放流水標準",
                            })
                        if pH_min and lo < pH_min:
                            findings.append({
                                "嚴重度": "不合理",
                                "類型": "水質標準",
                                "單元": code,
                                "標準槽體": "放流池",
                                "對照項目": f"pH ({stream_code})",
                                "描述": (
                                    f"放流口 {stream_code} pH 下限 {lo} < {business_label}業放流水標準 {pH_min} "
                                    f"(法規限值)"
                                ),
                                "依據": f"水污染防治法施行細則 - {business_label}業放流水標準",
                            })
                    continue

                # 2. 一般項目 — 比對最大限值
                # 統一項目名稱 (例如「懸浮固體（mg/L）」→ 「懸浮固體」)
                item_norm = re.sub(r"[（(].*?[）)]", "", item_name).strip()
                # 在 std_map 找對應 (試多個別名)
                std_entry = None
                for try_name in (item_name, item_norm,
                                 item_norm.replace("化學需氧量", "COD"),
                                 item_norm.replace("生化需氧量", "BOD"),
                                 item_norm.replace("懸浮固體", "SS")):
                    if try_name in std_map:
                        std_entry = std_map[try_name]
                        break
                if not std_entry:
                    continue

                # 拿濃度數值
                conc = _to_float(val_dict.get("濃度"))
                if conc is None or conc <= 0:
                    continue
                max_val = std_entry["max"]
                if conc > max_val:
                    findings.append({
                        "嚴重度": "不合理",
                        "類型": "水質標準",
                        "單元": code,
                        "標準槽體": "放流池",
                        "對照項目": f"{item_norm} ({stream_code})",
                        "描述": (
                            f"放流口 {stream_code} {item_norm} = {conc} {std_entry['unit']} "
                            f"超過 {business_label}業放流水標準 {max_val} {std_entry['unit']} "
                            f"({conc/max_val:.1f} 倍)"
                            + (f" · {std_entry['note']}" if std_entry.get("note") else "")
                        ),
                        "依據": f"水污染防治法施行細則 - {business_label}業放流水標準",
                    })
    return findings


if __name__ == "__main__":
    # 自我測試
    import sys, io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    s = load_standards()
    print(f"載入 {len(s)} 個業別的標準")
    for biz, items in s.items():
        print(f"\n=== {biz} ({len(items)} 項目) ===")
        for k, v in list(items.items())[:5]:
            print(f"  {k}: {v['max']} {v['unit']}")
