# -*- coding: utf-8 -*-
"""加藥規則 → 質量補償計算 (從規則庫 _加藥規則 讀, 非 hardcode)。

設計理念:
    - chemical_calc.py 用 hardcode 典型值 (v1 第一版)
    - 此模組改從 規則庫.xlsx 的 _加藥規則 讀 (v2)
    - chemical_calc.py 不動, 此模組純新增, 對外提供同名 API
    - tank_chemistry.py 之後 patch 時, 可選擇用哪個版本

讀的欄位:
    B 分流系統 / C 槽體類別 / D 藥劑 / E 加藥角色
    R(18) 典型劑量 (mg/L) / S(19) 引入水質項目 / T(20) 轉換係數

對外 API (跟 chemical_calc 一致, 可直接 swap):
    compute_chemical_mass(unit, item, q_cmd=None) → kg/d
    get_typical_dosing(std_tank) → dict[chem → mg/L]
    describe_chemical_contribution(unit, item, q_cmd=None) → str
"""
import os
import re
from functools import lru_cache

import openpyxl

DEFAULT_XLSX = os.path.join(os.path.dirname(__file__), "規則庫.xlsx")
SHEET_NAME = "_加藥規則"


def _split_multi(s):
    """處理多項目格式: 'a; b; c' → ['a', 'b', 'c']"""
    if not s:
        return []
    return [x.strip() for x in re.split(r"[;,，；]", str(s)) if x.strip()]


def _to_float(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


@lru_cache(maxsize=1)
def load_dosing_compensation(xlsx_path=None):
    """從 _加藥規則 讀「應加」藥劑 + 補償欄, 整理成可查表結構。

    回傳:
        {
            標準槽體名: [
                {
                    "drug": "PAC",
                    "system": "各系",
                    "mg_per_L": 100,
                    "items": ["懸浮固體（mg/L）", "氯離子"],
                    "coefs": [0.30, 0.05],
                },
                ...
            ],
            ...
        }
    """
    path = xlsx_path or DEFAULT_XLSX
    if not os.path.exists(path):
        return {}

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        print(f"[dosing_rules_loader] 讀 {path} 失敗: {e}")
        return {}

    if SHEET_NAME not in wb.sheetnames:
        return {}
    ws = wb[SHEET_NAME]

    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 20:
            continue
        system = row[1]   # B
        tank = row[2]     # C
        drug = row[3]     # D
        role = row[4]     # E
        mg_per_L = _to_float(row[17])  # R (idx 17 = col 18)
        items_str = row[18]              # S
        coefs_str = row[19]              # T

        if not tank or not drug:
            continue
        # 只處理「應加」(禁加/可加不算進補償)
        if str(role or "").strip() != "應加":
            continue
        if mg_per_L is None or mg_per_L <= 0:
            continue

        items = _split_multi(items_str)
        coefs = [_to_float(c) or 0.0 for c in _split_multi(coefs_str)]

        # 對齊 items 與 coefs 長度 (短的補 0)
        n = max(len(items), len(coefs))
        items = (items + [""] * n)[:n]
        coefs = (coefs + [0.0] * n)[:n]

        entry = {
            "drug": str(drug).strip(),
            "system": str(system or "").strip(),
            "mg_per_L": mg_per_L,
            "items": items,
            "coefs": coefs,
        }

        tank_key = str(tank).strip()
        out.setdefault(tank_key, []).append(entry)

    return out


# 槽體別名: PDF 上可能的槽體名 → _加藥規則 的標準名
TANK_ALIAS = {
    "pH調整槽": "pH調整池",
    "pH 調整槽": "pH調整池",
    "中和槽": "中和池",
    "快混槽": "快混池",
    "慢混槽": "慢混池",
    "氰系氧化槽": "氰系氧化池",
    "鉻系還原槽": "鉻系還原池",
    "硫化物沉澱槽": "硫化物沉澱池",
    "pH調整暨快混池": "快混池",       # 用 PAC 同
    "pH調整池暨快混池": "快混池",
    "混凝膠凝池": "混凝池",
    "混凝池": "混凝池",
}


def get_typical_dosing(std_tank):
    """回傳該標準槽體的藥劑 → mg/L (跟 chemical_calc API 相同)。"""
    db = load_dosing_compensation()
    key = TANK_ALIAS.get(std_tank, std_tank)
    entries = db.get(key, [])
    return {e["drug"]: e["mg_per_L"] for e in entries}


def compute_chemical_mass(unit, item, q_cmd=None):
    """計算該單元因加藥引入指定水質項目的質量 (kg/d)。

    Args:
        unit: 處理單元 dict (有 std_tank 欄)
        item: 水質項目名 (例: "懸浮固體（mg/L）")
        q_cmd: 進流總 Q (m³/d)。若 None 自動加總 stream_q 的 WTB。

    Returns:
        kg/d (float)
    """
    std_tank = unit.get("std_tank") if isinstance(unit, dict) else None
    if not std_tank:
        return 0.0

    db = load_dosing_compensation()
    key = TANK_ALIAS.get(std_tank, std_tank)
    entries = db.get(key, [])
    if not entries:
        return 0.0

    if q_cmd is None:
        sq = unit.get("stream_q") or {}
        q_cmd = 0.0
        for stream_code, info in sq.items():
            if stream_code.startswith("WTB") and isinstance(info, dict):
                q = info.get("q_cmd")
                if isinstance(q, (int, float)):
                    q_cmd += float(q)
    if not q_cmd or q_cmd <= 0:
        return 0.0

    total = 0.0
    for entry in entries:
        # 找該藥劑對該水質項目的轉換係數
        for it, coef in zip(entry["items"], entry["coefs"]):
            if it == item and coef > 0:
                # mg/L × m³/d / 1000 × 轉換係數 = kg/d
                total += entry["mg_per_L"] * q_cmd / 1000.0 * coef
    return total


def describe_chemical_contribution(unit, item, q_cmd=None):
    """描述加藥引入細節, 給 finding 用。

    回傳: str (可空字串), 例:
        "[加藥引入估算: PAC 100mg/L×0.3=710 → +710 kg/d]"
    """
    std_tank = unit.get("std_tank") if isinstance(unit, dict) else None
    if not std_tank:
        return ""

    db = load_dosing_compensation()
    key = TANK_ALIAS.get(std_tank, std_tank)
    entries = db.get(key, [])
    if not entries:
        return ""

    if q_cmd is None:
        sq = unit.get("stream_q") or {}
        q_cmd = 0.0
        for stream_code, info in sq.items():
            if stream_code.startswith("WTB") and isinstance(info, dict):
                q = info.get("q_cmd")
                if isinstance(q, (int, float)):
                    q_cmd += float(q)
    if not q_cmd or q_cmd <= 0:
        return ""

    parts = []
    total = 0.0
    for entry in entries:
        for it, coef in zip(entry["items"], entry["coefs"]):
            if it == item and coef > 0:
                added = entry["mg_per_L"] * q_cmd / 1000.0 * coef
                parts.append(f"{entry['drug']} {entry['mg_per_L']}mg/L×{coef}={added:.1f}")
                total += added

    if not parts:
        return ""
    return f"[加藥引入估算 {item}: {'; '.join(parts)} → +{total:.1f} kg/d]"


def _self_test():
    """跑跟 chemical_calc 一樣的 5 個測試, 確認從 Sheets 讀的結果一致。"""
    db = load_dosing_compensation()
    print(f"從 _加藥規則 讀到 {len(db)} 個標準槽體有加藥配方:")
    for k, v in db.items():
        chems = [(e["drug"], e["mg_per_L"]) for e in v]
        print(f"  {k}: {chems}")
    print()

    # Test 1: 快混池加 PAC 100 mg/L, Q = 23673 m³/d
    fake_unit = {
        "std_tank": "快混池",
        "stream_q": {"WTB01-01-1": {"q_cmd": 23673}},
    }
    ss = compute_chemical_mass(fake_unit, "懸浮固體（mg/L）")
    expected = 100 * 23673 / 1000 * 0.30
    print(f"Test 1 (快混池 PAC → SS): {ss:.2f} kg/d (預期 {expected:.2f})")
    assert abs(ss - expected) < 0.1

    # Test 2: 中和池 NaOH 不引入 SS
    n_unit = {"std_tank": "中和池", "stream_q": {"WTB": {"q_cmd": 1000}}}
    ss2 = compute_chemical_mass(n_unit, "懸浮固體（mg/L）")
    print(f"Test 2 (中和池 NaOH → SS): {ss2:.2f} kg/d (預期 0)")
    assert ss2 == 0.0

    # Test 3: 中和池 H2SO4 引入 SO4
    so4 = compute_chemical_mass(n_unit, "硫酸根")
    expected = 30 * 1000 / 1000 * 0.98
    print(f"Test 3 (中和池 H2SO4 → SO4): {so4:.2f} kg/d (預期 {expected:.2f})")
    assert abs(so4 - expected) < 0.1

    # Test 4: 描述
    desc = describe_chemical_contribution(fake_unit, "懸浮固體（mg/L）")
    print(f"Test 4 描述: {desc}")
    assert "PAC" in desc and "+710" in desc

    # Test 5: pH 槽 別名測試 - pH 調整池 含 Ca(OH)2 會引入 SS
    # (跟 hardcode v1 不同: v1 pH 槽只有 NaOH/H2SO4 不引 SS, v2 Sheets 多了鎳系 Ca(OH)2 → 會引)
    ph_unit = {"std_tank": "pH調整槽", "stream_q": {"WTB": {"q_cmd": 1000}}}
    ph_ss = compute_chemical_mass(ph_unit, "懸浮固體（mg/L）")
    # 預期: Ca(OH)2 60 mg/L × 0.50 = 30 kg/d
    print(f"Test 5 (pH調整槽 別名→pH調整池, Ca(OH)2 → SS): {ph_ss:.2f} kg/d (預期 30, 來自鎳系 Ca(OH)2)")
    assert ph_ss == 30.0

    print("\n[OK] 全部測試通過 — 從 Sheets 讀的補償計算等同 hardcode 版")


if __name__ == "__main__":
    _self_test()
