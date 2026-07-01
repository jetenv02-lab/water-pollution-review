# -*- coding: utf-8 -*-
"""進出水質「完全相同」偵測 — 廠商偷懶填表 (複製貼上) 偵測。

背景:
    部分廠商填水質表時, 直接把進流欄複製到出流欄, 導致 T01-XX 的
    進流/出流水質 20 個項目**完全相同** (連小數點都一樣)。
    質量平衡檢查看到「差 0%」以為 OK, 完全看不出來廠商沒實際計算。

判斷準則:
    - 只看「反應/分離」類槽體 (儲存/收集/放流本來就進=出)
    - 跳過 水溫 / pH (本來就應該幾乎不變)
    - 比對進出加權平均濃度
    - 若 >= 80% 項目「差 < 0.5%」→ 觸發提醒
    - 需至少 3 個項目才有意義

對外 API:
    check_identical_inout(unit) → list[finding]
    check_all_units_identical_inout(units) → list[finding]
    is_target_tank(std_tank) → bool

不動任何既有 code, 純新增模組。
整合位置: streamlit_app.py 主 findings 集結區。
"""

# ──────────────────────────────────────────────────
# 需要檢查的槽型 (反應/分離類)
# ──────────────────────────────────────────────────
TARGET_TANK_KEYWORDS = [
    # 反應類
    "快混", "慢混", "pH", "中和", "混凝", "膠凝",
    "氰系氧化", "鉻系還原", "氧化",
    # 分離類
    "化沉", "沉澱", "沉降", "浮除", "油脂分離",
    "砂濾", "活性碳", "離子交換",
    # 生物類
    "曝氣", "厭氧", "接觸氧化", "MBR",
    # 其他反應
    "預處理", "批次反應",
]

# 明確排除的槽型 (本來就進=出)
EXCLUDE_TANK_KEYWORDS = [
    "調勻", "貯留", "暫存", "廢水收集", "廢水調整",
    "放流",  # 放流池若進=出, 是「上游沒處理」的問題, 不是廠商填表問題
    "污泥儲", "濾液",
]

# 跳過比對的項目 (本來就幾乎不變)
SKIP_ITEMS = {
    "水溫", "水溫(攝氏)", "水溫（攝氏）", "溫度",
    "pH", "pH值", "PH", "ph",
}


def is_target_tank(std_tank):
    """判斷該槽體是否要跑進=出檢查。"""
    if not std_tank:
        return False
    s = str(std_tank)
    # 明確排除的先擋掉
    for bad in EXCLUDE_TANK_KEYWORDS:
        if bad in s:
            return False
    # 有 target 關鍵字
    for kw in TARGET_TANK_KEYWORDS:
        if kw in s:
            return True
    return False


def _to_float(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _get_stream_q(unit, stream_code):
    """從 unit['stream_q'][code]['q_cmd'] 拿 Q, 沒有回 None。"""
    sq = (unit.get("stream_q") or {}).get(stream_code)
    if not isinstance(sq, dict):
        return None
    return _to_float(sq.get("q_cmd"))


def _weighted_avg_concentrations(streams, unit):
    """算加權平均濃度。回傳 {item: 濃度平均}。

    若無 Q, fallback 用算術平均。
    """
    item_data = {}  # item -> [(c, q)]
    for stream_code, content in streams.items():
        if not isinstance(content, dict):
            continue
        q = _get_stream_q(unit, stream_code) or 1.0  # fallback
        for item, v in content.items():
            if item in SKIP_ITEMS:
                continue
            if isinstance(v, dict):
                c = _to_float(v.get("濃度"))
                if c is not None:
                    item_data.setdefault(item, []).append((c, q))

    out = {}
    for item, lst in item_data.items():
        q_sum = sum(q for _, q in lst)
        if q_sum <= 0:
            out[item] = sum(c for c, _ in lst) / len(lst)
        else:
            out[item] = sum(c * q for c, q in lst) / q_sum
    return out


def check_identical_inout(unit, min_ratio=0.90, threshold=0.005,
                          min_items=3, significant_change_pct=0.30):
    """檢查單一單元進=出情況。

    Args:
        unit: 處理單元 dict
        min_ratio: 相同項目比例門檻 (預設 90% — 較嚴, 避免誤判正常運作的沉澱池)
        threshold: 「完全相同」的相對差門檻 (預設 0.5%)
        min_items: 至少要有幾個項目才判斷 (預設 3)
        significant_change_pct: 顯著變動門檻 (預設 30%)
            若有任 1 個項目變動 > 30% (真正做了處理), 且項目數 ≥ 5,
            視為「該槽有實際運作」, 不觸發 finding

    Returns:
        list[finding] (最多 1 個 finding per unit)

    邏輯 (修訂 v2):
        1. 相同項目比例 < 90% → 不觸發 (代表有變動, 不是偷懶)
        2. 相同比例 ≥ 90% 但**只有 1~2 個項目變動** 且變動 > 30%
           → 若該槽本來就只該動這 1~2 項, 視為合理 (例: 氧化池只動 CN, 離子交換只動特定金屬)
           → 若共通項目 ≥ 5 且有顯著變動, 放行
        3. 100% 相同 或 相同比例 ≥ 90% 且無顯著變動 → 觸發提醒
    """
    std_tank = unit.get("std_tank") or ""
    code = unit.get("raw_code") or unit.get("code_id") or "?"

    if not is_target_tank(std_tank):
        return []

    influent = unit.get("influent") or {}
    effluent = unit.get("effluent") or {}
    if not influent or not effluent:
        return []

    in_avg = _weighted_avg_concentrations(influent, unit)
    out_avg = _weighted_avg_concentrations(effluent, unit)
    common_items = set(in_avg) & set(out_avg)

    if len(common_items) < min_items:
        return []

    identical = []
    different = []
    has_significant = False  # 是否有 1 項以上「顯著變動」
    for item in common_items:
        in_c = in_avg[item]
        out_c = out_avg[item]
        base = max(abs(in_c), abs(out_c))
        if base == 0:
            identical.append(item)
            continue
        rel_diff = abs(in_c - out_c) / base
        if rel_diff < threshold:
            identical.append(item)
        else:
            different.append((item, in_c, out_c))
            if rel_diff >= significant_change_pct:
                has_significant = True

    ratio = len(identical) / len(common_items)
    if ratio < min_ratio:
        return []

    # v2 新增: 若共通項目多 (≥ 5) 且有顯著變動的項目 (該槽確實在做處理), 放行
    # 例如: 沉澱池 SS 267→36 (-86%), 但其他 16 項不變 → 合理
    #       離子交換 銅 1.2→0.01 + 鎳 56→0.06, 其他不變 → 合理
    if len(common_items) >= 5 and has_significant:
        return []

    # 觸發 finding
    return [{
        "嚴重度": "提醒",  # 用「提醒」— 這是「疑似」不是硬性違規
        "類型": "文件一致性",
        "單元": code,
        "標準槽體": std_tank,
        "對照項目": "進出水質完全相同",
        "描述": (
            f"該單元進出水質 {len(identical)}/{len(common_items)} 項"
            f" ({ratio*100:.0f}%) 完全相同 (差異 < 0.5%), "
            f"廠商可能未實際計算加藥/處理後之出流水質, "
            f"僅將進流欄複製至出流欄。"
            + (f" 實際有變動的項目: "
               + ", ".join(f"{it}({inv:.2f}→{outv:.2f})" for it, inv, outv in different[:3])
               + ("..." if len(different) > 3 else "")
               if different else " (完全 100% 相同, 沒有任何項目變動)")
            + f" 建議廠商依 {std_tank} 之學理重新填寫。"
        ),
        "依據": "進出水質相同度自動偵測 (廠商填表完整性檢查)",
    }]


def check_all_units_identical_inout(units, min_ratio=0.80, threshold=0.005):
    """批次檢查所有單元。

    Args:
        units: dict[code → unit] 或 list[unit]

    Returns:
        list[finding]
    """
    findings = []
    if isinstance(units, dict):
        unit_iter = units.values()
    else:
        unit_iter = units

    for unit in unit_iter:
        if not isinstance(unit, dict):
            continue
        findings.extend(check_identical_inout(unit, min_ratio, threshold))
    return findings


def _self_test():
    """自我測試。"""
    # Test 1: 秋棠 T03-07 型 - 完全相同的 pH 調整暨快混池
    fake_unit_1 = {
        "std_tank": "pH調整暨快混池",
        "raw_code": "T03-07",
        "influent": {
            "WTB03-07-1": {
                "銅": {"濃度": 50.2, "質量": 5.0},
                "鎳": {"濃度": 8.3, "質量": 0.8},
                "COD": {"濃度": 300, "質量": 30},
                "SS": {"濃度": 120, "質量": 12},
                "pH": {"濃度": 7.5},
            }
        },
        "effluent": {
            "WTA03-07-1": {
                "銅": {"濃度": 50.2, "質量": 5.0},
                "鎳": {"濃度": 8.3, "質量": 0.8},
                "COD": {"濃度": 300, "質量": 30},
                "SS": {"濃度": 120, "質量": 12},
                "pH": {"濃度": 7.5},
            }
        },
    }
    fs1 = check_identical_inout(fake_unit_1)
    print(f"Test 1 (pH+快混 100% 相同): {len(fs1)} finding")
    for f in fs1:
        print(f"  {f['嚴重度']} - {f['對照項目']}: {f['描述'][:200]}")
    assert len(fs1) == 1

    # Test 2: 正常沉澱池 (SS 大幅下降)
    fake_unit_2 = {
        "std_tank": "沉澱池",
        "raw_code": "T01-05",
        "influent": {"WTB": {"SS": {"濃度": 500}, "COD": {"濃度": 300}, "銅": {"濃度": 5}}},
        "effluent": {"WTA": {"SS": {"濃度": 20}, "COD": {"濃度": 100}, "銅": {"濃度": 0.5}}},
    }
    fs2 = check_identical_inout(fake_unit_2)
    print(f"\nTest 2 (沉澱池正常運作): {len(fs2)} finding (預期 0)")
    assert len(fs2) == 0

    # Test 3: 貯留槽 (排除清單, 應跳過)
    fake_unit_3 = {
        "std_tank": "貯留槽",
        "raw_code": "T01-10",
        "influent": {"WTB": {"銅": {"濃度": 10}, "SS": {"濃度": 100}, "COD": {"濃度": 200}}},
        "effluent": {"WTA": {"銅": {"濃度": 10}, "SS": {"濃度": 100}, "COD": {"濃度": 200}}},
    }
    fs3 = check_identical_inout(fake_unit_3)
    print(f"Test 3 (貯留槽本來就進=出, 應排除): {len(fs3)} finding (預期 0)")
    assert len(fs3) == 0

    # Test 4: 快混池但只有 2 個項目 (< min_items)
    fake_unit_4 = {
        "std_tank": "快混槽",
        "raw_code": "T01-01",
        "influent": {"WTB": {"銅": {"濃度": 5}, "SS": {"濃度": 100}}},
        "effluent": {"WTA": {"銅": {"濃度": 5}, "SS": {"濃度": 100}}},
    }
    fs4 = check_identical_inout(fake_unit_4)
    print(f"Test 4 (快混池但只 2 項, < min): {len(fs4)} finding (預期 0)")
    assert len(fs4) == 0

    # Test 5: 有變動 (混合案例)
    fake_unit_5 = {
        "std_tank": "快混槽",
        "raw_code": "T01-02",
        "influent": {"WTB": {"銅": {"濃度": 5}, "鎳": {"濃度": 3}, "SS": {"濃度": 100},
                             "COD": {"濃度": 200}, "pH": {"濃度": 7}}},
        # 除了 SS (加 PAC 上升), 其他全部相同
        "effluent": {"WTA": {"銅": {"濃度": 5}, "鎳": {"濃度": 3}, "SS": {"濃度": 150},
                             "COD": {"濃度": 200}, "pH": {"濃度": 7}}},
    }
    fs5 = check_identical_inout(fake_unit_5)
    # 4 項共, 3 個相同 (COD, 銅, 鎳), 1 個變動 (SS)  → 75% < 80% → 不觸發
    # 注意: pH 已被排除, 所以 common_items = 4 (銅, 鎳, SS, COD)
    print(f"Test 5 (快混池 3/4=75%, 未達 80%): {len(fs5)} finding")
    # 這裡實際 3/4 = 75% < 80%, 應該不觸發
    assert len(fs5) == 0, f"預期 0 但得到 {len(fs5)}"

    print("\n[OK] 5 個自我測試全過")


if __name__ == "__main__":
    _self_test()
