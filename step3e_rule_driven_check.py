# -*- coding: utf-8 -*-
"""Step 3e: 規則庫驅動的檢查器。

核心想法:
之前的 step3b/3d 寫死 5-6 種檢查, 只能涵蓋規則庫的 10% 左右,
所以結果都集中在「快混槽展現重金屬去除」這一類。

本檢查器改成「規則庫驅動」:
- 讀規則庫.xlsx 各槽體分頁的 299 筆規則
- 對申請文件每個單元, 用該標準槽體類型對應的規則, 逐條嘗試比對
- 用「關鍵字模糊比對」啟發式判斷規則是否觸發
- 觸發 → 列出該條規則 + 標記為「需審查」

對照項目關鍵字 (例如):
  pH        → 看單元 measure_params 是否有 pH 範圍, 範圍是否 < 規則建議
  停留時間  → 看 design_params 是否有停留時間數值
  液位計    → 看 equipment 是否含液位計
  排泥      → 看 equipment 是否含排泥
  攪拌機    → 看 equipment 是否含攪拌
  污泥迴流  → 看是否有相關設施
  質量平衡  → 看進出流質量差異
  有效位數  → 看數值的小數位
  ...

依賴: openpyxl
使用:
    from step3e_rule_driven_check import run_rule_driven_check
    findings = run_rule_driven_check(app_data)
"""
import csv
import os
import re
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
RULES_CSV = os.path.join(BASE, "rules_extracted.csv")


# ──────────────────────────────────────────────────
# 規則庫載入 (直接讀 rules_extracted.csv, 避免依賴本機才有的 xlsx)
# ──────────────────────────────────────────────────

def load_rules_by_tank():
    """讀 rules_extracted.csv → { 標準槽體名稱: [規則 dict, ...] }。

    跳過「狀態 = ?」(待討論) 的規則。
    同義字: 「對照項目」會多存一個 normalized 版本到 `對照項目_標準` 欄。
    """
    if not os.path.exists(RULES_CSV):
        return {}
    # 嘗試載入同義字 (失敗不致命, 退化成 identity)
    try:
        import step3f_synonyms
        synonym_normalize = step3f_synonyms.normalize
    except Exception:
        synonym_normalize = lambda x: x

    rules_by_tank = defaultdict(list)
    skipped = 0
    with open(RULES_CSV, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 狀態欄: 空白 / V → 跑, ? → 跳過
            status = (row.get("狀態") or "").strip()
            if status == "?":
                skipped += 1
                continue
            tank = (row.get("標準槽體名稱") or "未分類").strip()
            raw_compare = row.get("對照項目", "")
            # 把 CSV 欄位名稱對應到原本 xlsx 的欄位名稱
            rule = {
                "缺失ID": row.get("缺失ID", ""),
                "來源": row.get("來源", ""),
                "原文缺失": row.get("原文缺失", ""),
                "檢查類型": row.get("檢查類型", ""),
                "對照項目": raw_compare,
                "對照項目_標準": synonym_normalize(raw_compare),  # 同義字展開後的標準詞
                "規則": row.get("規則", ""),
                "比對位置": row.get("比對位置", ""),
                "判定邏輯": row.get("判定邏輯", ""),
                "技師姓名": row.get("技師姓名", ""),
                "序號": row.get("序號", ""),
                "原始槽體代號": row.get("原始槽體代號", ""),
                "狀態": status,
            }
            rules_by_tank[tank].append(rule)
    # 把跳過數記到模組級全域 (Streamlit UI 可讀)
    globals()["LAST_SKIPPED_COUNT"] = skipped
    return dict(rules_by_tank)


# 上次載入規則時, 因「狀態 = ?」被跳過的筆數
LAST_SKIPPED_COUNT = 0


def get_last_skipped_count():
    """回傳上次載入規則時被「狀態 = ?」跳過的筆數。"""
    return globals().get("LAST_SKIPPED_COUNT", 0)


# ──────────────────────────────────────────────────
# 啟發式檢查器 (對照項目 → 檢查函式)
# ──────────────────────────────────────────────────

PH_RANGE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[~～\-]\s*(\d+(?:\.\d+)?)")
NUMBER_RE = re.compile(r"(\d+(?:\.\d+)?)")


def get_param_range(params_dict, key_substring):
    """從 design_params 或 measure_params 找名稱含 key_substring 的參數,回傳 (min, max)。"""
    for pname, pval in params_dict.items():
        if key_substring in pname:
            if isinstance(pval, dict):
                lo = pval.get("min")
                hi = pval.get("max")
                if lo is not None and hi is not None:
                    return float(lo), float(hi), pname
    return None, None, None


def has_equipment(equipment_list, name_keyword):
    """檢查機具清單是否含某關鍵字。"""
    for eq in equipment_list:
        eq_name = str(eq.get("name", ""))
        if name_keyword in eq_name:
            return True
    return False


def get_concentration_diff(unit, item_name):
    """取得某水質項目的進出流濃度差 (進 - 出)。回 (c_in, c_out) 或 (None, None)。"""
    c_in = c_out = None
    for code, items in unit.get("influent", {}).items():
        if item_name in items:
            v = items[item_name]
            if isinstance(v, dict):
                try:
                    c_in = float(v.get("濃度", 0))
                except (TypeError, ValueError):
                    pass
                break
    for code, items in unit.get("effluent", {}).items():
        if item_name in items:
            v = items[item_name]
            if isinstance(v, dict):
                try:
                    c_out = float(v.get("濃度", 0))
                except (TypeError, ValueError):
                    pass
                break
    return c_in, c_out


# ──────────────────────────────────────────────────
# 規則啟發式檢查
# ──────────────────────────────────────────────────

def check_rule_against_unit(rule, unit):
    """嘗試把一條規則套到一個單元, 啟發式判斷是否「應重點審查」。

    回傳 dict 或 None:
        {
            "嚴重度": "待人工",
            "對照項目": "...",
            "描述": "...",
            "規則來源": rule.get("缺失ID") + 來源,
        }
    """
    target_item = (rule.get("對照項目") or "").strip()
    rule_text = (rule.get("規則") or "").strip()
    check_type = (rule.get("檢查類型") or "").strip()
    judgment = (rule.get("判定邏輯") or "").strip()
    deficiency_id = rule.get("缺失ID", "")
    source = rule.get("來源", "")

    if not target_item:
        return None

    code = unit.get("raw_code", "")
    std_tank = unit.get("std_tank", "")

    design_params = unit.get("design_params", {})
    measure_params = unit.get("measure_params", {})
    equipment_list = unit.get("equipment", [])

    # ── 1. pH 規則 ──
    if "pH" in target_item or "pH" in rule_text:
        # 取 unit 的 pH 範圍
        for pname, pval in {**measure_params, **design_params}.items():
            if "pH" in pname:
                lo = pval.get("min") if isinstance(pval, dict) else None
                hi = pval.get("max") if isinstance(pval, dict) else None
                if lo is None:
                    continue
                # 規則裡的 pH 門檻
                threshold_match = re.search(
                    r"pH[\s\S]{0,20}?(?:下限|>|>=|超過|大於)\s*([0-9.]+)",
                    rule_text + judgment
                )
                if threshold_match:
                    try:
                        thr = float(threshold_match.group(1))
                        if lo < thr:
                            return _make_finding(
                                "待人工", check_type, code, std_tank, target_item,
                                f"{pname} 下限 {lo} 規則建議 > {thr} ({rule_text[:60]})",
                                deficiency_id, source
                            )
                    except ValueError:
                        pass
                else:
                    # 未抽到具體門檻 → 至少列出讓使用者人工檢查
                    return _make_finding(
                        "待人工", check_type, code, std_tank, target_item,
                        f"申請文件 {pname} 範圍 {lo}~{hi}, 規則: {rule_text[:80]}",
                        deficiency_id, source
                    )

    # ── 2. 停留時間 / 有效容量 / 有效水深 等設計參數 ──
    for design_key in ["停留時間", "有效容量", "有效水深", "表面溢流率", "溢流率",
                       "污泥迴流率", "迴流率", "MLSS", "DO", "溶氧", "F/M", "食微比",
                       "濾速", "上升流速", "體積負荷", "有機負荷", "攪拌"]:
        if design_key in target_item or design_key in rule_text:
            lo, hi, pname = get_param_range({**design_params, **measure_params}, design_key)
            if pname:
                return _make_finding(
                    "待人工", check_type, code, std_tank, target_item,
                    f"申請文件 {pname} = {lo}~{hi}, 規則: {rule_text[:80]}",
                    deficiency_id, source
                )
            # 規則指該項目但單元沒登載 → 也標記
            elif design_key in target_item and design_key in ["停留時間", "有效容量"]:
                return _make_finding(
                    "待人工", check_type, code, std_tank, target_item,
                    f"申請文件未登載『{design_key}』, 規則: {rule_text[:80]}",
                    deficiency_id, source
                )

    # ── 3. 機具設施 (液位計/pH計/攪拌機/排泥/流量計/液位/污泥泵...) ──
    eq_keywords = ["液位計", "pH計", "攪拌機", "鼓風機", "流量計", "加藥機",
                   "污泥泵", "排泥", "反洗", "鼓風機", "電磁式流量計"]
    for eq_kw in eq_keywords:
        if eq_kw in target_item or eq_kw in rule_text:
            if not has_equipment(equipment_list, eq_kw):
                return _make_finding(
                    "待人工", "機具設施", code, std_tank, target_item,
                    f"單元機具清單未列 {eq_kw}, 規則: {rule_text[:80]}",
                    deficiency_id, source
                )

    # ── 4. 加藥量 ──
    if "加藥" in target_item or "加藥量" in rule_text or "藥品" in rule_text:
        # 找量測參數中含「加藥量」者
        for pname, pval in measure_params.items():
            if "加藥" in pname:
                if isinstance(pval, dict):
                    raw = pval.get("raw", "")
                    return _make_finding(
                        "待人工", check_type, code, std_tank, target_item,
                        f"{pname} = {raw[:50]}, 規則: {rule_text[:80]}",
                        deficiency_id, source
                    )

    # ── 5. 質量平衡 ──
    if "質量平衡" in check_type or "質量平衡" in rule_text or "質量" in target_item:
        # 規則牽涉質量平衡, 多半需人工檢視
        # 只在 unit 有進出流時才列, 避免假陽性
        if unit.get("influent") and unit.get("effluent"):
            return _make_finding(
                "待人工", "質量平衡", code, std_tank, target_item,
                f"質量平衡需檢驗, 規則: {rule_text[:80]}",
                deficiency_id, source
            )

    # ── 6. 去除率 / 重金屬 ──
    heavy_metals = ["銅", "鎳", "鋅", "鉛", "鎘", "鉻", "總鉻", "六價鉻",
                    "錫", "汞", "總汞", "砷", "鉬"]
    for metal in heavy_metals:
        if metal in target_item:
            c_in, c_out = get_concentration_diff(unit, metal)
            if c_in is not None and c_out is not None and c_in > 0:
                removal = (c_in - c_out) / c_in * 100
                return _make_finding(
                    "待人工", "去除率", code, std_tank, target_item,
                    f"{metal} 進{c_in:.2f}→出{c_out:.2f} (去除 {removal:.1f}%), 規則: {rule_text[:80]}",
                    deficiency_id, source
                )

    # ── 7. SS / 懸浮固體 / BOD / COD ──
    for water_item in ["SS", "懸浮固體", "BOD", "COD", "氨氮", "總氮", "總磷",
                       "硝酸鹽氮", "硼", "氟鹽", "氰化物", "油脂", "真色色度"]:
        if water_item in target_item:
            c_in, c_out = get_concentration_diff(unit, water_item)
            if c_in is None or c_out is None:
                # 找含「mg/L」後綴的版本
                for variant in [water_item + "（mg/L）", water_item + "(mg/L)"]:
                    c_in, c_out = get_concentration_diff(unit, variant)
                    if c_in is not None:
                        break
            if c_in is not None and c_out is not None:
                return _make_finding(
                    "待人工", check_type or "水質標準", code, std_tank, target_item,
                    f"{water_item} 進{c_in:.2f}→出{c_out:.2f}, 規則: {rule_text[:80]}",
                    deficiency_id, source
                )

    # ── 8. 操作條件 (反洗頻率、更換頻率) ──
    for op_kw in ["反洗", "更換頻率", "操作參數", "操作時間"]:
        if op_kw in target_item:
            # 看 measure_params 或 design_params 有沒有相關記錄
            for pname in {**measure_params, **design_params}.keys():
                if op_kw in pname:
                    return _make_finding(
                        "待人工", check_type or "操作條件", code, std_tank, target_item,
                        f"申請文件有相關參數 ({pname}), 規則: {rule_text[:80]}",
                        deficiency_id, source
                    )

    # 沒命中任何啟發式 → 不產生 finding
    return None


def _make_finding(severity, check_type, code, std_tank, target_item, desc, deficiency_id, source):
    return {
        "嚴重度": severity,
        "類型": check_type or "其他",
        "單元": code,
        "標準槽體": std_tank,
        "對照項目": target_item,
        "描述": desc,
        "依據": f"規則庫 {deficiency_id} ({source})",
    }


# ──────────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────────

def run_rule_driven_check(app_data, max_findings_per_unit=20):
    """規則庫驅動的全面審查。

    對申請文件每個單元, 拿該標準槽體類型的所有規則跑啟發式檢查。

    Args:
        app_data: step2_extract_v2 輸出的 JSON
        max_findings_per_unit: 每單元最多 finding 數 (避免轟炸)

    Returns:
        list of finding dict
    """
    rules_by_tank = load_rules_by_tank()
    if not rules_by_tank:
        return []

    findings = []
    for code, unit in app_data.get("units", {}).items():
        std_tank = unit.get("std_tank", "")
        # 取該槽體類型的規則 (含通用的「文件類」「現場設備類」)
        applicable_rules = []
        applicable_rules.extend(rules_by_tank.get(std_tank, []))
        # 不額外加文件類/現場設備類, 避免重複過多

        unit_findings = []
        seen_keys = set()  # 避免同一單元同一對照項目重複
        for rule in applicable_rules:
            try:
                f = check_rule_against_unit(rule, unit)
                if f is None:
                    continue
                key = (f["對照項目"], f["描述"][:50])
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                unit_findings.append(f)
            except Exception as e:
                # 不讓單一規則失敗影響整體
                pass

        # 限制每單元 finding 數量
        unit_findings = unit_findings[:max_findings_per_unit]
        findings.extend(unit_findings)

    return findings


if __name__ == "__main__":
    import io
    import json
    import sys

    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    jsons = sorted([f for f in os.listdir(BASE) if f.startswith("application_") and f.endswith(".json")])
    if not jsons:
        print("找不到 application_*.json")
        sys.exit(0)
    with open(os.path.join(BASE, jsons[-1]), "r", encoding="utf-8") as f:
        app_data = json.load(f)

    print(f"=== 規則庫驅動檢查: {app_data.get('source_pdf', '?')} ===")
    print(f"單元數: {app_data.get('total_units', 0)}")
    findings = run_rule_driven_check(app_data)
    print(f"\n總 finding 數: {len(findings)}")

    # 統計
    from collections import Counter
    types = Counter()
    for f in findings:
        types[(f["嚴重度"], f["類型"])] += 1
    print("\n=== 類型分布 ===")
    for k, v in sorted(types.items()):
        print(f"  {k[0]:6s} | {k[1]:20s} | {v}")

    # 展示前 10 筆
    print("\n=== 前 10 筆 finding ===")
    for i, f in enumerate(findings[:10], 1):
        print(f"\n{i}. [{f['嚴重度']}|{f['類型']}] {f['單元']} ({f['標準槽體']}) - {f['對照項目']}")
        print(f"   {f['描述'][:120]}")
        print(f"   依據: {f['依據']}")
