# -*- coding: utf-8 -*-
"""加藥機制學理檢查 — 鎳系/輕系/各系分流系統。

學理:
    不同廢水分流系統對應不同加藥組合 (環工設計準則):

    | 分流 | 應加 | 不該加 | 學理目的 |
    |---|---|---|---|
    | 鎳系 | NaOH (大量, pH 10.5~11) | PAC/Polymer (沒沉澱) | Ni(OH)₂ 沉澱 |
    | 輕系 | NaOH (少量) | PAC/Polymer (沒必要) | 中和 |
    | 各系 | PAC + Polymer | 大量 NaOH (除非中和) | 混凝絮凝 |

判定:
    1. 從單元名稱 (name_in_doc) 偵測屬於哪個分流系統
        - 含「鎳系」「鎳水洗」「化銅」「化學鎳」「電鍍」 → 鎳系
        - 含「輕系」「輕細」「輕研磨」「酸鹼」 → 輕系
        - 含「各系」「綜合」「混合」「總匯」 → 各系
    2. 從 measure_params 看實際加什麼藥
    3. 比對學理應該加什麼, 不一致就標 ⚠️

同義字標準化:
    「聶系」→「鎳系」 (常見筆誤)
"""

# 分流系統 → (該加藥劑關鍵字, 不該加藥劑關鍵字)
DOSING_RULES = {
    "鎳系": {
        "should": ["NaOH", "氫氧化鈉", "燒鹼", "液鹼", "Ca(OH)", "氫氧化鈣", "石灰"],
        "should_not": ["PAC", "Polymer", "聚氯化鋁"],
        "ph_min": 10.5,  # pH 至少要拉到 10.5 才能讓 Ni(OH)₂ 完全沉澱
        "說明": "鎳系廢水含高濃度 Ni²⁺, 學理上應加大量 NaOH 將 pH 拉到 10.5~11, "
                "讓 Ni²⁺ → Ni(OH)₂ 沉澱 (綠色污泥)。若加 PAC/Polymer 沒鹼, 鎳無法去除",
    },
    "輕系": {
        "should": ["NaOH", "氫氧化鈉", "燒鹼", "液鹼", "硫酸", "H2SO4", "HCl"],
        "should_not": ["PAC", "Polymer", "聚氯化鋁"],
        "ph_min": None,  # 中和到 6~8 即可
        "說明": "輕系廢水污染輕、水量大, 學理上只需酸鹼中和 (pH 6~8)。"
                "加 PAC/Polymer 是浪費 (污染物濃度不夠形成絮羽)",
    },
    "各系": {
        "should": ["PAC", "Polymer", "聚氯化鋁", "硫酸鋁", "FeCl3", "高分子", "助凝劑"],
        "should_not": [],  # NaOH 可同時加 (中和+混凝)
        "ph_min": None,
        "說明": "各系是綜合廢水, 主要污染是膠體 + 細小 SS。學理上應加 PAC 混凝 + Polymer "
                "助凝形成大絮羽, 讓後端沉澱池能去除。只加 NaOH 沒混凝就沉不下來",
    },
}

# 鎳系/輕系/各系 偵測關鍵字
SYSTEM_DETECT = [
    # (關鍵字 list, 系統名)
    (["鎳系", "鎳水洗", "化銅", "化學鎳", "鍍鎳"], "鎳系"),
    (["聶系"], "鎳系"),  # 常見筆誤 (聶 = 鎳)
    (["輕系", "輕細", "輕研磨", "輕污染"], "輕系"),
    (["各系", "綜合", "混合廢水", "總匯", "總混"], "各系"),
]


def detect_system(unit):
    """從單元名稱偵測屬於哪個分流系統。"""
    name = unit.get("name_in_doc", "") or ""
    for keywords, system in SYSTEM_DETECT:
        for kw in keywords:
            if kw in name:
                return system, kw
    return None, None


def has_chemical(measure_params, keywords):
    """檢查 measure_params 中是否含某類藥劑。"""
    if not measure_params:
        return False
    for pname in measure_params.keys():
        for kw in keywords:
            if kw in str(pname):
                return True
    return False


def get_ph_max(measure_params):
    """從 measure_params 找 pH 值的最大值。"""
    if not measure_params:
        return None
    for pname, pval in measure_params.items():
        if "pH" in str(pname):
            if isinstance(pval, dict):
                m = pval.get("max")
                try:
                    return float(m) if m is not None else None
                except (TypeError, ValueError):
                    return None
    return None


def check_dosing_chemistry(unit):
    """檢查單元的加藥機制是否符合分流系統學理。

    Returns: list of findings
    """
    findings = []
    system, kw = detect_system(unit)
    if not system:
        return findings  # 無法判斷分流系統

    code = unit.get("raw_code") or unit.get("code") or "?"
    std_tank = unit.get("std_tank", "")
    rule = DOSING_RULES.get(system)
    if not rule:
        return findings

    measure = unit.get("measure_params") or {}

    # 1. 名稱含「聶系」筆誤 → 提示
    if kw == "聶系":
        findings.append({
            "嚴重度": "待人工",
            "類型": "文件一致性",
            "單元": code,
            "標準槽體": std_tank,
            "對照項目": "分流系統命名",
            "描述": (
                f"單元名稱「{unit.get('name_in_doc')}」含「聶系」, 應為「鎳系」筆誤 "
                f"(聶 vs 鎳, 同音字常見錯誤)。建議統一為「鎳系」。"
            ),
            "依據": "業界慣例: 鎳系 (Nickel) 廢水分流",
        })

    # 2. 鎳系應加 NaOH, 但 measure 沒有任何鹼劑 → 違反
    if rule["should"]:
        has_should = has_chemical(measure, rule["should"])
        if not has_should:
            findings.append({
                "嚴重度": "不合理",
                "類型": "加藥機制",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": f"{system}應加藥劑",
                "描述": (
                    f"此單元屬於「{system}」, 學理上應加 {', '.join(rule['should'][:3])} 等。"
                    f"但 PDF 量測參數沒看到任何對應加藥。{rule['說明']}"
                ),
                "依據": f"學理: {system}分流加藥規則",
            })

    # 3. 鎳系不該加 PAC/Polymer, 但 measure 有 → 警告 (可能整合單元?)
    if rule["should_not"]:
        has_not = has_chemical(measure, rule["should_not"])
        if has_not:
            findings.append({
                "嚴重度": "待人工",
                "類型": "加藥機制",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": f"{system}不該加的藥劑",
                "描述": (
                    f"此單元屬於「{system}」, 學理上不該加 {', '.join(rule['should_not'][:3])}。"
                    f"但 PDF 看到對應加藥, 請確認是否誤加或單元歸類錯誤。{rule['說明']}"
                ),
                "依據": f"學理: {system}分流加藥規則",
            })

    # 4. 鎳系 pH 應拉到 10.5+ 才能沉澱
    if rule.get("ph_min"):
        ph_max = get_ph_max(measure)
        if ph_max is not None and ph_max < rule["ph_min"]:
            findings.append({
                "嚴重度": "不合理",
                "類型": "加藥機制",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": "鎳系 pH 操作值",
                "描述": (
                    f"此單元屬於「{system}」, pH 量測最高 {ph_max} < 學理需求 {rule['ph_min']}。"
                    f"Ni²⁺ 沉澱需 pH ≥ 10.5, 否則放流會超標。"
                ),
                "依據": f"學理: Ni(OH)₂ 完全沉澱需 pH ≥ {rule['ph_min']}",
            })

    return findings


def run_dosing_checks(app_data):
    """對所有單元跑加藥機制檢查。"""
    findings = []
    for code, unit in app_data.get("units", {}).items():
        try:
            findings.extend(check_dosing_chemistry(unit))
        except Exception as e:
            findings.append({
                "嚴重度": "錯誤",
                "類型": "系統",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "check_dosing_chemistry",
                "描述": f"檢查器錯誤: {e}",
                "依據": "(內部)",
            })
    return findings
