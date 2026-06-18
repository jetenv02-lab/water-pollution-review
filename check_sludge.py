# -*- coding: utf-8 -*-
"""污泥質量平衡 — 推算固形物濃度與含水率合理性。

學理:
    - 一般廢水進濃縮槽前固形物: 0.5~2%
    - 濃縮後 / 終沉池排泥固形物: 2~6%
    - 脫水機進泥固形物: 2~6%
    - 脫水後濾餅含水率: 70~85% (即固形物 15~30%)

關鍵推算:
    Solids (kg/d) = Q_sludge (CMD) × Solids% × 10
       (Q × 1000 L/m³ × Solids% / 100 = Q × 10 × Solids%)

    反過來:
    Solids% = Solids (kg/d) / (Q × 10)

異常範例 (邑昇案 頁 94):
    脫水機進泥 4.3 CMD, 脫水後 2569.8 kg/d 固形物 (80% 含水率回推)
    → 推進泥固形物% = 2569.8 / (4.3 × 10) = 59.7%
    → 但一般進泥僅 1~5%
    → 嚴重不合理: 數字填錯
"""
import re


def _to_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        m = re.search(r"\d+(?:\.\d+)?", str(v))
        if m:
            try:
                return float(m.group())
            except ValueError:
                return None
        return None


# 含水率 / 固形物範圍 (學理)
SLUDGE_RANGES = {
    # 槽體類型: (進泥固形物%_min, 進泥固形物%_max, 說明)
    "脫水機":   (1.0, 8.0,   "脫水機進泥固形物應在 1~8% 範圍 (一般 2~6%)"),
    "濃縮槽":   (0.3, 3.0,   "濃縮槽進泥固形物應在 0.3~3% (一般 0.5~2%)"),
    "污泥儲槽": (0.3, 8.0,   "污泥儲槽固形物 0.3~8%"),
    "沉澱池":   (0.1, 2.0,   "沉澱池排泥固形物 0.1~2%"),
    "沉降池":   (0.1, 2.0,   "沉降池排泥固形物 0.1~2%"),
}

# 脫水後濾餅含水率 (脫水機特定)
DEWATERED_CAKE_MOISTURE_RANGE = (70, 90)  # % 含水率
DEWATERED_CAKE_SOLIDS_RANGE = (10, 30)   # % 固形物 (= 100 - 含水率)


def get_solids_concentration_percent(unit):
    """嘗試從 unit 抽出「固形物濃度%」(脫水機/濃縮槽常用的關鍵參數)。

    PDF 可能的命名: 含水率 / 含水率(%) / 固形物 / 固形物%
    """
    name_in_doc = unit.get("name_in_doc", "")
    measure_params = unit.get("measure_params") or {}
    design_params = unit.get("design_params") or {}
    all_params = {**design_params, **measure_params}

    moisture = None  # 含水率%
    solids = None    # 固形物%

    for pname, pval in all_params.items():
        if not isinstance(pval, dict):
            continue
        if "含水率" in pname or "含水量" in pname:
            v = _to_float(pval.get("max")) or _to_float(pval.get("min"))
            if v is not None:
                moisture = v
        if "固形物" in pname or "TS" in pname.upper():
            v = _to_float(pval.get("max")) or _to_float(pval.get("min"))
            if v is not None:
                solids = v

    # 互推: 含水率 80% → 固形物 20%
    if solids is None and moisture is not None:
        solids = 100 - moisture
    return solids


def check_sludge_balance(unit, app_data):
    """單元層級污泥質量平衡檢查。"""
    findings = []
    code = unit.get("raw_code") or "?"
    std = unit.get("std_tank") or ""
    name = unit.get("name_in_doc") or ""

    # 只看污泥相關單元
    is_sludge = (
        std in SLUDGE_RANGES
        or any(kw in name for kw in ["污泥", "脫水", "濃縮"])
    )
    if not is_sludge:
        return findings

    # 1) 檢查固形物濃度是否合理
    solids_pct = get_solids_concentration_percent(unit)
    if solids_pct is not None and solids_pct > 0:
        # 找對應範圍
        rng = SLUDGE_RANGES.get(std)
        if not rng:
            for kw, r in SLUDGE_RANGES.items():
                if kw in name or kw in std:
                    rng = r
                    break
        if rng:
            smin, smax, desc = rng
            if solids_pct < smin or solids_pct > smax:
                direction = "過低" if solids_pct < smin else "過高"
                severity = "不合理" if (solids_pct > smax * 3 or solids_pct < smin / 3) else "待人工"
                findings.append({
                    "嚴重度": severity,
                    "類型": "質量平衡",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": "污泥固形物濃度",
                    "描述": (
                        f"{name}: 固形物濃度 = {solids_pct:.2f}%, {direction}。"
                        f"學理範圍 {smin}~{smax}%。{desc}"
                    ),
                    "依據": "環工設計準則: 污泥側固形物濃度範圍",
                })

    # 2) 脫水機: 反推進泥固形物 = 脫水後 kg/d / (進泥 Q × 10)
    # 也順便檢查脫水後濾餅含水率
    if "脫水" in (std + name):
        # 找有沒有「進泥 Q」(從 stream_q) 跟「脫水後固形物 kg/d」(從 measure_params)
        stream_q = unit.get("stream_q") or {}
        # 取最大 stream_q 為進泥流量 (脫水機通常單一進)
        max_q = 0
        for s_code, qres in stream_q.items():
            if isinstance(qres, dict) and qres.get("ok"):
                q = qres.get("q_cmd") or 0
                if q > max_q:
                    max_q = q

        # 找「脫水後固形物 / 污泥餅 / 餅量 / 排出量」kg/d
        cake_kg = None
        for pname, pval in (unit.get("measure_params") or {}).items():
            if not isinstance(pval, dict):
                continue
            if any(kw in pname for kw in ["餅", "脫水後", "排出"]):
                v = _to_float(pval.get("max"))
                if v and v > 100:  # 排除「kg/m³」之類小單位
                    cake_kg = v
                    break

        if max_q > 0 and cake_kg:
            # 反推進泥固形物%
            implied_solids_pct = cake_kg / (max_q * 10)
            if implied_solids_pct > 15:
                findings.append({
                    "嚴重度": "不合理",
                    "類型": "質量平衡",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": "脫水機進泥反推固形物",
                    "描述": (
                        f"進泥 Q = {max_q:.2f} CMD, 脫水後乾固形物 ≈ {cake_kg:.1f} kg/d → "
                        f"反推進泥固形物 = {implied_solids_pct:.1f}%。"
                        f"一般未濃縮污泥固形物約 1~5%, 你的值 > 15% 嚴重不合常理, "
                        f"請檢查 (a) 進泥 Q 是否填錯 (b) 餅量是否漏寫 / 多寫 (c) 含水率單位"
                    ),
                    "依據": (
                        "環工設計準則: 污泥固形物濃度 (一般 1~5%, 濃縮後 2~6%, "
                        "若反推 > 15% 表示資料矛盾)"
                    ),
                })
            elif implied_solids_pct < 0.05:
                findings.append({
                    "嚴重度": "提醒",
                    "類型": "數據異常",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": "脫水機進泥反推固形物",
                    "描述": (
                        f"進泥 Q = {max_q:.2f} CMD, 脫水後乾固形物 ≈ {cake_kg:.1f} kg/d → "
                        f"反推進泥固形物 = {implied_solids_pct:.3f}% (幾乎是純水)。"
                        f"進泥要不要這麼大量? 還是餅量填太低?"
                    ),
                    "依據": "環工設計準則: 進泥固形物應 ≥ 0.5%",
                })

    return findings


def run_sludge_checks(app_data):
    """跑所有污泥相關單元的質量平衡檢查。"""
    findings = []
    for code, unit in app_data.get("units", {}).items():
        try:
            findings.extend(check_sludge_balance(unit, app_data))
        except Exception as e:
            findings.append({
                "嚴重度": "錯誤",
                "類型": "系統",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "check_sludge_balance",
                "描述": f"檢查器錯誤: {e}",
                "依據": "(內部)",
            })
    return findings
