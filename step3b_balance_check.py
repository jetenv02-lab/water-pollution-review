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

# ── pH 槽 / 快混池 學理分類 ──
# 兩種槽體的加藥方式不同, 學理上允許變動的水質項目也不同:
#
# 1. 純 pH 調整槽: 只加酸/鹼 (HCl, NaOH, Ca(OH)₂)
#    → 學理上「只有 pH 跟水溫」會變, 其他水質 (SS, 重金屬, COD, ...) 都不該變
#
# 2. pH 調整暨快混池: 加酸/鹼 + 混凝劑 (PAC, alum, FeCl₃)
#    → pH 會變 + 混凝劑本身是固體, 所以 SS 會增加 (合理)
#    → 但仍無固液分離, 重金屬/COD/BOD 等不應減少 (需後端沉澱池才能去除)

# 純 pH 調整槽 — 只允許 pH / 水溫變動
PURE_PH_TANK_TYPES = {
    "pH調整槽",
    "pH調整池",
    "中和池",  # 學理上跟純 pH 槽同類 (只加酸鹼)
}
PURE_PH_ALLOW_VARY = {"pH值", "pH", "水溫(攝氏)", "水溫"}

# pH 調整暨快混池 — pH/水溫 + SS 都允許變動 (因加了混凝劑)
PH_FAST_MIX_TYPES = {
    "pH調整暨快混池",      # 秋棠 T03-07 / T04-04 / T05-05
    "pH調整池暨快混池",    # 秋棠 T02-06
    "pH調整快混池",
    "pH調整與快混池",
}
PH_FAST_MIX_ALLOW_VARY = {
    "pH值", "pH", "水溫(攝氏)", "水溫",
    "懸浮固體", "懸浮固體（mg/L）", "懸浮固體(mg/L)", "SS", "ss",
}

# 法源用 (保留舊變數名給其他模組/檢查相容)
PH_ONLY_VARIABLE = PURE_PH_ALLOW_VARY

# 沉澱單元類型
SETTLING_TANK_TYPES = {"沉澱池", "沉降池", "浮除槽"}

# 快混槽類型 (不應展現去除率) — 涵蓋常見命名變體
FAST_MIX_TYPES = {
    "快混槽", "快混池",
    "慢混槽", "慢混池",  # 慢混也是「加藥+攪拌, 無固液分離」, 學理同類
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


def check_tank_chemistry(unit):
    """檢查 3: 各槽體進出流水質學理檢查。

    新版邏輯 (v2): 從規則庫.xlsx 的 _槽體學理 分頁讀規則, 涵蓋所有
    常見處理單元 (pH 槽、沉澱池、曝氣槽、脫水機、活性碳塔等 27+ 條規則)。
    若規則庫讀不到, fallback 到本檔內 hardcoded pH 槽常數 (舊行為)。

    判斷基準: Σ進質量 vs Σ出質量, 容忍度依槽體類型而定 (5%~30%)。
    """
    # 優先用新模組
    try:
        import tank_chemistry
        rules = tank_chemistry.load_rules()
        if rules:
            return tank_chemistry.check_unit(unit, rules)
    except Exception:
        pass

    # === Fallback: 舊版只覆蓋 pH 槽 / 暨快混池 ===
    findings = []
    code = unit["raw_code"]
    std_tank = unit["std_tank"]
    name_in_doc = unit.get("name_in_doc", "")
    if name_in_doc in PH_FAST_MIX_TYPES or std_tank in PH_FAST_MIX_TYPES:
        allow_vary = PH_FAST_MIX_ALLOW_VARY
        rule_desc = "pH 暨快混池加酸鹼+混凝劑, 除 pH/SS 外應守恆"
    elif name_in_doc in PURE_PH_TANK_TYPES or std_tank in PURE_PH_TANK_TYPES:
        allow_vary = PURE_PH_ALLOW_VARY
        rule_desc = "純 pH 調整槽只加酸/鹼, 除 pH 外應守恆"
    else:
        return findings

    influent = unit.get("influent", {}) or {}
    effluent = unit.get("effluent", {}) or {}
    if not influent or not effluent:
        return findings

    all_items = set()
    for stream in list(influent.values()) + list(effluent.values()):
        if isinstance(stream, dict):
            all_items.update(stream.keys())
    all_items = {i for i in all_items if i not in allow_vary}

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
        if diff_pct > 5:
            direction = "減少" if out_mass < in_mass else "增加"
            findings.append({
                "嚴重度": "不合理",
                "類型": "質量平衡",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": item,
                "描述": (
                    f"{item} 質量 進 {in_mass:.3f} → 出 {out_mass:.3f} kg/d "
                    f"({direction} {diff_pct:.1f}%)。{rule_desc}。"
                ),
                "依據": f"學理 (fallback): " + rule_desc,
            })
    return findings


# 舊名 alias — 維持向下相容 (給可能 import 舊名的外部呼叫者)
check_ph_tank_only_ph_change = check_tank_chemistry


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


def check_design_metrics(unit):
    """檢查設計參數體檢: HRT / SOR / G 值 是否在學理範圍。

    從 step3h_design_metrics 算指標, 從 tank_chemistry 規則表取學理範圍。
    """
    findings = []
    try:
        import step3h_design_metrics as _m
        import tank_chemistry as _tc
    except Exception:
        return findings

    rule = _tc.get_rule_for_unit(unit)
    if not rule:
        return findings

    code = unit.get("raw_code") or unit.get("code") or "?"
    std_tank = rule["標準槽體"]

    metrics = _m.compute_all_metrics(unit)
    hrt = metrics["hrt_hr"]
    sor = metrics["sor_m3_m2_d"]
    g = metrics["g_value_s_inv"]
    is_lamella = metrics["is_lamella"]
    sev = rule.get("嚴重度") or "待人工"

    # ── HRT 檢查 ──
    hrt_min = rule.get("HRT_min")
    hrt_max = rule.get("HRT_max")
    if hrt is not None and (hrt_min or hrt_max):
        hrt_min_v = hrt_min if hrt_min is not None else 0
        hrt_max_v = hrt_max if hrt_max is not None else float("inf")
        if hrt < hrt_min_v or hrt > hrt_max_v:
            direction = "過短" if hrt < hrt_min_v else "過長"
            findings.append({
                "嚴重度": sev,
                "類型": "設計參數",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": "水力停留時間 (HRT)",
                "描述": (
                    f"HRT = {hrt:.3f} hr ({hrt*60:.1f} 分鐘), {direction}。"
                    f"學理範圍 {hrt_min_v} ~ {hrt_max_v} hr。"
                    f"(V={metrics['volume_m3']:.2f} m³, Q={metrics['main_q_cmd']:.1f} CMD)"
                ),
                "依據": "_槽體學理 HRT 範圍 + 環工設計準則",
            })

    # ── SOR 檢查 (沉澱類) ──
    sor_max = rule.get("SOR_max")
    if sor is not None and sor_max:
        # 斜板池: SOR 上限放寬 (但有提醒)
        effective_max = sor_max
        sor_note = ""
        if is_lamella:
            effective_max = sor_max * 2.4  # 50 → 120
            sor_note = " (已偵測斜板, 上限放寬)"
        if sor > effective_max:
            findings.append({
                "嚴重度": sev,
                "類型": "設計參數",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": "表面溢流率 (SOR)",
                "描述": (
                    f"SOR = {sor:.1f} m³/m²·d, 超過學理上限 {effective_max:.0f}{sor_note}。"
                    f" 顆粒沉降速度可能跟不上水面上升, SS 會被沖走。"
                    f" (Q={metrics['main_q_cmd']:.1f} CMD, A={metrics['surface_area_m2']:.2f} m²)"
                ),
                "依據": "_槽體學理 SOR 範圍 + 環工設計準則",
            })

    # ── G 值檢查 (混合槽類) ──
    g_min = rule.get("G_min")
    g_max = rule.get("G_max")
    if g is not None and (g_min or g_max):
        g_min_v = g_min if g_min is not None else 0
        g_max_v = g_max if g_max is not None else float("inf")
        if g < g_min_v or g > g_max_v:
            direction = "過低" if g < g_min_v else "過高"
            warn_extra = ""
            if g > g_max_v and "慢混" in std_tank:
                warn_extra = " (慢混 G 過高會打散絮羽, 沉澱失敗)"
            findings.append({
                "嚴重度": sev,
                "類型": "設計參數",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": "G 值 (速度梯度)",
                "描述": (
                    f"G = {g:.0f} s⁻¹, {direction}。"
                    f"學理範圍 {g_min_v} ~ {g_max_v} s⁻¹。{warn_extra}"
                    f" (P={metrics['motor_power_w']:.0f} W, V={metrics['volume_m3']:.2f} m³)"
                ),
                "依據": "_槽體學理 G 值範圍 + 環工設計準則",
            })

    return findings


def check_fast_slow_hrt_ratio(app_data):
    """檢查快混 → 慢混的 HRT 比例 (跨單元檢查)。"""
    findings = []
    try:
        import step3h_design_metrics as _m
    except Exception:
        return findings

    pairs = _m.find_fast_slow_pairs(app_data)
    for fast_code, slow_code in pairs:
        result = _m.check_fast_slow_ratio(app_data, fast_code, slow_code)
        if result:
            findings.append(result)
    return findings


def check_cross_unit_q_consistency(app_data):
    """跨單元檢查: 同條 stream 在兩端的反推 Q 應相等。

    某條 stream code (例 WTA01-01-1) 同時出現在:
        - T01-01 的 effluent (作為出流)
        - T01-02 的 influent (作為進流)
    這是「同一條水」, 兩個單元反推出來的 Q 理論上應該相等。
    差 > 10% 表示水質表填寫不一致。
    """
    findings = []
    units = app_data.get("units", {}) or {}

    # 建立 stream_code → [(unit_code, q_cmd, side)]
    appearances = {}
    for code, unit in units.items():
        stream_q = unit.get("stream_q") or {}
        out_streams = list((unit.get("effluent") or {}).keys())
        in_streams = list((unit.get("influent") or {}).keys())
        for s_code in out_streams:
            qres = stream_q.get(s_code)
            if isinstance(qres, dict) and qres.get("ok"):
                q = qres.get("q_cmd")
                if q and q > 0:
                    appearances.setdefault(s_code, []).append((code, q, "out"))
        for s_code in in_streams:
            qres = stream_q.get(s_code)
            if isinstance(qres, dict) and qres.get("ok"):
                q = qres.get("q_cmd")
                if q and q > 0:
                    appearances.setdefault(s_code, []).append((code, q, "in"))

    for s_code, appears in appearances.items():
        if len(appears) < 2:
            continue
        qs = [a[1] for a in appears]
        q_max = max(qs)
        q_min = min(qs)
        if q_max <= 0:
            continue
        diff_pct = (q_max - q_min) / q_max * 100
        if diff_pct > 10:
            severity = "不合理" if diff_pct > 30 else "待人工"
            tags = ", ".join(f"{u}({side},Q={q:.2f})" for u, q, side in appears)
            findings.append({
                "嚴重度": severity,
                "類型": "文件一致性",
                "單元": " ↔ ".join(sorted(set(a[0] for a in appears))),
                "標準槽體": "",
                "對照項目": f"同條流 Q 不一致 ({s_code})",
                "描述": (
                    f"Stream {s_code} 在 {len(appears)} 個單元出現, 反推 Q 不一致: "
                    f"{tags}, 差 {diff_pct:.1f}%。"
                    f"同一條水的 Q 應守恆, 差 > 10% 表示某邊水質表填寫錯誤。"
                ),
                "依據": "學理: 同條 stream 兩端 Q 應相等",
            })
    return findings


def check_quality_table_consistency(unit):
    """檢查: 進/出流水質表的「質量÷濃度」反推 Q 應在 19 項間一致。

    學理: 同一條 stream, 所有水質項目的 Q = 質量/濃度×1000 理論上要相等
         (因為流量 Q 是同一條水)。若多項算出的 Q 差 > 5% (spread_pct),
         表示申請者填質量欄時算錯, 是常見的審查缺失。
    """
    findings = []
    code = unit.get("raw_code") or unit.get("code") or "?"
    std_tank = unit.get("std_tank", "")
    stream_q = unit.get("stream_q", {}) or {}

    for stream_code, qres in stream_q.items():
        if not isinstance(qres, dict):
            continue
        if not qres.get("ok"):
            continue
        spread = qres.get("spread_pct", 0)
        items_count = qres.get("items_count", 0)
        if spread > 5 and items_count >= 3:
            # 嚴重度依差異程度
            severity = "不合理" if spread > 20 else "待人工"
            findings.append({
                "嚴重度": severity,
                "類型": "文件一致性",
                "單元": code,
                "標準槽體": std_tank,
                "對照項目": f"水質表 {stream_code}",
                "描述": (
                    f"{stream_code}: 用 {items_count} 個水質項目「質量÷濃度」"
                    f"反推流量 Q, 算出 {qres['q_min']:.2f} ~ {qres['q_max']:.2f} CMD "
                    f"(差異 {spread:.1f}%, 中位數 {qres['q_cmd']:.2f})。"
                    f"理論上同條流的 Q 應一致, 差異 > 5% 表示水質表的質量欄可能填錯。"
                ),
                "依據": "學理: 質量 = Q × 濃度 × 1e-3, 同條流所有項目 Q 應相等",
            })
    return findings


# ─────────────────── 主入口 ───────────────────

def run_balance_checks(app_data):
    """對整份 app_data 跑所有檢查,回傳 findings list。"""
    findings = []
    checkers = [
        check_dissolved_concentration_up,
        check_fast_mix_metal_removal,
        check_tank_chemistry,            # 槽體學理規則 (規則庫驅動)
        check_settling_overflow_rate,
        check_required_equipment,
        check_quality_table_consistency, # 水質表 Q 反推一致性
        check_design_metrics,            # 新: HRT / SOR / G 值體檢
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

    # 跨單元檢查: 快/慢混 HRT 比例
    try:
        findings.extend(check_fast_slow_hrt_ratio(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(跨單元)",
            "標準槽體": "",
            "對照項目": "check_fast_slow_hrt_ratio",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 跨單元檢查: 同條流 Q 一致
    try:
        findings.extend(check_cross_unit_q_consistency(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(跨單元)",
            "標準槽體": "",
            "對照項目": "check_cross_unit_q_consistency",
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
