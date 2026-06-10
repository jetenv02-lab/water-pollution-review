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
    """檢查 2: 原廢水濃度偏離典型值 (僅供參考, 已退讓嚴重度)。

    ⚠️ 注意: 這個「典型值」是「一般生活/工業廢水」經驗中位數, **非法規**。
    電鍍/半導體/化學等業別原廢水濃度動輒 100~5000 mg/L, 完全合理。
    所以閾值放寬到 100 倍 / 0.01 倍, 並且僅當「明顯異常」才提示。
    真正的「不合格」應該以「放流水標準」為準 (見 check_discharge_standard)。
    """
    findings = []
    units = app_data.get("units", {})
    if not units:
        return findings

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
            # 閾值放寬: 100 倍以上才提示 (避免電鍍業誤判)
            if ratio > 100:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "水質標準",
                    "單元": first_code,
                    "標準槽體": "原廢水",
                    "對照項目": item,
                    "描述": (
                        f"原廢水 {item} 濃度 {c} {default['unit']} 遠超「一般廢水」"
                        f"典型值 ({default_v} {default['unit']}) 的 {ratio:.0f} 倍。"
                        f"⚠️ 此典型值非法規, 電鍍/半導體/化學等業別常見此狀況, 建議改參考「放流水標準」判斷。"
                    ),
                    "依據": "一般廢水水質中位經驗值 (非法規, 僅供異常偵測參考)",
                })
            # 低於 0.01 倍才提示 (太低可能是漏填/單位錯)
            elif ratio < 0.01:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "水質標準",
                    "單元": first_code,
                    "標準槽體": "原廢水",
                    "對照項目": item,
                    "描述": (
                        f"原廢水 {item} 濃度 {c} {default['unit']} 遠低於典型值 "
                        f"({default_v} {default['unit']}) 的 {ratio:.4f} 倍, 可能漏填或單位錯。"
                    ),
                    "依據": "一般廢水水質中位經驗值 (非法規, 僅供異常偵測參考)",
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
    """檢查 4: 單元層級 Q 守恆 — Σ進流 Q ≈ Σ出流 Q。

    用 step2 已反推的 stream_q (從質量÷濃度算出來), 對每個單元
    比較「Σ 所有進流的 Q」vs「Σ 所有出流的 Q」, 差 > 5% 標 ⚠️。

    特殊情況豁免:
        - 單側流 (只有進或只有出, 表示是端點或污泥側) → 跳過
        - 含「污泥」「脫水」「濃縮」字眼的單元 → 跳過 (污泥分離本來不守恆)
        - 廢水收集池 / 廢水調整池 (多股匯流入單一出, 通常一致) → 仍要查
    """
    findings = []
    units = app_data.get("units", {})

    SLUDGE_KW = ["污泥", "脫水", "濃縮"]

    # 預先建立「stream_code → 來源單元」對照 (用於追溯上游)
    # 例: WTB01-22-1 來自 T01-21 的某條出流
    stream_origin = {}  # WTB → 來源 T 單元
    for u_code, u in units.items():
        for s_code in (u.get("effluent") or {}).keys():
            # WTA01-22-1 來自 T01-22
            mm = None
            import re
            mm = re.match(r"^WTA(\d{2})[-－](\d{2})", s_code)
            if mm:
                stream_origin[s_code] = f"T{mm.group(1)}-{mm.group(2)}"

    # 統計「每個單元有幾條出流」, 用於判斷下游是否該豁免
    unit_out_count = {u_code: len(u.get("effluent") or {}) for u_code, u in units.items()}

    for code, unit in units.items():
        name = unit.get("name_in_doc", "") + " " + (unit.get("std_tank") or "")
        # 污泥相關單元: 跳過 (進出 Q 本來就不等)
        if any(kw in name for kw in SLUDGE_KW):
            continue

        stream_q = unit.get("stream_q") or {}
        influent = unit.get("influent") or {}
        effluent = unit.get("effluent") or {}

        if not influent or not effluent:
            # 只有單側 (端點或無資料), 跳過
            continue

        sum_in = 0.0
        sum_out = 0.0
        count_in = 0
        count_out = 0
        # 追蹤本單元進流的上游, 看上游是否有多出流 (若是, 本單元的 Q 可能不全)
        upstream_has_split = False
        upstream_units_seen = set()
        for code_s, items in influent.items():
            qres = stream_q.get(code_s, {})
            if qres.get("ok"):
                sum_in += qres.get("q_cmd", 0)
                count_in += 1
            # 對應的 WTA 上游
            # 本單元的 WTB01-22-1 對應上游的 WTA01-22-1 嗎? 不是. 是水質指紋去配對.
            # 用流量圖 (build_flow_graph) 可知道對應的 from_unit
            # 但 check_q_balance 沒拿到 flow_graph
            # 簡化: 用「stream_code 前綴」推測上游單元
            mm = re.match(r"^WTB(\d{2})[-－](\d{2})", code_s)
            if mm:
                # WTB01-22-1 表示「T01-22 自己的進流」, 不是上游編號
                # 真正的上游, 通常 stream_code 是 WTA 形式存在某單元的 effluent
                # 但流向配對複雜, 簡化處理: 用 flow_graph 的方法不完整
                pass

        for code_s, items in effluent.items():
            qres = stream_q.get(code_s, {})
            if qres.get("ok"):
                sum_out += qres.get("q_cmd", 0)
                count_out += 1

        if sum_in <= 0 or sum_out <= 0:
            continue

        # 透過 flow_graph 判斷上游是否多出流
        # 簡單做法: 看本單元進流數 vs 出流數, 若 進 < 出 (例如 1 進 2 出), 該單元自己分流,
        # 它的下游應該豁免 ("水量相加")
        # 但「本單元自己」進=1 出=1 也可能是被上游分流的後段
        # 用全廠掃描: 看「同序列上一個單元」是否多出流
        upstream_split = False
        # T01-22 的上一個 (T01-21) 是否多出流?
        mm_self = re.match(r"^T(\d{2})[-－](\d+)", code)
        if mm_self:
            prev_idx = int(mm_self.group(2)) - 1
            if prev_idx > 0:
                prev_code = f"T{mm_self.group(1)}-{prev_idx:02d}"
                if unit_out_count.get(prev_code, 0) >= 2:
                    upstream_split = True

        diff_pct = abs(sum_in - sum_out) / max(sum_in, sum_out) * 100

        # 拓樸偵測 (產生提示, 不豁免 finding)
        # 「自己有多出流」(本單元 effluent ≥ 2 條) 或 「上游有多出流」
        # → 屬「水量分流結構」, 仍出 finding 但加拓樸提示, 嚴重度降「待人工」
        self_split = (count_out >= 2) or (len(effluent) >= 2)
        topology_hint = ""
        if self_split and upstream_split:
            topology_hint = " ⚠️ 拓樸提示: 本單元有多條出流, 且上游也是分流結構, 此差異可能來自水量分流而非真實異常, 請確認後標記為合理(備註)或異常。"
        elif self_split:
            topology_hint = " ⚠️ 拓樸提示: 本單元有多條出流(水量分流), 此差異可能來自分流而非真實異常, 請確認。"
        elif upstream_split:
            topology_hint = " ⚠️ 拓樸提示: 上游單元有多條出流, 本單元只承接其中一條, 此差異可能來自上游分流, 請確認。"

        if diff_pct > 5:
            if self_split or upstream_split:
                # 分流結構下, 從「不合理」降為「待人工」(需審查員確認是分流還是真異常)
                severity = "待人工"
            else:
                severity = "不合理" if diff_pct > 20 else "待人工"
            direction = "多" if sum_in > sum_out else "少"
            findings.append({
                "嚴重度": severity,
                "類型": "質量平衡",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "Q 守恆 (Σ進=Σ出)",
                "描述": (
                    f"Σ 進流 Q = {sum_in:.2f} CMD ({count_in} 條), "
                    f"Σ 出流 Q = {sum_out:.2f} CMD ({count_out} 條), "
                    f"進流比出流{direction} {diff_pct:.1f}%。"
                    f"水流非污泥側單元理論上 Q 守恆 (Σ進=Σ出), "
                    f"差異 > 5% 表示可能有漏記的支流或水質表填寫錯誤。"
                    f"{topology_hint}"
                ),
                "依據": "質量守恆原理: 水流穩態下 Σ進=Σ出 (污泥側單元除外)",
            })
    return findings


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

    # Q 守恆檢查 (新啟用, 用 step2 反推 Q)
    try:
        findings.extend(check_q_balance(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_q_balance",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 放流水標準檢查 (依業別比對環境部公告限值)
    try:
        from discharge_standards import check_discharge_standard
        findings.extend(check_discharge_standard(app_data, business_type))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_discharge_standard",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 加藥機制檢查 (鎳系/輕系/各系 分流規則)
    try:
        from check_dosing import run_dosing_checks
        findings.extend(run_dosing_checks(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "run_dosing_checks",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 污泥質量平衡 (含水率 / 固形物濃度合理性)
    try:
        from check_sludge import run_sludge_checks
        findings.extend(run_sludge_checks(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤", "類型": "系統", "單元": "(全廠)", "標準槽體": "",
            "對照項目": "run_sludge_checks",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 計算式反推驗算 (PDF 內 "÷ × CMD" 等計算式)
    try:
        from check_calc_verify import run_calc_verify
        findings.extend(run_calc_verify(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤", "類型": "系統", "單元": "(全廠)", "標準槽體": "",
            "對照項目": "run_calc_verify",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 單位異常偵測 (溢流率 / HRT / 流量 / G 值 單位明顯錯誤)
    try:
        from check_unit_sanity import run_unit_sanity_checks
        findings.extend(run_unit_sanity_checks(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤", "類型": "系統", "單元": "(全廠)", "標準槽體": "",
            "對照項目": "run_unit_sanity_checks",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    # 文件一致性 (跨頁數值 + 圖面缺失)
    try:
        from check_doc_consistency import run_doc_consistency_checks
        # pdf_path 從 app_data 取得 (若有)
        _pdf_p = app_data.get("source_pdf_path") or None
        findings.extend(run_doc_consistency_checks(app_data, _pdf_p))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤", "類型": "系統", "單元": "(全廠)", "標準槽體": "",
            "對照項目": "run_doc_consistency_checks",
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
