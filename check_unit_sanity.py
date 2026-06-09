# -*- coding: utf-8 -*-
"""單位 / 數值合理性偵測 — 抓「數字看起來單位有誤」的情況。

例: 邑昇案頁 76 註解「單位是不是錯了 溢流率每日1.71」
    → 表面溢流率單位是 m³/m²·d, 一般 30~50, 太低 (1.71) → 可能誤填成 m³/m²·hr

常見「單位混淆」異常模式:
    1. 表面溢流率 0~5: 可能是 m³/m²·hr 跟 m³/m²·d 搞混 (24 倍差)
    2. 流量 < 1 CMD: 可能是 CMH 跟 CMD 搞混 (24 倍差)
    3. HRT < 0.01 hr 或 > 100 hr: 可能單位錯
    4. G 值 < 1: 不合理 (G 是 s⁻¹, 一般 10~1000)
"""


def _to_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def check_unit_sanity(unit):
    """檢查單元參數的單位合理性。"""
    findings = []
    code = unit.get("raw_code") or "?"
    std = unit.get("std_tank") or ""

    # 從 design_params + measure_params 找特定參數
    all_params = {**(unit.get("design_params") or {}), **(unit.get("measure_params") or {})}

    for pname, pval in all_params.items():
        if not isinstance(pval, dict):
            continue
        vmax = _to_float(pval.get("max"))
        vmin = _to_float(pval.get("min"))
        # 用最大值判斷範圍 (因 PDF 通常填區間, 但 max 較具代表性)
        v = vmax if vmax is not None else vmin
        if v is None or v <= 0:
            continue

        pname_str = str(pname)
        raw = str(pval.get("raw") or "")

        # ── 表面溢流率 / 水力負荷 ──
        if "溢流率" in pname_str or "水力負荷" in pname_str or "表面負荷" in pname_str:
            # 一般沉澱池 30~50 m³/m²·d, 但有時申報 < 5 表示單位錯
            if v < 5 and "/d" not in raw and "日" not in raw:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v}, 一般沉澱池表面溢流率約 30~50 m³/m²·d。"
                        f"值 < 5 可能單位寫成「每小時」(m³/m²·hr) 而非「每日」(m³/m²·d), "
                        f"差 24 倍。請核對單位。"
                    ),
                    "依據": "環工設計準則: 表面溢流率單位應為 m³/m²·d",
                })

        # ── HRT ──
        if "停留時間" in pname_str or "HRT" in pname_str.upper():
            if v > 100:  # 100 小時太離譜
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v}, 超過 100 小時 (4 天) 不合常理。"
                        f"可能 (a) 流量過小填錯 (b) 容積過大填錯 (c) 單位寫成「分鐘」當「小時」"
                    ),
                    "依據": "環工設計準則: HRT 超過 100 hr 應人工確認",
                })
            elif v < 0.005:  # < 0.005 hr = 18 秒
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v} hr (= {v*3600:.1f} 秒), 太短不合常理。"
                        f"可能單位寫成「小時」當「天」, 或數值漏一位。"
                    ),
                    "依據": "環工設計準則: HRT < 18 秒應人工確認",
                })

        # ── 流量 (CMD) ──
        if "流量" in pname_str or "處理水量" in pname_str or "處理量" in pname_str:
            if v > 1e6:  # > 100 萬 CMD 不合理
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v}, 超過 100 萬 (CMD) 不合常理。"
                        f"可能 (a) 多了 0 (b) 單位寫成 L/d 當 m³/d (差 1000 倍)"
                    ),
                    "依據": "環工設計準則: 單一工廠流量不應超過百萬 CMD",
                })

        # ── G 值 ──
        if "G值" in pname_str or "G 值" in pname_str or "速度梯度" in pname_str:
            if v < 1:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v} s⁻¹, 過低不合理。慢混 10~50, 快混 300~1000, "
                        f"< 1 表示沒攪拌 / 數值填錯。"
                    ),
                    "依據": "環工設計準則: G 值最低 10 s⁻¹",
                })
            elif v > 5000:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v} s⁻¹, 過高不合理。快混上限 1000~1500, "
                        f"> 5000 表示馬達過大 / 槽體過小 / 數值填錯。"
                    ),
                    "依據": "環工設計準則: G 值上限 1500 s⁻¹",
                })

        # ── 加藥量 (kg/Ton 廢水) ──
        if "加藥量" in pname_str and ("ton" in pname_str.lower() or "kg/ton" in pname_str.lower()):
            # kg/Ton 一般 < 1, 若 > 10 可能單位錯
            if v > 50:
                findings.append({
                    "嚴重度": "待人工",
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"單位異常: {pname}",
                    "描述": (
                        f"{pname} = {v} kg/Ton 廢水, 過高。"
                        f"一般加藥率 < 1 kg/Ton (PAC) 或 < 0.1 kg/Ton (Polymer)。"
                        f"> 50 表示單位可能搞錯 (kg/d vs kg/Ton)。"
                    ),
                    "依據": "環工設計準則: 加藥率合理範圍",
                })

    return findings


def run_unit_sanity_checks(app_data):
    """跑所有單元的單位合理性檢查。"""
    findings = []
    for code, unit in app_data.get("units", {}).items():
        try:
            findings.extend(check_unit_sanity(unit))
        except Exception as e:
            findings.append({
                "嚴重度": "錯誤",
                "類型": "系統",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "check_unit_sanity",
                "描述": f"檢查器錯誤: {e}",
                "依據": "(內部)",
            })
    return findings
