# -*- coding: utf-8 -*-
"""PDF 內計算式反推驗算。

PDF 的設計參數欄常見「8.28m³ ÷ 1.05CMD × 24小時/天」這種計算式,
旁邊還有結果 (例如 189.3 hr)。系統可以:
    1. 抽出計算式裡的數字
    2. 用四則運算重算
    3. 比對申報的結果跟重算結果, 若差 > 5% 標 ⚠️

這就是邑昇案頁 92 人工註解 "13.78? 分母應該是每日產出污泥4.365 ?373CMD是水量?"
要抓的問題。

支援的計算式類型 (從邑昇案常見):
    - "8.28m3 ÷ 1.05CMD × 24小時/天"        → HRT 計算
    - "底面積 × 有效水深"                       → 容積計算 (跳過, 非數值)
    - "馬達功率 / (μ × V) 開根號"               → G 值 (跳過, 太複雜)
"""
import re


_NUM_RE = re.compile(r"\d+(?:\.\d+)?")


def parse_calc_expression(text):
    """從文字抽出計算式 (限四則運算).

    支援:
        "8.28m³ ÷ 1.05CMD × 24小時/天"
        → numbers=[8.28, 1.05, 24], ops=['÷', '×']

    回傳 list of {numbers, ops, result_estimate} 或 None
    """
    if not text:
        return None
    s = str(text)
    # 找包含 ÷ × * / 的式子
    if not any(op in s for op in ["÷", "×", "*", "/"]):
        return None

    # 把全形運算符轉半形
    s_norm = s.replace("÷", "/").replace("×", "*")
    # 找數字 + 運算符序列
    tokens = re.findall(r"\d+(?:\.\d+)?|[*/+-]", s_norm)
    if len(tokens) < 3:
        return None

    # 用 eval 算 (限定純數字運算符, 安全)
    expr = "".join(tokens)
    # 安全檢查: 只能含 digit, dot, * / + -
    if not re.match(r"^[\d.*/+\-]+$", expr):
        return None

    try:
        result = eval(expr)
        return {
            "expression": expr,
            "result": float(result),
            "tokens": tokens,
        }
    except (ValueError, ZeroDivisionError, SyntaxError):
        return None


def check_parameter_calc(unit):
    """檢查單元的設計參數 / 量測參數中, 若有計算式 + 申報值, 是否一致。

    對 design_params 跟 measure_params 的 raw 欄位掃描:
        例: "60 189.3~ [086]小時 8.28m3 ÷ 1.05CMD × 24小時/天"
        → 申報 189.3 vs 計算 8.28/1.05*24 = 189.3 ✓ 一致
    """
    findings = []
    code = unit.get("raw_code") or "?"
    std = unit.get("std_tank") or ""

    for params_name, params in [
        ("設計", unit.get("design_params") or {}),
        ("量測", unit.get("measure_params") or {}),
    ]:
        for pname, pval in params.items():
            if not isinstance(pval, dict):
                continue
            raw = pval.get("raw") or ""
            if not raw:
                continue
            # 申報值 (min ~ max)
            reported_max = pval.get("max")
            reported_min = pval.get("min")

            # 嘗試抽出計算式
            calc = parse_calc_expression(raw)
            if not calc:
                continue
            calculated = calc["result"]
            if calculated <= 0:
                continue

            # 跟申報值比對 (用 max, 因 raw 通常是「~ max」形式)
            reported = reported_max if reported_max is not None else reported_min
            if reported is None or reported <= 0:
                continue

            diff_pct = abs(calculated - reported) / max(calculated, reported) * 100
            if diff_pct > 5:
                severity = "不合理" if diff_pct > 20 else "待確認"
                findings.append({
                    "嚴重度": severity,
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": std,
                    "對照項目": f"{params_name}參數計算驗算: {pname}",
                    "描述": (
                        f"{pname} 申報值 = {reported}, 但用計算式 ({calc['expression']}) "
                        f"重算 = {calculated:.3f}, 差 {diff_pct:.1f}%。"
                        f"可能原因: 分子/分母填錯、單位錯, 或計算式跟結果不匹配。"
                    ),
                    "依據": "PDF 內計算式自動驗算",
                })

    return findings


def run_calc_verify(app_data):
    """跑所有單元的計算式驗算。"""
    findings = []
    for code, unit in app_data.get("units", {}).items():
        try:
            findings.extend(check_parameter_calc(unit))
        except Exception as e:
            findings.append({
                "嚴重度": "錯誤",
                "類型": "系統",
                "單元": code,
                "標準槽體": unit.get("std_tank", ""),
                "對照項目": "check_parameter_calc",
                "描述": f"檢查器錯誤: {e}",
                "依據": "(內部)",
            })
    return findings
