# -*- coding: utf-8 -*-
"""從進出水質資料表反推 stream 流量 Q。

學理:
    質量 (kg/d) = 流量 Q (m³/d = CMD) × 濃度 (mg/L) × 1e-3

    →  Q = 質量 / 濃度 × 1000

每個水質項目都能算一個 Q, 理論上同一條流的所有項目算出來 Q 要相等
(因為它們的 Q 是同一條水)。實際會有四捨五入誤差, 用中位數較穩。

用途:
    第一階段就能拿到每條 WTA/WTB 流的 Q, 不用等示意圖解析。

特色:
    - 一致性檢驗: 若同一條流的多個項目算出來 Q 差異 > 5%, 表示水質表填錯
    - 來源備註: 結果會標 method = "reverse_from_quality" (誠實告知是反推)
    - 示意圖解析 (Gemini Vision) 結果可用來「驗算」反推值
"""
from statistics import median


def to_float(v):
    """容錯轉 float, 失敗回 None。"""
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def calc_q_from_stream(items, ignore_items=None):
    """從一條流的水質字典反推 Q。

    Args:
        items: dict {水質項目名: {"濃度": float, "質量": float, ...}}
                也接受 {"範圍": "1~7"} 這種 (會被略過)
        ignore_items: 要忽略的項目集合 (例如 pH 值不能用)

    Returns:
        dict {
            "ok": True/False,
            "q_cmd": median Q (CMD),
            "items_used": [(item_name, q_value), ...],  # 排序好
            "items_count": int,
            "spread_pct": float,  # (max - min) / median * 100, 一致性指標
            "consistent": bool,   # spread_pct < 5
            "method": "reverse_from_quality",
        }
    """
    if ignore_items is None:
        # 預設略過: pH (不能除)、水溫 (不能除)、含水率 (百分比不是 mg/L)
        ignore_items = {"pH值", "pH", "水溫", "水溫(攝氏)", "水溫（攝氏）",
                        "含水率", "含水率(%)", "含水率（%）"}

    if not items:
        return {"ok": False, "reason": "empty", "method": "reverse_from_quality"}

    qs = []
    for name, v in items.items():
        if name in ignore_items:
            continue
        if not isinstance(v, dict):
            continue
        c = to_float(v.get("濃度"))
        m = to_float(v.get("質量"))
        if c is None or m is None or c <= 0 or m <= 0:
            continue
        q = m / c * 1000  # m (kg/d) / c (mg/L) × 1000 = Q (m³/d = CMD)
        qs.append((str(name), round(q, 3)))

    if not qs:
        return {
            "ok": False,
            "reason": "no_valid_items",
            "method": "reverse_from_quality",
        }

    q_values = sorted([q for _, q in qs])
    q_median = median(q_values)
    q_min, q_max = q_values[0], q_values[-1]
    spread_pct = (q_max - q_min) / q_median * 100 if q_median > 0 else 0.0

    return {
        "ok": True,
        "q_cmd": round(q_median, 3),
        "q_min": q_min,
        "q_max": q_max,
        "items_used": sorted(qs, key=lambda x: x[0]),
        "items_count": len(qs),
        "spread_pct": round(spread_pct, 2),
        "consistent": spread_pct < 5.0,
        "method": "reverse_from_quality",
    }


def enrich_unit_with_q(unit):
    """給一個 unit (含 influent/effluent), 算每條流的 Q, 寫進 unit。

    會在 unit 加上 'stream_q' 欄:
        {
            "WTB01-01-1": {"q_cmd": 49.0, "consistent": True, "method": ...},
            "WTA01-01-1": {"q_cmd": 608.56, "consistent": True, ...},
        }
    """
    stream_q = {}
    for direction in ("influent", "effluent"):
        streams = unit.get(direction, {}) or {}
        for stream_code, items in streams.items():
            result = calc_q_from_stream(items)
            stream_q[stream_code] = result
    unit["stream_q"] = stream_q
    return unit


def enrich_app_data(app_data):
    """對整份 app_data 的每個 unit 算 stream_q。

    呼叫後 app_data['units'][code]['stream_q'] 就有完整對照表。
    """
    for code, unit in app_data.get("units", {}).items():
        enrich_unit_with_q(unit)
    return app_data


def build_stream_q_map(app_data):
    """組「stream_code → Q (CMD)」總對照表, 給 streamlit 顯示用。

    若同一個 stream_code 出現在多個單元 (一條流既是某單元出流又是另單元進流),
    取**一致性最好**的那個。
    """
    q_map = {}  # code → {"q_cmd": ..., "consistent": ..., "source_unit": ...}
    for unit_code, unit in app_data.get("units", {}).items():
        sq = unit.get("stream_q", {})
        for stream_code, res in sq.items():
            if not res.get("ok"):
                continue
            new_consistent = res.get("consistent", False)
            new_spread = res.get("spread_pct", 100)
            if stream_code not in q_map:
                q_map[stream_code] = {
                    "q_cmd": res["q_cmd"],
                    "consistent": new_consistent,
                    "spread_pct": new_spread,
                    "source_unit": unit_code,
                    "items_count": res.get("items_count", 0),
                    "method": "reverse_from_quality",
                }
            else:
                # 已存在, 比一致性: 取 spread 較小的
                if new_spread < q_map[stream_code]["spread_pct"]:
                    q_map[stream_code] = {
                        "q_cmd": res["q_cmd"],
                        "consistent": new_consistent,
                        "spread_pct": new_spread,
                        "source_unit": unit_code,
                        "items_count": res.get("items_count", 0),
                        "method": "reverse_from_quality",
                    }
    return q_map


def verify_with_diagram(stream_q_map, diagram_result, tolerance_pct=10.0):
    """用示意圖解析結果驗證反推 Q。

    Args:
        stream_q_map: build_stream_q_map() 的結果
        diagram_result: flow_diagram_extractor 的結果 (含 all_flows / external_inputs / discharge_points)
        tolerance_pct: 容忍誤差 %

    Returns:
        list of dict, 每筆是「兩來源不一致」的 stream
    """
    if not diagram_result or not diagram_result.get("ok"):
        return []

    # 從示意圖結果蒐集 stream_code → Q
    diagram_q = {}
    for f in diagram_result.get("all_flows", []) or []:
        q = f.get("Q_cmd")
        if q is None:
            continue
        for k in ("from_stream", "to_stream"):
            code = f.get(k)
            if code and code not in diagram_q:
                diagram_q[str(code)] = q
    for ei in diagram_result.get("external_inputs", []) or []:
        q = ei.get("Q_cmd")
        if q is None:
            continue
        for k in ("code", "to_stream"):
            code = ei.get(k)
            if code and code not in diagram_q:
                diagram_q[str(code)] = q
    for dp in diagram_result.get("discharge_points", []) or []:
        q = dp.get("Q_cmd")
        if q is None:
            continue
        for k in ("code", "from_stream"):
            code = dp.get(k)
            if code and code not in diagram_q:
                diagram_q[str(code)] = q

    # 比對
    mismatches = []
    for code, qinfo in stream_q_map.items():
        if code not in diagram_q:
            continue
        q_calc = qinfo["q_cmd"]
        q_diag = diagram_q[code]
        if q_calc <= 0:
            continue
        diff_pct = abs(q_calc - q_diag) / q_calc * 100
        if diff_pct > tolerance_pct:
            mismatches.append({
                "stream_code": code,
                "q_reversed_from_quality": q_calc,
                "q_from_diagram": q_diag,
                "diff_pct": round(diff_pct, 2),
                "source_unit": qinfo.get("source_unit"),
            })
    return mismatches
