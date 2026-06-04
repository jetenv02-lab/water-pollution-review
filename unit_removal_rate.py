# -*- coding: utf-8 -*-
"""計算每個處理單元的削減率 (去除率)。

資料結構 (來自 step2_extract_v2):
    unit["influent"] = {
        "WTB01-01-1": {
            "COD": {"濃度": 100, "質量": 4.9},
            "pH值": {"範圍": "6~9"},
            ...
        },
        "WTB01-01-2": {...}
    }
    unit["effluent"] = { 同結構 }

邏輯:
    對每個水質項目 (有「質量」欄的, 排除 pH/水溫/範圍):
        進流總質量 = Σ 所有進流的該項目「質量」
        出流總質量 = Σ 所有出流的該項目「質量」
        削減率 = (進流總質量 - 出流總質量) / 進流總質量 × 100%
"""

# 不算削減率的項目 (純範圍類)
SKIP_KEYWORDS = ("pH", "ph", "酸鹼", "水溫", "溫度", "溫", "鹼度")


def _should_skip(item_name):
    """判斷該水質項目要不要跳過。"""
    if not item_name:
        return True
    s = str(item_name).strip()
    for kw in SKIP_KEYWORDS:
        if kw in s:
            return True
    return False


def _get_mass(item_data):
    """從項目 dict 取出『質量』值, 沒有就 None。"""
    if not isinstance(item_data, dict):
        return None
    m = item_data.get("質量")
    if m is None or m == "":
        return None
    try:
        return float(m)
    except (ValueError, TypeError):
        return None


def compute_unit_removal(unit):
    """計算單個單元的削減率。

    Returns:
        {
            "items": {
                "COD": {"in_mass": 4.9, "out_mass": 54.3, "removal_pct": -1008.2},
                ...
            },
            "summary": {
                "items_count": int,
                "avg_removal_pct": float,
                "min_removal_pct": float,
                "max_removal_pct": float,
                "warnings": [str]
            }
        }
    """
    influent = unit.get("influent", {}) or {}
    effluent = unit.get("effluent", {}) or {}

    # 蒐集所有水質項目
    all_items = set()
    for stream in list(influent.values()) + list(effluent.values()):
        if isinstance(stream, dict):
            all_items.update(stream.keys())

    # 過濾掉 pH 等
    all_items = {i for i in all_items if not _should_skip(i)}

    items_result = {}
    for item in sorted(all_items):
        in_mass = 0.0
        out_mass = 0.0
        in_has = False
        out_has = False

        for stream in influent.values():
            if isinstance(stream, dict):
                m = _get_mass(stream.get(item))
                if m is not None:
                    in_mass += m
                    in_has = True

        for stream in effluent.values():
            if isinstance(stream, dict):
                m = _get_mass(stream.get(item))
                if m is not None:
                    out_mass += m
                    out_has = True

        # 只有兩邊都有資料才算 (有單側資料 → 無法計算削減率)
        if in_has and out_has and in_mass > 0:
            removal_pct = (in_mass - out_mass) / in_mass * 100
            items_result[item] = {
                "in_mass": round(in_mass, 4),
                "out_mass": round(out_mass, 4),
                "removal_pct": round(removal_pct, 2),
            }

    # summary + 警告
    removals = [r["removal_pct"] for r in items_result.values()]
    warnings = []
    for item, r in items_result.items():
        pct = r["removal_pct"]
        if pct < -10:
            warnings.append(
                f"{item}: 削減率 {pct:.1f}% (出流質量 > 進流質量, 可能水量增加或水質誤植)"
            )
        elif pct > 99.5 and not any(k in item for k in ("懸浮", "SS", "ss")):
            warnings.append(
                f"{item}: 削減率 {pct:.1f}% 過高 (>99.5%, 不合常理)"
            )

    if removals:
        summary = {
            "items_count": len(items_result),
            "avg_removal_pct": round(sum(removals) / len(removals), 2),
            "min_removal_pct": round(min(removals), 2),
            "max_removal_pct": round(max(removals), 2),
            "warnings": warnings,
        }
    else:
        summary = {
            "items_count": 0,
            "avg_removal_pct": None,
            "min_removal_pct": None,
            "max_removal_pct": None,
            "warnings": [],
        }

    return {"items": items_result, "summary": summary}


def compute_all_units_removal(app_data):
    """對 app_data["units"] 全部單元算削減率。"""
    results = {}
    for code, unit in app_data.get("units", {}).items():
        results[code] = compute_unit_removal(unit)
    return results


def format_removal_short(summary):
    """把 summary 格式化成單行短文字 (給表格用)。

    例: "+12.5% (5 項)"  or  "-25% ⚠️"  or  "-"
    """
    if not summary or summary.get("items_count", 0) == 0:
        return "-"
    avg = summary.get("avg_removal_pct")
    if avg is None:
        return "-"
    sign = "+" if avg >= 0 else ""
    text = f"{sign}{avg:.1f}% ({summary['items_count']} 項)"
    if summary.get("warnings"):
        text += " ⚠️"
    return text


# ──────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import json
    import sys
    import io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print("用法: python unit_removal_rate.py <app_data.json>")
        sys.exit(0)

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        app_data = json.load(f)

    print("=== 各單元削減率 ===")
    results = compute_all_units_removal(app_data)
    for code, r in results.items():
        s = r["summary"]
        if s["items_count"] > 0:
            print(f"\n{code}: {format_removal_short(s)}")
            print(f"  min {s['min_removal_pct']:.1f}% / max {s['max_removal_pct']:.1f}%")
            for w in s["warnings"]:
                print(f"  ⚠️ {w}")
        else:
            print(f"{code}: -")
