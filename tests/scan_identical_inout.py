"""掃描「進=出水質完全相同」可疑單元 (廠商可能未實際填寫,直接複製貼上)。

判斷準則:
  - 跳過 水溫 / pH (本來就應該幾乎不變)
  - 比對 濃度 (mg/L)
  - 若 >= 80% 的水質項目「進加權平均」與「出加權平均」完全相同 (差 < 0.5%), 視為可疑
  - 對「反應槽」/「分離槽」類型才掃 (儲存/集水/泵浦/管路 本來就可能進=出)

目標槽體類型:
  快混 / 慢混 / pH調整 / 中和 / 化沉 / 沉澱 / 浮除 / 油脂分離
  砂濾 / 活性碳 / 離子交換 / 混凝 / 膠凝 / 氧化 / 氰氧化 / 鉻還原 / 預處理
"""

import json
import sys
from pathlib import Path

SKIP_ITEMS = {"水溫", "pH", "PH", "ph"}

TARGET_TANK_KEYWORDS = [
    "快混", "慢混", "pH", "中和", "化沉", "沉澱", "浮除",
    "油脂", "砂濾", "活性碳", "離子交換", "混凝", "膠凝",
    "氧化", "氰氧化", "鉻還原", "預處理", "酸鹼", "曝氣",
    "生物", "脫氮", "脫磷", "MBR", "薄膜",
]


def is_target_tank(std_tank: str) -> bool:
    if not std_tank:
        return False
    for kw in TARGET_TANK_KEYWORDS:
        if kw in std_tank:
            return True
    return False


def get_stream_q(unit: dict, stream_code: str) -> float | None:
    """從 unit['stream_q'][code]['q_cmd'] 拿 Q (CMD)。"""
    sq = (unit.get("stream_q") or {}).get(stream_code)
    if not isinstance(sq, dict):
        return None
    q = sq.get("q_cmd")
    try:
        return float(q) if q is not None else None
    except (TypeError, ValueError):
        return None


def weighted_avg(streams: dict, unit: dict) -> dict[str, tuple[float, float]]:
    """回傳 {item: (加權平均濃度, 總質量)}。若無 Q,fallback 用算術平均。"""
    item_data: dict[str, list[tuple[float, float]]] = {}  # item -> [(c, q)]
    for stream_code, content in streams.items():
        if not isinstance(content, dict):
            continue
        q = get_stream_q(unit, stream_code) or 1.0  # fallback 1 (算術平均)
        for item, v in content.items():
            if item in SKIP_ITEMS:
                continue
            if isinstance(v, dict):
                c = v.get("濃度")
                try:
                    c = float(c)
                except (TypeError, ValueError):
                    continue
                item_data.setdefault(item, []).append((c, q))

    out = {}
    for item, lst in item_data.items():
        q_sum = sum(q for _, q in lst)
        if q_sum <= 0:
            out[item] = (sum(c for c, _ in lst) / len(lst), 0.0)
        else:
            c_avg = sum(c * q for c, q in lst) / q_sum
            out[item] = (c_avg, q_sum)
    return out


def is_identical_in_out(unit: dict, threshold: float = 0.005) -> dict:
    """回傳 {ratio, identical_items, diff_items, in_avg, out_avg}。

    ratio = 完全相同 (差 < threshold) 的項目數 / 總比對項目數。
    """
    inf = unit.get("influent") or {}
    eff = unit.get("effluent") or {}
    if not inf or not eff:
        return {"ratio": 0.0, "n_total": 0, "identical": [], "diff": [], "in_avg": {}, "out_avg": {}}

    in_avg = weighted_avg(inf, unit)
    out_avg = weighted_avg(eff, unit)
    common = set(in_avg) & set(out_avg)

    identical = []
    diff = []
    for item in common:
        in_c, _ = in_avg[item]
        out_c, _ = out_avg[item]
        if in_c == 0 and out_c == 0:
            identical.append((item, in_c, out_c))
            continue
        base = max(abs(in_c), abs(out_c))
        if base == 0:
            identical.append((item, in_c, out_c))
            continue
        rel_diff = abs(in_c - out_c) / base
        if rel_diff < threshold:
            identical.append((item, in_c, out_c))
        else:
            diff.append((item, in_c, out_c, rel_diff))

    n_total = len(common)
    ratio = len(identical) / n_total if n_total > 0 else 0.0
    return {
        "ratio": ratio,
        "n_total": n_total,
        "identical": identical,
        "diff": diff,
        "in_avg": {k: v[0] for k, v in in_avg.items()},
        "out_avg": {k: v[0] for k, v in out_avg.items()},
    }


def scan_file(path: str | Path, min_ratio: float = 0.80) -> list:
    p = Path(path)
    if not p.exists():
        print(f"[X] {p} 不存在")
        return []
    data = json.loads(p.read_text(encoding="utf-8"))
    units = data.get("units") or {}
    suspicious = []
    for code, unit in units.items():
        std_tank = unit.get("std_tank") or ""
        if not is_target_tank(std_tank):
            continue
        r = is_identical_in_out(unit)
        if r["n_total"] < 3:  # 至少 3 個項目才有意義
            continue
        if r["ratio"] >= min_ratio:
            suspicious.append({
                "code": code,
                "tank": std_tank,
                "ratio": r["ratio"],
                "n_total": r["n_total"],
                "n_identical": len(r["identical"]),
                "identical_items": [it[0] for it in r["identical"]],
                "diff_items": [(it[0], it[1], it[2]) for it in r["diff"]],
            })
    return suspicious


def main():
    files = sys.argv[1:] or [
        "application_05 申請資料(秋棠)(1150119)final.json",
    ]
    print(f"{'='*70}")
    print("掃描「進=出水質完全相同」可疑單元")
    print(f"門檻: 至少 80% 水質項目完全相同 (差 < 0.5%)")
    print(f"{'='*70}\n")

    for f in files:
        print(f"[案件] {f}")
        results = scan_file(f)
        if not results:
            print("  (無可疑單元)\n")
            continue
        for r in results:
            print(f"  ⚠️  {r['code']} ({r['tank']})  相同率 {r['ratio']*100:.0f}% "
                  f"({r['n_identical']}/{r['n_total']})")
            print(f"     完全相同: {', '.join(r['identical_items'][:8])}"
                  + (" ..." if len(r['identical_items']) > 8 else ""))
            if r['diff_items']:
                diffs = [f"{it[0]}({it[1]:.2f}→{it[2]:.2f})" for it in r['diff_items'][:5]]
                print(f"     有變動  : {', '.join(diffs)}")
        print()


if __name__ == "__main__":
    main()
