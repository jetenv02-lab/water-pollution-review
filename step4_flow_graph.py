# -*- coding: utf-8 -*-
"""Step 4: 水流串接圖建立器。

從 WTB/WTA 編號規則 + 水質匹配, 推導出「哪個單元的出流流到哪個單元的進流」。

編號規則 (從秋棠申請文件歸納):
    WTB{系統}-{單元}-{流水號}  ← B = Before, 進流
    WTA{系統}-{單元}-{流水號}  ← A = After, 出流
    WM{XX}                  ← 原廢水 (Wastewater Material)
    D{XX}                   ← 放流口 (Discharge)

水流串接邏輯 (參考 jetwatersystem 的線性串接概念):
    A 單元出流 WTA01-01-1 → 通常會變成下一單元的 WTB 進流
    但水措 PDF 可能用同一個 A 編號做 B (因為流到別單元時編號規則複雜)

本模組做兩件事:
1. 從編號規則「猜測」流向 (例如 WTA01-01-1 預期流到 T01-02 的 WTB01-02-X)
2. 用「水質完全一致」做精確比對:
   - 若 T01-01 的 WTA01-01-1 水質和 T02-08 的 WTB02-08-X 水質完全相同
     → 確認 T01-01 → T02-08 有水流
3. 建立有向圖: 節點 = 單元, 邊 = (from_unit, from_stream, to_unit, to_stream, 信心度)

使用:
    from step4_flow_graph import build_flow_graph
    graph = build_flow_graph(app_data)
"""
import json
import os
import re
import sys
from collections import defaultdict


# ────────────────────────────────────────────────
# 編號解析
# ────────────────────────────────────────────────

# WTB01-01-1 → ("B", "01", "01", "1")
STREAM_CODE_RE = re.compile(r"^WT([AB])(\d{2})-(\d{2})(?:-(\d+))?")
WM_CODE_RE = re.compile(r"^WM(\d{2})")
DISCHARGE_RE = re.compile(r"^D(\d{2})")


def parse_stream_code(code):
    """解析水流編號 → (種類, 系統, 單元, 流水號) 或 None。"""
    m = STREAM_CODE_RE.match(code)
    if m:
        return {
            "kind": "WTA" if m.group(1) == "A" else "WTB",
            "system": m.group(2),    # 01-05 (T01-T05)
            "unit_seq": m.group(3),  # 01, 02, ...
            "stream_seq": m.group(4) or "",
            "owner_unit": f"T{m.group(2)}-{m.group(3)}",
        }
    if WM_CODE_RE.match(code):
        return {"kind": "WM", "raw": code}
    if DISCHARGE_RE.match(code):
        return {"kind": "D", "raw": code}
    return None


# ────────────────────────────────────────────────
# 水質指紋 (用來判斷兩股流是不是同一股)
# ────────────────────────────────────────────────

def quality_fingerprint(quality_dict, max_items=8):
    """產生一股流的水質「指紋」 — 取主要項目濃度湊成 hashable tuple。

    支援兩種值格式:
      {"濃度": N, "質量": M}  (step2 抽水質表的 stream)
      "30" or "3~ 7"           (step2_raw_water 抽 WM 的 string)
    """
    if not quality_dict:
        return None
    items = []
    for item_name, v in quality_dict.items():
        c = None
        if isinstance(v, dict):
            c = v.get("濃度")
        elif isinstance(v, (int, float)):
            c = v
        elif isinstance(v, str):
            # WM 格式: 可能是 "30" 或 "3~ 7" 或 "15~ 35"
            s = v.strip()
            # 範圍類 (pH/水溫) 跳過, 因為不是濃度
            if "~" in s or "-" in s:
                continue
            c = s
        try:
            items.append((str(item_name), round(float(c), 3)))
        except (TypeError, ValueError):
            continue
    items.sort()
    return tuple(items[:max_items])


# ────────────────────────────────────────────────
# 水流串接圖建立
# ────────────────────────────────────────────────

def build_flow_graph(app_data):
    """建立水流串接圖。

    Returns:
        {
            "nodes": [{code, std_tank, name}],  # 處理單元
            "edges": [{
                from_unit, from_stream, to_unit, to_stream,
                confidence, method  # "編號規則" or "水質指紋" or "兩者"
            }],
            "streams_index": {WTB/WTA 編號: {owner_unit, kind, quality, ...}},
            "wm_sources": [WM01, WM02, ...],   # 原廢水
            "discharges": [D01, ...],         # 放流口
        }
    """
    units = app_data.get("units", {})

    # 1) 建立 streams_index: 編號 → 所屬單元/種類/水質
    streams_index = {}
    for code, unit in units.items():
        for infl_code, q in unit.get("influent", {}).items():
            streams_index[infl_code] = {
                "kind_actual": "WTB",
                "owner_unit": code,
                "quality": q,
                "fingerprint": quality_fingerprint(q),
            }
        for effl_code, q in unit.get("effluent", {}).items():
            streams_index[effl_code] = {
                "kind_actual": "WTA",
                "owner_unit": code,
                "quality": q,
                "fingerprint": quality_fingerprint(q),
            }

    # 1.5) 補加 WMxx (原廢水) 到 streams_index
    # 從 app_data["raw_water"] 來 (step2_raw_water 抽的)
    for wm_code, wm_data in (app_data.get("raw_water") or {}).items():
        if wm_code in streams_index:
            continue  # 已存在不覆蓋
        wm_quality = wm_data.get("quality") or {}
        streams_index[wm_code] = {
            "kind_actual": "WM",
            "owner_unit": wm_code,
            "quality": wm_quality,
            "fingerprint": quality_fingerprint(wm_quality),
            "q_cmd": wm_data.get("q_cmd"),
            "sources": wm_data.get("sources", []),
        }

    # 2) 找 WM (原廢水進入點) 和 D (放流)
    wm_sources = sorted({code for code in streams_index if code.startswith("WM")})
    discharges = sorted({code for code in streams_index if code.startswith("D")})

    # 3) 推導邊: 對每個 WTA, 找它流到哪
    edges = []
    # 方法 A: 水質指紋完全相符
    wta_codes = [c for c in streams_index if c.startswith("WTA")]
    wtb_codes = [c for c in streams_index if c.startswith("WTB")]

    def _unit_seq(unit_code):
        """T01-05 → 5; T02-03 → 203 (前綴 × 100 + 序號) 用於排序"""
        import re
        m = re.match(r"T(\d{2})-(\d+)", str(unit_code))
        if m:
            return int(m.group(1)) * 100 + int(m.group(2))
        return -1

    matched_wtb = set()  # 已被某 WTA 配對的 WTB
    for wta in wta_codes:
        wta_info = streams_index[wta]
        wta_fp = wta_info.get("fingerprint")
        if not wta_fp:
            continue
        wta_unit = wta_info["owner_unit"]
        wta_seq = _unit_seq(wta_unit)
        # 找所有指紋相同的候選 WTB
        candidates = []
        for wtb in wtb_codes:
            if wtb in matched_wtb:
                continue
            wtb_info = streams_index[wtb]
            if wtb_info["owner_unit"] == wta_unit:
                continue  # 不能流向自己
            if wtb_info.get("fingerprint") != wta_fp:
                continue
            wtb_unit = wtb_info["owner_unit"]
            wtb_seq = _unit_seq(wtb_unit)
            candidates.append((wtb, wtb_unit, wtb_seq))
        if not candidates:
            continue
        # 優先序: WTB unit 序號 > WTA unit 序號 (合理下游) + 序號最近
        # 計算「距離」: 同前綴下游 = 正距離; 不同前綴 = +1000; 往回 = +5000
        def _dist(c):
            wtb, wtb_unit, wtb_seq = c
            # 同前綴 (T0X) 且下游
            if wta_seq > 0 and wtb_seq > wta_seq:
                same_prefix = (wta_unit[:3] == wtb_unit[:3])
                return (wtb_seq - wta_seq) + (0 if same_prefix else 1000)
            # 往回 (應該避開) — 給很大距離
            return 5000 + abs(wtb_seq - wta_seq)
        candidates.sort(key=_dist)
        best_wtb, best_unit, _ = candidates[0]
        edges.append({
            "from_unit": wta_unit,
            "from_stream": wta,
            "to_unit": best_unit,
            "to_stream": best_wtb,
            "confidence": "高",
            "method": "水質指紋一致 (多候選時取最近下游)",
        })
        matched_wtb.add(best_wtb)

    # 3.5) 沒被配對的 WTB → 來自外部 (WM 原廢水)
    # 例如 WTB01-01-6 沒對應的 WTA, 通常是從 WMxx 直接進來的
    for wtb in wtb_codes:
        if wtb in matched_wtb:
            continue
        wtb_info = streams_index[wtb]
        # 嘗試從 WM sources 中找指紋一致的 (若 WM 有水質)
        wtb_fp = wtb_info.get("fingerprint")
        matched_wm = None
        if wtb_fp and wm_sources:
            # 策略 1: 完全相符
            for wm in wm_sources:
                wm_info = streams_index.get(wm, {})
                if wm_info.get("fingerprint") == wtb_fp:
                    matched_wm = wm
                    break
            # 策略 2: 部分相符 (共同項目 ≥ 2 個, 濃度差 ≤ 20% 視為配對)
            # 處理 PDF 填表不一致時的容錯
            if not matched_wm:
                wtb_dict = dict(wtb_fp)  # tuple → dict
                for wm in wm_sources:
                    wm_info = streams_index.get(wm, {})
                    wm_fp = wm_info.get("fingerprint")
                    if not wm_fp:
                        continue
                    wm_dict = dict(wm_fp)
                    common_keys = set(wtb_dict.keys()) & set(wm_dict.keys())
                    if len(common_keys) < 2:
                        continue
                    matches = 0
                    for k in common_keys:
                        v1, v2 = wtb_dict[k], wm_dict[k]
                        try:
                            if v1 > 0 and abs(v1 - v2) / max(v1, v2) <= 0.2:
                                matches += 1
                        except (TypeError, ZeroDivisionError):
                            pass
                    # 至少 50% 共同項目接近 → 視為配對
                    if matches >= max(1, len(common_keys) * 0.5):
                        matched_wm = wm
                        break

        # 策略 3: 全廠只有一個 WM 時直接給它 (廠商填表不完整, 但邏輯上必有來源)
        if not matched_wm and len(wm_sources) == 1:
            matched_wm = wm_sources[0]

        # 策略 4: 多 WM 但都對不上時 → 找「水質完全空的 WM」
        # (廠商可能漏填某個 WM 的水質, 該 WM 對應的下游無法靠指紋抓)
        if not matched_wm and len(wm_sources) > 1:
            empty_wm = []
            for wm in wm_sources:
                wm_info = streams_index.get(wm, {})
                wm_q = wm_info.get("quality") or {}
                # 算「有濃度的項目」數量
                has_conc = 0
                for v in wm_q.values():
                    if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip() and "~" not in v and "-" not in v):
                        has_conc += 1
                if has_conc == 0:
                    empty_wm.append(wm)
            # 只有一個空的 → 把這個未配對的 WTB 給它
            if len(empty_wm) == 1:
                matched_wm = empty_wm[0]
        if matched_wm:
            edges.append({
                "from_unit": matched_wm,   # 例: WM08
                "from_stream": matched_wm,
                "to_unit": wtb_info["owner_unit"],
                "to_stream": wtb,
                "confidence": "高",
                "method": "水質指紋一致 (來自原廢水)",
            })
            matched_wtb.add(wtb)
        else:
            # 配對失敗 — 系統不知道確切來源
            # 不裝懂寫「外部進入」, 改寫「來源未配對」+ 提示
            # (如果有開 Vision 解析, streamlit_app 會用 Vision 結果補真實編號;
            #  如果 Vision 也沒結果, 顯示「來源待確認確認」讓使用者去看 PDF)
            edges.append({
                "from_unit": "? 來源未配對",
                "from_stream": "?",
                "to_unit": wtb_info["owner_unit"],
                "to_stream": wtb,
                "confidence": "低",
                "method": "step2 配對失敗 (水質指紋不符任何 WTA 或 WM)",
            })
            matched_wtb.add(wtb)

    # 4) 整理節點
    nodes = []
    for code, unit in sorted(units.items()):
        nodes.append({
            "code": code,
            "std_tank": unit.get("std_tank", ""),
            "name": unit.get("name_in_doc", ""),
        })

    return {
        "nodes": nodes,
        "edges": edges,
        "wm_sources": wm_sources,
        "discharges": discharges,
        "streams_count": len(streams_index),
    }


def get_unit_neighbors(graph, unit_code):
    """取得某單元的上游 / 下游單元。

    Returns:
        {
            "upstream": [(from_unit, from_stream, to_stream)],  # 流入這單元
            "downstream": [(to_unit, from_stream, to_stream)],   # 流出到別處
        }
    """
    upstream = []
    downstream = []
    for e in graph.get("edges", []):
        if e["to_unit"] == unit_code:
            upstream.append({
                "from_unit": e["from_unit"],
                "from_stream": e["from_stream"],
                "to_stream": e["to_stream"],
                "confidence": e["confidence"],
            })
        if e["from_unit"] == unit_code:
            downstream.append({
                "to_unit": e["to_unit"],
                "from_stream": e["from_stream"],
                "to_stream": e["to_stream"],
                "confidence": e["confidence"],
            })
    return {"upstream": upstream, "downstream": downstream}


# ────────────────────────────────────────────────
# 主入口
# ────────────────────────────────────────────────

if __name__ == "__main__":
    import io
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

    graph = build_flow_graph(app_data)

    print(f"=== 水流串接圖: {app_data.get('source_pdf', '?')} ===")
    print(f"處理單元 (節點): {len(graph['nodes'])}")
    print(f"水流串接 (邊): {len(graph['edges'])}")
    print(f"原廢水進入點 (WM): {graph['wm_sources']}")
    print(f"放流口 (D): {graph['discharges']}")
    print()

    # 顯示前 20 條串接
    print("=== 已偵測的水流串接 ===")
    for i, e in enumerate(graph["edges"][:20], 1):
        print(f"{i}. {e['from_unit']} ({e['from_stream']}) → {e['to_unit']} ({e['to_stream']}) [{e['method']}, 信心: {e['confidence']}]")

    if len(graph["edges"]) > 20:
        print(f"... 還有 {len(graph['edges']) - 20} 條")

    # 範例: 某單元的上下游
    print()
    print("=== T01-01 的上下游 ===")
    nb = get_unit_neighbors(graph, "T01-01")
    print(f"上游 (流入 T01-01): {len(nb['upstream'])} 條")
    for u in nb["upstream"][:5]:
        print(f"  {u['from_unit']} ({u['from_stream']}) → T01-01 ({u['to_stream']})")
    print(f"下游 (T01-01 流出): {len(nb['downstream'])} 條")
    for d in nb["downstream"][:5]:
        print(f"  T01-01 ({d['from_stream']}) → {d['to_unit']} ({d['to_stream']})")

    # 輸出 JSON
    out = os.path.join(BASE, "flow_graph.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(graph, f, ensure_ascii=False, indent=2)
    print(f"\n已輸出: {out}")
