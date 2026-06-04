# -*- coding: utf-8 -*-
"""從申請文件 PDF 的「水量平衡示意圖」抽出結構化的流向資料。

申請文件結構:
    參、水污染防治措施資料 / 水質水量平衡示意圖   ← 多頁 (各 Txx 系統一頁)
        頁次: 7 (整體概要)
        頁次: 8 (T01 廢水處理系統)
        頁次: 9 (T02 / T03 / ...)
        ...

每頁是一張圖, 用 Gemini Vision 抽出:
    - 該系統的處理單元清單
    - 流向 (from → to, 流量 Q)
    - 跨系統的進入/離開點

最後合併所有頁的流向, 跟現有 step2 抽出的水質資料對照。
"""
import io
import json
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))


# ──────────────────────────────────────────────────
# Gemini Prompt
# ──────────────────────────────────────────────────

EXTRACTION_PROMPT = """這是台灣水污染防治措施申請文件的「水量平衡示意圖」(Water Balance Flow Diagram)。

請從圖中抽出每一條水流箭頭, 整理成結構化 JSON。

圖上會有的元素:
- 方框: 處理單元, 例 "T03-05 第一氧化池"
- 箭頭: 表示水流方向 (從一個方框流到另一個方框)
- 編號標籤 (在箭頭旁邊):
  - WTAxx-yy-z = 第 xx-yy 號處理單元的「第 z 條出流」(After / 出)
  - WTBxx-yy-z = 第 xx-yy 號處理單元的「第 z 條進流」(Before / 進)
  - WMxx = 原廢水 (Wastewater 直接從製程進入廠區, 是某條進流的「外部上游」)
  - Dxx = 放流口 (Discharge, 是某條出流的「外部下游」)
- 流量: Q = XX CMD (Cubic Meters per Day)
- 質量數據 (可能標在箭頭旁): kg/day 等

【最重要 — 同一條箭頭兩端的編號是「同一條水」, 只是命名規則不同】

情況 A — 兩端都是處理單元:
例: T01-01 → T01-02 之間的箭頭, 在上游叫 WTA01-01-1 (T01-01 的出流),
   在下游叫 WTB01-02-1 (T01-02 的進流), 兩個是同一條水。
所以一條 flow 應寫:
  from_unit="T01-01", from_stream="WTA01-01-1"
  to_unit="T01-02",   to_stream="WTB01-02-1"

情況 B — 來源是 WMxx 原廢水:
例: WM08 化撿廢水 → T01-01 的箭頭, 在上游叫 WM08, 在下游叫 WTB01-01-6,
   兩個編號也是同一條水, 只是上游沒有處理單元。
所以這條 flow 應寫:
  from_unit="WM08", from_stream="WM08"     (沒有處理單元時, from_unit 就是 WMxx)
  to_unit="T01-01", to_stream="WTB01-01-6"
  Q_cmd=47.5
**這條 WMxx 應該同時也寫進 external_inputs 裡, 並標明 to_stream:**
  external_inputs: [{"code":"WM08", "name":"化撿廢水", "Q_cmd":47.5, "to_unit":"T01-01", "to_stream":"WTB01-01-6"}]

情況 C — 目的是 Dxx 放流口:
例: T01-02 → D01 的箭頭, 在上游叫 WTA01-02-1 (T01-02 的出流), 下游就是 D01。
所以這條 flow 應寫:
  from_unit="T01-02", from_stream="WTA01-02-1"
  to_unit="D01",      to_stream="D01"
**同時也寫進 discharge_points, 並標明 from_stream:**
  discharge_points: [{"code":"D01", "name":"放流口", "Q_cmd":608.56, "from_unit":"T01-02", "from_stream":"WTA01-02-1"}]

- 流量、水質在同一條箭頭兩端應該完全一致

【質量平衡關係】
- 同單元: Σ 所有進流 Q = Σ 所有出流 Q (除非有蒸發/補水)
- T01-01 中和池可能有 6 條進流 (WTB01-01-1 ~ -6) 合流, 但只有 1 條出流 (WTA01-01-1)
  此時: Σ(WTB01-01-1..6) Q = WTA01-01-1 Q

【輸出格式 — JSON】
{
  "system_title": "圖標題, 例: T03 廢水處理系統",
  "page_note": "頁次資訊, 例: 9/87",
  "units": [
    {"code": "T03-05", "name": "第一氧化池"}
  ],
  "flows": [
    {
      "from_unit": "T03-04",         // 來源單元代號 (沒有上游就 null, WMxx = 廠外進入)
      "from_name": "氰系儲槽",
      "to_unit": "T03-05",            // 目的單元
      "to_name": "第一氧化池",
      "from_stream": "WTA03-04-1",   // 出流編號 (圖上沒就 null)
      "to_stream": "WTB03-05-1",      // 進流編號 (圖上沒就 null)
      "Q_cmd": 56.0125,               // 流量 (沒寫就 null)
      "notes": ""                      // 圖上特殊備註
    }
  ],
  "external_inputs": [
    {"code": "WM03", "name": "氰系廢液", "Q_cmd": 0.625, "to_unit": "T03-02", "to_stream": "WTB03-02-1"}
  ],
  "discharge_points": [
    {"code": "D01", "name": "放流口", "from_unit": "T01-02", "from_stream": "WTA01-02-1", "Q_cmd": 608.56}
  ],
  "annotations": ["其他關鍵備註, 例如 設計值 100% 呈現, 反洗水質計算等"]
}

【重要規則】
1. 每一條箭頭 = 一筆 flow (含上游 → 下游)
2. 若一條箭頭包含多個來源 (例 "來源包含 T02-08, T03-09, T04-06" 寫在一起), 拆成多筆 flow, 每個來源一筆, to_unit 都一樣
3. WMxx (廠外原廢水) 寫在 external_inputs 不寫在 flows
4. Dxx (放流口) 寫在 discharge_points 不寫在 flows
5. 不確定的數值留 null, 不要猜
6. 處理單元的 code 一律用 Txx-yy 格式 (兩位數)
7. 編號標籤跟流量盡量保留, 漏抽會嚴重影響質量平衡檢核

請輸出 JSON (絕對不要加 ```json``` 包裝, 直接給 JSON):
"""


# ──────────────────────────────────────────────────
# 找水量平衡圖的頁碼
# ──────────────────────────────────────────────────

def find_balance_diagram_pages(pdf_bytes):
    """從 PDF 找出所有「水量平衡示意圖」的圖片頁。

    判定方式:
        - 標題含「水量平衡示意圖」或「水質水量平衡示意圖」
        - 頁面文字 < 100 字 (主要是圖)

    Returns:
        {
            "ok": True,
            "image_pages": [11, 12, 13, ...],  # 純圖頁
            "title_pages": [11, 12, ...],       # 任何含此標題的頁
            "total_pages": N,
        }
    """
    try:
        import pdfplumber
    except ImportError as e:
        return {"ok": False, "error": f"缺少 pdfplumber: {e}"}

    try:
        image_pages = []
        title_pages = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                if "水量平衡示意圖" in text or "水質水量平衡" in text:
                    title_pages.append(i)
                    # 純圖頁: 文字 < 100 字 (排除水質資料表)
                    if len(text.strip()) < 100:
                        image_pages.append(i)
        return {
            "ok": True,
            "image_pages": image_pages,
            "title_pages": title_pages,
            "total_pages": total,
        }
    except Exception as e:
        return {"ok": False, "error": f"PDF 解析失敗: {e}"}


# ──────────────────────────────────────────────────
# 渲染 PDF 頁為圖片
# ──────────────────────────────────────────────────

def render_page_as_image(pdf_bytes, page_number, resolution=200):
    """把 PDF 某一頁渲染成 PIL Image。

    Args:
        page_number: 1-based 頁碼
        resolution: DPI (200 約 1654x2339, 對 Gemini 已足夠)
    """
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if page_number < 1 or page_number > len(pdf.pages):
                return None
            page = pdf.pages[page_number - 1]
            im = page.to_image(resolution=resolution)
            # 把 pdfplumber 的 PageImage 轉成 PIL Image
            return im.original
    except Exception as e:
        raise RuntimeError(f"渲染 p{page_number} 失敗: {e}")


# ──────────────────────────────────────────────────
# Gemini Vision 抽取
# ──────────────────────────────────────────────────

def extract_one_page(pil_image, api_key=None):
    """對一張圖呼叫 Gemini Vision, 回傳結構化結果。"""
    if api_key is None:
        try:
            from gemini_extractor import _get_gemini_api_key
            api_key, _ = _get_gemini_api_key()
        except Exception:
            return {"ok": False, "error": "無法取得 Gemini API key"}
    if not api_key:
        return {"ok": False, "error": "未設定 Gemini API key"}

    try:
        import google.generativeai as genai
        from gemini_extractor import GEMINI_MODEL_CANDIDATES
    except Exception as e:
        return {"ok": False, "error": f"無法載入 google-generativeai: {e}"}

    genai.configure(api_key=api_key)

    last_err = None
    response = None
    used_model = None
    for candidate in GEMINI_MODEL_CANDIDATES:
        try:
            model = genai.GenerativeModel(candidate)
            response = model.generate_content(
                [EXTRACTION_PROMPT, pil_image],
                generation_config={
                    "temperature": 0.1,
                    "max_output_tokens": 8192,
                    "response_mime_type": "application/json",
                },
            )
            used_model = candidate
            break
        except Exception as e:
            last_err = e
            continue

    if response is None:
        return {"ok": False, "error": f"Gemini 失敗: {last_err}"}

    raw = response.text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        return {"ok": False, "error": f"JSON 解析失敗: {e}", "raw_response": raw[:500]}

    usage = {}
    if response.usage_metadata:
        usage = {
            "input_tokens": response.usage_metadata.prompt_token_count,
            "output_tokens": response.usage_metadata.candidates_token_count,
        }

    return {
        "ok": True,
        "data": parsed,
        "model": used_model,
        "usage": usage,
    }


# ──────────────────────────────────────────────────
# 主入口: 抽整份 PDF 的所有水量平衡圖
# ──────────────────────────────────────────────────

def extract_all_balance_diagrams(pdf_bytes, max_pages=None, progress_callback=None):
    """掃整份 PDF, 抽取所有水量平衡示意圖, 合併成完整流向結構。

    Args:
        pdf_bytes: PDF 原始 bytes
        max_pages: 限制處理的圖頁數 (預設不限)
        progress_callback: 進度回呼 fn(current, total, message)

    Returns:
        {
            "ok": True,
            "pages_processed": N,
            "all_units": [...],     # 合併所有頁的單元 (去重)
            "all_flows": [...],     # 合併所有頁的流
            "all_external_inputs": [...],
            "all_discharge_points": [...],
            "per_page_results": [{"page": int, "data": {...}}, ...],
            "errors": [{"page": int, "error": str}],
            "gemini_usage": {"input_tokens": N, "output_tokens": M},
        }
    """
    # 找頁
    loc = find_balance_diagram_pages(pdf_bytes)
    if not loc.get("ok"):
        return {"ok": False, "stage": "locate", "error": loc.get("error")}

    image_pages = loc["image_pages"]
    if not image_pages:
        return {"ok": False, "stage": "locate",
                "error": "找不到水量平衡示意圖頁 (純圖頁)"}

    if max_pages:
        image_pages = image_pages[:max_pages]

    # 認證
    try:
        from gemini_extractor import _get_gemini_api_key
        api_key, source = _get_gemini_api_key()
    except Exception as e:
        return {"ok": False, "stage": "auth", "error": str(e)}
    if not api_key:
        return {"ok": False, "stage": "auth", "error": "未設定 Gemini API key"}

    # 累積結果
    per_page = []
    errors = []
    units_set = {}  # code → name (去重)
    flows = []
    external = []
    discharge = []
    total_in_tokens = 0
    total_out_tokens = 0

    for idx, pn in enumerate(image_pages):
        if progress_callback:
            progress_callback(idx, len(image_pages), f"處理 p{pn}…")
        try:
            img = render_page_as_image(pdf_bytes, pn)
        except Exception as e:
            errors.append({"page": pn, "error": f"渲染失敗: {e}"})
            continue

        result = extract_one_page(img, api_key)
        if not result.get("ok"):
            errors.append({"page": pn, "error": result.get("error")})
            continue

        data = result["data"]
        per_page.append({
            "page": pn,
            "system_title": data.get("system_title", ""),
            "data": data,
        })
        for u in data.get("units", []):
            code = (u.get("code") or "").strip()
            if code:
                units_set.setdefault(code, u.get("name", ""))
        for f in data.get("flows", []):
            flows.append({**f, "_page": pn})
        for e in data.get("external_inputs", []):
            external.append({**e, "_page": pn})
        for d in data.get("discharge_points", []):
            discharge.append({**d, "_page": pn})

        u = result.get("usage", {})
        total_in_tokens += u.get("input_tokens") or 0
        total_out_tokens += u.get("output_tokens") or 0

    if progress_callback:
        progress_callback(len(image_pages), len(image_pages), "完成")

    return {
        "ok": True,
        "pages_processed": len(per_page),
        "image_pages_found": image_pages,
        "all_units": [{"code": c, "name": n} for c, n in sorted(units_set.items())],
        "all_flows": flows,
        "all_external_inputs": external,
        "all_discharge_points": discharge,
        "per_page_results": per_page,
        "errors": errors,
        "gemini_usage": {
            "input_tokens": total_in_tokens,
            "output_tokens": total_out_tokens,
        },
        "auth_source": source,
    }


# ──────────────────────────────────────────────────
# 跨單元質平檢核
# ──────────────────────────────────────────────────

def check_water_balance(extract_result):
    """從抽出的流向資料, 算每個單元的水量平衡 (Σ進 ≈ Σ出)。

    質量平衡邏輯:
        對每個處理單元 X:
            Σ 所有進流 (WTBxx-yy-*) 的 Q  ≈  Σ 所有出流 (WTAxx-yy-*) 的 Q
        外部進入 (WMxx) 算 in
        放流口 (Dxx) 算 out

    Returns:
        {
            "by_unit": {
                "T03-05": {
                    "in_total_cmd": ...,
                    "out_total_cmd": ...,
                    "diff_pct": ...,
                    "warning": str or None,
                }
            },
            "summary": {
                "balanced_count": N,    # |diff| < 1%
                "warning_count": M,     # |diff| 1~5%
                "error_count": K,       # |diff| > 5%
                "total_units": int,
            }
        }
    """
    flows = extract_result.get("all_flows", [])
    external = extract_result.get("all_external_inputs", [])
    discharge = extract_result.get("all_discharge_points", [])
    units = extract_result.get("all_units", [])

    # 算每個單元的進/出
    in_by_unit = {}  # code → total Q
    out_by_unit = {}

    # 記錄已從 flows 加過的「from-to-stream」三元組, 避免 external/discharge 重複算
    seen_flow_keys = set()

    for f in flows:
        q = f.get("Q_cmd")
        if q is None:
            continue
        try:
            q = float(q)
        except (TypeError, ValueError):
            continue
        to_u = f.get("to_unit")
        from_u = f.get("from_unit")
        if to_u:
            in_by_unit[to_u] = in_by_unit.get(to_u, 0) + q
        if from_u:
            out_by_unit[from_u] = out_by_unit.get(from_u, 0) + q
        # 記下這條 flow 的識別 (用 from-to-streams 三元組)
        key = (f.get("from_stream"), f.get("to_stream"), to_u)
        if key != (None, None, None):
            seen_flow_keys.add(key)

    # 外部進入算進 — 但若已在 flows 出現過, 跳過避免重複
    for e in external:
        q = e.get("Q_cmd")
        if q is None:
            continue
        try:
            q = float(q)
        except (TypeError, ValueError):
            continue
        to_u = e.get("to_unit")
        # 看看是否 flows 裡已有這條 (用 from_stream=WMxx 或 to_stream 對照)
        from_s = e.get("code") or e.get("from_stream")
        to_s = e.get("to_stream")
        key = (from_s, to_s, to_u)
        if key in seen_flow_keys:
            continue
        # 也檢查另一種對應 (Gemini 可能把 WMxx 寫成 from_unit)
        key2 = (None, to_s, to_u)
        if to_s and key2 in seen_flow_keys:
            continue
        if to_u:
            in_by_unit[to_u] = in_by_unit.get(to_u, 0) + q

    # 放流口算出 — 同樣避免重複
    for d in discharge:
        q = d.get("Q_cmd")
        if q is None:
            continue
        try:
            q = float(q)
        except (TypeError, ValueError):
            continue
        from_u = d.get("from_unit")
        from_s = d.get("from_stream")
        to_s = d.get("code") or d.get("to_stream")
        key = (from_s, to_s, d.get("code"))
        if key in seen_flow_keys:
            continue
        key2 = (from_s, None, None)
        if from_s and any(from_s == k[0] for k in seen_flow_keys):
            continue
        if from_u:
            out_by_unit[from_u] = out_by_unit.get(from_u, 0) + q

    by_unit = {}
    balanced = 0
    warned = 0
    errored = 0

    all_codes = set(in_by_unit) | set(out_by_unit) | {u["code"] for u in units}
    for code in sorted(all_codes):
        in_total = in_by_unit.get(code, 0)
        out_total = out_by_unit.get(code, 0)
        if in_total == 0 and out_total == 0:
            continue
        if in_total == 0:
            diff_pct = None
            warning = "只有出流, 沒抽到進流"
            errored += 1
        elif out_total == 0:
            diff_pct = None
            warning = "只有進流, 沒抽到出流"
            errored += 1
        else:
            diff_pct = (out_total - in_total) / in_total * 100
            if abs(diff_pct) < 1:
                warning = None
                balanced += 1
            elif abs(diff_pct) < 5:
                warning = f"輕微偏差 {diff_pct:+.1f}%"
                warned += 1
            else:
                warning = f"嚴重不平衡 {diff_pct:+.1f}%"
                errored += 1
        by_unit[code] = {
            "in_total_cmd": round(in_total, 3),
            "out_total_cmd": round(out_total, 3),
            "diff_pct": round(diff_pct, 2) if diff_pct is not None else None,
            "warning": warning,
        }

    return {
        "by_unit": by_unit,
        "summary": {
            "balanced_count": balanced,
            "warning_count": warned,
            "error_count": errored,
            "total_units": len(by_unit),
        }
    }


# ──────────────────────────────────────────────────
# 跨單元 stream 對應一致性檢核
# ──────────────────────────────────────────────────

# 解析 stream code 的 regex
_STREAM_RE = re.compile(r"^WT([AB])(\d{2})-(\d{2})-?(\d*)$")
_WM_RE = re.compile(r"^WM(\d+)$")        # 原廢水 (外部上游)
_DISCHARGE_RE = re.compile(r"^D(\d+)$")  # 放流口 (外部下游)


def _parse_stream_code(code):
    """把 WTA01-01-1 拆成 (kind='A', unit='T01-01', idx='1')。
    也支援 WMxx (kind='WM') 和 Dxx (kind='D')。失敗回 None。"""
    if not code:
        return None
    s = str(code).strip()
    m = _STREAM_RE.match(s)
    if m:
        return {"kind": m.group(1), "unit": f"T{m.group(2)}-{m.group(3)}",
                "idx": m.group(4) or "", "full": s}
    m = _WM_RE.match(s)
    if m:
        return {"kind": "WM", "unit": s, "idx": m.group(1), "full": s}
    m = _DISCHARGE_RE.match(s)
    if m:
        return {"kind": "D", "unit": s, "idx": m.group(1), "full": s}
    return None


def check_stream_consistency(extract_result):
    """跨單元 stream 對應一致性檢核。

    每條 flow 應該有 from_stream (WTAxx-yy-z) 跟 to_stream (WTBxx-yy-z),
    且 from_stream 跟 to_stream 是「同一條水流」, 它們在文件其他地方應該對得起來。

    檢核項目:
    1. 流量一致: 若多條 flow 的 from_stream / to_stream 相同, 它們的 Q 應一致
    2. 對應存在: 若 WTAxx-yy-z 出現在某條 flow 的 from_stream, 通常它在另一條 flow 的 to_stream 也應該對應到一個 WTBaa-bb-c
    3. 編號歸屬: WTAxx-yy-z 的歸屬單元應該 = 該 flow 的 from_unit (例 WTA01-01-1 屬於 T01-01)

    Returns:
        {
            "ok": True,
            "code_groups": {                  # 每個 stream code 出現的位置
                "WTA01-01-1": {
                    "as_from": [{flow_idx, Q_cmd}],  # 出現在某些 flow 的 from_stream
                    "as_to":   [{flow_idx, Q_cmd}],  # 出現在某些 flow 的 to_stream
                    "Q_set": [608.56, ...],          # 涉及的 Q 值集合 (應該都一樣)
                }
            },
            "warnings": [
                {"type": "...", "message": "...", "stream": "..."}
            ],
            "summary": {
                "total_streams": N,
                "warnings_count": M,
            }
        }
    """
    flows = extract_result.get("all_flows", [])
    code_groups = {}  # code -> {"as_from": [...], "as_to": [...], "Q_set": [...]}
    warnings = []

    def _add(code, role, flow_idx, q, other_unit):
        if not code:
            return
        if code not in code_groups:
            code_groups[code] = {"as_from": [], "as_to": [], "Q_set": []}
        code_groups[code][role].append({
            "flow_idx": flow_idx, "Q_cmd": q, "other_unit": other_unit
        })
        if q is not None:
            try:
                code_groups[code]["Q_set"].append(float(q))
            except (TypeError, ValueError):
                pass

    for i, f in enumerate(flows):
        q = f.get("Q_cmd")
        _add(f.get("from_stream"), "as_from", i, q, f.get("to_unit"))
        _add(f.get("to_stream"), "as_to", i, q, f.get("from_unit"))

    # 檢核 1: 同一 stream code 在多處出現時, Q 應一致
    for code, info in code_groups.items():
        q_vals = info["Q_set"]
        if len(q_vals) >= 2:
            # 看差異 (相對)
            mn, mx = min(q_vals), max(q_vals)
            if mn > 0:
                diff_pct = (mx - mn) / mn * 100
                if diff_pct > 1:  # > 1% 偏差
                    warnings.append({
                        "type": "stream_q_inconsistent",
                        "stream": code,
                        "message": f"流量不一致: {q_vals} (差 {diff_pct:.1f}%)",
                    })

    # 檢核 2: stream code 歸屬 vs flow 的 from/to_unit 是否一致
    for code, info in code_groups.items():
        parsed = _parse_stream_code(code)
        if not parsed:
            warnings.append({
                "type": "stream_code_invalid",
                "stream": code,
                "message": f"編號格式不認得 (應為 WTA/WTBxx-yy-z)",
            })
            continue
        expected_unit = parsed["unit"]
        kind = parsed["kind"]

        # WM/D 不檢查 kind, 它們可以同時當 from/to
        for ev in info["as_from"]:
            flow = flows[ev["flow_idx"]]
            from_u = flow.get("from_unit")
            if from_u and from_u != expected_unit and kind in ("A", "WM"):
                # WTA01-01-1 的 from_unit 應是 T01-01
                # WM08 的 from_unit 應是 WM08 (或 None / 跟自己同名)
                warnings.append({
                    "type": "stream_unit_mismatch",
                    "stream": code,
                    "message": f"{code} 標為 from_stream, 但該 flow 的 from_unit={from_u} (應為 {expected_unit})",
                })
            if kind == "B":
                warnings.append({
                    "type": "stream_kind_mismatch",
                    "stream": code,
                    "message": f"{code} 用為 from_stream (出流), 但編號是 WTB 開頭 (應為 WTA)",
                })

        for ev in info["as_to"]:
            flow = flows[ev["flow_idx"]]
            to_u = flow.get("to_unit")
            if to_u and to_u != expected_unit and kind in ("B", "D"):
                warnings.append({
                    "type": "stream_unit_mismatch",
                    "stream": code,
                    "message": f"{code} 標為 to_stream, 但該 flow 的 to_unit={to_u} (應為 {expected_unit})",
                })
            if kind == "A":
                warnings.append({
                    "type": "stream_kind_mismatch",
                    "stream": code,
                    "message": f"{code} 用為 to_stream (進流), 但編號是 WTA 開頭 (應為 WTB)",
                })

    # 檢核 3: 出現在 from_stream 的 WTA, 應該在另一條 flow 的 to_stream 對應到 WTB
    # (即跨單元連續性)
    unmatched_from = []
    unmatched_to = []
    for code, info in code_groups.items():
        parsed = _parse_stream_code(code)
        if not parsed:
            continue
        # WM 跟 D 本來就是「外部端點」, 不需要對應到另一端的 WT 編號 → 跳過
        if parsed["kind"] in ("WM", "D"):
            continue
        # WTA + 只出現在 as_from + 沒對應到 WTB → 可能漏抽下游的 to_stream
        if parsed["kind"] == "A" and info["as_from"] and not info["as_to"]:
            # 看這條 flow 的 to_stream 有沒有對應到任何 WTB
            for ev in info["as_from"]:
                flow = flows[ev["flow_idx"]]
                to_stream = flow.get("to_stream")
                if not to_stream:
                    unmatched_from.append({
                        "stream": code,
                        "from_unit": parsed["unit"],
                        "to_unit": flow.get("to_unit"),
                        "Q_cmd": ev["Q_cmd"],
                    })
        # WTB + 只出現在 as_to + 該 flow 沒 from_stream → 可能漏抽上游編號
        if parsed["kind"] == "B" and info["as_to"] and not info["as_from"]:
            for ev in info["as_to"]:
                flow = flows[ev["flow_idx"]]
                from_stream = flow.get("from_stream")
                if not from_stream:
                    unmatched_to.append({
                        "stream": code,
                        "to_unit": parsed["unit"],
                        "from_unit": flow.get("from_unit"),
                        "Q_cmd": ev["Q_cmd"],
                    })

    return {
        "ok": True,
        "code_groups": code_groups,
        "warnings": warnings,
        "unmatched_from": unmatched_from,  # 上游有 WTA 編號但下游 to_stream 漏抽
        "unmatched_to": unmatched_to,      # 下游有 WTB 編號但上游 from_stream 漏抽
        "summary": {
            "total_streams": len(code_groups),
            "warnings_count": len(warnings),
            "unmatched_from_count": len(unmatched_from),
            "unmatched_to_count": len(unmatched_to),
        }
    }


# ──────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print("用法: python flow_diagram_extractor.py <PDF 路徑> [max_pages]")
        sys.exit(0)

    pdf_path = sys.argv[1]
    max_pages = int(sys.argv[2]) if len(sys.argv) > 2 else None

    if not os.path.exists(pdf_path):
        print(f"找不到 {pdf_path}")
        sys.exit(1)

    with open(pdf_path, "rb") as f:
        data = f.read()

    def cb(cur, tot, msg):
        print(f"[{cur}/{tot}] {msg}")

    print(f"=== 找圖頁 ({pdf_path}) ===")
    loc = find_balance_diagram_pages(data)
    print(f"找到 image_pages: {loc.get('image_pages')}")
    print()

    print(f"=== 抽取 (max_pages={max_pages}) ===")
    result = extract_all_balance_diagrams(data, max_pages=max_pages, progress_callback=cb)
    if not result.get("ok"):
        print(f"失敗: {result.get('error')}")
        sys.exit(1)

    print()
    print(f"處理 {result['pages_processed']} 頁")
    print(f"抽出 {len(result['all_units'])} 單元 / {len(result['all_flows'])} 流向")
    print(f"Token: in={result['gemini_usage']['input_tokens']} out={result['gemini_usage']['output_tokens']}")

    if result["errors"]:
        print(f"錯誤 {len(result['errors'])} 頁:")
        for e in result["errors"]:
            print(f"  p{e['page']}: {e['error'][:80]}")

    # 質平檢核
    print()
    print("=== 水量平衡檢核 ===")
    bal = check_water_balance(result)
    s = bal["summary"]
    print(f"總計 {s['total_units']} 單元 / 平衡 {s['balanced_count']} / 警告 {s['warning_count']} / 異常 {s['error_count']}")
    for code, info in bal["by_unit"].items():
        if info.get("warning"):
            print(f"  {code}: in={info['in_total_cmd']} out={info['out_total_cmd']} ⚠️ {info['warning']}")
