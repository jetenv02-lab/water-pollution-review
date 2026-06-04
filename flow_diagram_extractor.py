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
  - WTA01-01-1 = 出流 (After)
  - WTB01-01-1 = 進流 (Before)
  - WMxx = 原廢水 (Wastewater 直接從製程進入廠區)
  - Dxx = 放流口 (Discharge)
- 流量: Q = XX CMD (Cubic Meters per Day)
- 質量數據 (可能標在箭頭旁): kg/day 等

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
    {"code": "WM03", "name": "氰系廢液", "Q_cmd": 0.625, "to_unit": "T03-02"}
  ],
  "discharge_points": [
    {"code": "D01", "name": "放流口", "from_unit": "T01-02", "Q_cmd": 608.56}
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

    # 外部進入算進
    for e in external:
        q = e.get("Q_cmd")
        if q is None:
            continue
        try:
            q = float(q)
        except (TypeError, ValueError):
            continue
        to_u = e.get("to_unit")
        if to_u:
            in_by_unit[to_u] = in_by_unit.get(to_u, 0) + q

    # 放流口算出
    for d in discharge:
        q = d.get("Q_cmd")
        if q is None:
            continue
        try:
            q = float(q)
        except (TypeError, ValueError):
            continue
        from_u = d.get("from_unit")
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
