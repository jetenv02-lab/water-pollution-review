# -*- coding: utf-8 -*-
"""Step 2c: OCR 解析流向示意圖 / 水質水量平衡示意圖。

這些章節在原 PDF 通常是純圖片(pdfplumber 抽不到文字),
本模組用 RapidOCR 跑文字辨識,然後解析:
- 單元代號 (T01-03, T02-08, D01, WM01, WTB01-01-1...)
- 流量 (Q = 55 CMD, Q = 122.5 CMD...)
- 加藥量 (PAC 1284.7 kg/d, polymer 0.1% 770kg/d...)
- 含水率 (含水率 99.7%, 乾基 41.319 kg/day, 濕基 14218.38 kg/day)

依賴: rapidocr-onnxruntime, pdfplumber, Pillow
使用:
    from step2c_ocr_diagram import ocr_diagram_pages
    result = ocr_diagram_pages(pdf_path, page_numbers=[14,15,16,17,18,19])
"""
import io
import os
import re
import sys
from datetime import datetime

import pdfplumber

# ───────────────────────────── 正規式 ─────────────────────────────

# 單元代號 (含 OCR 容易誤讀的容錯)
UNIT_CODE_RE = re.compile(r"T\s?\d{1,2}\s?[-－]\s?\d{1,3}")
WATER_FLOW_RE = re.compile(r"(WM|WTA|WTB)\s?\d{1,2}(?:[-－]\d{1,2}[-－]?\d*[a-zA-Z]?)?")
DISCHARGE_RE = re.compile(r"D\s?\d{1,2}")

# 流量 Q = 55 CMD / Q=122.5 CMD / Q＝55公噸
FLOW_RATE_RE = re.compile(r"Q\s*[=＝]\s*(\d+(?:\.\d+)?)\s*(?:CMD|公噸|噸|m3/d)?", re.IGNORECASE)

# 加藥量 PAC 1284.7 kg/d / polymer 0.1% 770 kg/d
DOSE_RE = re.compile(
    r"(PAC|polymer|NaOH|H2SO4|FeCl3|碳酸鈉|碳酸鈣|硫酸|氫氧化鈉|尿素)\s*"
    r"(?:\(\d+(?:\.\d+)?%?\))?\s*"
    r"(\d+(?:\.\d+)?)\s*(kg/d|公斤/日|噸/日)?",
    re.IGNORECASE
)

# 含水率 含水率 99.7% / 含水率=85%
MOISTURE_RE = re.compile(r"含水率\s*[=＝:：]?\s*(\d+(?:\.\d+)?)\s*%")
DRY_BASE_RE = re.compile(r"乾基\s*[=＝:：]?\s*(\d+(?:\.\d+)?)\s*kg/?day", re.IGNORECASE)
WET_BASE_RE = re.compile(r"濕基\s*[=＝:：]?\s*(\d+(?:\.\d+)?)\s*kg/?day", re.IGNORECASE)


def normalize_text(text):
    """正規化 OCR 識別文字。"""
    if not text:
        return ""
    for d in ["－", "─", "—", "–", "‐", "‑"]:
        text = text.replace(d, "-")
    return text


def page_to_image(page, resolution=200):
    """把 PDF page 轉成 PIL Image(高解析度供 OCR)。"""
    img = page.to_image(resolution=resolution)
    return img.original  # PIL Image object


def run_ocr_on_image(pil_image, ocr_engine):
    """跑 RapidOCR,回傳 [(座標, 文字, 信心度), ...]。"""
    import numpy as np
    # RapidOCR 接受 numpy array
    img_array = np.array(pil_image)
    result, _ = ocr_engine(img_array)
    if result is None:
        return []
    return result  # [(box, text, score), ...]


def parse_ocr_results(ocr_lines, page_number):
    """從 OCR 文字列表中,抽出結構化資訊。

    Args:
        ocr_lines: [(box, text, score), ...] from RapidOCR
        page_number: 1-based 頁碼

    Returns:
        dict: {
          "units_found": [{code, x, y, text}],
          "flows": [{from_code, q, x, y, text}],
          "doses": [{chemical, amount, unit, x, y, text}],
          "moistures": [{value_pct, x, y, text}],
          "raw_lines": [text, ...]
        }
    """
    units = []
    flows = []
    doses = []
    moistures = []
    raw_lines = []

    for entry in ocr_lines:
        if entry is None or len(entry) < 2:
            continue
        box = entry[0]
        text = entry[1]
        text = normalize_text(text)
        raw_lines.append(text)

        # 取中心點
        if box and len(box) == 4:
            x = sum(p[0] for p in box) / 4
            y = sum(p[1] for p in box) / 4
        else:
            x, y = 0, 0

        # 單元代號
        for m in UNIT_CODE_RE.finditer(text):
            code = re.sub(r"\s+", "", m.group(0))
            units.append({"code": code, "x": x, "y": y, "text": text})

        # 水流標籤
        for m in WATER_FLOW_RE.finditer(text):
            code = re.sub(r"\s+", "", m.group(0))
            units.append({"code": code, "x": x, "y": y, "text": text})

        # 放流口
        for m in DISCHARGE_RE.finditer(text):
            code = re.sub(r"\s+", "", m.group(0))
            units.append({"code": code, "x": x, "y": y, "text": text})

        # 流量
        for m in FLOW_RATE_RE.finditer(text):
            flows.append({
                "q": float(m.group(1)),
                "x": x, "y": y, "text": text
            })

        # 加藥
        for m in DOSE_RE.finditer(text):
            doses.append({
                "chemical": m.group(1),
                "amount": float(m.group(2)),
                "unit": m.group(3) or "",
                "x": x, "y": y, "text": text
            })

        # 含水率
        for m in MOISTURE_RE.finditer(text):
            moistures.append({
                "value_pct": float(m.group(1)),
                "x": x, "y": y, "text": text
            })

    # 去重(同代號多次出現只記一次,以 OCR 信心度較高者為準)
    seen_codes = set()
    units_unique = []
    for u in units:
        if u["code"] not in seen_codes:
            seen_codes.add(u["code"])
            units_unique.append(u)

    return {
        "page": page_number,
        "units_found": units_unique,
        "flows": flows,
        "doses": doses,
        "moistures": moistures,
        "raw_line_count": len(raw_lines),
    }


def ocr_diagram_pages(pdf_path, page_numbers, verbose=True):
    """對指定的 PDF 頁面跑 OCR,回傳結構化結果。

    Args:
        pdf_path: PDF 路徑
        page_numbers: 要 OCR 的頁碼清單 (1-based)
        verbose: 是否印進度

    Returns:
        {
          "ocr_engine": "RapidOCR",
          "pages": [parse_ocr_results 結果, ...],
          "all_units": [所有頁去重的單元清單],
          "summary": {total_units, total_flows, total_doses, ...}
        }
    """
    # Lazy import + 初始化 - 用 Exception 而非 ImportError 抓全部錯誤
    try:
        from rapidocr_onnxruntime import RapidOCR
        if verbose:
            print(f"初始化 RapidOCR engine...")
        ocr_engine = RapidOCR()
    except ImportError as e:
        return {
            "error": f"無法 import rapidocr_onnxruntime: {e}. 請執行 pip install rapidocr-onnxruntime",
        }
    except Exception as e:
        import traceback
        return {
            "error": f"RapidOCR 初始化失敗 ({type(e).__name__}): {e}\n{traceback.format_exc()[:500]}",
        }

    if verbose:
        print(f"OCR 目標頁: {page_numbers}")

    pages_results = []
    with pdfplumber.open(pdf_path) as pdf:
        for pn in page_numbers:
            if pn < 1 or pn > len(pdf.pages):
                if verbose:
                    print(f"  跳過頁 {pn} (超出範圍)")
                continue
            if verbose:
                print(f"  頁 {pn} OCR 中...", flush=True)
            try:
                pil_img = page_to_image(pdf.pages[pn - 1], resolution=200)
                ocr_lines = run_ocr_on_image(pil_img, ocr_engine)
                parsed = parse_ocr_results(ocr_lines, pn)
                pages_results.append(parsed)
                if verbose:
                    print(f"    識別 {parsed['raw_line_count']} 行文字, "
                          f"{len(parsed['units_found'])} 單元, "
                          f"{len(parsed['flows'])} 流量, "
                          f"{len(parsed['doses'])} 加藥, "
                          f"{len(parsed['moistures'])} 含水率")
            except Exception as e:
                if verbose:
                    print(f"    OCR 失敗: {e}")

    # 跨頁彙整
    all_units_dict = {}
    all_flows = []
    all_doses = []
    all_moistures = []
    for pr in pages_results:
        for u in pr["units_found"]:
            if u["code"] not in all_units_dict:
                all_units_dict[u["code"]] = u
        all_flows.extend(pr["flows"])
        all_doses.extend(pr["doses"])
        all_moistures.extend(pr["moistures"])

    summary = {
        "total_pages_ocr": len(pages_results),
        "total_units": len(all_units_dict),
        "total_flows": len(all_flows),
        "total_doses": len(all_doses),
        "total_moistures": len(all_moistures),
    }

    return {
        "ocr_engine": "RapidOCR",
        "ocr_at": datetime.now().isoformat(),
        "pages": pages_results,
        "all_units": list(all_units_dict.values()),
        "all_flows": all_flows,
        "all_doses": all_doses,
        "all_moistures": all_moistures,
        "summary": summary,
    }


def main():
    BASE = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = os.path.join(BASE, "參考", "需審查之文件")

    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
        # 第二參數是頁碼 (逗號分隔)
        if len(sys.argv) >= 3:
            page_numbers = [int(p) for p in sys.argv[2].split(",")]
        else:
            # 預設用章節定位器找
            from step2b_locate_sections import locate_sections
            sections = locate_sections(pdf_path)
            page_numbers = sorted(set(
                sections.get("flow_diagram", []) +
                sections.get("balance_diagram", [])
            ))
    else:
        if not os.path.exists(APP_DIR):
            print(f"找不到 {APP_DIR}")
            return
        pdfs = [f for f in os.listdir(APP_DIR) if f.lower().endswith(".pdf")]
        if not pdfs:
            print("找不到 PDF")
            return
        pdf_path = os.path.join(APP_DIR, pdfs[0])
        from step2b_locate_sections import locate_sections
        sections = locate_sections(pdf_path)
        page_numbers = sorted(set(
            sections.get("flow_diagram", []) +
            sections.get("balance_diagram", [])
        ))

    import io as _io
    try:
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    print(f"=== OCR 解析: {os.path.basename(pdf_path)} ===")
    print(f"目標頁: {page_numbers}\n")

    result = ocr_diagram_pages(pdf_path, page_numbers)
    if "error" in result:
        print(f"錯誤: {result['error']}")
        return

    print("\n=== OCR 結果摘要 ===")
    for k, v in result["summary"].items():
        print(f"  {k}: {v}")

    print("\n=== 識別到的單元 ===")
    for u in result["all_units"]:
        print(f"  {u['code']} (頁文字: {u['text'][:30]})")

    print("\n=== 識別到的流量 ===")
    for f in result["all_flows"][:20]:
        print(f"  Q={f['q']} (頁文字: {f['text'][:50]})")

    print("\n=== 識別到的加藥 ===")
    for d in result["all_doses"][:20]:
        print(f"  {d['chemical']} = {d['amount']} {d['unit']} (頁文字: {d['text'][:50]})")

    print("\n=== 識別到的含水率 ===")
    for m in result["all_moistures"][:20]:
        print(f"  含水率 = {m['value_pct']}% (頁文字: {m['text'][:50]})")

    # 輸出 JSON
    import json
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = os.path.join(BASE, f"ocr_{base}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n已輸出 OCR JSON: {out_path}")


if __name__ == "__main__":
    main()
