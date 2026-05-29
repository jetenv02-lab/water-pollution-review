# -*- coding: utf-8 -*-
"""Step 2: 抽取申請 PDF 的「單元結構化資料」輸出成 JSON。

輸出格式:
{
  "source_pdf": "05 申請資料(秋棠).pdf",
  "total_pages": 421,
  "extracted_at": "2026-05-29T12:00:00",
  "units": {
    "T01-03": {
      "raw_code": "T01-03",
      "std_tank": "批次反應槽",
      "name_in_doc": "T01-03 批次反應槽",
      "design_params": { "pH": "6~11", "停留時間": "8.7", ... },
      "measure_params": { ... },
      "equipment": ["液位計", "攪拌機", ...],
      "influent": { "BOD": 158, "COD": 640, ... },
      "effluent": { ... },
      "pages_found": [15, 16, 17, ...]
    },
    ...
  }
}

依賴: pdfplumber
用法: python step2_extract_application.py "申請PDF路徑"
"""
import json
import os
import re
import sys
from datetime import datetime
import pdfplumber

BASE = r"C:\Users\jeten\Desktop\AI\水措審查"
APP_DIR = os.path.join(BASE, "參考", "需審查之文件")

# 跟 step1b 共用的槽體歸一邏輯(複製貼上避免相互 import 出錯)
TANK_KEYWORDS = [
    ("pH調整", "pH調整槽"), ("pH 調整", "pH調整槽"),
    ("廢水調整", "廢水調整池"), ("調勻", "調勻池"),
    ("快混", "快混槽"), ("慢混", "慢混池"),
    ("沉澱", "沉澱池"), ("沉降", "沉澱池"),
    ("中和", "中和池"), ("放流", "放流池"),
    ("曝氣", "曝氣槽"), ("活性污泥", "曝氣槽"),
    ("砂濾", "砂濾塔"), ("過濾", "砂濾塔"),
    ("活性碳吸附", "活性碳吸附塔"), ("活性碳", "活性碳吸附裝置"),
    ("批次反應", "批次反應槽"), ("反應槽", "批次反應槽"),
    ("脫水", "脫水機"), ("壓濾", "脫水機"),
    ("貯留", "貯留槽"), ("暫存", "暫存槽"),
    ("污泥儲", "污泥儲槽"), ("污泥貯", "污泥儲槽"),
    ("濃縮", "濃縮槽"),
]

# 章節標題定位用
SECTION_MARKERS = {
    "design_params": ["三、處理單元名稱及操作參數", "(二)設計操作參數", "設計操作參數"],
    "measure_params": ["(三)量測操作參數", "量測操作參數"],
    "equipment": ["(四)相關機具設施", "相關機具設施"],
    "water_quality": ["水質水量平衡示意圖", "處理單元之進出水質"],
}

# 設計參數常見欄位
DESIGN_PARAM_FIELDS = [
    "pH", "停留時間", "有效容量", "有效水深", "表面溢流率", "溢流率",
    "DO", "溶氧", "MLSS", "F/M", "食微比", "污泥迴流率", "迴流率",
    "加藥量", "濾速", "上升流速", "負荷", "體積負荷", "有機負荷",
    "進流水量", "處理量", "操作時間", "批次數"
]

CODE_PATTERN = re.compile(r"T\d{1,2}-\d{1,3}[a-zA-Z]?|T\d{2}|D\d{1,2}|WM\d{1,2}|WTB\d{1,2}(?:-\d{1,2}[a-z]?)?")


def normalize_code(code):
    return code.replace("–", "-").strip()


def guess_tank(text):
    for kw, std in TANK_KEYWORDS:
        if kw in text:
            return std
    return None


def detect_unit_blocks(pages_text):
    """掃 PDF 每頁,找出每個處理單元的所在頁碼。
    回傳: { "T01-03": {"std_tank": "批次反應槽", "name_in_doc": "...", "pages": [15,16]} }
    """
    units = {}
    for i, text in enumerate(pages_text):
        if not text:
            continue
        norm = text.replace("–", "-")
        # 尋找 T01-XX 之類的代號
        for m in CODE_PATTERN.finditer(norm):
            code = normalize_code(m.group(0))
            # 取代號後方 30 字當作槽體名稱線索
            ctx_start = m.end()
            ctx = norm[ctx_start:ctx_start + 30]
            std = guess_tank(ctx)
            if not std:
                # 沒找到就跳過(避免噪音)
                continue
            if code not in units:
                units[code] = {
                    "raw_code": code,
                    "std_tank": std,
                    "name_in_doc": (code + " " + ctx.split()[0] if ctx else code),
                    "pages_found": [],
                    "design_params": {},
                    "measure_params": {},
                    "equipment": [],
                    "influent": {},
                    "effluent": {},
                }
            if (i + 1) not in units[code]["pages_found"]:
                units[code]["pages_found"].append(i + 1)
    return units


def extract_params_near_code(pages_text, units):
    """對每個單元,從其所在頁面萃取設計參數鍵值對。
    啟發式: 找『pH 6~9』『停留時間 8.7 小時』之類的 pattern。
    """
    param_pattern = re.compile(
        r"(pH|停留時間|有效容量|有效水深|表面溢流率|溢流率|DO|溶氧|MLSS|F/M|食微比|"
        r"污泥迴流率|迴流率|加藥量|濾速|上升流速|體積負荷|有機負荷|處理量)"
        r"\s*[:：]?\s*([\d\.~\-—–]+(?:\s*[a-zA-Z/°%μ]+)?)"
    )
    for code, info in units.items():
        for pg in info["pages_found"]:
            text = pages_text[pg - 1] or ""
            for m in param_pattern.finditer(text):
                key = m.group(1)
                val = m.group(2).strip()
                # 同一個參數可能出現多次,保留第一個
                if key not in info["design_params"]:
                    info["design_params"][key] = val


def extract_application(pdf_path):
    print(f"=== 抽取申請 PDF: {pdf_path} ===")
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        print(f"總頁數: {total}")
        for i, page in enumerate(pdf.pages):
            try:
                pages_text.append(page.extract_text() or "")
            except Exception as e:
                print(f"  [頁 {i+1}] extract_text 失敗: {e}")
                pages_text.append("")
            if (i + 1) % 50 == 0:
                print(f"  已處理 {i+1}/{total} 頁")

    units = detect_unit_blocks(pages_text)
    print(f"偵測到 {len(units)} 個處理單元")

    extract_params_near_code(pages_text, units)

    result = {
        "source_pdf": os.path.basename(pdf_path),
        "total_pages": len(pages_text),
        "extracted_at": datetime.now().isoformat(),
        "units": units,
    }

    # 輸出 JSON
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = os.path.join(BASE, f"application_{base}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"已輸出 → {out_path}")

    # 印一個摘要
    print("\n=== 摘要 ===")
    for code, info in sorted(units.items()):
        print(f"  {code} ({info['std_tank']}) - 頁 {info['pages_found'][:3]}{'...' if len(info['pages_found']) > 3 else ''}, "
              f"參數: {list(info['design_params'].keys())[:5]}")


def main():
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
    else:
        # 預設掃 參考/需審查之文件/
        pdfs = [f for f in os.listdir(APP_DIR) if f.lower().endswith(".pdf")]
        if not pdfs:
            print(f"在 {APP_DIR} 找不到 PDF")
            return
        pdf_path = os.path.join(APP_DIR, pdfs[0])

    if not os.path.exists(pdf_path):
        print(f"找不到 PDF: {pdf_path}")
        return
    extract_application(pdf_path)


if __name__ == "__main__":
    main()
