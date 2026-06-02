# -*- coding: utf-8 -*-
"""Step 2b: 章節動態定位器。

不同申請 PDF 的頁碼會浮動（秋棠的水量平衡圖在頁 9-13，別家可能在頁 5-7 或 20-25）。
本模組用「章節標題」尋找各區段所在頁面，回傳:

{
  "flow_diagram": [12, 13, 14, 15, 16],         # 流向示意圖頁(通常是圖片,需 OCR)
  "balance_diagram": [12, 13, 14, 15, 16],       # 水量平衡示意圖頁
  "quality_data": [19, 20, 21, ...],             # 進出水質資料頁(文字)
  "facility_table": [72, 73, 76, 81, 82, ...],   # 處理設施資料表頁(文字)
  "raw_water": [11],                             # 原廢水水質頁
  "discharge": [73],                             # 放流口頁
  "emergency": [76],                             # 緊急應變頁
  "sludge": [...],                               # 污泥處理頁
}

對應的環保署標準章節標題:
- 參、水污染防治措施資料/廢(污)水產生與水污染防治措施流向示意圖
- 參、水污染防治措施資料/水質水量平衡示意圖
- 參、水污染防治措施資料/廢(污)水(前)處理設施資料表
- 參、水污染防治措施資料/處理單元之進出水質資料

使用:
    from step2b_locate_sections import locate_sections
    sections = locate_sections(pdf_path)
    print(sections["balance_diagram"])  # 例 [12, 13, 14, 15, 16]
"""
import os
import re
import sys
from collections import defaultdict
import pdfplumber

# 章節標題識別規則 — key 是區段類型,value 是該區段的關鍵字 list
# 一頁含任一關鍵字即視為該區段
SECTION_PATTERNS = {
    # 流向示意圖 (純圖片,需 OCR)
    "flow_diagram": [
        "廢(污)水產生與水污染防治措施流向示意圖",
        "廢污水產生與水污染防治措施流向示意圖",
        "處理流向示意圖",
        "廢水流向示意圖",
    ],
    # 水量平衡示意圖 (純圖片,需 OCR)
    "balance_diagram": [
        "水質水量平衡示意圖",
        "水量平衡示意圖",
        "水質水量平衡圖",
    ],
    # 進出水質資料表 (文字)
    "quality_data": [
        "處理單元之進出水質資料",
        "處理單元之進出水質",
        "進出水質資料",
    ],
    # 處理設施資料表 (文字)
    "facility_table": [
        "廢(污)水(前)處理設施資料表",
        "廢污水前處理設施資料表",
        "處理設施資料表",
        "處理單元名稱及操作參數",
    ],
    # 原廢水水質
    "raw_water": [
        "原廢水水質",
        "原廢(污)水水質",
        "廢水水質資料",
    ],
    # 放流口
    "discharge": [
        "放流口資料表",
        "放流口",
    ],
    # 緊急應變
    "emergency": [
        "緊急應變方法",
        "緊急應變",
    ],
    # 污泥處理
    "sludge": [
        "污泥處理",
        "污泥清運計畫",
    ],
}


def normalize(text):
    """正規化文字: 半形化破折號 + 移除空白 (用於關鍵字比對)。"""
    if not text:
        return ""
    for d in ["－", "─", "—", "–", "‐", "‑"]:
        text = text.replace(d, "-")
    # 移除空白 (中文章節標題在 PDF 抽取時常被空白切散)
    return re.sub(r"\s+", "", text)


def locate_sections(pdf_path, verbose=False):
    """掃整份 PDF,找出各章節所在頁碼。

    Returns:
        dict: {section_type: [page_1, page_2, ...]} (1-based 頁碼)
    """
    sections = defaultdict(list)
    page_titles = []  # [(page_idx, [titles_found])]

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if verbose:
            print(f"總頁數: {total}")
            print("掃描章節標題...")

        for i, page in enumerate(pdf.pages):
            try:
                text = page.extract_text() or ""
            except Exception as e:
                if verbose:
                    print(f"  頁 {i+1} extract_text 失敗: {e}")
                continue

            norm = normalize(text)
            titles_in_page = []

            for section_type, keywords in SECTION_PATTERNS.items():
                for kw in keywords:
                    if normalize(kw) in norm:
                        sections[section_type].append(i + 1)  # 1-based
                        titles_in_page.append((section_type, kw))
                        break  # 每個 section type 在同頁只記一次

            if titles_in_page:
                page_titles.append((i + 1, titles_in_page))

    if verbose:
        print(f"\n=== 章節定位結果 ===")
        for section_type in SECTION_PATTERNS.keys():
            pages = sections.get(section_type, [])
            if pages:
                # 合併連續頁
                ranges = compress_ranges(pages)
                print(f"  {section_type}: 共 {len(pages)} 頁 → {ranges}")
            else:
                print(f"  {section_type}: 未找到")

        print(f"\n=== 各頁標題 (前 20 頁) ===")
        for p, titles in page_titles[:20]:
            tnames = [f"{st}({kw[:15]}...)" if len(kw) > 15 else f"{st}({kw})" for st, kw in titles]
            print(f"  頁 {p}: {tnames}")

    return dict(sections)


def compress_ranges(pages):
    """[1,2,3,5,6,8] → '1-3, 5-6, 8' 方便閱讀。"""
    if not pages:
        return ""
    pages = sorted(set(pages))
    ranges = []
    start = pages[0]
    prev = pages[0]
    for p in pages[1:]:
        if p == prev + 1:
            prev = p
        else:
            ranges.append(f"{start}-{prev}" if start != prev else str(start))
            start = p
            prev = p
    ranges.append(f"{start}-{prev}" if start != prev else str(start))
    return ", ".join(ranges)


def main():
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
    else:
        # 預設掃 參考/需審查之文件/
        APP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "參考", "需審查之文件")
        pdfs = [f for f in os.listdir(APP_DIR) if f.lower().endswith(".pdf")] if os.path.exists(APP_DIR) else []
        if not pdfs:
            print("用法: python step2b_locate_sections.py <PDF路徑>")
            return
        pdf_path = os.path.join(APP_DIR, pdfs[0])

    if not os.path.exists(pdf_path):
        print(f"找不到 PDF: {pdf_path}")
        return

    print(f"=== 章節動態定位: {os.path.basename(pdf_path)} ===\n")
    sections = locate_sections(pdf_path, verbose=True)

    # 額外提示 (避免 emoji 在 cp950 終端機炸掉)
    import io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass
    print("\n=== 提示 ===")
    if sections.get("balance_diagram"):
        pages = compress_ranges(sections["balance_diagram"])
        print(f"[水量平衡示意圖] 在頁 {pages} (純圖片,需 OCR 才能讀取數據)")
    if sections.get("quality_data"):
        pages = compress_ranges(sections["quality_data"])
        print(f"[進出水質資料] 在頁 {pages} (文字,可直接抽取)")
    if sections.get("facility_table"):
        pages = compress_ranges(sections["facility_table"])
        print(f"[處理設施資料表] 在頁 {pages} (文字,可直接抽取)")


if __name__ == "__main__":
    main()
