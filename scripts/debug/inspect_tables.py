# -*- coding: utf-8 -*-
"""Phase 2:傾印秋棠 PDF 兩個關鍵區段的表格結構,供定案規則庫欄位與比對引擎。
   區段一:水質水量平衡示意圖 / 處理單元之進出水質 (約 p14-19)  → 槽體代號↔名稱、進出流水質
   區段二:廢(污)水(前)處理設施資料表 / 處理單元名稱及操作參數 (約 p68-139) → 操作參數、相關機具設施
"""
import sys
import io
import pdfplumber

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PDF = r"C:\Users\jeten\Desktop\AI\水措審查\參考\需審查之文件\05 申請資料(秋棠)(1150119)final.pdf"

# 先各取代表性頁面探查結構(不掃全部 421 頁,避免一次輸出太大)
SAMPLE_PAGES = [14, 15, 16, 17, 18, 19, 68, 72, 73, 81]  # 1-based


def dump_page(page, idx1):
    print("=" * 70)
    print(f"### 第 {idx1} 頁 ###")
    print("-" * 70)
    text = page.extract_text() or ""
    # 只印前 40 行文字,避免過長
    lines = [ln for ln in text.split("\n") if ln.strip()]
    print("【文字(前40行)】")
    for ln in lines[:40]:
        print("  " + ln)
    print()

    tables = page.extract_tables()
    print(f"【表格數: {len(tables)}】")
    for ti, tb in enumerate(tables):
        print(f"  -- 表格 {ti+1} (列數={len(tb)}) --")
        for ri, row in enumerate(tb[:25]):  # 每表最多印 25 列
            cells = [("" if c is None else c.replace("\n", "／")) for c in row]
            print(f"    R{ri}: {cells}")
        if len(tb) > 25:
            print(f"    ...(其餘 {len(tb)-25} 列省略)")
    print()


print("=== Phase 2: 傾印表格結構 ===")
with pdfplumber.open(PDF) as pdf:
    total = len(pdf.pages)
    print(f"總頁數: {total}")
    print(f"探查頁面(1-based): {SAMPLE_PAGES}")
    print()
    for p1 in SAMPLE_PAGES:
        if 1 <= p1 <= total:
            dump_page(pdf.pages[p1 - 1], p1)
