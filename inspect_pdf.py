# -*- coding: utf-8 -*-
"""探查秋棠申請資料 PDF:以「標題」定位兩個關鍵區段,印出該區段的文字與表格結構。"""
import sys
import io
import pdfplumber

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PDF = r"C:\Users\jeten\Desktop\AI\水措審查\參考\需審查之文件\05 申請資料(秋棠)(1150119)final.pdf"

# 想找的標題關鍵字(by heading, 不靠頁碼)
TARGETS = [
    "水質水量平衡示意圖",
    "處理單元之進出水質",
    "廢(污)水(前)處理設施資料表",
    "廢污水前處理設施資料表",
    "處理單元名稱及操作參數",
]

print("=== 開啟 PDF ===")
with pdfplumber.open(PDF) as pdf:
    total = len(pdf.pages)
    print(f"總頁數: {total}")
    print()

    # 第一階段:掃描每頁文字,找出含目標標題的頁碼
    hits = {}  # target -> [page_index,...]
    for i, page in enumerate(pdf.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            print(f"  [頁{i+1}] extract_text 失敗: {e}")
            continue
        norm = text.replace(" ", "").replace("\u3000", "")
        for t in TARGETS:
            key = t.replace(" ", "").replace("(", "(").replace(")", ")")
            # 同時嘗試全形/半形括號
            variants = {t, t.replace("(", "(").replace(")", ")"),
                        t.replace("(", "(").replace(")", ")")}
            if any(v.replace(" ", "") in norm for v in variants):
                hits.setdefault(t, []).append(i)

    print("=== 標題命中頁碼(1-based) ===")
    for t in TARGETS:
        pages = [p + 1 for p in hits.get(t, [])]
        print(f"  {t}: {pages}")
    print()
