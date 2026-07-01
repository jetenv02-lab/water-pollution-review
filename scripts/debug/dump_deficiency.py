# -*- coding: utf-8 -*-
"""傾印『水污染業務查核缺失.pdf』全文,供萃取更多審查意見寫入規則庫。"""
import sys
import io
import pdfplumber

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PDF = r"C:\Users\jeten\Desktop\AI\水措審查\參考\審查意見\水污染業務查核缺失.pdf"

with pdfplumber.open(PDF) as pdf:
    print(f"總頁數: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        print("=" * 70)
        print(f"### 第 {i+1} 頁 ###")
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            if ln.strip():
                print("  " + ln)
