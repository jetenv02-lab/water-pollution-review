# -*- coding: utf-8 -*-
"""Step 1b: 把新的審查意見 PDF 匯入規則庫。

未來流程:
1) 把新的審查意見 PDF 放到 參考/審查意見/ 資料夾
2) 此程式會 dump 該 PDF 全文 → 存成 deficiency_dump_{檔名}.txt
3) 用啟發式規則把文字切成「逐條缺失」(以 (1) (2) (3)... 為界)
4) 自動猜測標準槽體名稱、原始槽體代號、檢查類型
5) 追加寫入 rules_extracted.csv (來源代號自動編 S02、S03...)
6) 接著請手動跑 step1c 把 CSV 寫入 規則庫.xlsx

備註: 萃取品質有限,適合作為「初稿」,人工再修。若要精緻萃取請用 LLM。

依賴: pdfplumber
用法: python step1b_import_review_pdf.py "PDF路徑" [來源代號]
若不指定 PDF 路徑,會掃 參考/審查意見/ 內所有 PDF。
"""
import csv
import os
import re
import sys
from datetime import datetime
import pdfplumber

BASE = r"C:\Users\jeten\Desktop\AI\水措審查"
REVIEW_DIR = os.path.join(BASE, "參考", "審查意見")
CSV_OUT = os.path.join(BASE, "rules_extracted.csv")

CSV_HEADERS = ["缺失ID", "來源", "技師姓名", "序號", "原文缺失", "檢查類型",
               "對照項目", "規則", "比對位置", "判定邏輯",
               "標準槽體名稱", "原始槽體代號"]

# 槽體關鍵字 → 標準名稱(順序很重要,優先匹配前面的)
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

# 檢查類型關鍵字
TYPE_KEYWORDS = [
    (["pH", "停留時間", "有效容量", "溢流率", "DO", "MLSS", "迴流", "加藥量",
      "F/M", "食微比", "負荷", "濾速", "有效水深"], "設計參數"),
    (["液位計", "流量計", "馬達", "攪拌", "pH計", "管線", "泵"], "機具設施"),
    (["質量平衡", "質平", "進出流", "濃度", "計算有誤", "守恆"], "質量平衡"),
    (["去除率", "去除效果", "去除"], "去除率"),
    (["流向", "示意圖", "標示"], "流向/示意圖"),
    (["不一致", "誤植", "與.+不符", "未填", "未列", "缺"], "文件一致性"),
    (["有效位數", "單位"], "單位/有效位數"),
    (["放流水標準", "原廢水"], "水質標準"),
]

# 槽體代號正則: T01-03 / T01–03 / T02 / D01 / WM01 / WTB01 / S01-1 等
CODE_PATTERN = re.compile(r"(?:T\d{1,2}[-–]?\d{0,3}[a-zA-Z]?|D\d{1,2}|WM\d{1,2}|WTB\d{1,2}(?:-\d{1,2}[a-z]?)?|M\d{1,2})")


def dump_pdf(pdf_path):
    """傾印 PDF 文字到同目錄的 dump 檔。"""
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    dump_path = os.path.join(BASE, f"deficiency_dump_{base}.txt")
    with pdfplumber.open(pdf_path) as pdf, open(dump_path, "w", encoding="utf-8") as f:
        f.write(f"總頁數: {len(pdf.pages)}\n")
        for i, page in enumerate(pdf.pages):
            f.write("=" * 70 + "\n")
            f.write(f"### 第 {i+1} 頁 ###\n")
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                if ln.strip():
                    f.write("  " + ln + "\n")
    print(f"Dump → {dump_path}")
    return dump_path


def guess_tank(text):
    for kw, std in TANK_KEYWORDS:
        if kw in text:
            return std
    return "(文件類)"


def guess_type(text):
    for kws, tname in TYPE_KEYWORDS:
        for kw in kws:
            if re.search(kw, text):
                return tname
    return "其他"


def extract_codes(text):
    """抓 T01-03 / D01 / WM01 等代號;en-dash 轉成 hyphen。"""
    matches = CODE_PATTERN.findall(text.replace("–", "-"))
    seen = []
    for m in matches:
        m2 = m.replace("–", "-")
        if m2 not in seen:
            seen.append(m2)
    return " ; ".join(seen)


def next_deficiency_id():
    """掃現有 CSV,回傳下一個 Dxxx 流水號。"""
    if not os.path.exists(CSV_OUT):
        return 1
    max_id = 0
    with open(CSV_OUT, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            did = row.get("缺失ID", "")
            m = re.match(r"D(\d+)", did)
            if m:
                max_id = max(max_id, int(m.group(1)))
    return max_id + 1


def next_source_id():
    """掃現有 CSV,回傳下一個 Sxx 來源代號。"""
    if not os.path.exists(CSV_OUT):
        return "S02"  # S01 = 第一份審查意見
    max_id = 1
    with open(CSV_OUT, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            src = row.get("來源", "")
            m = re.match(r"S(\d+)", src)
            if m:
                max_id = max(max_id, int(m.group(1)))
    return f"S{max_id + 1:02d}"


def split_into_deficiencies(dump_text):
    """把 dump 文字粗略切成「逐條缺失」。
    啟發式: 以 (1) (2) (3)... 編號當作分隔點。
    """
    # 把多頁 dump 合併成單一文字流
    lines = []
    for ln in dump_text.split("\n"):
        ln = ln.strip()
        if not ln or ln.startswith("=") or ln.startswith("###") or ln.startswith("總頁數"):
            continue
        lines.append(ln)
    text = " ".join(lines)
    # 用 (1) ~ (99) 當分隔
    parts = re.split(r"\((\d{1,2})\)", text)
    # parts = [前言, "1", 第1條, "2", 第2條, ...]
    items = []
    for i in range(1, len(parts) - 1, 2):
        num = parts[i]
        body = parts[i + 1].strip()
        if len(body) > 10:  # 過濾雜訊
            items.append((num, body))
    return items


def import_pdf(pdf_path, source_id=None):
    print(f"\n=== 匯入: {pdf_path} ===")
    dump_path = dump_pdf(pdf_path)
    with open(dump_path, "r", encoding="utf-8") as f:
        dump_text = f.read()

    items = split_into_deficiencies(dump_text)
    print(f"切出 {len(items)} 條候選缺失")

    sid = source_id or next_source_id()
    src_label = f"{sid} {os.path.basename(pdf_path)}"

    next_id = next_deficiency_id()
    rows = []
    for num, body in items:
        # 截斷過長的原文(避免單筆超過 Excel 32k 字元限制)
        body_clean = body[:2000]
        tank = guess_tank(body_clean)
        ctype = guess_type(body_clean)
        codes = extract_codes(body_clean)
        rows.append({
            "缺失ID": f"D{next_id:03d}",
            "來源": src_label,
            "技師姓名": "(待確認填)",
            "序號": f"{sid} ({num})",
            "原文缺失": body_clean,
            "檢查類型": ctype,
            "對照項目": "",
            "規則": "",
            "比對位置": "",
            "判定邏輯": "",
            "標準槽體名稱": tank,
            "原始槽體代號": codes,
        })
        next_id += 1

    # 追加到 CSV
    write_header = not os.path.exists(CSV_OUT)
    with open(CSV_OUT, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        if write_header:
            w.writeheader()
        w.writerows(rows)
    print(f"已追加 {len(rows)} 筆至 {CSV_OUT}")
    print(f"來源代號: {sid}")
    print("提醒: 自動萃取結果僅為初稿,請人工檢視並填寫『規則』『判定邏輯』欄位。")
    print("提醒: 接著請執行 python step1c_csv_to_xlsx.py 把新資料寫入 規則庫.xlsx。")


def main():
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
        source_id = sys.argv[2] if len(sys.argv) >= 3 else None
        if not os.path.exists(pdf_path):
            print(f"找不到 PDF: {pdf_path}")
            return
        import_pdf(pdf_path, source_id)
        return

    # 沒指定就掃 參考/審查意見/ 內所有 PDF
    if not os.path.exists(REVIEW_DIR):
        print(f"找不到目錄: {REVIEW_DIR}")
        return
    pdfs = [os.path.join(REVIEW_DIR, f) for f in os.listdir(REVIEW_DIR) if f.lower().endswith(".pdf")]
    if not pdfs:
        print("沒找到 PDF")
        return
    print(f"找到 {len(pdfs)} 個 PDF")
    for p in pdfs:
        # 已經處理過第一份(水污染業務查核缺失.pdf)就跳過
        if "水污染業務查核缺失" in os.path.basename(p):
            print(f"跳過(已處理): {p}")
            continue
        import_pdf(p)


if __name__ == "__main__":
    main()
