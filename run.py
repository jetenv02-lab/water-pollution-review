# -*- coding: utf-8 -*-
"""水措審查系統 — 主流程。

一鍵跑完: python run.py [申請PDF路徑]
若不指定路徑,會掃 參考/需審查之文件/ 內所有 PDF。

流程:
  Step 1c → 從 rules_extracted.csv 更新規則庫.xlsx
  Step 2  → 抽取申請 PDF 為 JSON
  Step 3  → 規則 × 申請 → 比對結果
  Step 4  → 輸出 Excel + Word + 寫回審查紀錄
  Step 5  → (選用) 同步 Google Sheets

也可以個別執行:
  python step1b_import_review_pdf.py "新審查意見.pdf"  # 增加新審查意見
  python step1c_csv_to_xlsx.py                       # CSV 寫回 xlsx
  python step2_extract_application.py "申請.pdf"      # 抽申請資料
  python step3_compare.py                            # 比對
  python step4_output.py                             # 輸出報告
  python sync_to_sheets.py                           # 同步 Sheets
"""
import os
import sys
import subprocess

BASE = r"C:\Users\jeten\Desktop\AI\水措審查"
APP_DIR = os.path.join(BASE, "參考", "需審查之文件")


def run_step(script, *args):
    cmd = [sys.executable, os.path.join(BASE, script), *args]
    print(f"\n>>> 執行: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=BASE)
    if result.returncode != 0:
        print(f"!!! {script} 失敗 (returncode={result.returncode})")
        return False
    return True


def main():
    if len(sys.argv) >= 2:
        app_pdfs = [sys.argv[1]]
    else:
        if not os.path.exists(APP_DIR):
            print(f"找不到 {APP_DIR}")
            return
        app_pdfs = [os.path.join(APP_DIR, f) for f in os.listdir(APP_DIR) if f.lower().endswith(".pdf")]
        if not app_pdfs:
            print(f"在 {APP_DIR} 找不到 PDF")
            return

    print("=" * 60)
    print("水措審查系統 — 主流程")
    print("=" * 60)
    print(f"待審查 PDF: {len(app_pdfs)} 份")
    for p in app_pdfs:
        print(f"  - {p}")

    # Step 1c: CSV → xlsx (確保規則庫是最新的)
    if not run_step("step1c_csv_to_xlsx.py"):
        print("Step 1c 失敗,中止")
        return

    # 對每份申請 PDF 跑 Step 2-4
    for pdf in app_pdfs:
        print(f"\n{'=' * 60}\n處理: {pdf}\n{'=' * 60}")
        if not run_step("step2_extract_application.py", pdf):
            continue
        base = os.path.splitext(os.path.basename(pdf))[0]
        app_json = os.path.join(BASE, f"application_{base}.json")
        if not os.path.exists(app_json):
            print(f"!!! 找不到 step2 輸出 {app_json}")
            continue
        if not run_step("step3_compare.py", app_json):
            continue
        comp_json = os.path.join(BASE, f"comparison_application_{base}.json")
        if not os.path.exists(comp_json):
            print(f"!!! 找不到 step3 輸出 {comp_json}")
            continue
        run_step("step4_output.py", comp_json)

    print("\n" + "=" * 60)
    print("主流程完成!")
    print("=" * 60)
    print("輸出位置: 審查報告/")
    print("如要同步到 Google Sheets,請執行: python sync_to_sheets.py")


if __name__ == "__main__":
    main()
