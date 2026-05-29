# 水措審查系統

> 自動化「水污染防治措施」申請文件審查工具。根據環工技師查核缺失資料庫，比對申請文件中各處理單元（如 T01-03 批次反應槽、T01-04 中和槽）的設計參數、水質數據、機具設施等是否合理。

## 功能

- **規則庫**：累積環工技師審查缺失，依「標準槽體類型」分類（pH 調整槽、廢水調整池、慢混池、沉澱池…）
- **自動萃取**：從申請 PDF 抽出各處理單元的設計參數、進出流水質
- **自動比對**：每個單元 × 適用規則 → 觸發清單（不合理／合理／待人工）
- **報告輸出**：逐單元 Excel 比對表 + Word 審查意見書
- **審查紀錄**：累積同一案件多次審查的歷史
- **協作同步**：可同步到 Google Sheets 三人協作
- **Web UI**：本地啟動瀏覽器操作介面

## 系統需求

- Python 3.10 以上
- Windows / macOS / Linux

## 快速開始

```bash
# 1. clone 專案
git clone https://github.com/<your-username>/water-pollution-review.git
cd water-pollution-review

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 放入資料 (參考 下方「目錄結構」)
mkdir -p 參考/審查意見 參考/需審查之文件
# 把審查意見 PDF 放進 參考/審查意見/
# 把要審查的申請 PDF 放進 參考/需審查之文件/

# 4. 啟動 Web UI
python web_app.py
# 瀏覽器開啟 http://localhost:5000

# 或用命令列一鍵跑
python run.py "參考/需審查之文件/你的申請.pdf"
```

## 目錄結構

```
水措審查/
├── 參考/                            ← 輸入區 (不入版控)
│   ├── 審查意見/
│   │   └── 水污染業務查核缺失.pdf    ← 規則庫的原始來源
│   └── 需審查之文件/
│       └── *.pdf                    ← 待審查的申請文件
│
├── 審查報告/                         ← 輸出區 (不入版控)
│   ├── *_比對結果.xlsx
│   └── *_審查意見.docx
│
├── rules_extracted.csv              ← 規則中間檔
├── 規則庫.xlsx                       ← 規則主檔 (不入版控)
│
├── step1b_import_review_pdf.py      ← 匯入新審查意見 PDF
├── step1c_csv_to_xlsx.py            ← CSV → 規則庫.xlsx
├── step2_extract_application.py     ← 抽申請 PDF 結構化
├── step3_compare.py                 ← 比對引擎
├── step4_output.py                  ← 輸出 Excel + Word
├── sync_to_sheets.py                ← 同步 Google Sheets
├── run.py                           ← 命令列主流程
├── web_app.py                       ← Web UI (Flask)
└── templates/
    └── index.html                   ← 網頁介面
```

## 規則庫結構

`規則庫.xlsx` 採「**槽體類型分頁**」設計：

| 分頁 | 說明 |
|------|------|
| `_說明` | 使用說明 |
| `_來源清單` | 審查意見 PDF 來源 (S01, S02...) |
| `_槽體對照表` | 原始代號 (T01-03) ↔ 標準槽體名稱 |
| `_審查紀錄` | 每份申請的審查歷史 |
| `pH調整槽` / `廢水調整池` / `快混槽` / ... | 各槽體類型的規則 |
| `(文件類)` / `(現場設備類)` | 不特定槽體的規則 |

**每筆規則欄位**：
缺失ID｜來源｜原文缺失｜檢查類型｜對照項目｜規則｜比對位置｜判定邏輯｜技師姓名｜序號｜原始槽體代號

## 工作流程

```
       新審查意見 PDF
            │
            ▼
   step1b_import_review_pdf.py
            │
            ▼
   rules_extracted.csv  ←─── (可手動編修)
            │
            ▼
   step1c_csv_to_xlsx.py
            │
            ▼
       規則庫.xlsx ────────► sync_to_sheets.py ──► Google Sheets
            │
            │   申請 PDF
            │     │
            │     ▼
            │   step2_extract_application.py
            │     │
            │     ▼
            │   application_*.json
            │     │
            └─────▼
              step3_compare.py
                  │
                  ▼
              comparison_*.json
                  │
                  ▼
              step4_output.py
                  │
        ┌─────────┼──────────┐
        ▼         ▼          ▼
   比對結果.xlsx  審查意見.docx  寫回審查紀錄
```

## 「處理單元」概念

水污染防治措施裡的「處理單元」指廢水處理流程中的單一槽體/設備，常見編號：

- `T01-03` 批次反應槽
- `T01-04` 中和槽
- `T01-08` 沉澱池
- `T01-13` 砂濾塔
- `D01` 放流口
- `WM01`、`WTB01` 水流/水質代號

每個單元有自己的設計參數（有效容量、停留時間、表面溢流率、pH、加藥量、去除率等）。

## 去除率計算公式

```
單元去除率 (%) = (進流濃度 - 出流濃度) / 進流濃度 × 100%
質量去除率 (%) = (M_in - M_out) / M_in × 100%   其中 M = Q × C
```

**常見不合理狀況**（從查核缺失歸納）：

1. **快混槽展現重金屬去除率** — 快混只是加藥混合，沒固液分離
2. **溶解性物質出現去除率**（如導電度、Cl⁻、SO₄²⁻、Na⁺）— 一般處理不去除這些離子
3. **沉澱池前單元展現 SS 去除率** — SS 要靠重力沉降
4. **生物處理對重金屬有高去除率** — 生物處理主要去有機物
5. **質量不平衡** — 進流質量 ≠ 出流質量 + 污泥帶走

## 安全與隱私

- `參考/` 含敏感的事業申請資料，**已在 `.gitignore` 排除**
- `service_account.json` 等金鑰檔，**已在 `.gitignore` 排除**
- 推送前請執行 `git status` 確認沒有上傳 PDF

## 授權

MIT License — 詳見 [LICENSE](LICENSE)。

## 致謝

規則來源：環工技師執行水污染業務查核缺失彙整資料 (民國 111-114 年)。
