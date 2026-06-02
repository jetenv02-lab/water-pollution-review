# 🌐 線上版部署指南

> 把水措審查系統部署到 **Streamlit Cloud**，任何人連網址就能用，免費。

## 線上版特性

| 項目 | 內容 |
|------|------|
| 平台 | Streamlit Cloud (官方免費方案) |
| 入口檔 | `streamlit_app.py` |
| 網址 | 部署後會有 `https://你的app名稱.streamlit.app` |
| 費用 | 免費 |
| 更新方式 | 每次 `git push` 自動重新部署 |
| 隱私 | 上傳 PDF **不保存**，session 結束即刪除 |

## 部署步驟（你做）

### Step 1：到 Streamlit Cloud 註冊

1. 開 https://streamlit.io/cloud
2. 點 **Sign up** → 選 **Continue with GitHub**
3. 用你的 GitHub 帳號 (`jetenv02-lab`) 登入並授權

### Step 2：新建 App

1. 登入後點 **New app** (右上)
2. 選擇 **From existing repo**
3. 填入：
   - **Repository**: `jetenv02-lab/water-pollution-review`
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`
   - **App URL** (選填): 例如 `water-pollution-review`
4. 點 **Deploy!**

### Step 3：等待部署完成

- 第一次部署約需 3-5 分鐘（安裝 pdfplumber、openpyxl 等套件）
- 部署完成會顯示 **You're all set!** 並給你網址
- 之後你 `git push` 任何變更都會自動觸發重新部署（30 秒內）

### Step 4：分享網址

部署完成後，網址會是類似：
```
https://water-pollution-review.streamlit.app
```

把這個網址分享給協作者（技師、同事），他們不用裝任何東西就能用。

## 線上版能做什麼

✅ 上傳申請 PDF → 即時抽取處理單元 → 比對規則 → 下載 Excel/Word 報告
✅ 瀏覽 53 筆規則庫內容、依槽體篩選、關鍵字搜尋
✅ 看系統說明、常見問題
✅ 隨時更新規則（編輯 `rules_extracted.csv` push → 自動上線）

## 線上版不能做什麼

❌ 寫回審查紀錄到 `規則庫.xlsx`（Streamlit Cloud 是無狀態）
❌ 跑 Google Sheets 同步（需要 service account 金鑰）
❌ 處理超大檔案（>100MB 會逾時，建議用本地版）

## 三種使用情境

| 情境 | 用哪個版本 |
|------|----------|
| 快速試用、demo、給人看 | **線上版**（Streamlit Cloud） |
| 處理機密資料、大量檔案 | **本地版 Streamlit**：`streamlit run streamlit_app.py` |
| 完整工作流程、累積審查紀錄 | **本地版 Flask + 命令列**：`python web_app.py` 或 `python run.py` |

## 自訂網址

Streamlit Cloud 預設網址是 `<repo-name>.streamlit.app`。
若想改網址，App 設定 → **General** → 改 **App URL**。

## 升級到付費方案？

免費方案足夠一般使用。若需要：
- 私有 App（限定人員存取）
- 更高運算資源
- Custom domain（如 `review.yourcompany.com`）

可升級至 Streamlit Teams 方案。

## 疑難排解

### 部署失敗：`ModuleNotFoundError`
→ 檢查 `requirements.txt` 是否含該套件

### PDF 上傳逾時
→ 檔案太大。線上版上限 100MB，建議用本地版處理

### 中文字顯示亂碼
→ Streamlit 預設支援 UTF-8。如有亂碼是字型問題，可在 `.streamlit/config.toml` 改 `font`

### 規則庫沒更新
→ Streamlit Cloud 有 `@st.cache_data` 快取。重新部署或在 app 右上點 **Rerun**

## 相關連結

- GitHub Repo: https://github.com/jetenv02-lab/water-pollution-review
- Streamlit Cloud: https://streamlit.io/cloud
- Streamlit 文件: https://docs.streamlit.io
