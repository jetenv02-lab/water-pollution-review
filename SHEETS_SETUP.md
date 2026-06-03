# Google Sheets 同步設定指引

> 一次性設定,完成後本機+Streamlit Cloud 兩邊都能用「規則庫管理」分頁的同步按鈕。

---

## 為什麼要做這個

- 規則庫主檔在 `rules_extracted.csv` (GitHub)
- 兩位同事用 Google Sheet 線上協作編輯
- 你按按鈕做雙向同步:CSV ↔ Sheet
- 不用再寄 Excel 給同事/同事回傳檔案合併

---

## 你的 Sheet 資訊

- Sheet URL: <https://docs.google.com/spreadsheets/d/1FOx4Wu1PVidbaC-89HBzyQSK7cBOOvxfu4RxLoBqwBQ/edit>
- Sheet ID: `1FOx4Wu1PVidbaC-89HBzyQSK7cBOOvxfu4RxLoBqwBQ`
- 工作表分頁名稱: `rules` (程式會自動建)

---

## Step 1 — 在 GCP 申請 Service Account

### 1.1 開啟 GCP Console
- 去 <https://console.cloud.google.com/>
- 用你的 Google 帳號登入

### 1.2 建立專案 (如果還沒有)
- 左上角專案下拉 → 「新增專案」
- 名字隨意,例如 `water-review-sync`
- 建完後切換到這個專案

### 1.3 啟用兩個 API
依序開啟,每個都按「啟用」:
1. <https://console.cloud.google.com/apis/library/sheets.googleapis.com>
2. <https://console.cloud.google.com/apis/library/drive.googleapis.com>

### 1.4 建立 Service Account
- 去 <https://console.cloud.google.com/iam-admin/serviceaccounts>
- 「+ 建立服務帳戶」
- 名稱: `water-review-bot` (隨意)
- 描述: 留空或填「水措審查系統 Sheet 同步」
- 角色: **不用選**(直接跳過,Sheet 權限走 Sheet 端分享)
- 「完成」

### 1.5 下載 JSON 金鑰
- 點剛建好的 service account
- 上方「金鑰」分頁
- 「新增金鑰」→「建立新的金鑰」→「JSON」→下載
- **這個檔案就是 `service_account.json`,複製到專案根目錄**
- ⚠️ 重要:此檔案已在 `.gitignore`,**絕對不要 push 到 GitHub**

### 1.6 記下 Service Account Email
下載的 JSON 裡有一行:
```json
"client_email": "water-review-bot@xxxx.iam.gserviceaccount.com"
```
**這個 email 等下要分享給 Sheet**。

---

## Step 2 — 把 Sheet 分享給 Service Account

1. 打開你的 Sheet: <https://docs.google.com/spreadsheets/d/1FOx4Wu1PVidbaC-89HBzyQSK7cBOOvxfu4RxLoBqwBQ/edit>
2. 右上「共用」
3. 把 Step 1.6 的 email 貼進去
4. 權限給「**編輯者**」(因為要寫入)
5. 取消勾選「通知這些人」(它不是真人,通知會 bounce)
6. 「分享」

---

## Step 3 — 本機設定 (option A)

把 Step 1.5 下載的 JSON 直接放在專案根目錄,檔名改為:
```
C:\Users\jeten\Desktop\AI\水措審查\service_account.json
```

驗證:
```bash
cd C:\Users\jeten\Desktop\AI\水措審查
python sheets_sync.py
```
應該顯示:
```
{'ok': True, 'source': 'local', 'email': 'water-review-bot@...', 'message': '已認證 (來源: local)'}
```

---

## Step 4 — Streamlit Cloud 設定 (option B,要線上也能同步才需要)

1. 去 <https://share.streamlit.io/> → 找到你的 app
2. 右下角「⋯」→「Settings」→「Secrets」
3. 在內容貼上以下格式 (用你下載的 JSON 內容填):

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "xxxx"
private_key = """-----BEGIN PRIVATE KEY-----
xxxxxxxxxxx
-----END PRIVATE KEY-----
"""
client_email = "water-review-bot@xxxx.iam.gserviceaccount.com"
client_id = "12345"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
universe_domain = "googleapis.com"
```

⚠️ `private_key` 的 `\n` 換行字元要保留,用 `"""..."""` 三引號包起來。

4. 儲存 → app 會自動重啟

---

## Step 5 — 測試

### 本機 CLI 測試
```bash
python sheets_sync.py preview   # 比對差異 (不寫檔)
python sheets_sync.py upload    # CSV → Sheet
python sheets_sync.py download  # Sheet → CSV (會自動備份)
python sheets_sync.py backups   # 看備份清單
```

### 線上版 UI 測試
1. 打開 <https://water-review.streamlit.app/>
2. 左側分頁切到「📋 規則庫管理」
3. 按「⬆️ 上傳 CSV → Sheet」
4. 同事就能在 Sheet 上看到 299 筆規則

---

## 工作流範例

### 你新增規則 → 推給同事
1. 修改 `rules_extracted.csv` (本機或 Sheet 都可以)
2. 若改本機,git push → 按「上傳到 Sheet」
3. 同事在 Sheet 上看到最新

### 同事改完 → 收回主檔
1. 同事在 Sheet 上編輯 (多人即時協作)
2. 你按「⬇️ 從 Sheet 下載 → 覆寫 CSV」
3. 系統自動備份舊 CSV + XLSX 到 `backup/`
4. 你 git push 把新 CSV 上 GitHub

### 出包還原
1. 看 `backup/` 目錄
2. 把某個 `rules_YYYYMMDD_HHMMSS.csv` 複製覆寫 `rules_extracted.csv`
3. git push

---

## 常見問題

**Q: 同步會不會把同事正在編的內容蓋掉?**
A: 會。下載前先問同事是否在編,或用「預覽差異」先看。系統會自動備份。

**Q: Service Account 可以用我的個人 Google 帳號嗎?**
A: 不行。Service Account 是專門的「機器帳號」,跟你個人帳號分開。但你可以在 GCP 用個人帳號建專案來託管它。

**Q: 一定要兩邊都設定嗎?**
A: 不用。本機 + 線上版選一邊就好,但你要的「兩邊都能按」就需要兩邊都設定。

**Q: 費用?**
A: 完全免費。Sheets API 每天 300 次/分鐘,我們一天可能按 5 次。
