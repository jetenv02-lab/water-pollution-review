# -*- coding: utf-8 -*-
"""Google Sheets ←→ rules_extracted.csv 同步模組。

工作流:
    1. 上傳 (push): CSV → Sheets
       - 主檔在 GitHub 的 rules_extracted.csv 是 source of truth
       - 按按鈕推到 Sheet 給同事編輯
    2. 下載 (pull): Sheets → CSV
       - 同事在 Sheet 上改完後, 按按鈕拉回 CSV
       - 拉回前會自動備份成 backup/rules_YYYYMMDD_HHMMSS.xlsx

認證:
    Service Account JSON 從以下兩處讀 (依序):
    1. Streamlit secrets[gcp_service_account]  ← 線上版
    2. 本機 service_account.json               ← 本機開發
    都找不到 → 顯示申請流程說明

依賴: gspread, google-auth, openpyxl
"""
import csv
import json
import os
from datetime import datetime

# 預設 Sheet 設定
DEFAULT_SHEET_ID = "1FOx4Wu1PVidbaC-89HBzyQSK7cBOOvxfu4RxLoBqwBQ"
DEFAULT_WORKSHEET = "rules"  # 工作表分頁名稱

BASE = os.path.dirname(os.path.abspath(__file__))
RULES_CSV = os.path.join(BASE, "rules_extracted.csv")
BACKUP_DIR = os.path.join(BASE, "backup")
LOCAL_SA_PATH = os.path.join(BASE, "service_account.json")


# ──────────────────────────────────────────────────
# 認證
# ──────────────────────────────────────────────────

def _get_service_account_info():
    """從 Streamlit Secrets 或本機檔案取得 service account 資訊。

    Returns:
        (dict, str): (sa_info_dict, source) 其中 source ∈ {"streamlit", "local", None}
    """
    # 1) 嘗試 Streamlit secrets (線上版)
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"]), "streamlit"
    except Exception:
        pass

    # 2) 嘗試本機 service_account.json
    if os.path.exists(LOCAL_SA_PATH):
        try:
            with open(LOCAL_SA_PATH, "r", encoding="utf-8") as f:
                return json.load(f), "local"
        except Exception:
            pass

    return None, None


def _get_gspread_client():
    """取得已認證的 gspread client。

    Raises:
        RuntimeError: 找不到認證 / 套件未安裝 / 認證失敗
    """
    sa_info, source = _get_service_account_info()
    if not sa_info:
        raise RuntimeError(
            "找不到 Google Service Account 認證。\n\n"
            "請完成以下其中一種設定:\n"
            "1. 本機: 把 service_account.json 放在專案根目錄\n"
            "2. Streamlit Cloud: 在 Secrets 設定 [gcp_service_account] 區段\n\n"
            "申請流程請見 SHEETS_SETUP.md"
        )

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as e:
        raise RuntimeError(f"缺少套件: {e}. 請 pip install gspread google-auth")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    client = gspread.authorize(creds)
    return client, source


def get_service_account_email():
    """回傳 service account 的 email (給使用者去 Sheet 分享用)。"""
    sa_info, _ = _get_service_account_info()
    if sa_info:
        return sa_info.get("client_email", "")
    return None


def check_auth_status():
    """檢查目前的認證狀態, 給 UI 顯示。

    Returns:
        dict: {
            "ok": bool,
            "source": "streamlit" | "local" | None,
            "email": str | None,
            "message": str
        }
    """
    sa_info, source = _get_service_account_info()
    if not sa_info:
        return {
            "ok": False,
            "source": None,
            "email": None,
            "message": "未設定 Service Account",
        }
    email = sa_info.get("client_email", "(未知)")
    return {
        "ok": True,
        "source": source,
        "email": email,
        "message": f"已認證 (來源: {source})",
    }


# ──────────────────────────────────────────────────
# Sheets 工作表存取
# ──────────────────────────────────────────────────

def _get_or_create_worksheet(client, sheet_id, worksheet_name):
    """打開 Sheet, 找不到指定工作表就建一個。"""
    try:
        sh = client.open_by_key(sheet_id)
    except Exception as e:
        raise RuntimeError(
            f"無法打開 Sheet (ID: {sheet_id})\n"
            f"原因: {e}\n\n"
            f"請確認:\n"
            f"1. Sheet ID 正確\n"
            f"2. Sheet 已分享給 service account email (給編輯權限)"
        )

    try:
        ws = sh.worksheet(worksheet_name)
    except Exception:
        # 工作表不存在 → 建一個
        ws = sh.add_worksheet(title=worksheet_name, rows=500, cols=20)
    return sh, ws


# ──────────────────────────────────────────────────
# 上傳: CSV → Sheets
# ──────────────────────────────────────────────────

def upload_csv_to_sheets(sheet_id=None, worksheet_name=None, csv_path=None):
    """把 rules_extracted.csv 整份推到 Google Sheets。

    Returns:
        dict: {ok, rows_written, sheet_url, source, error}
    """
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    worksheet_name = worksheet_name or DEFAULT_WORKSHEET
    csv_path = csv_path or RULES_CSV

    if not os.path.exists(csv_path):
        return {"ok": False, "error": f"找不到 CSV: {csv_path}"}

    # 讀 CSV
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        return {"ok": False, "error": "CSV 是空的"}

    try:
        client, source = _get_gspread_client()
        sh, ws = _get_or_create_worksheet(client, sheet_id, worksheet_name)

        # 清空後整批寫入 (避免殘留舊資料)
        ws.clear()
        ws.update(values=rows, range_name="A1")

        return {
            "ok": True,
            "rows_written": len(rows) - 1,  # 扣掉表頭
            "cols_written": len(rows[0]) if rows else 0,
            "sheet_url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
            "source": source,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ──────────────────────────────────────────────────
# 下載: Sheets → CSV (自動備份)
# ──────────────────────────────────────────────────

def _backup_current_csv():
    """把現有 CSV 複製成 backup/rules_YYYYMMDD_HHMMSS.csv 和 .xlsx。"""
    if not os.path.exists(RULES_CSV):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # CSV 備份
    csv_bak = os.path.join(BACKUP_DIR, f"rules_{ts}.csv")
    with open(RULES_CSV, "rb") as src, open(csv_bak, "wb") as dst:
        dst.write(src.read())

    # XLSX 備份 (用 openpyxl 從 CSV 轉)
    xlsx_bak = os.path.join(BACKUP_DIR, f"rules_{ts}.xlsx")
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "rules"
        with open(RULES_CSV, "r", encoding="utf-8-sig", newline="") as f:
            for row in csv.reader(f):
                ws.append(row)
        wb.save(xlsx_bak)
    except Exception:
        xlsx_bak = None

    return {"csv": csv_bak, "xlsx": xlsx_bak}


def download_sheets_to_csv(sheet_id=None, worksheet_name=None, csv_path=None):
    """從 Sheet 拉資料覆寫 rules_extracted.csv (覆寫前自動備份)。

    Returns:
        dict: {ok, rows_read, backup, source, error}
    """
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    worksheet_name = worksheet_name or DEFAULT_WORKSHEET
    csv_path = csv_path or RULES_CSV

    try:
        client, source = _get_gspread_client()
        sh, ws = _get_or_create_worksheet(client, sheet_id, worksheet_name)

        # 取所有資料 (含表頭)
        all_values = ws.get_all_values()
        if not all_values:
            return {"ok": False, "error": "Sheet 是空的"}

        # 先備份
        backup = _backup_current_csv()

        # 寫回 CSV
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(all_values)

        return {
            "ok": True,
            "rows_read": len(all_values) - 1,
            "cols_read": len(all_values[0]) if all_values else 0,
            "backup": backup,
            "sheet_url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
            "source": source,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ──────────────────────────────────────────────────
# 比對: Sheet vs CSV (差異預覽, 不寫檔)
# ──────────────────────────────────────────────────

def preview_diff(sheet_id=None, worksheet_name=None, csv_path=None):
    """比對 Sheet 跟 CSV 的差異 — 用 "缺失ID" 當主鍵。

    Returns:
        dict: {
            ok, added (Sheet 比 CSV 多), removed (CSV 比 Sheet 多),
            changed (兩邊都有但內容不同), unchanged_count, error
        }
    """
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    worksheet_name = worksheet_name or DEFAULT_WORKSHEET
    csv_path = csv_path or RULES_CSV

    try:
        client, source = _get_gspread_client()
        sh, ws = _get_or_create_worksheet(client, sheet_id, worksheet_name)
        sheet_rows = ws.get_all_records()

        # 讀 CSV
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            csv_rows = list(csv.DictReader(f))

        # 建 ID → row 對照
        csv_by_id = {r.get("缺失ID", ""): r for r in csv_rows}
        sheet_by_id = {r.get("缺失ID", ""): r for r in sheet_rows}

        added = [rid for rid in sheet_by_id if rid not in csv_by_id]
        removed = [rid for rid in csv_by_id if rid not in sheet_by_id]
        changed = []
        unchanged = 0
        for rid in sheet_by_id:
            if rid in csv_by_id:
                c = csv_by_id[rid]
                s = sheet_by_id[rid]
                diffs = [k for k in s if str(s.get(k, "")) != str(c.get(k, ""))]
                if diffs:
                    changed.append({"id": rid, "fields": diffs})
                else:
                    unchanged += 1

        return {
            "ok": True,
            "added": added,
            "removed": removed,
            "changed": changed,
            "unchanged_count": unchanged,
            "source": source,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ──────────────────────────────────────────────────
# 列出備份
# ──────────────────────────────────────────────────

def list_backups():
    """列出 backup/ 下的所有備份檔, 依時間倒序。"""
    if not os.path.exists(BACKUP_DIR):
        return []
    items = []
    for f in os.listdir(BACKUP_DIR):
        if not (f.endswith(".csv") or f.endswith(".xlsx")):
            continue
        path = os.path.join(BACKUP_DIR, f)
        mtime = os.path.getmtime(path)
        items.append({
            "name": f,
            "path": path,
            "size_kb": round(os.path.getsize(path) / 1024, 1),
            "mtime": datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
        })
    items.sort(key=lambda x: x["mtime"], reverse=True)
    return items


# ──────────────────────────────────────────────────
# CLI 測試入口
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import io

    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    print("=== 認證狀態 ===")
    status = check_auth_status()
    print(status)

    if not status["ok"]:
        print("\n請先設定 Service Account。")
        sys.exit(0)

    cmd = sys.argv[1] if len(sys.argv) > 1 else "preview"

    if cmd == "upload":
        print("\n=== 上傳 CSV → Sheets ===")
        r = upload_csv_to_sheets()
        print(json.dumps(r, ensure_ascii=False, indent=2))
    elif cmd == "download":
        print("\n=== 下載 Sheets → CSV ===")
        r = download_sheets_to_csv()
        print(json.dumps(r, ensure_ascii=False, indent=2))
    elif cmd == "preview":
        print("\n=== 預覽差異 ===")
        r = preview_diff()
        print(json.dumps(r, ensure_ascii=False, indent=2))
    elif cmd == "backups":
        print("\n=== 備份清單 ===")
        for b in list_backups():
            print(f"  {b['mtime']}  {b['name']}  ({b['size_kb']} KB)")
    else:
        print(f"用法: python sheets_sync.py [upload|download|preview|backups]")
