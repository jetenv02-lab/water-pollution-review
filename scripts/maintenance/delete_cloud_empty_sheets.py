# -*- coding: utf-8 -*-
"""精準刪除雲端協作表的空殼分頁。

對比 sheets_sync.upload_xlsx_to_sheets():
    那個是「整本覆蓋」, 會 ws.clear() + ws.update() 所有分頁。
    雖然 Drive comments 理論上不會被刪, 但 anchor 可能 invalidate, 風險較高。

這支只做一件事: 刪除指定的分頁 (整個 sheet 移除)。
其他任何分頁的內容、格式、註解, 完全不動。

用法:
    python delete_cloud_empty_sheets.py            # dry run
    python delete_cloud_empty_sheets.py --apply    # 真執行
"""
import sys
import sheets_sync
import gspread
from google.oauth2.service_account import Credentials

TARGETS = ["文件類", "現場設備類"]   # 要從雲端刪掉的空殼分頁


def main():
    apply = "--apply" in sys.argv

    # 認證
    creds = Credentials.from_service_account_file(
        "service_account.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheets_sync.DEFAULT_SHEET_ID)

    existing = {ws.title: ws for ws in sh.worksheets()}
    print(f"雲端目前有 {len(existing)} 個分頁\n")

    to_delete = []
    for name in TARGETS:
        if name not in existing:
            print(f"  [SKIP] '{name}' 雲端沒有, 跳過")
            continue
        ws = existing[name]
        # 安全檢查: 確認真的是空的
        values = ws.get_all_values()
        # 扣掉表頭, 看有沒有資料
        n_data = sum(1 for r in values[1:] if any(c.strip() for c in r))
        if n_data == 0:
            to_delete.append(ws)
            print(f"  [DEL]  '{name}' (雲端 0 筆資料, 將刪除整個分頁)")
        else:
            print(f"  [KEEP] '{name}' (雲端 {n_data} 筆資料, 不刪)")

    if not to_delete:
        print("\n沒有需要刪除的分頁。")
        return

    if not apply:
        print(f"\n(dry-run) 預計刪除 {len(to_delete)} 個雲端分頁")
        print("加 --apply 才會真的執行。")
        return

    for ws in to_delete:
        sh.del_worksheet(ws)
        print(f"  已從雲端刪除: {ws.title}")
    print(f"\n完成! 雲端剩餘 {len(sh.worksheets())} 個分頁")
    print("其他分頁的內容、狀態欄、Drive 註解全部沒動到。")


if __name__ == "__main__":
    main()
