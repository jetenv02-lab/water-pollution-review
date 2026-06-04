# -*- coding: utf-8 -*-
"""審查紀錄落地到 Google Sheet 的 _審查紀錄 分頁。

設計:
    - 每次 Streamlit「開始完整審查」完成 → 自動 append 一列到 Sheet
    - 沿用現有 _審查紀錄 分頁的表頭 (6 欄): 審查文件 / 不合理數量 / 審查結果 / 審查時間 / 審查次數 / 比對結果檔案位置
    - 額外加 4 欄 (本次需要的): 待人工 / 處理單元數 / 耗時(秒) / 規則庫版本 / 同義字數
    - 失敗不致命 (沒設 service account 或 Sheet 沒分享, 都會 silent skip)
"""
import os
from datetime import datetime

DEFAULT_SHEET_ID = "1FOx4Wu1PVidbaC-89HBzyQSK7cBOOvxfu4RxLoBqwBQ"
WORKSHEET_NAME = "_審查紀錄"

# 完整欄位清單 (跟 Sheet 的對應)
HEADERS = [
    "審查文件",
    "不合理數量",
    "審查結果",
    "審查時間",
    "審查次數",
    "比對結果檔案位置",
    "待人工",        # 新增
    "處理單元數",     # 新增
    "耗時(秒)",       # 新增
    "規則庫版本",     # 新增 (commit hash 或 timestamp)
]


def _get_client():
    """取得 gspread client。失敗 raise RuntimeError。"""
    try:
        import sheets_sync
        return sheets_sync._get_gspread_client()
    except Exception as e:
        raise RuntimeError(f"無法取得 Sheets 連線: {e}")


def _get_or_create_worksheet(client, sheet_id):
    """打開 Sheet 並取得 _審查紀錄 分頁。

    若分頁不存在 → 建一個 + 寫表頭。
    若分頁存在但欄位數少於 HEADERS → 擴充表頭 (不動既有資料)。
    """
    sh = client.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
        # 檢查表頭是否需要擴充
        existing_headers = ws.row_values(1) if ws.row_count >= 1 else []
        if not existing_headers:
            ws.update(values=[HEADERS], range_name="A1")
        elif len(existing_headers) < len(HEADERS):
            # 補齊缺少的欄
            full_headers = list(existing_headers) + [
                h for h in HEADERS if h not in existing_headers
            ]
            ws.update(values=[full_headers], range_name="A1")
    except Exception:
        # 不存在 → 建
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=500, cols=len(HEADERS))
        ws.update(values=[HEADERS], range_name="A1")
    return sh, ws


def _count_review_times(ws, filename):
    """算這份文件已經被審查過幾次 (含這次)。"""
    try:
        all_data = ws.get_all_records()
        count = sum(1 for r in all_data if str(r.get("審查文件", "")).strip() == filename)
        return count + 1  # +1 = 本次
    except Exception:
        return 1


def append_review_record(record, sheet_id=None):
    """把一筆審查紀錄 append 到 Sheet 的 _審查紀錄 分頁。

    Args:
        record: dict, keys:
            - filename (str): 審查文件名
            - unreasonable (int): 不合理數量
            - manual (int): 待人工數量
            - units (int): 處理單元數
            - elapsed_sec (int): 耗時秒
            - result (str, 可選): 審查結果 (預設根據不合理數判斷)
            - report_path (str, 可選): 比對結果檔案位置
        sheet_id: 預設用 DEFAULT_SHEET_ID

    Returns:
        {"ok": bool, "review_times": int, "error": str}
    """
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    filename = record.get("filename", "")
    if not filename:
        return {"ok": False, "error": "缺 filename"}

    try:
        client, _source = _get_client()
        _sh, ws = _get_or_create_worksheet(client, sheet_id)

        # 算第幾次審查
        review_times = _count_review_times(ws, filename)

        # 自動判斷審查結果
        unreasonable = int(record.get("unreasonable", 0))
        manual = int(record.get("manual", 0))
        result = record.get("result")
        if not result:
            if unreasonable == 0 and manual == 0:
                result = "合格(自動)"
            elif unreasonable == 0:
                result = f"待人工 {manual} 項"
            else:
                result = f"不合理 {unreasonable} 項"

        # 規則庫版本: 用 rules_extracted.csv 最後修改時間
        rule_version = ""
        try:
            import sheets_sync
            csv_path = sheets_sync.RULES_CSV
            if os.path.exists(csv_path):
                mtime = os.path.getmtime(csv_path)
                rule_version = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

        row = [
            filename,
            unreasonable,
            result,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            review_times,
            record.get("report_path", ""),
            manual,
            int(record.get("units", 0)),
            int(record.get("elapsed_sec", 0)),
            rule_version,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")

        return {
            "ok": True,
            "review_times": review_times,
            "filename": filename,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def load_review_history(sheet_id=None, limit=100):
    """從 Sheet 讀 _審查紀錄 全部紀錄。

    Returns:
        {"ok": bool, "rows": [dict, ...], "total": int, "error": str}
    """
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    try:
        client, _ = _get_client()
        sh = client.open_by_key(sheet_id)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except Exception:
            return {"ok": True, "rows": [], "total": 0}
        records = ws.get_all_records()
        # 倒序 (最新在最上面)
        records = list(reversed(records))
        if limit > 0:
            records = records[:limit]
        return {"ok": True, "rows": records, "total": len(records)}
    except Exception as e:
        return {"ok": False, "rows": [], "total": 0, "error": str(e)}


# ──────────────────────────────────────────────────
# CLI 測試
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import io
    import json
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    cmd = sys.argv[1] if len(sys.argv) > 1 else "list"

    if cmd == "list":
        print("=== 載入歷史 ===")
        r = load_review_history(limit=10)
        if r["ok"]:
            print(f"總筆數 (前 10): {r['total']}")
            for row in r["rows"]:
                print(json.dumps(row, ensure_ascii=False))
        else:
            print(f"失敗: {r.get('error')}")
    elif cmd == "test-write":
        print("=== 測試寫一筆 ===")
        test = {
            "filename": "TEST_FILE.pdf",
            "unreasonable": 0,
            "manual": 3,
            "units": 38,
            "elapsed_sec": 120,
        }
        r = append_review_record(test)
        print(json.dumps(r, ensure_ascii=False, indent=2))
    else:
        print("用法: python review_history.py [list|test-write]")
