# -*- coding: utf-8 -*-
"""審查快照儲存 + 階段歸檔管理。

設計:
    每次審查匯出時, 可選擇「儲存本次內部覆核快照」, 把 Excel + JSON 各存一份到
    review_runs/ 資料夾, 並回傳 (run_id, report_path, json_path) 給呼叫端
    寫到 _審查紀錄 Sheets 的 run_id / report_path / json_path 欄。

資料夾結構:
    review_runs/
    ├── 邑昇_20260610_1630_R001.xlsx     ← 0~30 天 (完整 Excel + JSON)
    ├── 邑昇_20260610_1630_R001.json
    ├── 馥廷_20260610_1700_R002.xlsx
    ├── 馥廷_20260610_1700_R002.json
    │
    └── archived/                          ← 30~90 天 (壓縮)
        ├── 秋棠_20260510_1400_R000.zip   (含 .xlsx + .json)
        │
        └── deep/                          ← 90+ 天 (只留 .json)
            └── 邑昇_20260301_1000_R-old.json

對外 API:
    save_snapshot(target_bytes, json_bytes, base_name, target="internal", findings_count=...) → dict
    cleanup_old_snapshots(reviews_dir=None) → dict
    generate_run_id(filename) → str
    list_snapshots(reviews_dir=None) → list
"""
from __future__ import annotations

import json
import os
import shutil
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

try:
    from zoneinfo import ZoneInfo
    TPE_TZ = ZoneInfo("Asia/Taipei")
except Exception:
    from datetime import timezone
    TPE_TZ = timezone(timedelta(hours=8))


# ──────────────────────────────────────────────────
# 路徑
# ──────────────────────────────────────────────────

DEFAULT_REVIEWS_DIR = Path(__file__).parent / "review_runs"
ARCHIVED_DIR_NAME = "archived"
DEEP_DIR_NAME = "deep"

# 階段歸檔閾值
FRESH_DAYS = 30   # 0~30 天 不動
ARCHIVE_DAYS = 90  # 30~90 天 壓縮成 zip
# 90+ 天: 解壓 → 刪 .xlsx → 只留 .json


def _now_tpe():
    return datetime.now(TPE_TZ)


def _ensure_dir(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_reviews_dir(custom_dir: Optional[Path] = None) -> Path:
    """取得 review_runs 目錄, 若不存在則建立。"""
    d = Path(custom_dir) if custom_dir else DEFAULT_REVIEWS_DIR
    _ensure_dir(d)
    _ensure_dir(d / ARCHIVED_DIR_NAME)
    _ensure_dir(d / ARCHIVED_DIR_NAME / DEEP_DIR_NAME)
    return d


# ──────────────────────────────────────────────────
# run_id 產生
# ──────────────────────────────────────────────────

def generate_run_id(filename: str, reviews_dir: Optional[Path] = None) -> str:
    """產生 run_id, 格式 R0001 / R0002 ...

    用 review_runs/ 中已有的快照數量 + 1 決定。
    若有 1234 個快照, 就用 R1235。
    """
    d = get_reviews_dir(reviews_dir)
    # 找所有 R\d+ 命名的檔
    existing_nums = set()
    import re
    pattern = re.compile(r"_R(\d{4,})", re.IGNORECASE)
    for f in d.glob("*"):
        m = pattern.search(f.stem)
        if m:
            existing_nums.add(int(m.group(1)))
    # 也掃描 archived 內
    for f in (d / ARCHIVED_DIR_NAME).glob("*"):
        m = pattern.search(f.stem)
        if m:
            existing_nums.add(int(m.group(1)))
    for f in (d / ARCHIVED_DIR_NAME / DEEP_DIR_NAME).glob("*"):
        m = pattern.search(f.stem)
        if m:
            existing_nums.add(int(m.group(1)))

    next_num = (max(existing_nums) + 1) if existing_nums else 1
    return f"R{next_num:04d}"


# ──────────────────────────────────────────────────
# 主要 API: 存快照
# ──────────────────────────────────────────────────

def save_snapshot(
    excel_bytes: bytes,
    json_bytes: bytes,
    base_name: str,
    findings_count: dict = None,
    reviews_dir: Optional[Path] = None,
) -> dict:
    """儲存一次審查的快照 (Excel + JSON 配對)。

    Args:
        excel_bytes: internal Excel 的 bytes
        json_bytes: 整合 JSON 的 bytes
        base_name: 案件檔名前綴 (例: "邑昇")
        findings_count: {"不合理": N, "待確認": M, "錯誤": K} 用於 metadata
        reviews_dir: 自訂目錄, 預設 review_runs/

    Returns:
        {
            "ok": True,
            "run_id": "R0042",
            "report_path": "review_runs/邑昇_20260610_1630_R0042.xlsx",
            "json_path": "review_runs/邑昇_20260610_1630_R0042.json",
            "size_bytes": (excel+json sizes),
            "timestamp": "2026-06-10T16:30:00+08:00",
        }
    """
    d = get_reviews_dir(reviews_dir)
    ts = _now_tpe()
    ts_str = ts.strftime("%Y%m%d_%H%M")
    run_id = generate_run_id(base_name, d)

    safe_base = "".join(c if c.isalnum() or c in "_-" else "_" for c in base_name)
    stem = f"{safe_base}_{ts_str}_{run_id}"

    excel_path = d / f"{stem}.xlsx"
    json_path = d / f"{stem}.json"

    try:
        excel_path.write_bytes(excel_bytes)
        # JSON 多塞一層 metadata
        try:
            existing_data = json.loads(json_bytes.decode("utf-8"))
        except Exception:
            existing_data = {"_raw_failed_to_parse": True}
        existing_data["_snapshot"] = {
            "run_id": run_id,
            "timestamp": ts.isoformat(),
            "base_name": base_name,
            "findings_count": findings_count or {},
        }
        json_path.write_text(
            json.dumps(existing_data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        return {"ok": False, "error": f"快照儲存失敗: {e}"}

    return {
        "ok": True,
        "run_id": run_id,
        "report_path": str(excel_path.relative_to(Path(__file__).parent).as_posix()),
        "json_path": str(json_path.relative_to(Path(__file__).parent).as_posix()),
        "size_bytes": excel_path.stat().st_size + json_path.stat().st_size,
        "timestamp": ts.isoformat(),
    }


# ──────────────────────────────────────────────────
# 階段歸檔: 30~90 天 壓縮 / 90+ 天 留 JSON
# ──────────────────────────────────────────────────

def cleanup_old_snapshots(reviews_dir: Optional[Path] = None) -> dict:
    """執行階段歸檔. 可手動跑或排程跑。

    階段:
        0~30 天 (FRESH_DAYS): 不動
        30~90 天 (ARCHIVE_DAYS): xlsx + json 壓縮成 zip 移到 archived/
        90+ 天: 解壓 zip → 只留 json → 移到 archived/deep/

    Returns:
        {
            "ok": True,
            "archived_count": N,    # 30~90 天區的處理筆數
            "deep_count": M,        # 90+ 天區的處理筆數
            "freed_kb": K,          # 釋出空間 KB
        }
    """
    d = get_reviews_dir(reviews_dir)
    now = _now_tpe()
    archived_dir = d / ARCHIVED_DIR_NAME
    deep_dir = archived_dir / DEEP_DIR_NAME

    archived_count = 0
    deep_count = 0
    freed_bytes = 0

    # ── 階段 1: 0~30 天 → 30~90 天 (壓縮) ──
    # 找根目錄的 .xlsx, mtime 超過 FRESH_DAYS → 壓縮
    for xlsx in d.glob("*.xlsx"):
        try:
            mtime = datetime.fromtimestamp(xlsx.stat().st_mtime, TPE_TZ)
            age = (now - mtime).days
            if age < FRESH_DAYS:
                continue
            if age >= ARCHIVE_DAYS:
                continue  # 跳過, 給階段 2 處理 (此檔可能還沒進 archived)
            json_partner = xlsx.with_suffix(".json")
            stem = xlsx.stem
            zip_path = archived_dir / f"{stem}.zip"
            old_size = xlsx.stat().st_size + (json_partner.stat().st_size if json_partner.exists() else 0)
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(xlsx, arcname=xlsx.name)
                if json_partner.exists():
                    zf.write(json_partner, arcname=json_partner.name)
            new_size = zip_path.stat().st_size
            freed_bytes += old_size - new_size
            xlsx.unlink()
            if json_partner.exists():
                json_partner.unlink()
            archived_count += 1
        except Exception:
            continue

    # ── 階段 2: 30~90 天區的 zip → 90+ 天 (解壓只留 json) ──
    for zip_f in archived_dir.glob("*.zip"):
        try:
            mtime = datetime.fromtimestamp(zip_f.stat().st_mtime, TPE_TZ)
            age = (now - mtime).days
            if age < ARCHIVE_DAYS:
                continue
            # 解壓只取 .json
            stem = zip_f.stem
            json_target = deep_dir / f"{stem}.json"
            with zipfile.ZipFile(zip_f, "r") as zf:
                json_in_zip = [n for n in zf.namelist() if n.endswith(".json")]
                if json_in_zip:
                    with zf.open(json_in_zip[0]) as src, open(json_target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
            old_size = zip_f.stat().st_size
            zip_f.unlink()
            new_size = json_target.stat().st_size if json_target.exists() else 0
            freed_bytes += old_size - new_size
            deep_count += 1
        except Exception:
            continue

    # ── 階段 3: 根目錄已超過 90 天的 .xlsx (異常情況, 直接刪 xlsx 留 json 在根) ──
    # 一般不會發生 (階段 1 會處理掉), 防呆
    for xlsx in d.glob("*.xlsx"):
        try:
            mtime = datetime.fromtimestamp(xlsx.stat().st_mtime, TPE_TZ)
            age = (now - mtime).days
            if age < ARCHIVE_DAYS:
                continue
            # 90+ 天直接刪 xlsx, json 留著
            old_size = xlsx.stat().st_size
            xlsx.unlink()
            freed_bytes += old_size
            deep_count += 1
        except Exception:
            continue

    return {
        "ok": True,
        "archived_count": archived_count,
        "deep_count": deep_count,
        "freed_kb": round(freed_bytes / 1024, 1),
    }


# ──────────────────────────────────────────────────
# 列出快照 (給 UI 顯示用)
# ──────────────────────────────────────────────────

def list_snapshots(reviews_dir: Optional[Path] = None, limit: int = 50) -> list:
    """列出所有快照 (含 archived). 按時間倒序。

    Returns:
        [{
            "run_id": "R0042",
            "base_name": "邑昇",
            "timestamp": "...",
            "stage": "fresh" / "archived" / "deep",
            "size_kb": ...,
            "report_path": "...",
            "json_path": "...",
        }, ...]
    """
    d = get_reviews_dir(reviews_dir)
    items = []

    # 階段 1: 根目錄
    for xlsx in d.glob("*.xlsx"):
        json_partner = xlsx.with_suffix(".json")
        items.append({
            "stage": "fresh",
            "report_path": str(xlsx),
            "json_path": str(json_partner) if json_partner.exists() else None,
            "mtime": xlsx.stat().st_mtime,
            "size_kb": round((xlsx.stat().st_size + (json_partner.stat().st_size if json_partner.exists() else 0)) / 1024, 1),
        })

    # 階段 2: archived/
    for zip_f in (d / ARCHIVED_DIR_NAME).glob("*.zip"):
        items.append({
            "stage": "archived",
            "report_path": str(zip_f),
            "json_path": None,
            "mtime": zip_f.stat().st_mtime,
            "size_kb": round(zip_f.stat().st_size / 1024, 1),
        })

    # 階段 3: archived/deep/
    for js in (d / ARCHIVED_DIR_NAME / DEEP_DIR_NAME).glob("*.json"):
        items.append({
            "stage": "deep",
            "report_path": None,
            "json_path": str(js),
            "mtime": js.stat().st_mtime,
            "size_kb": round(js.stat().st_size / 1024, 1),
        })

    # 從檔名解析 base_name / timestamp / run_id
    import re
    for it in items:
        name = Path(it.get("report_path") or it.get("json_path") or "").stem
        m = re.match(r"(.+)_(\d{8}_\d{4})_(R\d+)", name)
        if m:
            it["base_name"] = m.group(1)
            it["timestamp"] = m.group(2)
            it["run_id"] = m.group(3)
        else:
            it["base_name"] = name
            it["timestamp"] = ""
            it["run_id"] = ""

    items.sort(key=lambda x: x["mtime"], reverse=True)
    return items[:limit]


# ──────────────────────────────────────────────────
# CLI 測試
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "cleanup":
        result = cleanup_old_snapshots()
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif len(sys.argv) > 1 and sys.argv[1] == "list":
        items = list_snapshots()
        print(f"共 {len(items)} 筆快照:")
        for it in items:
            print(f"  {it.get('run_id', '?'):>8} | {it.get('stage', '?'):<8} | "
                  f"{it.get('base_name', '?'):<10} | {it.get('timestamp', '?'):<14} | {it.get('size_kb', 0):>6.1f} KB")
    else:
        # 測試: 創一個假快照
        excel_data = b"FAKE EXCEL DATA"
        json_data = json.dumps({"test": "data"}).encode("utf-8")
        result = save_snapshot(excel_data, json_data, "測試案", findings_count={"不合理": 5, "待確認": 3})
        print(json.dumps(result, ensure_ascii=False, indent=2))
