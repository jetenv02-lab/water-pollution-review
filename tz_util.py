# -*- coding: utf-8 -*-
"""時區工具 — 統一用台北時區 (UTC+8)。

問題:
    Streamlit Cloud / GitHub Actions 等雲端跑的伺服器預設時區通常是 UTC,
    所以 `datetime.now()` / `date.today()` 會回 UTC 時間, 落差 8 小時。
    這會讓「審查時間」、「上傳時間」、「規則庫版本」等所有顯示給使用者的時間
    都晚了 8 小時 (看起來像「明明剛剛跑完, 顯示卻是 8 小時前」)。

用法:
    from tz_util import now_tpe, today_tpe, fmt_tpe

    now_tpe()                       # datetime(2026, 6, 4, 15, 30, ..., tzinfo=TPE)
    today_tpe()                     # date(2026, 6, 4)
    fmt_tpe()                       # "2026-06-04 15:30:00"
    fmt_tpe("%Y%m%d_%H%M%S")        # "20260604_153000"
    fmt_tpe_from_ts(mtime)          # 把 unix timestamp 轉成台北時區字串
"""
from datetime import datetime, date

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    TPE_TZ = ZoneInfo("Asia/Taipei")
except Exception:
    # 後備: Python <3.9 或缺 tzdata
    from datetime import timezone, timedelta
    TPE_TZ = timezone(timedelta(hours=8))


def now_tpe():
    """目前時間 (台北時區, aware datetime)。"""
    return datetime.now(TPE_TZ)


def today_tpe():
    """今天的日期 (台北時區)。"""
    return now_tpe().date()


def fmt_tpe(fmt="%Y-%m-%d %H:%M:%S"):
    """目前時間格式化字串 (預設 'YYYY-MM-DD HH:MM:SS' 台北時區)。"""
    return now_tpe().strftime(fmt)


def fmt_tpe_from_ts(ts, fmt="%Y-%m-%d %H:%M"):
    """把 unix timestamp 轉成台北時區字串。"""
    return datetime.fromtimestamp(ts, TPE_TZ).strftime(fmt)


def iso_tpe():
    """ISO 8601 字串 (帶 +08:00 時區資訊)。"""
    return now_tpe().isoformat()
