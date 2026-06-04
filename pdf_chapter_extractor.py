# -*- coding: utf-8 -*-
"""申請文件 PDF 章節擷取 — 只切「參、水污染防治措施資料」段。

申請文件通常有完整章節:
    壹、申請文件目錄
    貳、廢(污)水基本資料
    參、水污染防治措施資料      ← 我們要這段
    肆、廢(污)水處理設施操作條件
    伍、附件

抽出來那段給後續 OCR / Gemini 用, 加速 + 省 token。
"""
import io
import os
import re

# 章節標題正則 (參章 + 變體)
CHAPTER_III_PATTERNS = [
    re.compile(r"^\s*[參參]\s*[、,.]\s*水污染防治措施"),
    re.compile(r"^\s*第\s*[參参三3]\s*章\s*水污染防治措施"),
    re.compile(r"^\s*III[\s\.、,]+水污染防治措施"),
    re.compile(r"^\s*3[\s\.、,]+水污染防治措施"),
]

# 下一章 (肆) 的正則 — 用來標記參章結束
CHAPTER_IV_PATTERNS = [
    re.compile(r"^\s*[肆肆]\s*[、,.]"),
    re.compile(r"^\s*第\s*[肆四4]\s*章"),
    re.compile(r"^\s*IV[\s\.、,]"),
    re.compile(r"^\s*4[\s\.、,]"),
]


def _is_chapter_iii(line):
    """判斷一行是不是「參、水污染防治措施」開頭。"""
    line_clean = line.strip()
    return any(p.match(line_clean) for p in CHAPTER_III_PATTERNS)


def _is_chapter_iv(line):
    """判斷一行是不是「肆」或「第四章」開頭 (參章結束)。"""
    line_clean = line.strip()
    if not line_clean:
        return False
    # 避免誤判 (例如「肆陸」開頭的詞), 看後面接什麼
    return any(p.match(line_clean) for p in CHAPTER_IV_PATTERNS)


def find_chapter_iii_pages(pdf_bytes):
    """掃整份 PDF, 找參章的 起始頁 / 結束頁 (1-based)。

    Returns:
        {
            "ok": True,
            "found": bool,
            "start_page": int or None,
            "end_page": int or None,
            "total_pages": int,
            "title_line": str (找到的標題原文)
        }
    """
    try:
        import pdfplumber
    except ImportError as e:
        return {"ok": False, "error": f"缺少 pdfplumber: {e}"}

    try:
        start_page = None
        end_page = None
        title_line = ""
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    if start_page is None and _is_chapter_iii(line):
                        start_page = i
                        title_line = line.strip()[:50]
                        break
                    if start_page is not None and i > start_page and _is_chapter_iv(line):
                        end_page = i - 1
                        break
                if end_page is not None:
                    break

            if start_page is not None and end_page is None:
                end_page = total  # 沒找到肆章 → 一路到底

        return {
            "ok": True,
            "found": start_page is not None,
            "start_page": start_page,
            "end_page": end_page,
            "total_pages": total,
            "title_line": title_line,
        }
    except Exception as e:
        return {"ok": False, "error": f"PDF 解析失敗: {e}"}


def extract_chapter_iii_text(pdf_bytes, location=None):
    """抽參章的純文字。

    Args:
        pdf_bytes: PDF 原始 bytes
        location: 可選, 已知的 {"start_page": N, "end_page": M};
                  沒給就自動找

    Returns:
        {
            "ok": True,
            "text": str,
            "start_page": int,
            "end_page": int,
            "total_pages": int,
            "chars": int,
            "found": bool,
            "fallback_used": bool   (找不到參章時, 是否用全文當 fallback)
        }
    """
    if location is None:
        location = find_chapter_iii_pages(pdf_bytes)
        if not location.get("ok"):
            return location

    if not location.get("found"):
        # 找不到 → fallback 用全文
        return _extract_all_text(pdf_bytes, fallback=True)

    start = location["start_page"]
    end = location["end_page"]

    try:
        import pdfplumber
    except ImportError as e:
        return {"ok": False, "error": f"缺少 pdfplumber: {e}"}

    try:
        parts = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                if i < start or i > end:
                    continue
                text = page.extract_text() or ""
                parts.append(f"\n---PAGE {i}---\n{text}")
                # 表格
                try:
                    tables = page.extract_tables()
                    for t_idx, table in enumerate(tables):
                        parts.append(f"\n[表 {i}-{t_idx + 1}]\n")
                        for row in table:
                            parts.append("\t".join(str(c) if c is not None else "" for c in row))
                            parts.append("\n")
                except Exception:
                    pass
        text = "".join(parts)
        return {
            "ok": True,
            "text": text,
            "start_page": start,
            "end_page": end,
            "total_pages": total,
            "chars": len(text),
            "found": True,
            "fallback_used": False,
        }
    except Exception as e:
        return {"ok": False, "error": f"PDF 解析失敗: {e}"}


def _extract_all_text(pdf_bytes, fallback=False):
    """全文抽取 (沒找到參章時用)。"""
    try:
        import pdfplumber
    except ImportError as e:
        return {"ok": False, "error": f"缺少 pdfplumber: {e}"}

    try:
        parts = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                parts.append(f"\n---PAGE {i}---\n{text}")
        text = "".join(parts)
        return {
            "ok": True,
            "text": text,
            "start_page": 1,
            "end_page": total,
            "total_pages": total,
            "chars": len(text),
            "found": False,
            "fallback_used": fallback,
        }
    except Exception as e:
        return {"ok": False, "error": f"PDF 解析失敗: {e}"}


# ──────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print("用法: python pdf_chapter_extractor.py <PDF 路徑>")
        sys.exit(0)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"找不到 {pdf_path}")
        sys.exit(1)

    with open(pdf_path, "rb") as f:
        data = f.read()

    print(f"=== 找參章 ({pdf_path}) ===")
    loc = find_chapter_iii_pages(data)
    print(loc)

    if loc.get("ok") and loc.get("found"):
        print()
        print(f"=== 抽參章內容 (頁 {loc['start_page']}~{loc['end_page']}) ===")
        result = extract_chapter_iii_text(data, loc)
        if result.get("ok"):
            print(f"字數: {result['chars']}")
            print(f"前 500 字:")
            print(result["text"][:500])
        else:
            print(f"失敗: {result.get('error')}")
