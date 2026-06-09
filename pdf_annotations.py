# -*- coding: utf-8 -*-
"""抽 PDF 上的「技師審查註解」(annotations / sticky notes / highlight comments)。

申請文件 PDF 上, 審查技師會用 Adobe Acrobat / Foxit / 其他 PDF 工具加上
黃色註解、便利貼、高亮等標註, 寫上自己的疑問或修正建議。例如:
    - 「確認?」
    - 「水量?129?」
    - 「13.78?分母應該是每日產出污泥4.365 ?373CMD是水量?」
    - 「污泥單元請重新釐清, 每天有4.3CMD污泥進入脫水機 ...」

這些註解**不是 PDF 內文**, 是 PDF Annotation 物件 (PDF spec 標準), 一般
text extractor (pdfplumber.extract_text) 抓不到, 需要用 pypdf 的
page.annotations 介面。

注意:
    pdfplumber 的 page.annots 在某些 PDF 會遞迴爆炸 (resolve_all 迴圈),
    所以這裡用 pypdf 取代。
"""
import os


def extract_reviewer_notes(pdf_path):
    """從 PDF 抽所有技師審查註解。

    Returns:
        list of {
            "page": int,         # 頁碼 (1-based)
            "subtype": str,      # PDF annotation 類型 (/Text, /Highlight, /FreeText...)
            "author": str,       # 註解作者 (PDF 的 /T 欄)
            "contents": str,     # 註解內容
            "x": float,          # 註解 X 座標 (PDF 內部單位)
            "y": float,          # 註解 Y 座標
        }
    """
    notes = []
    if not os.path.exists(pdf_path):
        return notes

    try:
        from pypdf import PdfReader
    except ImportError:
        return notes

    try:
        reader = PdfReader(pdf_path)
    except Exception:
        return notes

    for pn, page in enumerate(reader.pages, start=1):
        try:
            annots = page.annotations
        except Exception:
            continue
        if not annots:
            continue
        for a in annots:
            try:
                obj = a.get_object()
            except Exception:
                continue
            try:
                subtype = str(obj.get("/Subtype") or "")
                contents = str(obj.get("/Contents") or "").strip()
                author = str(obj.get("/T") or "").strip()

                # 沒內容就跳過 (像是純粹的高亮框, 不是文字註解)
                if not contents:
                    continue

                # 取座標 (Rect = [x1, y1, x2, y2])
                x = y = None
                try:
                    rect = obj.get("/Rect")
                    if rect and len(rect) >= 2:
                        x = float(rect[0])
                        y = float(rect[1])
                except Exception:
                    pass

                notes.append({
                    "page": pn,
                    "subtype": subtype,
                    "author": author,
                    "contents": contents,
                    "x": x,
                    "y": y,
                })
            except Exception:
                continue
    return notes


def group_notes_by_page(notes):
    """依頁碼分組。"""
    by_page = {}
    for n in notes:
        by_page.setdefault(n["page"], []).append(n)
    return by_page


def find_notes_near_unit(notes, unit_pages):
    """找跟某單元相關的註解 (依頁碼比對)。

    Args:
        notes: extract_reviewer_notes 的結果
        unit_pages: 該單元出現的頁碼列表 (從 step2 的 pages_found)

    Returns:
        list of notes that are in/near unit_pages
    """
    if not unit_pages:
        return []
    page_set = set(unit_pages)
    # 也接受 ±1 頁的鄰近註解 (有時技師在頁尾批註下一頁的東西)
    nearby = set()
    for p in unit_pages:
        nearby.update([p - 1, p, p + 1])

    direct = [n for n in notes if n["page"] in page_set]
    near = [n for n in notes if n["page"] in nearby and n["page"] not in page_set]
    return {"direct": direct, "nearby": near}


if __name__ == "__main__":
    import sys, io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    pdf = sys.argv[1] if len(sys.argv) > 1 else r"參考/需審查之文件/01 線上資料(邑昇)(1150529).pdf"
    notes = extract_reviewer_notes(pdf)
    print(f"抽到 {len(notes)} 筆技師註解\n")
    for n in notes:
        author_tag = f" ({n['author']})" if n['author'] else ""
        print(f"📌 頁 {n['page']}{author_tag} [{n['subtype']}]")
        print(f"   {n['contents'][:200]}")
        print()
