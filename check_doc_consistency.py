# -*- coding: utf-8 -*-
"""PDF 文件一致性檢查:
    1. 跨頁數值一致 — 同一單元的 Q / 有效容量 在多頁出現時應一致
    2. 圖面缺失偵測 — 「附件」「附圖」「圖示」字眼出現但該頁無圖
"""
import re


def _to_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        m = re.search(r"\d+(?:\.\d+)?", str(v))
        if m:
            try:
                return float(m.group())
            except ValueError:
                return None
        return None


# ─────────────────────────────────────────────
# #2 跨頁數值一致 (Q / 有效容量 / 有效水深)
# ─────────────────────────────────────────────


def check_cross_page_consistency(app_data):
    """同單元的 Q / 容積在「設施資料表」「水質表」「示意圖」三處出現,
    用 stream_q 反推的 Q 跟 size.有效容量 比對, 看「計算式裡的隱含 Q」是否一致。
    """
    findings = []
    units = app_data.get("units", {})

    for code, unit in units.items():
        # 從 design_params 找含「× CMD」或「÷ CMD」的計算式中隱含的 Q
        embedded_q_values = []
        for params in [unit.get("design_params") or {}, unit.get("measure_params") or {}]:
            for pname, pval in params.items():
                if not isinstance(pval, dict):
                    continue
                raw = str(pval.get("raw") or "")
                # 找 "÷ XX CMD" 或 "× XX CMD" 中的 XX
                for m in re.finditer(r"[÷/×*]\s*(\d+(?:\.\d+)?)\s*CMD", raw):
                    try:
                        q = float(m.group(1))
                        embedded_q_values.append((pname, q, raw[:60]))
                    except ValueError:
                        pass

        if not embedded_q_values:
            continue

        # 比對「stream_q 反推」的 Q
        stream_q = unit.get("stream_q") or {}
        reverse_qs = []
        for s_code, qres in stream_q.items():
            if isinstance(qres, dict) and qres.get("ok"):
                q = qres.get("q_cmd")
                if q and q > 0:
                    reverse_qs.append((s_code, q))

        if not reverse_qs:
            continue

        # 用最大的反推 Q 跟設計參數裡寫的 Q 比對
        max_reverse_q = max(q for _, q in reverse_qs)

        for pname, q_from_calc, raw_snippet in embedded_q_values:
            if max_reverse_q <= 0:
                continue
            diff_pct = abs(q_from_calc - max_reverse_q) / max(q_from_calc, max_reverse_q) * 100
            if diff_pct > 10:
                severity = "不合理" if diff_pct > 30 else "待確認"
                findings.append({
                    "嚴重度": severity,
                    "類型": "文件一致性",
                    "單元": code,
                    "標準槽體": unit.get("std_tank", ""),
                    "對照項目": f"跨資料 Q 不一致: {pname}",
                    "描述": (
                        f"設計參數「{pname}」計算式裡寫 Q = {q_from_calc} CMD, "
                        f"但從水質表反推 Q = {max_reverse_q:.2f} CMD, 差 {diff_pct:.0f}%。"
                        f"申請文件不同處的水量應一致, 請核對。"
                    ),
                    "依據": "PDF 設計參數 Q vs 反推 Q",
                })

    return findings


# ─────────────────────────────────────────────
# #5 圖面缺失偵測
# ─────────────────────────────────────────────


def check_diagram_presence(app_data, pdf_path=None):
    """檢查 PDF 中「附件」「附圖」「圖示」等出現的頁是否有圖。

    使用方式:
        從 reviewer_notes 找含「圖面無法顯示」「圖缺失」「附件無」等的人工註解,
        直接列為提示 (因為自動偵測「該頁有沒有圖」需要 pdfplumber image API,
        實作較複雜, 先用人工註解優先)。

    若有 pdf_path, 進階版可掃所有頁面, 找含「圖」字眼但 page.images 為空的頁。
    """
    findings = []
    notes = app_data.get("reviewer_notes") or []

    # 從技師註解找跟圖面相關的關鍵字
    diagram_keywords = ["圖面無法顯示", "圖面缺失", "圖示缺", "附件無",
                        "圖看不到", "圖呢", "圖在哪", "無圖", "缺圖"]

    for n in notes:
        contents = n.get("contents", "")
        if any(kw in contents for kw in diagram_keywords):
            findings.append({
                "嚴重度": "待確認",
                "類型": "文件一致性",
                "單元": f"頁 {n.get('page', '?')}",
                "標準槽體": "",
                "對照項目": "圖面 / 附件缺失",
                "描述": (
                    f"PDF 上技師註解: 「{contents[:80]}」"
                    f" (頁 {n.get('page')})。請確認該頁是否有圖未顯示, "
                    f"或附件是否漏附。"
                ),
                "依據": "PDF 技師審查註解 + 圖面完整性檢查",
            })

    # 進階版 (有 pdf_path 時): 掃 PDF 找含「圖」字眼但無圖的頁
    #
    # 排除清單: 這類頁本來就會提到「示意圖」等字眼但不需要真的長一張圖
    # (例如檢核表列出「應檢附示意圖」是描述性文字, 不是「該頁應該是張圖」)
    #
    # 誤判修正 (2026-06-30):
    # - 排除申請目錄/變更項目表/檢核表/簽證表 等純文字章節
    # - 頁碼改雙標示 (物理頁 N / 內部頁次 M) 讓技師對得上 PDF reader
    # - 嚴重度降為「提醒」(這是自動偵測, 誤判率高, 不宜「待確認」)
    EXCLUDE_TITLES = [
        "變更項目表",           # p5 型: 勾選清單有「水質水量平衡示意圖」字眼
        "水污染防治措施計畫及許可申請文件檢核表",  # 檢核表列出應附的文件
        "水污染防治措施資料技師簽證表",  # 簽證表提到附件目錄
        "簽證工作底稿",         # 簽證附件目錄
        "涉及變更之相關附件",   # 檢核表項目行
        "應檢附之相關附件",     # 檢核表項目行
        "檢附之相關附件",       # 檢核表項目行
        "檢核表",              # 通用檢核表
        "附件清單",            # 附件目錄
        "頁至",                # 「附 1 頁至 附 76 頁」這種描述性內容
        # 章節總稱 (示意圖字眼出現在章節大標, 但該章實際內容是表格說明)
        "處理單元之進出水質資料",  # p13 型: 章節總稱含「示意圖」但內容是水質表
        "進出處理單元之水質資料",  # 同上變體
        "進流水流編號",         # 有 WTB/WTA 編號 = 這頁是水質表, 不是圖
        "水質項目",             # 水質表表頭
    ]
    if pdf_path:
        try:
            import pdfplumber, re
            with pdfplumber.open(pdf_path) as pdf:
                for pn, page in enumerate(pdf.pages, start=1):
                    try:
                        text = page.extract_text() or ""
                    except Exception:
                        continue
                    # 是否有「示意圖 / 附圖 / 流向圖 / 平衡圖」標題
                    if not any(kw in text for kw in ["示意圖", "附圖", "流向圖", "水量平衡圖"]):
                        continue
                    # 排除純文字章節 (檢核表 / 簽證表 / 申請目錄 等)
                    if any(bad in text for bad in EXCLUDE_TITLES):
                        continue
                    # 看該頁有沒有圖
                    try:
                        n_images = len(page.images or [])
                    except Exception:
                        n_images = 0
                    # 也用「文字字數 < 200」當作純圖頁
                    text_len = len(text.strip())
                    if n_images != 0 or text_len <= 500:
                        continue

                    # 抓「頁次: N/M」讓 finding 描述雙標示
                    inner_page = ""
                    m_page = re.search(r"頁次[:：]\s*(\d+)\s*[／/]\s*(\d+)", text)
                    if m_page:
                        inner_page = f" (內部頁次 {m_page.group(1)}/{m_page.group(2)})"
                    page_label = f"頁 {pn}{inner_page}"

                    findings.append({
                        "嚴重度": "提醒",  # 降級: 自動偵測誤判率高
                        "類型": "文件一致性",
                        "單元": page_label,
                        "標準槽體": "",
                        "對照項目": "可能漏附圖",
                        "描述": (
                            f"物理頁 p{pn}{inner_page} 含「示意圖 / 流向圖」等標題, "
                            f"但偵測不到圖片物件。請確認該頁是否漏附圖, "
                            f"或圖以 PDF 文字向量方式呈現 (此情況可忽略)。"
                        ),
                        "依據": "PDF 圖頁完整性自動偵測",
                    })
        except Exception:
            pass

    return findings


# ─────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────


def run_doc_consistency_checks(app_data, pdf_path=None):
    """跑所有文件一致性檢查。"""
    findings = []
    try:
        findings.extend(check_cross_page_consistency(app_data))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_cross_page_consistency",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    try:
        findings.extend(check_diagram_presence(app_data, pdf_path))
    except Exception as e:
        findings.append({
            "嚴重度": "錯誤",
            "類型": "系統",
            "單元": "(全廠)",
            "標準槽體": "",
            "對照項目": "check_diagram_presence",
            "描述": f"檢查器錯誤: {e}",
            "依據": "(內部)",
        })

    return findings
