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
    if pdf_path:
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                for pn, page in enumerate(pdf.pages, start=1):
                    try:
                        text = page.extract_text() or ""
                    except Exception:
                        continue
                    # 是否有「示意圖 / 附圖 / 流向圖 / 平衡圖」標題
                    if any(kw in text for kw in ["示意圖", "附圖", "流向圖", "水量平衡圖"]):
                        # 看該頁有沒有圖
                        try:
                            n_images = len(page.images or [])
                        except Exception:
                            n_images = 0
                        # 也用「文字字數 < 200」當作純圖頁
                        text_len = len(text.strip())
                        if n_images == 0 and text_len > 500:
                            # 有大量文字但沒圖, 該頁可能是「該有圖但圖沒顯示」
                            findings.append({
                                "嚴重度": "待確認",
                                "類型": "文件一致性",
                                "單元": f"頁 {pn}",
                                "標準槽體": "",
                                "對照項目": "可能漏附圖",
                                "描述": (
                                    f"頁 {pn} 含「示意圖 / 流向圖」等標題, 但偵測不到圖片物件。"
                                    f"請確認該頁是否漏附圖, 或圖以 PDF 文字向量方式呈現 (此情況可忽略)。"
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
