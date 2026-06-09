# -*- coding: utf-8 -*-
"""審查報告匯出產生器。

對外 API:
    build_export(target, app_data, findings=None, options=None, base_name=None)
        target: "internal" / "vendor" / "json"
        回傳 (bytes_data, file_name, mime)

對象 → 格式:
    internal  → Excel (多分頁: 摘要/不合理/待人工/拓樸備註/單元表/水質/規則對照)
    vendor    → Word (改善建議書, 廠商可編輯回覆)
    json      → JSON (整合 app_data + findings + topology_notes + 統計)
"""
from __future__ import annotations

import io
import json
from datetime import datetime
from typing import Any


# ─────────────────────────────────────────
# 共用: 收集 topology_notes
# ─────────────────────────────────────────

def collect_topology_notes(app_data: dict) -> dict[str, list[str]]:
    """從 units[*]["topology_notes"] 收集所有拓樸備註。"""
    out = {}
    for code, unit in (app_data.get("units") or {}).items():
        notes = unit.get("topology_notes") or []
        if notes:
            out[code] = list(notes)
    return out


def summarize_findings(findings: list[dict]) -> dict[str, int]:
    """統計 findings 各嚴重度數量。"""
    summary = {"不合理": 0, "待人工": 0, "提醒": 0, "錯誤": 0, "其他": 0}
    for f in findings or []:
        sev = f.get("嚴重度", "")
        if sev in summary:
            summary[sev] += 1
        else:
            summary["其他"] += 1
    return summary


# ─────────────────────────────────────────
# 1. Excel (內部覆核)
# ─────────────────────────────────────────

def build_internal_excel(app_data: dict, findings: list[dict] | None = None,
                         options: dict | None = None) -> bytes:
    """產生「內部覆核」Excel。

    分頁:
        1. 審查摘要
        2. 🔴 不合理 findings
        3. 🟡 待人工 findings
        4. ℹ️ 拓樸備註
        5. 全廠單元表
        6. 各單元水質 (進/出 質量平衡)
        7. 觸發規則對照
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    opts = options or {}
    include_topology = opts.get("include_topology", True)
    include_water_quality = opts.get("include_water_quality", True)
    include_rules = opts.get("include_rules", False)

    wb = openpyxl.Workbook()
    findings = findings or []
    topology = collect_topology_notes(app_data)
    summary = summarize_findings(findings)

    # 樣式
    HEAD_FILL = PatternFill("solid", fgColor="305496")
    HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)
    RED_FILL = PatternFill("solid", fgColor="FFE6E6")
    YEL_FILL = PatternFill("solid", fgColor="FFF2CC")
    INFO_FILL = PatternFill("solid", fgColor="DDEBF7")
    WRAP_LEFT = Alignment(wrap_text=True, vertical="top", horizontal="left")

    def _write_header(ws, headers):
        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.font = HEAD_FONT
            c.fill = HEAD_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"

    def _autosize(ws, max_width=60):
        for col in ws.columns:
            try:
                length = max(len(str(c.value or "")) for c in col)
            except Exception:
                length = 12
            letter = col[0].column_letter
            ws.column_dimensions[letter].width = min(max(length + 2, 8), max_width)

    # ─── 分頁 1: 摘要 ───
    ws1 = wb.active
    ws1.title = "1.審查摘要"
    rows = [
        ["項目", "值"],
        ["案件名稱", app_data.get("source_pdf", "")],
        ["業別", opts.get("business_type", "(未指定)")],
        ["匯出時間", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["", ""],
        ["━━━━━━━━━ 審查結果統計 ━━━━━━━━━", ""],
        ["🔴 不合理 (高優先)", summary["不合理"]],
        ["🟡 待人工 (需人工複核)", summary["待人工"]],
        ["💡 提醒 / 其他", summary["提醒"] + summary["其他"]],
        ["ℹ️ 拓樸備註 (水量分流, 非異常)", sum(len(v) for v in topology.values())],
        ["⚠️ 系統錯誤", summary["錯誤"]],
        ["", ""],
        ["━━━━━━━━━ 結構統計 ━━━━━━━━━", ""],
        ["處理單元數", len(app_data.get("units") or {})],
        ["有拓樸備註的單元", len(topology)],
    ]
    for r_idx, r in enumerate(rows, start=1):
        for c_idx, v in enumerate(r, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=v)
            if r_idx == 1:
                cell.font = HEAD_FONT
                cell.fill = HEAD_FILL
    ws1.column_dimensions["A"].width = 40
    ws1.column_dimensions["B"].width = 50

    # ─── 分頁 2: 不合理 findings ───
    ws2 = wb.create_sheet("2.不合理🔴")
    headers = ["嚴重度", "類型", "單元", "標準槽體", "對照項目", "描述", "依據"]
    _write_header(ws2, headers)
    row = 2
    for f in findings:
        if f.get("嚴重度") != "不合理":
            continue
        for col_idx, key in enumerate(headers, start=1):
            cell = ws2.cell(row=row, column=col_idx, value=f.get(key, ""))
            cell.alignment = WRAP_LEFT
            cell.fill = RED_FILL
        row += 1
    _autosize(ws2)

    # ─── 分頁 3: 待人工 findings ───
    ws3 = wb.create_sheet("3.待人工🟡")
    _write_header(ws3, headers)
    row = 2
    for f in findings:
        if f.get("嚴重度") != "待人工":
            continue
        for col_idx, key in enumerate(headers, start=1):
            cell = ws3.cell(row=row, column=col_idx, value=f.get(key, ""))
            cell.alignment = WRAP_LEFT
            cell.fill = YEL_FILL
        row += 1
    _autosize(ws3)

    # ─── 分頁 4: 拓樸備註 ───
    if include_topology and topology:
        ws4 = wb.create_sheet("4.拓樸備註ℹ️")
        _write_header(ws4, ["單元", "備註"])
        row = 2
        for code, notes in sorted(topology.items()):
            for n in notes:
                ws4.cell(row=row, column=1, value=code).fill = INFO_FILL
                cell = ws4.cell(row=row, column=2, value=n)
                cell.alignment = WRAP_LEFT
                cell.fill = INFO_FILL
                row += 1
        ws4.column_dimensions["A"].width = 12
        ws4.column_dimensions["B"].width = 80

    # ─── 分頁 5: 全廠單元表 ───
    ws5 = wb.create_sheet("5.單元表")
    unit_headers = ["代號", "原始名稱", "標準槽體", "進流數", "出流數",
                    "有效容量(m³)", "HRT(hr)", "SOR(m³/m²·d)", "G(1/s)",
                    "拓樸備註數"]
    _write_header(ws5, unit_headers)
    row = 2
    for code, unit in sorted((app_data.get("units") or {}).items()):
        size = unit.get("size") or {}
        dp = unit.get("design_params") or {}
        def _get_dp(*keys):
            for k in keys:
                v = dp.get(k) or {}
                if isinstance(v, dict):
                    raw = v.get("raw") or v.get("min") or v.get("max")
                    if raw:
                        return str(raw)
            return ""
        notes = len(unit.get("topology_notes") or [])
        ws5.cell(row=row, column=1, value=code)
        ws5.cell(row=row, column=2, value=unit.get("name_in_doc", ""))
        ws5.cell(row=row, column=3, value=unit.get("std_tank", ""))
        ws5.cell(row=row, column=4, value=len(unit.get("influent") or {}))
        ws5.cell(row=row, column=5, value=len(unit.get("effluent") or {}))
        ws5.cell(row=row, column=6, value=size.get("有效容量", ""))
        ws5.cell(row=row, column=7, value=_get_dp("水力停留時間", "HRT"))
        ws5.cell(row=row, column=8, value=_get_dp("表面溢流率", "SOR"))
        ws5.cell(row=row, column=9, value=_get_dp("G值", "速度梯度", "G"))
        ws5.cell(row=row, column=10, value=notes)
        if notes:
            for c in ws5[row]:
                c.fill = INFO_FILL
        row += 1
    _autosize(ws5)

    # ─── 分頁 6: 各單元水質 ───
    if include_water_quality:
        ws6 = wb.create_sheet("6.各單元水質")
        wq_headers = ["單元", "方向", "Stream編號", "水質項目", "濃度", "質量"]
        _write_header(ws6, wq_headers)
        row = 2
        for code, unit in sorted((app_data.get("units") or {}).items()):
            for direction, key in [("進", "influent"), ("出", "effluent")]:
                for stream_code, items in (unit.get(key) or {}).items():
                    if not isinstance(items, dict):
                        continue
                    for item_name, val in items.items():
                        if not isinstance(val, dict):
                            continue
                        ws6.cell(row=row, column=1, value=code)
                        ws6.cell(row=row, column=2, value=direction)
                        ws6.cell(row=row, column=3, value=stream_code)
                        ws6.cell(row=row, column=4, value=item_name)
                        ws6.cell(row=row, column=5, value=val.get("濃度"))
                        ws6.cell(row=row, column=6, value=val.get("質量"))
                        row += 1
        _autosize(ws6)

    # ─── 分頁 7: 觸發規則對照 ───
    if include_rules:
        ws7 = wb.create_sheet("7.規則對照")
        _write_header(ws7, ["規則來源", "觸發次數"])
        row = 2
        from collections import Counter
        ctr = Counter()
        for f in findings:
            ctr[f.get("依據", "(無依據)")] += 1
        for rule_src, cnt in ctr.most_common():
            ws7.cell(row=row, column=1, value=rule_src)
            ws7.cell(row=row, column=2, value=cnt)
            row += 1
        ws7.column_dimensions["A"].width = 80
        ws7.column_dimensions["B"].width = 12

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────
# 2. Word (廠商通知)
# ─────────────────────────────────────────

def build_vendor_word(app_data: dict, findings: list[dict] | None = None,
                      options: dict | None = None) -> bytes:
    """產生「廠商審查意見」Word docx。

    結構:
        - 案件基本資料
        - 一、待修正項目 (紅字)
        - 二、待澄清項目 (黃字)
        - 三、拓樸備註 (供參, 不需修正)
        - 廠商回覆欄
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        raise RuntimeError(
            "未安裝 python-docx。請執行: pip install python-docx 後再試。"
        )

    opts = options or {}
    findings = findings or []
    topology = collect_topology_notes(app_data)

    doc = Document()

    # 標題
    title = doc.add_heading("水污染防治措施申請書 - 審查意見書", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 基本資料
    doc.add_paragraph()
    info_p = doc.add_paragraph()
    info_p.add_run("案件名稱: ").bold = True
    info_p.add_run(str(app_data.get("source_pdf", "")) + "\n")
    info_p.add_run("業別: ").bold = True
    info_p.add_run(str(opts.get("business_type", "(未指定)")) + "\n")
    info_p.add_run("審查日期: ").bold = True
    info_p.add_run(datetime.now().strftime("%Y-%m-%d") + "\n")

    # 一、待修正項目
    red = [f for f in findings if f.get("嚴重度") == "不合理"]
    doc.add_heading(f"一、待修正項目 ({len(red)} 項)", level=1)
    if not red:
        doc.add_paragraph("(本案無此類項目)")
    else:
        for i, f in enumerate(red, start=1):
            p = doc.add_paragraph(style="List Number")
            run = p.add_run(f"{f.get('單元', '')} / {f.get('對照項目', '')}")
            run.bold = True
            run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            doc.add_paragraph(f"問題描述: {f.get('描述', '')}")
            doc.add_paragraph(f"依據: {f.get('依據', '')}")
            doc.add_paragraph("廠商回覆: " + " " * 60).italic = True

    # 二、待澄清項目
    yel = [f for f in findings if f.get("嚴重度") == "待人工"]
    doc.add_heading(f"二、待澄清項目 ({len(yel)} 項)", level=1)
    if not yel:
        doc.add_paragraph("(本案無此類項目)")
    else:
        for i, f in enumerate(yel, start=1):
            p = doc.add_paragraph(style="List Number")
            run = p.add_run(f"{f.get('單元', '')} / {f.get('對照項目', '')}")
            run.bold = True
            run.font.color.rgb = RGBColor(0xBF, 0x90, 0x00)
            doc.add_paragraph(f"說明: {f.get('描述', '')}")
            doc.add_paragraph("廠商說明: " + " " * 60).italic = True

    # 三、拓樸備註
    if topology:
        doc.add_heading(f"三、拓樸備註 (供參, 不需修正)", level=1)
        doc.add_paragraph(
            "下列為系統偵測到的「水量分流 / 多源滙流」結構提示, "
            "為審查系統的自動標記, 不屬於需修正項目。"
        )
        for code, notes in sorted(topology.items()):
            doc.add_heading(code, level=2)
            for n in notes:
                doc.add_paragraph(n, style="List Bullet")

    # 簽核
    doc.add_paragraph()
    doc.add_paragraph("─" * 40)
    sig = doc.add_paragraph()
    sig.add_run("審查人: __________________  日期: __________\n").bold = True
    sig.add_run("覆核人: __________________  日期: __________").bold = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────
# 3. JSON (整合 dump, 給 AI 再分析)
# ─────────────────────────────────────────

def build_integrated_json(app_data: dict, findings: list[dict] | None = None,
                          options: dict | None = None) -> bytes:
    """整合 JSON: app_data + findings + topology + 統計。"""
    opts = options or {}
    findings = findings or []
    topology = collect_topology_notes(app_data)
    summary = summarize_findings(findings)

    out = {
        "案件": app_data.get("source_pdf", ""),
        "業別": opts.get("business_type", "(未指定)"),
        "匯出時間": datetime.now().isoformat(),
        "統計": {
            "不合理": summary["不合理"],
            "待人工": summary["待人工"],
            "提醒": summary["提醒"],
            "錯誤": summary["錯誤"],
            "拓樸備註_總則數": sum(len(v) for v in topology.values()),
            "拓樸備註_單元數": len(topology),
            "處理單元數": len(app_data.get("units") or {}),
        },
        "findings": findings,
        "topology_notes": topology,
        "units": app_data.get("units") or {},
        "metadata": {
            "source_pdf": app_data.get("source_pdf"),
            "extracted_at": app_data.get("extracted_at"),
            "total_units": app_data.get("total_units"),
        },
    }
    return json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8")


# ─────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────

def build_export(target: str, app_data: dict, findings: list[dict] | None = None,
                 options: dict | None = None,
                 base_name: str = "report") -> tuple[bytes, str, str]:
    """產生指定對象的匯出檔。

    Args:
        target: "internal" / "vendor" / "json"
        app_data: step2 抽取的 application_*.json
        findings: 智能審查 findings (可為 None)
        options: 額外選項 dict (business_type / include_topology / ...)
        base_name: 檔名前綴

    Returns:
        (bytes_data, file_name, mime_type)
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    if target == "internal":
        data = build_internal_excel(app_data, findings, options)
        fname = f"{base_name}_內部覆核_{ts}.xlsx"
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif target == "vendor":
        data = build_vendor_word(app_data, findings, options)
        fname = f"{base_name}_廠商通知_{ts}.docx"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif target == "json":
        data = build_integrated_json(app_data, findings, options)
        fname = f"{base_name}_整合資料_{ts}.json"
        mime = "application/json"
    else:
        raise ValueError(f"未知 target: {target} (允許: internal / vendor / json)")

    return data, fname, mime


# ─────────────────────────────────────────
# CLI 測試
# ─────────────────────────────────────────

if __name__ == "__main__":
    import os
    import sys

    base = os.path.dirname(os.path.abspath(__file__))
    jsons = sorted([f for f in os.listdir(base) if f.startswith("application_") and f.endswith(".json")])
    if not jsons:
        print("找不到 application_*.json")
        sys.exit(0)
    with open(os.path.join(base, jsons[-1]), encoding="utf-8") as f:
        app = json.load(f)

    # 模擬 findings
    from step3d_principle_check import run_advanced_checks
    findings = run_advanced_checks(app, business_type="電鍍業")

    print(f"案件: {app.get('source_pdf')}")
    print(f"findings: {len(findings)}")

    for tgt in ["internal", "json"]:
        try:
            data, fname, mime = build_export(tgt, app, findings,
                                              options={"business_type": "電鍍業"},
                                              base_name="test")
            print(f"  {tgt}: {len(data)/1024:.1f} KB → {fname} ({mime})")
            with open(os.path.join(base, fname), "wb") as f:
                f.write(data)
        except Exception as e:
            print(f"  {tgt}: ERROR {e}")
