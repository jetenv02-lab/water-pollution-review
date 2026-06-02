# -*- coding: utf-8 -*-
"""水措審查系統 — Streamlit 線上版。

本地測試:
    streamlit run streamlit_app.py

部署到 Streamlit Cloud:
    1) https://streamlit.io/cloud → Sign in with GitHub
    2) New app → 選 jetenv02-lab/water-pollution-review
    3) Main file path: streamlit_app.py
    4) Deploy

線上版特性:
- 規則庫直接從 repo 的 rules_extracted.csv 讀取
- 使用者上傳申請 PDF → 暫存於記憶體/temp → 抽取 + 比對
- 比對結果可下載 Excel/Word
- 不寫回任何持久檔(每次 session 獨立)
"""
import csv
import io
import json
import os
import re
import tempfile
from collections import defaultdict
from datetime import datetime

import streamlit as st
import openpyxl
import pdfplumber
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ───────────────────────────── 頁面設定 ─────────────────────────────

st.set_page_config(
    page_title="水措審查系統",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 自訂樣式
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #2F5496 0%, #1a365d 100%);
        color: white;
        padding: 20px 30px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .main-header h1 { margin: 0; font-size: 28px; }
    .main-header p { margin: 4px 0 0 0; opacity: 0.9; }
    .stat-card {
        background: #f7fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .stat-number {
        font-size: 32px;
        font-weight: 700;
        color: #2F5496;
    }
    .stat-label {
        font-size: 13px;
        color: #718096;
    }
    .finding-unreasonable { background: #fed7d7; padding: 8px; border-radius: 4px; }
    .finding-manual { background: #fefcbf; padding: 8px; border-radius: 4px; }
    .finding-reasonable { background: #c6f6d5; padding: 8px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>💧 水措審查系統</h1>
    <p>自動化「水污染防治措施」申請文件審查 · 比對環工技師查核缺失資料庫</p>
</div>
""", unsafe_allow_html=True)

# ───────────────────────────── 共用資源 ─────────────────────────────

CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules_extracted.csv")

# 槽體關鍵字 → 標準名稱
TANK_KEYWORDS = [
    ("pH調整", "pH調整槽"), ("pH 調整", "pH調整槽"),
    ("廢水調整", "廢水調整池"), ("調勻", "調勻池"),
    ("快混", "快混槽"), ("慢混", "慢混池"),
    ("沉澱", "沉澱池"), ("沉降", "沉澱池"),
    ("中和", "中和池"), ("放流", "放流池"),
    ("曝氣", "曝氣槽"), ("活性污泥", "曝氣槽"),
    ("砂濾", "砂濾塔"), ("過濾", "砂濾塔"),
    ("活性碳吸附", "活性碳吸附塔"), ("活性碳", "活性碳吸附裝置"),
    ("批次反應", "批次反應槽"), ("反應槽", "批次反應槽"),
    ("脫水", "脫水機"), ("壓濾", "脫水機"),
    ("貯留", "貯留槽"), ("暫存", "暫存槽"),
    ("污泥儲", "污泥儲槽"), ("污泥貯", "污泥儲槽"),
    ("濃縮", "濃縮槽"),
]
CODE_PATTERN = re.compile(r"T\d{1,2}-\d{1,3}[a-zA-Z]?|T\d{2}|D\d{1,2}|WM\d{1,2}|WTB\d{1,2}(?:-\d{1,2}[a-z]?)?")
PH_LOWER_PATTERN = re.compile(r"pH\s*下限\s*[<≤]=?\s*(\d+(?:\.\d+)?)")


@st.cache_data(show_spinner=False)
def load_rules():
    """讀取 rules_extracted.csv → 依標準槽體分組。"""
    if not os.path.exists(CSV_PATH):
        return [], {}
    rules = []
    with open(CSV_PATH, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rules.append(row)
    by_tank = defaultdict(list)
    for r in rules:
        tank = (r.get("標準槽體名稱") or "未分類").strip()
        by_tank[tank].append(r)
    return rules, dict(by_tank)


def guess_tank(text):
    for kw, std in TANK_KEYWORDS:
        if kw in text:
            return std
    return None


def extract_application_from_pdf(pdf_file):
    """從 PDF (BytesIO 或檔案路徑) 抽取單元結構化資料。"""
    pages_text = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            try:
                pages_text.append(page.extract_text() or "")
            except Exception:
                pages_text.append("")
    total = len(pages_text)

    units = {}
    for i, text in enumerate(pages_text):
        if not text:
            continue
        norm = text.replace("–", "-")
        for m in CODE_PATTERN.finditer(norm):
            code = m.group(0)
            ctx_start = m.end()
            ctx = norm[ctx_start:ctx_start + 30]
            std = guess_tank(ctx)
            if not std:
                continue
            if code not in units:
                first_word = ctx.split()[0] if ctx.strip() else ""
                units[code] = {
                    "raw_code": code,
                    "std_tank": std,
                    "name_in_doc": (code + " " + first_word),
                    "pages_found": [],
                    "design_params": {},
                }
            if (i + 1) not in units[code]["pages_found"]:
                units[code]["pages_found"].append(i + 1)

    # 抽參數
    param_pattern = re.compile(
        r"(pH|停留時間|有效容量|有效水深|表面溢流率|溢流率|DO|溶氧|MLSS|F/M|食微比|"
        r"污泥迴流率|迴流率|加藥量|濾速|上升流速|體積負荷|有機負荷|處理量)"
        r"\s*[:：]?\s*([\d\.~\-—–]+(?:\s*[a-zA-Z/°%μ]+)?)"
    )
    for code, info in units.items():
        for pg in info["pages_found"]:
            text = pages_text[pg - 1] or ""
            for m in param_pattern.finditer(text):
                key, val = m.group(1), m.group(2).strip()
                if key not in info["design_params"]:
                    info["design_params"][key] = val

    return {"total_pages": total, "units": units}


def parse_ph_range(s):
    if not s:
        return None, None
    s = str(s).replace("～", "~").replace("－", "-").replace("–", "-")
    m = re.search(r"(\d+(?:\.\d+)?)\s*[~\-]\s*(\d+(?:\.\d+)?)", s)
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None


def evaluate_rule(rule, unit_info):
    logic = (rule.get("判定邏輯") or "").strip()
    rule_text = (rule.get("規則") or "").strip()
    ph_value = unit_info.get("design_params", {}).get("pH", "")
    if "pH" in (rule.get("對照項目") or "") or "pH" in logic:
        lo, hi = parse_ph_range(ph_value)
        if lo is None:
            return None, f"無法解析 pH(申請文件: {ph_value!r}) — 待人工"
        m = PH_LOWER_PATTERN.search(logic) or re.search(r"pH\s*下限\s*<\s*(\d+(?:\.\d+)?)", logic)
        if m:
            threshold = float(m.group(1))
            if lo < threshold:
                return True, f"申請 pH 下限 {lo} < 規則門檻 {threshold}"
            else:
                return False, f"申請 pH 下限 {lo} ≥ 規則門檻 {threshold}"
    return None, "規則需人工判定: " + rule_text[:80]


def compare_application(app_data, rules_by_tank):
    findings = []
    for code, info in app_data["units"].items():
        std_tank = info.get("std_tank", "")
        for rule in rules_by_tank.get(std_tank, []):
            triggered, desc = evaluate_rule(rule, info)
            findings.append({
                "申請單元": code,
                "標準槽體": std_tank,
                "缺失ID": rule.get("缺失ID", ""),
                "規則來源": rule.get("來源", ""),
                "檢查類型": rule.get("檢查類型", ""),
                "對照項目": rule.get("對照項目", ""),
                "判定": ("不合理" if triggered else "合理" if triggered is False else "待人工"),
                "描述": desc,
                "規則原文": rule.get("規則", ""),
                "原文缺失": (rule.get("原文缺失") or "")[:200],
            })
    return findings


# ───────────────────────────── 輸出產生器 ─────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SEVERITY_FILL = {
    "不合理": PatternFill("solid", fgColor="FCE4D6"),
    "待人工": PatternFill("solid", fgColor="FFF2CC"),
    "合理": PatternFill("solid", fgColor="E2EFDA"),
}
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_excel(findings, app_filename):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    fields = ["申請單元", "標準槽體", "缺失ID", "規則來源", "檢查類型",
              "對照項目", "判定", "描述", "規則原文", "原文缺失"]
    widths = [12, 14, 10, 24, 12, 14, 10, 36, 36, 36]

    # 總表
    ws = wb.create_sheet("_總表")
    ws.append(fields)
    for r in findings:
        ws.append([r.get(f, "") for f in fields])
        fill = SEVERITY_FILL.get(r.get("判定"))
        if fill:
            for c in range(1, len(fields) + 1):
                ws.cell(row=ws.max_row, column=c).fill = fill
    for c in range(1, len(fields) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(fields)):
        for cell in row:
            cell.alignment = WRAP
            cell.border = BORDER

    # 各單元分頁
    by_unit = defaultdict(list)
    for r in findings:
        by_unit[r["申請單元"]].append(r)
    for unit, rows in sorted(by_unit.items()):
        sname = unit[:31].replace("/", "-")
        ws = wb.create_sheet(sname)
        ws.append(fields)
        for r in rows:
            ws.append([r.get(f, "") for f in fields])
            fill = SEVERITY_FILL.get(r.get("判定"))
            if fill:
                for c in range(1, len(fields) + 1):
                    ws.cell(row=ws.max_row, column=c).fill = fill
        for c in range(1, len(fields) + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
            cell.border = BORDER
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
        ws.freeze_panes = "A2"
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(fields)):
            for cell in row:
                cell.alignment = WRAP
                cell.border = BORDER

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_word(findings, app_filename):
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        return None
    doc = Document()
    doc.add_heading("水措審查意見書", level=0)
    doc.add_paragraph(f"審查文件: {app_filename}")
    doc.add_paragraph(f"審查時間: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"比對總數: {len(findings)}")

    not_ok = [r for r in findings if r["判定"] == "不合理"]
    manual = [r for r in findings if r["判定"] == "待人工"]

    doc.add_heading(f"一、自動判定不合理項目 ({len(not_ok)} 項)", level=1)
    if not_ok:
        for i, r in enumerate(not_ok, 1):
            p = doc.add_paragraph(style="List Number")
            p.add_run(f"【{r['申請單元']} {r['標準槽體']}】").bold = True
            p.add_run(f" 對照項目: {r['對照項目']} | 描述: {r['描述']}\n")
            p.add_run(f"規則原文: {r['規則原文']}\n").italic = True
            p.add_run(f"出處: {r['規則來源']} (缺失ID {r['缺失ID']})").font.size = Pt(9)
    else:
        doc.add_paragraph("(無)")

    doc.add_heading(f"二、待人工複核項目 ({len(manual)} 項)", level=1)
    by_unit = defaultdict(list)
    for r in manual:
        by_unit[r["申請單元"]].append(r)
    for unit, rows in sorted(by_unit.items()):
        doc.add_heading(f"{unit} ({rows[0]['標準槽體']})", level=2)
        for r in rows:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(f"[{r['檢查類型']}/{r['對照項目']}] {r['規則原文']}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ───────────────────────────── 介面 ─────────────────────────────

rules, rules_by_tank = load_rules()

# Sidebar
with st.sidebar:
    st.header("📋 系統狀態")
    st.metric("規則總數", len(rules))
    st.metric("槽體分類數", len(rules_by_tank))
    st.divider()
    st.header("🔍 各槽體規則數")
    for tank, rs in sorted(rules_by_tank.items(), key=lambda x: -len(x[1])):
        st.text(f"  {tank}: {len(rs)} 筆")
    st.divider()
    st.markdown("**專案連結**")
    st.markdown("[GitHub Repo](https://github.com/jetenv02-lab/water-pollution-review)")
    st.markdown("[使用說明 (README)](https://github.com/jetenv02-lab/water-pollution-review#readme)")

# 主畫面分頁
tab1, tab2, tab3, tab4 = st.tabs(["🚀 開始審查", "📊 規則庫瀏覽", "📖 系統說明", "❓ 常見問題"])

# ───── Tab 1: 開始審查 ─────
with tab1:
    st.subheader("上傳申請文件 PDF")
    st.caption("拖拉或點擊選擇要審查的申請 PDF 檔案。系統會自動抽取每個處理單元(T01-03、T01-08 等)的資料，並比對規則庫。")

    uploaded = st.file_uploader(
        "選擇 PDF",
        type=["pdf"],
        help="支援單一 PDF 檔。較大檔案(>50MB)可能需要等待較久。",
    )

    if uploaded is not None:
        st.success(f"已上傳: **{uploaded.name}** ({uploaded.size // 1024} KB)")

        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            run_btn = st.button("🚀 開始審查", type="primary", use_container_width=True)
        with col2:
            preview_only = st.button("👀 僅抽取（不比對）", use_container_width=True)

        if run_btn or preview_only:
            with st.spinner("📖 正在解析 PDF... (大檔案可能需要 1-3 分鐘)"):
                pdf_bytes = uploaded.read()
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                    tf.write(pdf_bytes)
                    tmp_path = tf.name
                try:
                    app_data = extract_application_from_pdf(tmp_path)
                finally:
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass

            st.success(f"✅ 解析完成! 共 {app_data['total_pages']} 頁，偵測到 **{len(app_data['units'])}** 個處理單元")

            if app_data["units"]:
                with st.expander("📋 偵測到的處理單元", expanded=True):
                    unit_rows = []
                    for code, info in sorted(app_data["units"].items()):
                        unit_rows.append({
                            "單元代號": code,
                            "標準槽體": info["std_tank"],
                            "出現頁數": ", ".join(map(str, info["pages_found"][:5])) + ("..." if len(info["pages_found"]) > 5 else ""),
                            "已抽取參數": ", ".join(info["design_params"].keys()) or "(無)",
                        })
                    st.dataframe(unit_rows, use_container_width=True, hide_index=True)

            if run_btn:
                with st.spinner("⚙ 正在比對規則..."):
                    findings = compare_application(app_data, rules_by_tank)

                stats = {"不合理": 0, "合理": 0, "待人工": 0}
                for f in findings:
                    stats[f["判定"]] += 1

                st.subheader("📊 比對結果")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("總比對數", len(findings))
                c2.metric("🔴 不合理", stats["不合理"])
                c3.metric("🟡 待人工", stats["待人工"])
                c4.metric("🟢 合理", stats["合理"])

                # 篩選器
                st.divider()
                col1, col2, col3 = st.columns([1, 1, 2])
                with col1:
                    f_unit = st.selectbox("依單元篩選", ["(全部)"] + sorted({f["申請單元"] for f in findings}))
                with col2:
                    f_severity = st.selectbox("依判定篩選", ["(全部)", "不合理", "待人工", "合理"])
                with col3:
                    f_search = st.text_input("關鍵字搜尋（描述/規則原文）")

                filtered = findings
                if f_unit != "(全部)":
                    filtered = [f for f in filtered if f["申請單元"] == f_unit]
                if f_severity != "(全部)":
                    filtered = [f for f in filtered if f["判定"] == f_severity]
                if f_search:
                    s = f_search.lower()
                    filtered = [f for f in filtered if s in (f.get("描述", "") + f.get("規則原文", "")).lower()]

                st.caption(f"顯示 {len(filtered)} / {len(findings)} 筆")

                # 列表
                for f in filtered[:200]:
                    severity_emoji = {"不合理": "🔴", "待人工": "🟡", "合理": "🟢"}.get(f["判定"], "⚪")
                    with st.expander(f"{severity_emoji} **{f['申請單元']}** ({f['標準槽體']}) — {f['對照項目']} | {f['判定']}"):
                        st.markdown(f"**描述**: {f['描述']}")
                        st.markdown(f"**規則原文**: {f['規則原文']}")
                        st.caption(f"出處: {f['規則來源']} (缺失ID {f['缺失ID']}) | 檢查類型: {f['檢查類型']}")
                        if f.get("原文缺失"):
                            with st.container():
                                st.caption("📄 原始缺失內容:")
                                st.text(f["原文缺失"])

                if len(filtered) > 200:
                    st.warning(f"⚠ 結果太多，僅顯示前 200 筆。請使用上方篩選器縮小範圍，或下載完整 Excel。")

                # 下載
                st.divider()
                st.subheader("📥 下載報告")
                col1, col2 = st.columns(2)
                base_name = os.path.splitext(uploaded.name)[0]
                with col1:
                    excel_buf = build_excel(findings, uploaded.name)
                    st.download_button(
                        "📊 下載 Excel 比對結果",
                        data=excel_buf,
                        file_name=f"{base_name}_比對結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col2:
                    word_buf = build_word(findings, uploaded.name)
                    if word_buf:
                        st.download_button(
                            "📝 下載 Word 審查意見書",
                            data=word_buf,
                            file_name=f"{base_name}_審查意見.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )

# ───── Tab 2: 規則庫瀏覽 ─────
with tab2:
    st.subheader("規則庫內容")
    st.caption(f"目前載入 {len(rules)} 筆規則，分布於 {len(rules_by_tank)} 個槽體分類。")

    if not rules:
        st.warning("找不到 rules_extracted.csv，請確認 repo 內檔案完整。")
    else:
        col1, col2 = st.columns([1, 3])
        with col1:
            selected_tank = st.selectbox("選擇槽體", ["(全部)"] + sorted(rules_by_tank.keys()))
        with col2:
            keyword = st.text_input("關鍵字搜尋")

        display_rules = rules
        if selected_tank != "(全部)":
            display_rules = rules_by_tank.get(selected_tank, [])
        if keyword:
            kw = keyword.lower()
            display_rules = [r for r in display_rules if kw in str(r).lower()]

        st.caption(f"顯示 {len(display_rules)} / {len(rules)} 筆")

        st.dataframe(
            [
                {
                    "ID": r.get("缺失ID"),
                    "槽體": r.get("標準槽體名稱"),
                    "技師": r.get("技師姓名"),
                    "檢查類型": r.get("檢查類型"),
                    "對照項目": r.get("對照項目"),
                    "規則": r.get("規則"),
                    "原文缺失": (r.get("原文缺失") or "")[:100],
                }
                for r in display_rules[:500]
            ],
            use_container_width=True,
            hide_index=True,
        )
        if len(display_rules) > 500:
            st.warning(f"⚠ 結果太多，僅顯示前 500 筆。請使用篩選器縮小範圍。")

        with open(CSV_PATH, "rb") as f:
            st.download_button(
                "📥 下載完整 CSV",
                data=f.read(),
                file_name="rules_extracted.csv",
                mime="text/csv",
            )

# ───── Tab 3: 系統說明 ─────
with tab3:
    st.subheader("水措審查系統 — 使用說明")
    st.markdown("""
### 系統用途

自動化「水污染防治措施」(水措) 申請文件審查工具。根據環工技師過往查核累積的缺失資料庫，
比對申請文件中各處理單元(如 T01-03 批次反應槽、T01-04 中和槽) 的設計參數、水質數據、
機具設施等是否合理。

### 處理單元概念

水措審查中的「處理單元」指廢水處理流程中的單一槽體/設備，常見編號：
- `T01-03` 批次反應槽
- `T01-04` 中和槽
- `T01-08` 沉澱池
- `T01-13` 砂濾塔
- `D01` 放流口
- `WM01`、`WTB01` 水流/水質代號

每個單元有自己的設計參數(有效容量、停留時間、表面溢流率、pH、加藥量、去除率等)。

### 工作流程

1. **上傳申請 PDF** → 系統用 pdfplumber 抽取全文
2. **抽取處理單元** → 用正則 + 關鍵字偵測 T01-XX 代號 + 自動歸類到「標準槽體名稱」
3. **抽取設計參數** → 從各單元頁面萃取 pH、停留時間、DO、MLSS 等
4. **規則比對** → 對每個單元，套用該「標準槽體」分類下的所有規則
5. **輸出結果** → Excel 比對表 + Word 審查意見書

### 去除率計算

```
單元去除率 (%) = (進流濃度 - 出流濃度) / 進流濃度 × 100%
質量去除率 (%) = (M_in - M_out) / M_in × 100%   其中 M = Q × C
```

### 常見不合理狀況

從查核缺失歸納的學理錯誤：
1. 快混槽展現重金屬去除率(沒固液分離)
2. 溶解性物質(導電度、Cl⁻、SO₄²⁻、Na⁺)出現去除率
3. 沉澱池前單元展現 SS 去除率
4. 生物處理對重金屬有高去除率
5. 質量不平衡(進流質量 ≠ 出流 + 污泥帶走)

### 限制

- 本線上版只跑「線上即時審查」，不會把申請 PDF 存到伺服器
- 規則庫內建 53 筆環工技師查核缺失(D001-D047)
- 比對引擎目前只機械化解析 pH 範圍，其他規則標「待人工」
- 完整 235 筆規則的萃取待後續補完
- 申請 PDF 超過 100MB 可能會逾時，建議在本地版執行

### 本地版

如需處理大量檔案或敏感資料不便上傳，可下載原始碼到本機跑：
[GitHub Repo](https://github.com/jetenv02-lab/water-pollution-review)

```bash
git clone https://github.com/jetenv02-lab/water-pollution-review.git
cd water-pollution-review
pip install -r requirements.txt
streamlit run streamlit_app.py
```

或用 Flask Web UI：`python web_app.py`
""")

# ───── Tab 4: 常見問題 ─────
with tab4:
    st.subheader("常見問題")
    with st.expander("Q: 為什麼很多規則都標「待人工」？"):
        st.markdown("""
A: 因為比對引擎目前只機械化解析 **pH 範圍** 這一類規則。其他類型(質量平衡、機具設施、
有效位數等)需要更複雜的解析或人工判讀，所以暫時標「待人工」讓你人工檢視。
未來會逐步擴充機械化解析的範圍。
""")
    with st.expander("Q: 上傳的 PDF 會被保存嗎？隱私安全嗎？"):
        st.markdown("""
A: **不會保存**。上傳的 PDF 在處理完後立即從暫存區刪除。每次重新整理頁面就是全新 session。
不過 Streamlit Cloud 是公開部署，建議敏感檔案不要上傳，用本地版即可。
""")
    with st.expander("Q: 為什麼某些單元沒被偵測到？"):
        st.markdown("""
A: 系統用啟發式偵測 `T01-XX`、`D01`、`WM01` 之類代號。若 PDF 內有特殊命名(如全形數字、
中文夾代號)可能漏抓。可以告訴開發者新增關鍵字。
""")
    with st.expander("Q: 我可以加入新規則嗎？"):
        st.markdown("""
A: 可以！流程：
1. 編輯 `rules_extracted.csv`(用 Excel 或文字編輯器)
2. 加入新的一列，按欄位填寫
3. 推送到 GitHub，Streamlit Cloud 會自動重新部署
""")
    with st.expander("Q: 規則庫的「標準槽體名稱」清單是什麼？"):
        st.markdown("""
A: 目前支援的標準槽體名稱：
`pH調整槽`、`廢水調整池`、`快混槽`、`慢混池`、`沉澱池`、`中和池`、`放流池`、
`曝氣槽`、`砂濾塔`、`活性碳吸附塔`、`活性碳吸附裝置`、`批次反應槽`、`脫水機`、
`暫存槽`、`貯留槽`、`污泥儲槽`、`濃縮槽`、`調勻池`

不特定槽體的填 `(文件類)` 或 `(現場設備類)`。
""")

st.divider()
st.caption("水措審查系統 v1.0 · Made with 💧 by jetenv02 · [GitHub](https://github.com/jetenv02-lab/water-pollution-review)")
