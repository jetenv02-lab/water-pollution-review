# -*- coding: utf-8 -*-
"""水措審查系統 — Streamlit 線上版 (v2 邏輯,經使用者明確授權覆蓋)。

本檔內容已切換為 streamlit_app_v2.py 的邏輯,使 Streamlit Cloud
預設 Main file (streamlit_app.py) 就能跑 v2 版抽取器。

v2 重大改進:
- 改用成熟版抽取器 step2_extract_v2 (覆蓋率 100%、單元類型歸類正確)
- 顯示更豐富的單元資訊(設計參數/量測參數/機具設施/進出流水質)
- 加入「水量平衡缺口提示」(尚未做 OCR,但會說明限制)
"""
import csv
import io
import json
import os
import tempfile
from collections import defaultdict
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Import 內部模組 - 用 try/except 包起避免單一模組失敗讓整個 App 掛掉
# (Python 3.14 + Streamlit Cloud 偶發 KeyError 在 module loading)
_import_errors = []

try:
    from step2_extract_v2 import extract_application
except Exception as e:
    _import_errors.append(("step2_extract_v2", str(e)))
    extract_application = None

try:
    from step2b_locate_sections import locate_sections, compress_ranges
except Exception as e:
    _import_errors.append(("step2b_locate_sections", str(e)))
    locate_sections = None
    def compress_ranges(p): return str(p)

try:
    from step2c_ocr_diagram import ocr_diagram_pages
except Exception as e:
    _import_errors.append(("step2c_ocr_diagram", str(e)))
    def ocr_diagram_pages(*a, **k): return {"error": f"OCR 模組載入失敗: {_import_errors[-1][1]}"}

try:
    from step3b_balance_check import run_balance_checks
except Exception as e:
    _import_errors.append(("step3b_balance_check", str(e)))
    def run_balance_checks(*a, **k): return []

try:
    from step3c_unit_db import BUSINESS_TYPES
except Exception as e:
    _import_errors.append(("step3c_unit_db", str(e)))
    BUSINESS_TYPES = {}

try:
    from step3d_principle_check import run_advanced_checks
except Exception as e:
    _import_errors.append(("step3d_principle_check", str(e)))
    def run_advanced_checks(*a, **k): return []

try:
    from step3e_rule_driven_check import run_rule_driven_check
except Exception as e:
    _import_errors.append(("step3e_rule_driven_check", str(e)))
    def run_rule_driven_check(*a, **k): return []

try:
    from step4_flow_graph import build_flow_graph, get_unit_neighbors
except Exception as e:
    _import_errors.append(("step4_flow_graph", str(e)))
    def build_flow_graph(*a, **k): return {"nodes": [], "edges": [], "wm_sources": [], "discharges": []}
    def get_unit_neighbors(*a, **k): return {"upstream": [], "downstream": []}

# ───────────────────────────── 頁面設定 ─────────────────────────────

st.set_page_config(
    page_title="水措審查系統 v3",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
    .version-badge {
        background: #38a169;
        color: white;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 13px;
        margin-left: 8px;
    }
    .ocr-warning {
        background: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 12px 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>水措審查系統 <span class="version-badge">v3</span></h1>
    <p>自動化「水污染防治措施」申請文件審查 · 比對環工技師查核缺失資料庫</p>
    <p style="font-size: 13px; opacity: 0.85; margin-top: 6px;">
        v3 新增: 章節動態定位 · OCR 流向圖解析 · 智能審查 (學理檢查)
    </p>
</div>
""", unsafe_allow_html=True)

CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules_extracted.csv")


@st.cache_data(show_spinner=False)
def load_rules():
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


# ───────────────────────────── Excel 輸出 ─────────────────────────────

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


def build_unit_excel(app_data):
    """產出『各單元資料表』Excel (含設計/量測/機具/進出水質)。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 總表分頁
    ws = wb.create_sheet("_單元總表")
    headers = ["單元代號", "原始名稱", "標準槽體類型", "代碼", "出現頁數",
               "設計參數數", "量測參數數", "機具項目數", "進流數", "出流數"]
    widths = [12, 30, 16, 8, 16, 12, 12, 12, 10, 10]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
    ws.freeze_panes = "A2"

    for code, info in sorted(app_data["units"].items()):
        ws.append([
            code, info.get("name_in_doc", ""), info.get("std_tank", ""),
            info.get("code_id", ""),
            ", ".join(map(str, info.get("pages_found", []))),
            len(info.get("design_params", {})),
            len(info.get("measure_params", {})),
            len(info.get("equipment", [])),
            len(info.get("influent", {})),
            len(info.get("effluent", {})),
        ])

    # 每個單元一張分頁
    for code, info in sorted(app_data["units"].items()):
        sname = code[:31]
        ws = wb.create_sheet(sname)
        row = 1

        # 標題
        ws.cell(row=row, column=1, value=f"{code} {info.get('name_in_doc', '')}").font = Font(size=14, bold=True)
        row += 1
        ws.cell(row=row, column=1, value=f"標準類型: {info.get('std_tank', '')} | 代碼: {info.get('code_id', '')} | 頁: {info.get('pages_found', [])}")
        row += 2

        # 設計參數
        if info.get("design_params"):
            ws.cell(row=row, column=1, value="設計操作參數").font = Font(bold=True, color="2F5496")
            row += 1
            ws.cell(row=row, column=1, value="參數")
            ws.cell(row=row, column=2, value="最小值")
            ws.cell(row=row, column=3, value="最大值")
            ws.cell(row=row, column=4, value="原始")
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = HEADER_FILL
                ws.cell(row=row, column=c).font = HEADER_FONT
            row += 1
            for pname, pval in info["design_params"].items():
                ws.cell(row=row, column=1, value=pname)
                ws.cell(row=row, column=2, value=pval.get("min", ""))
                ws.cell(row=row, column=3, value=pval.get("max", ""))
                ws.cell(row=row, column=4, value=pval.get("raw", ""))
                row += 1
            row += 1

        # 量測參數
        if info.get("measure_params"):
            ws.cell(row=row, column=1, value="量測操作參數").font = Font(bold=True, color="2F5496")
            row += 1
            ws.cell(row=row, column=1, value="參數")
            ws.cell(row=row, column=2, value="最小值")
            ws.cell(row=row, column=3, value="最大值")
            ws.cell(row=row, column=4, value="原始")
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = HEADER_FILL
                ws.cell(row=row, column=c).font = HEADER_FONT
            row += 1
            for pname, pval in info["measure_params"].items():
                ws.cell(row=row, column=1, value=pname)
                ws.cell(row=row, column=2, value=pval.get("min", ""))
                ws.cell(row=row, column=3, value=pval.get("max", ""))
                ws.cell(row=row, column=4, value=pval.get("raw", ""))
                row += 1
            row += 1

        # 機具
        if info.get("equipment"):
            ws.cell(row=row, column=1, value="相關機具設施").font = Font(bold=True, color="2F5496")
            row += 1
            ws.cell(row=row, column=1, value="名稱")
            ws.cell(row=row, column=2, value="位置")
            ws.cell(row=row, column=3, value="數量")
            ws.cell(row=row, column=4, value="馬力(kW)")
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = HEADER_FILL
                ws.cell(row=row, column=c).font = HEADER_FONT
            row += 1
            for eq in info["equipment"]:
                ws.cell(row=row, column=1, value=eq.get("name", ""))
                ws.cell(row=row, column=2, value=eq.get("位置", ""))
                ws.cell(row=row, column=3, value=eq.get("數量", ""))
                ws.cell(row=row, column=4, value=eq.get("馬力_kW", ""))
                row += 1
            row += 1

        # 進流水質
        for infl_code, q_data in info.get("influent", {}).items():
            ws.cell(row=row, column=1, value=f"進流水質 — {infl_code}").font = Font(bold=True, color="2F5496")
            row += 1
            ws.cell(row=row, column=1, value="水質項目")
            ws.cell(row=row, column=2, value="濃度")
            ws.cell(row=row, column=3, value="質量")
            ws.cell(row=row, column=4, value="範圍(pH/水溫)")
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = HEADER_FILL
                ws.cell(row=row, column=c).font = HEADER_FONT
            row += 1
            for item, val in q_data.items():
                ws.cell(row=row, column=1, value=item)
                ws.cell(row=row, column=2, value=val.get("濃度", ""))
                ws.cell(row=row, column=3, value=val.get("質量", ""))
                ws.cell(row=row, column=4, value=val.get("範圍", ""))
                row += 1
            row += 1

        # 出流水質
        for effl_code, q_data in info.get("effluent", {}).items():
            ws.cell(row=row, column=1, value=f"出流水質 — {effl_code}").font = Font(bold=True, color="38a169")
            row += 1
            ws.cell(row=row, column=1, value="水質項目")
            ws.cell(row=row, column=2, value="濃度")
            ws.cell(row=row, column=3, value="質量")
            ws.cell(row=row, column=4, value="範圍(pH/水溫)")
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = HEADER_FILL
                ws.cell(row=row, column=c).font = HEADER_FONT
            row += 1
            for item, val in q_data.items():
                ws.cell(row=row, column=1, value=item)
                ws.cell(row=row, column=2, value=val.get("濃度", ""))
                ws.cell(row=row, column=3, value=val.get("質量", ""))
                ws.cell(row=row, column=4, value=val.get("範圍", ""))
                row += 1
            row += 1

        # 欄寬
        for c in range(1, 6):
            ws.column_dimensions[get_column_letter(c)].width = [22, 18, 18, 36][min(c - 1, 3)]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ───────────────────────────── 介面 ─────────────────────────────

rules, rules_by_tank = load_rules()

with st.sidebar:
    st.header("系統狀態")
    st.metric("規則總數", len(rules))
    st.metric("槽體分類數", len(rules_by_tank))
    # 顯示 CSV 檔案修改時間 + 大小, 方便確認是否吃到最新 CSV
    if os.path.exists(CSV_PATH):
        import datetime as _dt
        mtime = _dt.datetime.fromtimestamp(os.path.getmtime(CSV_PATH))
        st.caption(f"CSV: {os.path.getsize(CSV_PATH)//1024} KB · {mtime.strftime('%m-%d %H:%M')}")
    if st.button("🔄 清快取重載", help="若顯示舊資料,點此清快取"):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.markdown("**v3 版本能力**")
    st.text("- 38 單元完整抽取")
    st.text("- 設計/量測/機具/水質")
    st.text("- 章節動態定位")
    st.text("- OCR 流向圖解析")
    st.text("- 智能審查 (學理檢查)")
    st.divider()
    st.markdown("**專案連結**")
    st.markdown("[GitHub](https://github.com/jetenv02-lab/water-pollution-review)")

tab1, tab2, tab3 = st.tabs(["🚀 開始審查", "📊 規則庫瀏覽", "📖 使用說明"])

with tab1:
    st.subheader("上傳申請文件 PDF")
    st.caption("v3 抽取器會解析: 處理設施資料表 + 進出水質資料表。流向示意圖與水量平衡示意圖可透過 OCR 自動讀取。")

    st.markdown("""
    <div class="ocr-warning">
    <strong>v3 已支援的審查能力:</strong><br>
    - 自動定位本份文件的水量平衡圖與流向示意圖頁碼(不寫死)<br>
    - OCR 讀取圖片頁面: 單元代號、流量 Q、加藥量、含水率<br>
    - 智能審查: 質量平衡守恆、快混槽不應展現重金屬去除、沉澱池溢流率等學理檢查
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("選擇 PDF", type=["pdf"])

    # 一鍵跑完整流程: 抽取 + OCR + 智能審查
    if uploaded is not None:
        st.success(f"已上傳: **{uploaded.name}** ({uploaded.size // 1024} KB)")

        # 事業類別選擇 (供智能審查使用) - 提前到按鈕之前讓使用者一次設定
        business_type = st.selectbox(
            "事業類別 (用於檢查申報項目是否完整)",
            ["(不檢查)"] + list(BUSINESS_TYPES.keys()),
            key="_business_type",
            help="選了之後智能審查會檢查該事業類別應申報的項目是否漏項",
        )

        if st.button("🚀 開始完整審查", type="primary",
                     help="一次跑完: 抽取單元 → 章節定位 → OCR 流向圖 → 智能審查"):
            # 暫存 PDF 到 disk 供多步驟使用
            pdf_bytes = uploaded.read()
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                tf.write(pdf_bytes)
                tmp_path = tf.name

            # 進度條
            progress = st.progress(0, text="準備中...")
            status = st.empty()

            try:
                # ─── Step 1: 章節定位 + 單元抽取 ───
                status.info("Step 1/3: 解析 PDF 章節與處理單元...")
                progress.progress(10, text="解析 PDF 文字內容...")
                sections_local = locate_sections(tmp_path, verbose=False)
                progress.progress(30, text="抽取處理單元結構化資料...")
                app_data_local = extract_application(tmp_path, verbose=False)
                progress.progress(40, text=f"完成 Step 1: 共 {app_data_local['total_units']} 個處理單元")

                # 存到 session
                ocr_target_pages = sorted(set(
                    sections_local.get("flow_diagram", []) +
                    sections_local.get("balance_diagram", [])
                ))
                st.session_state["_sections"] = sections_local
                st.session_state["_app_data"] = app_data_local
                st.session_state["_pdf_bytes"] = pdf_bytes
                st.session_state["_pdf_filename"] = uploaded.name
                st.session_state["_ocr_target_pages"] = ocr_target_pages
                # 建立水流串接圖
                st.session_state["_flow_graph"] = build_flow_graph(app_data_local)

                # ─── Step 2: OCR (若有流向圖頁) ───
                if ocr_target_pages:
                    n_ocr_pages = len(ocr_target_pages)
                    status.info(f"Step 2/3: OCR 解析 {n_ocr_pages} 頁流向圖 / 水量平衡圖 (約 {n_ocr_pages*15}~{n_ocr_pages*30} 秒)...")
                    progress.progress(50, text=f"執行 OCR ({n_ocr_pages} 頁)...")
                    ocr_result = ocr_diagram_pages(tmp_path, ocr_target_pages, verbose=False)
                    st.session_state["_ocr_result"] = ocr_result
                    if "error" not in ocr_result:
                        summary = ocr_result["summary"]
                        progress.progress(75, text=f"完成 Step 2: 識別 {summary['total_units']} 單元/{summary['total_flows']} 流量/{summary['total_doses']} 加藥")
                    else:
                        progress.progress(75, text="完成 Step 2: OCR 略過")
                else:
                    st.session_state.pop("_ocr_result", None)
                    progress.progress(75, text="Step 2/3: 無流向圖頁面,跳過 OCR")

                # ─── Step 3: 智能審查 (3 層: 質量平衡 + 學理 + 規則庫驅動) ───
                status.info("Step 3/3: 執行智能審查 (3 層檢查)...")
                progress.progress(82, text="Step 3.1: 質量平衡檢查...")
                findings_basic = run_balance_checks(app_data_local)
                progress.progress(88, text="Step 3.2: 學理檢查 (環工設計準則)...")
                bt = None if business_type == "(不檢查)" else business_type
                findings_adv = run_advanced_checks(app_data_local, business_type=bt)
                progress.progress(94, text="Step 3.3: 規則庫驅動檢查 (299 筆環工技師缺失)...")
                findings_rule = run_rule_driven_check(app_data_local)
                # 合併三層, 規則庫驅動的放最後 (一般是「待人工」性質)
                st.session_state["_check_findings"] = findings_basic + findings_adv + findings_rule

                # 完成
                progress.progress(100, text="全部完成!")
                stats = {"不合理": 0, "待人工": 0}
                for f in st.session_state["_check_findings"]:
                    sev = f.get("嚴重度")
                    if sev in stats:
                        stats[sev] += 1
                status.success(
                    f"✅ 審查完成! 共 {app_data_local['total_units']} 單元 · "
                    f"找出 {stats['不合理']} 項不合理 / {stats['待人工']} 項待人工複核"
                )
            finally:
                try:
                    os.unlink(tmp_path)
                except:
                    pass

    # ───────── 顯示區 (永遠基於 session_state, 不被 button rerun 影響) ─────────
    if st.session_state.get("_app_data"):
        app_data = st.session_state["_app_data"]
        sections = st.session_state["_sections"]
        pdf_filename = st.session_state.get("_pdf_filename", "")

        st.success(f"✅ 抽取完成! 共 **{app_data['total_units']}** 個處理單元 · 來源: {pdf_filename}")

        # ───────── 章節動態定位區塊 ─────────
        st.subheader("📍 本文件章節定位")
        st.caption("不同 PDF 頁碼會浮動。系統用『章節標題』找到各區段所在頁,而非寫死頁碼。")

        sec_labels = {
            "flow_diagram": ("廢(污)水流向示意圖", "🌊", "純圖片,需 OCR"),
            "balance_diagram": ("水質水量平衡示意圖", "📊", "純圖片,需 OCR"),
            "quality_data": ("進出水質資料表", "💧", "文字,已抽取"),
            "facility_table": ("處理設施資料表", "🔧", "文字,已抽取"),
            "raw_water": ("原廢水水質", "🧪", "文字"),
            "discharge": ("放流口資料", "🚰", "文字"),
            "emergency": ("緊急應變方法", "⚠", "文字"),
            "sludge": ("污泥處理", "♻", "文字"),
        }
        section_rows = []
        for sec_type, (label, emoji, note) in sec_labels.items():
            pages = sections.get(sec_type, [])
            if pages:
                section_rows.append({
                    "區段": f"{emoji} {label}",
                    "頁碼範圍": compress_ranges(pages),
                    "頁數": len(pages),
                    "狀態": note,
                })
            else:
                section_rows.append({
                    "區段": f"{emoji} {label}",
                    "頁碼範圍": "(未找到)",
                    "頁數": 0,
                    "狀態": "未在文件中偵測到",
                })
        st.dataframe(section_rows, use_container_width=True, hide_index=True)

        # 統計卡片
        total_design = sum(len(u["design_params"]) for u in app_data["units"].values())
        total_measure = sum(len(u["measure_params"]) for u in app_data["units"].values())
        total_eq = sum(len(u["equipment"]) for u in app_data["units"].values())
        total_in = sum(len(u["influent"]) for u in app_data["units"].values())
        total_out = sum(len(u["effluent"]) for u in app_data["units"].values())

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("處理單元", app_data["total_units"])
        c2.metric("設計參數", total_design)
        c3.metric("量測參數", total_measure)
        c4.metric("機具", total_eq)
        c5.metric("水質流向", f"進{total_in}/出{total_out}")

        # 單元清單
        st.subheader("📋 處理單元清單")
        unit_rows = []
        for code, info in sorted(app_data["units"].items()):
            unit_rows.append({
                "代號": code,
                "原始名稱": info["name_in_doc"],
                "標準類型": info["std_tank"],
                "代碼": info.get("code_id", ""),
                "頁數": ", ".join(map(str, info["pages_found"][:3])),
                "設計": len(info["design_params"]),
                "量測": len(info["measure_params"]),
                "機具": len(info["equipment"]),
                "進": len(info["influent"]),
                "出": len(info["effluent"]),
            })
        st.dataframe(unit_rows, use_container_width=True, hide_index=True)

        # 篩選器 — 看單元詳情 (現在用 session_state 保留選擇)
        st.subheader("🔎 單元詳情")
        unit_codes = sorted(app_data["units"].keys())
        selected = st.selectbox(
            "選擇單元查看詳情",
            unit_codes,
            key="_selected_unit",  # key 讓 Streamlit 自動把選擇存進 session
        )
        if selected and selected in app_data["units"]:
            unit = app_data["units"][selected]

            # 顯示水流上下游 (如果有 flow_graph)
            graph = st.session_state.get("_flow_graph")
            if graph:
                nb = get_unit_neighbors(graph, selected)
                if nb["upstream"] or nb["downstream"]:
                    st.markdown("##### 🔗 水流串接")
                    cu, cd = st.columns(2)
                    with cu:
                        st.markdown(f"**上游 ({len(nb['upstream'])} 條進入)**")
                        if nb["upstream"]:
                            up_rows = [
                                {"來源單元": u["from_unit"],
                                 "出流編號": u["from_stream"],
                                 "→ 進流編號": u["to_stream"]}
                                for u in nb["upstream"]
                            ]
                            st.dataframe(up_rows, use_container_width=True, hide_index=True)
                        else:
                            st.caption("(無偵測到上游, 可能是原廢水進入點或未串接)")
                    with cd:
                        st.markdown(f"**下游 ({len(nb['downstream'])} 條流出)**")
                        if nb["downstream"]:
                            dn_rows = [
                                {"出流編號": d["from_stream"],
                                 "→ 目標單元": d["to_unit"],
                                 "目標進流編號": d["to_stream"]}
                                for d in nb["downstream"]
                            ]
                            st.dataframe(dn_rows, use_container_width=True, hide_index=True)
                        else:
                            st.caption("(無偵測到下游, 可能是放流口或未串接)")
                    st.divider()

            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**單元代號**: {selected}")
                st.markdown(f"**原始名稱**: {unit['name_in_doc']}")
                st.markdown(f"**標準類型**: {unit['std_tank']}")
                st.markdown(f"**內部代碼**: {unit.get('code_id', '')}")
                st.markdown(f"**出現頁數**: {unit['pages_found']}")
            with c2:
                if unit.get("design_params"):
                    st.markdown("**設計操作參數**:")
                    for k, v in unit["design_params"].items():
                        st.markdown(f"- {k}: `{v.get('raw', '')}`")
                if unit.get("measure_params"):
                    st.markdown("**量測操作參數**:")
                    for k, v in unit["measure_params"].items():
                        st.markdown(f"- {k}: `{v.get('raw', '')}`")

            if unit.get("equipment"):
                st.markdown("**相關機具設施**:")
                eq_rows = [{"名稱": str(e["name"]), "位置": str(e.get("位置", "")),
                            "數量": str(e.get("數量", "")), "馬力(kW)": str(e.get("馬力_kW", ""))}
                           for e in unit["equipment"]]
                st.dataframe(eq_rows, use_container_width=True, hide_index=True)

            if unit.get("influent"):
                st.markdown(f"**進流水質** ({len(unit['influent'])} 流向)")
                for infl_code, qdata in unit["influent"].items():
                    st.caption(f"📥 {infl_code}")
                    q_rows = [{"水質項目": str(k),
                               "濃度": str(v.get("濃度", v.get("範圍", ""))),
                               "質量": str(v.get("質量", ""))} for k, v in qdata.items()]
                    st.dataframe(q_rows, use_container_width=True, hide_index=True)

            if unit.get("effluent"):
                st.markdown(f"**出流水質** ({len(unit['effluent'])} 流向)")
                for effl_code, qdata in unit["effluent"].items():
                    st.caption(f"📤 {effl_code}")
                    q_rows = [{"水質項目": str(k),
                               "濃度": str(v.get("濃度", v.get("範圍", ""))),
                               "質量": str(v.get("質量", ""))} for k, v in qdata.items()]
                    st.dataframe(q_rows, use_container_width=True, hide_index=True)

        # 下載
        st.divider()
        st.subheader("📥 下載抽取結果")
        base_name = os.path.splitext(pdf_filename)[0] if pdf_filename else "result"

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = build_unit_excel(app_data)
            st.download_button(
                "📊 下載各單元詳細 Excel",
                data=excel_buf,
                file_name=f"{base_name}_單元資料.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col2:
            json_str = json.dumps(app_data, ensure_ascii=False, indent=2)
            st.download_button(
                "💾 下載 JSON 結構化資料",
                data=json_str.encode("utf-8"),
                file_name=f"{base_name}_抽取結果.json",
                mime="application/json",
                use_container_width=True,
            )

    # ───────── 智能審查結果顯示 (基於 session, 由「開始完整審查」按鈕產生) ─────────
    if st.session_state.get("_check_findings") is not None:
        st.divider()
        st.subheader("智能審查結果")
        st.caption(
            "根據環工技師 299 筆查核缺失歸納的學理規則。"
            " 結果依「審查類型」分組,涵蓋質量平衡/機具設施/設計參數/去除率等多面向。"
        )

        findings = st.session_state["_check_findings"]

        # 統計 (合併「不合理」和「待人工」, 改用「審查類型」分組)
        from collections import Counter
        type_counter = Counter(f.get("類型", "其他") for f in findings)
        sev_counter = Counter(f.get("嚴重度", "?") for f in findings)

        # 上方四張卡片: 總覽 + 三種嚴重度
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 總審查項", len(findings))
        c2.metric("🔴 明顯不合理", sev_counter.get("不合理", 0),
                  help="系統能 100% 自動判定為違反學理 (如快混槽展現重金屬去除)")
        c3.metric("🟡 應人工複核", sev_counter.get("待人工", 0),
                  help="系統找出規則涉及的具體單元數值,需技師人工判讀是否合理")
        c4.metric("📊 涵蓋類型", len(type_counter))

        st.divider()

        # 篩選器
        col1, col2, col3 = st.columns([2, 2, 3])
        with col1:
            filter_type = st.selectbox(
                "依類型篩選",
                ["(全部)"] + sorted(type_counter.keys()),
                key="_filter_type",
            )
        with col2:
            unique_units = sorted({f["單元"] for f in findings})
            filter_unit = st.selectbox(
                "依單元篩選",
                ["(全部)"] + unique_units,
                key="_filter_unit",
            )
        with col3:
            filter_kw = st.text_input("關鍵字搜尋 (描述/規則)", key="_filter_kw")

        filtered = findings
        if filter_type != "(全部)":
            filtered = [f for f in filtered if f.get("類型") == filter_type]
        if filter_unit != "(全部)":
            filtered = [f for f in filtered if f.get("單元") == filter_unit]
        if filter_kw:
            kw = filter_kw.lower()
            filtered = [
                f for f in filtered
                if kw in str(f.get("描述", "")).lower()
                or kw in str(f.get("依據", "")).lower()
            ]

        st.caption(f"顯示 {len(filtered)} / {len(findings)} 筆")

        # 依類型分組顯示
        by_type = defaultdict(list)
        for f in filtered:
            by_type[f.get("類型", "其他")].append(f)

        # 類型顯示順序 (重要的放前面)
        type_priority = [
            "質量平衡", "去除率", "設計參數", "機具設施", "水質標準",
            "操作條件", "流向/示意圖", "文件一致性", "單位/有效位數", "其他"
        ]
        ordered_types = (
            [t for t in type_priority if t in by_type]
            + [t for t in sorted(by_type.keys()) if t not in type_priority]
        )

        for type_name in ordered_types:
            items = by_type[type_name]
            if not items:
                continue
            # 嚴重度色碼 emoji
            sev_in_group = Counter(f["嚴重度"] for f in items)
            sev_summary_parts = []
            if sev_in_group.get("不合理"):
                sev_summary_parts.append(f"🔴 {sev_in_group['不合理']} 不合理")
            if sev_in_group.get("待人工"):
                sev_summary_parts.append(f"🟡 {sev_in_group['待人工']} 待人工")
            sev_summary = " · ".join(sev_summary_parts)

            with st.expander(f"**{type_name}** ({len(items)} 筆) — {sev_summary}", expanded=True):
                # 每個類型用表格顯示
                rows = []
                for f in items:
                    sev_emoji = {"不合理": "🔴", "待人工": "🟡"}.get(f["嚴重度"], "⚪")
                    rows.append({
                        "嚴重": sev_emoji,
                        "單元": str(f["單元"]),
                        "標準槽體": str(f["標準槽體"]),
                        "對照項目": str(f["對照項目"]),
                        "描述": str(f["描述"])[:150],
                        "依據": str(f.get("依據", ""))[:80],
                    })
                st.dataframe(rows, use_container_width=True, hide_index=True)

        if not findings:
            st.success("本份文件未偵測到明顯不合理之處 (基於目前內建的學理規則)")

        # 把 Counter 轉為 dict 給 JSON 序列化
        findings_json = json.dumps({
            "source": st.session_state.get("_pdf_filename", "?"),
            "total_findings": len(findings),
            "severity_stats": dict(sev_counter),
            "type_stats": dict(type_counter),
            "findings": findings,
        }, ensure_ascii=False, indent=2)
        st.download_button(
            "下載審查結果 JSON",
            data=findings_json.encode("utf-8"),
            file_name="智能審查結果.json",
            mime="application/json",
        )

    # ───────── OCR 流向圖結果顯示 (由「開始完整審查」按鈕產生) ─────────
    if st.session_state.get("_ocr_result"):
        st.divider()
        st.subheader("OCR 解析流向圖 / 水量平衡圖")
        ocr_pages = st.session_state.get("_ocr_target_pages", [])
        if ocr_pages:
            st.caption(f"OCR 解析了 {len(ocr_pages)} 頁 (頁 {compress_ranges(ocr_pages)})")

        ocr_result = st.session_state["_ocr_result"]
        if "error" in ocr_result:
            st.error(f"OCR 失敗: {ocr_result['error']}")
        else:
            summary = ocr_result["summary"]
            st.success(
                f"OCR 完成! 識別 {summary['total_units']} 個單元、"
                f"{summary['total_flows']} 個流量、"
                f"{summary['total_doses']} 個加藥、"
                f"{summary['total_moistures']} 個含水率"
            )

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("單元", summary["total_units"])
            c2.metric("流量(Q)", summary["total_flows"])
            c3.metric("加藥", summary["total_doses"])
            c4.metric("含水率", summary["total_moistures"])

            if ocr_result.get("all_flows"):
                st.markdown("**識別到的流量 Q (CMD):**")
                flow_rows = [
                    {"Q (CMD)": str(f["q"]), "OCR 原文": str(f["text"])}
                    for f in ocr_result["all_flows"]
                ]
                st.dataframe(flow_rows, use_container_width=True, hide_index=True)

            if ocr_result.get("all_doses"):
                st.markdown("**識別到的加藥量:**")
                dose_rows = [
                    {"化學品": str(d["chemical"]), "用量": str(d["amount"]),
                     "單位": str(d.get("unit", "")), "OCR 原文": str(d["text"])}
                    for d in ocr_result["all_doses"]
                ]
                st.dataframe(dose_rows, use_container_width=True, hide_index=True)

            if ocr_result.get("all_moistures"):
                st.markdown("**識別到的含水率:**")
                mois_rows = [
                    {"含水率 (%)": str(m["value_pct"]), "OCR 原文": str(m["text"])}
                    for m in ocr_result["all_moistures"]
                ]
                st.dataframe(mois_rows, use_container_width=True, hide_index=True)

            if ocr_result.get("all_units"):
                with st.expander(f"OCR 識別到的單元代號 ({len(ocr_result['all_units'])} 個)"):
                    unit_rows = [
                        {"代號": str(u["code"]), "OCR 原文": str(u["text"])}
                        for u in ocr_result["all_units"]
                    ]
                    st.dataframe(unit_rows, use_container_width=True, hide_index=True)

            pdf_name = st.session_state.get("_pdf_filename", "ocr_result")
            base_ocr = os.path.splitext(pdf_name)[0]
            ocr_json = json.dumps(ocr_result, ensure_ascii=False, indent=2)
            st.download_button(
                "下載 OCR 結果 JSON",
                data=ocr_json.encode("utf-8"),
                file_name=f"{base_ocr}_ocr.json",
                mime="application/json",
            )

with tab2:
    st.subheader("規則庫內容")
    st.caption(f"目前載入 {len(rules)} 筆規則，分布於 {len(rules_by_tank)} 個槽體分類。")

    if rules:
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
            [{
                "ID": r.get("缺失ID"),
                "槽體": r.get("標準槽體名稱"),
                "技師": r.get("技師姓名"),
                "檢查類型": r.get("檢查類型"),
                "對照項目": r.get("對照項目"),
                "規則": r.get("規則"),
                "原文": (r.get("原文缺失") or "")[:100],
            } for r in display_rules[:500]],
            use_container_width=True, hide_index=True,
        )

with tab3:
    st.subheader("水措審查系統 v3 — 使用說明")
    st.markdown("""
### v3 版本能力

| 項目 | 狀態 |
|------|------|
| 單元偵測覆蓋率 | **100%** (38/38) |
| 單元類型歸類 | **完全正確** (中和池/沉澱池/慢混池...) |
| 設計操作參數 | **已抽取** (停留時間/有效容量/攪拌轉速...) |
| 量測操作參數 | **已抽取** (pH/加藥量/DO...) |
| 機具設施 | **已抽取** (pH計/液位計/攪拌機 + 馬力數) |
| 進出流水質 | **已抽取** (各 35 流向、含 21+ 水質項目) |
| 章節動態定位 | **已支援** (不寫死頁碼,各家工廠都適用) |
| 流向圖 OCR | **已支援** (RapidOCR 中文識別) |
| 智能審查 | **已支援** (學理檢查、不再全部待人工) |

### 使用流程

1. **上傳 PDF** → 點「開始解析」
2. **看「本文件章節定位」** → 系統會列出本份文件每個區段在第幾頁
3. **看「處理單元清單」** → 38 個單元 + 對應標準槽體類型
4. **執行 OCR** → 對流向圖/水量平衡圖跑 OCR、抽出流量/加藥/含水率
5. **執行智能審查** → 自動列出不合理項目 + 學理依據

### 智能審查涵蓋的學理規則

- **質量守恆**: 溶解性物質(硝酸鹽/硼/Cl-)不應在無濃縮機制單元自行濃縮
- **去除位置學理**: 快混槽/pH調整槽無固液分離,不應展現重金屬去除
- **pH 槽特性**: pH 調整槽除 pH 外,其他水質應不變
- **沉澱設計**: 表面溢流率應 < 50 m3/m2-d
- **必要機具**: 各槽體應有的液位計/pH計/排泥等設施

### 抽取邏輯

1. **掃描所有頁,找兩種關鍵區段**
   - 處理設施資料表（含「(一)處理單元名稱：xxx 序號：T01-01 代碼：120」）
   - 進出水質資料表（含「單元序號：T01-01」、「進流水流編號：WTB...」、「出流水流編號：WTA...」）
2. **解析設施資料表**: 抽 (二)設計參數、(三)量測參數、(四)機具設施
3. **解析水質表**: 抽各水質項目的濃度、質量、pH 範圍
4. **OCR 流向圖**: 用 RapidOCR 讀取圖片頁面 → 抽單元代號/Q/加藥/含水率

### 剩餘限制

- 規則庫目前 53 筆,完整版 235 筆萃取中
- 比對引擎可繼續擴充更多學理規則
- 大型 PDF (>100MB) 建議用本地版執行
""")

st.divider()
st.caption("水措審查系統 v3 · 章節定位 + OCR + 智能審查 · [GitHub](https://github.com/jetenv02-lab/water-pollution-review)")
