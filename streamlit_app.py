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

    # ───────── 本次 session 審查歷史 ─────────
    history = st.session_state.get("_review_history", [])
    if history:
        st.markdown("**📋 本次審查歷史**")
        st.caption("(本次瀏覽 session · 永久紀錄請看「📊 規則庫瀏覽」)")
        # 最近一次的 Sheet 寫入狀態
        log_status = st.session_state.get("_last_sheet_log_status")
        if log_status:
            st.caption(log_status)
        for i, h in enumerate(history[:5]):
            with st.container(border=True):
                fname_short = h["filename"]
                if len(fname_short) > 22:
                    fname_short = fname_short[:20] + "..."
                st.caption(f"⏱ {h['time']} · {h['elapsed_sec']}s")
                st.markdown(f"**{fname_short}**")
                st.caption(
                    f"{h['units']} 單元 · 🔴 {h['unreasonable']} · 🟡 {h['manual']}"
                )
        if len(history) > 5:
            st.caption(f"還有 {len(history)-5} 筆紀錄...")
        if st.button("🗑 清除歷史", key="_clear_history"):
            st.session_state.pop("_review_history", None)
            st.rerun()
        st.divider()

    st.markdown("**v3 版本能力**")
    st.text("- 38 單元完整抽取")
    st.text("- 設計/量測/機具/水質")
    st.text("- 章節動態定位")
    st.text("- OCR 流向圖解析")
    st.text("- 智能審查 (學理檢查)")
    st.text("- 水流串接圖")
    st.divider()
    st.markdown("**專案連結**")
    st.markdown("[GitHub](https://github.com/jetenv02-lab/water-pollution-review)")

    # 顯示 import 失敗模組 (debug 用)
    if _import_errors:
        with st.expander(f"⚠ 模組載入問題 ({len(_import_errors)})"):
            for mod, err in _import_errors:
                st.caption(f"**{mod}**: {err[:150]}")

# ───────── 全頁面 busy 橫幅 (置頂,不論在哪個 tab 都會看到) ─────────
if st.session_state.get("_busy", False):
    st.error(
        "⏳ **系統忙碌中,請勿關閉/重整頁面,避免操作其他按鈕** — 可能會中斷處理"
    )

tab1, tab2, tab_sync, tab_import, tab3 = st.tabs(["🚀 開始審查", "📊 規則庫瀏覽", "🔄 規則庫管理", "📥 匯入新規則", "📖 使用說明"])

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

        # 章節擷取 toggle — 只處理「參、水污染防治措施資料」段
        col_opt1, col_opt2 = st.columns([2, 3])
        with col_opt1:
            only_chapter_iii = st.checkbox(
                "✂️ 只處理「參、水污染防治措施資料」章節",
                value=True,
                key="_only_ch3",
                help="自動偵測該章節, 只抽該段內容 (省 token + 加速)。"
                     "找不到章節會自動 fallback 用全文。",
            )
        with col_opt2:
            if only_chapter_iii:
                st.caption(
                    "✅ 系統會先掃 PDF 找「參、」章節, 切出該段給 OCR/Gemini, 通常加速 5-10 倍"
                )
            else:
                st.caption("⚠️ 處理整本 PDF (慢、貴, 但完整)")

        # 事業類別選擇 (供智能審查使用) - 提前到按鈕之前讓使用者一次設定
        business_type = st.selectbox(
            "事業類別 (用於檢查申報項目是否完整)",
            ["(不檢查)"] + list(BUSINESS_TYPES.keys()),
            key="_business_type",
            help="選了之後智能審查會檢查該事業類別應申報的項目是否漏項",
        )

        # 🔒 若正在跑審查, button 禁用
        _is_busy = st.session_state.get("_busy", False)
        if _is_busy:
            st.warning(
                "⏳ **審查進行中,請勿關閉頁面或操作其他功能** — 此操作通常需要 30秒~2分鐘"
            )
        if st.button("🚀 開始完整審查", type="primary",
                     disabled=_is_busy,
                     help="一次跑完: 抽取單元 → 章節定位 → OCR 流向圖 → 智能審查"):
            # 標記正在跑審查 (鎖其他主要按鈕)
            st.session_state["_busy"] = True
            # 暫存 PDF 到 disk 供多步驟使用
            pdf_bytes = uploaded.read()
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                tf.write(pdf_bytes)
                tmp_path = tf.name

            # 若啟用章節擷取, 先找參章位置 (給後續 step 參考)
            chapter_info = None
            if only_chapter_iii:
                try:
                    import pdf_chapter_extractor
                    chapter_info = pdf_chapter_extractor.find_chapter_iii_pages(pdf_bytes)
                    if chapter_info.get("ok") and chapter_info.get("found"):
                        st.info(
                            f"✂️ 找到「參、水污染防治措施資料」: "
                            f"頁 {chapter_info['start_page']}~{chapter_info['end_page']} "
                            f"(全本 {chapter_info['total_pages']} 頁, "
                            f"節省 {100 - (chapter_info['end_page'] - chapter_info['start_page'] + 1) * 100 // chapter_info['total_pages']}%)"
                        )
                    elif chapter_info.get("ok"):
                        st.warning("⚠️ 找不到「參、水污染防治措施資料」章節, 改用全文")
                        chapter_info = None
                except Exception as _e:
                    st.caption(f"章節定位失敗 ({_e}), 改用全文")
                    chapter_info = None
            # 把章節範圍存到 session_state, 給其他模組讀
            st.session_state["_chapter_iii_info"] = chapter_info

            # 進度條 + 剩餘時間估算
            import time as _time
            t_start = _time.time()

            # 基準預估時間 (秒, 來自實測秋棠案例)
            # Step1: 60s, Step2 OCR: 90s, Step3: 5s, 總共 ~155s
            _BASELINE_TOTAL = 155

            def _eta(percent_done):
                """估算剩餘秒數。

                早期 (已耗時 < 3 秒) 用基準預估值,
                晚期改用「已耗時 / 已完成%」實測推估 (更準)。
                """
                if percent_done <= 0:
                    return "預估中..."
                elapsed = _time.time() - t_start
                # 早期用基準, 避免「跑 0.5 秒 就以為總共 5 秒」
                if elapsed < 3:
                    remaining = _BASELINE_TOTAL * (100 - percent_done) / 100
                else:
                    total_est = elapsed / (percent_done / 100)
                    remaining = max(0, total_est - elapsed)
                if remaining < 1:
                    return "即將完成"
                if remaining < 60:
                    return f"約 {remaining:.0f} 秒"
                return f"約 {remaining/60:.1f} 分鐘"

            progress = st.progress(0, text="準備中...")
            status = st.empty()

            try:
                # ─── Step 1: 章節定位 + 單元抽取 ───
                status.info("Step 1/3: 解析 PDF 章節與處理單元...")
                progress.progress(10, text=f"[{_eta(10)}] 解析 PDF 文字內容...")
                sections_local = locate_sections(tmp_path, verbose=False)
                progress.progress(30, text=f"[{_eta(30)}] 抽取處理單元結構化資料...")
                app_data_local = extract_application(tmp_path, verbose=False)
                progress.progress(40, text=f"[{_eta(40)}] 完成 Step 1: 共 {app_data_local['total_units']} 個處理單元")

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
                    status.info(f"Step 2/3: OCR 解析 {n_ocr_pages} 頁流向圖 / 水量平衡圖")
                    progress.progress(50, text=f"[{_eta(50)}] 執行 OCR ({n_ocr_pages} 頁)...")
                    ocr_result = ocr_diagram_pages(tmp_path, ocr_target_pages, verbose=False)
                    st.session_state["_ocr_result"] = ocr_result
                    if "error" not in ocr_result:
                        summary = ocr_result["summary"]
                        progress.progress(75, text=f"[{_eta(75)}] 完成 Step 2: 識別 {summary['total_units']} 單元/{summary['total_flows']} 流量/{summary['total_doses']} 加藥")
                    else:
                        progress.progress(75, text=f"[{_eta(75)}] 完成 Step 2: OCR 略過")
                else:
                    st.session_state.pop("_ocr_result", None)
                    progress.progress(75, text=f"[{_eta(75)}] Step 2/3: 無流向圖頁面,跳過 OCR")

                # ─── Step 3: 智能審查 (3 層: 質量平衡 + 學理 + 規則庫驅動) ───
                status.info("Step 3/3: 執行智能審查 (3 層檢查)...")
                progress.progress(82, text=f"[{_eta(82)}] Step 3.1: 質量平衡檢查...")
                findings_basic = run_balance_checks(app_data_local)
                progress.progress(88, text=f"[{_eta(88)}] Step 3.2: 學理檢查 (環工設計準則)...")
                bt = None if business_type == "(不檢查)" else business_type
                findings_adv = run_advanced_checks(app_data_local, business_type=bt)
                progress.progress(94, text=f"[{_eta(94)}] Step 3.3: 規則庫驅動檢查 (299 筆環工技師缺失)...")
                findings_rule = run_rule_driven_check(app_data_local)
                # 合併三層, 規則庫驅動的放最後 (一般是「待人工」性質)
                st.session_state["_check_findings"] = findings_basic + findings_adv + findings_rule

                # 完成
                total_elapsed = _time.time() - t_start
                progress.progress(100, text=f"全部完成! 共耗時 {total_elapsed:.0f} 秒")
                stats = {"不合理": 0, "待人工": 0}
                for f in st.session_state["_check_findings"]:
                    sev = f.get("嚴重度")
                    if sev in stats:
                        stats[sev] += 1
                status.success(
                    f"✅ **基本審查已完成**! 共 {app_data_local['total_units']} 單元 · "
                    f"找出 {stats['不合理']} 項不合理 / {stats['待人工']} 項待人工複核 · 耗時 {total_elapsed:.0f} 秒"
                )
                # 提示下一步
                st.info(
                    "💡 **下一步**: 你可以繼續往下捲動,展開「📊 水量平衡示意圖解析」 "
                    "讓 Gemini Vision 讀流向圖, 得到跨單元的完整流向結構 + 質量平衡檢核"
                )

                # 記錄到 session 歷史 (重整就消失, 但本次 session 可看)
                from datetime import datetime as _dt
                if "_review_history" not in st.session_state:
                    st.session_state["_review_history"] = []
                review_record = {
                    "time": _dt.now().strftime("%H:%M:%S"),
                    "filename": uploaded.name,
                    "units": app_data_local["total_units"],
                    "unreasonable": stats["不合理"],
                    "manual": stats["待人工"],
                    "elapsed_sec": int(total_elapsed),
                }
                st.session_state["_review_history"].insert(0, review_record)
                # 限制歷史只保留最近 10 筆
                st.session_state["_review_history"] = st.session_state["_review_history"][:10]

                # 同時把這筆寫到 Sheet 的 _審查紀錄 分頁 (跨 session 持久化)
                try:
                    import review_history
                    sheet_result = review_history.append_review_record(review_record)
                    if sheet_result.get("ok"):
                        st.session_state["_last_sheet_log_status"] = (
                            f"✅ 紀錄已存到 Sheet (第 {sheet_result['review_times']} 次審查此文件)"
                        )
                    else:
                        st.session_state["_last_sheet_log_status"] = (
                            f"⚠️ Sheet 紀錄失敗: {sheet_result.get('error', '?')[:80]}"
                        )
                except Exception as _e:
                    st.session_state["_last_sheet_log_status"] = f"⚠️ Sheet 紀錄失敗: {_e}"
            finally:
                try:
                    os.unlink(tmp_path)
                except:
                    pass
                # 解鎖 UI
                st.session_state["_busy"] = False

    # ───────── 顯示區 (永遠基於 session_state, 不被 button rerun 影響) ─────────
    # ── 水量平衡示意圖解析 (Gemini Vision 抽結構化流向) ──
    pdf_bytes_for_flow = st.session_state.get("_pdf_bytes")
    if pdf_bytes_for_flow:
        with st.expander("📊 水量平衡示意圖解析 (跨單元流向 + 質平檢核)", expanded=False):
            st.caption(
                "從申請文件的「參、水污染防治措施資料 / 水質水量平衡示意圖」抽出流向結構,"
                " 可得知: 每個單元有幾條進流/出流、來源/去處、流量 Q, 並做跨單元質量平衡檢核。"
            )

            try:
                import flow_diagram_extractor as _fde
                import gemini_extractor as _ge_chk
                fde_ok = True
            except Exception as _fe:
                st.error(f"無法載入 flow_diagram_extractor: {_fe}")
                fde_ok = False

            if fde_ok:
                _fstat = _ge_chk.check_gemini_status()
                if not _fstat["ok"]:
                    st.warning(f"⚠️ {_fstat['message']} — 此功能需要 Gemini Vision")
                else:
                    # 偵測圖頁
                    loc = _fde.find_balance_diagram_pages(pdf_bytes_for_flow)
                    if loc.get("ok") and loc.get("image_pages"):
                        img_pages = loc["image_pages"]
                        st.info(
                            f"📍 找到 **{len(img_pages)}** 張水量平衡示意圖頁: "
                            f"p{', p'.join(map(str, img_pages))}"
                        )
                        max_pages = st.slider(
                            "處理頁數 (省 token, 先測一張看效果)",
                            min_value=1, max_value=len(img_pages),
                            value=min(2, len(img_pages)),
                            key="_flow_max_pages",
                            help=f"每張圖約 $0.002~0.005, 共 {len(img_pages)} 張全跑約 ${len(img_pages) * 0.005:.3f}"
                        )

                        if "_flow_extract_result" not in st.session_state:
                            st.session_state["_flow_extract_result"] = None

                        if st.button("🤖 開始解析", type="primary", key="_btn_flow_extract",
                                     width="stretch",
                                     disabled=st.session_state.get("_busy", False)):
                            st.session_state["_busy"] = True
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            def _cb(cur, tot, msg):
                                pct = (cur / tot) if tot else 0
                                progress_bar.progress(min(pct, 1.0), text=msg)
                                status_text.caption(f"[{cur}/{tot}] {msg}")
                            with st.spinner("Gemini Vision 處理中…"):
                                fr = _fde.extract_all_balance_diagrams(
                                    pdf_bytes_for_flow,
                                    max_pages=max_pages,
                                    progress_callback=_cb,
                                )
                            progress_bar.empty()
                            status_text.empty()
                            st.session_state["_flow_extract_result"] = fr
                            # 解鎖
                            st.session_state["_busy"] = False

                        fr = st.session_state.get("_flow_extract_result")
                        if fr:
                            if not fr.get("ok"):
                                st.error(f"❌ 失敗 ({fr.get('stage')}): {fr.get('error')}")
                            else:
                                usage = fr.get("gemini_usage", {})
                                st.success(
                                    f"✅ **水量平衡示意圖解析完成** · "
                                    f"處理 {fr['pages_processed']} 頁 / "
                                    f"抽出 **{len(fr['all_units'])}** 單元 / "
                                    f"**{len(fr['all_flows'])}** 流向 / "
                                    f"**{len(fr['all_external_inputs'])}** 外部輸入 / "
                                    f"**{len(fr['all_discharge_points'])}** 放流口"
                                )
                                st.caption(
                                    f"Token: in={usage.get('input_tokens', '?')} "
                                    f"out={usage.get('output_tokens', '?')}"
                                )

                                if fr.get("errors"):
                                    with st.expander(f"⚠️ {len(fr['errors'])} 頁解析失敗"):
                                        for e in fr["errors"]:
                                            st.caption(f"p{e['page']}: {e['error'][:100]}")

                                tabs = st.tabs([
                                    "🌊 流向總表", "📦 單元清單",
                                    "⚖️ 質平檢核", "🔗 編號一致性",
                                    "🔄 跟系統比對", "📄 各頁原始 JSON",
                                ])

                                with tabs[0]:
                                    import pandas as _pd
                                    if fr["all_flows"]:
                                        df_f = _pd.DataFrame(fr["all_flows"]).astype(str)
                                        st.dataframe(df_f, width="stretch", hide_index=True)
                                    else:
                                        st.info("沒有抽到流向")

                                    if fr["all_external_inputs"]:
                                        st.markdown("**📥 外部原廢水 (WM)**")
                                        st.dataframe(
                                            _pd.DataFrame(fr["all_external_inputs"]).astype(str),
                                            width="stretch", hide_index=True
                                        )

                                    if fr["all_discharge_points"]:
                                        st.markdown("**📤 放流口 (D)**")
                                        st.dataframe(
                                            _pd.DataFrame(fr["all_discharge_points"]).astype(str),
                                            width="stretch", hide_index=True
                                        )

                                with tabs[1]:
                                    if fr["all_units"]:
                                        import pandas as _pd
                                        df_u = _pd.DataFrame(fr["all_units"]).astype(str)
                                        st.dataframe(df_u, width="stretch", hide_index=True)
                                        st.caption(f"圖中共 {len(fr['all_units'])} 個處理單元")
                                    else:
                                        st.info("沒有抽到單元")

                                with tabs[2]:
                                    bal = _fde.check_water_balance(fr)
                                    s = bal["summary"]
                                    cb1, cb2, cb3, cb4 = st.columns(4)
                                    cb1.metric("總單元", s["total_units"])
                                    cb2.metric("✅ 平衡 (<1%)", s["balanced_count"])
                                    cb3.metric("⚠️ 偏差 (1~5%)", s["warning_count"])
                                    cb4.metric("❌ 異常 (>5%)", s["error_count"])

                                    rows_b = []
                                    for code, info in bal["by_unit"].items():
                                        rows_b.append({
                                            "單元": code,
                                            "Σ進流 (CMD)": info["in_total_cmd"],
                                            "Σ出流 (CMD)": info["out_total_cmd"],
                                            "差異 (%)": info["diff_pct"],
                                            "警告": info["warning"] or "",
                                        })
                                    if rows_b:
                                        import pandas as _pd
                                        st.dataframe(
                                            _pd.DataFrame(rows_b).astype(str),
                                            width="stretch", hide_index=True
                                        )

                                with tabs[3]:
                                    # 編號一致性檢核
                                    consistency = _fde.check_stream_consistency(fr)
                                    s = consistency["summary"]
                                    sc1, sc2, sc3, sc4 = st.columns(4)
                                    sc1.metric("總編號數", s["total_streams"])
                                    sc2.metric("⚠️ 警告", s["warnings_count"])
                                    sc3.metric("🔍 上游有編號\n但下游漏抽", s["unmatched_from_count"])
                                    sc4.metric("🔍 下游有編號\n但上游漏抽", s["unmatched_to_count"])

                                    if consistency["warnings"]:
                                        st.markdown("**⚠️ 一致性警告**")
                                        warn_rows = []
                                        for w in consistency["warnings"]:
                                            warn_rows.append({
                                                "類型": w["type"],
                                                "編號": w.get("stream", ""),
                                                "訊息": w["message"],
                                            })
                                        import pandas as _pd
                                        st.dataframe(
                                            _pd.DataFrame(warn_rows).astype(str),
                                            width="stretch", hide_index=True
                                        )
                                    else:
                                        st.success("✅ 沒有編號一致性問題")

                                    if consistency["unmatched_from"]:
                                        with st.expander(
                                            f"🔍 上游 WTA 編號 {len(consistency['unmatched_from'])} 條, 下游沒抽到對應 WTB"
                                        ):
                                            import pandas as _pd
                                            st.dataframe(
                                                _pd.DataFrame(consistency["unmatched_from"]).astype(str),
                                                width="stretch", hide_index=True
                                            )
                                            st.caption("可能原因: 圖中沒標下游 WTB 編號 / Gemini 漏抽 / 該箭頭實際進入外部")

                                    if consistency["unmatched_to"]:
                                        with st.expander(
                                            f"🔍 下游 WTB 編號 {len(consistency['unmatched_to'])} 條, 上游沒抽到對應 WTA"
                                        ):
                                            import pandas as _pd
                                            st.dataframe(
                                                _pd.DataFrame(consistency["unmatched_to"]).astype(str),
                                                width="stretch", hide_index=True
                                            )
                                            st.caption("可能原因: 來自外部原廢水 (WM) / Gemini 漏抽上游編號")

                                    with st.expander("📋 全部編號詳細位置 (debug)"):
                                        cg = consistency["code_groups"]
                                        cg_rows = []
                                        for code, info in cg.items():
                                            cg_rows.append({
                                                "編號": code,
                                                "as_from 次數": len(info["as_from"]),
                                                "as_to 次數": len(info["as_to"]),
                                                "Q 值集合": str(info["Q_set"]),
                                            })
                                        if cg_rows:
                                            import pandas as _pd
                                            st.dataframe(
                                                _pd.DataFrame(cg_rows).astype(str),
                                                width="stretch", hide_index=True
                                            )

                                with tabs[4]:
                                    # 跟現有 app_data 比對
                                    app_data_cmp = st.session_state.get("_app_data") or {}
                                    if not app_data_cmp:
                                        st.info("尚未做完整審查, 先按上方「開始完整審查」")
                                    else:
                                        sys_units = set(app_data_cmp.get("units", {}).keys())
                                        img_units = {u["code"] for u in fr["all_units"]}
                                        only_in_sys = sys_units - img_units
                                        only_in_img = img_units - sys_units
                                        both = sys_units & img_units

                                        cmp1, cmp2, cmp3 = st.columns(3)
                                        cmp1.metric("共同有", len(both))
                                        cmp2.metric("僅系統有", len(only_in_sys))
                                        cmp3.metric("僅圖中有", len(only_in_img))

                                        if only_in_img:
                                            st.warning(
                                                f"圖上有但系統沒抽到的單元: {', '.join(sorted(only_in_img))}"
                                            )
                                        if only_in_sys:
                                            st.caption(
                                                f"系統有但圖上沒看到 (可能是污泥/反洗等支線): "
                                                f"{', '.join(sorted(only_in_sys))}"
                                            )

                                        # 比對進流數
                                        st.markdown("**進流數比對 (圖中 vs 系統抽取)**")
                                        cmp_rows = []
                                        in_cnt_img = {}
                                        for flow in fr["all_flows"]:
                                            to_u = flow.get("to_unit")
                                            if to_u:
                                                in_cnt_img[to_u] = in_cnt_img.get(to_u, 0) + 1
                                        for ext in fr["all_external_inputs"]:
                                            to_u = ext.get("to_unit")
                                            if to_u:
                                                in_cnt_img[to_u] = in_cnt_img.get(to_u, 0) + 1
                                        for code in sorted(both):
                                            sys_in = len(app_data_cmp["units"][code].get("influent", {}))
                                            img_in = in_cnt_img.get(code, 0)
                                            cmp_rows.append({
                                                "單元": code,
                                                "圖中進流數": img_in,
                                                "系統抽取進流數": sys_in,
                                                "差異": img_in - sys_in,
                                                "備註": "✅" if img_in == sys_in else "⚠️ 不一致",
                                            })
                                        if cmp_rows:
                                            import pandas as _pd
                                            st.dataframe(
                                                _pd.DataFrame(cmp_rows).astype(str),
                                                width="stretch", hide_index=True
                                            )

                                with tabs[5]:
                                    for pr in fr.get("per_page_results", []):
                                        st.markdown(f"**p{pr['page']} - {pr.get('system_title', '')}**")
                                        st.json(pr["data"])
                    else:
                        st.caption("此 PDF 沒找到水量平衡示意圖純圖頁")

    # ── 補充: 圖片局部判讀 (Gemini Vision) ──
    with st.expander("📷 補充: 上傳圖片做局部判讀 (Gemini Vision)", expanded=False):
        st.caption(
            "用途: 上傳申請文件的截圖 (流向圖 / 水量平衡 / 數據表 / 設計尺寸) 讓 Gemini 判讀。"
            " 適用於 PDF 抽取不夠完整、想針對特定區塊深入分析的時候。"
        )
        try:
            import gemini_vision as _gv
            _gv_ok = True
        except Exception as _e:
            st.error(f"無法載入 gemini_vision: {_e}")
            _gv_ok = False

        if _gv_ok:
            # 認證狀態
            try:
                import gemini_extractor as _ge
                _gs = _ge.check_gemini_status()
            except Exception:
                _gs = {"ok": False, "message": "?"}

            if not _gs.get("ok"):
                st.warning(f"⚠️ {_gs.get('message')} (請見『匯入新規則』分頁設定)")
            else:
                imgs = st.file_uploader(
                    "拖入 1~N 張圖片 (PNG / JPG / WEBP)",
                    type=["png", "jpg", "jpeg", "webp"],
                    accept_multiple_files=True,
                    key="_review_img_uploader",
                )
                focus_hint = st.text_input(
                    "提示 (可選, 告訴 Gemini 該注意什麼)",
                    placeholder="例: 請特別注意流向是否標示反洗水來源",
                    key="_review_img_hint",
                )

                if imgs:
                    cols_thumb = st.columns(min(len(imgs), 4))
                    for i, img in enumerate(imgs):
                        with cols_thumb[i % 4]:
                            st.image(img, caption=img.name, width=150)

                    if st.button("🔍 開始判讀", key="_btn_interpret_images",
                                 type="primary"):
                        with st.spinner(f"Gemini Vision 判讀 {len(imgs)} 張圖片…"):
                            img_bytes_list = [img.getvalue() for img in imgs]
                            _gr = _gv.process_images(
                                img_bytes_list,
                                mode="interpret_diagram",
                                focus_hint=focus_hint,
                            )
                        if _gr.get("ok"):
                            result = _gr.get("result", {})
                            st.success(
                                f"✅ 判讀完成 ({_gr['image_count']} 張) · "
                                f"信心度 {result.get('confidence', '?')} · "
                                f"tokens in {_gr['gemini_usage'].get('input_tokens')} "
                                f"out {_gr['gemini_usage'].get('output_tokens')}"
                            )
                            st.markdown(f"**圖表類型**: {result.get('diagram_type', '(未判斷)')}")

                            if result.get("observations"):
                                st.markdown("**🔍 關鍵發現**")
                                for obs in result["observations"]:
                                    st.markdown(f"- {obs}")

                            if result.get("concerns"):
                                st.markdown("**⚠️ 可能不合理之處**")
                                for c in result["concerns"]:
                                    st.markdown(f"- {c}")

                            if result.get("units"):
                                st.markdown(f"**📦 槽體資料 ({len(result['units'])} 個)**")
                                import pandas as _pd
                                df_u = _pd.DataFrame(result["units"]).astype(str)
                                st.dataframe(df_u, width="stretch", hide_index=True)

                            if result.get("flows"):
                                st.markdown(f"**🌊 流向資料 ({len(result['flows'])} 條)**")
                                import pandas as _pd
                                df_f = _pd.DataFrame(result["flows"]).astype(str)
                                st.dataframe(df_f, width="stretch", hide_index=True)

                            with st.expander("Raw text (圖片裡看得到的所有文字)"):
                                st.text(result.get("raw_text", ""))
                        else:
                            st.error(f"❌ 失敗 ({_gr.get('stage')}): {_gr.get('error')}")

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
        st.dataframe(section_rows, width="stretch", hide_index=True)

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

        # 單元清單 (精簡 + 加削減率)
        st.subheader("📋 處理單元清單")
        st.caption(
            "**進/出** = 進出流水流股數　·　"
            "**削減率** = 質量平衡: Σ(進流質量) → Σ(出流質量), 跨各水質項目平均 "
            "(排除 pH/水溫)。⚠️ = 有項目異常 (出 > 進 或 > 99.5%)"
        )

        # 先算所有單元的削減率
        try:
            import unit_removal_rate as _urr
            removal_results = _urr.compute_all_units_removal(app_data)
        except Exception as _re:
            st.caption(f"削減率計算失敗: {_re}")
            removal_results = {}
            _urr = None

        unit_rows = []
        for code, info in sorted(app_data["units"].items()):
            rr = removal_results.get(code, {})
            removal_text = (_urr.format_removal_short(rr.get("summary", {}))
                            if _urr and removal_results else "-")
            unit_rows.append({
                "代號": code,
                "原始名稱": info["name_in_doc"],
                "標準類型": info["std_tank"],
                "頁數": ", ".join(map(str, info["pages_found"][:3])),
                "📥進": len(info["influent"]),
                "📤出": len(info["effluent"]),
                "📉削減率": removal_text,
            })
        st.dataframe(unit_rows, width="stretch", hide_index=True)

        # 削減率異常清單
        if _urr and removal_results:
            warned_units = [
                (code, rr["summary"]) for code, rr in removal_results.items()
                if rr.get("summary", {}).get("warnings")
            ]
            if warned_units:
                with st.expander(f"⚠️ {len(warned_units)} 個單元有削減率異常項目"):
                    for code, summ in warned_units[:30]:
                        unit_info = app_data["units"].get(code, {})
                        unit_name = unit_info.get("name_in_doc", "")
                        std_name = unit_info.get("std_tank", "")
                        st.markdown(
                            f"**{code} {unit_name}** ({std_name}) — "
                            f"平均削減率 {summ.get('avg_removal_pct', '?')}%"
                        )
                        for w in summ["warnings"][:5]:
                            st.caption(f"　• {w}")
                        if len(summ["warnings"]) > 5:
                            st.caption(f"　… 還有 {len(summ['warnings']) - 5} 項異常")

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
                            st.dataframe(up_rows, width="stretch", hide_index=True)
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
                            st.dataframe(dn_rows, width="stretch", hide_index=True)
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
                st.dataframe(eq_rows, width="stretch", hide_index=True)

            if unit.get("influent"):
                st.markdown(f"**進流水質** ({len(unit['influent'])} 流向)")
                for infl_code, qdata in unit["influent"].items():
                    st.caption(f"📥 {infl_code}")
                    q_rows = [{"水質項目": str(k),
                               "濃度": str(v.get("濃度", v.get("範圍", ""))),
                               "質量": str(v.get("質量", ""))} for k, v in qdata.items()]
                    st.dataframe(q_rows, width="stretch", hide_index=True)

            if unit.get("effluent"):
                st.markdown(f"**出流水質** ({len(unit['effluent'])} 流向)")
                for effl_code, qdata in unit["effluent"].items():
                    st.caption(f"📤 {effl_code}")
                    q_rows = [{"水質項目": str(k),
                               "濃度": str(v.get("濃度", v.get("範圍", ""))),
                               "質量": str(v.get("質量", ""))} for k, v in qdata.items()]
                    st.dataframe(q_rows, width="stretch", hide_index=True)

            # 削減率分項表 (本單元的所有水質項目)
            if _urr:
                unit_rr = removal_results.get(selected, {})
                items = unit_rr.get("items", {})
                summ = unit_rr.get("summary", {})
                if items:
                    st.markdown("**📉 削減率分項 (Σ進流質量 → Σ出流質量)**")
                    rr_rows = []
                    for item, r in sorted(items.items()):
                        pct = r["removal_pct"]
                        flag = ""
                        if pct < -10:
                            flag = "⚠️ 出流>進流"
                        elif pct > 99.5 and not any(k in item for k in ("懸浮", "SS")):
                            flag = "⚠️ 過高"
                        rr_rows.append({
                            "水質項目": item,
                            "進流總質量 (g/d)": r["in_mass"],
                            "出流總質量 (g/d)": r["out_mass"],
                            "削減率 (%)": pct,
                            "備註": flag,
                        })
                    st.dataframe(rr_rows, width="stretch", hide_index=True)
                    cc1, cc2, cc3, cc4 = st.columns(4)
                    cc1.metric("涵蓋項目", summ["items_count"])
                    cc2.metric("平均削減率", f"{summ['avg_removal_pct']}%")
                    cc3.metric("最低", f"{summ['min_removal_pct']}%")
                    cc4.metric("最高", f"{summ['max_removal_pct']}%")
                else:
                    st.caption("(此單元無進+出兩邊都有資料的水質項目, 無法計算削減率)")

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
                width="stretch",
            )
        with col2:
            json_str = json.dumps(app_data, ensure_ascii=False, indent=2)
            st.download_button(
                "💾 下載 JSON 結構化資料",
                data=json_str.encode("utf-8"),
                file_name=f"{base_name}_抽取結果.json",
                mime="application/json",
                width="stretch",
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
                st.dataframe(rows, width="stretch", hide_index=True)

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
                st.dataframe(flow_rows, width="stretch", hide_index=True)

            if ocr_result.get("all_doses"):
                st.markdown("**識別到的加藥量:**")
                dose_rows = [
                    {"化學品": str(d["chemical"]), "用量": str(d["amount"]),
                     "單位": str(d.get("unit", "")), "OCR 原文": str(d["text"])}
                    for d in ocr_result["all_doses"]
                ]
                st.dataframe(dose_rows, width="stretch", hide_index=True)

            if ocr_result.get("all_moistures"):
                st.markdown("**識別到的含水率:**")
                mois_rows = [
                    {"含水率 (%)": str(m["value_pct"]), "OCR 原文": str(m["text"])}
                    for m in ocr_result["all_moistures"]
                ]
                st.dataframe(mois_rows, width="stretch", hide_index=True)

            if ocr_result.get("all_units"):
                with st.expander(f"OCR 識別到的單元代號 ({len(ocr_result['all_units'])} 個)"):
                    unit_rows = [
                        {"代號": str(u["code"]), "OCR 原文": str(u["text"])}
                        for u in ocr_result["all_units"]
                    ]
                    st.dataframe(unit_rows, width="stretch", hide_index=True)

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
            width="stretch", hide_index=True,
        )

    # ── 跨 session 永久審查紀錄 (從 Sheet 讀) ──
    st.divider()
    st.subheader("📋 歷次審查紀錄 (跨 session 永久)")
    st.caption("從 Google Sheet 的 `_審查紀錄` 分頁讀取。每次「開始完整審查」完成會自動 append 一列。")

    col_h1, col_h2 = st.columns([1, 4])
    with col_h1:
        refresh_history = st.button("🔄 重新載入", key="_refresh_history_btn")
    if refresh_history or "_review_history_sheet" not in st.session_state:
        try:
            import review_history as _rh
            with st.spinner("從 Sheet 載入歷史…"):
                result = _rh.load_review_history(limit=200)
            st.session_state["_review_history_sheet"] = result
        except Exception as e:
            st.session_state["_review_history_sheet"] = {"ok": False, "error": str(e)}

    sheet_history = st.session_state.get("_review_history_sheet", {})
    if sheet_history.get("ok"):
        rows = sheet_history.get("rows", [])
        if rows:
            with col_h2:
                st.caption(f"共 {len(rows)} 筆 (最新在最上面)")
            try:
                import pandas as _pd
                df_h = _pd.DataFrame(rows)
                # 把所有欄位轉成字串 (避免 PyArrow 錯誤)
                df_h = df_h.astype(str)
                st.dataframe(df_h, width="stretch", hide_index=True)
            except Exception as _e:
                st.error(f"無法顯示表格: {_e}")
                st.json(rows[:10])
        else:
            st.info("尚無歷史紀錄。執行一次「🚀 開始審查」後會自動寫進來。")
    else:
        err = sheet_history.get("error", "未知錯誤")
        st.warning(f"無法載入: {err[:200]}")
        st.caption("可能原因: 未設定 Service Account / Sheet 沒分享 / 網路問題")

# ─────────────────────────────────────────────
# 規則庫管理 (Google Sheets 同步) — v2: xlsx 為主, csv 自動衍生
# ─────────────────────────────────────────────
with tab_sync:
    st.subheader("🔄 規則庫 ↔ Google Sheets 同步")
    st.markdown(
        "**規則庫.xlsx** 是主檔 (跟同事編輯的那份一樣)。\n"
        "**Google Sheet** 是同事的線上編輯場 (結構同 xlsx, 每槽體一分頁)。\n"
        "**rules_extracted.csv** 是「衍生物」, 自動由 xlsx 產出, 給審查程式跑。"
    )

    # 載入同步模組
    try:
        import sheets_sync
        sync_ok = True
    except Exception as e:
        st.error(f"無法載入 sheets_sync 模組: {e}")
        st.info("請先 `pip install gspread google-auth` (或檢查 requirements.txt)")
        sync_ok = False

    if sync_ok:
        # ── 認證狀態 ──
        status = sheets_sync.check_auth_status()
        if status["ok"]:
            st.success(f"✅ {status['message']}")
            st.caption(f"Service Account: `{status['email']}`")
        else:
            st.warning(f"⚠️ {status['message']}")
            st.info(
                "請先按照 **SHEETS_SETUP.md** 完成 Service Account 申請設定。"
            )

        st.divider()

        # ── Sheet 連結 ──
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheets_sync.DEFAULT_SHEET_ID}/edit"
        col_link_a, col_link_b = st.columns([3, 2])
        with col_link_a:
            st.markdown(f"**📄 Google Sheet** (給同事編): [打開 Sheet]({sheet_url})")
            st.caption(f"Sheet ID: `{sheets_sync.DEFAULT_SHEET_ID}`")
        with col_link_b:
            xlsx_exists = os.path.exists(sheets_sync.RULES_XLSX)
            csv_exists = os.path.exists(sheets_sync.RULES_CSV)
            st.markdown("**📁 本機檔案狀態**")
            st.caption(f"規則庫.xlsx: {'✅ 存在' if xlsx_exists else '❌ 不存在'}")
            st.caption(f"rules_extracted.csv: {'✅ 存在' if csv_exists else '❌ 不存在'}")

        st.divider()

        # ── 四個動作按鈕 ──
        col_a, col_b, col_c, col_d = st.columns(4)

        with col_a:
            st.markdown("#### ⬆️ 上傳")
            st.caption("xlsx → Sheet (清空後整批寫入,自動加狀態下拉選單)")
            if st.button("上傳 xlsx → Sheet", type="primary",
                         disabled=not status["ok"] or st.session_state.get("_busy", False),
                         width="stretch", key="upload_btn"):
                st.session_state["_busy"] = True
                with st.spinner("上傳中… (29 個分頁)"):
                    r = sheets_sync.upload_xlsx_to_sheets()
                st.session_state["_busy"] = False
                if r.get("ok"):
                    st.success(
                        f"✅ 上傳成功: {len(r['sheets_written'])} 個分頁 / "
                        f"{r['total_data_rows']} 筆資料"
                    )
                    if r.get("sheets_removed"):
                        st.caption(f"清掉舊分頁: {', '.join(r['sheets_removed'])}")
                    st.caption(f"狀態欄下拉選單: 已設 {r.get('validations_added', 0)} 個分頁")
                    st.caption(f"時間: {r['timestamp']}")
                else:
                    st.error(f"❌ 失敗: {r.get('error', '?')}")

        with col_b:
            st.markdown("#### 🔍 預覽差異")
            st.caption("比對 Sheet vs xlsx 哪些規則有差異 (依槽體+缺失ID)")
            if st.button("預覽差異",
                         disabled=not status["ok"] or st.session_state.get("_busy", False),
                         width="stretch", key="preview_btn"):
                st.session_state["_busy"] = True
                with st.spinner("比對中…"):
                    r = sheets_sync.preview_diff()
                st.session_state["_busy"] = False
                if r.get("ok"):
                    st.success(
                        f"Sheet {r['sheet_total']} 筆 vs xlsx {r['xlsx_total']} 筆 / "
                        f"新增 {r['total_added']} / 刪除 {r['total_removed']} / "
                        f"異動 {r['total_changed']}"
                    )
                    if r.get("by_tank"):
                        with st.expander(f"差異明細 ({len(r['by_tank'])} 個槽體有變化)"):
                            for tank, d in r["by_tank"].items():
                                line = f"**{tank}**: "
                                parts = []
                                if d["added"]:
                                    parts.append(f"➕ {len(d['added'])} 筆 ({', '.join(d['added'][:3])}…)")
                                if d["removed"]:
                                    parts.append(f"➖ {len(d['removed'])} 筆 ({', '.join(d['removed'][:3])}…)")
                                if d["changed"]:
                                    parts.append(f"✏️ {len(d['changed'])} 筆 ({', '.join(d['changed'][:3])}…)")
                                st.markdown(line + " / ".join(parts))
                    else:
                        st.info("Sheet 跟 xlsx 完全一致 ✅")
                else:
                    st.error(f"❌ 失敗: {r.get('error', '?')}")

        with col_c:
            st.markdown("#### ⬇️ 下載")
            st.caption("Sheet → xlsx (自動備份舊版 + 自動產 csv)")
            confirm_dl = st.checkbox("我確定要覆寫 xlsx", key="confirm_download")
            if st.button("下載 Sheet → xlsx", type="secondary",
                         disabled=not (status["ok"] and confirm_dl) or st.session_state.get("_busy", False),
                         width="stretch", key="download_btn"):
                st.session_state["_busy"] = True
                with st.spinner("下載 + 備份中…"):
                    r = sheets_sync.download_sheets_to_xlsx()
                st.session_state["_busy"] = False
                if r.get("ok"):
                    st.success(
                        f"✅ 下載成功: {r['sheets_read']} 個分頁 / "
                        f"{r['total_data_rows']} 筆資料"
                    )
                    if r.get("backup"):
                        st.caption(f"已備份: `{os.path.basename(r['backup'].get('xlsx',''))}`")
                    if r.get("csv_export", {}).get("ok"):
                        st.caption(f"已產出 csv: {r['csv_export']['rows_written']} 筆")
                    st.warning("⚠️ 記得 git push xlsx + csv 上 GitHub, 線上版才會生效")
                else:
                    st.error(f"❌ 失敗: {r.get('error', '?')}")

        with col_d:
            st.markdown("#### 🔄 重產 csv")
            st.caption("從 xlsx 重新產出 csv (不動 Sheet)")
            if st.button("重產 csv", width="stretch", key="export_btn"):
                with st.spinner("產出中…"):
                    r = sheets_sync.export_xlsx_to_csv()
                if r.get("ok"):
                    st.success(f"✅ 完成: {r['rows_written']} 筆規則寫入 csv")
                else:
                    st.error(f"❌ 失敗: {r.get('error', '?')}")

        st.divider()

        # ── 狀態欄統計 ──
        try:
            import step3e_rule_driven_check as _s3e
            rules = _s3e.load_rules_by_tank()
            # 排除 _ 開頭的 key
            tank_rules = {k: v for k, v in rules.items() if not k.startswith("_")}
            total = sum(len(v) for v in tank_rules.values())
            skipped = _s3e.get_last_skipped_count()
            st.markdown("#### 📊 規則狀態統計")
            cols = st.columns(3)
            cols[0].metric("可用規則 (V/空白)", total)
            cols[1].metric("跳過 (狀態=?)", skipped)
            cols[2].metric("規則庫總數", total + skipped)
        except Exception as e:
            st.caption(f"無法載入規則統計: {e}")

        st.divider()

        # ── 備份清單 ──
        st.markdown("#### 📦 本機備份歷史")
        st.caption("每次「下載」前自動備份 xlsx + csv (timestamped)")
        backups = sheets_sync.list_backups()
        if backups:
            import pandas as _pd
            df_bk = _pd.DataFrame(backups)
            df_bk = df_bk[["mtime", "name", "size_kb"]]
            df_bk.columns = ["時間", "檔名", "大小 KB"]
            st.dataframe(df_bk, width="stretch", hide_index=True)
            st.caption(f"備份目錄: `{sheets_sync.BACKUP_DIR}` (已加 .gitignore, 不會推到 GitHub)")
        else:
            st.info("尚無備份。第一次「下載 Sheet → CSV」後會自動產生。")

# ─────────────────────────────────────────────
# 📥 匯入新規則 (半自動,從 NotebookLM 結果貼上)
# ─────────────────────────────────────────────
with tab_import:
    st.subheader("📥 匯入新審查意見規則")
    st.markdown(
        "**半自動流程**: 你 (或同事) 用 NotebookLM 把審查意見 PDF 抽出結構化規則 → 複製貼上 → "
        "系統解析 + 預覽 + 衝突檢查 → 確認後自動寫入主檔。"
    )

    try:
        import rule_importer
        importer_ok = True
    except Exception as e:
        st.error(f"無法載入 rule_importer: {e}")
        importer_ok = False

    if importer_ok:
        from datetime import date as _date

        # 預先讀現有 _來源清單 算下個 S 編號
        try:
            _state = rule_importer._get_existing_state()
            next_s = f"S{_state['max_s_num'] + 1:02d}"
            next_d = f"D{_state['max_d_num'] + 1:03d}"
        except Exception:
            next_s = "S?"
            next_d = "D?"

        st.info(f"📌 系統會自動分配: 來源代號 = `{next_s}` / 第一筆缺失ID = `{next_d}` 起跳")

        # ── Step 1: 提供資料 (兩種方式) ──
        st.markdown("### Step 1 — 提供規則資料")

        input_mode = st.radio(
            "選擇輸入方式",
            [
                "🤖 上傳 PDF (Gemini 自動抽)",
                "📷 上傳圖片 (Gemini Vision 抽規則)",
                "📎 上傳檔案 (CSV / xlsx)",
                "📝 貼上文字",
            ],
            key="imp_input_mode",
            horizontal=True,
        )

        parse_result = None
        auto_filename = ""
        auto_technicians = ""
        is_gemini_result = False  # 標記是否來自 Gemini (給後面複核 UI 用)

        if input_mode.startswith("🤖"):
            # Gemini PDF 模式
            try:
                import gemini_extractor
                gemini_ok = True
            except Exception as e:
                st.error(f"無法載入 gemini_extractor: {e}")
                gemini_ok = False

            if gemini_ok:
                g_status = gemini_extractor.check_gemini_status()
                if g_status["ok"]:
                    st.success(
                        f"✅ Gemini {g_status['message']} "
                        f"(key: `{g_status['key_preview']}`)"
                    )
                else:
                    st.warning(f"⚠️ {g_status['message']}")
                    st.info(
                        "請在 Streamlit Cloud Secrets 加上一行 "
                        "(必須在 `[gcp_service_account]` block **之後** 留空一行再加):\n\n"
                        "```toml\n"
                        "[gcp_service_account]\n"
                        "...\n"
                        "universe_domain = \"googleapis.com\"\n"
                        "\n"
                        "gemini_api_key = \"AQ.xxxxx\"\n"
                        "```\n\n"
                        "貼完按 Save → 線上版會自動重啟 (或右上「⋮ → Reboot app」強制)"
                    )
                    # debug: 顯示 st.secrets 看得到的 top-level key 名稱 (不顯示值)
                    available = g_status.get("available_keys", [])
                    if available:
                        st.caption(
                            f"🔍 **Debug** — 系統在 `st.secrets` 看到的 key 名稱: "
                            f"`{available}`"
                        )
                        if "gemini_api_key" in available:
                            st.error(
                                "⚠️ key 看得到, 但讀取失敗 (可能是程式 bug, 截圖給開發者)"
                            )
                        elif "gcp_service_account" in available and len(available) == 1:
                            st.warning(
                                "⚠️ 只看到 `gcp_service_account`, 沒有 `gemini_api_key`. "
                                "可能 `gemini_api_key` 跑到 block **裡面** 了, "
                                "請在 `universe_domain = \"...\"` 那行**下一行**留空, "
                                "再下一行才寫 `gemini_api_key = \"...\"`"
                            )
                        else:
                            st.info(
                                "可能 Secrets 還沒生效, 試試右上 ⋮ → Reboot app"
                            )

                uploaded_pdf = st.file_uploader(
                    "拖 PDF 到這裡 (審查意見書)",
                    type=["pdf"],
                    key="imp_pdf_uploader",
                    disabled=not g_status["ok"],
                    help="Gemini 會自動讀整份 PDF, 抽出每筆缺失成結構化規則"
                )

                if "gemini_extract_result" not in st.session_state:
                    st.session_state["gemini_extract_result"] = None
                    st.session_state["gemini_pdf_key"] = ""

                if uploaded_pdf:
                    pdf_key = f"{uploaded_pdf.name}_{uploaded_pdf.size}"
                    last_key = st.session_state.get("gemini_pdf_key", "")

                    col_btn_e, col_st = st.columns([1, 3])
                    with col_btn_e:
                        re_extract = st.button(
                            "🔄 重抽" if last_key == pdf_key else "🤖 開始抽取",
                            type="primary",
                            disabled=not g_status["ok"] or st.session_state.get("_busy", False),
                            width="stretch",
                            key="imp_gemini_btn",
                        )
                    with col_st:
                        if last_key == pdf_key and st.session_state["gemini_extract_result"]:
                            st.caption(
                                f"✅ 已抽過: {uploaded_pdf.name} (按重抽會再呼叫 Gemini, 會花費 token)"
                            )
                        else:
                            st.caption(f"準備抽: {uploaded_pdf.name} ({uploaded_pdf.size/1024:.0f} KB)")

                    if re_extract:
                        st.session_state["_busy"] = True
                        with st.spinner("📄 讀 PDF → 🤖 呼叫 Gemini → 解析… (10-30 秒)"):
                            ex_result = gemini_extractor.extract_rules_from_pdf(
                                uploaded_pdf.getvalue(), uploaded_pdf.name
                            )
                        st.session_state["gemini_extract_result"] = ex_result
                        st.session_state["gemini_pdf_key"] = pdf_key
                        st.session_state["_busy"] = False

                    ex_result = st.session_state.get("gemini_extract_result")
                    if ex_result and st.session_state.get("gemini_pdf_key") == pdf_key:
                        if not ex_result.get("ok"):
                            st.error(
                                f"❌ 失敗 ({ex_result.get('stage')}): {ex_result.get('error')}"
                            )
                            if ex_result.get("raw_response"):
                                with st.expander("Gemini raw response (debug)"):
                                    st.code(ex_result["raw_response"])
                        else:
                            cd = ex_result.get("confidence_dist", {})
                            usage = ex_result.get("gemini_usage", {})
                            st.success(
                                f"✅ Gemini 抽出 **{ex_result['row_count']}** 筆 "
                                f"(高信心 {cd.get('high', 0)} / "
                                f"中 {cd.get('medium', 0)} / "
                                f"低 {cd.get('low', 0)})"
                            )
                            st.caption(
                                f"PDF: {ex_result['pdf_pages']} 頁 {ex_result['pdf_chars']} 字 / "
                                f"Gemini tokens: in {usage.get('input_tokens', '?')} "
                                f"out {usage.get('output_tokens', '?')}"
                            )

                            parse_result = {
                                "ok": True,
                                "rows": ex_result["rows"],
                                "format": "gemini-pdf",
                                "row_count": ex_result["row_count"],
                                "filename": uploaded_pdf.name,
                            }
                            auto_filename = uploaded_pdf.name.rsplit(".", 1)[0]
                            is_gemini_result = True
        elif input_mode.startswith("📷"):
            # Gemini Vision 圖片模式: 把審查意見書的截圖 → 結構化規則
            try:
                import gemini_vision
                import gemini_extractor
                vision_ok = True
            except Exception as _ve:
                st.error(f"無法載入 gemini_vision: {_ve}")
                vision_ok = False

            if vision_ok:
                _vstat = gemini_extractor.check_gemini_status()
                if _vstat["ok"]:
                    st.success(
                        f"✅ Gemini Vision 已啟用 (key: `{_vstat['key_preview']}`)"
                    )
                else:
                    st.warning(f"⚠️ {_vstat['message']}")

                uploaded_imgs = st.file_uploader(
                    "拖入 1~N 張審查意見書截圖 (PNG / JPG / WEBP)",
                    type=["png", "jpg", "jpeg", "webp"],
                    accept_multiple_files=True,
                    key="imp_img_uploader",
                    disabled=not _vstat["ok"],
                    help="可一次上傳多張, Gemini 會綜合判斷"
                )

                if "vision_extract_result" not in st.session_state:
                    st.session_state["vision_extract_result"] = None
                    st.session_state["vision_img_key"] = ""

                if uploaded_imgs:
                    # session key 用「檔名 + size 列表」(換不同圖片組就重抽)
                    img_key = "|".join(f"{i.name}_{i.size}" for i in uploaded_imgs)
                    last_key = st.session_state.get("vision_img_key", "")

                    # 縮圖
                    cols_t = st.columns(min(len(uploaded_imgs), 4))
                    for i, img in enumerate(uploaded_imgs):
                        with cols_t[i % 4]:
                            st.image(img, caption=img.name, width=130)

                    btn_label = "🔄 重抽" if last_key == img_key else f"📷 開始抽取 ({len(uploaded_imgs)} 張)"
                    if st.button(btn_label, type="primary",
                                 disabled=not _vstat["ok"], width="stretch",
                                 key="imp_vision_btn"):
                        with st.spinner(f"Gemini Vision 處理 {len(uploaded_imgs)} 張…"):
                            img_bytes_list = [img.getvalue() for img in uploaded_imgs]
                            _vr = gemini_vision.process_images(
                                img_bytes_list, mode="extract_rules"
                            )
                        st.session_state["vision_extract_result"] = _vr
                        st.session_state["vision_img_key"] = img_key

                    ex_result = st.session_state.get("vision_extract_result")
                    if ex_result and st.session_state.get("vision_img_key") == img_key:
                        if not ex_result.get("ok"):
                            st.error(
                                f"❌ 失敗 ({ex_result.get('stage')}): {ex_result.get('error')}"
                            )
                            if ex_result.get("raw_response"):
                                with st.expander("Gemini raw response (debug)"):
                                    st.code(ex_result["raw_response"])
                        else:
                            cd = ex_result.get("confidence_dist", {})
                            usage = ex_result.get("gemini_usage", {})
                            st.success(
                                f"✅ Vision 抽出 **{ex_result['row_count']}** 筆 "
                                f"(高 {cd.get('high', 0)} / 中 {cd.get('medium', 0)} / 低 {cd.get('low', 0)})"
                            )
                            st.caption(
                                f"圖片: {ex_result['image_count']} 張 / "
                                f"tokens in {usage.get('input_tokens', '?')} "
                                f"out {usage.get('output_tokens', '?')}"
                            )

                            parse_result = {
                                "ok": True,
                                "rows": ex_result["rows"],
                                "format": "gemini-vision",
                                "row_count": ex_result["row_count"],
                            }
                            # 用第一張圖的名稱當預設檔名
                            first_img_name = uploaded_imgs[0].name
                            auto_filename = f"圖片_{first_img_name.rsplit('.', 1)[0]}"
                            is_gemini_result = True
        elif input_mode.startswith("📎"):
            uploaded = st.file_uploader(
                "拖檔案到這裡, 或點選檔案",
                type=["csv", "xlsx", "xls"],
                key="imp_uploader",
                help="CSV: 單一扁平表 / xlsx: 可單一分頁也可多分頁 (每分頁名稱 = 槽體名稱)"
            )
            if uploaded:
                with st.spinner(f"解析 {uploaded.name}…"):
                    parse_result = rule_importer.parse_uploaded_file(uploaded)
                auto_filename = uploaded.name.rsplit(".", 1)[0]

                if parse_result.get("ok") and parse_result.get("sheets"):
                    sheets_info = parse_result["sheets"]
                    if len(sheets_info) > 1:
                        st.caption(
                            f"📑 偵測到 {len(sheets_info)} 個分頁: "
                            + ", ".join(f"{s['name']}({s['rows']})" for s in sheets_info[:8])
                            + ("..." if len(sheets_info) > 8 else "")
                        )
        else:
            st.caption(
                "**支援**: TSV (Excel 複製貼上預設) / CSV / Markdown 表格。"
                " 必要欄位: `原文缺失` / `檢查類型` / `對照項目` / `規則` / `標準槽體名稱`。"
            )

            with st.expander("📋 範本格式 (點開複製)"):
                st.code("""缺失ID\t原文缺失\t檢查類型\t對照項目\t規則\t比對位置\t判定邏輯\t技師姓名\t序號\t標準槽體名稱\t原始槽體代號
\t出水池排放口缺裝液位計\t機具設施\t液位計\t放流池排放口應設液位計監測水位\t廢污水處理設施操作條件\t若 放流池 且 無液位計 → 標記:缺機具\t範例技師\t序1 範例技師 (1)\t放流池\tT01-15
""", language="text")

            pasted_text = st.text_area(
                "貼上你的規則表 (TSV / CSV / Markdown)",
                key="imp_pasted",
                height=200,
                placeholder="從 NotebookLM 或 Excel/Sheet 複製整個表格 (含表頭) 直接貼這裡",
            )
            if pasted_text.strip():
                parse_result = rule_importer.parse_input_text(pasted_text)

        # ── 後續步驟 (只有 parse_result OK 才往下) ──
        if parse_result:
            if not parse_result.get("ok"):
                st.error(f"❌ 解析失敗: {parse_result.get('error')}")
            else:
                rows = parse_result["rows"]

                # 自動推導: 從上傳資料抽出技師姓名 (去重)
                techs_in_data = sorted({
                    (r.get("技師姓名") or "").strip()
                    for r in rows
                    if (r.get("技師姓名") or "").strip()
                })
                if techs_in_data:
                    auto_technicians = " / ".join(techs_in_data[:5])
                    if len(techs_in_data) > 5:
                        auto_technicians += f" 等 {len(techs_in_data)} 位"

                st.success(
                    f"✅ 解析成功: {parse_result['row_count']} 筆 "
                    f"({parse_result['format'].upper()})"
                )

                # ── Gemini 結果 → 顯示複核 UI (data_editor) ──
                if is_gemini_result:
                    st.markdown("### Step 1.5 — 人工複核 (預設全勾選)")
                    st.caption(
                        "Gemini 抽出的結果可能有誤, 請逐筆確認:\n"
                        "- ☑️ **選取** 欄: 取消勾選不要匯入的\n"
                        "- 其他欄位都可直接點擊修改\n"
                        "- **信心度** = Gemini 自評, low 建議仔細看"
                    )

                    import pandas as _pd
                    # 把 rows 轉成 DataFrame, 加「選取」欄預設 True, 加「信心度」欄
                    df_review = _pd.DataFrame(rows).fillna("")
                    df_review.insert(0, "選取", True)
                    if "_confidence" in df_review.columns:
                        df_review.rename(columns={"_confidence": "信心度"}, inplace=True)
                    else:
                        df_review["信心度"] = "unknown"

                    # 欄位順序
                    ordered_cols = [
                        "選取", "信心度", "標準槽體名稱", "原始槽體代號",
                        "原文缺失", "檢查類型", "對照項目", "規則",
                        "比對位置", "判定邏輯", "技師姓名", "序號", "缺失ID", "狀態",
                    ]
                    ordered_cols = [c for c in ordered_cols if c in df_review.columns]
                    df_review = df_review[ordered_cols]

                    # 排序: 低信心度排前面 (讓使用者先注意)
                    conf_order = {"low": 0, "medium": 1, "high": 2, "unknown": 3}
                    df_review["_sort"] = df_review["信心度"].map(lambda x: conf_order.get(x, 4))
                    df_review = df_review.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

                    # 槽體下拉選單選項
                    try:
                        import gemini_extractor as _ge
                        tank_options = _ge.STANDARD_TANKS
                        check_options = _ge.CHECK_TYPES
                    except Exception:
                        tank_options = []
                        check_options = []

                    # data_editor 設定
                    column_config = {
                        "選取": st.column_config.CheckboxColumn(
                            "✓", help="勾選 = 匯入", default=True, width="small"
                        ),
                        "信心度": st.column_config.SelectboxColumn(
                            "信心",
                            options=["high", "medium", "low", "unknown"],
                            width="small",
                        ),
                        "標準槽體名稱": st.column_config.SelectboxColumn(
                            "標準槽體",
                            options=tank_options,
                            required=False,
                        ) if tank_options else st.column_config.TextColumn("標準槽體"),
                        "檢查類型": st.column_config.SelectboxColumn(
                            "檢查類型",
                            options=check_options,
                        ) if check_options else st.column_config.TextColumn("檢查類型"),
                        "原文缺失": st.column_config.TextColumn(
                            "原文缺失", width="large"
                        ),
                        "規則": st.column_config.TextColumn(
                            "規則", width="medium"
                        ),
                        "原始槽體代號": st.column_config.TextColumn("原代號", width="small"),
                        "對照項目": st.column_config.TextColumn("對照項目", width="small"),
                        "比對位置": st.column_config.TextColumn("比對位置"),
                        "判定邏輯": st.column_config.TextColumn("判定邏輯"),
                        "技師姓名": st.column_config.TextColumn("技師", width="small"),
                        "序號": st.column_config.TextColumn("序號", width="small"),
                        "缺失ID": st.column_config.TextColumn("缺失ID", width="small"),
                        "狀態": st.column_config.SelectboxColumn(
                            "狀態", options=["", "V", "?"], width="small",
                        ),
                    }

                    edited_df = st.data_editor(
                        df_review,
                        column_config=column_config,
                        hide_index=True,
                        width="stretch",
                        num_rows="fixed",  # 不讓使用者亂加列
                        key="imp_gemini_editor",
                    )

                    # 統計勾選
                    selected_count = int(edited_df["選取"].sum())
                    total_count = len(edited_df)
                    col_sel1, col_sel2, col_sel3 = st.columns(3)
                    col_sel1.metric("勾選筆數", selected_count)
                    col_sel2.metric("總筆數", total_count)
                    col_sel3.metric(
                        "取消勾選", total_count - selected_count,
                        delta=f"-{total_count - selected_count}" if selected_count < total_count else None,
                    )

                    # 把編輯+勾選後的結果回灌成 rows (給後續流程用)
                    selected_rows = edited_df[edited_df["選取"] == True].drop(columns=["選取"])
                    if "信心度" in selected_rows.columns:
                        selected_rows = selected_rows.drop(columns=["信心度"])
                    rows = selected_rows.to_dict(orient="records")

                    if not rows:
                        st.warning("⚠️ 沒有勾選任何規則, 請至少勾一筆")

                else:
                    # 非 Gemini 路徑: 顯示傳統前 5 筆預覽
                    with st.expander(f"📋 前 5 筆預覽 (共 {len(rows)} 筆)"):
                        import pandas as _pd
                        df_preview = _pd.DataFrame(rows).fillna("")
                        st.dataframe(df_preview.head(5), width="stretch", hide_index=True)

                # ── Step 2: 預覽匯入結果 ──
                if rows:
                    st.markdown("### Step 2 — 預覽匯入結果")
                    preview = rule_importer.preview_import(rows)

                    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
                    col_p1.metric("總筆數", preview["total"])
                    col_p2.metric("可匯入", preview["ok_to_import"])
                    col_p3.metric("涵蓋槽體", len(preview["tanks_in_import"]))
                    col_p4.metric("新槽體", len(preview["new_tanks"]))

                    if preview["new_tanks"]:
                        st.warning(
                            f"⚠️ 將新建 {len(preview['new_tanks'])} 個槽體分頁: "
                            f"{', '.join(preview['new_tanks'])}\n\n"
                            f"請確認名稱跟 RULE_AUTHORING.md 第 1 章的標準槽體一致。"
                        )

                    if preview["id_conflicts"]:
                        st.info(
                            f"ℹ️ {len(preview['id_conflicts'])} 筆缺失ID 跟現有重複, "
                            f"系統會自動改成 `{next_d}` 起的新編號: "
                            f"{', '.join(preview['id_conflicts'][:5])}"
                            f"{'...' if len(preview['id_conflicts']) > 5 else ''}"
                        )

                    if preview["missing_required"]:
                        with st.expander(f"⚠️ {len(preview['missing_required'])} 筆缺必填欄 (會被跳過)"):
                            for m in preview["missing_required"][:20]:
                                st.write(f"- 第 {m['row_idx']} 列: 缺 {', '.join(m['missing'])} (預覽: {m['row_preview']})")
                            if len(preview["missing_required"]) > 20:
                                st.caption(f"… 還有 {len(preview['missing_required']) - 20} 筆")

                    if preview["tanks_in_import"]:
                        with st.expander(f"📂 涵蓋槽體清單 ({len(preview['tanks_in_import'])} 個)"):
                            st.write(", ".join(preview["tanks_in_import"]))

                    # ── Step 3: 基本資料 (自動帶入, 可選修改) ──
                    with st.expander("📝 基本資料 (系統已自動帶入, 可選擇修改)", expanded=False):
                        col_m1, col_m2 = st.columns(2)
                        today_str = _date.today().strftime("%Y-%m-%d")
                        with col_m1:
                            src_filename = st.text_input(
                                "審查意見檔名 (自動)", value=auto_filename,
                                key="imp_filename",
                                help="預設使用上傳檔案的檔名 (扣副檔名)",
                            )
                            src_technician = st.text_input(
                                "技師姓名 (自動)", value=auto_technicians,
                                key="imp_technician",
                                help="從上傳資料的「技師姓名」欄自動統計",
                            )
                            src_cert = st.text_input("技師證書字號 (可選)", value="",
                                                      key="imp_cert", placeholder="留空或 (見原文)")
                        with col_m2:
                            src_date = st.text_input("查核日期 (預設今天)", value=today_str,
                                                      key="imp_date")
                            src_company = st.text_input("簽證事業名稱 (可選)", value="",
                                                         key="imp_company", placeholder="例: (多家) 或 公司名")
                            src_note = st.text_input("備註 (可選)", value="",
                                                      key="imp_note", placeholder="留空使用預設")

                    # 沒展開的話 src_xxx 變數不會被設, 在這邊補預設
                    if "imp_filename" not in st.session_state:
                        src_filename = auto_filename or "未命名審查意見"
                        src_technician = auto_technicians
                        src_cert = ""
                        src_date = _date.today().strftime("%Y-%m-%d")
                        src_company = ""
                        src_note = ""
                    else:
                        src_filename = st.session_state.get("imp_filename", auto_filename) or auto_filename or "未命名審查意見"
                        src_technician = st.session_state.get("imp_technician", auto_technicians)
                        src_cert = st.session_state.get("imp_cert", "")
                        src_date = st.session_state.get("imp_date", _date.today().strftime("%Y-%m-%d"))
                        src_company = st.session_state.get("imp_company", "")
                        src_note = st.session_state.get("imp_note", "")

                    # ── Step 4: 確認匯入 ──
                    st.markdown("### Step 3 — 確認匯入")
                    st.info(
                        f"即將匯入 **{preview['ok_to_import']}** 筆 → "
                        f"來源代號 **{preview['next_source_code']}** ({src_filename}) / "
                        f"從 **{next_d}** 開始分配 ID"
                    )

                    confirm_import = st.checkbox(
                        "我確定要寫入 規則庫.xlsx (會自動備份舊版)",
                        key="imp_confirm",
                    )

                    if st.button("📥 執行匯入", type="primary",
                                 disabled=not confirm_import or st.session_state.get("_busy", False),
                                 width="stretch"):
                        st.session_state["_busy"] = True
                        metadata = {
                            "檔名": src_filename,
                            "技師姓名": src_technician,
                            "技師證書字號": src_cert,
                            "查核日期": src_date,
                            "簽證事業名稱": src_company,
                            "備註": src_note or "Streamlit 半自動匯入",
                        }
                        with st.spinner("寫入中…"):
                            result = rule_importer.commit_import(rows, metadata, skip_missing=True)
                        st.session_state["_busy"] = False

                        if result.get("ok"):
                            st.success(
                                f"✅ 匯入成功! 來源 **{result['source_code']}** / "
                                f"匯入 **{result['imported_count']}** 筆 / "
                                f"跳過 **{result['skipped_count']}** 筆"
                            )
                            if result.get("new_tanks_created"):
                                st.info(f"🆕 新建分頁: {', '.join(result['new_tanks_created'])}")
                            if result.get("backup"):
                                st.caption(f"備份: `{os.path.basename(result['backup'])}`")
                            st.warning(
                                "⚠️ 下一步:\n"
                                "1. 切到「🔄 規則庫管理」分頁\n"
                                "2. 按「⬆️ 上傳 xlsx → Sheet」(讓同事看到新規則)\n"
                                "3. 本機 git push (推到 GitHub, 線上版生效)"
                            )
                        else:
                            st.error(f"❌ 失敗: {result.get('error', '?')}")

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
