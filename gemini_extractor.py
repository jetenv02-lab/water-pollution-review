# -*- coding: utf-8 -*-
"""Gemini 自動從審查意見 PDF 抽取規則。

工作流:
    1. pdfplumber 抽 PDF 全文
    2. 把全文 + 標準槽體清單 + JSON schema 送給 Gemini Flash
    3. Gemini 回傳結構化 JSON list of rules
    4. 每筆規則含 confidence (high/medium/low)
    5. 上層 (Streamlit) 顯示 data_editor 給使用者人工複核

認證:
    1. Streamlit secrets["gemini_api_key"]
    2. 環境變數 GEMINI_API_KEY (僅本機開發)
"""
import json
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))

# 標準槽體清單 (跟 RULE_AUTHORING.md 一致, 不要加 _說明 等 meta 分頁)
STANDARD_TANKS = [
    # 預處理 / 前處理
    "廢水收集池", "廢水調整池", "調勻池", "暫存槽", "中和池",
    # 化學處理
    "pH調整槽", "快混槽", "慢混池", "沉澱池",
    # 生物處理
    "曝氣槽", "氧化池",
    # 高級處理
    "活性碳吸附塔", "活性碳吸附裝置", "砂濾塔", "離子交換樹脂塔",
    # 污泥處理
    "濃縮槽", "污泥儲槽", "脫水機",
    # 排放
    "放流池", "貯留槽",
    # 批次
    "批次反應槽",
    # 其他 / 文件類
    "(文件類)", "(現場設備類)", "文件類", "現場設備類",
]

CHECK_TYPES = [
    "設計參數", "機具設施", "質量平衡", "操作條件",
    "流向示意圖", "水質標準", "文件一致性", "去除率", "其他",
]

GEMINI_MODEL = "gemini-2.0-flash-exp"  # Flash 便宜快


# ──────────────────────────────────────────────────
# 認證
# ──────────────────────────────────────────────────

def _get_gemini_api_key():
    """取得 Gemini API key。優先 Streamlit Secrets, 再 env var。"""
    try:
        import streamlit as st
        if "gemini_api_key" in st.secrets:
            return str(st.secrets["gemini_api_key"]), "streamlit"
    except Exception:
        pass

    key = os.environ.get("GEMINI_API_KEY")
    if key:
        return key, "env"

    return None, None


def check_gemini_status():
    """檢查 Gemini API 設定狀態。"""
    key, source = _get_gemini_api_key()
    if not key:
        return {
            "ok": False,
            "source": None,
            "message": "未設定 Gemini API key",
        }
    # 不顯示完整 key, 只顯示前後 4 字元
    masked = f"{key[:6]}...{key[-4:]}" if len(key) > 15 else "***"
    return {
        "ok": True,
        "source": source,
        "key_preview": masked,
        "message": f"已設定 (來源: {source})",
    }


# ──────────────────────────────────────────────────
# PDF 文字抽取
# ──────────────────────────────────────────────────

def _extract_pdf_text(pdf_bytes):
    """用 pdfplumber 抽出 PDF 全文 (含表格)。

    Returns:
        {"ok": True, "text": str, "pages": N}
        {"ok": False, "error": str}
    """
    try:
        import pdfplumber
        import io
    except ImportError as e:
        return {"ok": False, "error": f"缺少 pdfplumber: {e}"}

    try:
        parts = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page_count = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                parts.append(f"\n---PAGE {i}---\n{text}")
                # 補抓表格 (純文字抽不到的)
                try:
                    tables = page.extract_tables()
                    for t_idx, table in enumerate(tables):
                        parts.append(f"\n[表 {i}-{t_idx + 1}]\n")
                        for row in table:
                            parts.append("\t".join(str(c) if c is not None else "" for c in row))
                            parts.append("\n")
                except Exception:
                    pass
        return {"ok": True, "text": "".join(parts), "pages": page_count}
    except Exception as e:
        return {"ok": False, "error": f"PDF 解析失敗: {e}"}


# ──────────────────────────────────────────────────
# Gemini Prompt
# ──────────────────────────────────────────────────

EXTRACTION_PROMPT = """你是台灣環工技師, 專門審查「水污染防治措施」申請文件。

我會給你一份「審查意見書」的文字內容。請你抽取出每一筆技師指出的缺失,
整理成結構化資料。

【輸出格式】
回傳一個 JSON array, 每個元素代表一筆缺失:

{
  "技師姓名": "從原文抽出, 例 方天志",
  "序號": "技師標的序號, 例 序1 方天志技師 (1)",
  "原文缺失": "技師原文 (盡量逐字, 不要改寫). 例: (1)頁次:7/34, 廢(污)水產生與水污染防治措施流向示意圖, 未標示T01-05活性碳吸附裝置反洗水來源。",
  "檢查類型": "從下列選一個: %s",
  "對照項目": "簡短說在查什麼, 例: 反洗水來源 / pH / 攪拌轉速 / 液位計",
  "規則": "白話描述, 例: 流向示意圖需標示反洗水來源",
  "比對位置": "在申請文件的哪段查, 例: 廢(污)水產生與水污染防治措施流向示意圖",
  "判定邏輯": "什麼條件下標記什麼, 例: 若 設備具反洗功能 且 未標示來源 → 標記:未標示來源",
  "標準槽體名稱": "從下列選一個: %s. 若涉及多槽體, 各寫一筆. 若是文件層級的缺失, 用 (文件類) ; 若是現場機具設施類, 用 (現場設備類).",
  "原始槽體代號": "原文寫的槽體序號, 例: T01-05",
  "confidence": "你對這筆抽取的信心: high (清楚明確) / medium (有點不確定) / low (內容模糊)"
}

【重要規則】
1. 每筆缺失對應一個槽體。若原文涉及多個槽體 (如 T01-09、T01-10、T01-11), 拆成多筆。
2. 標準槽體名稱**必須**從給的清單裡選, 不要自創。
3. 「原文缺失」盡量保留原文, 包括序號 (例 (1) (2) ...) 跟頁次資訊。
4. 信心度判斷:
   - high: 明確指明槽體跟問題
   - medium: 槽體可推測但不百分百確定
   - low: 內容模糊或跨多個系統
5. 若一份審查意見有多位技師, 每筆要正確標出技師姓名。

【審查意見全文】
%s

【任務】
請抽取所有缺失, 回傳 JSON array (不要加 ```json 標記, 直接給 JSON):
""" % (
    " / ".join(CHECK_TYPES),
    " / ".join(STANDARD_TANKS),
    "%s",
)


# ──────────────────────────────────────────────────
# 呼叫 Gemini
# ──────────────────────────────────────────────────

def _call_gemini(pdf_text, api_key):
    """呼叫 Gemini API 抽取規則。"""
    try:
        import google.generativeai as genai
    except ImportError as e:
        return {"ok": False, "error": f"缺少 google-generativeai: {e}. pip install google-generativeai"}

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(GEMINI_MODEL)

        prompt = EXTRACTION_PROMPT % pdf_text

        response = model.generate_content(
            prompt,
            generation_config={
                "temperature": 0.1,  # 低溫度, 求穩定
                "max_output_tokens": 32768,
                "response_mime_type": "application/json",
            },
        )

        raw = response.text.strip()
        # 移除可能的 ``` 包裝
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        rules = json.loads(raw)
        if not isinstance(rules, list):
            return {"ok": False, "error": "Gemini 回傳格式不是 array",
                    "raw_response": raw[:500]}

        return {
            "ok": True,
            "rules": rules,
            "usage": {
                "input_tokens": response.usage_metadata.prompt_token_count if response.usage_metadata else None,
                "output_tokens": response.usage_metadata.candidates_token_count if response.usage_metadata else None,
            },
        }
    except json.JSONDecodeError as e:
        return {"ok": False, "error": f"Gemini 回傳的不是合法 JSON: {e}",
                "raw_response": raw[:500] if 'raw' in dir() else ""}
    except Exception as e:
        return {"ok": False, "error": f"Gemini 呼叫失敗: {type(e).__name__}: {e}"}


# ──────────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────────

def extract_rules_from_pdf(pdf_bytes, filename="(uploaded.pdf)"):
    """完整 PDF → 規則 流程。

    Returns:
        {
            "ok": True,
            "rows": [...],                    # 跟 rule_importer 同格式
            "confidence_dist": {"high": N, "medium": M, "low": K},
            "pdf_pages": N,
            "pdf_chars": N,
            "gemini_usage": {...},
            "filename": str,
        }
        {"ok": False, "error": str, "stage": str}
    """
    # Step 1: 認證
    api_key, source = _get_gemini_api_key()
    if not api_key:
        return {"ok": False, "stage": "auth",
                "error": "未設定 Gemini API key\n\n"
                         "請在 Streamlit Cloud Secrets 加: gemini_api_key = \"...\"\n"
                         "或設環境變數 GEMINI_API_KEY"}

    # Step 2: 抽 PDF 全文
    pdf_result = _extract_pdf_text(pdf_bytes)
    if not pdf_result.get("ok"):
        return {"ok": False, "stage": "pdf", "error": pdf_result.get("error")}

    pdf_text = pdf_result["text"]
    if not pdf_text.strip():
        return {"ok": False, "stage": "pdf",
                "error": "PDF 沒有可抽取的文字 (可能是純圖檔, 需要 OCR)"}

    # 限制長度 (Gemini Flash context window 是 1M, 但 prompt 太長會慢)
    if len(pdf_text) > 200000:
        pdf_text = pdf_text[:200000] + "\n\n[... 文字過長, 已截斷 ...]"

    # Step 3: 呼叫 Gemini
    gemini_result = _call_gemini(pdf_text, api_key)
    if not gemini_result.get("ok"):
        return {"ok": False, "stage": "gemini",
                "error": gemini_result.get("error"),
                "raw_response": gemini_result.get("raw_response", "")}

    rules = gemini_result["rules"]

    # Step 4: 後處理 — 對應到 rule_importer 期望的欄位
    rows = []
    confidence_dist = {"high": 0, "medium": 0, "low": 0, "unknown": 0}
    for r in rules:
        if not isinstance(r, dict):
            continue
        row = {
            "缺失ID": "",  # 留空, 系統會自動分配
            "原文缺失": r.get("原文缺失", "").strip(),
            "檢查類型": r.get("檢查類型", "").strip(),
            "對照項目": r.get("對照項目", "").strip(),
            "規則": r.get("規則", "").strip(),
            "比對位置": r.get("比對位置", "").strip(),
            "判定邏輯": r.get("判定邏輯", "").strip(),
            "技師姓名": r.get("技師姓名", "").strip(),
            "序號": r.get("序號", "").strip(),
            "標準槽體名稱": r.get("標準槽體名稱", "").strip(),
            "原始槽體代號": r.get("原始槽體代號", "").strip(),
            "狀態": "",
            "_confidence": r.get("confidence", "unknown").lower().strip(),
        }
        conf = row["_confidence"]
        if conf not in confidence_dist:
            conf = "unknown"
        confidence_dist[conf] += 1
        rows.append(row)

    return {
        "ok": True,
        "rows": rows,
        "row_count": len(rows),
        "confidence_dist": confidence_dist,
        "pdf_pages": pdf_result["pages"],
        "pdf_chars": len(pdf_text),
        "gemini_usage": gemini_result.get("usage", {}),
        "auth_source": source,
        "filename": filename,
    }


# ──────────────────────────────────────────────────
# CLI 測試
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    print("=== Gemini API 狀態 ===")
    status = check_gemini_status()
    print(status)

    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        if not os.path.exists(pdf_path):
            print(f"找不到 {pdf_path}")
            sys.exit(1)
        with open(pdf_path, "rb") as f:
            data = f.read()
        print(f"\n=== 抽取 {pdf_path} ({len(data)} bytes) ===")
        result = extract_rules_from_pdf(data, os.path.basename(pdf_path))
        if result.get("ok"):
            print(f"✓ 抽出 {result['row_count']} 筆")
            print(f"  信心度: {result['confidence_dist']}")
            print(f"  PDF: {result['pdf_pages']} 頁 / {result['pdf_chars']} 字")
            print(f"  Gemini tokens: {result['gemini_usage']}")
            print(f"\n前 2 筆:")
            for r in result["rows"][:2]:
                print(json.dumps(r, ensure_ascii=False, indent=2))
        else:
            print(f"✗ 失敗 ({result.get('stage')}): {result.get('error')}")
