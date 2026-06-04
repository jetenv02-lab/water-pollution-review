# -*- coding: utf-8 -*-
"""Gemini Vision — 圖片局部處理。

兩種模式:
    1. EXTRACT_RULES — 把審查意見書的截圖 → 結構化規則 (走 rule_importer 流程)
    2. INTERPRET_DIAGRAM — 把申請文件的截圖 → 文字判讀 (流向圖/水量平衡/數據表)

支援:
    - 單張或多張圖片 (多張時, Gemini 會綜合判斷)
    - PNG / JPG / WEBP

認證: 沿用 gemini_extractor._get_gemini_api_key()
"""
import io
import json
import os
import re

# 從 gemini_extractor 共用模型候選清單
try:
    from gemini_extractor import GEMINI_MODEL_CANDIDATES
except Exception:
    GEMINI_MODEL_CANDIDATES = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-flash-latest"]
GEMINI_MODEL = GEMINI_MODEL_CANDIDATES[0]


# ──────────────────────────────────────────────────
# Prompt 模板
# ──────────────────────────────────────────────────

def _build_extract_rules_prompt():
    """模式 1: 從審查意見書截圖抽規則 (跟 gemini_extractor 同邏輯, 但接受圖片)。"""
    from gemini_extractor import CHECK_TYPES, STANDARD_TANKS, _build_synonyms_hint
    synonyms_hint = _build_synonyms_hint()
    return f"""你是台灣環工技師, 專門審查「水污染防治措施」申請文件。

我會給你一張 (或多張) 「審查意見書」的截圖。請從圖片裡讀出每一筆技師指出的缺失,
整理成結構化資料。

【輸出格式】
回傳一個 JSON array, 每個元素代表一筆缺失:

{{
  "技師姓名": "從圖片抽出, 例 方天志",
  "序號": "技師標的序號, 例 序1 方天志技師 (1)",
  "原文缺失": "技師原文 (盡量逐字, 不要改寫)",
  "檢查類型": "從下列選一個: {" / ".join(CHECK_TYPES)}",
  "對照項目": "簡短說在查什麼, 例: 反洗水來源 / pH / 攪拌轉速 / 液位計. 請優先用標準詞",
  "規則": "白話描述",
  "比對位置": "在申請文件的哪段查",
  "判定邏輯": "什麼條件下標記什麼",
  "標準槽體名稱": "從下列選一個: {" / ".join(STANDARD_TANKS)}",
  "原始槽體代號": "原文寫的槽體序號, 例: T01-05",
  "confidence": "high / medium / low"
}}

【重要規則】
1. 多個槽體 → 拆成多筆
2. 標準槽體名稱必須從清單選, 不要自創
3. 若多張圖, 整合所有缺失一起抽
{synonyms_hint}

【任務】
請從附上的圖片中抽取所有缺失, 回傳 JSON array (不要加 ```json 標記):
"""


def _build_interpret_diagram_prompt(focus_hint=""):
    """模式 2: 判讀申請文件的圖表 (流向圖/水量平衡圖/數據表/設計尺寸)。"""
    extra = f"\n【特別注意】\n{focus_hint}\n" if focus_hint else ""
    return f"""你是台灣環工技師, 專門審查「水污染防治措施」。

我會給你一張 (或多張) 申請文件的截圖。請判讀內容, 整理成結構化資料。

可能的圖表類型:
- 廢(污)水水質水量平衡示意圖 — 標示各槽體進出水量、水質濃度
- 廢(污)水產生及水污染防治措施流向示意圖 — 標示槽體連接、流向、迴流/反洗
- 處理設施資料表 — 槽體尺寸、有效容量、設計參數
- 進出處理單元之水質資料 — 各項水質濃度
- 機具設施清單 — 設備、馬力、規格
- 加藥資料 — 藥劑名稱、加藥率

【輸出格式 — JSON】
{{
  "diagram_type": "你判斷的圖表類型",
  "units": [
    {{
      "code": "槽體序號 (T01-05)",
      "name": "槽體名稱 (活性碳吸附塔)",
      "dimensions": "尺寸 (例: 1.2L × 1.2W × 1.5H m)",
      "effective_volume": "有效容量 (例: 1.8 m³)",
      "design_params": ["設計參數1: 值", "設計參數2: 值"],
      "equipment": ["機具1", "機具2"]
    }}
  ],
  "flows": [
    {{
      "from": "T01-05",
      "to": "T01-06",
      "stream_code": "WTA01-05 → WTB01-06",
      "flow_rate": "Q = 30 m³/d",
      "concentration": {{"COD": "150 mg/L", "pH": "7.5"}}
    }}
  ],
  "observations": [
    "你看到的關鍵發現 (數值異常、缺漏標示等), 每條一句話"
  ],
  "concerns": [
    "可能的不合理之處 (要請技師審查的點)"
  ],
  "raw_text": "圖片裡看得到的所有文字 (給後續系統用)",
  "confidence": "high / medium / low"
}}

【重要】
- 不確定的數值留空, 不要猜
- units 跟 flows 抽到幾個就寫幾個
- observations 重點: 數值是否合理、有沒有明顯錯誤、有沒有跟其他資料矛盾
{extra}

【任務】
請判讀附上的圖片, 回傳 JSON (不要加 ```json 標記):
"""


# ──────────────────────────────────────────────────
# 圖片載入
# ──────────────────────────────────────────────────

def _load_pil_image(file_bytes_or_path):
    """把 bytes 或 path 轉成 PIL.Image。"""
    from PIL import Image
    if isinstance(file_bytes_or_path, (bytes, bytearray)):
        return Image.open(io.BytesIO(file_bytes_or_path))
    elif isinstance(file_bytes_or_path, str) and os.path.exists(file_bytes_or_path):
        return Image.open(file_bytes_or_path)
    elif hasattr(file_bytes_or_path, "read"):
        return Image.open(file_bytes_or_path)
    else:
        raise ValueError(f"無法載入圖片: {type(file_bytes_or_path)}")


# ──────────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────────

def process_images(images, mode="extract_rules", focus_hint=""):
    """處理一張或多張圖片。

    Args:
        images: list of bytes/path/file-like (一張也要包成 list)
        mode: "extract_rules" (抽規則) 或 "interpret_diagram" (判讀圖表)
        focus_hint: 給 Gemini 的額外提示 (例: "請特別注意流向是否標示反洗水來源")

    Returns:
        {
            "ok": True,
            "mode": str,
            "image_count": N,
            "result": dict (mode=interpret_diagram) or list[dict] (mode=extract_rules),
            "rows": [...] (僅 mode=extract_rules, 給 rule_importer 用),
            "gemini_usage": {...},
        }
        {"ok": False, "stage": str, "error": str}
    """
    if not images:
        return {"ok": False, "stage": "input", "error": "沒有圖片"}
    if isinstance(images, (bytes, bytearray, str)) or hasattr(images, "read"):
        images = [images]

    # 認證
    try:
        from gemini_extractor import _get_gemini_api_key
        api_key, source = _get_gemini_api_key()
    except Exception as e:
        return {"ok": False, "stage": "auth", "error": f"無法取得 API key: {e}"}
    if not api_key:
        return {"ok": False, "stage": "auth", "error": "未設定 Gemini API key"}

    # 載入圖片
    try:
        pil_images = []
        for img in images:
            pil_images.append(_load_pil_image(img))
    except Exception as e:
        return {"ok": False, "stage": "load", "error": f"圖片載入失敗: {e}"}

    # 組 prompt
    if mode == "extract_rules":
        prompt = _build_extract_rules_prompt()
    elif mode == "interpret_diagram":
        prompt = _build_interpret_diagram_prompt(focus_hint)
    else:
        return {"ok": False, "stage": "mode", "error": f"未知模式: {mode}"}

    # 呼叫 Gemini
    try:
        import google.generativeai as genai
    except ImportError as e:
        return {"ok": False, "stage": "import", "error": f"缺少 google-generativeai: {e}"}

    try:
        genai.configure(api_key=api_key)

        # 把 prompt 跟圖片一起送
        content = [prompt] + pil_images

        # 依序試多個模型 (跟 gemini_extractor 同步)
        last_err = None
        response = None
        for candidate in GEMINI_MODEL_CANDIDATES:
            try:
                model = genai.GenerativeModel(candidate)
                response = model.generate_content(
                    content,
                    generation_config={
                        "temperature": 0.1,
                        "max_output_tokens": 16384,
                        "response_mime_type": "application/json",
                    },
                )
                break
            except Exception as e:
                last_err = e
                continue
        if response is None:
            return {"ok": False, "stage": "gemini",
                    "error": f"所有 Gemini 模型都失敗, 最後錯誤: {last_err}"}

        raw = response.text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        parsed = json.loads(raw)

        usage = {}
        if response.usage_metadata:
            usage = {
                "input_tokens": response.usage_metadata.prompt_token_count,
                "output_tokens": response.usage_metadata.candidates_token_count,
            }
    except json.JSONDecodeError as e:
        return {"ok": False, "stage": "parse",
                "error": f"Gemini 回傳的不是合法 JSON: {e}",
                "raw_response": raw[:500] if 'raw' in dir() else ""}
    except Exception as e:
        return {"ok": False, "stage": "gemini",
                "error": f"Gemini 呼叫失敗: {type(e).__name__}: {e}"}

    # 後處理
    result = {
        "ok": True,
        "mode": mode,
        "image_count": len(pil_images),
        "auth_source": source,
        "gemini_usage": usage,
        "raw_result": parsed,
    }

    if mode == "extract_rules":
        # 對齊 rule_importer 期望的格式 (跟 gemini_extractor 一樣)
        if not isinstance(parsed, list):
            return {"ok": False, "stage": "format", "error": "預期 array 但收到 dict"}

        try:
            import step3f_synonyms
            synonym_normalize = step3f_synonyms.normalize
        except Exception:
            synonym_normalize = lambda x: x

        rows = []
        confidence_dist = {"high": 0, "medium": 0, "low": 0, "unknown": 0}
        for r in parsed:
            if not isinstance(r, dict):
                continue
            compare_raw = r.get("對照項目", "").strip()
            compare_std = synonym_normalize(compare_raw) if compare_raw else ""
            row = {
                "缺失ID": "",
                "原文缺失": r.get("原文缺失", "").strip(),
                "檢查類型": r.get("檢查類型", "").strip(),
                "對照項目": compare_std,
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

        result["rows"] = rows
        result["row_count"] = len(rows)
        result["confidence_dist"] = confidence_dist
    else:
        # interpret_diagram: parsed 是 dict
        result["result"] = parsed

    return result


# ──────────────────────────────────────────────────
# CLI 測試
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print("用法: python gemini_vision.py <圖片路徑> [extract_rules|interpret_diagram]")
        sys.exit(0)

    img_path = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "interpret_diagram"

    if not os.path.exists(img_path):
        print(f"找不到 {img_path}")
        sys.exit(1)

    with open(img_path, "rb") as f:
        data = f.read()

    print(f"=== 處理 {img_path} (mode={mode}) ===")
    result = process_images([data], mode=mode)
    print(json.dumps(result, ensure_ascii=False, indent=2)[:3000])
