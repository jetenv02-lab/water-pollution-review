# -*- coding: utf-8 -*-
"""容錯 JSON 解析 — 用於 Gemini 回傳的 JSON 偶爾有格式問題時嘗試修復。

常見問題:
  1. trailing comma: {"a": 1,}
  2. 中文字串內含未跳脫的 "
  3. 結尾被截斷 (response 長度限制)
  4. 多了一段非 JSON 的解釋文字

策略: 多種修復方法依序嘗試, 任何一種成功就回傳。
"""
import json
import re


def parse_json_tolerant(raw):
    """嘗試解析 JSON 字串, 多重容錯。

    Returns:
        (parsed_dict_or_None, error_or_None)
    """
    # 1. 直接解析
    try:
        return json.loads(raw), None
    except json.JSONDecodeError as e:
        first_err = e

    # 2. 移除尾部多餘逗號 (trailing comma)
    fixed = re.sub(r",(\s*[}\]])", r"\1", raw)
    try:
        return json.loads(fixed), None
    except json.JSONDecodeError:
        pass

    # 3. 截到最後一個 } 或 ] (避免被截斷的 response)
    last_brace = raw.rfind("}")
    last_bracket = raw.rfind("]")
    last_close = max(last_brace, last_bracket)
    if last_close > 0:
        truncated = raw[: last_close + 1]
        # 同時也應用修正 2
        truncated = re.sub(r",(\s*[}\]])", r"\1", truncated)
        try:
            return json.loads(truncated), None
        except json.JSONDecodeError:
            pass

    # 4. 中文字串內可能的未跳脫 "  → 找 key/value 字串內的 " 並 escape
    #    粗略法: 找出每對 "..." 並把中間的 " 跳脫
    escaped = _escape_inner_quotes(raw)
    try:
        return json.loads(escaped), None
    except json.JSONDecodeError:
        pass

    # 5. 找第一個 { 跟最後一個 } 之間的內容 (清掉前後雜訊)
    first_open = raw.find("{")
    if first_open > 0 and last_brace > first_open:
        mid = raw[first_open : last_brace + 1]
        mid = re.sub(r",(\s*[}\]])", r"\1", mid)
        try:
            return json.loads(mid), None
        except json.JSONDecodeError:
            pass

    # 全部失敗, 回傳原始錯誤
    return None, first_err


def _escape_inner_quotes(text):
    # 粗略把字串內部的雙引號 escape
    DQ = chr(34)
    BS = chr(92)
    result = []
    in_str = False
    i = 0
    while i < len(text):
        c = text[i]
        if c == DQ and (i == 0 or text[i - 1] != BS):
            if in_str:
                # 判斷是不是真結尾: 後面 (跳過空白) 是不是 , } ] : 或換行
                j = i + 1
                while j < len(text) and text[j] in " \t":
                    j += 1
                if j >= len(text) or text[j] in ",}]:\n":
                    in_str = False
                    result.append(c)
                else:
                    # 內部的 ", 跳脫
                    result.append(BS + DQ)
            else:
                in_str = True
                result.append(c)
        else:
            result.append(c)
        i += 1
    return "".join(result)


if __name__ == "__main__":
    # 測試
    tests = [
        '{"a": 1, "b": 2}',                                     # 正常
        '{"a": 1, "b": 2,}',                                    # trailing comma
        '{"a": "說 "hello" 給你", "b": 2}',                     # 內部 "
        '{"a": 1, "b": {"c":',                                  # 截斷
        'Some text {"a": 1} more text',                         # 前後雜訊
    ]
    for t in tests:
        parsed, err = parse_json_tolerant(t)
        status = "OK" if parsed else f"FAIL: {err}"
        print(f"{status}: {t[:50]}")
