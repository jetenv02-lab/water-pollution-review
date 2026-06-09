# -*- coding: utf-8 -*-
"""抽取 PDF「參、第六項 原廢(污)水水量水質資料」區段的 WMxx / Dxx 資料。

PDF 結構:
    參、水污染防治措施資料/用水、廢(污)水及生產、服務量彙總登記事項
    六、原廢(污)水水量、水質資料
      (一) 原廢(污)水編號: WM01
      (二) 原廢(污)水來源: ✓作業廢水, ✓水洗廢水 ...
      (三) 原廢水產生之每日最大量: 申請每日最大量 752 (立方公尺/日)
      (四) 水質項目:
            氨氮     45
            油脂     8
            錫       1.5
            銅       48
            ...

    放流口 (Dxx) 結構類似, 在「七、放流口資料」區段。

回傳結構:
    {
        "raw_water": {
            "WM01": {
                "code": "WM01",
                "q_cmd": 44.0,
                "sources": ["作業廢水", "水洗"],
                "quality": {"氰化物": "1", "硝酸鹽氮": "50", ...},
                "pages": [54],
            },
            ...
        },
        "discharge": {
            "D01": { ... 同上 ... },
        },
    }

使用情境:
    1. 補水流串接 — WMxx → 處理單元 / 處理單元 → Dxx
    2. 法規檢查 — 比對放流口水質 vs 放流水標準
    3. 質量平衡 — Σ WMxx 進量 = Σ Dxx 放流量 + 污泥
"""
import re


# 編號 header (WM/D, 接受全形括號)
WM_HEADER_RE = re.compile(
    r"\(\s*[一壹]\s*\)\s*原廢\(?\s*[污貨]?\s*\)?水編號[：:]\s*(WM\d+)"
)
DISCHARGE_HEADER_RE = re.compile(
    r"\(\s*[一壹]\s*\)\s*放流口編號[：:]\s*(D\d+)"
)
# 也接受第七項放流口的不同寫法
DISCHARGE_HEADER_ALT_RE = re.compile(
    r"放流口[編代]?號[：:]\s*(D\d+)"
)

# 每日最大量 Q
Q_RE = re.compile(r"申請每日最大量\s*(\d+(?:\.\d+)?)\s*\(立方公尺[／/]日\)")

# 水質列: 「氰化物 1」「pH值 1~ 7」「水溫(攝氏) 15~ 35」「懸浮固體（mg/L） 50」
QUALITY_LINE_RE = re.compile(
    r"^([\u4e00-\u9fff（）()pHＨ值\s\-\/0-9㎎/Ll]+?)\s+"  # 項目名
    r"(\d+(?:\.\d+)?(?:\s*[~～-]\s*\d+(?:\.\d+)?)?)\s*$"  # 數值 (可範圍)
)

# 來源勾選的關鍵字
SOURCE_KEYWORDS = [
    "作業廢水", "洩放廢水", "未接觸冷卻水", "逕流廢水", "污水",
    "研磨或切割", "氟系", "TMAH", "氰系", "鉻系", "銅系", "水洗",
    "氟系(含氟)", "氰系(含氰)", "鉻系(含鉻)", "銅系(含銅)",
]


def _is_checked(line):
    """判斷該行是否含「ˇ」「✓」「V」勾選符號。"""
    if not line:
        return False
    # ˇ 是常見的 PDF 勾選符號, V 在開頭也算
    stripped = line.strip()
    return ("ˇ" in line) or ("✓" in line) or ("✔" in line) \
        or stripped.startswith("V ") or stripped.startswith("∨")


def _extract_block_data(block_text, page_num):
    """從 (一) 編號到下一個 (一) 之間的內容抽資料。"""
    data = {
        "q_cmd": None,
        "sources": [],
        "quality": {},
    }

    # Q
    mq = Q_RE.search(block_text)
    if mq:
        try:
            data["q_cmd"] = float(mq.group(1))
        except ValueError:
            pass

    # 水質: 找「(四) 水質項目」之後到下一個 (xxx) 之前
    quality_section = None
    quality_match = re.search(r"\(\s*四\s*\)\s*水質項目", block_text)
    if quality_match:
        after = block_text[quality_match.end():]
        # 切到下一個 (xxx) (例: (一)下一個 WM / 七、放流口)
        next_section = re.search(r"\(\s*[一壹七]\s*\)|七、|\Z", after)
        if next_section:
            quality_section = after[: next_section.start()]
        else:
            quality_section = after

    if quality_section:
        for ln in quality_section.split("\n"):
            ln = ln.strip()
            if not ln:
                continue
            if "水質項目" in ln or "數值" in ln:
                continue
            m = QUALITY_LINE_RE.match(ln)
            if m:
                item = m.group(1).strip()
                val = m.group(2).strip()
                # 過濾掉太短或非水質的
                if len(item) < 1 or "編號" in item or "原廢" in item:
                    continue
                # 標準化常見項目名
                item = item.replace(" ", "")
                data["quality"][item] = val

    # 來源勾選
    for ln in block_text.split("\n"):
        if not _is_checked(ln):
            continue
        for kw in SOURCE_KEYWORDS:
            if kw in ln and kw not in data["sources"]:
                data["sources"].append(kw)

    return data


def extract_raw_water_and_discharge(pdf_path):
    """從 PDF 抽 WMxx (原廢水) 跟 Dxx (放流口) 資料。

    Returns:
        {
            "raw_water": {WMxx: {code, q_cmd, sources, quality, pages}},
            "discharge": {Dxx: {code, q_cmd, sources, quality, pages}},
        }
    """
    try:
        import pdfplumber
    except ImportError:
        return {"raw_water": {}, "discharge": {}}

    result = {"raw_water": {}, "discharge": {}}

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                try:
                    text = page.extract_text() or ""
                except Exception:
                    continue

                # 找 WM
                wm_blocks = WM_HEADER_RE.split(text)
                # wm_blocks: [前文, WMnn, 該段內容, WMnn, 該段內容, ...]
                for idx in range(1, len(wm_blocks) - 1, 2):
                    wm_code = wm_blocks[idx]
                    block = wm_blocks[idx + 1]
                    block_data = _extract_block_data(block, i)

                    if wm_code not in result["raw_water"]:
                        result["raw_water"][wm_code] = {
                            "code": wm_code,
                            "q_cmd": block_data["q_cmd"],
                            "sources": block_data["sources"],
                            "quality": block_data["quality"],
                            "pages": [i],
                        }
                    else:
                        # 跨頁合併
                        rec = result["raw_water"][wm_code]
                        if i not in rec["pages"]:
                            rec["pages"].append(i)
                        if block_data["q_cmd"] and not rec["q_cmd"]:
                            rec["q_cmd"] = block_data["q_cmd"]
                        for s in block_data["sources"]:
                            if s not in rec["sources"]:
                                rec["sources"].append(s)
                        rec["quality"].update(block_data["quality"])

                # 找 D (放流口)
                for hdr_re in (DISCHARGE_HEADER_RE, DISCHARGE_HEADER_ALT_RE):
                    d_blocks = hdr_re.split(text)
                    for idx in range(1, len(d_blocks) - 1, 2):
                        d_code = d_blocks[idx]
                        block = d_blocks[idx + 1]
                        block_data = _extract_block_data(block, i)

                        if d_code not in result["discharge"]:
                            result["discharge"][d_code] = {
                                "code": d_code,
                                "q_cmd": block_data["q_cmd"],
                                "sources": block_data["sources"],
                                "quality": block_data["quality"],
                                "pages": [i],
                            }
                        else:
                            rec = result["discharge"][d_code]
                            if i not in rec["pages"]:
                                rec["pages"].append(i)
                            if block_data["q_cmd"] and not rec["q_cmd"]:
                                rec["q_cmd"] = block_data["q_cmd"]
                            rec["quality"].update(block_data["quality"])

    except Exception as e:
        return {"raw_water": {}, "discharge": {}, "error": str(e)}

    return result


if __name__ == "__main__":
    import sys, io, json
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

    pdf = sys.argv[1] if len(sys.argv) > 1 else "參考/需審查之文件/申請文件(秋棠)(1150519).pdf"
    r = extract_raw_water_and_discharge(pdf)

    print(f"=== 原廢水 (WMxx): {len(r['raw_water'])} 個 ===")
    for code, d in r["raw_water"].items():
        print(f"\n  {code}: Q={d['q_cmd']} CMD, 頁={d['pages']}")
        print(f"    來源: {d['sources']}")
        print(f"    水質 ({len(d['quality'])} 項): {dict(list(d['quality'].items())[:5])}")

    print(f"\n\n=== 放流口 (Dxx): {len(r['discharge'])} 個 ===")
    for code, d in r["discharge"].items():
        print(f"\n  {code}: Q={d['q_cmd']} CMD, 頁={d['pages']}")
        print(f"    水質 ({len(d['quality'])} 項): {dict(list(d['quality'].items())[:5])}")
