# -*- coding: utf-8 -*-
"""抽取 PDF 「五、與廢(污)水、污泥產生量有關之製程設施、生產或服務規模」原料表。

用途:
    - 從廠商 PDF 抽出全廠原料代碼 + 名稱 + 用量
    - 對照「原料 → 應檢測水質」映射, 推導廠商「應該檢測但漏測」的水質項目
    - 純新增, 不動既有 step2

輸出:
    {
      "案件名": {
        "業別": "金屬電鍍處理程序",
        "編號": "M01",
        "產品": [(代碼, 名稱, 量, 單位), ...],
        "原料": [(代碼, 名稱, 量, 單位), ...],
        "推導應檢測水質": [水質項目, ...],
      }
    }
"""
import json
import re
import sys
from pathlib import Path

import pdfplumber

BASE = Path(__file__).parent
PDF_DIR = BASE / "參考" / "需審查之文件"
OUTPUT = BASE / "production_scale.json"


# ──────────────────────────────────────────────────
# 原料代碼 → 應檢測水質項目映射
# 代碼為環境部 IPC 物質代碼; 規則依物質學理推導
# ──────────────────────────────────────────────────
MATERIAL_TO_WATER_QUALITY = {
    # 銅化合物 → 銅
    "180052": ["銅"],                # 氧化銅
    "180071": ["銅"],                # 硫酸銅
    "186099": ["銅", "總磷"],         # 焦磷酸銅
    "251799": ["銅"],                # 銅球
    "240112": ["銅"],                # 銅箔
    "260057": ["銅"],                # 印刷電路用銅箔基板

    # 錫化合物 → 錫
    "180080": ["錫", "硫酸根"],       # 硫酸亞錫
    "190186": ["錫"],                # 剝錫液
    "241299": ["錫"],                # 錫球
    "190178": ["錫"],                # 錫膏

    # 鎳化合物 → 鎳
    "180070": ["鎳", "硫酸根"],       # 硫酸鎳

    # 氰化合物 → 氰化物 + 對應金屬
    "180088": ["氰化物", "鉀"],       # 氰化鉀
    "180089": ["氰化物", "鋅"],       # 氰化鋅
    "180090": ["氰化物", "銀"],       # 氰化銀

    # 鈷化合物 → 鈷
    "181599": ["鈷", "硫酸根"],       # 硫酸鈷

    # 鋁化合物 → 鋁 + SS
    "180067": ["鋁", "硫酸根", "懸浮固體（mg/L）"],  # 硫酸鋁

    # 強酸 → pH, 對應陰離子
    "180030": ["pH", "硫酸根"],       # 濃硫酸
    "180031": ["pH", "氯離子"],       # HCl
    "180032": ["pH", "硝酸鹽氮"],     # 硝酸
    "181199": ["pH", "硫酸根"],       # 其他無機酸: 硫酸

    # 強鹼 → pH
    "180041": ["pH"],                # NaOH
    "180140": ["pH", "鉀"],           # 碳酸鉀
    "180139": ["pH"],                # 碳酸鈉 (純鹼)

    # 氧化劑
    "180060": ["化學需氧量（mg/L）"], # H2O2 → 影響 COD
    "180081": ["硫酸根", "懸浮固體（mg/L）"],  # 過硫酸鈉
    "180505": ["懸浮固體（mg/L）"],   # 矽酸鹽 → SS

    # 其他無機鹽
    "181099": ["氯離子"],            # 氯酸鈉
    "180120": ["氟鹽"],              # 氟化銨
    "180141": ["鈣", "懸浮固體（mg/L）"],  # 碳酸鈣

    # 有機溶劑 → COD
    "180259": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 丁酮
    "180293": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 冰醋酸
    "190131": ["化學需氧量（mg/L）"], # 稀釋劑
    "190046": ["化學需氧量（mg/L）", "真色色度"],  # 油墨

    # 油脂 / 食品 / 紙
    "081399": ["油脂"],              # 食用植物油
    "180928": ["化學需氧量（mg/L）"], # 聚醯胺樹脂
    "180454": ["化學需氧量（mg/L）"], # 橡膠乳液
    "190288": ["化學需氧量（mg/L）"], # 松香

    # 食品原料 → BOD/COD/SS
    "080104": ["化學需氧量（mg/L）", "生化需氧量（mg/L）", "懸浮固體（mg/L）"],  # 米
    "080106": ["化學需氧量（mg/L）", "生化需氧量（mg/L）", "懸浮固體（mg/L）"],  # 麵粉
    "080135": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 食用澱粉
    "080142": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 樹薯粉
    "080028": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 魚漿
    "010006": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 紅豆
    "010036": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 仙草
    "080133": ["化學需氧量（mg/L）", "生化需氧量（mg/L）"],  # 玉米粉
    "150001": ["化學需氧量（mg/L）", "懸浮固體（mg/L）"],   # 木漿
}


# 只抓「五、與廢(污)水、污泥產生量有關之製程設施」這個明確標題, 跳過申請目錄的勾選列
# 容忍各種頓號變體 + 空白
SECTION_START = re.compile(r"五[、,]\s*與廢\s*[\(（]\s*污\s*[\)）]\s*水.{0,5}污泥產生量")
TABLE_ROW = re.compile(
    r"[\[【]\s*(\d{6})\s*[\]】]\s*"      # 代碼 [XXXXXX]
    r"(.+?)\s+"                          # 名稱 (greedy)
    r"(\d+(?:\.\d+)?)\s+"                # 數量
    r"由核發機關"                         # 後面接 "由核發機關"
)
# 編號 M01 可能被 PDF 拆成 "M0\n1", 用 \s* 容忍空白換行
SERVICE_NAME_RE = re.compile(r"製程設施[、,]\s*生產或服務名稱\s+(.+?)\s+編號\s*([A-Z]\s*\d+)")


def extract_section_5(pdf_path):
    """從 PDF 抽出「五、」段落, 解析原料表。

    策略: 必須先找到「製程設施、生產或服務名稱」+「(一)」這種真正進入第五章的標記,
    才認為是真標題, 避開「變更項目表」之類的勾選清單誤觸發。
    """
    pdf = pdfplumber.open(pdf_path)
    found_pages = []
    text_buf = []

    in_section = False
    for i, p in enumerate(pdf.pages):
        t = p.extract_text() or ""
        # 真正的第五章「五、與廢(污)水...製程設施、生產或服務規模」會緊跟著
        # 「(一)製程設施、生產或服務名稱」+「編號 MXX」, 用這 3 個條件確認
        is_real_section = (
            SECTION_START.search(t)
            and "(一)" in t
            and ("製程設施、生產或服務名稱" in t or "生產或服務名稱" in t)
            and re.search(r"編號\s*[A-Z]\s*\d", t)
        )
        if is_real_section:
            in_section = True
            found_pages.append(i + 1)
            text_buf.append(t)
        elif in_section:
            # 連續抽取直到下個「六、」或「七、」標題
            if re.search(r"六、\s*原廢", t) or re.search(r"六、\s*[原核]", t):
                cutoff = re.search(r"六、", t)
                text_buf.append(t[:cutoff.start()] if cutoff else t)
                break
            else:
                found_pages.append(i + 1)
                text_buf.append(t)
                if len(found_pages) >= 8:  # 安全停止
                    break

    pdf.close()

    full_text = "\n".join(text_buf)

    # 找業別 + 編號
    m = SERVICE_NAME_RE.search(full_text)
    service_name = m.group(1).strip() if m else None
    service_code = m.group(2).strip() if m else None

    # 找產品/原料分隔: 看每行的 ˇ 字元
    products = []
    materials = []

    # 把跨行黏起來 (PDF 抽出來常被換行打散)
    flat = re.sub(r"\s+", " ", full_text)

    # 找所有 [XXXXXX]
    pos = 0
    while True:
        m = re.search(r"[\[【](\d{6})[\]】]", flat[pos:])
        if not m:
            break
        code = m.group(1)
        start = pos + m.start()
        # 後面找下個 [XXXXXX] 或結尾
        next_m = re.search(r"[\[【]\d{6}[\]】]", flat[pos + m.end():])
        end = pos + m.end() + next_m.start() if next_m else len(flat)
        chunk = flat[start:end]

        # 從 chunk 找名稱 + 數量 + 單位
        # 名稱: [XXXXXX] 後面到第一個數字之前 (扣掉空白)
        name_m = re.match(r"[\[【]\d{6}[\]】]\s*(.+?)\s+(\d+(?:\.\d+)?)", chunk)
        if not name_m:
            pos = pos + m.end()
            continue
        name = name_m.group(1).strip()
        try:
            amount = float(name_m.group(2))
        except ValueError:
            pos = pos + m.end()
            continue

        # 單位: 找 "公斤／日" / "公噸／日" / "立方公尺／日" / "平方公尺／日" / 個/日 等
        # 容忍跨行: "公斤" 跟 "／日" 可能被換行/空白分開
        unit_m = re.search(r"(公斤|公噸|立方公尺|平方公尺|個|噸)[\s\n]*[／/][\s\n]*日", chunk[name_m.end():])
        unit = (unit_m.group(0) if unit_m else "未知").replace(" ", "").replace("\n", "").replace("／", "/")

        # 看前面文字決定產品/原料
        prev_chunk = flat[max(0, start - 50):start]
        is_product = "ˇ產品量" in prev_chunk or "✓產品量" in prev_chunk

        entry = {
            "代碼": code,
            "名稱": name,
            "量": amount,
            "單位": unit,
        }
        (products if is_product else materials).append(entry)

        pos = pos + m.end()

    # 用代碼推導應檢測水質項目
    derived_items = set()
    for ent in products + materials:
        if ent["代碼"] in MATERIAL_TO_WATER_QUALITY:
            derived_items.update(MATERIAL_TO_WATER_QUALITY[ent["代碼"]])

    return {
        "業別": service_name,
        "編號": service_code,
        "頁碼": found_pages,
        "產品": products,
        "原料": materials,
        "推導應檢測水質": sorted(derived_items),
    }


def main():
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    results = {}
    print(f"=== 抽取「五、生產規模」 ({len(pdfs)} 案) ===\n")
    for pdf in pdfs:
        case_name = pdf.stem
        try:
            r = extract_section_5(str(pdf))
        except Exception as e:
            print(f"[X] {case_name}: {e}")
            continue
        results[case_name] = r
        print(f"\n📋 {case_name}")
        print(f"  業別:      {r['業別']}")
        print(f"  編號:      {r['編號']}")
        print(f"  頁碼:      {r['頁碼']}")
        print(f"  產品 ({len(r['產品'])} 項):")
        for p in r['產品'][:5]:
            print(f"    [{p['代碼']}] {p['名稱']}: {p['量']} {p['單位']}")
        if len(r['產品']) > 5:
            print(f"    ... 另 {len(r['產品']) - 5} 項")
        print(f"  原料 ({len(r['原料'])} 項):")
        for m in r['原料'][:5]:
            print(f"    [{m['代碼']}] {m['名稱']}: {m['量']} {m['單位']}")
        if len(r['原料']) > 5:
            print(f"    ... 另 {len(r['原料']) - 5} 項")
        print(f"  ⚠️ 推導應檢測水質 ({len(r['推導應檢測水質'])} 項):")
        for w in r['推導應檢測水質']:
            print(f"    - {w}")

    OUTPUT.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n✅ 結果寫入 {OUTPUT.name}")


if __name__ == "__main__":
    main()
