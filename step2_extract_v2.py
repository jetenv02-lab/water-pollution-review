# -*- coding: utf-8 -*-
"""Step 2 v2: 成熟版申請 PDF 抽取器。

重大修正(相對 v1):
1. 支援全形破折號 T01－01 (不只是半形 T01-01)
2. 從『廢(污)水(前)處理設施資料表』表頭(含「處理單元名稱：xxx 序號：T01－01」)抽取
   單元的標準名稱與序號(而不是用啟發式關鍵字猜)
3. 從『處理單元之進出水質資料』(含「單元序號：T01-01」「進流水流編號：WTB...」)抽取
   進出流水質數據(濃度、質量)
4. 抽取設計參數(液位/停留時間/有效容量)、量測參數(pH/加藥量)、機具設施(液位計/攪拌機)
5. 過濾雜訊(如 D97 這種誤抓的代號)

輸出:
{
  "source_pdf": "...",
  "units": {
    "T01-01": {
      "raw_code": "T01-01",
      "name_in_doc": "中和池",           ← 從 PDF 表頭抽
      "std_tank": "中和池",
      "code_id": "120",
      "size": {"長": "1", "寬": "1.2", "高": "1.2", "有效水深": "1.2", "有效容量": "1.44"},
      "design_params": {"攪拌機轉速": "180~220 rpm"},
      "measure_params": {"pH": "6~9", "加藥量(NaOH)": "63.77~637.66 kg/日"},
      "equipment": [{"name": "pH計", "位置": "池內", "數量": 1}, ...],
      "influent": {  ← 進流水質
        "WTB01-01-1": {"硝酸鹽氮": {"濃度": 50, "質量": 2.45}, ...}
      },
      "effluent": {  ← 出流水質
        "WTA01-01-1": {"硝酸鹽氮": {"濃度": 50, "質量": 30.428}, ...}
      },
      "pages_found": [19, 72]
    },
    ...
  }
}
"""
import json
import os
import re
import sys
from datetime import datetime
import pdfplumber

BASE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE, "參考", "需審查之文件")

# 全形破折號 normalize
DASH_VARIANTS = ["－", "─", "—", "–", "‐", "‑"]


def normalize_text(text):
    """半形化破折號 + 移除多餘空白。"""
    if not text:
        return ""
    s = text
    for d in DASH_VARIANTS:
        s = s.replace(d, "-")
    return s


# 槽體類型關鍵字 (按優先序匹配)
TANK_CLASSIFIERS = [
    ("pH調整暨快混", "pH調整暨快混池"),
    ("pH調整", "pH調整槽"),
    ("快混", "快混槽"),
    ("慢混", "慢混池"),
    ("沉澱", "沉澱池"),
    ("沉降", "沉澱池"),
    ("中和", "中和池"),
    ("放流", "放流池"),
    ("曝氣", "曝氣槽"),
    ("活性污泥", "曝氣槽"),
    ("還原", "還原池"),
    ("氧化", "氧化池"),
    ("陽離子交換", "離子交換樹脂塔"),
    ("陰離子交換", "離子交換樹脂塔"),
    ("離子交換", "離子交換樹脂塔"),
    ("活性碳吸附", "活性碳吸附塔"),
    ("活性碳", "活性碳吸附塔"),
    ("批次反應", "批次反應槽"),
    ("反應槽", "批次反應槽"),
    ("污泥濃縮", "污泥濃縮池"),
    ("濃縮", "濃縮槽"),
    ("污泥脫水", "脫水機"),
    ("污泥烘乾", "污泥烘乾機"),
    ("脫水", "脫水機"),
    ("壓濾", "脫水機"),
    ("砂濾", "砂濾塔"),
    ("過濾", "砂濾塔"),
    ("廢水收集", "廢水收集池"),
    ("廢液收集", "廢液收集池"),
    ("廢水調整", "廢水調整池"),
    ("廢液調整", "廢液調整池"),
    ("調勻", "調勻池"),
    ("調節", "調節池"),
    ("中間池", "中間池"),
    ("濾液", "濾液池"),
    ("貯留", "貯留槽"),
    ("暫存", "暫存槽"),
    ("污泥儲", "污泥儲槽"),
    ("污泥貯", "污泥儲槽"),
    ("收集池", "廢水收集池"),
    ("調整池", "廢水調整池"),
    # ── 邑昇案常見變體 (廢水/廢液貯槽 + 中間槽 + 計量槽) ──
    # 「XX 廢水貯槽 / XX 廢液貯槽」當作「廢水收集池」(都是收集 + 緩衝)
    ("廢水貯槽", "廢水收集池"),
    ("廢液貯槽", "廢水收集池"),
    ("廢水貯", "廢水收集池"),
    ("廢液貯", "廢水收集池"),
    # 「中間槽 / 中間池」(連續槽體間的緩衝) — 學理同「暫存槽」
    ("中間槽", "暫存槽"),
    # 「計量槽」 — 通常是批次取樣 / 量化用, 視為批次反應槽
    ("計量", "批次反應槽"),
    # 「批次槽」(無「反應」字) — 也歸批次反應槽
    ("批次", "批次反應槽"),
    # 「酸化反應槽」/「酸化池」(預酸化、調 pH) — 歸 pH 調整槽
    ("酸化", "pH調整槽"),
]


def classify_tank(name_in_doc):
    """根據 PDF 中讀到的單元名稱,歸類到標準槽體類型。"""
    if not name_in_doc:
        return "未分類"
    for kw, std in TANK_CLASSIFIERS:
        if kw in name_in_doc:
            return std
    return name_in_doc.strip() or "未分類"


# ─────────────────── 各區段的解析 ───────────────────

# 頁 72+: (一)處理單元名稱: xxx 序號: T01－01 代碼: 120
# 注意 pdfplumber 抽出來可能是 "T01- 01"(破折號後有空格), 需放寬
UNIT_HEADER_PATTERN = re.compile(
    r"\(一\)\s*處理單元名稱[：:\s]+(.+?)\s+序號[：:\s]+(T\d{2}\s*-\s*\d{2})\s+代碼[：:\s]+(\d+)"
)

# 頁 19+: 單元序號：T01-01
UNIT_SEQ_PATTERN = re.compile(r"單元序號[：:]\s*(T\d{2}-\d{2})")
# 進出流編號只接受 WT/WM/D 開頭的合法代號, 避免空值後面誤吞「水質項目」標題
INFL_CODE_PATTERN = re.compile(r"進流水流編號[：:]\s*((?:WT[AB]|WM|D|T)\S*)")
EFFL_CODE_PATTERN = re.compile(r"出流水流編號[：:]\s*((?:WT[AB]|WM|D|T)\S*)")

# 尺寸列: 長/直徑 寬 高 有效水深 有效容量 數量
SIZE_DIM_PATTERN = re.compile(
    r"(\d+(?:\.\d+)?)\s*\(?\s*公尺\s*\)?\s+"
    r"(\d+(?:\.\d+)?)\s*\(?\s*公尺?\s*\)?\s+"
    r"(\d+(?:\.\d+)?)\s*\(?\s*公尺?\s*\)?"
)

# (二)設計操作參數 / (三)量測操作參數 中的數值列
# 例: "攪拌機轉速 09 180～ 220 [547]轉／分(rpm) 機具規格"
# 例: "pH值 03 6～ 9 [000]無單位 pH計 1次/日"
PARAM_LINE_PATTERN = re.compile(
    r"^(.+?)\s+(\d{2})\s+(\d+(?:\.\d+)?)\s*[~～]\s*(\d+(?:\.\d+)?)\s+\[(\d+)\](.+)$"
)

# (四)機具設施: pH計 池內 1 0.37 KW
EQUIPMENT_PATTERN = re.compile(
    r"^([^\d\s][^\d]*?)\s+([^\d\s]+)\s+(\d+)\s+(?:(\d+(?:\.\d+)?)\s*)?[Kk][Ww]"
)

# 水質列: "硝酸鹽氮 50 2.45 50 30.428"
# 或不完整: "硝酸鹽氮 50 6.387"
QUALITY_LINE_PATTERN = re.compile(
    r"^(\D+?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*(?:(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?))?\s*$"
)


def find_section_pages(pdf):
    """掃所有頁,找出兩種關鍵區段:
       - 'facility' 頁: 含『(一)處理單元名稱：xxx 序號：T01－XX』
       - 'quality' 頁: 含『單元序號：T01-XX』 或 『進出處理單元之水質資料』

    回傳:
        facility_pages: [(page_index, normalized_text, page_obj)]
        quality_pages: [(page_index, normalized_text)]
    """
    facility_pages = []
    quality_pages = []
    for i, page in enumerate(pdf.pages):
        try:
            text = page.extract_text() or ""
        except Exception:
            continue
        norm = normalize_text(text)
        if "(一)處理單元名稱" in norm and "序號" in norm:
            facility_pages.append((i, norm, page))
        if "單元序號" in norm or "進出處理單元之水質資料" in norm:
            quality_pages.append((i, norm))
    return facility_pages, quality_pages


def _num(s):
    """從字串抽第一個浮點數, 失敗回 None。

    例: "1.69(公\\n尺)" → 1.69
        "(公尺)" → None
        "2.305\\n(公尺)" → 2.305
    """
    if s is None:
        return None
    m = re.search(r"\d+(?:\.\d+)?", str(s))
    if m:
        try:
            return float(m.group())
        except ValueError:
            return None
    return None


def extract_unit_size(page):
    """從 pdfplumber page 物件抽單元尺寸 (用 extract_tables, 比 regex 純文字穩)。

    處理設施資料表的 (一) 區塊有一個 8 欄表:
        材質 | 長/直徑 | 寬 | 高 | 有效水深 | 有效容量 | 數量 | 其他

    PDF 表格的「值列」會散在不同欄, 而且常含換行/單位字串, 用 _num() 容錯抽。

    Returns:
        dict {"材質": "塑膠", "長/直徑": 1.69, "寬": ..., ...} 或 {} (沒找到)
    """
    try:
        tables = page.extract_tables()
    except Exception:
        return {}

    for tbl in tables:
        # 找含「單元尺寸」表頭的表
        # tbl 是 list of rows, 每 row 是 list of cell str
        # 表頭通常是 ['材質', '單元尺寸', None, ...] 或 [None, '長/直徑', '寬', '高', '有效水深', '有效容量', '數量', '其他']
        # 邑昇案 PDF 的表頭被換行: 「長/直\n徑」, 所以要去 \n 和空白後再 match
        header_idx = None
        for ri, row in enumerate(tbl):
            row_str_raw = " ".join(str(c or "") for c in row)
            # 去掉換行/空白 (PDF 可能把「長/直徑」拆成「長/直\n徑」)
            row_str_clean = re.sub(r"\s+", "", row_str_raw)
            if ("長/直徑" in row_str_clean or "長／直徑" in row_str_clean) and "有效容量" in row_str_clean:
                header_idx = ri
                break
        if header_idx is None:
            continue

        # 表頭 + 1 = 資料列
        if header_idx + 1 >= len(tbl):
            continue
        data_row = tbl[header_idx + 1]
        if len(data_row) < 8:
            continue

        # 8 欄: 材質, 長/直徑, 寬, 高, 有效水深, 有效容量, 數量, 其他
        size = {}
        material_raw = data_row[0]
        if material_raw:
            # "ˇ塑膠" → "塑膠"; 去 v/ˇ 勾選符號
            m = re.sub(r"[ˇvVˆ✓☑]", "", str(material_raw)).strip()
            if m:
                size["材質"] = m

        # 數值欄位 (用 _num 容錯抽)
        FIELD_MAP = {
            1: "長/直徑",
            2: "寬",
            3: "高",
            4: "有效水深",
            5: "有效容量",
            6: "數量",
        }
        for col_idx, field in FIELD_MAP.items():
            val = _num(data_row[col_idx]) if col_idx < len(data_row) else None
            if val is not None:
                size[field] = val

        # 其他 (字串, 整列原文)
        if len(data_row) > 7 and data_row[7]:
            other = str(data_row[7]).strip()
            if other:
                size["其他"] = other

        if size:
            return size

    return {}


def _parse_params_section(section_text):
    """解析 (二)/(三) 操作參數區, 處理「PDF 把一筆參數拆成 2-3 行」的常見問題。

    PDF 表格的單一儲存格如果太長 (例如「加藥量(H2SO4（45%）)」), pdfplumber
    抽出來會變成多行:
        加藥量(H2SO4（45        1次/         ← 行 A: 名稱開頭 + 頻率開頭
        05 10.81～ 108.072 [075]公斤／日 依藥劑桶液位差    ← 行 B: 數值列
        ％）) 日                                            ← 行 C: 名稱結尾 + 頻率結尾

    策略: 先找「數值列」(代碼 + 範圍), 再回推「名稱列」(往前找到沒被吃掉的中文文字),
    然後組合成 (name, info) 對。

    Returns: list of (param_name, {"min": float, "max": float, "raw": str})
    """
    results = []
    lines = [ln.rstrip() for ln in section_text.split("\n")]

    # 數值列特徵: 開頭就是「2位數代碼 + 數字 ~ 數字 + [3位數]單位」
    # 例: "05 10.81～ 108.072 [075]公斤／日 依藥劑桶液位差"
    value_pattern = re.compile(
        r"^(\d{2})\s+(\d+(?:\.\d+)?)\s*[~～]\s*(\d+(?:\.\d+)?)\s+\[(\d+)\](.+)$"
    )
    # 表頭關鍵字 — 用來跳過
    header_kw = ("處理單元", "屬處理設施", "數值", "操作參數",
                 "設計參數名稱", "量測參數名稱", "最小值", "最大值", "代碼", "單位")

    used = set()  # 已被「組合進某筆」的行 index, 避免重複用

    for i, ln in enumerate(lines):
        s = ln.strip()
        if not s or i in used:
            continue

        # Case 1: 整行就符合 (名稱 + 數值都在同一行)
        # 例: "pH值 03 6～ 9 [000]無單位 pH計"
        m_full = PARAM_LINE_PATTERN.match(s)
        if m_full:
            pname = m_full.group(1).strip()
            # 跳過表頭
            if any(kw in pname for kw in header_kw):
                continue
            pmin, pmax = m_full.group(3), m_full.group(4)
            tail = m_full.group(6).strip()
            results.append((pname, {
                "min": float(pmin), "max": float(pmax),
                "raw": f"{pmin}~{pmax} {tail}",
            }))
            used.add(i)
            continue

        # Case 2: 只有「數值列」(以代碼開頭) → 名稱在前面幾行
        mv = value_pattern.match(s)
        if not mv:
            continue

        pmin, pmax = mv.group(2), mv.group(3)
        unit_part = mv.group(5).strip()

        # 往前找名稱: 跳過表頭關鍵字, 取「有中文且不是表頭」的最近 1-2 行組合
        name_parts = []
        j = i - 1
        # 名稱可能跨 2 行 (行 A: 名稱開頭, 行 C: 名稱結尾)
        # 但行 C 通常在數值列「後面」, 不是前面 → 改成往後找
        # 結構: 行 A (前) → 行 B (數值, 當前 i) → 行 C (後)
        while j >= 0 and len(name_parts) < 1:
            prev = lines[j].strip()
            if not prev or j in used:
                j -= 1
                continue
            if any(kw in prev for kw in header_kw):
                break
            # 不要把上一筆數值列拿來
            if value_pattern.match(prev) or PARAM_LINE_PATTERN.match(prev):
                break
            # 找到含中文的有效行
            if re.search(r"[\u4e00-\u9fff]", prev):
                name_parts.append((j, prev))
                used.add(j)
                break
            j -= 1

        # 往後找可能的「名稱結尾」(例: "％）) 日")
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            # 結尾特徵: 含「)」「）」「日」「％」, 但**不**含「代碼」或數值範圍
            if (nxt and not value_pattern.match(nxt)
                    and not any(kw in nxt for kw in header_kw)
                    and re.search(r"[)）％]", nxt)
                    and len(nxt) < 20):
                name_parts.append((i + 1, nxt))
                used.add(i + 1)

        if not name_parts:
            continue

        # 組合名稱: 行 A + 行 C (去掉常見尾巴「日」「1次/」)
        # 重組原則: 把所有 name_parts 連起來, 去除尾部頻率字眼
        name_parts.sort(key=lambda x: x[0])
        combined = "".join(p[1] for p in name_parts)
        # 移除「N次/日」「日」這種頻率字眼 (它們是別欄的)
        combined = re.sub(r"\d+\s*次\s*/\s*日?", "", combined)
        combined = re.sub(r"[ \t]+", "", combined)
        # 結尾的孤立「日」也去掉
        combined = re.sub(r"日$", "", combined).strip()

        if any(kw in combined for kw in header_kw):
            continue
        if not combined:
            continue

        results.append((combined, {
            "min": float(pmin), "max": float(pmax),
            "raw": f"{pmin}~{pmax} {unit_part}",
        }))
        used.add(i)

    return results


def parse_facility_page(text, page_index):
    """解析一頁『處理設施資料表』，回傳 unit dict 或 None。"""
    # 找單元表頭
    m = UNIT_HEADER_PATTERN.search(text)
    if not m:
        return None
    name = m.group(1).strip()
    code = re.sub(r"\s+", "", m.group(2))  # T01- 01 → T01-01
    code_id = m.group(3).strip()  # 120

    unit = {
        "raw_code": code,
        "name_in_doc": name,
        "std_tank": classify_tank(name),
        "code_id": code_id,
        "size": {},
        "design_params": {},
        "measure_params": {},
        "equipment": [],
        "influent": {},
        "effluent": {},
        "pages_found": [page_index + 1],
    }

    # 把頁面切成 4 段: (二)設計、(三)量測、(四)機具
    parts = re.split(r"\((二|三|四|五)\)", text)
    # parts = [前面..., '二', 設計內容, '三', 量測內容, '四', 機具內容, '五', ...]

    section_map = {}
    for j in range(1, len(parts) - 1, 2):
        section_map[parts[j]] = parts[j + 1]

    # (二) 設計操作參數
    if "二" in section_map:
        for pname, pinfo in _parse_params_section(section_map["二"]):
            unit["design_params"][pname] = pinfo

    # (三) 量測操作參數
    if "三" in section_map:
        for pname, pinfo in _parse_params_section(section_map["三"]):
            unit["measure_params"][pname] = pinfo

    # (四) 機具設施
    if "四" in section_map:
        for ln in section_map["四"].split("\n"):
            me = EQUIPMENT_PATTERN.match(ln.strip())
            if me:
                eqname = me.group(1).strip()
                # 避免抓到表頭 "名稱"
                if eqname in ("名稱", "設施名稱", ""):
                    continue
                pos = me.group(2)
                qty = int(me.group(3))
                hp = float(me.group(4)) if me.group(4) else None
                unit["equipment"].append({
                    "name": eqname, "位置": pos, "數量": qty, "馬力_kW": hp
                })

    return unit


def parse_quality_page(text):
    """解析『進出水質資料』(可跨頁), 回傳 [(unit_code, infl_code, effl_code, quality_data)]。

    結構:
        單元序號：T01-01
        進出處理單元之水質資料   ← 子區塊 1
        進流水流編號：WTB01-01-1 出流水流編號：WTA01-01-1
        [水質表格]
        進出處理單元之水質資料   ← 子區塊 2 (同一單元的另一股進流)
        進流水流編號：WTB01-01-2 出流水流編號：
        [水質表格]
        進出處理單元之水質資料   ← 子區塊 3...
        ...

    所以要用「進出處理單元之水質資料」切, 而非「單元序號」(否則一個單元只能抽 1 股)。
    """
    results = []
    # 用「進出處理單元之水質資料」(或別名) 切; 同時保留切點前的「單元序號」標記
    # 這樣每個 block 內最多只有一組「進流水流編號」「出流水流編號」, 才能正確抽出多股
    sub_pattern = re.compile(r"進出處理單元之水質資料|進出處理單元水質資料")
    parts = sub_pattern.split(text)
    # parts[0] 是第一個小標題之前的內容(可能含「單元序號：T01-01」)
    # parts[1..] 每段都是一個「進出處理...」標題之後的內容

    current_unit = None
    # 第一段可能含初始 current_unit
    m_first = UNIT_SEQ_PATTERN.search(parts[0]) if parts else None
    if m_first:
        current_unit = m_first.group(1).strip()

    pending_unit_for_next = None  # 上一個 block 結尾出現的「單元序號」, 留給下個 block
    for blk in parts[1:]:
        # 找該 block 的 進流/出流編號 (優先, 因為比「單元序號」更準確)
        m_in = INFL_CODE_PATTERN.search(blk)
        m_out = EFFL_CODE_PATTERN.search(blk)
        infl_code = m_in.group(1) if m_in else None
        effl_code = m_out.group(1) if (m_out and m_out.group(1)) else None

        # 在 blk 內找新的「單元序號：T0X-XX」
        m_unit = UNIT_SEQ_PATTERN.search(blk)
        unit_in_blk = m_unit.group(1).strip() if m_unit else None

        # 從進流編號反推真正的歸屬單元 (例: WTB01-01-6 → T01-01)
        infl_implied_unit = None
        if infl_code:
            mm = re.match(r"^WT[AB](\d{2})[-－](\d{2})", infl_code)
            if mm:
                infl_implied_unit = f"T{mm.group(1)}-{mm.group(2)}"

        # 決定 current_unit 優先序:
        # 1. 進流編號隱含的單元 (最準, 直接綁定)
        # 2. 上個 block 留下來的 pending_unit_for_next
        # 3. block 內出現的 單元序號
        # 4. 沿用之前的 current_unit
        new_pending = None
        if infl_implied_unit:
            current_unit = infl_implied_unit
            # 若 block 內也有 unit_in_blk 但跟進流編號不一致, 留給下個 block
            if unit_in_blk and unit_in_blk != infl_implied_unit:
                new_pending = unit_in_blk
        elif pending_unit_for_next:
            current_unit = pending_unit_for_next
        elif unit_in_blk:
            current_unit = unit_in_blk

        pending_unit_for_next = new_pending
        if not current_unit:
            continue

        # 抽水質列
        infl_q, effl_q = {}, {}
        in_data = False
        for ln in blk.split("\n"):
            ln = ln.strip()
            if "濃度" in ln and "質量" in ln:
                in_data = True
                continue
            if not in_data or not ln:
                continue
            # 把全形空白也當分隔
            ln_norm = re.sub(r"[\s\u3000]+", " ", ln)
            # 4 個數字 = 進+出
            mq = re.match(
                r"^([^\d\s][^\d]*?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*$",
                ln_norm
            )
            if mq:
                item = mq.group(1).strip()
                infl_q[item] = {"濃度": float(mq.group(2)), "質量": float(mq.group(3))}
                effl_q[item] = {"濃度": float(mq.group(4)), "質量": float(mq.group(5))}
                continue
            # 2 個數字 = 單側 (沒出流)
            mq2 = re.match(
                r"^([^\d\s][^\d]*?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*$",
                ln_norm
            )
            if mq2:
                item = mq2.group(1).strip()
                infl_q[item] = {"濃度": float(mq2.group(2)), "質量": float(mq2.group(3))}
                continue
            # pH/水溫: "pH值 1 ~ 7 - 6 ~ 9 -"
            mph = re.match(
                r"^(pH值|水溫.*?)\s+([\d\.]+)\s*[~～-]\s*([\d\.]+)\s*-?\s*([\d\.]+)\s*[~～-]\s*([\d\.]+)\s*-?\s*$",
                ln_norm
            )
            if mph:
                item = mph.group(1).strip()
                infl_q[item] = {"範圍": f"{mph.group(2)}~{mph.group(3)}"}
                effl_q[item] = {"範圍": f"{mph.group(4)}~{mph.group(5)}"}

        if current_unit:
            results.append({
                "unit_code": current_unit,
                "infl_code": infl_code,
                "effl_code": effl_code,
                "infl_quality": infl_q,
                "effl_quality": effl_q,
            })
    return results


# ─────────────────── 主函式 ───────────────────


def extract_application(pdf_path, verbose=True):
    if verbose:
        print(f"=== 開啟 PDF: {pdf_path} ===")

    units = {}

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if verbose:
            print(f"總頁數: {total}")
            print("掃描章節...")

        facility_pages, quality_pages = find_section_pages(pdf)
        if verbose:
            print(f"  處理設施資料表: {len(facility_pages)} 頁")
            print(f"  進出水質資料: {len(quality_pages)} 頁")

        # 解析設施資料表 → 取單元 metadata + 設計/量測/機具 + 尺寸
        for page_idx, text, page_obj in facility_pages:
            unit = parse_facility_page(text, page_idx)
            if unit:
                code = unit["raw_code"]
                # 從表格抽單元尺寸 (材質 / 長/直徑 / 寬 / 高 / 有效水深 / 有效容量 / 數量)
                try:
                    size_info = extract_unit_size(page_obj)
                except Exception:
                    size_info = {}
                if size_info:
                    unit["size"] = size_info

                if code in units:
                    # 已存在:合併 (有時跨頁)
                    units[code]["pages_found"].append(page_idx + 1)
                    units[code]["design_params"].update(unit["design_params"])
                    units[code]["measure_params"].update(unit["measure_params"])
                    units[code]["equipment"].extend(unit["equipment"])
                    # 尺寸: 只在原本沒有時填入 (避免覆蓋已抽到的)
                    if size_info and not units[code].get("size"):
                        units[code]["size"] = size_info
                else:
                    units[code] = unit

        # 解析進出水質 → 加進對應單元
        # 修: 把所有 quality_pages 串成一段, 讓 parse_quality_page 的「current_unit」
        # 狀態能跨頁保留 (例如 T01-01 的多股進流可能跨 3 頁)
        if quality_pages:
            # 用特殊分隔符記錄頁碼, 才能還原 page_idx
            combined_text_parts = []
            page_markers = []  # [(start_pos, page_idx)]
            for page_idx, text in quality_pages:
                page_markers.append((len("\n".join(combined_text_parts)), page_idx))
                combined_text_parts.append(text)
            combined_text = "\n".join(combined_text_parts)

            block_results = parse_quality_page(combined_text)
            for r in block_results:
                code = r["unit_code"]
                # 找該 result 屬於哪一頁 (用 r 內的 _pos 或 fallback 用第一頁)
                page_for_record = r.get("_page_idx", quality_pages[0][0])
                if code not in units:
                    units[code] = {
                        "raw_code": code,
                        "name_in_doc": "(僅水質資料)",
                        "std_tank": "未分類",
                        "code_id": "",
                        "size": {},
                        "design_params": {},
                        "measure_params": {},
                        "equipment": [],
                        "influent": {},
                        "effluent": {},
                        "pages_found": [page_for_record + 1],
                    }
                if (page_for_record + 1) not in units[code]["pages_found"]:
                    units[code]["pages_found"].append(page_for_record + 1)
                if r["infl_code"]:
                    units[code]["influent"][r["infl_code"]] = r["infl_quality"]
                if r["effl_code"]:
                    units[code]["effluent"][r["effl_code"]] = r["effl_quality"]

    result = {
        "source_pdf": os.path.basename(pdf_path),
        "extracted_at": datetime.now().isoformat(),
        "total_units": len(units),
        "units": units,
    }

    # 反推每條 stream 的流量 Q (從質量÷濃度×1000 算)
    # 完整覆蓋 100% 有水質資料的 stream, 不需要 Gemini Vision 跑示意圖解析
    # 若 19 項算出來收斂 → 確信值; 若分散 → 水質表填錯, 後續學理檢查會抓
    try:
        import stream_q_calculator as _sqc
        _sqc.enrich_app_data(result)
        # 順手組「stream_code → Q 對照表」放頂層, 給 UI 直接用
        result["stream_q_map"] = _sqc.build_stream_q_map(result)
    except Exception as _q_err:
        # 反推失敗不致命 — 主流程繼續, 沒 Q 而已
        result["stream_q_error"] = str(_q_err)

    return result


def main():
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
    else:
        pdfs = [f for f in os.listdir(APP_DIR) if f.lower().endswith(".pdf")]
        if not pdfs:
            print(f"在 {APP_DIR} 找不到 PDF")
            return
        pdf_path = os.path.join(APP_DIR, pdfs[0])

    if not os.path.exists(pdf_path):
        print(f"找不到 PDF: {pdf_path}")
        return

    result = extract_application(pdf_path)

    # 輸出 JSON
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = os.path.join(BASE, f"application_{base}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n已輸出 JSON: {out_path}")

    # 印摘要
    print(f"\n=== 摘要 ===")
    print(f"共偵測 {result['total_units']} 個處理單元")
    for code, info in sorted(result["units"].items()):
        n_design = len(info["design_params"])
        n_measure = len(info["measure_params"])
        n_eq = len(info["equipment"])
        n_in = len(info["influent"])
        n_out = len(info["effluent"])
        print(f"  {code} {info['name_in_doc']:30s} → {info['std_tank']:15s} "
              f"頁{info['pages_found'][:3]} | 設計{n_design} 量測{n_measure} "
              f"機具{n_eq} 進{n_in} 出{n_out}")


if __name__ == "__main__":
    main()
