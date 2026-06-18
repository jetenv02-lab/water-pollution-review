# -*- coding: utf-8 -*-
"""Step 3c: 處理單元預設值資料庫 + 事業類別申報項目對照。

資料來源: jetwatersystem (Nick Chang / ZN Studio) 的設計準則,
經授權整合至本系統供智能審查使用。

提供:
- 各處理單元類型對各水質項目的預設削減率 (用來判斷申請文件填的去除率是否合理)
- 各水質項目的預設原廢水濃度 (用來比對申請文件原廢水水質是否在合理範圍)
- 事業類別對應的申報水質項目清單 (檢查申請文件水質項目是否漏項)
- 額外進流類型 (RAS / 化學藥劑 / 上清液 等)
- 計算公式: 質量、混合濃度、出流濃度
"""

# ──────────────────────────────────────────────────
# A. 處理單元 × 水質項目 預設削減率 (%)
# ──────────────────────────────────────────────────
# 用法: 申請文件如果填的去除率與預設值差異 > 20%, 標「待確認複核」
UNIT_DEFAULT_REMOVAL = {
    # 進流類
    "進流":              {},
    "攔污柵":            {"SS": 5},
    "沉砂池":            {"SS": 5},
    "調勻池":            {},
    "廢水調整池":         {},

    # 初級處理
    "初級沉澱池":         {"BOD": 30, "SS": 50, "COD": 25, "總磷": 10},
    "沉澱池":             {"BOD": 30, "SS": 85, "COD": 25, "總磷": 10},  # 一般稱呼

    # 生物處理
    "標準活性污泥池":      {"BOD": 85, "SS": 85, "COD": 80, "氨氮": 30, "總氮": 20, "總磷": 20},
    "延長曝氣池":         {"BOD": 90, "SS": 90, "COD": 85, "氨氮": 85, "總氮": 30, "總磷": 25},
    "A2O生物池":          {"BOD": 90, "SS": 90, "COD": 85, "氨氮": 90, "總氮": 70, "總磷": 80},
    "SBR反應槽":          {"BOD": 90, "SS": 90, "COD": 85, "氨氮": 85, "總氮": 60, "總磷": 50},
    "MBR膜生物反應器":     {"BOD": 95, "SS": 99, "COD": 90, "氨氮": 95, "總氮": 70, "總磷": 85, "大腸桿菌群": 99.9},
    "曝氣槽":             {"BOD": 85, "SS": 85, "COD": 80, "氨氮": 30},  # 一般稱呼

    # 二級沉澱
    "二級沉澱池":         {"SS": 90, "BOD": 5, "COD": 5},

    # 物化處理
    "快濾池":             {"SS": 50, "總磷": 20},
    "砂濾塔":             {"SS": 50, "總磷": 20},  # 別名
    "加藥混凝池":         {"SS": 60, "總磷": 70, "真色色度": 50},
    "快混槽":             {},  # 快混槽本身不應有去除 (學理: 無固液分離)
    "慢混池":             {},
    "pH調整槽":           {},  # 只調整 pH, 不去除
    "pH調整暨快混池":      {},  # 同上
    "中和池":             {},

    # 消毒
    "消毒池":             {"大腸桿菌群": 99.9, "BOD": 5, "COD": 5},

    # 過濾 / 吸附
    "活性碳吸附塔":        {"COD": 50, "真色色度": 80, "BOD": 30},
    "活性碳吸附裝置":      {"COD": 50, "真色色度": 80, "BOD": 30},

    # 重金屬處理 (學理參考)
    "離子交換樹脂塔":      {"銅": 95, "鎳": 95, "鋅": 95, "鎘": 95, "六價鉻": 95},
    "還原池":              {"六價鉻": 95},  # 鉻還原: Cr(VI) → Cr(III)
    "氧化池":              {"氰化物": 90, "BOD": 20, "COD": 15},

    # 污泥處理 (污泥側、出流濃度上升合理)
    "污泥濃縮池":         {"_sludge_side": True},
    "濃縮槽":             {"_sludge_side": True},
    "脫水機":             {"_sludge_side": True},
    "污泥儲槽":           {"_sludge_side": True},
    "貯留槽":             {},
    "暫存槽":             {},
    "放流池":             {},
}

# ──────────────────────────────────────────────────
# A.1 水質項目預設原廢水濃度
# ──────────────────────────────────────────────────
# 用法: 申請文件原廢水濃度遠超這些值, 可能設計值偏離實測
DEFAULT_RAW_CONCENTRATIONS = {
    "pH":               {"value": "6-9",    "unit": "-",          "is_range": True},
    "水溫":             {"value": 25,       "unit": "°C",         "is_range": False},
    "BOD":              {"value": 200,      "unit": "mg/L",       "is_range": False},
    "COD":              {"value": 350,      "unit": "mg/L",       "is_range": False},
    "SS":               {"value": 250,      "unit": "mg/L",       "is_range": False},
    "懸浮固體":          {"value": 250,      "unit": "mg/L",       "is_range": False},
    "氨氮":             {"value": 30,       "unit": "mg/L",       "is_range": False},
    "總氮":             {"value": 40,       "unit": "mg/L",       "is_range": False},
    "總磷":             {"value": 8,        "unit": "mg/L",       "is_range": False},
    "真色色度":         {"value": 100,      "unit": "ADMI",       "is_range": False},
    "自由有效餘氯":      {"value": 1,        "unit": "mg/L",       "is_range": False},
    "大腸桿菌群":        {"value": 200000,   "unit": "CFU/100mL",  "is_range": False},
    "油脂":             {"value": 30,       "unit": "mg/L",       "is_range": False},
    "陰離子界面活性劑":  {"value": 10,       "unit": "mg/L",       "is_range": False},
    "硝酸鹽氮":         {"value": 10,       "unit": "mg/L",       "is_range": False},
    "氟鹽":             {"value": 5,        "unit": "mg/L",       "is_range": False},
    "氰化物":           {"value": 0.5,      "unit": "mg/L",       "is_range": False},
    "總鉻":             {"value": 1,        "unit": "mg/L",       "is_range": False},
    "六價鉻":           {"value": 0.5,      "unit": "mg/L",       "is_range": False},
    "鎘":               {"value": 0.03,     "unit": "mg/L",       "is_range": False},
    "鎳":               {"value": 1,        "unit": "mg/L",       "is_range": False},
    "銅":               {"value": 3,        "unit": "mg/L",       "is_range": False},
    "總汞":             {"value": 0.005,    "unit": "mg/L",       "is_range": False},
    "鉛":               {"value": 1,        "unit": "mg/L",       "is_range": False},
    "砷":               {"value": 0.5,      "unit": "mg/L",       "is_range": False},
    "鋅":               {"value": 5,        "unit": "mg/L",       "is_range": False},
    "硼":               {"value": 1,        "unit": "mg/L",       "is_range": False},
    "錫":               {"value": 1,        "unit": "mg/L",       "is_range": False},
    "鉬":               {"value": 0.6,      "unit": "mg/L",       "is_range": False},
}


# ──────────────────────────────────────────────────
# D. 事業類別 → 申報水質項目 + 申報頻率
# ──────────────────────────────────────────────────
# 用法: 知道事業類別後, 檢查申請 PDF 的水質項目是否完整
BUSINESS_TYPES = {
    "製糖業": {
        "id": 1,
        "general": ["pH", "水溫", "BOD", "COD", "SS"],
        "specific_1": [],
        "specific_2": [],
    },
    "紡織業": {
        "id": 2,
        "general": ["pH", "水溫", "BOD", "COD", "SS", "真色色度", "自由有效餘氯"],
        "specific_1": [],
        "specific_2": [],
    },
    "電鍍業": {
        "id": 19,
        "general": ["pH", "水溫", "COD", "SS", "氨氮"],
        "specific_1": ["總鉻", "鎘", "鎳", "銅", "總汞", "鉛", "砷", "鋅", "氰化物", "硝酸鹽氮", "氟鹽"],
        "specific_2": ["六價鉻", "硼", "錫", "鉬"],
    },
    "晶圓製造及半導體製造業": {
        "id": 20,
        "general": ["pH", "水溫", "COD", "SS", "氨氮", "總磷"],
        "specific_1": ["總鉻", "鎘", "鎳", "銅", "總汞", "鉛", "砷", "鋅", "氰化物",
                       "硝酸鹽氮", "氟鹽", "陰離子界面活性劑"],
        "specific_2": ["六價鉻", "硼", "錫", "鉬"],
    },
    "光電材料及元件製造業": {
        "id": 21,  # 對應水措類似於 20
        "general": ["pH", "水溫", "COD", "SS", "氨氮", "總磷"],
        "specific_1": ["總鉻", "鎘", "鎳", "銅", "總汞", "鉛", "砷", "鋅",
                       "氰化物", "硝酸鹽氮", "氟鹽", "陰離子界面活性劑"],
        "specific_2": ["六價鉻", "硼", "錫", "鉬"],
    },
    "印刷電路板製造業": {
        "id": 22,
        "general": ["pH", "水溫", "COD", "SS"],
        "specific_1": ["總鉻", "鎘", "鎳", "銅", "鉛", "砷", "鋅", "氰化物", "氟鹽"],
        "specific_2": ["六價鉻", "硼"],
    },
    "食品製造業": {
        "id": 42,
        "general": ["pH", "水溫", "BOD", "COD", "SS"],
        "specific_1": ["油脂"],
        "specific_2": [],
    },
    "餐飲業/觀光旅館": {
        "id": 55,
        "general": ["pH", "水溫", "BOD", "COD", "SS", "大腸桿菌群", "總氮", "總磷"],
        "specific_1": ["油脂"],
        "specific_2": [],
    },
    "醫院/醫事機構": {
        "id": 53,
        "general": ["pH", "水溫", "BOD", "COD", "SS", "大腸桿菌群", "自由有效餘氯", "氨氮"],
        "specific_1": [],
        "specific_2": [],
    },
    "公共污水下水道": {
        "id": 63,
        "general": ["pH", "水溫", "BOD", "COD", "SS", "大腸桿菌群", "自由有效餘氯",
                    "總氮", "氨氮", "總磷"],
        "specific_1": [],
        "specific_2": [],
    },
    "社區專用污水下水道": {
        "id": 64,
        "general": ["pH", "水溫", "BOD", "SS", "大腸桿菌群"],
        "specific_1": [],
        "specific_2": [],
    },
    "其他(自訂)": {
        "id": 99,
        "general": ["pH", "水溫", "BOD", "COD", "SS"],
        "specific_1": [],
        "specific_2": [],
    },
}

REPORT_FREQUENCY = {
    "general": "每三個月",
    "specific_1": "每六個月",
    "specific_2": "每年",
}


# ──────────────────────────────────────────────────
# F. 額外進流類型 (RAS, 化學藥劑, 上清液...)
# ──────────────────────────────────────────────────
INLET_TYPES = {
    "RAS": {
        "name": "迴流污泥",
        "description": "Return Activated Sludge",
        "default_flow_cmd": 300,
        "expected_ratio_to_main": (0.3, 1.5),  # RAS / Q_in 的合理範圍
    },
    "化學藥劑": {
        "name": "化學藥劑",
        "description": "PAC、聚合物、NaOH 等",
        "default_flow_cmd": 10,
    },
    "上清液": {
        "name": "上清液",
        "description": "污泥濃縮/脫水上清液",
        "default_flow_cmd": 50,
    },
    "其他處理線": {
        "name": "其他處理線",
        "description": "來自其他處理線的水流",
        "default_flow_cmd": 100,
    },
}


# ──────────────────────────────────────────────────
# B. 計算公式 (純函式)
# ──────────────────────────────────────────────────

def calculate_mass_kg_day(flow_cmd, concentration_mg_per_l):
    """質量 (kg/day) = 流量 (CMD) × 濃度 (mg/L) × 0.001。"""
    try:
        return float(flow_cmd) * float(concentration_mg_per_l) * 0.001
    except (TypeError, ValueError):
        return None


def calculate_mixed_concentration(main_flow, main_conc, additional_inlets):
    """混合進流濃度 = Σ(Q × C) / Σ(Q)。

    Args:
        main_flow: 主進流量 (CMD)
        main_conc: 主進流濃度 (mg/L)
        additional_inlets: [{flow, concentration}, ...]
    Returns:
        混合濃度 (mg/L), 或 None 若計算不出
    """
    try:
        total_q = float(main_flow)
        total_qc = float(main_flow) * float(main_conc)
        for inlet in additional_inlets:
            q = float(inlet.get("flow", 0))
            c = float(inlet.get("concentration", 0))
            total_q += q
            total_qc += q * c
        if total_q <= 0:
            return None
        return total_qc / total_q
    except (TypeError, ValueError):
        return None


def calculate_outlet_concentration(inlet_conc, removal_rate_pct):
    """出流濃度 = 進流濃度 × (1 - 削減率/100)。"""
    try:
        return float(inlet_conc) * (1 - float(removal_rate_pct) / 100)
    except (TypeError, ValueError):
        return None


# ──────────────────────────────────────────────────
# 查詢輔助函式
# ──────────────────────────────────────────────────

def get_default_removal(std_tank, item_name):
    """取得標準槽體對某水質項目的預設削減率(%);找不到回 None。"""
    rates = UNIT_DEFAULT_REMOVAL.get(std_tank, {})
    if "_sludge_side" in rates:
        return None  # 污泥側豁免
    return rates.get(item_name)


def is_sludge_side_unit(std_tank):
    """判斷是否為污泥側單元(出流濃縮為合理)。"""
    rates = UNIT_DEFAULT_REMOVAL.get(std_tank, {})
    return rates.get("_sludge_side", False)


def get_business_required_items(business_type):
    """取得事業類別應申報的所有水質項目(扁平 list)。"""
    bt = BUSINESS_TYPES.get(business_type)
    if not bt:
        return []
    return bt["general"] + bt["specific_1"] + bt["specific_2"]


def check_missing_report_items(business_type, declared_items):
    """檢查申請文件是否漏報該事業類別應有的項目。

    Args:
        business_type: 事業類別名稱
        declared_items: 申請文件實際申報的水質項目 list

    Returns:
        list of {"item": ..., "category": ..., "frequency": ...}
    """
    bt = BUSINESS_TYPES.get(business_type)
    if not bt:
        return []
    missing = []
    declared_set = set(declared_items)
    for category in ("general", "specific_1", "specific_2"):
        for item in bt[category]:
            if item not in declared_set:
                missing.append({
                    "item": item,
                    "category": category,
                    "frequency": REPORT_FREQUENCY[category],
                })
    return missing


def detect_business_type(application_text):
    """從申請文件文字中啟發式偵測事業類別。"""
    text = application_text or ""
    # 簡易關鍵字比對
    if "印刷電路板" in text or "PCB" in text:
        return "印刷電路板製造業"
    if "晶圓" in text or "半導體" in text:
        return "晶圓製造及半導體製造業"
    if "光電" in text:
        return "光電材料及元件製造業"
    if "電鍍" in text:
        return "電鍍業"
    if "餐飲" in text or "觀光旅館" in text:
        return "餐飲業/觀光旅館"
    if "醫院" in text or "醫事" in text:
        return "醫院/醫事機構"
    if "食品" in text:
        return "食品製造業"
    if "紡織" in text or "染整" in text:
        return "紡織業"
    if "製糖" in text:
        return "製糖業"
    return None  # 未識別
