# -*- coding: utf-8 -*-
"""加藥質量補償計算模組 (Step 1 — 純新增, 不動任何既有檔案)。

用途:
    - 質量平衡檢查時, 把加藥引入的質量補上來
    - 例: 快混槽加 PAC 100 mg/L, Q = 1000 m³/d
      → PAC 引入 SS 質量 = 100 × 1000 / 1000 × 0.30 = 30 kg/d
    - 系統用法: in_mass_total = in_mass + chemical_added
    - 然後比較 out_mass vs in_mass_total

第一版設計 (典型值粗估, 不依賴 step2 PDF 加藥抽取):
    - 每個 std_tank 設 default 加藥配方 (mg/L 為單位的劑量)
    - 透過化學計量算引入質量

未來升級 (精算):
    - 改為優先讀 step2 抽出的「加藥表 + 各槽劑量」
    - fallback 才用典型值

對外 API:
    compute_chemical_mass(unit, item, q_cmd) → kg/d
    get_typical_dosing(std_tank) → dict[chem → mg/L]
    get_conversion_factor(chem, item) → float (0~1)
    list_dosing_chemicals(std_tank) → list[str] (該槽體典型加哪些藥)

注意:
    - 不引入 SS 的藥劑 (例: NaOH 純鹼) 轉換係數 = 0, 不會影響 SS 質平
    - PAM 純品 100% 變固體, 但實際進槽是 0.1% 稀釋液 → 引入量是 PAM 純品的量
    - 化學計量是「上限」, 實際可能 50~80%
"""

# ──────────────────────────────────────────────────
# 典型加藥配方 (mg/L, 對進流水的劑量)
# 來源: 環工設計慣用值 + 污水處理廠設計手冊 + 學理教材
# ──────────────────────────────────────────────────
TYPICAL_DOSING = {
    # 反應類 (酸鹼/混凝/氧化還原)
    "pH調整槽":         {"NaOH": 50, "H2SO4": 30},
    "pH調整池":         {"NaOH": 50, "H2SO4": 30},
    "中和池":           {"NaOH": 50, "H2SO4": 30},
    "pH調整暨快混池":   {"NaOH": 50, "H2SO4": 30, "PAC": 100},
    "pH調整池暨快混池": {"NaOH": 50, "H2SO4": 30, "PAC": 100},
    "快混槽":           {"PAC": 100},
    "快混池":           {"PAC": 100},
    "慢混池":           {"PAM": 2},
    "慢混槽":           {"PAM": 2},
    "混凝膠凝池":       {"PAC": 100, "PAM": 2},
    "氧化池":           {"NaClO": 200},
    "氰系氧化槽":       {"NaClO": 300, "NaOH": 80},
    "鉻系還原槽":       {"NaHSO3": 200, "H2SO4": 100},
    "批次反應槽":       {},  # 視批次配方, 不假設

    # 分離類 (本身通常不加藥, 但有些含過濾的會反洗)
    "沉澱池":           {},
    "沉降池":           {},
    "浮除槽":           {"PAM": 1},  # 少量加幫絮凝上浮
    "砂濾塔":           {},  # 反洗用清水, 不加藥
    "砂濾器":           {},
    "活性碳吸附塔":     {},
    "活性碳吸附裝置":   {},
    "離子交換樹脂塔":   {},  # 再生用酸鹼, 但日常運轉不加
    "油脂分離槽":       {},
    "預處理池":         {},

    # 生物類 (有時加碳源/磷源)
    "曝氣槽":           {"尿素": 5, "磷酸": 1},  # 補 N, P 元素 (BOD:N:P=100:5:1)
    "厭氧池":           {},
    "接觸氧化池":       {},
    "MBR":              {},

    # 污泥類
    "脫水機":           {"PAM": 5},  # 污泥脫水陽離子 PAM
    "污泥離心式脫水機": {"PAM": 5},
    "污泥帶濾式脫水機": {"PAM": 5},
    "濃縮槽":           {},
    "污泥濃縮設施":     {},
    "污泥儲槽":         {},
    "濾液池":           {},

    # 通用/中性類 (不加藥)
    "暫存槽":           {},
    "貯留槽":           {},
    "廢水調整池":       {},
    "廢水收集池":       {},
    "調勻池":           {},
    "放流池":           {},
}


# ──────────────────────────────────────────────────
# 化學計量轉換係數 (kg 引入水質項目 / kg 藥劑原料)
# 來源: 化學計量 + 典型反應產物
# ──────────────────────────────────────────────────
CONVERSION = {
    # 凝集劑 (引入 SS 為主)
    "PAC": {
        "懸浮固體（mg/L）": 0.30,  # Al₂O₃·5H₂O 形成 Al(OH)₃ 沉澱 ≈ 30% mass
        "懸浮固體(mg/L)":  0.30,
        "SS":              0.30,
        "氯離子":          0.05,  # PAC 含部分 Cl⁻
        "氯鹽":            0.05,
        "Cl":              0.05,
    },
    "FeCl3": {
        "懸浮固體（mg/L）": 0.35,  # 形成 Fe(OH)₃
        "懸浮固體(mg/L)":  0.35,
        "SS":              0.35,
        "氯離子":          0.65,  # 3 個 Cl⁻
        "氯鹽":            0.65,
        "Cl":              0.65,
    },
    "Al2(SO4)3": {
        "懸浮固體（mg/L）": 0.40,
        "懸浮固體(mg/L)":  0.40,
        "SS":              0.40,
        "硫酸根":          0.85,
        "硫酸鹽":          0.85,
        "SO4":             0.85,
    },
    "PAM": {
        "懸浮固體（mg/L）": 1.00,  # 純品幾乎全變 SS
        "懸浮固體(mg/L)":  1.00,
        "SS":              1.00,
    },

    # 酸鹼
    "NaOH": {},  # 純鹼不引入 SS, 只改 pH
    "Ca(OH)2": {
        "懸浮固體（mg/L）": 0.50,  # 形成 CaCO₃/CaF₂ 沉澱
        "懸浮固體(mg/L)":  0.50,
        "SS":              0.50,
    },
    "H2SO4": {
        "硫酸根":          0.98,
        "硫酸鹽":          0.98,
        "SO4":             0.98,
    },
    "HCl": {
        "氯離子":          0.97,
        "氯鹽":            0.97,
        "Cl":              0.97,
    },

    # 氧化/還原劑
    "NaClO": {
        "氯離子":          0.48,  # NaClO 反應後變 Cl⁻
        "氯鹽":            0.48,
        "Cl":              0.48,
    },
    "NaHSO3": {
        "硫酸根":          0.92,  # 反應後變 SO₄²⁻
        "硫酸鹽":          0.92,
        "SO4":             0.92,
    },
    "Na2S2O5": {
        "硫酸根":          1.00,
        "硫酸鹽":          1.00,
        "SO4":             1.00,
    },

    # 生物補劑 (N / P)
    "尿素": {
        "氨氮":            0.47,  # 尿素 CO(NH₂)₂ 含 47% N
        "氨氮（mg/L）":    0.47,
    },
    "磷酸": {
        "總磷":            0.32,  # H₃PO₄ 含 32% P
    },
}


def get_typical_dosing(std_tank):
    """取得該標準槽體的典型加藥配方。

    回傳: dict[藥劑名 → mg/L 劑量], 若無資料回 {}。
    """
    return TYPICAL_DOSING.get(std_tank, {}) or {}


def get_conversion_factor(chem, item):
    """取得某藥劑對某水質項目的轉換係數。

    回傳: 0~1 之間的 float (kg 引入項目 / kg 藥劑原料)。
    無資料 (= 不引入該項目) 回 0。
    """
    return CONVERSION.get(chem, {}).get(item, 0.0)


def list_dosing_chemicals(std_tank):
    """列出該標準槽體會加哪些藥劑 (用於 finding 描述)。"""
    return list(get_typical_dosing(std_tank).keys())


def compute_chemical_mass(unit, item, q_cmd=None):
    """計算該單元因加藥而引入指定水質項目的質量 (kg/d)。

    Args:
        unit: 處理單元 dict (有 std_tank 欄)
        item: 水質項目名 (例: "懸浮固體（mg/L）")
        q_cmd: 進流總 Q (m³/d)。若 None, 嘗試從 unit['stream_q'] 反推。

    Returns:
        kg/d (float), 若無加藥配方或無對應轉換係數則 0.0。

    計算公式:
        kg/d = Σ_chemicals (劑量_mg/L × Q_m³/d / 1000 × 轉換係數)

    例:
        快混槽 加 PAC 100 mg/L, Q = 23673 m³/d
        → PAC 引入 SS = 100 × 23673 / 1000 × 0.30 = 710.19 kg/d
    """
    std_tank = unit.get("std_tank") if isinstance(unit, dict) else None
    if not std_tank:
        return 0.0

    dosing = get_typical_dosing(std_tank)
    if not dosing:
        return 0.0

    # 取 Q
    if q_cmd is None:
        # 從 stream_q 加總所有 WTB (進流) Q
        sq = unit.get("stream_q") or {}
        q_cmd = 0.0
        for stream_code, info in sq.items():
            if stream_code.startswith("WTB") and isinstance(info, dict):
                q = info.get("q_cmd")
                if isinstance(q, (int, float)):
                    q_cmd += float(q)
    if not q_cmd or q_cmd <= 0:
        return 0.0

    total = 0.0
    for chem, mg_per_L in dosing.items():
        conv = get_conversion_factor(chem, item)
        if conv <= 0:
            continue
        # mg/L × m³/d / 1000 = kg/d 藥劑原料量 × 轉換係數 = kg/d 引入項目量
        total += float(mg_per_L) * float(q_cmd) / 1000.0 * float(conv)
    return total


def describe_chemical_contribution(unit, item, q_cmd=None):
    """描述加藥引入的細節, 給 finding 描述用。

    回傳: str (可空字串), 例:
        "[加藥估算: PAC 100mg/L × Q 23673 × 0.3 = +710 kg/d SS]"
    """
    std_tank = unit.get("std_tank") if isinstance(unit, dict) else None
    if not std_tank:
        return ""
    dosing = get_typical_dosing(std_tank)
    if not dosing:
        return ""
    if q_cmd is None:
        sq = unit.get("stream_q") or {}
        q_cmd = 0.0
        for stream_code, info in sq.items():
            if stream_code.startswith("WTB") and isinstance(info, dict):
                q = info.get("q_cmd")
                if isinstance(q, (int, float)):
                    q_cmd += float(q)
    if not q_cmd or q_cmd <= 0:
        return ""

    parts = []
    total = 0.0
    for chem, mg_per_L in dosing.items():
        conv = get_conversion_factor(chem, item)
        if conv <= 0:
            continue
        added = mg_per_L * q_cmd / 1000.0 * conv
        if added <= 0:
            continue
        parts.append(f"{chem} {mg_per_L}mg/L×{conv}={added:.1f}")
        total += added

    if not parts:
        return ""
    return f"[加藥引入估算 {item}: {', '.join(parts)} → +{total:.1f} kg/d]"


# ──────────────────────────────────────────────────
# 自我測試 (執行此檔可直接驗證)
# ──────────────────────────────────────────────────
def _self_test():
    # Test 1: 快混槽加 PAC 100 mg/L, Q = 23673 m³/d
    fake_unit = {
        "std_tank": "快混槽",
        "stream_q": {"WTB01-01-1": {"q_cmd": 23673}},
    }
    ss_added = compute_chemical_mass(fake_unit, "懸浮固體（mg/L）")
    expected = 100 * 23673 / 1000 * 0.30  # = 710.19
    print(f"Test 1 (快混槽 PAC → SS): {ss_added:.2f} kg/d (預期 {expected:.2f})")
    assert abs(ss_added - expected) < 0.01

    # Test 2: 純 pH 槽不引入 SS
    ph_unit = {"std_tank": "pH調整槽", "stream_q": {"WTB": {"q_cmd": 1000}}}
    ss = compute_chemical_mass(ph_unit, "懸浮固體（mg/L）")
    print(f"Test 2 (pH調整槽 NaOH/H2SO4 → SS): {ss:.2f} kg/d (預期 0)")
    assert ss == 0.0

    # Test 3: pH 槽會引入 SO4 (從 H2SO4)
    so4 = compute_chemical_mass(ph_unit, "硫酸根")
    expected = 30 * 1000 / 1000 * 0.98  # = 29.4
    print(f"Test 3 (pH調整槽 H2SO4 → SO4): {so4:.2f} kg/d (預期 {expected:.2f})")
    assert abs(so4 - expected) < 0.01

    # Test 4: 描述
    desc = describe_chemical_contribution(fake_unit, "懸浮固體（mg/L）")
    print(f"Test 4 描述: {desc}")
    assert "PAC" in desc and "+710.2" in desc

    # Test 5: 沒有加藥配方的槽 (放流池)
    pool_unit = {"std_tank": "放流池", "stream_q": {"WTB": {"q_cmd": 1000}}}
    ss = compute_chemical_mass(pool_unit, "懸浮固體（mg/L）")
    print(f"Test 5 (放流池 不加藥 → SS): {ss:.2f} kg/d (預期 0)")
    assert ss == 0.0

    print("\n[OK] 全部自我測試通過")


if __name__ == "__main__":
    _self_test()
