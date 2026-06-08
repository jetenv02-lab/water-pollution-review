# -*- coding: utf-8 -*-
"""設計參數體檢 — 算每個單元的 HRT / SOR / G 值, 對照學理範圍。

3 個指標:
    1. HRT (水力停留時間, hr) = 有效容積 / Q
       Q 轉小時: Q (CMD) → m³/hr 要 ÷ 24
       → HRT (hr) = V (m³) / Q (CMD) * 24

    2. SOR (表面溢流率, m³/m²·d) = Q / A_俯視
       A_俯視 (矩形) = 長 × 寬
       A_俯視 (圓池) = π × (直徑/2)²

    3. G 值 (速度梯度, s⁻¹) = √(P / (μ × V))
       P  = 馬達總功率 (W)  (注意: equipment 的「馬力_kW」要 ×1000)
       μ  = 水動黏滯係數 (20°C ≈ 1.002e-3 Pa·s)
       V  = 有效容積 (m³)

斜板偵測:
    若 name_in_doc / equipment / design_params 含「斜板 / Lamella / Tube settler」
    → SOR_max 改用較寬鬆值 (一般 50 → 斜板 120)

跨單元比例:
    慢混 HRT 應 ≈ 快混 HRT × 10 (8~12 倍合理, 5~20 倍 ⚠️, 否則 🔴)
    慢混 HRT 應 ≥ 快混 HRT (絕對學理)
"""
import math
import re

# 20°C 水的動黏滯係數 (Pa·s)
WATER_MU_20C = 1.002e-3


def to_float(v):
    """容錯轉 float, 失敗回 None。"""
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        m = re.search(r"\d+(?:\.\d+)?", str(v))
        if m:
            try:
                return float(m.group())
            except ValueError:
                return None
    return None


def get_volume(unit):
    """取得單元有效容積 (m³)。優先 size.有效容量, 後備從 長×寬×高 / 長×寬×水深 算。"""
    size = unit.get("size") or {}
    v = to_float(size.get("有效容量"))
    if v is not None and v > 0:
        return v
    # 後備: 用尺寸算
    L = to_float(size.get("長/直徑"))
    W = to_float(size.get("寬"))
    h = to_float(size.get("有效水深")) or to_float(size.get("高"))
    if L and h:
        if W and W > 0:
            return L * W * h
        # 圓池: 用直徑算
        # 但 size 表的「長/直徑」可能兩者都用同一欄, 不容易判斷
        # 保守: 矩形若沒寬, 不算
    return None


def get_main_q(unit):
    """取得單元主流量 (CMD)。

    取 stream_q 中最大的 in_mass 對應的 Q (代表「主要進流」)。
    若 stream_q 拿不到, 退用 effluent 最大的 Q。
    """
    stream_q = unit.get("stream_q") or {}
    if not stream_q:
        return None

    # 區分進/出流: 看 influent / effluent 的 key
    in_streams = set((unit.get("influent") or {}).keys())
    out_streams = set((unit.get("effluent") or {}).keys())

    # 主流量 = max(Σ進流 Q, Σ出流 Q), 兩者理論上要相等
    sum_in_q = 0.0
    sum_out_q = 0.0
    for code, qinfo in stream_q.items():
        if not isinstance(qinfo, dict) or not qinfo.get("ok"):
            continue
        q = qinfo.get("q_cmd", 0)
        if code in in_streams:
            sum_in_q += q
        if code in out_streams:
            sum_out_q += q

    if sum_in_q > 0:
        return sum_in_q
    if sum_out_q > 0:
        return sum_out_q
    return None


def get_surface_area(unit):
    """取得池俯視面積 (m²)。

    矩形池: 長 × 寬
    圓池: π × (直徑/2)²  (但 size 表很少明確分長/直徑, 保守只算矩形)
    """
    size = unit.get("size") or {}
    L = to_float(size.get("長/直徑"))
    W = to_float(size.get("寬"))
    if L and W and L > 0 and W > 0:
        return L * W
    # 圓池 fallback: 若名稱含「圓」「塔」或 W 是 0, 用直徑算
    if L and (not W or W == 0):
        name = unit.get("name_in_doc", "")
        if any(kw in name for kw in ["塔", "圓", "槽桶"]):
            return math.pi * (L / 2) ** 2
    return None


def get_motor_power_w(unit):
    """從 equipment 中找攪拌機馬力, 轉成 W。

    equipment 結構: [{"name": "攪拌機", "位置": "...", "數量": 1, "馬力_kW": 0.37}, ...]
    取所有「攪拌機 / 鼓風機」累計, 轉成 W (kW × 1000)。
    """
    total_w = 0.0
    found = False
    for eq in unit.get("equipment") or []:
        if not isinstance(eq, dict):
            continue
        name = str(eq.get("name") or "")
        if not any(kw in name for kw in ["攪拌", "鼓風"]):
            continue
        kw = to_float(eq.get("馬力_kW"))
        qty = to_float(eq.get("數量")) or 1
        if kw is not None and kw > 0:
            total_w += kw * qty * 1000  # kW → W
            found = True
    return total_w if found else None


def detect_lamella(unit):
    """偵測是否為斜板沉澱池。"""
    text_to_check = " ".join([
        str(unit.get("name_in_doc") or ""),
        " ".join(str(e.get("name") or "") for e in (unit.get("equipment") or [])),
        " ".join((unit.get("design_params") or {}).keys()),
    ])
    keywords = ["斜板", "斜管", "Lamella", "lamella", "Tube settler", "tube settler"]
    return any(kw in text_to_check for kw in keywords)


# ─────────────────────────────────────────────
# 主要計算函式
# ─────────────────────────────────────────────

def compute_hrt(unit):
    """算 HRT (小時)。"""
    V = get_volume(unit)
    Q = get_main_q(unit)
    if V is None or Q is None or Q <= 0:
        return None
    return V / Q * 24.0


def compute_sor(unit):
    """算表面溢流率 SOR (m³/m²·d)。"""
    A = get_surface_area(unit)
    Q = get_main_q(unit)
    if A is None or Q is None or A <= 0:
        return None
    return Q / A


def compute_g_value(unit):
    """算 G 值 (s⁻¹)。"""
    V = get_volume(unit)
    P = get_motor_power_w(unit)
    if V is None or P is None or V <= 0 or P <= 0:
        return None
    # G = √(P / (μ × V))
    return math.sqrt(P / (WATER_MU_20C * V))


def compute_all_metrics(unit):
    """一次算 3 個指標 + 斜板偵測, 回傳 dict。"""
    return {
        "hrt_hr": compute_hrt(unit),
        "sor_m3_m2_d": compute_sor(unit),
        "g_value_s_inv": compute_g_value(unit),
        "is_lamella": detect_lamella(unit),
        "volume_m3": get_volume(unit),
        "surface_area_m2": get_surface_area(unit),
        "main_q_cmd": get_main_q(unit),
        "motor_power_w": get_motor_power_w(unit),
    }


# ─────────────────────────────────────────────
# 跨單元: 快/慢混 HRT 比例
# ─────────────────────────────────────────────

def find_fast_slow_pairs(app_data):
    """從 app_data 找「同序列」的快混-慢混配對。

    啟發式: T01-08 (快混) → T01-09 (慢混), 序號相鄰且槽體類型對得上
    """
    units = app_data.get("units", {})
    pairs = []  # [(fast_code, slow_code)]
    # 快混類: 含「快混」的, 也接受「pH 調整暨快混池」(常見合二為一)
    fast_set = {"快混槽", "快混池",
                "pH調整暨快混池", "pH調整池暨快混池",
                "pH調整快混池", "pH調整與快混池"}
    slow_set = {"慢混池", "慢混槽"}

    # 依代號排序, 看相鄰的 fast → slow
    sorted_codes = sorted(units.keys())
    for i, code in enumerate(sorted_codes):
        u = units[code]
        std = u.get("std_tank")
        if std not in fast_set:
            continue
        # 看下一個或下幾個是不是慢混 (允許中間隔 1-2 個)
        for j in range(i + 1, min(i + 5, len(sorted_codes))):
            ncode = sorted_codes[j]
            nu = units[ncode]
            # 序號要同 series (T01-XX → T01-YY)
            if code.split("-")[0] != ncode.split("-")[0]:
                break
            if nu.get("std_tank") in slow_set:
                pairs.append((code, ncode))
                break
    return pairs


def check_fast_slow_ratio(app_data, fast_code, slow_code):
    """檢查 (快混, 慢混) 配對的 HRT 比例。

    Returns: finding dict 或 None
    """
    units = app_data.get("units", {})
    u_fast = units.get(fast_code, {})
    u_slow = units.get(slow_code, {})

    hrt_fast = compute_hrt(u_fast)
    hrt_slow = compute_hrt(u_slow)

    if hrt_fast is None or hrt_slow is None or hrt_fast <= 0:
        return None

    # 違反絕對學理: 慢混 < 快混
    if hrt_slow < hrt_fast:
        return {
            "嚴重度": "不合理",
            "類型": "設計參數",
            "單元": f"{fast_code} → {slow_code}",
            "標準槽體": "快混+慢混",
            "對照項目": "慢混 HRT < 快混 HRT",
            "描述": (
                f"快混 {fast_code} HRT={hrt_fast*60:.1f} 分 ({hrt_fast:.3f} hr), "
                f"慢混 {slow_code} HRT={hrt_slow*60:.1f} 分 ({hrt_slow:.3f} hr)。"
                f"學理: 慢混停留時間應「大於或等於」快混 (絮羽需時間長大)。"
            ),
            "依據": "規則庫 D148 / 環工設計準則",
        }

    # 比例檢查 (你的經驗法則: 慢混 / 快混 應 8~12 倍)
    ratio = hrt_slow / hrt_fast
    if ratio < 5 or ratio > 20:
        return {
            "嚴重度": "不合理",
            "類型": "設計參數",
            "單元": f"{fast_code} → {slow_code}",
            "標準槽體": "快混+慢混",
            "對照項目": "慢混/快混 HRT 比例",
            "描述": (
                f"快混 {fast_code} HRT={hrt_fast*60:.1f} 分, "
                f"慢混 {slow_code} HRT={hrt_slow*60:.1f} 分, "
                f"比例 = {ratio:.1f} 倍 (學理常見 8~12 倍, "
                f"<5 或 >20 表示設計失衡)。"
            ),
            "依據": "業界經驗: 慢混 HRT ≈ 快混 × 10",
        }
    elif ratio < 8 or ratio > 12:
        return {
            "嚴重度": "待人工",
            "類型": "設計參數",
            "單元": f"{fast_code} → {slow_code}",
            "標準槽體": "快混+慢混",
            "對照項目": "慢混/快混 HRT 比例",
            "描述": (
                f"快混 {fast_code} HRT={hrt_fast*60:.1f} 分, "
                f"慢混 {slow_code} HRT={hrt_slow*60:.1f} 分, "
                f"比例 = {ratio:.1f} 倍 (學理理想 8~12 倍, 偏離但仍可接受)。"
            ),
            "依據": "業界經驗: 慢混 HRT ≈ 快混 × 10",
        }
    return None  # 比例正常


if __name__ == "__main__":
    # 自我測試
    import json, sys, io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass

    with open('application_05 申請資料(秋棠)(1150119)final.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 跑全部單元
    print(f"{'單元':<10} {'類型':<14} {'HRT (hr)':<12} {'SOR (m/d)':<12} {'G (1/s)':<10}")
    print("-" * 70)
    for code, u in sorted(data['units'].items()):
        m = compute_all_metrics(u)
        std = u.get('std_tank') or '?'
        hrt = f"{m['hrt_hr']:.2f}" if m['hrt_hr'] else '-'
        sor = f"{m['sor_m3_m2_d']:.1f}" if m['sor_m3_m2_d'] else '-'
        g = f"{m['g_value_s_inv']:.0f}" if m['g_value_s_inv'] else '-'
        print(f"{code:<10} {std[:12]:<14} {hrt:<12} {sor:<12} {g:<10}")
