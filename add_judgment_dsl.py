# -*- coding: utf-8 -*-
"""一次性腳本: 在 rules_extracted.csv 新增 `_可比對判斷式` 欄位, 並填 10 個範例。

DSL 規範請見 RULE_AUTHORING.md。基本語法:
    IF <condition> THEN FLAG "<訊息>"

變數 (可從 step2_extract_v2 / OCR 結果取得):
    unit.std_tank             標準槽體名稱 (例 "活性碳吸附塔")
    unit.name_in_doc          申請文件填的槽體名稱
    unit.code                 T01-05
    unit.influent[X].濃度     進流水質
    unit.effluent[X].濃度     出流水質
    unit.equipment[]          機具清單 (字串陣列)
    unit.design.<param>       設計參數 (滯留時間, 表面溢流率...)
    unit.dose                 加藥資訊
    unit.purpose              處理用途 (字串)

執行: python add_judgment_dsl.py
輸出: 直接覆寫 rules_extracted.csv (備份成 rules_extracted.csv.bak)
"""
import csv
import os
import shutil

CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules_extracted.csv")
NEW_COL = "_可比對判斷式"

# ─────────────────────────────────────────────────
# 10 個範例 (對應實際缺失 ID)
# ─────────────────────────────────────────────────
EXAMPLES = {
    # D001: 活性碳吸附裝置反洗水來源未標示
    "D001": 'IF unit.std_tank IN ["活性碳吸附塔","砂濾塔","多介質過濾器"] AND unit.equipment CONTAINS "反洗" AND NOT flow_graph.has_upstream(unit, label="反洗水") THEN FLAG "反洗水來源未標示於流向圖"',

    # D002: 批次反應槽處理量無法達成
    "D002": 'IF unit.std_tank == "批次反應槽" AND unit.design.daily_capacity > unit.design.batch_volume * unit.design.batches_per_day THEN FLAG "單日處理量超出批次容量×日批次數"',

    # D004: 活性碳吸附裝置缺更換頻率
    "D004": 'IF unit.std_tank IN ["活性碳吸附塔","離子交換樹脂塔"] AND unit.design.replacement_frequency IS NULL THEN FLAG "吸附/離子交換設施缺更換頻率"',

    # D005: 批次反應槽 pH 設定不符學理 (重金屬去除)
    "D005": 'IF unit.std_tank == "批次反應槽" AND unit.purpose CONTAINS "重金屬" AND (unit.design.pH < 8.5 OR unit.design.pH > 11.0) THEN FLAG "重金屬沉澱 pH 應落在 8.5~11.0"',

    # D015: 砂濾塔反洗水未納入總水量
    "D015": 'IF unit.std_tank IN ["砂濾塔","多介質過濾器"] AND unit.equipment CONTAINS "反洗" AND NOT mass_balance.includes(unit, stream="反洗水") THEN FLAG "反洗水量未計入質量平衡"',

    # D018: 砂濾塔未登載反洗馬達
    "D018": 'IF unit.std_tank IN ["砂濾塔","活性碳吸附塔","多介質過濾器"] AND unit.equipment CONTAINS "反洗" AND NOT unit.equipment CONTAINS_ANY ["反洗馬達","反洗泵","反洗風機"] THEN FLAG "登載反洗但未列反洗動力機具"',

    # D021a: 快混槽不該展現金屬去除率 (應該是慢混+沉澱才會)
    "D021a": 'IF unit.std_tank == "快混槽" AND removal_rate(unit, "重金屬") > 0.10 THEN FLAG "快混槽 (僅加藥混合) 不應展現顯著重金屬去除率"',

    # D025: 沉澱池有污泥流向但未登載污泥泵
    "D025": 'IF unit.std_tank IN ["沉澱池","澄清池","化學沉澱池"] AND flow_graph.has_downstream(unit, label="污泥") AND NOT unit.equipment CONTAINS_ANY ["污泥泵","刮泥機","螺旋輸送機"] THEN FLAG "沉澱池有污泥流向但未列污泥輸送機具"',

    # D026a: 快混槽流向不能雙向
    "D026a": 'IF unit.std_tank == "快混槽" AND flow_graph.is_bidirectional(unit) THEN FLAG "流向示意圖標示為雙向, 應為單向"',

    # D011: 放流池鋅濃度數值明顯錯誤
    "D011": 'IF unit.std_tank == "放流池" AND unit.effluent.鋅.濃度 > unit.influent.鋅.濃度 * 1.05 THEN FLAG "放流池鋅出流濃度高於進流, 數值疑似誤植"',
}


def main():
    # 讀入
    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = list(reader.fieldnames)

    if NEW_COL in fieldnames:
        print(f"欄位 `{NEW_COL}` 已存在 — 只更新範例值, 不動其他列。")
    else:
        fieldnames.append(NEW_COL)
        print(f"新增欄位 `{NEW_COL}`。")

    # 備份
    bak = CSV_PATH + ".bak"
    if not os.path.exists(bak):
        shutil.copy2(CSV_PATH, bak)
        print(f"備份: {bak}")

    # 填值
    filled = 0
    for row in rows:
        if NEW_COL not in row:
            row[NEW_COL] = ""
        rid = row.get("缺失ID", "").strip()
        if rid in EXAMPLES:
            row[NEW_COL] = EXAMPLES[rid]
            filled += 1

    # 寫回 (UTF-8 BOM, Excel 直接打得開)
    with open(CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"完成。填入範例 {filled} 筆 / 總共 {len(rows)} 筆。")
    print()
    print("=== 範例對照 ===")
    for rid, dsl in EXAMPLES.items():
        print(f"{rid}: {dsl[:90]}...")


if __name__ == "__main__":
    main()
