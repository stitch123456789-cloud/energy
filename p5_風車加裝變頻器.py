import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心格式工具 ---

def set_table_border(table):
    """手動為表格添加黑色框線，防止格子不見"""
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') 
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

def fix_cell_font(cell, size=10, is_bold=False):
    """表格內容強制格式化為標楷體"""
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace_keep_style(doc, data_map):
    """強化版替換邏輯：保留紅色、粗體等原始格式"""
    # 處理段落
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                # 遍歷段落中的每個 Run，只改文字，不改屬性
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
    
    # 處理現有表格（例如看板表格）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        if key in p.text:
                            for run in p.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(val))

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0)
    elec_val = st.session_state.get('auto_avg_price', 3.5)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    rt_info = st.text_input("容量 (RT)", value="1500RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = motor_hp * 0.746 
    details = []
    total_old, total_new = 0, 0
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h 
        details.append({
            "季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%",
            "舊": o_kwh, "新": n_kwh, "省": o_kwh - n_kwh
        })
        total_old += o_kwh
        total_new += n_kwh
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    save_rate = (save_kwh / total_old * 100) if total_old > 0 else 0
    return {
        "old_total": total_old, "save_kwh": save_kwh, "save_money": save_money, 
        "payback": payback, "save_rate": save_rate, "details": details
    }

# --- 4. 生成按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告 (完整整合版)", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        doc = Document("template_p5.docx")
        
        # 精準標籤對齊
        data_map = {
            "{{UN}}": unit_name, "{{COUNT}}": "2", "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info, "{{MT}}": f"三台 {int(motor_hp)}hp",
            "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{results['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_RATE}}": f"{results['save_rate']:.2f}",
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}",
            "{{SUPPRESS_KW}}": "13",
            "{{13}}": "13"
        }
        
        # A. 執行保留格式的文字替換
        safe_replace_keep_style(doc, data_map)
        
        # B. 清除標籤定位符號
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text: p.text = ""
            if "[[NEW_TABLE]]" in p.text: p.text = ""

        # 3. 在文末生成【橫向】表格 (完全對齊您截圖的格式)
        doc.add_page_break()
        doc.add_paragraph("--- 自動生成的橫向表格 (請剪下並貼至指定位置) ---")

        # 準備資料：馬力轉換為 kW
        base_kw = motor_hp * 0.746
        
        # 欄數：1(標題欄) + 資料筆數 (春秋/夏/冬) + 1(合計欄)
        num_cols = len(results['details']) + 2
        
        # 建立一個 7 橫列的表格
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫第一列：編號 (合併中間儲存格)
        fix_cell_font(table.cell(0, 0).text = "編號", is_bold=True)
        # 合併中間格子並填入 CT-1
        merged_header = table.cell(0, 1)
        for i in range(2, num_cols - 1):
            merged_header.merge(table.cell(0, i))
        merged_header.text = ch_info # 使用您輸入的主機編號
        fix_cell_font(merged_header, is_bold=True)
        # 最後一格填合計
        table.cell(0, num_cols-1).text = "合計"
        fix_cell_font(table.cell(0, num_cols-1), is_bold=True)

        # 定義左側標題 labels
        labels = [
            "水塔散熱噸數(RT)", 
            "額定馬力(hp)", 
            "實際耗功(kW)", 
            "全年使用時數(hr)", 
            "負載率(%)", 
            "全年耗電(kWh)"
        ]

        # 逐列填寫資料 (從 Row 1 到 Row 6)
        for r_idx, label in enumerate(labels, start=1):
            # 填寫左邊標題
            table.cell(r_idx, 0).text = label
            fix_cell_font(table.cell(r_idx, 0), is_bold=True)
            
            row_total = 0
            # 填寫中間各季節資料
            for c_idx, d in enumerate(results['details'], start=1):
                cell = table.cell(r_idx, c_idx)
                if r_idx == 1: cell.text = str(rt_info)
                elif r_idx == 2: cell.text = f"{motor_hp:.1f}"
                elif r_idx == 3: cell.text = f"{base_kw:.1f}"
                elif r_idx == 4: cell.text = f"{d['時數']:,.0f}"
                elif r_idx == 5: cell.text = "100%"
                elif r_idx == 6: 
                    cell.text = f"{d['舊']:,.0f}"
                    row_total += d['舊']
                fix_cell_font(cell, is_bold=(r_idx==6)) # 最後一列加粗
            
            # 填寫右邊合計
            last_cell = table.cell(r_idx, num_cols-1)
            if r_idx == 3: last_cell.text = f"{base_kw * len(results['details']):.1f}"
            elif r_idx == 6: last_cell.text = f"{results['old_total']:,.0f}"
            else: last_cell.text = "" # 其他格留空
            fix_cell_font(last_cell, is_bold=True)
