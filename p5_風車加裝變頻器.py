import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心格式工具 ---

def set_table_border(table):
    """手動強制繪製表格黑色框線，確保表格不會隱形"""
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

def fix_cell_format(cell, is_bold=False, size=10, align_center=True):
    """統一格式化表格內容：標楷體、字體大小、置中"""
    for paragraph in cell.paragraphs:
        if align_center:
            paragraph.alignment = 1 # 置中
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """替換段落與表格中的標籤"""
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        if key in p.text:
                            for run in p.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(val))
                                    run.font.color.rgb = RGBColor(0, 0, 0)
                                    run.font.name = '標楷體'
                                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=15.0)
    elec_input = st.number_input("平均電費 (元/度)", value=4.45, step=0.01)
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="300RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("運轉說明", value="僅開啟一台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "運轉時數(hr)": [4380, 3285, 2190],
        "平均負載率(%)": [100, 100, 100]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = motor_hp * 0.746 
    details = []
    total_old = 0
    for _, row in df.iterrows():
        # 使用字串 key 讀取，確保與 DataFrame 欄位一致
        h = float(row["運轉時數(hr)"])
        l = float(row["平均負載率(%)"]) / 100
        o_kwh = base_kw * h
        details.append({
            "季節": row["季節"], "時數": h, "負載": f"{row['平均負載率(%)']}%",
            "舊": o_kwh
        })
        total_old += o_kwh
    return {
        "old_total": total_old, "details": details
    }

# --- 4. 生成按鈕 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    res = run_calculation(current_op_df)
    save_money = (res['old_total'] * 0.3) * elec_input / 10000 # 假設節省 30%
    payback = invest_amt / save_money if save_money > 0 else 0
    
    try:
        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name, "{{CH_INFO}}": ch_info, "{{RT_INFO}}": rt_info,
            "{{MT}}": f"{int(motor_hp)}hp", "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{res['old_total']:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}"
        }
        
        # 1. 文字替換
        safe_replace(doc, data_map)

        # 2. 在文末生成橫向表格 (對齊截圖格式)
        doc.add_page_break()
        doc.add_paragraph("【表一、現況耗電明細分析表】")
        
        # 欄數：標題欄 + 季節資料欄 + 合計欄
        num_cols = len(res['details']) + 2
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        side_headers = [
            "編號", "水塔散熱頓數(RT)", "額定馬力(hp)", "實際耗功(kW)", 
            "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"
        ]
        
        for i, text in enumerate(side_headers):
            cell = table.cell(i, 0)
            cell.text = text
            fix_cell_format(cell, is_bold=True)

        # 填寫季節數據與合計
        total_kw = 0
        total_kwh = 0
        
        # 設備編號橫跨
        merged_header = table.cell(0, 1)
        for i in range(2, num_cols - 1):
            merged_header.merge(table.cell(0, i))
        merged_header.text = "CT-1" # 範例編號
        fix_cell_format(merged_header, is_bold=True)
        
        table.cell(0, num_cols-1).text = "合計"
        fix_cell_format(table.cell(0, num_cols-1), is_bold=True)

        for idx, d in enumerate(res['details'], start=1):
            base_kw = motor_hp * 0.746
            table.cell(1, idx).text = str(rt_info)
            table.cell(2, idx).text = f"{motor_hp:.1f}"
            table.cell(3, idx).text = f"{base_kw:.1f}"
            table.cell(4, idx).text = f"{d['時數']:,.0f}"
            table.cell(5, idx).text = "100%"
            table.cell(6, idx).text = f"{d['舊']:,.0f}"
            
            total_kw += base_kw
            total_kwh += d['舊']
            
            # 設定格式
            for r in range(1, 7):
                fix_cell_format(table.cell(r, idx), is_bold=(r==6))

        # 填寫合計欄位
        table.cell(3, num_cols-1).text = f"{total_kw:.1f}"
        table.cell(6, num_cols-1).text = f"{total_kwh:,.0f}"
        fix_cell_format(table.cell(3, num_cols-1), is_bold=True)
        fix_cell_format(table.cell(6, num_cols-1), is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！橫向表格已生成於文末。")
        st.download_button("📥 下載修正版 Word 報告", buf.getvalue(), "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"出錯了: {e}")
