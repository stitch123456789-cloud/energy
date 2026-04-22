import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 通用工具函數 ---
def set_table_border(table):
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') 
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

def fix_cell_font(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 
        if not paragraph.runs: paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold

# --- 2. 介面設計 ---
st.set_page_config(layout="wide")
st.title("💡 P6. 照明系統節能效益分析")

# 基礎參數
c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    elec_price = st.number_input("平均電費 (元/度)", value=3.5)
with c2:
    report_date = st.text_input("報告日期", value="2024/05/20")
    wage_cost = st.number_input("平均每盞施工費", value=150)
with c3:
    manager_name = st.text_input("專案負責人", value="工程部")

# 區域與燈具數據 (對應截圖中的 Excel 欄位)
if "lighting_data" not in st.session_state:
    st.session_state.lighting_data = [
        {"area": "地下室停車場", "type": "T8 20W*2", "old_w": 46, "qty": 150, "hr": 8760, "led_w": 18, "led_price": 300},
        {"area": "1F 辦公室", "type": "T5 28W*3", "old_w": 95, "qty": 80, "hr": 3000, "led_w": 40, "led_price": 1200}
    ]

st.subheader("📝 燈具汰換數據明細")
# 使用 data_editor 模擬 Excel 操作感
df_lighting = st.data_editor(pd.DataFrame(st.session_state.lighting_data), num_rows="dynamic", use_container_width=True)

# --- 3. 核心邏輯與報告生成 ---
if st.button("🚀 生成 P6 照明節能報告", use_container_width=True):
    try:
        doc = Document() # 若有範本可改為 Document("template_p6.docx")
        
        # 標題
        title = doc.add_heading('照明系統 LED 汰換效益分析表', 0)
        
        # 計算總體數據
        total_invest = 0
        total_save_kwh = 0
        
        # 建立表格 (對齊截圖中的 Word 樣式)
        # 欄位：區域 | 燈具形式 | 數量 | 時數 | 改善前(W/kwh) | 改善後(W/kwh) | 節電量 | 投資額
        table = doc.add_table(rows=1, cols=10)
        table.style = 'Table Grid'
        set_table_border(table)
        
        headers = ["區域", "原燈具形式", "數量", "年時數", "原單盞(W)", "改善後(W)", "原耗電(kWh)", "新耗電(kWh)", "節電量", "投資額"]
        for i, h in enumerate(headers):
            table.cell(0, i).text = h
            fix_cell_font(table.cell(0, i), is_bold=True)

        # 填入數據列
        for _, row in df_lighting.iterrows():
            cells = table.add_row().cells
            
            # 計算該區域數值
            old_kwh = (row['old_w'] * row['qty'] * row['hr']) / 1000
            new_kwh = (row['led_w'] * row['qty'] * row['hr']) / 1000
            save_kwh = old_kwh - new_kwh
            invest = (row['led_price'] + wage_cost) * row['qty']
            
            total_invest += invest
            total_save_kwh += save_kwh
            
            cells[0].text = str(row['area'])
            cells[1].text = str(row['type'])
            cells[2].text = str(row['qty'])
            cells[3].text = str(row['hr'])
            cells[4].text = str(row['old_w'])
            cells[5].text = str(row['led_w'])
            cells[6].text = f"{old_kwh:,.0f}"
            cells[7].text = f"{new_kwh:,.0f}"
            cells[8].text = f"{save_kwh:,.0f}"
            cells[9].text = f"{invest:,.0f}"
            
            for c in cells: fix_cell_font(c)

        # 合計列
        last_row = table.add_row().cells
        last_row[0].text = "合計"
        last_row[8].text = f"{total_save_kwh:,.0f}"
        last_row[9].text = f"{total_invest:,.0f}"
        for i in [0, 8, 9]: fix_cell_font(last_row[i], is_bold=True)

        # 效益總結文字
        save_money = total_save_kwh * elec_price / 10000 # 萬元
        payback = total_invest / (save_money * 10000) if save_money > 0 else 0
        
        doc.add_paragraph(f"\n1. 預估年節電量：{total_save_kwh:,.0f} kWh")
        doc.add_paragraph(f"2. 預估年節省電費：{save_money:.2f} 萬元")
        doc.add_paragraph(f"3. 總投資金額：{total_invest:,.0f} 元")
        doc.add_paragraph(f"4. 回收年限：約 {payback:.1f} 年")

        # 下載
        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ P6 報告生成成功！")
        st.download_button("📥 下載 P6 照明報告", buf.getvalue(), f"照明節能分析_{unit_name}.docx")

    except Exception as e:
        st.error(f"生成報告時出錯：{e}")
