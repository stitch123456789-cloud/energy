import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 介面標題 ---
st.title("💡 設備系統資料庫")
st.info("此模組將自動抓取『表九（照明）』與『表十（空調）』等數據並生成 Word 表格。")

# --- 2. 數據抓取函數 (範例：照明系統) ---
def fetch_lighting_data(file):
    try:
        # 假設分頁名稱包含 "九"
        all_sheets = pd.ExcelFile(file).sheet_names
        sheet_name = [s for s in all_sheets if "九" in s][0]
        df = pd.read_excel(file, sheet_name=sheet_name, skipfooter=1)
        # 這裡需要根據你的 Excel 欄位名稱做對應
        # 返回一個串列方便 Word 生成
        return df.to_dict('records')
    except:
        return []

# --- 3. Word 表格生成邏輯 ---
def add_lighting_table(doc, data):
    doc.add_paragraph("2. 照明系統：", style='List Number')
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    # 設定表頭
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '燈具種類'
    hdr_cells[1].text = '燈具形式(規格/數量)'
    hdr_cells[2].text = '數量'
    hdr_cells[3].text = '運轉時數(小時/年)'
    
    # 填充數據 (範例循環)
    for item in data:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item.get('種類', ''))
        row_cells[1].text = str(item.get('規格', ''))
        row_cells[2].text = str(item.get('數量', ''))
        row_cells[3].text = str(item.get('時數', ''))

# --- 4. 主程式邏輯 ---
if st.session_state.get('global_excel'):
    file = st.session_state['global_excel']
    
    # 這裡放你的展示邏輯 (像是照明、空調的預覽)
    st.subheader("🔦 照明系統預覽")
    lighting_data = fetch_lighting_data(file)
    if lighting_data:
        st.write(pd.DataFrame(lighting_data))
    
    # 生成按鈕
    if st.button("📝 生成設備資料 Word 報告"):
        doc = Document()
        # 呼叫各個表格生成函數
        add_lighting_table(doc, lighting_data)
        
        # 儲存與打包邏輯 (同之前 p2)
        # ...
else:
    st.warning("請先在側邊欄上傳 Excel 檔案。")
