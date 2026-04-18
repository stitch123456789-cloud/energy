import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 工具函數：字體設定 ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False

def set_font_kai_bold_14(run):
    # 設定標楷體 14號 加粗
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(14)
    run.font.bold = True

# --- 2. 數據抓取與自動加總邏輯 ---
def fetch_and_aggregate_lighting(file):
    try:
        xl = pd.ExcelFile(file)
        # 模糊搜尋所有包含「表九之二」的分頁
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        aggregated_data = {}

        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 數據從 index 6 開始讀取
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()   # B欄
                spec = str(df.iloc[i, 5]).strip()   # F欄
                count_str = str(df.iloc[i, 9]).strip()  # J欄
                hours_str = str(df.iloc[i, 11]).strip() # L欄
                
                if kind == "nan" or "註" in kind or "合計" in kind: continue
                if spec == "nan" or spec == "": continue
                
                # 清洗種類名稱：移除「1. 」等數字開頭
                if '.' in kind:
                    kind = kind.split('.')[-1].strip()
                
                try:
                    count = int(float(count_str.replace(',', '')))
                    hours = int(float(hours_str.replace(',', '')))
                except:
                    continue

                # 以 (種類, 規格, 時數) 為 Key 進行加總
                key = (kind, spec, hours)
                aggregated_data[key] = aggregated_data.get(key, 0) + count
                    
        return aggregated_data
    except Exception as e:
        return None

# --- 3. Word 生成邏輯 ---
def get_lighting_word_bytes(aggregated_data):
    doc = Document()
    
    # 標題：2.照明系統： (標楷體 14號 加粗)
    p = doc.add_paragraph()
    run_title = p.add_run("2.照明系統：")
    set_font_kai_bold_14(run_title)

    # 建立 2層表頭表格
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    table.cell(0, 1).merge(table.cell(0, 2)) # 橫向合併：燈具形式
    table.cell(0, 0).merge(table.cell(1, 0)) # 縱向合併：種類
    table.cell(0, 3).merge(table.cell(1, 3)) # 縱向合併：時數
    
    headers = [
        (table.cell(0, 0), "燈具種類"),
        (table.cell(0, 1), "燈具形式"),
        (table.cell(0, 3), "運轉時數(小時/年)"),
        (table.cell(1, 1), "容量規格"),
        (table.cell(1, 2), "數量")
    ]

    for cell, text in headers:
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cell = cell.paragraphs[0]
        p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_cell.add_run(text)
        set_font_kai_11(r)

    # 填入數據 (按種類排序)
    sorted_items = sorted(aggregated_data.items(), key=lambda x: x[0][0])

    for (kind, spec, hours), count in sorted_items:
        row_cells = table.add_row().cells
        row_data = [kind, spec, str(count), str(hours)]
        for idx, val in enumerate(row_data):
            cp = row_cells[idx].paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cp.add_run(val)
            set_font_kai_11(r)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# --- 4. Streamlit 介面渲染 (自動掃描) ---
st.subheader("⚙️ 設備系統資料庫")

# 獲取檔案
up_file = st.file_uploader("請上傳單張或整份能源查核 Excel", type=["xlsx"])
final_file = up_file if up_file else st.session_state.get('global_excel')

if final_file:
    # --- 背景自動掃描 ---
    agg_data = fetch_and_aggregate_lighting(final_file)
    
    if agg_data:
        # --- 唯一的按鍵：生成並下載 ---
        word_data = get_lighting_word_bytes(agg_data)
        
        st.download_button(
            label="🚀 生成並下載照明系統 Word 報告",
            data=word_data,
            file_name="照明系統總表.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True # 讓按鈕寬一點比較專業
        )
        
        # 同步存入左側打包中心
        if 'report_warehouse' not in st.session_state:
            st.session_state['report_warehouse'] = {}
        st.session_state['report_warehouse']['照明系統'] = word_data
        
    else:
        st.error("查無符合的分頁內容（需含『表九之二』）。")
else:
    st.info("💡 請先上傳 Excel 檔案以啟用報告生成功能。")
