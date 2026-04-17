import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 工具函數：設定標楷體 11 號 (不加粗) ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False  # 確保不加粗

# --- 1. 數據抓取核心 ---
def fetch_lighting_data(file):
    try:
        xl = pd.ExcelFile(file)
        # 尋找所有名稱包含 "表九之二" 的分頁
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        all_buildings_results = {}

        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 提取建築物名稱 (從分頁名稱抓取較準確)
            b_name = sheet.split('、')[-1] if '、' in sheet else sheet
            
            lighting_items = []
            # 數據從第 7 列開始 (index 6)
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()   # B欄:種類
                spec = str(df.iloc[i, 5]).strip()   # F欄:容量規格
                count = str(df.iloc[i, 9]).strip()  # J欄:數量
                hours = str(df.iloc[i, 11]).strip() # L欄:運轉時數
                
                # --- 數據清洗邏輯 ---
                # 1. 過濾掉空值 (nan)
                if kind == "nan" or spec == "nan":
                    continue
                # 2. 過濾掉包含「註」或「合計」的行
                if "註" in kind or "合計" in kind or "合計" in spec:
                    continue
                
                lighting_items.append({
                    "kind": kind,
                    "spec": spec,
                    "count": count,
                    "hours": hours
                })
            
            if lighting_items:
                all_buildings_results[b_name] = lighting_items
                
        return all_buildings_results
    except Exception as e:
        st.error(f"抓取照明資料失敗：{e}")
        return None

# --- 2. Word 生成邏輯 ---
def generate_lighting_word(data):
    doc = Document()
    
    # 標題：2.照明系統 (11號標楷體，不加粗)
    p = doc.add_paragraph()
    run = p.add_run("2.照明系統：")
    set_font_kai_11(run)

    for b_name, items in data.items():
        # 建築物子名稱 (例如：建築物 編號-1)
        sub_p = doc.add_paragraph()
        run_sub = sub_p.add_run(f"({b_name})")
        set_font_kai_11(run_sub)

        # 建立表格 (與圖片一致的結構)
        # 圖片中表頭分為兩層，這裡用標準表格模擬
        table = doc.add_table(rows=2, cols=4)
        table.style = 'Table Grid'
        
        # --- 處理複雜表頭 (合併儲存格模擬圖片) ---
        # 合併第一列的第2,3欄 (燈具形式)
        hdr_cell_form = table.cell(0, 1).merge(table.cell(0, 2))
        
        # 填寫第一層表頭
        headers_top = [
            (table.cell(0, 0), "燈具種類"),
            (hdr_cell_form, "燈具形式"),
            (table.cell(0, 3), "運轉時數(小時/年)")
        ]
        # 填寫第二層表頭
        headers_bottom = [
            (table.cell(1, 1), "容量規格"),
            (table.cell(1, 2), "數量")
        ]
        
        # 合併第一欄與第四欄的垂直儲存格 (讓標題置中)
        table.cell(0, 0).merge(table.cell(1, 0))
        table.cell(0, 3).merge(table.cell(1, 3))

        # 統一設定表頭文字格式
        all_header_cells = [
            (table.cell(0, 0), "燈具種類"),
            (hdr_cell_form, "燈具形式"),
            (table.cell(0, 3), "運轉時數(小時/年)"),
            (table.cell(1, 1), "容量規格"),
            (table.cell(1, 2), "數量")
        ]

        for cell, text in all_header_cells:
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_font_kai_11(run)

        # --- 填入數據內容 ---
        for item in items:
            row_cells = table.add_row().cells
            # 依序：種類, 規格, 數量, 時數
            row_data = [item['kind'], item['spec'], item['count'], item['hours']]
            for idx, val in enumerate(row_data):
                cell_p = row_cells[idx].paragraphs[0]
                cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # 去除數字後的 .0 
                clean_val = str(val).replace('.0', '')
                run = cell_p.add_run(clean_val)
                set_font_kai_11(run)
                
    return doc

# --- 3. Streamlit 介面 ---
if st.session_state.get('global_excel'):
    if st.button("🚀 生成照明系統表格報告"):
        data = fetch_lighting_data(st.session_state['global_excel'])
        if data:
            doc = generate_lighting_word(data)
            buffer = io.BytesIO()
            doc.save(buffer)
            st.download_button(
                label="📥 下載照明系統報告",
                data=buffer.getvalue(),
                file_name="2_照明系統.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("✅ 報告已生成，格式已調整為標楷體11號、不加粗、置中。")
