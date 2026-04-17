import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 工具函數：設定標楷體 11 號 (不加粗) ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False

# --- 2. 數據抓取核心：模糊匹配分頁 ---
def fetch_lighting_data_flexible(file):
    try:
        xl = pd.ExcelFile(file)
        # 💡 模糊功能：只要名稱包含 "表九之二" 就算數
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        all_buildings_results = {}

        for sheet in target_sheets:
            # 讀取該分頁
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 從分頁名稱提取建築物標題 (例如從 "表九之二、建築物 編號-1" 提取 "建築物 編號-1")
            display_name = sheet.split('、')[-1] if '、' in sheet else sheet
            
            lighting_items = []
            # 數據通常從第 7 列開始 (index 6)
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()   # B欄: 種類
                spec = str(df.iloc[i, 5]).strip()   # F欄: 容量規格
                count = str(df.iloc[i, 9]).strip()  # J欄: 數量
                hours = str(df.iloc[i, 11]).strip() # L欄: 運轉時數
                
                # --- 數據清洗 ---
                if kind == "nan" or "註" in kind or "合計" in kind:
                    continue
                
                # 如果規格是空的也跳過
                if spec == "nan" or spec == "":
                    continue

                lighting_items.append({
                    "kind": kind,
                    "spec": spec,
                    "count": count,
                    "hours": hours
                })
            
            # 只有當該建築物有抓到資料時才放入結果
            if lighting_items:
                all_buildings_results[display_name] = lighting_items
                
        return all_buildings_results
    except Exception as e:
        st.error(f"模糊抓取照明資料失敗：{e}")
        return None

# --- 3. Word 生成邏輯 (保持原樣，確保格式一致) ---
def generate_lighting_word(data):
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("2.照明系統：")
    set_font_kai_11(run)

    for b_name, items in data.items():
        sub_p = doc.add_paragraph()
        run_sub = sub_p.add_run(f"({b_name})")
        set_font_kai_11(run_sub)

        # 建立兩層表頭表格 (4欄)
        table = doc.add_table(rows=2, cols=4)
        table.style = 'Table Grid'
        
        # 合併儲存格還原圖片格式
        table.cell(0, 1).merge(table.cell(0, 2)) # 燈具形式
        table.cell(0, 0).merge(table.cell(1, 0)) # 種類垂直合併
        table.cell(0, 3).merge(table.cell(1, 3)) # 時數垂直合併

        header_cells = [
            (table.cell(0, 0), "燈具種類"),
            (table.cell(0, 1), "燈具形式"),
            (table.cell(0, 3), "運轉時數(小時/年)"),
            (table.cell(1, 1), "容量規格"),
            (table.cell(1, 2), "數量")
        ]

        for cell, text in header_cells:
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cell = cell.paragraphs[0]
            p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_cell.add_run(text)
            set_font_kai_11(r)

        for item in items:
            row_cells = table.add_row().cells
            row_data = [item['kind'], item['spec'], item['count'], item['hours']]
            for idx, val in enumerate(row_data):
                cp = row_cells[idx].paragraphs[0]
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                clean_txt = str(val).replace('.0', '') if val != "nan" else "-"
                r = cp.add_run(clean_txt)
                set_font_kai_11(r)
                
    return doc

# --- 4. Streamlit 介面 ---
st.header("💡 照明系統設備庫 (模糊抓取版)")

if st.session_state.get('global_excel'):
    excel = st.session_state['global_excel']
    
    if st.button("🔍 執行模糊掃描並生成"):
        data = fetch_lighting_data_flexible(excel)
        
        if data:
            st.success(f"✅ 成功找到 {len(data)} 個照明分頁！")
            for name in data.keys():
                st.write(f"• 已解析：{name}")
            
            # 生成 Word
            doc_obj = generate_lighting_word(data)
            buffer = io.BytesIO()
            doc_obj.save(buffer)
            
            if 'report_warehouse' not in st.session_state:
                st.session_state['report_warehouse'] = {}
            st.session_state['report_warehouse']['2.照明系統'] = buffer.getvalue()
            st.info("報告已存入『報告輸出中心』。")
        else:
            st.error("❌ 找不到任何名稱包含『表九之二』的分頁。")
