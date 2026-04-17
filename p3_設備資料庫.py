import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 工具函數：設定標楷體 11 號 ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False # 依照圖片，不加粗

# --- 2. 數據抓取：掃描所有「表九之二」分頁 ---
def fetch_lighting_data(file):
    try:
        xl = pd.ExcelFile(file)
        # 抓取所有包含「表九之二」的分頁名稱
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        all_buildings_results = {}

        for sheet in target_sheets:
            # 讀取 Excel (header=None 方便定位)
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 提取建築物名稱 (分頁名稱通常是：表九之二、建築物 編號-1)
            b_name = sheet.split('、')[-1] if '、' in sheet else sheet
            
            lighting_items = []
            # 從第 7 列開始讀取數據 (index 6)
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()   # B欄:種類
                spec = str(df.iloc[i, 5]).strip()   # F欄:容量規格
                count = str(df.iloc[i, 9]).strip()  # J欄:數量
                hours = str(df.iloc[i, 11]).strip() # L欄:運轉時數
                
                # --- 數據清洗 ---
                if kind == "nan" or "註" in kind or "合計" in kind or "合計" in spec:
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
        st.error(f"解析 Excel 出錯: {e}")
        return None

# --- 3. Word 表格生成邏輯 ---
def generate_word_report(data):
    doc = Document()
    
    # 章節標題
    p = doc.add_paragraph()
    run = p.add_run("2.照明系統：")
    set_font_kai_11(run)

    for b_name, items in data.items():
        # 建築物標題
        sub_p = doc.add_paragraph()
        run_sub = sub_p.add_run(f"({b_name})")
        set_font_kai_11(run_sub)

        # 建立表格 (與圖片一致：2層表頭)
        table = doc.add_table(rows=2, cols=4)
        table.style = 'Table Grid'
        
        # 1. 第一層表頭合併
        # 合併「燈具形式」 (第2、3欄)
        header_form = table.cell(0, 1).merge(table.cell(0, 2))
        
        # 2. 垂直合併第一欄(種類)與第四欄(時數)
        table.cell(0, 0).merge(table.cell(1, 0))
        table.cell(0, 3).merge(table.cell(1, 3))

        # 3. 填入表頭文字
        header_map = [
            (table.cell(0, 0), "燈具種類"),
            (header_form, "燈具形式"),
            (table.cell(0, 3), "運轉時數(小時/年)"),
            (table.cell(1, 1), "容量規格"),
            (table.cell(1, 2), "數量")
        ]

        for cell, text in header_map:
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cell = cell.paragraphs[0]
            p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_cell.add_run(text)
            set_font_kai_11(r)

        # 4. 填入數據內容
        for item in items:
            row_cells = table.add_row().cells
            row_vals = [item['kind'], item['spec'], item['count'], item['hours']]
            for idx, val in enumerate(row_vals):
                cp = row_cells[idx].paragraphs[0]
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # 清除數字尾端的 .0
                clean_txt = str(val).replace('.0', '') if val != "nan" else "-"
                r = cp.add_run(clean_txt)
                set_font_kai_11(r)
                
    return doc

# --- 4. Streamlit 介面渲染 ---
st.header("💡 照明系統設備庫")

if st.session_state.get('global_excel'):
    excel = st.session_state['global_excel']
    
    # 執行數據抓取
    lighting_data = fetch_lighting_data(excel)
    
    if lighting_data:
        st.success(f"✅ 已偵測到 {len(lighting_data)} 個建築物的照明分頁")
        
        # 預覽數據 (選擇其中一個建築物展示)
        first_b = list(lighting_data.keys())[0]
        st.write(f"預覽數據內容 ({first_b}):")
        st.dataframe(pd.DataFrame(lighting_data[first_b]))

        # 生成報告按鈕
        if st.button("📝 生成並儲存照明系統報告"):
            doc_obj = generate_word_report(lighting_data)
            
            # 存入記憶體
            output = io.BytesIO()
            doc_obj.save(output)
            doc_bytes = output.getvalue()
            
            # 存入倉庫，讓左側打包下載中心可以抓到
            if 'report_warehouse' not in st.session_state:
                st.session_state['report_warehouse'] = {}
            st.session_state['report_warehouse']['2.照明系統'] = doc_bytes
            
            st.success("🎉 報告已成功生成！請至左側側邊欄點選「一鍵打包」下載報告。")
    else:
        st.error("❌ 找不到包含『表九之二』名稱的分頁，請檢查 Excel 內容。")
else:
    st.info("請先在左側上傳能源查核 Excel 檔案。")
