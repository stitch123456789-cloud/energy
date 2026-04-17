import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 工具函數：設定標楷體 ---
def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold

# --- 1. 數據抓取核心：照明系統 (表九之二) ---
def fetch_lighting_data(file):
    try:
        xl = pd.ExcelFile(file)
        all_sheets = xl.sheet_names
        
        # 尋找所有名稱包含 "表九之二" 的分頁 (因可能有多個建築物分頁)
        target_sheets = [s for s in all_sheets if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        all_buildings_results = {}

        for sheet in target_sheets:
            # 讀取 Excel，不設 header 方便精準定位
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 抓取建築物名稱 (假設在 A4 或 B4 位置，根據截圖，通常在標題列下方)
            # 這裡我們掃描前 5 列找 "建築物名稱" 字眼
            b_name = sheet # 預設用分頁名稱
            for r in range(10):
                row_str = str(df.iloc[r, :].values)
                if "編號" in row_str:
                    # 嘗試從分頁名稱或內容提取更乾淨的名稱
                    b_name = sheet.split('、')[-1] if '、' in sheet else sheet
                    break
            
            lighting_items = []
            # 數據通常從第 7 列開始 (index 6)
            for i in range(6, len(df)):
                # B欄:種類(1), F欄:規格(5), J欄:數量(9), L欄:時數(11)
                kind = str(df.iloc[i, 1])   # B
                spec = str(df.iloc[i, 5])   # F
                count = str(df.iloc[i, 9])  # J
                hours = str(df.iloc[i, 11]) # L
                
                # 過濾無效行 (如果種類是空值或是 "註：" 就跳過)
                if kind == "nan" or "註" in kind or "合計" in kind:
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

# --- 2. Word 表格生成：2. 照明系統 ---
def add_lighting_section(doc, data):
    # 標題：2.照明系統
    p = doc.add_paragraph()
    run = p.add_run("2.照明系統：")
    set_font_kai(run, size=12, is_bold=True)

    for b_name, items in data.items():
        # 建築物小標 (例如：建築物 編號-1)
        sub_p = doc.add_paragraph()
        sub_p.paragraph_format.left_indent = Pt(20) # 稍微縮排
        run_sub = sub_p.add_run(f"({b_name})")
        set_font_kai(run_sub, size=11, is_bold=True)

        # 建立 4 欄表格
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        # 設定表頭
        hdr_cells = table.rows[0].cells
        hdr_labels = ['燈具種類', '燈具形式\n(容量規格)', '數量', '運轉時數\n(小時/年)']
        
        for idx, label in enumerate(hdr_labels):
            cell_p = hdr_cells[idx].paragraphs[0]
            cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cell_p.add_run(label)
            set_font_kai(r, size=10, is_bold=True)

        # 填入數據
        for item in items:
            row_cells = table.add_row().cells
            row_data = [item['kind'], item['spec'], item['count'], item['hours']]
            for idx, val in enumerate(row_data):
                cell_p = row_cells[idx].paragraphs[0]
                cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # 去除 .0 (針對數量和時數是整數的情況)
                clean_val = str(val).replace('.0', '') if val != "nan" else "-"
                r = cell_p.add_run(clean_val)
                set_font_kai(r, size=10)

# --- 3. Streamlit 執行邏輯 ---
if st.session_state.get('global_excel'):
    excel = st.session_state['global_excel']
    
    if st.button("🚀 開始解析並生成照明系統表格"):
        lighting_data = fetch_lighting_data(excel)
        
        if lighting_data:
            doc = Document()
            add_lighting_section(doc, lighting_data)
            
            # 提供下載
            buffer = io.BytesIO()
            doc.save(buffer)
            st.download_button(
                label="📥 下載照明系統報告 (Word)",
                data=buffer.getvalue(),
                file_name="2_照明系統報告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("✅ 解析完成！已根據 Excel 分頁生成多個建築物表格。")
        else:
            st.error("❌ 找不到『表九之二』相關分頁，請檢查 Excel 工作表名稱。")
