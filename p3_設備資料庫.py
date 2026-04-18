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

def set_font_bold_black_14(run):
    # 設定黑體 14號 加粗
    run.font.name = '微軟正黑體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    run.font.size = Pt(14)
    run.font.bold = True

# --- 2. 數據抓取與自動加總邏輯 ---
def fetch_and_aggregate_lighting(file):
    try:
        xl = pd.ExcelFile(file)
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        
        if not target_sheets:
            return None
        
        # 使用字典來做加總，Key 為 (種類, 規格, 時數)
        aggregated_data = {}

        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # 數據通常從第 7 列 (index 6) 開始
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()   # B欄
                spec = str(df.iloc[i, 5]).strip()   # F欄
                count_str = str(df.iloc[i, 9]).strip()  # J欄
                hours_str = str(df.iloc[i, 11]).strip() # L欄
                
                # 數據清洗
                if kind == "nan" or "註" in kind or "合計" in kind: continue
                if spec == "nan" or spec == "": continue
                
                # 去除種類前面的數字 (例如 "1. 日光燈" -> "日光燈")
                # 使用 split('.') 處理
                if '.' in kind:
                    kind = kind.split('.')[-1].strip()
                
                try:
                    count = int(float(count_str.replace(',', '')))
                    hours = int(float(hours_str.replace(',', '')))
                except:
                    continue

                # 建立唯一 Key
                key = (kind, spec, hours)
                
                if key in aggregated_data:
                    aggregated_data[key] += count
                else:
                    aggregated_data[key] = count
                    
        return aggregated_data
    except Exception as e:
        st.error(f"數據加總失敗：{e}")
        return None

# --- 3. Word 生成邏輯 (全域總表) ---
def generate_aggregated_word(aggregated_data):
    doc = Document()
    
    # 標題：2.照明系統： (黑體 14號 加粗)
    p = doc.add_paragraph()
    run_title = p.add_run("2.照明系統：")
    set_font_bold_black_14(run_title)

    # 建立表格
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    # 合併表頭
    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 3).merge(table.cell(1, 3))
    
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

    # 填入加總後的數據
    # 將字典轉回列表並排序（可選，按種類排序）
    sorted_items = sorted(aggregated_data.items(), key=lambda x: x[0][0])

    for (kind, spec, hours), count in sorted_items:
        row_cells = table.add_row().cells
        row_data = [kind, spec, str(count), str(hours)]
        for idx, val in enumerate(row_data):
            cp = row_cells[idx].paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cp.add_run(val)
            set_font_kai_11(r)
                
    return doc

# --- 4. Streamlit 主介面 ---
st.header("⚙️ 設備系統資料庫 (全域加總版)")

up_file = st.file_uploader("請上傳 Excel (支援多建築物自動合併)", type=["xlsx"])
final_file = up_file if up_file else st.session_state.get('global_excel')

if final_file:
    if st.button("🚀 掃描所有建築物並生成唯一總表"):
        agg_data = fetch_and_aggregate_lighting(final_file)
        
        if agg_data:
            st.success(f"✅ 已完成跨建築物加總！共合併為 {len(agg_data)} 個項目。")
            
            doc_obj = generate_aggregated_word(agg_data)
            buffer = io.BytesIO()
            doc_obj.save(buffer)
            
            st.download_button(
                label="📥 下載全域照明總表 (.docx)",
                data=buffer.getvalue(),
                file_name="照明系統合併報告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("查無符合的分頁內容。")
