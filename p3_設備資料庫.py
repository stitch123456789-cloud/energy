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

# --- 2. 子函數：專門處理照明(九之二)的內容解析 ---
def process_lighting_sheet(df):
    items = []
    # 根據你的 Excel 結構，通常從 index 6 或 7 開始
    for i in range(6, len(df)):
        try:
            kind = str(df.iloc[i, 1]).strip() # B欄：種類
            if kind == "nan" or "合計" in kind or "註" in kind: 
                continue
            
            items.append({
                "kind": kind,
                "spec": str(df.iloc[i, 5]),   # F欄：規格
                "count": str(df.iloc[i, 9]),  # J欄：數量
                "hours": str(df.iloc[i, 11])  # L欄：時數
            })
        except:
            continue
    return items

# --- 3. Word 生成邏輯 (兩層表頭合併儲存格) ---
def add_lighting_table_to_doc(doc, b_name, items):
    # 建築物子標題
    sub_p = doc.add_paragraph()
    run_sub = sub_p.add_run(f"({b_name})")
    set_font_kai_11(run_sub)

    # 建立兩層表頭表格 (4欄)
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    # 合併儲存格還原圖片格式
    table.cell(0, 1).merge(table.cell(0, 2)) # 第一列：燈具形式 (跨兩欄)
    table.cell(0, 0).merge(table.cell(1, 0)) # 第一欄：種類垂直合併
    table.cell(0, 3).merge(table.cell(1, 3)) # 第四欄：時數垂直合併

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

    # 填入數據內容
    for item in items:
        row_cells = table.add_row().cells
        row_data = [item['kind'], item['spec'], item['count'], item['hours']]
        for idx, val in enumerate(row_data):
            cp = row_cells[idx].paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clean_txt = str(val).replace('.0', '') if str(val) != "nan" else "-"
            r = cp.add_run(clean_txt)
            set_font_kai_11(r)

# --- 4. 主介面與執行邏輯 ---
st.header("⚙️ 設備系統資料庫 (全自動融合版)")

# 優先檢查單獨上傳，否則抓全域
uploaded_file = st.file_uploader("若要單張處理，請在此上傳單獨的表九 Excel", type=["xlsx"])
final_file = uploaded_file if uploaded_file else st.session_state.get('global_excel')

if final_file:
    # 唯一的大按鈕
    if st.button("🚀 立即掃描並下載完整設備報告"):
        with st.spinner("正在自動識別分頁內容並生成 Word..."):
            try:
                xl = pd.ExcelFile(final_file)
                sheet_names = xl.sheet_names
                
                # 初始化 Word
                doc = Document()
                # 設定大標題
                p = doc.add_paragraph()
                run_title = p.add_run("2.照明系統：")
                set_font_kai_11(run_title)
                
                found_any = False
                
                # 遍歷所有分頁進行模糊抓取
                for sheet in sheet_names:
                    # --- 處理九之二 (照明) ---
                    if "表九之二" in sheet:
                        df = pd.read_excel(final_file, sheet_name=sheet, header=None)
                        items = process_lighting_sheet(df)
                        
                        if items:
                            # 提取乾淨的建築物名稱
                            b_name = sheet.split('、')[-1] if '、' in sheet else sheet
                            # 加入表格到 Word
                            add_lighting_table_to_doc(doc, b_name, items)
                            found_any = True
                            st.write(f"✅ 已解析照明分頁：{sheet}")
                    
                    # --- 預留位置給未來的九之一與九之三 ---
                    # elif "表九之一" in sheet: ...
                
                if found_any:
                    # 輸出下載按鈕
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    st.download_button(
                        label="📥 點此下載設備系統報告 (.docx)",
                        data=buffer.getvalue(),
                        file_name="設備系統報告.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    # 同時存一份到倉庫 (備份用)
                    if 'report_warehouse' not in st.session_state:
                        st.session_state['report_warehouse'] = {}
                    st.session_state['report_warehouse']['設備系統報告'] = buffer.getvalue()
                else:
                    st.warning("查無符合的分頁名稱（需含『表九之二』）。")
                    
            except Exception as e:
                st.error(f"執行失敗：{e}")
else:
    st.info("請先在側邊欄或上方上傳能源查核 Excel 檔案。")
