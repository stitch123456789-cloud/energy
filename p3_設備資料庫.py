import streamlit as st
import pandas as pd
from docx import Document
import io

# --- 核心邏輯：全能掃描器 ---
def fetch_all_equipment_flexible(file):
    try:
        xl = pd.ExcelFile(file)
        sheet_names = xl.sheet_names
        results = {
            "九之一": [], # 空調
            "九之二": [], # 照明
            "九之三": []  # 其他
        }

        for sheet in sheet_names:
            # 1. 抓取表九之一 (空調系統)
            if "表九之一" in sheet:
                df = pd.read_excel(file, sheet_name=sheet, header=None)
                # 這裡加入你原本九之一的解析邏輯...
                results["九之一"].append({"name": sheet, "data": "解析後的空調數據"})

            # 2. 抓取表九之二 (照明系統)
            elif "表九之二" in sheet:
                df = pd.read_excel(file, sheet_name=sheet, header=None)
                # 這裡執行剛才寫的模糊抓取照明邏輯...
                lighting_data = process_lighting_sheet(df) # 呼叫處理照明的子函數
                results["九之二"].append({"name": sheet, "items": lighting_data})

            # 3. 抓取表九之三 (其他設備)
            elif "表九之三" in sheet:
                df = pd.read_excel(file, sheet_name=sheet, header=None)
                # 這裡加入九之三的解析邏輯...
                results["九之三"].append({"name": sheet, "data": "解析後的其他數據"})

        return results
    except Exception as e:
        st.error(f"掃描檔案失敗：{e}")
        return None

# --- 子函數：專門處理照明(九之二)的內容 ---
def process_lighting_sheet(df):
    items = []
    # 根據你的 Excel 結構，通常從 index 6 或 7 開始
    for i in range(6, len(df)):
        kind = str(df.iloc[i, 1]).strip() # 種類
        if kind == "nan" or "合計" in kind: continue
        items.append({
            "kind": kind,
            "spec": str(df.iloc[i, 5]),   # 規格
            "count": str(df.iloc[i, 9]),  # 數量
            "hours": str(df.iloc[i, 11])  # 時數
        })
    return items

# --- Streamlit 介面調整 ---
st.header("⚙️ 設備系統資料庫 (全自動生成版)")

# 優先檢查是否有單獨上傳的檔案，沒有則抓全域檔案
uploaded_file = st.file_uploader("若要單張處理，請在此上傳單獨的表九 Excel", type=["xlsx"])
final_file = uploaded_file if uploaded_file else st.session_state.get('global_excel')

if final_file:
    if st.button("🚀 立即掃描並生成完整設備報告"):
        with st.spinner("正在自動識別分頁內容..."):
            all_data = fetch_all_equipment_flexible(final_file)
            
            if all_data:
                # 建立一個新的 Word 文件
                doc = Document()
                doc.add_heading('九、使用能源設備統計', 0)
                
                # 根據抓到的資料，自動按順序生成 Word 內容
                if all_data["九之一"]:
                    st.write("✅ 偵測到表九之一 (空調)")
                    # 執行生成空調 Word 的代碼...
                    
                if all_data["九之二"]:
                    st.write("✅ 偵測到表九之二 (照明)")
                    # 執行生成照明表格的代碼... (使用剛才提供的 generate_lighting_word 邏輯)
                
                if all_data["九之三"]:
                    st.write("✅ 偵測到表九之三 (其他)")
                    # 執行生成其他設備 Word 的代碼...

                # 儲存與輸出
                # ... (存入 report_warehouse 的邏輯)

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
