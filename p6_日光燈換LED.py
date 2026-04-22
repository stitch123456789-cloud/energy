import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io

# --- 1. Word 工具函式：在模板中插入表格 ---
def add_equipment_table(doc, df):
    """
    在 Word 文件末尾或特定位置增加設備清單表格
    """
    st.info("正在將設備資料寫入 Word 表格...")
    # 建立表格 (列數為 df 行數 + 1 個表頭，欄數為 df 欄數)
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid' # 使用 Word 內建網格樣式

    # 寫入表頭
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)
    
    # 寫入資料
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, item in enumerate(row):
            row_cells[i].text = str(item)

# --- 2. 主程式邏輯 ---
st.title("⚙️ 表九之二：動力設備自動產製")

uploaded_file = st.session_state.get('global_excel')

if uploaded_file is None:
    st.warning("⚠️ 請先在左側邊欄上傳「完整能源查核 Excel」檔案。")
else:
    # A. 讀取並提取表九之二
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    all_building_data = []

    for sheet_name, df_raw in all_sheets.items():
        if '表九之二' in sheet_name:
            # 尋找表頭座標 (鎖定關鍵字：設備名稱 或 種類)
            header_idx = -1
            for idx, row in df_raw.head(15).iterrows():
                row_str = "".join([str(val) for val in row.values if pd.notna(val)])
                if '設備' in row_str or '種類' in row_str or '名稱' in row_str:
                    header_idx = idx
                    break
            
            if header_idx != -1:
                df = df_raw.iloc[header_idx+1:].copy()
                df.columns = df_raw.iloc[header_idx]
                df = df.dropna(how='all').dropna(axis=1, how='all') # 去除空行空欄
                df['來源建築物'] = sheet_name
                all_building_data.append(df)

    if not all_building_data:
        st.error("❌ 在 Excel 中找不到任何名為「表九之二」的分頁。")
    else:
        # B. 合併與編輯
        final_df = pd.concat(all_building_data, ignore_index=True)
        
        st.success(f"✅ 已成功從各分頁提取 {len(final_df)} 筆動力設備資料。")
        
        # 讓你在網頁上直接編輯內容
        st.write("📋 **待產出設備清單：**")
        edited_df = st.data_editor(final_df, use_container_width=True, num_rows="dynamic")

        # C. 產出 Word
        st.divider()
        report_title = st.text_input("報告檔案名稱", value="動力系統設備查核清單")
        
        if st.button("🚀 生成 Word 設備報告", use_container_width=True):
            try:
                # 這裡建議你放一個對應表九之二的 Word 模板檔案
                # 如果沒有模板，程式會自動建立一個新的空白 Word
                try:
                    doc = Document("表九之二模板.docx")
                except:
                    doc = Document()
                    doc.add_heading('表九之二、動力系統設備資料', 0)

                # 將編輯後的表格內容寫入 Word
                add_equipment_table(doc, edited_df)

                # 存檔至記憶體
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                # 提供下載
                st.session_state['report_warehouse'][report_title] = doc_io.getvalue()
                st.success(f"✅ 「{report_title}」已生成！")
                st.download_button(
                    label="📥 下載 Word 檔案",
                    data=doc_io.getvalue(),
                    file_name=f"{report_title}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"產出 Word 失敗：{e}")
