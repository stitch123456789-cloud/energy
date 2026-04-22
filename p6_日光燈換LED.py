import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re

# --- 1. Word 工具：格式鎖定 (確保產出的 Word 標楷體、不跑位) ---
def set_cell_style(cell, text, size=10, is_bold=False):
    """設定儲存格內容與格式"""
    # 確保儲存格內有段落
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.text = str(text) if pd.notna(text) else ""
    # 設定字體
    for run in p.runs:
        run.font.name = '標楷體'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        run.font.size = Pt(size)
        run.font.bold = is_bold

# --- 2. 核心提取邏輯 ---
st.title("⚙️ 表九之二：動力系統設備提取器")

uploaded_file = st.session_state.get('global_excel')

if uploaded_file is None:
    st.warning("⚠️ 請先在左側邊欄上傳「完整能源查核 Excel」檔案。")
else:
    # A. 自動讀取與定位
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    all_extracted_dfs = []

    for sheet_name, df_raw in all_sheets.items():
        
        # ✨ 1. 分頁名稱模糊清洗：去除所有空白、換行符號
        clean_name = re.sub(r'\s+', '', sheet_name)
        
        # ✨ 2. 模糊關鍵字比對：只要包含以下任何一種寫法就抓！
        # 涵蓋：表九之二、表9之2、表9-2、表九-2... 等等
        is_target_sheet = bool(re.search(r'表九之二|表9之2|表9-2|表九-2', clean_name))
        
        if is_target_sheet:
            st.toast(f"🔍 模糊比對成功：正在讀取 [{sheet_name}]...") # 畫面右下角跳出提示
            
            # 尋找表頭座標 (掃描前20列)
            header_row_idx = -1
            for idx, row in df_raw.head(20).iterrows():
                row_str = "".join([str(v) for v in row.values if pd.notna(v)])
                # 根據你截圖的關鍵字：編號、設備系統、數量
                if '編號' in row_str and ('設備系統' in row_str or '種類' in row_str):
                    header_row_idx = idx
                    break
            
            if header_row_idx != -1:
                # 重新定義 DataFrame (表頭在那一列，資料從下一列開始)
                df = df_raw.iloc[header_row_idx+1:].copy()
                df.columns = df_raw.iloc[header_row_idx]
                
                # 清除無效行 (例如編號是空的)
                df = df[df.iloc[:, 0].notna()]
                # 清除完全空白的列
                df = df.dropna(axis=1, how='all')
                
                df['來源建築物'] = sheet_name # 這裡保留原本長長的名字，讓你知道資料是哪來的
                all_extracted_dfs.append(df)

    if not all_extracted_dfs:
        st.error("❌ 找不到名為「表九之二」的分頁，或表格內缺少關鍵欄位名稱。")
    else:
        # B. 數據合併與展示
        final_df = pd.concat(all_extracted_dfs, ignore_index=True)
        
        st.success(f"✅ 成功掃描！共發現 {len(final_df)} 筆動力系統設備資料。")
        
        # 移除一些全空或不必要的隱藏欄位
        final_df = final_df.loc[:, final_df.columns.notna()]
        
        st.write("📊 **已提取資料預覽 (你可以在此直接修改內容)：**")
        edited_df = st.data_editor(final_df, use_container_width=True, num_rows="dynamic")

        # C. 生成 Word 報告
        st.divider()
        if st.button("🚀 生成 Word 設備清單報告", use_container_width=True):
            try:
                # 建立新 Word
                doc = Document()
                doc.add_heading('表九之二、動力系統設備資料清單', 0)
                
                # 建立表格
                table = doc.add_table(rows=1, cols=len(edited_df.columns))
                table.style = 'Table Grid'
                
                # 寫入表頭 (粗體)
                hdr_cells = table.rows[0].cells
                for i, col_name in enumerate(edited_df.columns):
                    set_cell_style(hdr_cells[i], col_name, is_bold=True)
                
                # 寫入每一列資料
                for _, row in edited_df.iterrows():
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        set_cell_style(row_cells[i], value)
                
                # 存檔處理
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                # 下載與通知
                report_title = "表九之二動力系統清單"
                st.session_state['report_warehouse'][report_title] = doc_io.getvalue()
                
                st.success("✅ 報告生成成功！")
                st.download_button(
                    label="📥 下載 Word 設備報告",
                    data=doc_io.getvalue(),
                    file_name=f"{report_title}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Word 生成過程發生錯誤：{e}")
