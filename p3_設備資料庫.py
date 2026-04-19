import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體函數 ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False

def set_font_kai_bold_14(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(14)
    run.font.bold = True

# --- 2. 數據抓取：(2) 冰水主機規格 ---
def fetch_chiller_spec(file):
    try:
        xl = pd.ExcelFile(file)
        # 抓取所有表九之一的分頁
        target_sheets = [s for s in xl.sheet_names if "表九之一" in s]
        if not target_sheets: return None
        
        all_chillers = []
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            # 數據通常從 index 6 開始 (B欄是 index 1)
            for i in range(6, len(df)):
                name = str(df.iloc[i, 1]).strip() # B: 設備名稱
                if "冰水主機" not in name and "空調系統" not in name: continue
                
                sn = str(df.iloc[i, 2]).strip()      # C: 設備編號
                form = str(df.iloc[i, 5]).strip()    # F: 形式
                inverter_raw = str(df.iloc[i, 7]).strip() # H: 有無 (變頻)
                volt = str(df.iloc[i, 11]).strip()   # L: 電壓
                power = str(df.iloc[i, 12]).strip()  # M: 功率
                year = str(df.iloc[i, 13]).strip()   # N: 製造年份
                cap_val = str(df.iloc[i, 14]).strip() # O: 容量數值
                cap_unit = str(df.iloc[i, 15]).strip() # P: 單位
                qty = str(df.iloc[i, 21]).strip()    # V: 數量
                
                if cap_val == "nan" or cap_val == "": continue

                # A. 變頻/定頻判定 (H欄)
                type_tag = "變頻" if inverter_raw == "有" else "定頻"
                
                # B. 單位換算 RT
                try:
                    val = float(cap_val.replace(',', ''))
                    if "kW" in cap_unit:
                        rt_val = round(val / 3.517, 1)
                    elif "kcal" in cap_unit:
                        rt_val = round(val / 3024, 1)
                    else: # 假設原本就是 RT
                        rt_val = val
                except:
                    rt_val = cap_val
                
                all_chillers.append([sn, form, type_tag, volt, power, year, rt_val, qty])
        
        return all_chillers
    except Exception as e:
        st.error(f"冰水主機抓取失敗: {e}")
        return None

# --- 3. Word 表格生成：(2) 冰水主機規格 ---
def add_chiller_spec_table(doc, chiller_data):
    p2 = doc.add_paragraph()
    p2.paragraph_format.left_indent = Pt(20)
    run2 = p2.add_run("(2) 冰水主機規格：")
    set_font_kai_bold_14(run2)

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    headers = ["設備編號", "形式", "備註", "電壓\n(V)", "功率\n(kW)", "製造年份", "容量\n(RT)", "現有數量"]
    
    # 設定表頭
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font_kai_11(p.runs[0])

    # 填入數據
    for row_data in chiller_data:
        row_cells = table.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = str(val)
            row_cells[i].vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(p.add_run("")) # 確保字體設定成功

# --- 4. 原有的照明與空調開啟模式函數 (略，請保留之前版本) ---
# [保留 fetch_and_aggregate_lighting, add_lighting_table, add_ac_mode_table]

# --- 5. Streamlit 介面渲染 ---
st.subheader("⚙️ 設備系統資料庫")

# [保留之前的 3. 空調主機開啟模式設定 UI]
# (這部分與上次代碼相同，省略顯示以節省篇幅)

st.markdown("---")
up_file = st.file_uploader("請上傳能源查核 Excel", type=["xlsx"])
final_file = up_file if up_file else st.session_state.get('global_excel')

if final_file:
    if st.button("🚀 生成並下載設備系統報告", use_container_width=True):
        doc = Document()
        
        # 1. 插入照明系統 (編號 2)
        light_data = fetch_and_aggregate_lighting(final_file)
        if light_data:
            # 這裡我們需要一個 add_lighting_table 函數
            pass 
        
        doc.add_paragraph() 
        
        # 2. 插入空調系統 (編號 3)
        # (1) 開啟模式
        # add_ac_mode_table(doc, ac_rows)
        
        # (2) 冰水主機規格 (新加入)
        chiller_data = fetch_chiller_spec(final_file)
        if chiller_data:
            add_chiller_spec_table(doc, chiller_data)
        
        # 下載輸出
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "設備報告.docx", use_container_width=True)
