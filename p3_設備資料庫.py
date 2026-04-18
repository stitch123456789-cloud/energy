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

# --- 2. 數據抓取：照明系統自動加總 ---
def fetch_and_aggregate_lighting(file):
    try:
        xl = pd.ExcelFile(file)
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        if not target_sheets: return None
        
        aggregated_data = {}
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            for i in range(6, len(df)):
                kind = str(df.iloc[i, 1]).strip()
                spec = str(df.iloc[i, 5]).strip()
                count_str = str(df.iloc[i, 9]).strip()
                hours_str = str(df.iloc[i, 11]).strip()
                
                if kind == "nan" or "註" in kind or "合計" in kind: continue
                if spec == "nan" or spec == "": continue
                if '.' in kind: kind = kind.split('.')[-1].strip()
                
                try:
                    count = int(float(count_str.replace(',', '')))
                    hours = int(float(hours_str.replace(',', '')))
                    key = (kind, spec, hours)
                    aggregated_data[key] = aggregated_data.get(key, 0) + count
                except: continue
        return aggregated_data
    except: return None

# --- 3. Word 表格生成：空調開啟模式 ---
def add_ac_mode_table(doc, ac_data):
    # 標題 3. 空調系統
    p3 = doc.add_paragraph()
    run3 = p3.add_run("3. 空調系統：")
    set_font_kai_bold_14(run3)

    # 標題 (1) 空調主機開啟模式
    p1 = doc.add_paragraph()
    p1.paragraph_format.left_indent = Pt(20)
    run1 = p1.add_run("(1) 空調主機開啟模式：")
    set_font_kai_bold_14(run1)

    table = doc.add_table(rows=4, cols=6)
    table.style = 'Table Grid'
    headers = ["季節", "主機總容量\n(RT)", "冰機總開啟台數", "負載率\n(%)", "合計容量\n(RT)", "出水溫度設定 (°C)"]
    
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_font_kai_11(run)

    for r_idx, row_vals in enumerate(ac_data, start=1):
        for c_idx, val in enumerate(row_vals):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            set_font_kai_11(run)

# --- 4. Word 表格生成：照明系統 ---
def add_lighting_table(doc, lighting_data):
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("2. 照明系統：")
    set_font_kai_bold_14(run_title)

    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 3).merge(table.cell(1, 3))
    
    headers = [(0,0,"燈具種類"),(0,1,"燈具形式"),(0,3,"運轉時數(小時/年)"),(1,1,"容量規格"),(1,2,"數量")]
    for r, c, txt in headers:
        cell = table.cell(r, c)
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(txt)
        set_font_kai_11(run)

    sorted_items = sorted(lighting_data.items(), key=lambda x: x[0][0])
    for (kind, spec, hours), count in sorted_items:
        row_cells = table.add_row().cells
        for idx, val in enumerate([kind, spec, str(count), str(hours)]):
            p = row_cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            set_font_kai_11(run)

# --- 5. Streamlit 介面渲染 ---
st.subheader("⚙️ 設備系統資料庫")

# --- 空調手動輸入區 ---
st.markdown("### ❄️ 3. 空調主機開啟模式設定")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.write("**季節**")
    st.caption("夏季"); st.caption("春秋"); st.caption("冬季")
with c2:
    st.write("**主機總容量(RT)**")
    rt_s = st.number_input("夏季容量", value=600, label_visibility="collapsed")
    rt_sp = st.number_input("春秋容量", value=450, label_visibility="collapsed")
    rt_w = st.number_input("冬季容量", value=450, label_visibility="collapsed")
with c3:
    st.write("**負載率(%)**")
    ld_s = st.number_input("夏季負載", value=70, label_visibility="collapsed")
    ld_sp = st.number_input("春秋負載", value=70, label_visibility="collapsed")
    ld_w = st.number_input("冬季負載", value=60, label_visibility="collapsed")
with c4:
    st.write("**出水溫度(°C)**")
    tp_s = st.number_input("夏季溫度", value=7, label_visibility="collapsed")
    tp_sp = st.number_input("春秋溫度", value=7, label_visibility="collapsed")
    tp_w = st.number_input("冬季溫度", value=7, label_visibility="collapsed")

st.write("**冰機總開啟台數**")
tc1, tc2, tc3 = st.columns(3)
ct_s = tc1.number_input("夏季台數", value=1)
ct_sp = tc2.number_input("春秋台數", value=1)
ct_w = tc3.number_input("冬季台數", value=1)

# 計算合計容量
ac_rows = [
    ["夏季", rt_s, ct_s, f"{ld_s}%", round(rt_s*ld_s/100, 1), tp_s],
    ["春秋", rt_sp, ct_sp, f"{ld_sp}%", round(rt_sp*ld_sp/100, 1), tp_sp],
    ["冬季", rt_w, ct_w, f"{ld_w}%", round(rt_w*ld_w/100, 1), tp_w]
]

# 檔案上傳
up_file = st.file_uploader("請上傳能源查核 Excel", type=["xlsx"])
final_file = up_file if up_file else st.session_state.get('global_excel')

if final_file:
    if st.button("🚀 生成並下載設備系統報告", use_container_width=True):
        doc = Document()
        
        # 1. 插入照明系統 (根據你之前編號為 2)
        light_data = fetch_and_aggregate_lighting(final_file)
        if light_data:
            add_lighting_table(doc, light_data)
        
        doc.add_paragraph() # 隔行
        
        # 2. 插入空調系統 (編號為 3)
        add_ac_mode_table(doc, ac_rows)
        
        # 下載
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "設備報告.docx", use_container_width=True)
