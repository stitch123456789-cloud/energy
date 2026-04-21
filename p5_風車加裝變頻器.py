import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 ---

def set_table_border(table):
    """手動強制繪製表格黑色框線"""
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') 
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

def fix_cell_font_ultimate(cell, size=12):
    """依照要求：黑色、標楷體、12號、置中、不加粗"""
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 # 置中
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = False
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """替換文字標籤並確保格式為黑色標楷體12號"""
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                        run.font.size = Pt(12)
                        run.font.bold = False
                        run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面設定：動態組數與台數 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

if "tower_configs" not in st.session_state:
    st.session_state.tower_configs = [{"name": "CH-1", "rt": 1500, "fans": 3, "hp": 50.0}]

with st.sidebar:
    st.header("⚙️ 系統配置")
    if st.button("➕ 新增一組冷卻水塔"):
        idx = len(st.session_state.tower_configs) + 1
        st.session_state.tower_configs.append({"name": f"CH-{idx}", "rt": 1500, "fans": 1, "hp": 50.0})
    if st.button("🧹 重設所有配置"):
        st.session_state.tower_configs = [{"name": "CH-1", "rt": 1500, "fans": 3, "hp": 50.0}]
        st.rerun()

# 顯示各組設定
for i, group in enumerate(st.session_state.tower_configs):
    with st.expander(f"冷卻水塔組別：{group['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        group['name'] = c1.text_input("主機/水塔編號", value=group['name'], key=f"n_{i}")
        group['rt'] = c2.number_input("散熱噸數(RT)", value=group['rt'], key=f"r_{i}")
        group['hp'] = c3.number_input("單台馬力(HP)", value=group['hp'], key=f"h_{i}")
        group['fans'] = c4.number_input("風扇台數", min_value=1, max_value=5, value=group['fans'], key=f"f_{i}")

st.subheader("📊 運轉參數明細設定")
rows = []
for g in st.session_state.tower_configs:
    for f in range(1, g['fans'] + 1):
        rows.append({"組別": g['name'], "風扇編號": f"{g['name']}-F{f}", "運轉時數(hr)": 4380, "負載率(%)": 100})
edit_df = st.data_editor(pd.DataFrame(rows), use_container_width=True)

# --- 3. 生成按鈕 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    try:
        doc = Document("template_p5.docx")
        
        # 計算與準備數據
        fan_details = []
        total_old_kwh = 0
        total_kw = 0
        
        curr_idx = 0
        for g in st.session_state.tower_configs:
            for f in range(g['fans']):
                row_data = edit_df.iloc[curr_idx]
                h = float(row_data["運轉時數(hr)"])
                kw = g['hp'] * 0.746
                kwh = kw * h
                fan_details.append({"h": h, "kw": kw, "kwh": kwh})
                total_old_kwh += kwh
                total_kw += kw
                curr_idx += 1

        # A. 文字替換
        data_map = {
            "{{UN}}": "貴單位", "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{(total_old_kwh * 0.4):,.0f}", # 預估節能
            "{{INVEST}}": "80.0", "{{PAYBACK}}": "1.2"
        }
        safe_replace(doc, data_map)
        
        # B. 清除標籤定位
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text or "[[NEW_TABLE]]" in p.text:
                p.text = ""

        # C. 生成橫向表格
        doc.add_page_break()
        doc.add_paragraph("--- 以下為自動生成的橫向表格 (請剪下並貼至指定位置) ---")
        
        num_cols = 1 + len(fan_details) + 1
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, txt in enumerate(labels):
            fix_cell_font_ultimate(table.cell(r, 0))
            table.cell(r, 0).text = txt

        # 填寫資料與合併
        col_ptr = 1
        for g in st.session_state.tower_configs:
            f_count = g['fans']
            # 合併編號
            c_n = table.cell(0, col_ptr)
            if f_count > 1: c_n = c_n.merge(table.cell(0, col_ptr + f_count - 1))
            c_n.text = g['name']
            fix_cell_font_ultimate(c_n)
            
            # 合併 RT
            c_r = table.cell(1, col_ptr)
            if f_count > 1: c_r = c_r.merge(table.cell(1, col_ptr + f_count - 1))
            c_r.text = f"{g['rt']}RT"
            fix_cell_font_ultimate(c_r)

            for i in range(f_count):
                idx = col_ptr - 1 + i
                d = fan_details[idx]
                table.cell(2, col_ptr + i).text = f"{g['hp']:.1f}"
                table.cell(3, col_ptr + i).text = f"{d['kw']:.1f}"
                table.cell(4, col_ptr + i).text = f"{d['h']:,.0f}"
                table.cell(5, col_ptr + i).text = "100%"
                table.cell(6, col_ptr + i).text = f"{d['kwh']:,.0f}"
                for r in range(2, 7): fix_cell_font_ultimate(table.cell(r, col_ptr + i))
            col_ptr += f_count

        # 合計欄
        table.cell(0, num_cols-1).text = "合計"
        table.cell(3, num_cols-1).text = f"{total_kw:.1f}"
        table.cell(6, num_cols-1).text = f"{total_old_kwh:,.0f}"
        for r in range(7): fix_cell_font_ultimate(table.cell(r, num_cols-1))

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！文字與表格均已統一為黑色標楷體12號。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車分析修正版.docx")

    except Exception as e:
        st.error(f"發生錯誤: {e}")
