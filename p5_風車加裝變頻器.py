import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 ---
def set_table_border(table):
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

def fix_cell_font(cell, size=12, is_bold=False):
    """依照要求：標楷體、12號、置中"""
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 # 置中
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """暴力替換法：確保標籤碎裂也能換掉，並鎖定黑色標楷體"""
    for p in doc.paragraphs:
        full_p_text = "".join(run.text for run in p.runs)
        for key, val in data_map.items():
            if key in full_p_text:
                full_p_text = full_p_text.replace(key, str(val))
                p.clear() # 清空舊的 runs
                new_run = p.add_run(full_p_text)
                new_run.font.name = '標楷體'
                new_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                new_run.font.color.rgb = RGBColor(0, 0, 0)
                new_run.font.size = Pt(12)

# --- 2. 介面設定：動態設備管理 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

# 使用 session_state 儲存設備清單
if "towers" not in st.session_state:
    st.session_state.towers = [{"name": "CT-1", "rt": 300, "hp": 15.0, "fans": 3}]

with st.sidebar:
    st.header("⚙️ 設備管理")
    if st.button("➕ 新增一組水塔"):
        idx = len(st.session_state.towers) + 1
        st.session_state.towers.append({"name": f"CT-{idx}", "rt": 300, "hp": 15.0, "fans": 1})
        st.rerun()
    if st.button("❌ 刪除最後一組"):
        if len(st.session_state.towers) > 1:
            st.session_state.towers.pop()
            st.rerun()

# 編輯設備參數
for i, t in enumerate(st.session_state.towers):
    with st.expander(f"配置組別：{t['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        t['name'] = c1.text_input("編號", value=t['name'], key=f"n_{i}")
        t['rt'] = c2.number_input("噸數(RT)", value=t['rt'], key=f"r_{i}")
        t['hp'] = c3.number_input("馬力(HP)", value=t['hp'], key=f"h_{i}")
        t['fans'] = c4.number_input("台數", min_value=1, max_value=5, value=t['fans'], key=f"f_{i}")

# 運轉時數編輯表
st.subheader("📊 運轉參數明細設定")
rows = []
for t in st.session_state.towers:
    for f in range(1, t['fans'] + 1):
        rows.append({"組別": t['name'], "編號": f"{t['name']}-F{f}", "時數(hr)": 4380, "負載(%)": 100})
edit_df = st.data_editor(pd.DataFrame(rows), use_container_width=True)

# --- 3. 生成按鈕 ---
if st.button("🚀 生成 P5 專業報告", use_container_width=True):
    try:
        # 計算
        total_old_kwh = 0
        total_kw = 0
        fan_list = []
        
        curr_row = 0
        for t in st.session_state.towers:
            for f in range(t['fans']):
                h = float(edit_df.iloc[curr_row]["時數(hr)"])
                kw = t['hp'] * 0.746
                kwh = kw * h
                fan_list.append({"h": h, "kw": kw, "kwh": kwh})
                total_old_kwh += kwh
                total_kw += kw
                curr_row += 1

        doc = Document("template_p5.docx")
        
        # 文字標籤地圖
        data_map = {
            "{{UN}}": "貴單位", "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{(total_old_kwh*0.4):,.0f}", "{{PAYBACK}}": "1.2",
            "{{INVEST}}": "80.0", "{{SAVE_MONEY}}": "97.73", "{{SUPPRESS_KW}}": "13"
        }
        
        safe_replace(doc, data_map)

        # 生成橫向表格
        doc.add_page_break()
        doc.add_paragraph("【表一、現況耗電明細分析表 (橫向擴展)】")
        
        num_fans = len(fan_list)
        num_cols = 1 + num_fans + 1
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, label in enumerate(labels):
            fix_cell_font(table.cell(r, 0), is_bold=True)
            table.cell(r, 0).text = label

        # 填寫數據與合併
        col_ptr = 1
        for t in st.session_state.towers:
            f_count = t['fans']
            # 編號合併
            c_n = table.cell(0, col_ptr).merge(table.cell(0, col_ptr + f_count - 1))
            c_n.text = t['name']
            fix_cell_font(c_n, is_bold=True)
            
            # RT合併
            c_r = table.cell(1, col_ptr).merge(table.cell(1, col_ptr + f_count - 1))
            c_r.text = f"{t['rt']}RT"
            fix_cell_font(c_r)

            for i in range(f_count):
                d = fan_list[col_ptr - 1 + i]
                table.cell(2, col_ptr + i).text = f"{t['hp']:.1f}"
                table.cell(3, col_ptr + i).text = f"{d['kw']:.1f}"
                table.cell(4, col_ptr + i).text = f"{d['h']:,.0f}"
                table.cell(5, col_ptr + i).text = "100%"
                table.cell(6, col_ptr + i).text = f"{d['kwh']:,.0f}"
                for r in range(2, 7): fix_cell_font(table.cell(r, col_ptr + i), is_bold=(r==6))
            col_ptr += f_count

        # 合計
        table.cell(0, num_cols-1).text = "合計"
        table.cell(3, num_cols-1).text = f"{total_kw:.1f}"
        table.cell(6, num_cols-1).text = f"{total_old_kwh:,.0f}"
        for r in [0, 3, 6]: fix_cell_font(table.cell(r, num_cols-1), is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 生成成功！文字已替換，動態表格已完成。")
        st.download_button("📥 下載修正版報告", buf.getvalue(), "風車效益整合分析.docx")
        
    except Exception as e:
        st.error(f"發生錯誤: {e}")
