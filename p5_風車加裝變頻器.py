import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心格式工具 ---
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

def fix_cell_format(cell, text, is_bold=False, size=10):
    """置中、標楷體、設定文字"""
    p = cell.paragraphs[0]
    p.alignment = 1 # 置中
    p.clear()
    run = p.add_run(str(text))
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace_keep_style(doc, data_map):
    """替換文字並保留原始紅字/粗體格式"""
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

# 多組水塔設定
if "tower_groups" not in st.session_state:
    st.session_state.tower_groups = [{"name": "CT-1", "rt": 300, "fans": 3, "hp": 15.0}]

st.subheader("🏗️ 冷卻水塔組別配置")
for i, group in enumerate(st.session_state.tower_groups):
    with st.expander(f"組別：{group['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        group['name'] = c1.text_input("水塔編號", value=group['name'], key=f"gn_{i}")
        group['rt'] = c2.number_input("散熱噸數(RT)", value=group['rt'], key=f"gr_{i}")
        group['hp'] = c3.number_input("單扇馬力(HP)", value=group['hp'], key=f"gh_{i}")
        group['fans'] = c4.number_input("風扇台數", min_value=1, max_value=5, value=group['fans'], key=f"gf_{i}")

if st.button("➕ 新增一組水塔"):
    st.session_state.tower_groups.append({"name": f"CT-{len(st.session_state.tower_groups)+1}", "rt": 300, "fans": 1, "hp": 15.0})
    st.rerun()

st.subheader("📅 運轉時數與負載設定")
rows = []
for g in st.session_state.tower_groups:
    for f in range(1, g['fans'] + 1):
        rows.append({"水塔組別": g['name'], "風扇編號": f"{g['name']}-F{f}", "時數(hr)": 4380, "負載率(%)": 100})
edit_df = st.data_editor(pd.DataFrame(rows), use_container_width=True)

# --- 3. 生成按鈕 ---
if st.button("🚀 生成報告 (橫向合併表格)", use_container_width=True):
    try:
        doc = Document("template_p5.docx")
        
        # 計算總耗電
        total_kwh = 0
        total_kw = 0
        fan_details = [] # 用來存每一台的計算結果
        
        # 依照介面設定進行計算
        curr_fan_idx = 0
        for g in st.session_state.tower_groups:
            group_kwh = 0
            for f in range(g['fans']):
                h = float(edit_df.iloc[curr_fan_idx]["時數(hr)"])
                kw = g['hp'] * 0.746
                kwh = kw * h
                total_kwh += kwh
                total_kw += kw
                fan_details.append({"h": h, "kw": kw, "kwh": kwh})
                curr_fan_idx += 1

        # 文字標籤替換
        data_map = {
            "{{UN}}": "貴單位", "{{OLD_KWH}}": f"{total_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{(total_kwh*0.35):,.0f}", "{{PAYBACK}}": "1.2"
        }
        safe_replace_keep_style(doc, data_map)

        # 生成橫向表格
        doc.add_page_break()
        doc.add_paragraph("【現況耗電明細分析表 (橫向合併版)】")
        
        num_fans = len(fan_details)
        num_cols = 1 + num_fans + 1
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, txt in enumerate(labels):
            fix_cell_format(table.cell(r, 0), txt, is_bold=True)

        # 填寫資料與橫向合併
        col_ptr = 1
        for g in st.session_state.tower_groups:
            f_count = g['fans']
            # 合併編號列
            cell_n = table.cell(0, col_ptr)
            if f_count > 1: cell_n = cell_n.merge(table.cell(0, col_ptr + f_count - 1))
            fix_cell_format(cell_n, g['name'], is_bold=True)
            
            # 合併 RT 列
            cell_r = table.cell(1, col_ptr)
            if f_count > 1: cell_r = cell_r.merge(table.cell(1, col_ptr + f_count - 1))
            fix_cell_format(cell_r, g['rt'])

            # 填寫每一台風扇的細節
            for i in range(f_count):
                detail = fan_details[col_ptr - 1 + i]
                fix_cell_format(table.cell(2, col_ptr + i), f"{g['hp']:.1f}")
                fix_cell_format(table.cell(3, col_ptr + i), f"{detail['kw']:.1f}")
                fix_cell_format(table.cell(4, col_ptr + i), f"{detail['h']:,.0f}")
                fix_cell_format(table.cell(5, col_ptr + i), "100%")
                fix_cell_format(table.cell(6, col_ptr + i), f"{detail['kwh']:,.0f}", is_bold=True)
            col_ptr += f_count

        # 合計欄
        fix_cell_format(table.cell(0, num_cols-1), "合計", is_bold=True)
        fix_cell_format(table.cell(3, num_cols-1), f"{total_kw:.1f}", is_bold=True)
        fix_cell_format(table.cell(6, num_cols-1), f"{total_kwh:,.0f}", is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！請下載檢查最後一頁的橫向表格。")
        st.download_button("📥 下載修正版 Word 報告", buf.getvalue(), "風車分析報告.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤，請檢查代碼或範本檔: {e}")
