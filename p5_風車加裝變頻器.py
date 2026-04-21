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

def fix_style(obj):
    """強效格式化：黑色、標楷體、12號、不加粗、置中"""
    for paragraph in obj.paragraphs:
        paragraph.alignment = 1 
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(12)
            run.font.bold = False
            run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面設定：動態組數增減 ---
st.title("🌀 P5. 冷卻水塔風車變頻專業分析")

# 初始化 session_state
if "tower_groups" not in st.session_state:
    st.session_state.tower_groups = [{"name": "CT-1", "rt": 300, "fans": 3, "hp": 15.0}]

# 側邊欄控制增減
with st.sidebar:
    st.header("⚙️ 設備管理")
    if st.button("➕ 新增一組水塔"):
        new_name = f"CT-{len(st.session_state.tower_groups)+1}"
        st.session_state.tower_groups.append({"name": new_name, "rt": 300, "fans": 1, "hp": 15.0})
        st.rerun()
    if st.button("❌ 刪除最後一組"):
        if len(st.session_state.tower_groups) > 1:
            st.session_state.tower_groups.pop()
            st.rerun()

# 顯示與編輯組別
for i, group in enumerate(st.session_state.tower_groups):
    with st.expander(f"冷卻水塔組別配置：{group['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        group['name'] = c1.text_input("編號", value=group['name'], key=f"n_{i}")
        group['rt'] = c2.number_input("噸數(RT)", value=group['rt'], key=f"r_{i}")
        group['hp'] = c3.number_input("馬力(HP)", value=group['hp'], key=f"h_{i}")
        group['fans'] = c4.number_input("風扇數", min_value=1, max_value=5, value=group['fans'], key=f"f_{i}")

# 建立明細表格
st.subheader("📊 運轉參數明細設定")
rows = []
for g in st.session_state.tower_groups:
    for f in range(1, g['fans'] + 1):
        rows.append({"組別": g['name'], "編號": f"{g['name']}-F{f}", "時數(hr)": 4380, "負載(%)": 100})
edit_df = st.data_editor(pd.DataFrame(rows), use_container_width=True, key="main_editor")

# --- 3. 生成邏輯 ---
if st.button("🚀 生成 P5 專業報告", use_container_width=True):
    try:
        # A. 數據預計算
        total_old_kwh = 0
        total_kw = 0
        fan_data_list = []
        
        curr_row = 0
        for g in st.session_state.tower_groups:
            for f in range(g['fans']):
                h = float(edit_df.iloc[curr_row]["時數(hr)"])
                kw = g['hp'] * 0.746
                kwh = kw * h
                fan_data_list.append({"h": h, "kw": kw, "kwh": kwh})
                total_old_kwh += kwh
                total_kw += kw
                curr_row += 1

        # B. 準備替換資料 (這部分標籤要精準)
        save_kwh = total_old_kwh * 0.4 # 假設值
        save_money = save_kwh * 3.5 / 10000
        
        data_map = {
            "{{UN}}": "貴單位", 
            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{INVEST}}": "80.0",
            "{{PAYBACK}}": "1.2",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{SUPPRESS_KW}}": "13"
        }

        doc = Document("template_p5.docx")

        # C. 強化版文字替換：暴力掃描
        for p in doc.paragraphs:
            # 獲取段落完整文字，防止被 Run 切斷
            p_text = "".join(run.text for run in p.runs)
            original_style_run = p.runs[0] if p.runs else None
            
            for k, v in data_map.items():
                if k in p_text:
                    p_text = p_text.replace(k, str(v))
                    # 重寫段落
                    p.clear()
                    new_run = p.add_run(p_text)
                    # 強制修正格式
                    new_run.font.name = '標楷體'
                    new_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                    new_run.font.size = Pt(12)
                    new_run.font.color.rgb = RGBColor(0, 0, 0)
        
        # D. 橫向表格生成 (完全對齊截圖)
        doc.add_page_break()
        doc.add_paragraph("--- 自動生成的現況明細表 ---")
        
        num_fans = len(fan_data_list)
        num_cols = 1 + num_fans + 1
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫標題列
        headers = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, txt in enumerate(headers):
            fix_style(table.cell(r, 0))
            table.cell(r, 0).text = txt

        # 橫向填寫組別資料與合併
        col_ptr = 1
        for g in st.session_state.tower_groups:
            f_count = g['fans']
            # 編號合併
            c_n = table.cell(0, col_ptr)
            if f_count > 1: c_n = c_n.merge(table.cell(0, col_ptr + f_count - 1))
            c_n.text = g['name']
            fix_style(c_n)
            
            # RT合併
            c_r = table.cell(1, col_ptr)
            if f_count > 1: c_r = c_r.merge(table.cell(1, col_ptr + f_count - 1))
            c_r.text = f"{g['rt']}RT"
            fix_style(c_r)

            # 個別風扇數據
            for i in range(f_count):
                idx = col_ptr - 1 + i
                d = fan_data_list[idx]
                table.cell(2, col_ptr + i).text = f"{g['hp']:.1f}"
                table.cell(3, col_ptr + i).text = f"{d['kw']:.1f}"
                table.cell(4, col_ptr + i).text = f"{d['h']:,.0f}"
                table.cell(5, col_ptr + i).text = "100%"
                table.cell(6, col_ptr + i).text = f"{d['kwh']:,.0f}"
                for r in range(2, 7): fix_style(table.cell(r, col_ptr + i))
            col_ptr += f_count

        # 最後合计欄
        table.cell(0, num_cols-1).text = "合計"
        table.cell(3, num_cols-1).text = f"{total_kw:.1f}"
        table.cell(6, num_cols-1).text = f"{total_old_kwh:,.0f}"
        for r in range(7): fix_style(table.cell(r, num_cols-1))

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！文字與動態表格均已修正。")
        st.download_button("📥 下載修正版報告", buf.getvalue(), "P5_風車分析修正版.docx")

    except Exception as e:
        st.error(f"❌ 發生致命錯誤: {e}")
