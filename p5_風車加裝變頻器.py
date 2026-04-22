import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 (保留您原本正確的邏輯) ---

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
    """保留您原本強化的碎裂標籤替換邏輯"""
    for p in doc.paragraphs:
        inline_text = "".join([run.text for run in p.runs])
        for key, val in data_map.items():
            if key in inline_text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                    elif key[0:2] in run.text:
                        p.text = p.text.replace(key, str(val))
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.font.name = '標楷體'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_cell_text = "".join([run.text for run in p.runs])
                    for key, val in data_map.items():
                        if key in full_cell_text:
                            p.text = p.text.replace(key, str(val))
                            for run in p.runs:
                                run.font.color.rgb = RGBColor(0, 0, 0)
                                run.font.name = '標楷體'
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 介面設定 (融合手繪稿風格) ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

# A. 基礎資訊
with st.container():
    c1, c2 = st.columns(2)
    unit_name = c1.text_input("單位名稱", value="貴單位")
    avg_price = c2.number_input("平均電費 (元/度)", value=3.50, step=0.01)

# B. 設備清單管理
if "towers" not in st.session_state:
    st.session_state.towers = [{"name": "CT-1", "rt": 300, "hp": 15.0, "fans": 3}]

st.subheader("🏗️ 冷卻水塔設備清單")
for i, t in enumerate(st.session_state.towers):
    with st.expander(f"組別 {t['name']}", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        t['name'] = col1.text_input("編號", value=t['name'], key=f"n_{i}")
        t['rt'] = col2.number_input("噸數(RT)", value=t['rt'], key=f"r_{i}")
        t['hp'] = col3.number_input("馬力(HP)", value=t['hp'], key=f"h_{i}")
        t['fans'] = col4.number_input("台數", min_value=1, value=t['fans'], key=f"f_{i}")

c_add, c_del = st.columns(2)
if c_add.button("➕ 新增組別"):
    st.session_state.towers.append({"name": f"CT-{len(st.session_state.towers)+1}", "rt": 300, "hp": 15.0, "fans": 1})
    st.rerun()
if c_del.button("❌ 刪除組別"):
    if len(st.session_state.towers) > 1:
        st.session_state.towers.pop()
        st.rerun()

# C. 季節參數編輯 (您要求的變數調整處)
st.subheader("📅 季節運轉參數設定")
season_init = pd.DataFrame({
    "季節": ["夏季 (6-9月)", "春秋季 (3-5, 10-11月)", "冬季 (12-2月)"],
    "時數 (hr)": [2920, 3650, 2190],
    "負載率 (%)": [85.0, 70.0, 50.0]
})
edit_season = st.data_editor(season_init, use_container_width=True)

# --- 3. 計算與生成 ---
if st.button("🚀 生成效益報告", use_container_width=True):
    try:
        # 1. 投資額計算：抓取總馬力 * 1.3萬
        total_hp = sum(t['hp'] * t['fans'] for t in st.session_state.towers)
        invest_amt = total_hp * 1.3 
        
        # 2. 核心計算邏輯
        total_old_kwh = 0
        total_save_kwh = 0
        season_results = []
        total_kw = total_hp * 0.746
        
        for _, s in edit_season.iterrows():
            h = float(s["時數 (hr)"])
            load = float(s["負載率 (%)"]) / 100
            o_kwh = total_kw * h
            n_kwh = total_kw * (load**3) * 1.06 * h # 變頻公式
            s_kwh = o_kwh - n_kwh
            
            season_results.append({
                "季節": s["季節"], "時數": h, "負載": f"{s['負載率 (%)']}%",
                "舊": o_kwh, "新": n_kwh, "省": s_kwh
            })
            total_old_kwh += o_kwh
            total_save_kwh += s_kwh

        save_money = total_save_kwh * avg_price / 10000
        payback = invest_amt / save_money if save_money > 0 else 0

        # 3. Word 替換
        doc = Document("template_p5.docx")
        data_map = {
            "{{UN}}": unit_name, 
            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{total_save_kwh:,.0f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{SUPPRESS_KW}}": f"{(total_kw * 0.15):,.1f}"
        }
        
        safe_replace(doc, data_map)

        # 4. 生成【季節性效益明細表】(符合手繪稿要求的橫向表)
        doc.add_page_break()
        doc.add_paragraph("【表、季節性節能效益分析表】")
        table_s = doc.add_table(rows=6, cols=5)
        set_table_border(table_s)
        
        headers = ["項目", "夏季", "春秋季", "冬季", "合計"]
        for i, h in enumerate(headers):
            table_s.cell(0, i).text = h
            fix_cell_font(table_s.cell(0, i), is_bold=True)

        labels = ["運轉時數(hr)", "平均負載(%)", "現況耗電(kWh)", "變頻耗電(kWh)", "節電量(kWh)"]
        for r_idx, label in enumerate(labels, 1):
            table_s.cell(r_idx, 0).text = label
            fix_cell_font(table_s.cell(r_idx, 0))
            
            row_sum = 0
            for c_idx in range(1, 4):
                res = season_results[c_idx-1]
                val = ""
                if r_idx == 1: val = f"{res['時數']:,.0f}"; row_sum += res['時數']
                elif r_idx == 2: val = res['負載']
                elif r_idx == 3: val = f"{res['舊']:,.0f}"; row_sum += res['舊']
                elif r_idx == 4: val = f"{res['新']:,.0f}"; row_sum += res['新']
                elif r_idx == 5: val = f"{res['省']:,.0f}"; row_sum += res['省']
                
                table_s.cell(r_idx, c_idx).text = val
                fix_cell_font(table_s.cell(r_idx, c_idx))
            
            if r_idx != 2:
                table_s.cell(r_idx, 4).text = f"{row_sum:,.0f}"
                fix_cell_font(table_s.cell(r_idx, 4), is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success(f"✅ 生成成功！總投資額為：{invest_amt:.1f} 萬元")
        st.download_button("📥 下載專業報告", buf.getvalue(), "風車變頻效益分析.docx")

    except Exception as e:
        st.error(f"執行發生錯誤: {e}")
