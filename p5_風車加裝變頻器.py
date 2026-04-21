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

def fix_cell(cell, text, is_bold=False, size=10):
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.alignment = 1 # 置中
    p.clear()
    run = p.add_run(str(text))
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面設計：多組動態設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻專業分析系統")

if "tower_configs" not in st.session_state:
    st.session_state.tower_configs = [{"name": "CT-1", "rt": 300, "fans": 3, "hp": 15.0}]

st.subheader("🏗️ 冷卻水塔配置")
for i, config in enumerate(st.session_state.tower_configs):
    with st.expander(f"配置組別：{config['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        config['name'] = c1.text_input("水塔編號", value=config['name'], key=f"name_{i}")
        config['rt'] = c2.number_input("散熱噸數(RT)", value=config['rt'], key=f"rt_{i}")
        config['hp'] = c3.number_input("風扇馬力(HP)", value=config['hp'], key=f"hp_{i}")
        config['fans'] = c4.number_input("風扇台數", min_value=1, max_value=5, value=config['fans'], key=f"fans_{i}")

if st.button("➕ 新增水塔組別"):
    st.session_state.tower_configs.append({"name": f"CT-{len(st.session_state.tower_configs)+1}", "rt": 300, "fans": 1, "hp": 15.0})
    st.rerun()

st.subheader("📅 運轉時數設定")
# 建立一個動態表格，讓使用者手動輸入每一台的時數
rows = []
for config in st.session_state.tower_configs:
    for f_idx in range(1, config['fans'] + 1):
        rows.append({"水塔": config['name'], "風扇編號": f"{config['name']}-F{f_idx}", "運轉時數(hr)": 4380, "負載率(%)": 100})

hours_df = st.data_editor(pd.DataFrame(rows), use_container_width=True)

# --- 3. 生成 Word 報告 ---
if st.button("🚀 生成專業核核報告"):
    try:
        doc = Document("template_p5.docx")
        
        # 文字替換邏輯 (保留您原本的部分)
        # ... (此處可加入原本的 safe_replace)

        doc.add_page_break()
        doc.add_paragraph("【表一、現況耗電明細分析表】", style='Normal')

        # 計算總欄數：1 (標題欄) + 所有風扇台數 + 1 (合計欄)
        total_fans = int(hours_df.shape[0])
        num_cols = 1 + total_fans + 1
        
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r_idx, label in enumerate(labels):
            fix_cell(table.cell(r_idx, 0), label, is_bold=True)

        # 填寫資料與合併儲存格
        current_col = 1
        grand_total_kwh = 0
        
        for config in st.session_state.tower_configs:
            fans_in_group = config['fans']
            # 1. 處理編號列與合併
            start_cell = table.cell(0, current_col)
            end_cell = table.cell(0, current_col + fans_in_group - 1)
            merged_header = start_cell.merge(end_cell)
            fix_cell(merged_header, config['name'], is_bold=True)
            
            # 2. 處理 RT 列與合併
            rt_start = table.cell(1, current_col)
            rt_end = table.cell(1, current_col + fans_in_group - 1)
            merged_rt = rt_start.merge(rt_end)
            fix_cell(merged_rt, config['rt'])

            # 3. 填寫該組風扇的個別數據
            group_kwh = 0
            for f_idx in range(fans_in_group):
                col_idx = current_col + f_idx
                # 從 hours_df 找出對應資料
                fan_data = hours_df.iloc[col_idx - 1]
                h = float(fan_data["運轉時數(hr)"])
                kw = config['hp'] * 0.746
                kwh = kw * h
                group_kwh += kwh
                
                fix_cell(table.cell(2, col_idx), f"{config['hp']:.1f}")
                fix_cell(table.cell(3, col_idx), f"{kw:.2f}")
                fix_cell(table.cell(4, col_idx), f"{h:,.0f}")
                fix_cell(table.cell(5, col_idx), f"{fan_data['負載率(%)']}%")
                fix_cell(table.cell(6, col_idx), f"{kwh:,.0f}", is_bold=True)
            
            current_col += fans_in_group
            grand_total_kwh += group_kwh

        # 4. 填寫最後一欄：合計
        fix_cell(table.cell(0, num_cols-1), "合計", is_bold=True)
        # 實際耗功合計與總耗電合計
        fix_cell(table.cell(3, num_cols-1), f"{(total_fans * config['hp'] * 0.746):.1f}")
        fix_cell(table.cell(6, num_cols-1), f"{grand_total_kwh:,.0f}", is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告已生成！支援多組水塔合併與自定義風扇台數。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車多組分析.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤: {e}")
