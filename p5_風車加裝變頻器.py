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
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """強化版標籤替換：處理碎裂標籤與格式鎖定"""
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

# --- 2. 介面設定 (融合你指定的排版與 Dataframe) ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=50.0)
    elec_val = st.session_state.get('auto_avg_price', 3.5)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="1500RT")
    invest_amt_input = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# 動態設備管理 (側邊欄)
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

for i, t in enumerate(st.session_state.towers):
    with st.expander(f"設備組別：{t['name']}", expanded=False):
        tc1, tc2, tc3, tc4 = st.columns(4)
        t['name'] = tc1.text_input("編號", value=t['name'], key=f"n_{i}")
        t['rt'] = tc2.number_input("噸數(RT)", value=t['rt'], key=f"r_{i}")
        t['hp'] = tc3.number_input("馬力(HP)", value=t['hp'], key=f"h_{i}")
        t['fans'] = tc4.number_input("台數", min_value=1, max_value=5, value=t['fans'], key=f"f_{i}")

# --- 3. 生成與計算 ---
if st.button("🚀 生成專業效益報告", use_container_width=True):
    try:
        # 核心計算：根據馬力自動計算投資額 (1.3萬/HP)
        total_hp = sum(t['hp'] * t['fans'] for t in st.session_state.towers)
        auto_invest = total_hp * 1.3 
        
        # 季節性計算
        total_old_kwh = 0
        total_new_kwh = 0
        total_kw = total_hp * 0.746
        
        for _, row in current_op_df.iterrows():
            h = float(row["時數(hr)"])
            l = float(row["負載率(%)"]) / 100
            o_kwh = total_kw * h
            n_kwh = total_kw * (l**3) * 1.06 * h
            total_old_kwh += o_kwh
            total_new_kwh += n_kwh

        save_kwh = total_old_kwh - total_new_kwh
        save_money = save_kwh * elec_input / 10000

        doc = Document("template_p5.docx")
        
        # 文字標籤地圖
        data_map = {
            "{{UN}}": unit_name, "{{CH_INFO}}": ch_info, "{{RT_INFO}}": rt_info,
            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}", "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}", "{{INVEST}}": f"{auto_invest:.1f}",
            "{{PAYBACK}}": f"{(auto_invest/save_money if save_money > 0 else 0):.1f}"
        }
        
        safe_replace(doc, data_map)

        # 橫向合併表格 (完全對齊 CT-1 格式)
        doc.add_page_break()
        doc.add_paragraph("【表一、現況耗電明細分析表 (橫向擴展)】")
        
        # 這裡會依據你新增的水塔組別動態長大
        fan_count = sum(t['fans'] for t in st.session_state.towers)
        num_cols = 1 + fan_count + 1
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, label in enumerate(labels):
            fix_cell_font(table.cell(r, 0), is_bold=True)
            table.cell(r, 0).text = label

        col_ptr = 1
        for t in st.session_state.towers:
            f_count = t['fans']
            c_n = table.cell(0, col_ptr).merge(table.cell(0, col_ptr + f_count - 1))
            c_n.text = t['name']
            fix_cell_font(c_n, is_bold=True)
            
            c_r = table.cell(1, col_ptr).merge(table.cell(1, col_ptr + f_count - 1))
            c_r.text = f"{t['rt']}RT"
            fix_cell_font(c_r)

            for i in range(f_count):
                kw = t['hp'] * 0.746
                table.cell(2, col_ptr + i).text = f"{t['hp']:.1f}"
                table.cell(3, col_ptr + i).text = f"{kw:.1f}"
                table.cell(4, col_ptr + i).text = "4,380" # 範例
                table.cell(5, col_ptr + i).text = "100%"
                table.cell(6, col_ptr + i).text = f"{(kw*4380):,.0f}"
                for r in range(2, 7): fix_cell_font(table.cell(r, col_ptr + i), is_bold=(r==6))
            col_ptr += f_count

        buf = io.BytesIO()
        doc.save(buf)
        st.success(f"✅ 報告生成完畢！總投資額預估：{auto_invest:.1f} 萬元")
        st.download_button("📥 下載完整整合報告", buf.getvalue(), "風車分析整合版.docx")

    except Exception as e:
        st.error(f"融合出錯: {e}")
