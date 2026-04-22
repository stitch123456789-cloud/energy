import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 格式與替換核心工具 ---
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
    """鎖定：標楷體、黑色、12號、置中"""
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 
        if not paragraph.runs: paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        for key, val in data_map.items():
            if key in full_text:
                full_text = full_text.replace(key, str(val))
                p.clear()
                new_run = p.add_run(full_text)
                new_run.font.name = '標楷體'
                new_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                new_run.font.size = Pt(12)
                new_run.font.color.rgb = RGBColor(0, 0, 0)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_text = "".join(run.text for run in p.runs)
                    for key, val in data_map.items():
                        if key in full_text:
                            full_text = full_text.replace(key, str(val))
                            p.text = full_text
                            for run in p.runs:
                                run.font.name = '標楷體'
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                                run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. Streamlit 介面設計 ---
st.set_page_config(layout="wide") # 使用寬螢幕模式更適合多台設備排版
st.title("🌀 P5. 冷卻水塔風車變頻專業分析")

# 初始化 Session State
if "towers" not in st.session_state:
    st.session_state.towers = [{"name": "CT-1", "rt": 300, "hp": 15.0, "fans": 3}]

# A. 基礎參數
col_a, col_b, col_c = st.columns(3)
with col_a:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with col_b:
    motor_hp = st.number_input("基準馬力 (HP)", value=50.0)
    elec_val = st.number_input("平均電費 (元/度)", value=3.5, step=0.1)
with col_c:
    rt_info = st.text_input("水塔總容量", value="1500RT")
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

# B. 設備管理 (側邊欄)
with st.sidebar:
    st.header("⚙️ 設備管理")
    if st.button("➕ 新增一組水塔"):
        idx = len(st.session_state.towers) + 1
        st.session_state.towers.append({"name": f"CT-{idx}", "rt": 300, "hp": 15.0, "fans": 1})
        st.rerun()
    if st.button("❌ 刪除最後一組"):
        if len(st.session_state.towers) > 1:
            st.session_state.towers.pop(); st.rerun()

# C. 顯示並編輯設備
for i, t in enumerate(st.session_state.towers):
    with st.expander(f"配置：{t['name']}", expanded=True):
        tc1, tc2, tc3, tc4 = st.columns(4)
        t['name'] = tc1.text_input("編號", value=t['name'], key=f"n_{i}")
        t['rt'] = tc2.number_input("噸數(RT)", value=t['rt'], key=f"r_{i}")
        t['hp'] = tc3.number_input("馬力(HP)", value=t['hp'], key=f"h_{i}")
        t['fans'] = tc4.number_input("台數", min_value=1, max_value=5, value=t['fans'], key=f"f_{i}")

st.divider()

# D. 改善後參數設定 (直式塊狀輸入)
st.subheader("⚙️ 改善後：各台風扇運轉參數設定")
st.info("請設定每一台風扇在不同季節的運轉參數（對齊表二格式）")

after_config_results = []
for t in st.session_state.towers:
    f_count = int(t['fans'])
    st.markdown(f"#### 🏗️ 組別：{t['name']}")
    cols = st.columns(f_count)
    for i in range(f_count):
        with cols[i]:
            # 讓使用者可以手動定義風扇名稱
            fan_name = st.text_input("風扇名稱", value=f"{t['name']}-F{i+1}", key=f"fname_{t['name']}_{i}")
            
            sp_l = st.number_input(f"春秋負載(%)", value=70, key=f"sp_l_{t['name']}_{i}")
            sp_h = st.number_input(f"春秋時數(hr)", value=4380, key=f"sp_h_{t['name']}_{i}")
            su_l = st.number_input(f"夏季負載(%)", value=85, key=f"su_l_{t['name']}_{i}")
            su_h = st.number_input(f"夏季時數(hr)", value=2190, key=f"su_h_{t['name']}_{i}")
            wi_l = st.number_input(f"冬季負載(%)", value=60, key=f"wi_l_{t['name']}_{i}")
            wi_h = st.number_input(f"冬季時數(hr)", value=2190, key=f"wi_h_{t['name']}_{i}")
            
            after_config_results.append({
                "name": fan_name,
                "parent_name": t['name'],
                "hp": t['hp'],
                "sp_l": sp_l, "sp_h": sp_h,
                "su_l": su_l, "su_h": su_h,
                "wi_l": wi_l, "wi_h": wi_h
            })
    st.divider()

# --- 3. 生成按鈕與核心邏輯 ---


# --- 3. 生成按鈕與核心邏輯 ---
if st.button("🚀 生成專業效益報告", use_container_width=True):
    try:
        # 1. 核心計算 (與文字替換標籤連動)
        total_hp = sum(t['hp'] * t['fans'] for t in st.session_state.towers)
        calc_invest = total_hp * 1.3
        
        # 這裡的計算需要與 after_config_results 同步
        total_old_kwh = sum((t['hp']*0.746) * 8760 * t['fans'] for t in st.session_state.towers) # 簡化假設
        total_after_kwh = 0
        for fan in after_config_results:
            kw_b = fan['hp'] * 0.746
            total_after_kwh += (kw_b * (fan['sp_l']/100)**3 * 1.06 * fan['sp_h'])
            total_after_kwh += (kw_b * (fan['su_l']/100)**3 * 1.06 * fan['su_h'])
            total_after_kwh += (kw_b * (fan['wi_l']/100)**3 * 1.06 * fan['wi_h'])
        
        save_kwh = total_old_kwh - total_after_kwh
        save_money = save_kwh * elec_val / 10000
        payback = calc_invest / save_money if save_money > 0 else 0

        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name,
            "{{COUNT}}": str(len(st.session_state.towers)),
            "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info,
            "{{MT}}": f"{motor_hp}hp",
            "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{MOTOR_SPEC}}": f"{motor_hp}HP x {int(total_hp/motor_hp)}台",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{INVEST}}": f"{calc_invest:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}",
            "{{SUPPRESS_KW}}": f"{total_hp * 0.746 * 0.15:.1f}"
        }
        safe_replace(doc, data_map)

        # 2. 生成表一 (現況)
        doc.add_page_break()
        doc.add_paragraph("【表一、現況耗電明細分析表】")
        num_fans = len(after_config_results)
        num_cols = 1 + num_fans + 1
        table1 = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table1)
        
        labels1 = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r, txt in enumerate(labels1): 
            table1.cell(r, 0).text = txt
            fix_cell_font(table1.cell(r, 0), is_bold=True)

        col_ptr = 1
        total_old_kw_sum = 0
        total_old_kwh_sum = 0
        for t in st.session_state.towers:
            f_count = t['fans']
            c_n = table1.cell(0, col_ptr).merge(table1.cell(0, col_ptr + f_count - 1))
            c_n.text = t['name']; fix_cell_font(c_n, is_bold=True)
            for i in range(f_count):
                kw_f = t['hp'] * 0.746
                kwh_f = kw_f * 8760
                table1.cell(2, col_ptr+i).text = f"{t['hp']}"
                table1.cell(3, col_ptr+i).text = f"{kw_f:.1f}"
                table1.cell(4, col_ptr+i).text = "8,760"
                table1.cell(5, col_ptr+i).text = "100%"
                table1.cell(6, col_ptr+i).text = f"{kwh_f:,.0f}"
                total_old_kw_sum += kw_f
                total_old_kwh_sum += kwh_f
                for r in range(1, 7): fix_cell_font(table1.cell(r, col_ptr+i))
            col_ptr += f_count
            
        table1.cell(0, num_cols-1).text = "合計"
        table1.cell(3, num_cols-1).text = f"{total_old_kw_sum:.1f}"
        table1.cell(6, num_cols-1).text = f"{total_old_kwh_sum:,.0f}"
        for r in [0, 3, 6]: fix_cell_font(table1.cell(r, num_cols-1), is_bold=True)

        # 3. 生成表二 (改善後)
        doc.add_paragraph("\n【表二、改善後變頻節能效益分析表】")
        table2 = doc.add_table(rows=20, cols=num_cols)
        set_table_border(table2)
        
        labels2 = ["編號", "春秋季負載率(%)", "夏季負載率(%)", "冬季負載率(%)", "春秋季使用時數(hr)", "夏季使用時數(hr)", "冬季使用時數(hr)", "春秋季耗功(kW)", "夏季耗功(kW)", "冬季耗功(kW)", "春秋季耗電(kWh)", "夏季耗電(kWh)", "冬季耗電(kWh)", "改善後總耗電(kWh)", "節省耗電(kWh)", "抑低需量(kW)", "節約金額(萬元/年)", "節能比(%)", "註1：變頻器損失", "註2：平均電費(元/kWh)"]
        for r, txt in enumerate(labels2):
            table2.cell(r, 0).text = txt
            fix_cell_font(table2.cell(r, 0), is_bold=True)

        final_after_kwh = 0
        final_save_kwh = 0
        final_suppress_kw = 0

        # 填寫每一台數據
        # 填寫每一台數據
        for idx, fan in enumerate(after_config_results):
            c_idx = idx + 1
            kw_b = fan['hp'] * 0.746
                                
            # 計算數據
            sp_kw = kw_b * (fan['sp_l']/100)**3 * 1.06
            su_kw = kw_b * (fan['su_l']/100)**3 * 1.06
            wi_kw = kw_b * (fan['wi_l']/100)**3 * 1.06
            
            sp_kwh = sp_kw * fan['sp_h']
            su_kwh = su_kw * fan['su_h']
            wi_kwh = wi_kw * fan['wi_h']
            
            f_after_kwh = sp_kwh + su_kwh + wi_kwh
            f_old_kwh = kw_b * (fan['sp_h'] + fan['su_h'] + fan['wi_h'])
            f_save_kwh = f_old_kwh - f_after_kwh
            f_suppress = kw_b * 0.15
            
            # 填表
            table2.cell(0, c_idx).text = fan['name']  # ✨ 確保這一行是寫 fan['name']
            table2.cell(1, c_idx).text = f"{fan['sp_l']}%"
            table2.cell(2, c_idx).text = f"{fan['su_l']}%"
            table2.cell(3, c_idx).text = f"{fan['wi_l']}%"
            table2.cell(4, c_idx).text = f"{fan['sp_h']:,.0f}"
            table2.cell(5, c_idx).text = f"{fan['su_h']:,.0f}"
            table2.cell(6, c_idx).text = f"{fan['wi_h']:,.0f}"
            table2.cell(7, c_idx).text = f"{sp_kw:.2f}"
            table2.cell(8, c_idx).text = f"{su_kw:.2f}"
            table2.cell(9, c_idx).text = f"{wi_kw:.2f}"
            table2.cell(10, c_idx).text = f"{sp_kwh:,.0f}"
            table2.cell(11, c_idx).text = f"{su_kwh:,.0f}"
            table2.cell(12, c_idx).text = f"{wi_kwh:,.0f}"
            table2.cell(13, c_idx).text = f"{f_after_kwh:,.0f}"
            table2.cell(14, c_idx).text = f"{f_save_kwh:,.0f}"
            table2.cell(15, c_idx).text = f"{f_suppress:.1f}"
            
            final_after_kwh += f_after_kwh
            final_save_kwh += f_save_kwh
            final_suppress_kw += f_suppress
            for r in range(16): fix_cell_font(table2.cell(r, c_idx))

        # 表二合計與底部合併
        table2.cell(0, num_cols-1).text = "合計"
        table2.cell(13, num_cols-1).text = f"{final_after_kwh:,.0f}"
        table2.cell(14, num_cols-1).text = f"{final_save_kwh:,.0f}"
        table2.cell(15, num_cols-1).text = f"{final_suppress_kw:.1f}"
        for r in [0, 13, 14, 15]: fix_cell_font(table2.cell(r, num_cols-1), is_bold=True)
        
        # 底部合併欄位 (16-19列)
        for row_idx in range(16, 20):
            merged = table2.cell(row_idx, 1).merge(table2.cell(row_idx, num_cols-1))
            if row_idx == 16: merged.text = f"{(final_save_kwh * elec_val / 10000):.1f}"
            elif row_idx == 17: merged.text = f"{(final_save_kwh / (final_after_kwh + final_save_kwh) * 100):.1f}%"
            elif row_idx == 18: merged.text = "6.0%"
            elif row_idx == 19: merged.text = f"{elec_val:.2f}"
            fix_cell_font(merged, is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！")
        st.download_button("📥 下載完整報告", buf.getvalue(), "風車分析專業版.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤: {e}")
