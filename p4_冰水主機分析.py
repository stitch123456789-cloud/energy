import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體工具函數 ---
def add_run_kai(paragraph, text, size=12, is_bold=False):
    run = paragraph.add_run(text)
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    return run

# --- 2. Streamlit 介面佈局 ---
st.title("❄️ P4. 冰水主機汰換效益分析")

# 第一排：基本環境設定
c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
with c2:
    val_from_app = st.session_state.get('auto_avg_price', 4.48)
    elec_price = st.number_input("平均電費 (元/度)", value=float(val_from_app), step=0.01)
with c3:
    setup_year = st.number_input("原主機設置年份 (民國)", value=104)

# 第二排：效率基準快速設定 (連動下方表格)
st.markdown("---")
st.subheader("🎯 效率與名稱快速設定")
ca, cb, cc = st.columns(3)
with ca:
    suggest_ch_name = st.text_input("建議更換主機名稱", value="CH-1")
with cb:
    base_old_eff = st.number_input("現況夏季效率基準 (kW/RT)", value=0.95, step=0.01)
with cc:
    base_new_eff = st.number_input("改善後夏季效率基準 (kW/RT)", value=0.50, step=0.01)

# 第三排：【改善前】與【改善後】配置與數據 (並聯顯示)
st.markdown("---")
left_col, right_col = st.columns(2)

# --- 1. 初始化 Session State (加入 RT 與 台數 到運轉表格) ---
if "old_op_data" not in st.session_state:
    st.session_state.old_op_data = pd.DataFrame({
        "季節": ["春秋", "夏季", "冬季"],
        "RT": [500, 500, 500],        # 新增：讓您手動填
        "台數": [1, 1, 1],             # 新增：讓您手動填
        "時數(hr/y)": [2190, 1095, 1095],
        "負載率(%)": [60, 70, 50],
        "效率(kW/RT)": [round(base_old_eff*0.96,3), base_old_eff, round(base_old_eff*0.94,3)]
    })

if "new_op_data" not in st.session_state:
    # 預設複製改善前的所有數值
    st.session_state.new_op_data = st.session_state.old_op_data.copy()
    st.session_state.new_op_data["效率(kW/RT)"] = [round(base_new_eff*0.96,3), base_new_eff, round(base_new_eff*0.94,3)]

# --- 2. 介面佈局與同步邏輯 ---
left_col, right_col = st.columns(2)

with left_col:
    st.subheader("🧊 1. 改善前 (現況)")
    # 配置表 (僅作為文字敘述用)
    old_cfg = st.data_editor(st.session_state.old_cfg_data, num_rows="dynamic", key="old_cfg_edit")
    # 運轉表 (包含 RT 與 台數)
    old_op = st.data_editor(st.session_state.old_op_data, use_container_width=True, key="old_op_edit")

    # 同步邏輯：左邊改，右邊跟著改 (RT, 台數, 時數, 負載)
    if not old_op.equals(st.session_state.old_op_data):
        st.session_state.old_op_data = old_op
        for col in ["RT", "台數", "時數(hr/y)", "負載率(%)"]:
            st.session_state.new_op_data[col] = old_op[col]
        st.rerun()

with right_col:
    st.subheader("✨ 2. 改善後 (預期)")
    new_cfg = st.data_editor(st.session_state.new_cfg_data, num_rows="dynamic", key="new_cfg_edit")
    # 改善後運轉表 (也可獨立修改 RT 與 台數)
    new_op = st.data_editor(st.session_state.new_op_data, use_container_width=True, key="new_op_edit")
    
    st.session_state.new_op_data = new_op

# --- 3. Word 表格生成函數 (修改計算邏輯) ---
def build_word_table(doc, op_df):
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    hd = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "負載率", "耗電\n(kWh/年)"]
    for i, h in enumerate(hd):
        cp = table.cell(0,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_kai(cp, h, size=10, is_bold=True)

    total_kwh = 0
    for _, row in op_df.iterrows():
        # 重要：現在直接讀取表格內的 RT 與 台數
        current_rt = row["RT"]
        current_qty = row["台數"]
        
        # 計算：RT * 台數 * 效率 * 時數 * 負載率
        kwh = current_rt * current_qty * row["效率(kW/RT)"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
        total_kwh += kwh
        
        r_cells = table.add_row().cells
        vals = [
            row["季節"], 
            f"{current_rt:,.0f}", 
            f"{current_qty:,.0f}", 
            f"{row['效率(kW/RT)']:.3f}", 
            f"{row['時數(hr/y)']:,.0f}", 
            f"{row['負載率(%)']}%", 
            f"{kwh:,.0f}"
        ]
        for i, v in enumerate(vals):
            cp = r_cells[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run_kai(cp, v)
    
    # 合計列
    row_sum = table.add_row().cells
    row_sum[0].merge(row_sum[5])
    p_sum = row_sum[0].paragraphs[0]; p_sum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(p_sum, "總耗電量(kWh/年)", is_bold=True)
    p_val = row_sum[6].paragraphs[0]; p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(p_val, f"{total_kwh:,.0f}", is_bold=True)
    return total_kwh

# --- 呼叫函數處也要簡化 ---
total_old_kwh = build_word_table(doc, old_op)
# ... 略 ...
total_new_kwh = build_word_table(doc, new_op)

# 生成現況表格
total_old_kwh = build_word_table(doc, old_cfg, old_op)
p_old_sum = doc.add_paragraph(); p_old_sum.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p_old_sum, f"2.推估耗電量：{total_old_kwh:,.0f} kWh/年。")

# B. 二、改善方案
doc.add_paragraph()
add_run_kai(doc.add_heading('', level=1), "二、改善方案", size=14, is_bold=True)
p2 = doc.add_paragraph(); p2.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p2, f"1. 建議編列經費汰換為高效率冰水主機，目前新型高效率 1 級能效離心式冰水主機之運轉效率可達 {base_new_eff:.2f} kW/RT，如與以上大樓現況冰水主機運轉效率相比，有節能空間。")
p3 = doc.add_paragraph(); p3.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p3, f"2.參考(附件A-4)『冰水機組製冷能源效率分級基準表』，擬建議貴單位優先將現況低效率之冰水主機{suggest_ch_name}，汰換為符合建議標準的冰水主機，以節省主機運轉耗能。")

# C. 三、預期效益
doc.add_paragraph()
add_run_kai(doc.add_heading('', level=1), "三、預期效益", size=14, is_bold=True)
p_res_title = doc.add_paragraph(); p_res_title.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p_res_title, "改善後冰水主機耗電量計算如下：")

new_desc = " + ".join([f"{r['容量(RT)']}RT×{r['台數']}" for _, r in new_cfg.iterrows()])
p5 = doc.add_paragraph(); p5.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p5, f"1. 採用高效率離心式冰水主機 {new_desc} 台，推估年度耗電量如下表：")

# 生成改善後表格
total_new_kwh = build_word_table(doc, new_cfg, new_op)

# D. 最終結算
save_kwh = total_old_kwh - total_new_kwh
save_money = save_kwh * elec_price / 10000
res_p = doc.add_paragraph(); res_p.paragraph_format.first_line_indent = Pt(24)
add_run_kai(res_p, f"預估年節電量約 {save_kwh:,.0f} kWh，年節省電費約 {save_money:.1f} 萬元。")

# --- 5. 報告輸出中心 ---
st.markdown("---")
st.subheader("🚀 報告輸出中心")
buf = io.BytesIO()
doc.save(buf)
current_word_data = buf.getvalue()

col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("🔄 確認數值並同步至打包中心", use_container_width=True):
        if 'report_warehouse' not in st.session_state:
            st.session_state['report_warehouse'] = {}
        st.session_state['report_warehouse']["4. 冰水主機效益分析"] = current_word_data
        st.success("✅ 數據已鎖定！")
        st.rerun()
with col_btn2:
    st.download_button("💾 下載目前的 Word 報告", current_word_data, "冰水主機汰換效益分析.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
