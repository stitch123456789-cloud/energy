import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
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

# 第一排：基本參數
col1_1, col1_2, col1_3 = st.columns(3)
with col1_1:
    unit_name = st.text_input("單位名稱", value="貴單位")
with col1_2:
    val_from_app = st.session_state.get('auto_avg_price', 4.48)
    elec_price = st.number_input("平均電費 (元/度)", value=float(val_from_app), step=0.01)
with col1_3:
    setup_year = st.number_input("原主機設置年份 (民國)", value=104)

# 第二排：現有主機配置
st.subheader("🧊 現有主機配置")
df_init = pd.DataFrame([
    {"編號": "CH-1", "台數": 1, "容量(RT)": 350, "型式": "螺旋式"},
    {"編號": "CH-2", "台數": 1, "容量(RT)": 350, "型式": "螺旋式"}
])
chiller_config = st.data_editor(df_init, num_rows="dynamic", use_container_width=True, key="ch_cfg_editor")

# 第三排：改善方案與效率基準設定
st.subheader("⚙️ 效率基準與改善設定")
col2_1, col2_2, col2_3, col2_4 = st.columns(4)
with col2_1:
    suggest_ch_name = st.text_input("建議更換主機名稱", value="CH-1")
with col2_2:
    base_old_eff = st.number_input("現況夏季效率基準 (kW/RT)", value=0.95, step=0.01)
with col2_3:
    base_new_eff = st.number_input("改善夏季效率基準 (kW/RT)", value=0.50, step=0.01)
with col2_4:
    # 新增：讓您可以改改善後的台數
    new_ch_qty = st.number_input("改善後汰換台數", value=1, min_value=1)

# 第四排：核心運轉數據編輯區
st.subheader("📅 運轉參數預覽")
op_data = {
    "季節": ["春秋", "夏季", "冬季"],
    "時數(hr/y)": [2190, 1095, 1095],
    "平均負載率(%)": [60, 70, 50],
    "現況kW/RT": [round(base_old_eff * 0.96, 3), base_old_eff, round(base_old_eff * 0.94, 3)],
    "改善後kW/RT": [round(base_new_eff * 0.96, 3), base_new_eff, round(base_new_eff * 0.94, 3)] 
}
df_op = st.data_editor(pd.DataFrame(op_data), use_container_width=True, key="op_cfg_editor")

# ---------------------------------------------------------
# --- 3. Word 生成邏輯 ---
# ---------------------------------------------------------
doc = Document()

# 數據預處理
desc_list = [f"{r['台數']}台{r['容量(RT)']}RT {r['型式']}" for _, r in chiller_config.iterrows()]
chiller_desc_text = "、".join(desc_list)
base_rt = chiller_config.iloc[0]['容量(RT)'] if not chiller_config.empty else 0

# --- A. 一、現況說明 ---
add_run_kai(doc.add_heading('', level=1), "一、現況說明", size=14, is_bold=True)
p1 = doc.add_paragraph()
p1.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p1, f"1. {unit_name}空調系統有{chiller_desc_text}冰水主機(設置年份{setup_year}年)，推估年度耗電量如下表：")

# 表格生成函數
def create_energy_table(doc, data_df, rt_val, qty, mode="old"):
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    headers = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "負載率", "耗電\n(kWh/年)"]
    for i, h in enumerate(headers):
        cp = table.cell(0,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_kai(cp, h, size=10, is_bold=True)

    total_kwh = 0
    eff_col = "現況kW/RT" if mode == "old" else "改善後kW/RT"
    for _, row in data_df.iterrows():
        # 使用傳入的 qty (現況通常是1, 改善後依輸入)
        kwh = rt_val * qty * row[eff_col] * row["時數(hr/y)"] * (row["平均負載率(%)"]/100)
        total_kwh += kwh
        r_cells = table.add_row().cells
        vals = [row["季節"], str(rt_val), str(qty), f"{row[eff_col]:.3f}", f"{row['時數(hr/y)']:,.0f}", f"{row['平均負載率(%)']}%", f"{kwh:,.0f}"]
        for i, v in enumerate(vals):
            cp = r_cells[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run_kai(cp, v)
            
    row_sum = table.add_row().cells
    row_sum[0].merge(row_sum[5])
    p_sum = row_sum[0].paragraphs[0]; p_sum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(p_sum, "總耗電量(kWh/年)", is_bold=True)
    p_val = row_sum[6].paragraphs[0]; p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(p_val, f"{total_kwh:,.0f}", is_bold=True)
    return total_kwh

# 插入現況表格 (現況台數固定為 1)
total_old_kwh = create_energy_table(doc, df_op, base_rt, 1, mode="old")

# 修正：補上「2.推估耗電量」字樣
p_old_sum = doc.add_paragraph()
p_old_sum.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p_old_sum, f"2.推估耗電量：{total_old_kwh:,.0f} kWh/年。")

# --- B. 二、改善方案 ---
doc.add_paragraph()
add_run_kai(doc.add_heading('', level=1), "二、改善方案", size=14, is_bold=True)
p2 = doc.add_paragraph(); p2.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p2, f"1. 建議編列經費汰換為高效率冰水主機，目前新型高效率 1 級能效離心式冰水主機之運轉效率可達 {base_new_eff:.2f} kW/RT，如與以上大樓現況冰水主機運轉效率相比，有節能空間。")
p3 = doc.add_paragraph(); p3.paragraph_format.first_line_indent = Pt(24)
# 修正文字：完全依照您要求的附件 A-4 文字
add_run_kai(p3, f"2.參考(附件A-4)『冰水機組製冷能源效率分級基準表』，擬建議貴單位優先將現況低效率之冰水主機{suggest_ch_name}，汰換為符合建議標準的冰水主機，以節省主機運轉耗能。")

# --- C. 三、預期效益 ---
doc.add_paragraph()
add_run_kai(doc.add_heading('', level=1), "三、預期效益", size=14, is_bold=True)
doc.add_paragraph().paragraph_format.first_line_indent = Pt(24)
add_run_kai(doc.paragraphs[-1], "改善後冰水主機耗電量計算如下：")
p5 = doc.add_paragraph(); p5.paragraph_format.first_line_indent = Pt(24)
# 修正：採用指定的台數
add_run_kai(p5, f"1. 採用高效率離心式冰水主機 {base_rt}RT×{new_ch_qty} 台，推估年度耗電量如下表：")

# 插入改善後表格 (使用 new_ch_qty)
total_new_kwh = create_energy_table(doc, df_op, base_rt, new_ch_qty, mode="new")

# --- D. 效益結算 ---
save_kwh = total_old_kwh - total_new_kwh
save_money = save_kwh * elec_price / 10000
res_p = doc.add_paragraph(); res_p.paragraph_format.first_line_indent = Pt(24)
add_run_kai(res_p, f"預估年節電量約 {save_kwh:,.0f} kWh，年節省電費約 {save_money:.1f} 萬元。")

# --- 4. 輸出中心 ---
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
        st.success("✅ 數據已同步！")
        st.rerun()
with col_btn2:
    st.download_button("💾 下載目前的 Word 報告", current_word_data, "冰水主機汰換效益分析.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
