import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體工具函數 ---
def add_run_kai(paragraph, text, size=12, is_bold=False, is_red=False):
    run = paragraph.add_run(text)
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    if is_red:
        run.font.color.rgb = RGBColor(255, 0, 0)
    else:
        run.font.color.rgb = RGBColor(0, 0, 0)
    return run

# --- 2. Streamlit 介面佈局 ---
st.title("❄️ P4. 冰水主機汰換效益分析")

# 第一排：橫向三格
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

# 第三排：橫向二格 (改善後參數)
st.subheader("⚙️ 改善後配置參數")
col2_1, col2_2 = st.columns(2)
with col2_1:
    target_eff = st.number_input("預期改善後效率 (kW/RT)", value=0.50, step=0.01)
with col2_2:
    suggest_ch_name = st.text_input("建議更換主機名稱", value="冰水主機(CH-1)")

# 第四排：運轉參數設定 (連動表格)
st.subheader("📅 運轉參數設定")
op_data = {
    "季節": ["春秋", "夏季", "冬季"],
    "時數(hr/y)": [2190, 1095, 1095],
    "平均負載率(%)": [60, 70, 50],
    "現況kW/RT": [0.864, 0.90, 0.846],
    "改善後kW/RT": [target_eff, target_eff, target_eff] 
}
df_op = st.data_editor(pd.DataFrame(op_data), use_container_width=True, key="op_cfg_editor")

# ---------------------------------------------------------
# --- 3. Word 生成與計算邏輯 ---
# ---------------------------------------------------------
doc = Document()

# A. 數據預處理
desc_list = []
total_rt_sum = 0
for _, r in chiller_config.iterrows():
    desc_list.append(f"{r['台數']}台{r['容量(RT)']}RT {r['型式']}")
    total_rt_sum += r['容量(RT)']
chiller_desc_text = "、".join(desc_list)
# 取得第一台主機容量作為表格基準
base_rt = chiller_config.iloc[0]['容量(RT)'] if not chiller_config.empty else 0

# B. 一、現況說明
add_run_kai(doc.add_paragraph(), "一、現況說明", size=14, is_bold=True)
p1 = doc.add_paragraph()
p1.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p1, f"1. {unit_name}空調系統有")
add_run_kai(p1, chiller_desc_text, is_red=True)
add_run_kai(p1, "冰水主機(設置年份")
add_run_kai(p1, str(setup_year), is_red=True)
add_run_kai(p1, "年)，推估年度耗電量如下表：")

# --- C. 改善前耗能推估表 (新增) ---
t1 = doc.add_table(rows=1, cols=6); t1.style = 'Table Grid'
h_names = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "負載率", "耗電\n(kWh/年)"]
# 修正為 7 欄以符合截圖需求
t1 = doc.add_table(rows=1, cols=7); t1.style = 'Table Grid'
headers = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "負載率", "耗電\n(kWh/年)"]
for i, h in enumerate(headers):
    cp = t1.cell(0,i).paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(cp, h, size=10, is_bold=True)

total_old_kwh = 0
for _, row in df_op.iterrows():
    # 計算：RT * 台數(1) * kW/RT * 時數 * 負載率
    kwh = base_rt * 1 * row["現況kW/RT"] * row["時數(hr/y)"] * (row["平均負載率(%)"]/100)
    total_old_kwh += kwh
    r_cells = t1.add_row().cells
    add_run_kai(r_cells[0].paragraphs[0], row["季節"])
    add_run_kai(r_cells[1].paragraphs[0], str(base_rt), is_red=True)
    add_run_kai(r_cells[2].paragraphs[0], "1", is_red=True)
    add_run_kai(r_cells[3].paragraphs[0], f"{row['現況kW/RT']:.3f}", is_red=True)
    add_run_kai(r_cells[4].paragraphs[0], str(row["時數(hr/y)"]), is_red=True)
    add_run_kai(r_cells[5].paragraphs[0], f"{row['平均負載率(%)']}%", is_red=True)
    add_run_kai(r_cells[6].paragraphs[0], f"{kwh:,.0f}", is_red=True)

# 合計列
row_sum1 = t1.add_row().cells
row_sum1[0].merge(row_sum1[5])
add_run_kai(row_sum1[0].paragraphs[0], "總耗電量(kWh/年)", is_bold=True)
add_run_kai(row_sum1[6].paragraphs[0], f"{total_old_kwh:,.0f}", is_red=True, is_bold=True)

# D. 二、改善方案
doc.add_paragraph()
add_run_kai(doc.add_paragraph(), "二、改善方案", size=14, is_bold=True)
p2 = doc.add_paragraph()
p2.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p2, "1. 建議編列經費汰換為高效率冰水主機，目前新型高效率 1 級能效離心式冰水主機之運轉效率可達 ")
add_run_kai(p2, f"{target_eff:.2f}", is_red=True)
add_run_kai(p2, " kW/RT，如與以上大樓現況冰水主機運轉效率相比，有節能空間。")
p3 = doc.add_paragraph()
p3.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p3, "2. 參考附件建議優先將現況低效率之")
add_run_kai(p3, suggest_ch_name, is_red=True)
add_run_kai(p3, "，汰換為符合建議標準之冰水主機，以節省主機運轉耗能。")

# E. 三、預期效益
doc.add_paragraph()
add_run_kai(doc.add_paragraph(), "三、預期效益", size=14, is_bold=True)
p4 = doc.add_paragraph(); p4.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p4, "改善後冰水主機耗電量計算如下：")
p5 = doc.add_paragraph(); p5.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p5, f"1. 採用高效率離心式冰水主機 ")
add_run_kai(p5, f"{base_rt}RT×1", is_red=True)
add_run_kai(p5, " 台，推估年度耗電量如下表：")

# --- 改善後耗能推估表 ---
t2 = doc.add_table(rows=1, cols=7); t2.style = 'Table Grid'
for i, h in enumerate(headers):
    cp = t2.cell(0,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_kai(cp, h, size=10, is_bold=True)

total_new_kwh = 0
for _, row in df_op.iterrows():
    kwh_new = base_rt * 1 * row["改善後kW/RT"] * row["時數(hr/y)"] * (row["平均負載率(%)"]/100)
    total_new_kwh += kwh_new
    r_cells = t2.add_row().cells
    add_run_kai(r_cells[0].paragraphs[0], row["季節"])
    add_run_kai(r_cells[1].paragraphs[0], str(base_rt), is_red=True)
    add_run_kai(r_cells[2].paragraphs[0], "1", is_red=True)
    add_run_kai(r_cells[3].paragraphs[0], f"{row['改善後kW/RT']:.3f}", is_red=True)
    add_run_kai(r_cells[4].paragraphs[0], str(row["時數(hr/y)"]), is_red=True)
    add_run_kai(r_cells[5].paragraphs[0], f"{row['平均負載率(%)']}%", is_red=True)
    add_run_kai(r_cells[6].paragraphs[0], f"{kwh_new:,.0f}", is_red=True)

row_sum2 = t2.add_row().cells
row_sum2[0].merge(row_sum2[5])
add_run_kai(row_sum2[0].paragraphs[0], "總耗電量(kWh/年)", is_bold=True)
add_run_kai(row_sum2[6].paragraphs[0], f"{total_new_kwh:,.0f}", is_red=True, is_bold=True)

# 最終節電效益
save_kwh = total_old_kwh - total_new_kwh
save_money = save_kwh * elec_price / 10000
doc.add_paragraph()
res_p = doc.add_paragraph(); res_p.paragraph_format.first_line_indent = Pt(24)
add_run_kai(res_p, "預估年節電量約 ")
add_run_kai(res_p, f"{save_kwh:,.0f} kWh", is_red=True)
add_run_kai(res_p, "，年節省電費約 ")
add_run_kai(res_p, f"{save_money:.1f} 萬元", is_red=True)

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
    st.download_button(
        label="💾 下載目前的 Word 報告",
        data=current_word_data,
        file_name="冰水主機汰換效益分析.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
