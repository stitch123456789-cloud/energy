import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體工具函數 (支援黑紅變色) ---
def add_run_kai(paragraph, text, size=12, is_bold=False, is_red=False):
    run = paragraph.add_run(text)
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    if is_red:
        run.font.color.rgb = RGBColor(255, 0, 0) # 紅色
    else:
        run.font.color.rgb = RGBColor(0, 0, 0)   # 黑色
    return run

# --- 2. Streamlit 介面輸入區 ---
st.title("❄️ P4. 冰水主機汰換效益分析")

col1, col2 = st.columns(2)
with col1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    setup_year = st.number_input("原主機設置年份 (民國)", value=104)

with col2:
    # 自動帶入全域計算的電費，若無則預設 4.48
    val_from_app = st.session_state.get('auto_avg_price', 4.48)
    elec_price = st.number_input(
        "平均電費 (元/度)", 
        min_value=0.0,
        value=float(val_from_app),
        step=0.01,
        key="p4_elec_price"
    )

st.subheader("🧊 現有主機配置")
df_init = pd.DataFrame([
    {"編號": "CH-1", "台數": 1, "容量(RT)": 350, "型式": "螺旋式"},
    {"編號": "CH-2", "台數": 1, "容量(RT)": 350, "型式": "離心式"}
])
chiller_config = st.data_editor(df_init, num_rows="dynamic", use_container_width=True)

st.subheader("📅 運轉參數設定")
op_data = {
    "季節": ["春秋", "夏季", "冬季"],
    "時數(hr/y)": [2190, 1095, 1095],
    "負載率(%)": [60, 70, 50],
    "現況kW/RT": [0.864, 0.90, 0.846],
    "改善後kW/RT": [0.48, 0.50, 0.47]
}
df_op = st.data_editor(pd.DataFrame(op_data), use_container_width=True)

# ---------------------------------------------------------
# --- 3. 核心計算與 Word 即時生成邏輯 (已移出按鈕) ---
# ---------------------------------------------------------
doc = Document()

# A. 數據預處理
desc_list = []
main_rt = 0
for _, r in chiller_config.iterrows():
    desc_list.append(f"{r['台數']}台{r['容量(RT)']}RT {r['型式']}")
    main_rt = r['容量(RT)']
chiller_desc_text = "、".join(desc_list)

# B. Word 內容：一、現況說明
add_run_kai(doc.add_paragraph(), "一、現況說明", size=14, is_bold=True)
p1 = doc.add_paragraph()
p1.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p1, f"1. {unit_name}空調系統有")
add_run_kai(p1, chiller_desc_text, is_red=True)
add_run_kai(p1, "冰水主機(設置年份")
add_run_kai(p1, str(setup_year), is_red=True)
add_run_kai(p1, "年)，推估年度耗電量如下表：")

# C. Word 內容：現況表格 (t1)
t1 = doc.add_table(rows=1, cols=6); t1.style = 'Table Grid'
h_names = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "耗電\n(kWh/年)"]
for i, h in enumerate(h_names):
    add_run_kai(t1.cell(0,i).paragraphs[0], h, size=10, is_bold=True)

total_old_kwh = 0
for _, row in df_op.iterrows():
    kwh = main_rt * 1 * row["現況kW/RT"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
    total_old_kwh += kwh
    r_cells = t1.add_row().cells
    add_run_kai(r_cells[0].paragraphs[0], row["季節"])
    add_run_kai(r_cells[1].paragraphs[0], str(main_rt), is_red=True)
    add_run_kai(r_cells[2].paragraphs[0], "1", is_red=True)
    add_run_kai(r_cells[3].paragraphs[0], str(row["現況kW/RT"]), is_red=True)
    add_run_kai(r_cells[4].paragraphs[0], f"{row['時數(hr/y)']:,.0f}", is_red=True)
    add_run_kai(r_cells[5].paragraphs[0], f"{kwh:,.0f}", is_red=True)

# D. Word 內容：二、改善方案
doc.add_paragraph()
add_run_kai(doc.add_paragraph(), "二、改善方案", size=14, is_bold=True)
p2 = doc.add_paragraph()
p2.paragraph_format.first_line_indent = Pt(24)
add_run_kai(p2, "建議更換為高效率離心式冰水主機，預期改善後數據如下表：")

# E. Word 內容：改善後表格 (t2)
total_new_kwh = 0
t2 = doc.add_table(rows=1, cols=6); t2.style = 'Table Grid'
for i, h in enumerate(h_names):
    add_run_kai(t2.cell(0,i).paragraphs[0], h, size=10, is_bold=True)

for _, row in df_op.iterrows():
    kwh_new = main_rt * 1 * row["改善後kW/RT"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
    total_new_kwh += kwh_new
    r_cells = t2.add_row().cells
    add_run_kai(r_cells[0].paragraphs[0], row["季節"])
    add_run_kai(r_cells[1].paragraphs[0], str(main_rt), is_red=True)
    add_run_kai(r_cells[2].paragraphs[0], "1", is_red=True)
    add_run_kai(r_cells[3].paragraphs[0], str(row["改善後kW/RT"]), is_red=True)
    add_run_kai(r_cells[4].paragraphs[0], f"{row['時數(hr/y)']:,.0f}", is_red=True)
    add_run_kai(r_cells[5].paragraphs[0], f"{kwh_new:,.0f}", is_red=True)

# F. Word 內容：三、預期效益
save_kwh = total_old_kwh - total_new_kwh
save_money = save_kwh * elec_price / 10000

doc.add_paragraph()
add_run_kai(doc.add_paragraph(), "三、預期效益", size=14, is_bold=True)
res_p = doc.add_paragraph()
res_p.paragraph_format.first_line_indent = Pt(24)
add_run_kai(res_p, "預估年節電量約 ")
add_run_kai(res_p, f"{save_kwh:,.0f} kWh", is_red=True)
add_run_kai(res_p, "，年省電費約 ")
add_run_kai(res_p, f"{save_money:.1f} 萬元", is_red=True)

# ---------------------------------------------------------
# --- 4. 報告輸出中心 (雙按鈕直接顯示) ---
# ---------------------------------------------------------
st.markdown("---")
st.subheader("🚀 報告輸出中心")

# 將當前即時生成的 doc 轉為二進位數據
buf = io.BytesIO()
doc.save(buf)
current_word_data = buf.getvalue()

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    # 按鈕 1：同步到打包中心
    if st.button("🔄 確認數值並同步至打包中心", use_container_width=True):
        if 'report_warehouse' not in st.session_state:
            st.session_state['report_warehouse'] = {}
        # 將目前的 Word 數據存入暫存，供側邊欄 ZIP 下載
        st.session_state['report_warehouse']["4. 冰水主機效益分析"] = current_word_data
        st.success("✅ 數據已鎖定，左側打包清單已更新！")
        st.rerun()

with col_btn2:
    # 按鈕 2：直接下載單份 Word 報告
    st.download_button(
        label="💾 下載目前的 Word 報告",
        data=current_word_data,
        file_name="冰水主機汰換效益分析.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
