import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體工具函數 (支援變色) ---
def add_run_kai(paragraph, text, size=12, is_bold=False, is_red=False):
    run = paragraph.add_run(text)
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    if is_red:
        run.font.color.rgb = RGBColor(255, 0, 0) # 設定紅色
    else:
        run.font.color.rgb = RGBColor(0, 0, 0)   # 設定黑色
    return run

# --- 2. Streamlit 介面輸入 ---
st.title("❄️ P4. 冰水主機汰換效益分析")

col1, col2 = st.columns(2)
with col1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    setup_year = st.number_input("原主機設置年份 (民國)", value=104)
with col2:
# 1. 先從 Session State 拿自動計算的值，拿不到就用 4.48 (預設值)
val_from_app = st.session_state.get('auto_avg_price', 4.48)

# 2. 設定輸入框，讓它的預設值 (value) 等於自動抓到的值
elec_price = st.number_input(
    "平均電費 (元/度)", 
    min_value=0.0,
    value=float(val_from_app), # 自動帶入表五之二的結果
    step=0.01,
    key="p4_elec_price"
)

st.subheader("🧊 現有主機配置")
# 讓使用者可以填寫多台主機
df_init = pd.DataFrame([
    {"編號": "CH-1", "台數": 1, "容量(RT)": 350, "型式": "螺旋式"},
    {"編號": "CH-2", "台數": 1, "容量(RT)": 350, "型式": "離心式"}
])
chiller_config = st.data_editor(df_init, num_rows="dynamic", use_container_width=True)

# 季節參數設定
st.subheader("📅 運轉參數設定")
op_data = {
    "季節": ["春秋", "夏季", "冬季"],
    "時數(hr/y)": [2190, 1095, 1095],
    "負載率(%)": [60, 70, 50],
    "現況kW/RT": [0.864, 0.90, 0.846],
    "改善後kW/RT": [0.48, 0.50, 0.47]
}
df_op = st.data_editor(pd.DataFrame(op_data), use_container_width=True)

# --- 3. 執行產出邏輯 ---
if st.button("🚀 生成 P4 效益報告並同步"):
    doc = Document()
    
    # 計算主機描述文字
    desc_list = []
    total_old_kwh = 0
    total_new_kwh = 0
    main_rt = 0 # 假設以第一台為基準計算表格
    
    for _, r in chiller_config.iterrows():
        desc_list.append(f"{r['台數']}台{r['容量(RT)']}RT {r['型式']}")
        main_rt = r['容量(RT)'] # 取最後一個
    chiller_desc_text = "、".join(desc_list)

    # --- 一、 現況說明 ---
    p1_title = doc.add_paragraph()
    add_run_kai(p1_title, "一、現況說明", size=14, is_bold=True)
    
    p1_text = doc.add_paragraph()
    p1_text.paragraph_format.first_line_indent = Pt(24)
    add_run_kai(p1_text, f"1. {unit_name}空調系統有")
    add_run_kai(p1_text, chiller_desc_text, is_red=True) # 變數變紅
    add_run_kai(p1_text, "冰水主機(設置年份")
    add_run_kai(p1_text, str(setup_year), is_red=True)   # 變數變紅
    add_run_kai(p1_text, "年)，推估年度耗電量如下表：")

    # --- 建立現況耗電表格 ---
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    headers = ["季節", "製冷量(RT)", "台數", "運轉耗電率(kW/RT)", "時數(時/年)", "負載率"] # 先簡化
    # 此處省略表格繪製細節，先專注在文字...

    # --- 二、 預期效益 (黑紅計算結果) ---
    # 假設計算總量
    for _, row in df_op.iterrows():
        total_old_kwh += main_rt * 1 * row["現況kW/RT"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
        total_new_kwh += main_rt * 1 * row["改善後kW/RT"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
    
    save_kwh = total_old_kwh - total_new_kwh
    save_money = save_kwh * elec_price / 10000

    p_benefit = doc.add_paragraph()
    add_run_kai(p_benefit, "推估耗電量：")
    add_run_kai(p_benefit, f"{total_old_kwh:,.0f} kWh/年", is_red=True)
    add_run_kai(p_benefit, "。節省電費約 ")
    add_run_kai(p_benefit, f"{save_money:.1f} 萬元/年", is_red=True)

    # 存檔與預覽
    buf = io.BytesIO()
    doc.save(buf)
    st.session_state['p4_report'] = buf.getvalue()
    st.success("✅ P4 報告生成成功！文字已根據變數自動標註紅色。")
