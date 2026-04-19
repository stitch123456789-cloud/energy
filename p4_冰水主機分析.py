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
        run.font.color.rgb = RGBColor(255, 0, 0) # 紅色
    else:
        run.font.color.rgb = RGBColor(0, 0, 0)   # 黑色
    return run

# --- 2. Streamlit 介面輸入 ---
st.title("❄️ P4. 冰水主機汰換效益分析")

col1, col2 = st.columns(2)
with col1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    setup_year = st.number_input("原主機設置年份 (民國)", value=104)

with col2:
    # 修正縮進：確保以下內容在 with col2 內
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
    {"編號": "CH-2", "台數": 1, "容量(RT)": 350, "型式": "螺旋式"}
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

# --- 3. 執行產出邏輯 ---
if st.button("🚀 生成 P4 效益報告並同步"):
    doc = Document()
    
    # 預運算數據
    desc_list = []
    main_rt = 0
    total_qty = 0
    for _, r in chiller_config.iterrows():
        desc_list.append(f"{r['台數']}台{r['容量(RT)']}RT {r['型式']}")
        main_rt = r['容量(RT)']
        total_qty += r['台數']
    chiller_desc_text = "、".join(desc_list)

    # 一、 現況說明
    add_run_kai(doc.add_paragraph(), "一、現況說明", size=14, is_bold=True)
    p1 = doc.add_paragraph()
    p1.paragraph_format.first_line_indent = Pt(24)
    add_run_kai(p1, f"1. {unit_name}空調系統有")
    add_run_kai(p1, chiller_desc_text, is_red=True)
    add_run_kai(p1, "冰水主機(設置年份")
    add_run_kai(p1, str(setup_year), is_red=True)
    add_run_kai(p1, "年)，推估年度耗電量如下表：")

    # 現況表格
    t1 = doc.add_table(rows=1, cols=6); t1.style = 'Table Grid'
    h1 = ["季節", "製冷量\n(RT)", "台數", "運轉耗電率\n(kW/RT)", "時數\n(時/年)", "耗電\n(kWh/年)"]
    for i, h in enumerate(h1):
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
        add_run_kai(r_cells[4].paragraphs[0], str(row["時數(hr/y)"]), is_red=True)
        add_run_kai(r_cells[5].paragraphs[0], f"{kwh:,.0f}", is_red=True)

    # 二、 改善方案
    doc.add_paragraph()
    add_run_kai(doc.add_paragraph(), "二、改善方案", size=14, is_bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.first_line_indent = Pt(24)
    add_run_kai(p2, "建議更換為高效率離心式冰水主機，預期改善後數據如下表：")

    # 三、 預期效益
    doc.add_paragraph()
    add_run_kai(doc.add_paragraph(), "三、預期效益", size=14, is_bold=True)
    
    total_new_kwh = 0
    t2 = doc.add_table(rows=1, cols=6); t2.style = 'Table Grid'
    for i, h in enumerate(h1): # 使用相同表頭
        add_run_kai(t2.cell(0,i).paragraphs[0], h, size=10, is_bold=True)
        
    for _, row in df_op.iterrows():
        kwh_new = main_rt * 1 * row["改善後kW/RT"] * row["時數(hr/y)"] * (row["負載率(%)"]/100)
        total_new_kwh += kwh_new
        r_cells = t2.add_row().cells
        add_run_kai(r_cells[0].paragraphs[0], row["季節"])
        add_run_kai(r_cells[1].paragraphs[0], str(main_rt), is_red=True)
        add_run_kai(r_cells[2].paragraphs[0], "1", is_red=True)
        add_run_kai(r_cells[3].paragraphs[0], str(row["改善後kW/RT"]), is_red=True)
        add_run_kai(r_cells[4].paragraphs[0], str(row["時數(hr/y)"]), is_red=True)
        add_run_kai(r_cells[5].paragraphs[0], f"{kwh_new:,.0f}", is_red=True)

    # 結算文字
    save_kwh = total_old_kwh - total_new_kwh
    save_money = save_kwh * elec_price / 10000
    
    res_p = doc.add_paragraph()
    add_run_kai(res_p, "預估年節電量約 ")
    add_run_kai(res_p, f"{save_kwh:,.0f} kWh", is_red=True)
    add_run_kai(res_p, "，年省電費約 ")
    add_run_kai(res_p, f"{save_money:.1f} 萬元", is_red=True)

    # 同步到打包中心
   # ... (前面的 Word 生成邏輯保持不變) ...

    # 同步與下載按鈕區塊
    st.markdown("---")
    col_btn1, col_btn2 = st.columns(2)

    buf = io.BytesIO()
    doc.save(buf)
    report_data = buf.getvalue()

    with col_btn1:
        # 第一步：確認數值並存入打包中心
        if st.button("🔄 確認數值並同步至打包中心", use_container_width=True):
            if 'report_warehouse' not in st.session_state:
                st.session_state['report_warehouse'] = {}
            st.session_state['report_warehouse']["4. 冰水主機效益分析"] = report_data
            st.success("✅ 數據已鎖定，左側打包下載已更新！")
            st.rerun()

    with col_btn2:
        # 第二步：直接下載 (只有在 session 裡有資料時才顯示)
        if 'report_warehouse' in st.session_state and "4. 冰水主機效益分析" in st.session_state['report_warehouse']:
            st.download_button(
                label="📥 下載目前的 Word 報告",
                data=st.session_state['report_warehouse']["4. 冰水主機效益分析"],
                file_name="冰水主機汰換效益分析.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.button("📥 請先點擊左側確認同步", disabled=True, use_container_width=True)
