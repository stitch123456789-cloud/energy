import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 表格格式工具 (確保生成出來的表格是標楷體) ---
def format_cell(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車報告生成器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0)
with c2:
    motor_count = st.number_input("風車台數", value=3, step=1)
    elec_input = st.number_input("平均電費 (元/度)", value=3.5)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, num_rows="dynamic", use_container_width=True)

# --- 3. 生成按鈕 ---
if st.button("🚀 生成報告 (表格生成於文件末尾)"):
    # 計算邏輯
    base_kw = motor_hp * 0.746
    details = []
    total_old = 0
    total_new = 0
    
    for _, row in current_op_df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h
        details.append({"季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%", "舊": o_kwh, "新": n_kwh, "省": o_kwh - n_kwh})
        total_old += o_kwh
        total_new += n_kwh

    try:
        doc = Document("template_p5.docx")
        
        # --- A. 文字替換 (簡單替換，不傷排版) ---
        data_map = {
            "{{UN}}": unit_name,
            "{{MT}}": f"{int(motor_count)}台 {int(motor_hp)}hp",
            "{{OLD_KWH}}": f"{total_old:,.0f}",
            "{{SAVE_KWH}}": f"{(total_old - total_new):,.0f}",
            "{{SAVE_MONEY}}": f"{((total_old - total_new) * elec_input / 10000):.2f}",
            "{{PAYBACK}}": f"{(invest_amt / ((total_old - total_new) * elec_input / 10000)):.1f}"
        }
        
        for p in doc.paragraphs:
            for k, v in data_map.items():
                if k in p.text:
                    p.text = p.text.replace(k, str(v))

        # --- B. 在文件最後增加提示文字 ---
        doc.add_page_break()
        doc.add_paragraph("以下為生成的表格，請剪下並貼上到指定位置：", style='Normal')

        # --- C. 生成【現況耗電表格】 ---
        doc.add_paragraph("【現況耗電明細表】")
        table_old = doc.add_table(rows=1, cols=4)
        table_old.style = 'Table Grid'
        hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
        for i, text in enumerate(hdr):
            table_old.cell(0, i).text = text
            format_cell(table_old.cell(0, i), is_bold=True)
            
        for d in details:
            row = table_old.add_row().cells
            row[0].text = d['季節']
            row[1].text = f"{d['時數']:,.0f}"
            row[2].text = "100%"
            row[3].text = f"{d['舊']:,.0f}"
            for c in row: format_cell(c)
            
        # 合計
        tot = table_old.add_row().cells
        tot[0].text = "合計"; tot[3].text = f"{total_old:,.0f}"
        for c in tot: format_cell(c, is_bold=True)

        doc.add_paragraph("\n") # 隔開

        # --- D. 生成【預期效益表格】 ---
        doc.add_paragraph("【預期節能效益表】")
        table_new = doc.add_table(rows=1, cols=5)
        table_new.style = 'Table Grid'
        hdr_new = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
        for i, text in enumerate(hdr_new):
            table_new.cell(0, i).text = text
            format_cell(table_new.cell(0, i), is_bold=True)
            
        for d in details:
            row = table_new.add_row().cells
            row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"
            row[2].text = d['負載']; row[3].text = f"{d['新']:,.0f}"; row[4].text = f"{d['省']:,.0f}"
            for c in row: format_cell(c)
            
        # 合計
        tot_new = table_new.add_row().cells
        tot_new[0].text = "合計"; tot_new[4].text = f"{(total_old - total_new):,.0f}"
        for c in tot_new: format_cell(c, is_bold=True)

        # 輸出檔案
        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！表格已附加在文件最後一頁。")
        st.download_button("📥 下載報告", buf.getvalue(), "風車效益報告_含表格.docx")

    except Exception as e:
        st.error(f"發生錯誤: {e}")
