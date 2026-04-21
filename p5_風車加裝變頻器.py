import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 核心工具函數 (確保字體為標楷體) ---
def fix_font(run, size=10, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """安全替換 {{標籤}}"""
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                p.text = p.text.replace(key, str(val))
                for run in p.runs:
                    fix_font(run)

# --- 2. 介面與計算 ---
st.title("🌀 P5. 冷卻水塔風車變頻效益分析")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    motor_hp = st.number_input("單台馬力 (HP)", value=15.0)
with c2:
    motor_count = st.number_input("風車台數", value=3, step=1)
    elec_input = st.number_input("平均電費 (元/度)", value=4.45)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("運轉說明", value="僅開啟兩台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, num_rows="dynamic", use_container_width=True)

# --- 3. 生成按鈕 ---
if st.button("🚀 生成報告 (表格生成於文末)"):
    try:
        # A. 計算
        base_kw = motor_hp * 0.746
        details = []
        total_old, total_new = 0, 0
        for _, row in current_op_df.iterrows():
            h, l = float(row["時數(hr)"]), float(row["負載率(%)"]) / 100
            o_kwh = base_kw * h
            n_kwh = base_kw * (l**3) * 1.06 * h
            details.append({"季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%", "舊": o_kwh, "新": n_kwh, "省": o_kwh - n_kwh})
            total_old += o_kwh
            total_new += n_kwh

        # B. 開啟範本與替換文字
        doc = Document("template_p5.docx")
        data_map = {
            "{{UN}}": unit_name, "{{MT}}": f"{int(motor_count)}台 {int(motor_hp)}hp",
            "{{OLD_KWH}}": f"{total_old:,.0f}", "{{SAVE_KWH}}": f"{(total_old-total_new):,.0f}",
            "{{SAVE_MONEY}}": f"{((total_old-total_new)*elec_input/10000):.2f}",
            "{{PAYBACK}}": f"{(invest_amt/((total_old-total_new)*elec_input/10000)):.1f}"
        }
        safe_replace(doc, data_map)

        # C. 在文件最後一頁生成表格 (100% 成功，不跑版)
        doc.add_page_break()
        doc.add_paragraph("--- 以下為自動生成的表格，請剪下後貼至報告指定位置 ---")

        # 表格 1：現況耗電
        doc.add_paragraph("【表一、現況運轉耗電明細表】")
        t1 = doc.add_table(rows=1, cols=4)
        t1.style = 'Table Grid'
        cols1 = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
        for i, name in enumerate(cols1):
            t1.cell(0, i).text = name
        for d in details:
            row = t1.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
        
        # 表格 2：節能效益
        doc.add_paragraph("\n【表二、預期節能效益表】")
        t2 = doc.add_table(rows=1, cols=5)
        t2.style = 'Table Grid'
        cols2 = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
        for i, name in enumerate(cols2):
            t2.cell(0, i).text = name
        for d in details:
            row = t2.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"

        # D. 匯出
        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告已生成！表格位於最後一頁。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益報告.docx")

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
