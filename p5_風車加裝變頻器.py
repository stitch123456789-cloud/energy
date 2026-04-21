import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 極簡格式工具 (確保標楷體) ---
def apply_font_style(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_table_border(table):
    """手動強制繪製表格黑色框線"""
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

def fix_cell_format(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            apply_font_style(run, size, is_bold)

# --- 2. 介面與計算邏輯 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位", key="p5_unit")
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0, key="p5_hp")
with c2:
    motor_count = st.number_input("風車台數", value=3, step=1, key="p5_cnt")
    elec_val = st.number_input("平均電費 (元/度)", value=3.5, step=0.01, key="p5_elec")
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0, key="p5_inv")
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台", key="p5_note")

if "p5_data" not in st.session_state:
    st.session_state.p5_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_df = st.data_editor(st.session_state.p5_data, use_container_width=True, key="p5_edit")

# --- 3. 生成按鈕 ---
if st.button("🚀 生成報告 (表格生成於文件末尾)", use_container_width=True):
    try:
        # 計算結果
        base_kw = motor_hp * 0.746
        details = []
        total_old, total_new = 0, 0
        for _, row in current_df.iterrows():
            h, l = float(row["時數(hr)"]), float(row["負載率(%)"])/100
            o_kwh = base_kw * h
            n_kwh = base_kw * (l**3) * 1.06 * h
            details.append({"季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%", "舊": o_kwh, "新": n_kwh, "省": o_kwh-n_kwh})
            total_old += o_kwh
            total_new += n_kwh

        save_kwh = total_old - total_new
        save_money = save_kwh * elec_val / 10000
        payback = invest_amt / save_money if save_money > 0 else 0

        # 開啟 Word
        doc = Document("template_p5.docx")

        # A. 執行文字替換 (不理會表格標籤)
        data_map = {
            "{{UN}}": unit_name, "{{MT}}": f"{int(motor_count)}台 {int(motor_hp)}hp",
            "{{OLD_KWH}}": f"{total_old:,.0f}", "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}", "{{PAYBACK}}": f"{payback:.1f}"
        }
        for p in doc.paragraphs:
            for k, v in data_map.items():
                if k in p.text:
                    p.text = p.text.replace(k, str(v))
                    for run in p.runs: apply_font_style(run)

        # B. 在文件最後強制增加分頁，並生成表格
        doc.add_page_break()
        doc.add_paragraph("--- 自動生成的表格 (請剪下並貼至指定位置) ---")

        # 建立現況表
        doc.add_paragraph("【現況耗電明細表】")
        t1 = doc.add_table(rows=1, cols=4)
        set_table_border(t1)
        hdr1 = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
        for i, text in enumerate(hdr1):
            t1.cell(0, i).text = text
            fix_cell_format(t1.cell(0, i), is_bold=True)
        for d in details:
            row = t1.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
            for c in row: fix_cell_format(c)
        
        # 建立效益表
        doc.add_paragraph("\n【預期節能效益表】")
        t2 = doc.add_table(rows=1, cols=5)
        set_table_border(t2)
        hdr2 = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
        for i, text in enumerate(hdr2):
            t2.cell(0, i).text = text
            fix_cell_format(t2.cell(0, i), is_bold=True)
        for d in details:
            row = t2.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"
            for c in row: fix_cell_format(c)

        # C. 導出
        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告已生成！表格位於 Word 最後一頁。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤: {e}")
