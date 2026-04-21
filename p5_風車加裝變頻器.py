import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 極簡格式工具 ---
def apply_font(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_table_border(table):
    """手動為表格添加框線，防止樣式錯誤"""
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') # 1/8 pt
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析")

col1, col2, col3 = st.columns(3)
with col1:
    unit_name = st.text_input("單位名稱", value="貴單位", key="p5_unit_final")
with col2:
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0, key="p5_hp_final")
    elec_price = st.number_input("平均電費 (元/度)", value=3.5, key="p5_elec_final")
with col3:
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0, key="p5_invest_final")
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台", key="p5_note_final")

# 參數設定表格
if "p5_data_final" not in st.session_state:
    st.session_state.p5_data_final = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_df = st.data_editor(st.session_state.p5_data_final, use_container_width=True, key="p5_editor_final")

# --- 3. 生成邏輯 ---
if st.button("🚀 生成報告 (修正 apply_font 錯誤)", use_container_width=True):
    try:
        # 計算
        base_kw = motor_hp * 0.746
        total_old, total_new = 0, 0
        rows_data = []
        for _, r in current_df.iterrows():
            h, l = float(r["時數(hr)"]), float(r["負載率(%)"])/100
            o, n = base_kw * h, base_kw * (l**3) * 1.06 * h
            rows_data.append([r["季節"], f"{h:,.0f}", f"{r['負載率(%)']}%", f"{o:,.0f}", f"{n:,.0f}", f"{o-n:,.0f}"])
            total_old += o
            total_new += n
        
        save_kwh = total_old - total_new
        save_money = save_kwh * elec_price / 10000
        payback = invest_amt / save_money if save_money > 0 else 0

        # 開啟文件
        doc = Document("template_p5.docx")
        
        # 文字替換
        data_map = {
            "{{UN}}": unit_name,
            "{{OLD_KWH}}": f"{total_old:,.0f}",
            "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{PAYBACK}}": f"{payback:.1f}"
        }
        for p in doc.paragraphs:
            for k, v in data_map.items():
                if k in p.text:
                    p.text = p.text.replace(k, v)
                    for run in p.runs: apply_font(run)

        # 在文末生成表格
        doc.add_page_break()
        doc.add_paragraph("--- 自動生成效益表 (請剪下貼上) ---")
        
        table = doc.add_table(rows=1, cols=6)
        set_table_border(table) # 強制畫出框線
        
        # 標題
        hdr = ["季節", "時數", "負載", "現況耗電", "預期耗電", "節電量"]
        for i, name in enumerate(hdr):
            cell = table.cell(0, i)
            cell.text = name
            for p in cell.paragraphs:
                for run in p.runs: apply_font(run, is_bold=True)
        
        # 數據
        for row_vals in rows_data:
            cells = table.add_row().cells
            for i, val in enumerate(row_vals):
                cells[i].text = val
                for p in cells[i].paragraphs:
                    for run in p.runs: apply_font(run)
        
        # 下載
        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車分析報告.docx")

    except Exception as e:
        st.error(f"❌ 執行出錯: {str(e)}")
