import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 極簡格式工具 ---
def apply_font_style(run, size=11, is_bold=False):
    """設定標楷體與字體大小"""
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_table_border(table):
    """手動為表格添加框線，防止 Word 樣式缺失報錯"""
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

def fix_cell_format(cell, size=10, is_bold=False):
    """確保表格每一格都是標楷體"""
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            apply_font_style(run, size, is_bold)

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析系統")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位", key="p5_u")
    ch_info = st.text_input("主機編號", value="CH-1", key="p5_ch")
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=50.0, key="p5_hp")
    elec_val = st.session_state.get('auto_avg_price', 3.5)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01, key="p5_e")
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="1500RT", key="p5_rt")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0, key="p5_inv")
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台", key="p5_note")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True, key="p5_editor")

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = motor_hp * 0.746 
    details = []
    total_old, total_new = 0, 0
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h # 立方定律 + 6% 損耗
        details.append({
            "季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%",
            "舊": o_kwh, "新": n_kwh, "省": o_kwh - n_kwh
        })
        total_old += o_kwh
        total_new += n_kwh
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    save_rate = (save_kwh / total_old * 100) if total_old > 0 else 0
    return {
        "old_total": total_old, "save_kwh": save_kwh, "save_money": save_money, 
        "payback": payback, "save_rate": save_rate, "details": details
    }

# --- 4. 生成按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告 (含動態表格)", use_container_width=True):
    res = run_calculation(current_op_df)
    try:
        doc = Document("template_p5.docx")
        
        # A. 文字替換
        data_map = {
            "{{UN}}": unit_name, "{{CH_INFO}}": ch_info, "{{RT_INFO}}": rt_info,
            "{{MT}}": f"三台 {int(motor_hp)}hp", "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{res['old_total']:,.0f}", "{{SAVE_KWH}}": f"{res['save_kwh']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台", "{{SAVE_RATE}}": f"{res['save_rate']:.2f}",
            "{{SAVE_MONEY}}": f"{res['save_money']:.2f}", "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{res['payback']:.1f}", "{{SUPPRESS_KW}}": "13", "{{COUNT}}": "2"
        }
        
        for p in doc.paragraphs:
            # 暴力替換文字
            for k, v in data_map.items():
                if k in p.text:
                    p.text = p.text.replace(k, str(v))
                    for run in p.runs: apply_font_style(run)

            # B. 偵測標籤並插入表格
            full_p_text = "".join(run.text for run in p.runs).strip()
            
            if "[[OLD_TABLE]]" in full_p_text:
                p.text = "" 
                table = doc.add_table(rows=1, cols=4)
                set_table_border(table)
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
                    for c in row: fix_cell_format(c)
                tot = table.add_row().cells
                tot[0].text, tot[3].text = "合計", f"{res['old_total']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

            if "[[NEW_TABLE]]" in full_p_text:
                p.text = ""
                table = doc.add_table(rows=1, cols=5)
                set_table_border(table)
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"
                    for c in row: fix_cell_format(c)
                tot = table.add_row().cells
                tot[0].text, tot[4].text = "合計", f"{res['save_kwh']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

        # 處理現有看板表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for k, v in data_map.items():
                        if k in cell.text:
                            cell.text = cell.text.replace(k, str(v))
                            fix_cell_format(cell, size=10)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！文字與動態表格已完成。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
