import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心格式工具 ---

def set_table_border(table):
    """手動強制繪製表格黑色框線，確保表格不會隱形"""
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

def fix_cell(cell, text, is_bold=False, size=10):
    p = cell.paragraphs[0]
    p.alignment = 1 # 置中
    run = p.add_run(str(text))
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
def safe_replace(doc, data_map):
    """您原本測試成功的替換工具：處理段落與看板表格"""
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        if key in p.text:
                            for run in p.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(val))
                                    run.font.color.rgb = RGBColor(0, 0, 0)
                                    run.font.name = '標楷體'
                                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=50.0)
    elec_input = st.number_input("平均電費 (元/度)", value=3.5, step=0.01)
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="1500RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "運轉時數(hr)": [4380, 2190, 2190],
        "平均負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = motor_hp * 0.746 
    details = []
    total_old, total_new = 0, 0
    for _, row in df.iterrows():
        h = float(row["運轉時數(hr)"])
        l = float(row["平均負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h # 立方律 + 損失
        details.append({
            "季節": row["季節"], "時數": h, "負載": f"{row['平均負載率(%)']}%",
            "舊": o_kwh, "新": n_kwh, "省": o_kwh - n_kwh
        })
        total_old += o_kwh
        total_new += n_kwh
    return {
        "old_total": total_old, "new_total": total_new,
        "save_kwh": total_old - total_new, "details": details
    }

# --- 4. 生成按鈕 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    res = run_calculation(current_op_df)
    save_money = res['save_kwh'] * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    
    try:
        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name, "{{CH_INFO}}": ch_info, "{{RT_INFO}}": rt_info,
            "{{MT}}": f"三台 {int(motor_hp)}hp", "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{res['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{res['save_kwh']:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}",
            "{{COUNT}}": "2", "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SUPPRESS_KW}}": "13"
        }
        
        # 1. 文字替換
        safe_replace(doc, data_map)

        # 2. 在文末生成完全符合截圖的表格
        doc.add_page_break()
        doc.add_paragraph("--- 自動生成的表格 (請剪下並貼至指定位置) ---")

        # 表一：現況耗電明細表
        doc.add_paragraph("【表一、現況耗電明細表】")
        t1 = doc.add_table(rows=1, cols=4)
        set_table_border(t1)
        hdr1 = ["季節", "運轉時數(hr)", "平均負載率(%)", "耗電量(kWh)"]
        for i, text in enumerate(hdr1):
            t1.cell(0, i).text = text
            fix_cell_format(t1.cell(0, i), is_bold=True)
        for d in res['details']:
            row = t1.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
            for c in row: fix_cell_format(c)
        tot1 = t1.add_row().cells
        tot1[0].text, tot1[3].text = "合計", f"{res['old_total']:,.0f}"
        for c in tot1: fix_cell_format(c, is_bold=True)

        # 表二：預期節能效益表
        doc.add_paragraph("\n【表二、預期節能效益表】")
        t2 = doc.add_table(rows=1, cols=5)
        set_table_border(t2)
        hdr2 = ["季節", "運轉時數(hr)", "平均負載率(%)", "改善後耗電(kWh)", "節電量(kWh)"]
        for i, text in enumerate(hdr2):
            t2.cell(0, i).text = text
            fix_cell_format(t2.cell(0, i), is_bold=True)
        for d in res['details']:
            row = t2.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"
            for c in row: fix_cell_format(c)
        tot2 = t2.add_row().cells
        tot2[0].text, tot2[4].text = "合計", f"{res['save_kwh']:,.0f}"
        for c in tot2: fix_cell_format(c, is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！表格已按截圖格式生成於文末。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"出錯了: {e}")
