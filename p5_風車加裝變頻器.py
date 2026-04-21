import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 ---

def set_table_border(table):
    """手動為表格添加黑色框線，防止部分 Word 版本看不到格子"""
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

def fix_cell_font(cell, size=10, is_bold=False):
    """表格內容強制格式化為標楷體"""
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """您的核心替換工具：處理段落與現有表格，不傷格式"""
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
    elec_val = st.session_state.get('auto_avg_price', 3.5)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="1500RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

st.subheader("📊 運轉參數設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = motor_hp * 0.746 
    details = []
    total_old, total_new = 0, 0
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h 
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
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name, "{{COUNT}}": "2", "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info, "{{MT}}": f"三台 {int(motor_hp)}hp",
            "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{results['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_RATE}}": f"{results['save_rate']:.2f}",
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}",
            "{{SUPPRESS_KW}}": "13",
            "{{13}}": "13"
        }
        
        # 1. 執行文字替換
        safe_replace(doc, data_map)
        
        # 2. 清除文中原本的表格標籤文字
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text: p.text = ""
            if "[[NEW_TABLE]]" in p.text: p.text = ""

        # 3. 在文末生成新表格 (確保絕對能看到數據)
        doc.add_page_break()
        doc.add_paragraph("--- 以下為自動生成的表格 (請剪下貼上至指定位置) ---")

        # 生成現況表
        doc.add_paragraph("【表一、現況耗電明細表】")
        t1 = doc.add_table(rows=1, cols=4)
        set_table_border(t1)
        hdr1 = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
        for i, text in enumerate(hdr1):
            t1.cell(0, i).text = text
            fix_cell_font(t1.cell(0, i), is_bold=True)
        for d in results['details']:
            row = t1.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
            for c in row: fix_cell_font(c)
        tot1 = t1.add_row().cells
        tot1[0].text, tot1[3].text = "合計", f"{results['old_total']:,.0f}"
        for c in tot1: fix_cell_font(c, is_bold=True)

        # 生成效益表
        doc.add_paragraph("\n【表二、預期節能效益表】")
        t2 = doc.add_table(rows=1, cols=5)
        set_table_border(t2)
        hdr2 = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
        for i, text in enumerate(hdr2):
            t2.cell(0, i).text = text
            fix_cell_font(t2.cell(0, i), is_bold=True)
        for d in results['details']:
            row = t2.add_row().cells
            row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"
            for c in row: fix_cell_font(c)
        tot2 = t2.add_row().cells
        tot2[0].text, tot2[4].text = "合計", f"{results['save_kwh']:,.0f}"
        for c in tot2: fix_cell_font(c, is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告生成成功！請至最後一頁查看表格。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"執行出錯: {e}")
