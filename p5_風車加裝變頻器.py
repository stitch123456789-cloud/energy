import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 格式與替換工具 ---
def safe_replace(doc, data_map):
    """安全替換文字，保留段落縮排"""
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

def fix_cell_format(cell, size=10, is_bold=False):
    """表格內容強制格式化"""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=15.0)
    elec_val = st.session_state.get('auto_avg_price', 4.45)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    rt_info = st.text_input("冷卻水塔容量", value="1200RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("運轉說明", value="僅開啟一台")

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
    total_old = 0
    total_new = 0
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h 
        s_kwh = o_kwh - n_kwh
        details.append({
            "季節": row["季節"], 
            "時數": h, 
            "負載": f"{row['負載率(%)']}%", 
            "舊": o_kwh, 
            "新": n_kwh, 
            "省": s_kwh
        })
        total_old += o_kwh
        total_new += n_kwh
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    save_rate = (save_kwh / total_old * 100) if total_old > 0 else 0
    
    return {
        "old_total": total_old, 
        "save_kwh": save_kwh, 
        "save_money": save_money, 
        "payback": payback,
        "save_rate": save_rate,
        "details": details
    }

# --- 4. 生成報告邏輯 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        doc = Document("template_p5.docx")
        
        # A. 文字替換
        data_map = {
            "{{UN}}": unit_name,
            "{{COUNT}}": "2",
            "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info,
            "{{MT}}": f"三台 {int(motor_hp)}hp",
            "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{results['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_RATE}}": f"{results['save_rate']:.2f}",
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}",
            "{{SUPPRESS_KW}}": "13"
        }
        safe_replace(doc, data_map)
        
        # B. 動態表格生成
        for p in doc.paragraphs:
            # 1. 現況表格 (OLD)
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" 
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                # 標頭
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                # 資料列
                for d in results['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']
                    row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = "100%"
                    row[3].text = f"{d['舊']:,.0f}"
                    for c in row: fix_cell_format(c)
                # 合計列
                tot = table.add_row().cells
                tot[0].text = "合計"; tot[3].text = f"{results['old_total']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

            # 2. 效益表格 (NEW)
            if "[[NEW_TABLE]]" in p.text:
                p.text = ""
                table = doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                for d in results['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = d['負載']; row[3].text = f"{d['新']:,.0f}"
                    row[4].text = f"{d['省']:,.0f}"
                    for c in row: fix_cell_format(c)
                tot = table.add_row().cells
                tot[0].text = "合計"; tot[4].text = f"{results['save_kwh']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

        # C. 儲存與下載
        buf = io.BytesIO()
        doc.save(buf)
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = buf.getvalue()
        st.success("✅ 報告生成成功！")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"發生錯誤: {e}")
