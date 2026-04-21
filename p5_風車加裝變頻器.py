import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 格式修正工具 ---
def fix_run_format(run, size=10, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def fix_cell_format(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            fix_run_format(run, size, is_bold)

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
    return {"old_total": total_old, "save_kwh": save_kwh, "save_money": save_money, 
            "payback": payback, "save_rate": save_rate, "details": details}

# --- 4. 核心生成邏輯 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    res = run_calculation(current_op_df)
    try:
        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name, "{{CH_INFO}}": ch_info, "{{RT_INFO}}": rt_info,
            "{{MT}}": f"三台 {int(motor_hp)}hp", "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{res['old_total']:,.0f}", "{{SAVE_KWH}}": f"{res['save_kwh']:,.0f}",
            "{{SAVE_RATE}}": f"{res['save_rate']:.2f}", "{{SAVE_MONEY}}": f"{res['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}", "{{PAYBACK}}": f"{res['payback']:.1f}",
            "{{SUPPRESS_KW}}": "13"
        }

        # 一次性掃描所有段落進行 替換 或 插入表格
        for p in doc.paragraphs:
            # 處理文字標籤替換
            for key, val in data_map.items():
                if key in p.text:
                    for run in p.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, str(val))
                            fix_run_format(run, size=12)

            # 處理現況表格 [[OLD_TABLE]]
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" # 清除標籤
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = "100%"; row[3].text = f"{d['舊']:,.0f}"
                    for c in row: fix_cell_format(c)
                tot = table.add_row().cells
                tot[0].text = "合計"; tot[3].text = f"{res['old_total']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

            # 處理效益表格 [[NEW_TABLE]]
            if "[[NEW_TABLE]]" in p.text:
                p.text = "" # 清除標籤
                table = doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = d['負載']; row[3].text = f"{d['新']:,.0f}"
                    row[4].text = f"{d['省']:,.0f}"
                    for c in row: fix_cell_format(c)
                tot = table.add_row().cells
                tot[0].text = "合計"; tot[4].text = f"{res['save_kwh']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

        # 處理現有表格內的文字替換 (例如看板表格)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, val in data_map.items():
                        if key in cell.text:
                            for p_in_cell in cell.paragraphs:
                                for run in p_in_cell.runs:
                                    if key in run.text:
                                        run.text = run.text.replace(key, str(val))
                                        fix_run_format(run, size=10)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 報告已生成！請點擊下方按鈕下載。")
        st.download_button("📥 下載修正版 Word 報告", buf.getvalue(), "風車效益分析_修正版.docx")
        
    except Exception as e:
        st.error(f"發生錯誤: {e}")
