import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 格式修正工具 (確保字體變回黑色標楷體) ---
def fix_run_format(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0) # 強制變回黑色

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    elec_val = st.session_state.get('auto_avg_price', 4.63)
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=15.0)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)

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
    base_kw = motor_hp * 0.746 # 15HP 約 11.2kW
    total_old = 0
    total_new = 0
    details = []
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h # 立方定律 + 6%變頻損失
        total_old += o_kwh
        total_new += n_kwh
        details.append({
            "季節": row["季節"], 
            "時數": h, 
            "負載": f"{row['負載率(%)']}%", 
            "舊耗電": o_kwh, 
            "新耗電": n_kwh,
            "節電": o_kwh - n_kwh
        })
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    save_rate = (save_kwh / total_old * 100) if total_old > 0 else 0
    payback = invest_amt / save_money if save_money > 0 else 0
    return {
        "old_total": total_old, 
        "save_kwh": save_kwh, 
        "save_money": save_money, 
        "save_rate": save_rate,
        "payback": payback, 
        "details": details
    }

# --- 4. Word 生成 (精準對齊你的新標籤) ---
def build_report(res):
    try:
        doc = Document("template_p5.docx")
    except:
        st.error("找不到 template_p5.docx")
        return None

    # 此處字典的 Key 必須跟 Word 裡的紅色文字完全一樣
    d_map = {
        "{{貴單位}}": unit_name,
        "{{OLD_KWH}}": f"{res['old_total']:,.0f}",
        "{{SAVE_KWH}}": f"{res['save_kwh']:,.0f}",
        "{{SAVE_RATE}}": f"{res['save_rate']:.1f}",
        "{{SAVE_MONEY}}": f"{res['save_money']:.2f}",
        "{{INVEST}}": f"{invest_amt:.1f}",
        "{{PAYBACK}}": f"{res['payback']:.1f}"
    }

    # 1. 處理段落文字替換
    for p in doc.paragraphs:
        # 替換文字標籤
        for k, v in d_map.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, str(v))
                        fix_run_format(run)
        
        # 插入現況表格 [[OLD_TABLE]]
        if "[[OLD_TABLE]]" in p.text:
            p.text = "" # 刪除標籤文字
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
            for i, h in enumerate(hdr):
                table.cell(0, i).text = h
                fix_run_format(table.cell(0, i).paragraphs[0].runs[0], is_bold=True)
            for d in res['details']:
                row = table.add_row().cells
                row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"; row[2].text = "100%"; row[3].text = f"{d['舊耗電']:,.0f}"
                for cell in row: fix_run_format(cell.paragraphs[0].runs[0], size=10)
            tot = table.add_row().cells; tot[0].text = "合計"; tot[3].text = f"{res['old_total']:,.0f}"
            fix_run_format(tot[0].paragraphs[0].runs[0], is_bold=True); fix_run_format(tot[3].paragraphs[0].runs[0], is_bold=True)

        # 插入效益表格 [[NEW_TABLE]]
        if "[[NEW_TABLE]]" in p.text:
            p.text = "" # 刪除標籤文字
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量(kWh)"]
            for i, h in enumerate(hdr):
                table.cell(0, i).text = h
                fix_run_format(table.cell(0, i).paragraphs[0].runs[0], is_bold=True)
            for d in res['details']:
                row = table.add_row().cells
                row[0].text = d['季節']; row[1].text = f"{d['時數']:,.0f}"; row[2].text = d['負載']
                row[3].text = f"{d['新耗電']:,.0f}"; row[4].text = f"{d['節電']:,.0f}"
                for cell in row: fix_run_format(cell.paragraphs[0].runs[0], size=10)
            tot = table.add_row().cells; tot[0].text = "合計"; tot[4].text = f"{res['save_kwh']:,.0f}"
            fix_run_format(tot[0].paragraphs[0].runs[0], is_bold=True); fix_run_format(tot[4].paragraphs[0].runs[0], is_bold=True)

    # 2. 處理表格內的文字替換 (例如回收年限格子裡的 {{}})
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for k, v in d_map.items():
                    if k in cell.text:
                        cell.text = cell.text.replace(k, str(v))
                        if cell.paragraphs: fix_run_format(cell.paragraphs[0].runs[0], size=10)

    return doc

# --- 5. 輸出按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    final_doc = build_report(results)
    if final_doc:
        buf = io.BytesIO(); final_doc.save(buf)
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = buf.getvalue()
        st.success("✅ 報告生成成功！標籤已替換且表格已分開插入。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")
