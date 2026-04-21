import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 格式修正工具 ---
def fix_run_format(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面與參數 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機資訊", value="CH-1 / 1200RT")
with c2:
    motor_spec = st.text_input("風車規格", value="三台 15hp")
    elec_val = st.session_state.get('auto_avg_price', 4.63)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_val), step=0.01)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("運轉說明", value="僅開啟一台")

st.subheader("📊 運轉時數與負載設定")
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })

# 重要：確保變數名稱一致
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 計算邏輯 ---
def run_calculation(df):
    base_kw = 11.2 # 15HP 基準
    total_old = 0
    total_new = 0
    details = []
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        # 立方定律計算
        n_kwh = base_kw * (l**3) * 1.06 * h
        total_old += o_kwh
        total_new += n_kwh
        details.append({"季節": row["季節"], "時數": h, "負載": f"{row['負載率(%)']}%", "舊耗電": o_kwh})
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    return {"old_kwh": total_old, "save_kwh": save_kwh, "save_money": save_money, "payback": payback, "details": details}

# --- 4. Word 生成 (修正 NameError 與表格消失問題) ---
def build_report(res_dict, op_table_df):
    try:
        doc = Document("template_p5.docx")
    except:
        st.error("找不到 template_p5.docx")
        return None

    # 文字替換
    d_map = {
        "{{貴單位}}": unit_name,
        "{{110, 277}}": f"{res_dict['old_kwh']:,.0f}",
        "{{42, 054}}": f"{res_dict['save_kwh']:,.0f}",
        "{{19.5}}": f"{res_dict['save_money']:.2f}",
        "{{58.5}}": f"{invest_amt:.1f}",
        "{{3}}": f"{res_dict['payback']:.1f}"
    }

    for p in doc.paragraphs:
        for k, v in d_map.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, str(v))
                        fix_run_format(run)

    # 表格插入邏輯 (修正關鍵點：直接傳入運轉資料)
    for p in doc.paragraphs:
        if "[[OLD_TABLE_PLACEHOLDER]]" in p.text:
            p.text = "" 
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'
            # 標頭
            hdr = ["季節", "馬力(hp)", "台數", "時數(hr)", "負載(%)", "耗電(kWh)"]
            for i, h in enumerate(hdr):
                table.cell(0, i).text = h
                fix_run_format(table.cell(0, i).paragraphs[0].runs[0], is_bold=True)
            
            # 內容填寫
            for d in res_dict['details']:
                row = table.add_row().cells
                row[0].text = d['季節']; row[1].text = "15"; row[2].text = "1"
                row[3].text = f"{d['時數']:,.0f}"; row[4].text = d['負載']; row[5].text = f"{d['舊耗電']:,.0f}"
                for c in row: fix_run_format(c.paragraphs[0].runs[0], size=10)
            
            # 合計
            tot = table.add_row().cells
            tot[0].text = "合計"; tot[5].text = f"{res_dict['old_kwh']:,.0f}"
            fix_run_format(tot[0].paragraphs[0].runs[0], is_bold=True)
            fix_run_format(tot[5].paragraphs[0].runs[0], is_bold=True)

    return doc

# --- 5. 執行按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    # 先計算，再傳入結果生成報告
    final_res = run_calculation(current_op_df)
    report_doc = build_report(final_res, current_op_df)
    
    if report_doc:
        buf = io.BytesIO()
        report_doc.save(buf)
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = buf.getvalue()
        st.success("✅ 報告生成成功！")
        st.download_button("📥 下載 Word", buf.getvalue(), "風車效益分析.docx")
