import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體與格式修正工具 ---
def fix_run_format(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0) # 強制轉回黑色

# --- 2. 介面與參數設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號/容量", value="CH-1 / 1200RT")
with c2:
    motor_spec = st.text_input("風車規格", value="三台 15hp")
    elec_price = st.session_state.get('auto_avg_price', 4.63)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_price), step=0.01)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("開機說明", value="僅開啟一台")

# 運轉時數與負載設定
st.subheader("📊 運轉時數設定 (三台風車合計)")
op_df = pd.DataFrame({
    "季節": ["春秋季", "夏季", "冬季"],
    "時數(hr)": [4380, 2190, 2190],
    "負載率(%)": [70, 85, 60]
})
edit_op = st.data_editor(op_df, use_container_width=True)

# --- 3. 核心計算邏輯 ---
def do_calc():
    # 假設 15hp * 0.746 = 11.2kW (單台)
    # 改善前：全速運轉；改善後：立方定律計算
    total_old_kwh = 0
    total_new_kwh = 0
    
    for _, row in edit_op.iterrows():
        hr = row["時數(hr)"]
        load = row["負載率(%)"] / 100
        # 改善前 (定頻)
        old_kwh = 11.2 * hr 
        # 改善後 (變頻 + 6%損失)
        new_kwh = 11.2 * (load**3) * 1.06 * hr
        
        total_old_kwh += old_kwh
        total_new_kwh += new_kwh
        
    save_kwh = total_old_kwh - total_new_kwh
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    save_rate = (save_kwh / total_old_kwh * 100) if total_old_kwh > 0 else 0
    
    return {
        "old_kwh": total_old_kwh,
        "save_kwh": save_kwh,
        "save_money": save_money,
        "save_rate": save_rate,
        "payback": payback
    }

# --- 4. 報告生成函數 ---
def generate_p5_report(res):
    doc = Document("template_p5.docx")
    
    # A. 替換紅色標籤 {{}}
    data_map = {
        "{{貴單位}}": unit_name,
        "{{CH-1}}": ch_info.split("/")[0].strip(),
        "{{1200RT}}": ch_info.split("/")[1].strip() if "/" in ch_info else ch_info,
        "{{三台 15hp}}": motor_spec,
        "{{僅開啟一台}}": setup_note,
        "{{110,277}}": f"{res['old_kwh']:,.0f}",
        "{{42,054}}": f"{res['save_kwh']:,.0f}",
        "{{19.5}}": f"{res['save_money']:.1f}",
        "{{SAVE_RATE}}": f"{res['save_rate']:.1f}",
        "{{INVEST}}": f"{invest_amt:.1f}",
        "{{PAYBACK}}": f"{res['payback']:.1f}"
    }

    # 執行段落替換
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)
                        fix_run_format(run)

    # B. 插入動態表格 [[OLD_TABLE_PLACEHOLDER]]
    for p in doc.paragraphs:
        if "[[OLD_TABLE_PLACEHOLDER]]" in p.text:
            p.text = "" # 清空標籤
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'
            # 填寫標頭
            hdr = ["季節", "馬力", "台數", "時數", "負載", "耗電(kWh)"]
            for i, h in enumerate(hdr):
                cell = table.cell(0, i)
                cell.text = h
                fix_run_format(cell.paragraphs[0].runs[0], is_bold=True)
            
            # 填寫數據邏輯 (省略細節，以此類推)...
            
    return doc

# --- 5. 輸出中心 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告"):
    res = do_calc()
    doc = generate_p5_report(res)
    
    buf = io.BytesIO()
    doc.save(buf)
    st.session_state['report_warehouse']["5. 風車變頻器分析"] = buf.getvalue()
    st.success("✅ 報告已生成並鎖定至打包中心！")
    st.download_button("💾 下載此份 Word", buf.getvalue(), "風車變頻器效益分析.docx")
