import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 核心替換工具 (僅更換文字，保留段落格式) ---
def safe_replace(doc, data_map):
    """
    這是一個精確的替換函數。它不會動到 paragraph 層級(縮排、行距)，
    只會動到 run 層級(文字內容)。
    """
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        # 執行替換
                        run.text = run.text.replace(key, str(val))
                        # 強制改回黑色標楷體
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    # 同時處理表格內部的格子文字替換
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
with c2:
    motor_hp = st.number_input("單台風車馬力 (HP)", value=15.0)
    elec_val = st.session_state.get('auto_avg_price', 4.63)
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
    base_kw = motor_hp * 0.746 
    total_old = 0
    total_new = 0
    for _, row in df.iterrows():
        h = float(row["時數(hr)"])
        l = float(row["負載率(%)"]) / 100
        o_kwh = base_kw * h
        n_kwh = base_kw * (l**3) * 1.06 * h 
        total_old += o_kwh
        total_new += n_kwh
    
    save_kwh = total_old - total_new
    save_money = save_kwh * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    return {
        "old_total": total_old, 
        "save_kwh": save_kwh, 
        "save_money": save_money, 
        "payback": payback
    }

# --- 4. 輸出按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        # 讀取模板
        doc = Document("template_p5.docx")
        
        # 定義替換地圖 (請確保 Word 裡的標籤名稱與這裡完全一致)
        data_map = {
            "{{貴單位}}": unit_name,
            "{{COUNT}}": "2",
            "{{CH_INFO}}": "CH-1",
            "{{RT_INFO}}": "1200RT",
            "{{MOTOR_INFO}}": f"三台 {int(motor_hp)}hp",
            "{{OP_NOTE}}": "僅開啟一台",
            "{{OLD_KWH}}": f"{results['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}"
        }
        
        # 執行安全替換
        safe_replace(doc, data_map)
        
        # 儲存結果
        buf = io.BytesIO()
        doc.save(buf)
        report_data = buf.getvalue()
        
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = report_data
        st.success("✅ 報告生成成功！僅替換文字，保留原始排版。")
        st.download_button("📥 下載 Word 報告", report_data, "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"錯誤: {e}")
