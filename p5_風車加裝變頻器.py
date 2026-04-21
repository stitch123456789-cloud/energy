import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 核心替換工具 (不傷格式) ---
def safe_replace(doc, data_map):
    """
    精確替換文字並強制校正字體，完全不觸動段落格式(縮排/行距)。
    """
    # 處理一般段落
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        run.font.color.rgb = RGBColor(0, 0, 0) # 轉黑色
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    # 處理表格內的所有格子
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
    elec_val = st.session_state.get('auto_avg_price', 4.45)
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
        n_kwh = base_kw * (l**3) * 1.06 * h # 立方定律 + 6%損耗
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

# --- 4. 生成按鈕 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        doc = Document("template_p5.docx")
        
        # 定義全部文字替換對應表
        full_data_map = {
            "{{貴單位}}": unit_name,
            "{{OLD_KWH_TEXT}}": f"{results['old_total']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{SUPPRESS_KW}}": "13", # 需量抑低可依需求改為計算值
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}"
        }
        
        # 執行文字替換 (保留所有縮排格式)
        safe_replace(doc, full_data_map)
        
        # --- 處理表格：僅在標籤處插入簡單表格 ---
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" # 移除標籤
                # 這裡暫時插入一個極簡提示，確認位置正確
                # (如需完整數據表格，建議直接在 Word 畫好並用 {{}} 填空)
            if "[[NEW_TABLE]]" in p.text:
                p.text = ""

        buf = io.BytesIO()
        doc.save(buf)
        report_data = buf.getvalue()
        
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = report_data
        st.success("✅ 報告生成成功！文字已填入並修正為標楷體。")
        st.download_button("📥 下載 Word 報告", report_data, "風車效益分析.docx")
        
    except Exception as e:
        st.error(f"執行出錯: {e}")
