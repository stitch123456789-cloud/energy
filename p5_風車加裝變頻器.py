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

# --- 在 p5_風車加裝變頻器.py 的生成報告函數內 ---

def generate_step_by_step(res):
    doc = Document("template_p5.docx")
    
    # 建立精準對應字典
    # 這裡的 Key 必須跟你 Word 裡的紅字完全一樣
    data_map = {
        "{{貴單位}}": unit_name,
        "{{COUNT}}": str(2), # 如果是固定值也可以直接寫死
        "{{CH_INFO}}": "CH-1",
        "{{RT_INFO}}": "1200RT",
        "{{MOTOR_INFO}}": "三台 15hp",
        "{{OP_NOTE}}": "僅開啟一台"
    }

    # 執行「不傷格式」的替換
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                # 只有當標籤存在於這個段落時，才去拆解內部的 Runs
                for run in p.runs:
                    if key in run.text:
                        # 替換文字
                        run.text = run.text.replace(key, val)
                        # 強制修正顏色與字體 (這會覆蓋你的紅色，變成黑色)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = '標楷體'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

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
