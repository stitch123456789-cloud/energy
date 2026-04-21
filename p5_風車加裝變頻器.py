import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 格式修正工具 ---
def fix_run_style(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def fix_cell_style(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            fix_run_style(run, size, is_bold)

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻效益分析")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-5")
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
current_op_df = st.data_editor(st.session_state.p5_op_data, num_rows="dynamic", use_container_width=True)

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

# --- 4. 生成按鈕與核心邏輯 ---
st.markdown("---")
# --- 修正後的動態表格處理邏輯 (只改這裡) ---

if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    res = run_calculation(current_op_df)
    try:
        doc = Document("template_p5.docx")
        
        # 1. 執行您原本成功的文字替換 (safe_replace)
        safe_replace(doc, data_map)

        # 2. 專門處理表格插入 (這段代碼必須獨立於 safe_replace 之外)
        for p in doc.paragraphs:
            # 處理現況表格
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" # 標籤清空
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                # 填入表頭
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_style(table.cell(0, i), is_bold=True) # 確保格式
                # 填入細節
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']
                    row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = "100%"
                    row[3].text = f"{d['舊']:,.0f}"
                    for c in row: fix_cell_style(c)
                # 填入合計
                tot = table.add_row().cells
                tot[0].text = "合計"
                tot[3].text = f"{res['old_total']:,.0f}"
                for c in tot: fix_cell_style(c, is_bold=True)

            # 處理效益表格
            if "[[NEW_TABLE]]" in p.text:
                p.text = "" # 標籤清空
                table = doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_style(table.cell(0, i), is_bold=True)
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']
                    row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = d['負載']
                    row[3].text = f"{d['新']:,.0f}"
                    row[4].text = f"{d['省']:,.0f}"
                    for c in row: fix_cell_style(c)
                tot = table.add_row().cells
                tot[0].text = "合計"
                tot[4].text = f"{res['save_kwh']:,.0f}"
                for c in tot: fix_cell_style(c, is_bold=True)

        # 3. 儲存下載 (保持不變)
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")

    except Exception as e:
        st.error(f"錯誤: {e}")
        doc.save(buf)
        st.success("✅ 報告已成功生成！表格與文字已對齊。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤: {e}")
