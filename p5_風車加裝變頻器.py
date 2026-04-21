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
    run.font.color.rgb = RGBColor(0, 0, 0) # 強制變回黑色

# --- 2. 介面與參數設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機資訊", value="CH-1 / 1200RT")
with c2:
    motor_spec = st.text_input("風車規格", value="三台 15hp")
    elec_price_val = st.session_state.get('auto_avg_price', 4.63)
    elec_input = st.number_input("平均電費 (元/度)", value=float(elec_price_val), step=0.01)
with c3:
    invest_amt = st.number_input("投資金額 (萬元)", value=58.5)
    setup_note = st.text_input("運轉說明", value="僅開啟一台")

st.markdown("---")
st.subheader("📊 運轉時數與負載設定 (改善前後對照)")

# 初始化資料表
if "p5_op_data" not in st.session_state:
    st.session_state.p5_op_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })

edit_op = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

# --- 3. 核心計算邏輯 ---
def calculate_results(df):
    # 假設單台風車 15HP 耗功 11.2 kW
    base_kw = 11.2
    total_old_kwh = 0
    total_new_kwh = 0
    
    # 用於表格顯示的細節資料
    calc_details = []

    for _, row in df.iterrows():
        hr = float(row["時數(hr)"])
        load_pct = float(row["負載率(%)"]) / 100
        
        # 改善前：定頻全速運轉
        old_kwh = base_kw * hr
        # 改善後：立方定律 P2 = P1 * (RPM2/RPM1)^3 * (1 + 變頻器損失6%)
        new_kwh = base_kw * (load_pct**3) * 1.06 * hr
        
        total_old_kwh += old_kwh
        total_new_kwh += new_kwh
        
        calc_details.append({
            "季節": row["季節"],
            "時數": hr,
            "負載": f"{row['負載率(%)']}%",
            "改善前kwh": old_kwh,
            "改善後kwh": new_kwh
        })
        
    save_kwh = total_old_kwh - total_new_kwh
    save_money = save_kwh * elec_input / 10000
    save_rate = (save_kwh / total_old_kwh * 100) if total_old_kwh > 0 else 0
    payback = invest_amt / save_money if save_money > 0 else 0
    suppress_kw = base_kw - (base_kw * (0.85**3) * 1.06) # 假設夏季抑低需量

    return {
        "old_kwh": total_old_kwh,
        "save_kwh": save_kwh,
        "save_money": save_money,
        "save_rate": save_rate,
        "payback": payback,
        "details": calc_details,
        "suppress_kw": suppress_kw
    }

# --- 4. Word 生成函數 (混合文字替換與表格插入) ---
def generate_word(res):
    try:
        doc = Document("template_p5.docx")
    except:
        st.error("❌ 找不到 template_p5.docx 檔案，請確認已上傳至 GitHub。")
        return None

    # A. 文字標籤替換
    data_map = {
        "{{貴單位}}": unit_name,
        "{{CH-1}}": ch_info.split("/")[0].strip(),
        "{{1200RT}}": ch_info.split("/")[1].strip() if "/" in ch_info else "",
        "{{三台 15hp}}": motor_spec,
        "{{僅開啟一台}}": setup_note,
        "{{110, 277}}": f"{res['old_kwh']:,.0f}",
        "{{110277}}": f"{res['old_kwh']:,.0f}",
        "{{42, 054}}": f"{res['save_kwh']:,.0f}",
        "{{42054}}": f"{res['save_kwh']:,.0f}",
        "{{19.5}}": f"{res['save_money']:.2f}",
        "{{13}}": f"{res['suppress_kw']:.1f}",
        "{{58.5}}": f"{invest_amt:.1f}",
        "{{3}}": f"{res['payback']:.1f}"
    }

    # 遍歷段落
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        fix_run_format(run)

    # 遍歷表格中的標籤
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in data_map.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, str(val))
                        if cell.paragraphs:
                            fix_run_format(cell.paragraphs[0].runs[0], size=10)

    # B. 動態表格插入 (取代 [[OLD_TABLE_PLACEHOLDER]])
    for p in doc.paragraphs:
        if "[[OLD_TABLE_PLACEHOLDER]]" in p.text:
            p.text = "" # 刪除標籤文字
            
            # 建立表格
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'
            
            # 填寫標頭
            hdr = ["季節", "馬力(hp)", "台數", "時數(hr)", "負載(%)", "耗電(kWh)"]
            for i, h in enumerate(hdr):
                cell = table.cell(0, i)
                cell.text = h
                fix_run_format(cell.paragraphs[0].runs[0], size=10, is_bold=True)
            
            # 填寫計算後的每一行數據
            for d in res['details']:
                row_cells = table.add_row().cells
                row_cells[0].text = d['季節']
                row_cells[1].text = "15.0"
                row_cells[2].text = "1.0"
                row_cells[3].text = f"{d['時數']:,.0f}"
                row_cells[4].text = d['負載']
                row_cells[5].text = f"{d['改善前kwh']:,.0f}"
                for c in row_cells:
                    if c.paragraphs: fix_run_format(c.paragraphs[0].runs[0], size=10)

            # 合計行
            total_row = table.add_row().cells
            total_row[0].text = "合計"
            total_row[5].text = f"{res['old_kwh']:,.0f}"
            fix_run_format(total_row[0].paragraphs[0].runs[0], size=10, is_bold=True)
            fix_run_format(total_row[5].paragraphs[0].runs[0], size=10, is_bold=True)

    return doc

# --- 5. 輸出中心 ---
st.markdown("---")
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    res = calculate_results(edit_op)
    final_doc = generate_word(res)
    
    if final_doc:
        buf = io.BytesIO()
        final_doc.save(buf)
        report_data = buf.getvalue()
        
        # 存入打包倉庫
        if 'report_warehouse' not in st.session_state:
            st.session_state['report_warehouse'] = {}
        st.session_state['report_warehouse']["5. 風車加裝變頻器"] = report_data
        
        st.success("✅ 報告生成成功！數據已填入範本並修正為黑色標楷體。")
        st.download_button("📥 下載此份 Word 報告", report_data, "風車變頻器效益分析.docx", use_container_width=True)
