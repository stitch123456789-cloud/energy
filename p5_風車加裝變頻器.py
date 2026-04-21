import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 格式保護工具 ---
def set_table_border(table):
    """手動強制繪製表格黑色框線 (防止樣式報錯)"""
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') 
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

def fix_cell_font(cell):
    """表格內容固定字體"""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 核心：保留格式替換函數 ---
def smart_replace(doc, data_map):
    """這是不會破壞原本紅色、粗體等格式的替換法"""
    # 處理段落
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if key in p.text:
                # 遍歷段落中的每個 Run
                for run in p.runs:
                    if key in run.text:
                        # 只替換文字，Run 內建的顏色/字體會被保留
                        run.text = run.text.replace(key, str(val))
    
    # 處理表格格子內的標籤
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        if key in p.text:
                            for run in p.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(val))

# --- 3. 介面與計算 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0)
    elec_val = st.number_input("平均電費 (元/度)", value=3.5)
with c3:
    rt_info = st.text_input("容量 (RT)", value="1500RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

if "p5_data" not in st.session_state:
    st.session_state.p5_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"], "時數(hr)": [4380, 2190, 2190], "負載率(%)": [70, 85, 60]
    })
current_df = st.data_editor(st.session_state.p5_data, use_container_width=True)

# --- 4. 生成按鈕 ---
if st.button("🚀 生成報告 (保留原始格式)", use_container_width=True):
    try:
        # 計算
        base_kw = motor_hp * 0.746
        total_old, total_new = 0, 0
        details = []
        for _, r in current_df.iterrows():
            h, l = float(r["時數(hr)"]), float(r["負載率(%)"])/100
            o, n = base_kw * h, base_kw * (l**3) * 1.06 * h
            details.append({"季節": r["季節"], "時數": h, "負載": f"{r['負載率(%)']}%", "舊": o, "新": n, "省": o-n})
            total_old += o
            total_new += n

        save_kwh = total_old - total_new
        save_money = save_kwh * elec_val / 10000
        payback = invest_amt / save_money if save_money > 0 else 0

        doc = Document("template_p5.docx")

        # A. 文字替換 (使用 smart_replace，保留紅色格式)
        data_map = {
            "{{UN}}": unit_name, "{{COUNT}}": "2", "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info, "{{MT}}": f"{int(motor_hp)}hp",
            "{{ON}}": setup_note, "{{OLD_KWH}}": f"{total_old:,.0f}",
            "{{SAVE_KWH}}": f"{save_kwh:,.0f}", "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_MONEY}}": f"{save_money:.2f}", "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}", "{{SUPPRESS_KW}}": "13"
        }
        smart_replace(doc, data_map)

        # B. 插入表格 (這段邏輯不影響文字格式)
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" 
                tbl = doc.add_table(rows=1, cols=4)
                set_table_border(tbl)
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, txt in enumerate(hdr):
                    tbl.cell(0, i).text = txt
                for d in details:
                    row = tbl.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
                for cell in tbl._element.xpath('.//w:tc'): fix_cell_font(tbl.cell(0,0)) # 簡化調用

            if "[[NEW_TABLE]]" in p.text:
                p.text = ""
                tbl = doc.add_table(rows=1, cols=5)
                set_table_border(tbl)
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, txt in enumerate(hdr):
                    tbl.cell(0, i).text = txt
                for d in details:
                    row = tbl.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 格式完美保留！標籤已換為數字。")
        st.download_button("📥 下載成果檔案", buf.getvalue(), "風車分析報告_格式保留版.docx")

    except Exception as e:
        st.error(f"錯誤: {e}")
