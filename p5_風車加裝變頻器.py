import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 表格格式修正工具 ---
def fix_cell_format(cell, size=10, is_bold=False):
    """確保表格內的每一格都是標楷體、黑色"""
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 介面與計算 (保持不變) ---
st.title("🌀 P5. 風車加裝變頻器分析")
# ... (這裡請保留您原本的 c1, c2, c3 輸入框代碼) ...

# 假設計算結果 res 已經產出 (內含 res['details'], res['old_total'], res['save_kwh'])

# --- 3. 核心生成邏輯 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    # 執行您的計算函數獲取 res
    res = run_calculation(current_op_df) 
    
    try:
        doc = Document("template_p5.docx")

        # --- A. 表格插入 (優先處理，避免標籤被文字替換干擾) ---
        for p in doc.paragraphs:
            # 1. 現況表格
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" # 清除標籤
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                # 標頭
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    table.cell(0, i).text = text
                    fix_cell_format(table.cell(0, i), is_bold=True)
                # 資料列
                for d in res['details']:
                    row = table.add_row().cells
                    row[0].text = d['季節']
                    row[1].text = f"{d['時數']:,.0f}"
                    row[2].text = "100%"
                    row[3].text = f"{d['舊']:,.0f}"
                    for c in row: fix_cell_format(c)
                # 合計
                tot = table.add_row().cells
                tot[0].text = "合計"; tot[3].text = f"{res['old_total']:,.0f}"
                for c in tot: fix_cell_format(c, is_bold=True)

            # 2. 效益表格
            if "[[NEW_TABLE]]" in p.text:
                p.text = ""
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

        # --- B. 文字標籤替換 (最後處理) ---
        # 這裡放入您原本成功的文字替換代碼 (safe_replace 邏輯)
        # ... 

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 表格已動態生成！")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車效益分析.docx")

    except Exception as e:
        st.error(f"錯誤: {e}")
