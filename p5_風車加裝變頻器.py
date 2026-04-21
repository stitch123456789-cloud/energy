import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 定義字體修正函數 ---
def fix_format(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0) # 強制轉回黑色

# --- 2. 執行替換的邏輯 ---
def process_report(template_path, results):
    doc = Document(template_path)
    
    # 建立與你 Word 標籤對應的數據地圖
    data_map = {
        "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
        "{{SAVE_RATE}}": f"{results['save_rate']:.1f}",
        "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
        "{{INVEST}}": f"{results['invest']:.1f}",
        "{{PAYBACK}}": f"{results['payback']:.1f}"
    }

    # A. 替換段落文字 (處理 {{...}})
    for p in doc.paragraphs:
        for key, value in data_map.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
                        fix_format(run, size=12) # 修正顏色為黑色

    # B. 替換表格內文字
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in data_map.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, str(value))
                        # 表格通常字級稍小，設為 10
                        if cell.paragraphs:
                            for run in cell.paragraphs[0].runs:
                                fix_format(run, size=10)

    # C. 處理 [[CT_TABLE]] 插入動態表格
    # (此處可加入 add_table 邏輯，或如果你已經在 Word 畫好格子就直接用 {{}} 填入)
    
    return doc

# --- 3. Streamlit 輸出中心 ---
# 計算完 results 後呼叫...
# final_doc = process_report("template_p5.docx", results)
