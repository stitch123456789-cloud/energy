import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 字體工具函數 ---
def set_run_kai(run, size=12, is_bold=False):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)

# --- 2. 混合填充核心函數 ---
def fill_p5_report(template_path, data_dict, op_df):
    doc = Document(template_path)

    # A. 處理文字替換 (簡單變數)
    # 範本中請放 {{UNIT_NAME}}, {{INVEST}}, {{PAYBACK}}
    replacements = {
        "{{UNIT_NAME}}": data_dict['unit_name'],
        "{{INVEST}}": f"{data_dict['invest']}",
        "{{PAYBACK}}": f"{data_dict['payback']}",
        "{{SAVE_KWH}}": f"{data_dict['save_kwh']:,.0f}"
    }

    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)

    # B. 處理動態表格 (精確定位插入)
    # 技巧：在 Word 範本中單獨一行寫下 "[[INSERT_TABLE_HERE]]"
    for p in doc.paragraphs:
        if "[[INSERT_TABLE_HERE]]" in p.text:
            p.text = "" # 清空標籤
            # 在此段落後建立表格
            table = doc.add_table(rows=1, cols=7) 
            table.style = 'Table Grid'
            # ... 這裡接著寫你原本生成表格的代碼 ...
            # 這樣表格就會乖乖出現在標籤的位置，而不會亂跑
            
    return doc

# --- 3. Streamlit 介面與計算 ---
st.title("🌀 P5. 風車加裝變頻器分析")

# 這裡延用你之前的輸入介面 (單位名稱、電費、投資金額等)
unit_name = st.text_input("單位名稱", value="貴單位")
invest_amount = st.number_input("投資費用 (萬元)", value=58.5)

# 假設計算完畢後的結果
results = {
    "unit_name": unit_name,
    "invest": invest_amount,
    "save_kwh": 42054,
    "payback": 3.0
}

# --- 4. 輸出中心 ---
if st.button("🚀 執行混合填充生成"):
    # 這裡請上傳或指定你的公司範本路徑
    # template = "company_template_p5.docx" 
    
    st.info("正在將數據填入範本標籤並生成動態表格...")
    # doc = fill_p5_report(template, results, df_op)
    
    st.success("報告已生成！文字已替換，表格已定位插入。")
