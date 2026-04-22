import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 ---

def set_table_border(table):
    """手動為表格添加黑色框線，防止部分 Word 版本看不到格子"""
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

def fix_cell_font(cell, size=10, is_bold=False):
    """表格內容強制格式化為標楷體"""
    for paragraph in cell.paragraphs:
        if not paragraph.runs:
            paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

def safe_replace(doc, data_map):
    """強化版替換工具：先處理碎裂標籤，再處理表格內標籤"""
    # 處理所有段落
    for p in doc.paragraphs:
        # 檢查段落內是否含有標籤
        inline_text = "".join([run.text for run in p.runs])
        for key, val in data_map.items():
            if key in inline_text:
                # 重新組合文字並替換
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                    elif key[0:2] in run.text: # 處理碎裂標籤的保險邏輯
                        # 如果標籤被切斷，直接在段落層級處理
                        p.text = p.text.replace(key, str(val))
                # 強制修正格式
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.font.name = '標楷體'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    # 處理所有表格內容
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_cell_text = "".join([run.text for run in p.runs])
                    for key, val in data_map.items():
                        if key in full_cell_text:
                            # 直接對段落文字進行替換以保證成功
                            p.text = p.text.replace(key, str(val))
                            # 補回格式
                            for run in p.runs:
                                run.font.color.rgb = RGBColor(0, 0, 0)
                                run.font.name = '標楷體'
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車加裝變頻器")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
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
current_op_df = st.data_editor(st.session_state.p5_op_data, use_container_width=True)

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
    
    save_money = (total_old - total_new) * elec_input / 10000
    payback = invest_amt / save_money if save_money > 0 else 0
    return {
        "old_total": total_old, "save_kwh": total_old - total_new, "save_money": save_money, 
        "payback": payback, "save_rate": ((total_old-total_new)/total_old*100), "details": details
    }

# --- 4. 生成按鈕 ---
if st.button("🚀 生成 P5 變頻器報告", use_container_width=True):
    results = run_calculation(current_op_df)
    
    try:
        doc = Document("template_p5.docx")
        
        data_map = {
            "{{UN}}": unit_name, "{{COUNT}}": "2", "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info, "{{MT}}": f"三台 {int(motor_hp)}hp",
            "{{ON}}": setup_note,
            "{{OLD_KWH}}": f"{results['old_total']:,.0f}",
            "{{SAVE_KWH}}": f"{results['save_kwh']:,.0f}",
            "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_RATE}}": f"{results['save_rate']:.2f}",
            "{{SAVE_MONEY}}": f"{results['save_money']:.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}",
            "{{PAYBACK}}": f"{results['payback']:.1f}",
            "{{SUPPRESS_KW}}": "13",
            "{{13}}": "13"
        }
        
        # 1. 執行核心文字替換
        safe_replace(doc, data_map)
        
        # 2. 清除標籤定位符號
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text: p.text = ""
            if "[[NEW_TABLE]]" in p.text: p.text = ""

        # 3. 在文末生成橫向橫式表格 (依照您要求的 CT-1 格式)
        doc.add_page_break()
        doc.add_paragraph("--- 以下為自動生成的橫向表格 (請剪下並貼至指定位置) ---")

        base_kw = motor_hp * 0.746
        num_cols = len(results['details']) + 2
        table = doc.add_table(rows=7, cols=num_cols)
        set_table_border(table)

        # 填寫左側標題
        labels = ["編號", "水塔散熱噸數(RT)", "額定馬力(hp)", "實際耗功(kW)", "全年使用時數(hr)", "負載率(%)", "全年耗電(kWh)"]
        for r_idx, label in enumerate(labels):
            fix_cell_font(table.cell(r_idx, 0), is_bold=True)
            table.cell(r_idx, 0).text = label

        # 合併第一列編號與第二列 RT
        merged_header = table.cell(0, 1).merge(table.cell(0, num_cols-2))
        merged_header.text = ch_info
        fix_cell_font(merged_header, is_bold=True)
        
        merged_rt = table.cell(1, 1).merge(table.cell(1, num_cols-2))
        merged_rt.text = rt_info
        fix_cell_font(merged_rt)

        # 填寫數據
        for c_idx, d in enumerate(results['details'], start=1):
            table.cell(2, c_idx).text = f"{motor_hp:.1f}"
            table.cell(3, c_idx).text = f"{base_kw:.1f}"
            table.cell(4, c_idx).text = f"{d['時數']:,.0f}"
            table.cell(5, c_idx).text = "100%"
            table.cell(6, c_idx).text = f"{d['舊']:,.0f}"
            for r in range(2, 7):
                fix_cell_font(table.cell(r, c_idx), is_bold=(r==6))

        # 合計欄
        table.cell(0, num_cols-1).text = "合計"
        table.cell(3, num_cols-1).text = f"{base_kw * len(results['details']):.1f}"
        table.cell(6, num_cols-1).text = f"{results['old_total']:,.0f}"
        for r in [0, 3, 6]:
            fix_cell_font(table.cell(r, num_cols-1), is_bold=True)

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 融合報告生成成功！文字與橫向表格皆已處理。")
        st.download_button("📥 下載 Word 報告", buf.getvalue(), "風車分析整合版.docx")
        
    except Exception as e:
        st.error(f"融合執行失敗: {e}")
