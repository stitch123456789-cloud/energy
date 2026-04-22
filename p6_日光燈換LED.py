import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 核心工具函數 ---
def set_table_border(table):
    tbl = table._tbl
    ptr = tbl.find(qn('w:tblPr'))
    if ptr is not None:
        borders = OxmlElement('w:tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            edge = OxmlElement(f'w:{b}')
            edge.set(qn('w:val'), 'single')
            edge.set(qn('w:sz'), '4') 
            edge.set(qn('w:color'), '000000')
            borders.append(edge)
        ptr.append(borders)

def fix_cell_font(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        paragraph.alignment = 1 
        if not paragraph.runs: paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold

# --- 2. 數據抓取邏輯 (對接 app.py) ---

# 初始化數據容器
if "lighting_data" not in st.session_state:
    st.session_state.lighting_data = []

# 嘗試從 app.py 的全域 Excel 中抓取「表九之二」
global_file = st.session_state.get('global_excel')

if global_file is not None and not st.session_state.lighting_data:
    try:
        # 讀取所有 Sheet 名字
        all_sheets = pd.ExcelFile(global_file).sheet_names
        # 尋找包含 "九之二" 字眼的 Sheet
        target_s = [s for s in all_sheets if "九之二" in s]
        
        if target_s:
            # 讀取表九之二 (假設數據從第 4 行標題開始，請根據實況調整 skiprows)
            df_raw = pd.read_excel(global_file, sheet_name=target_s[0], skiprows=2)
            
            # 能源局標準格式對接
            mapping = {
                '區域名稱': 'area', '現有燈具形式': 'type',
                '現有燈具功率(W)': 'old_w', '現有數量': 'qty', '年運轉時數': 'hr'
            }
            df_raw = df_raw.rename(columns=mapping)
            
            # 過濾出必要的欄位
            valid_cols = [c for c in mapping.values() if c in df_raw.columns]
            df_final = df_raw[valid_cols].copy()
            
            # 補足 LED 預設參數
            df_final['led_w'] = 18
            df_final['led_price'] = 250
            
            st.session_state.lighting_data = df_final.to_dict('records')
            st.toast("✅ 已成功從總表抓取照明數據 (表九之二)")
    except Exception as e:
        st.warning(f"自動抓取總表失敗，請改用手動編輯：{e}")

# --- 3. 介面設計 ---
st.title("💡 P6. 照明系統節能效益分析")

# 基礎參數 (平均電費自動對接 app.py 的計算結果)
avg_p = st.session_state.get('auto_avg_price', 3.5)

c1, c2, c3 = st.columns(3)
u_name = c1.text_input("單位名稱", value="貴單位")
e_price = c2.number_input("平均電費 (元/度)", value=float(avg_p))
w_cost = c3.number_input("平均施工費 (元/盞)", value=150)

# 數據編輯區
st.subheader("📝 燈具數據確認與編輯")
if not st.session_state.lighting_data:
    st.info("💡 目前全域資料庫無照明數據，請在下方表格直接貼上資料或手動輸入。")
    # 給一個預設空白列
    st.session_state.lighting_data = [{"area": "", "type": "", "old_w": 46, "qty": 0, "hr": 3000, "led_w": 18, "led_price": 250}]

df_lighting = st.data_editor(pd.DataFrame(st.session_state.lighting_data), num_rows="dynamic", use_container_width=True)

# --- 4. 報告生成與存入 warehouse ---
if st.button("🚀 生成 P6 照明報告", use_container_width=True):
    try:
        doc = Document()
        # 設定橫向
        section = doc.sections[0]
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

        doc.add_heading(f'{u_name} 照明節能效益分析', 0)

        # 表格生成
        table = doc.add_table(rows=1, cols=10)
        table.style = 'Table Grid'
        set_table_border(table)
        
        headers = ["區域", "原形式", "數量", "時數", "原W", "新W", "原kWh", "新kWh", "節電量", "投資額"]
        for i, h in enumerate(headers):
            table.cell(0, i).text = h
            fix_cell_font(table.cell(0, i), is_bold=True)

        t_old_kwh, t_new_kwh, t_inv = 0, 0, 0

        for _, row in df_lighting.iterrows():
            qty, hr = float(row.get('qty', 0)), float(row.get('hr', 0))
            o_w, n_w = float(row.get('old_w', 0)), float(row.get('led_w', 0))
            lp = float(row.get('led_price', 0))

            o_kwh = (o_w * qty * hr) / 1000
            n_kwh = (n_w * qty * hr) / 1000
            inv = (lp + w_cost) * qty

            t_old_kwh += o_kwh
            t_new_kwh += n_kwh
            t_inv += inv

            cells = table.add_row().cells
            cells[0].text = str(row.get('area', ''))
            cells[1].text = str(row.get('type', ''))
            cells[2].text = f"{qty:,.0f}"
            cells[3].text = f"{hr:,.0f}"
            cells[4].text = str(o_w)
            cells[5].text = str(n_w)
            cells[6].text = f"{o_kwh:,.0f}"
            cells[7].text = f"{n_kwh:,.0f}"
            cells[8].text = f"{(o_kwh - n_kwh):,.0f}"
            cells[9].text = f"{inv:,.0f}"
            for c in cells: fix_cell_font(c)

        # 存入 warehouse (讓 app.py 的側邊欄可以打包)
        buf = io.BytesIO()
        doc.save(buf)
        report_name = f"P6_照明分析_{u_name}"
        st.session_state['report_warehouse'][report_name] = buf.getvalue()
        
        st.success(f"🎉 報告已生成並存入輸出中心！(目前共 {len(st.session_state['report_warehouse'])} 份)")
        st.download_button("📥 直接下載此份報告", buf.getvalue(), f"{report_name}.docx")

    except Exception as e:
        st.error(f"生成失敗：{e}")
