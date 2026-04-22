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

# --- 2. 介面設計 ---
st.set_page_config(layout="wide")
st.title("💡 P6. 照明系統節能效益分析 (表九之二整合版)")

# A. 基礎參數設定
with st.container():
    c1, c2, c3 = st.columns(3)
    unit_name = c1.text_input("單位名稱", value="貴單位")
    elec_price = c2.number_input("平均電費 (元/度)", value=3.5)
    wage_cost = c3.number_input("平均每盞施工費 (元)", value=150)

# B. 資料來源載入 (Excel 上傳)
st.subheader("📂 步驟一：載入數據")
uploaded_file = st.file_uploader("請上傳「表九之二」或照明明細 Excel", type=["xlsx"])

# 初始化照明數據清單
if "lighting_data" not in st.session_state:
    st.session_state.lighting_data = [
        {"area": "範例區域", "type": "T8 20W*2", "old_w": 46, "qty": 10, "hr": 3000, "led_w": 18, "led_price": 250}
    ]

# 處理 Excel 上傳邏輯
if uploaded_file:
    try:
        # 讀取 Excel (假設數據在第一個 Sheet)
        df_import = pd.read_excel(uploaded_file)
        
        # 欄位對接地圖 (依據能源局標準表九之二常見關鍵字)
        mapping = {
            '區域名稱': 'area', '區域': 'area',
            '現有燈具形式': 'type', '燈具形式': 'type',
            '現有燈具功率(W)': 'old_w', '單盞功率': 'old_w',
            '現有數量': 'qty', '數量': 'qty',
            '年運轉時數': 'hr', '運轉時數': 'hr'
        }
        df_import = df_import.rename(columns=mapping)
        
        # 只取我們需要的欄位，並補足 LED 相關空白欄位
        needed_cols = ['area', 'type', 'old_w', 'qty', 'hr']
        # 過濾出存在的欄位
        existing_cols = [c for c in needed_cols if c in df_import.columns]
        df_final = df_import[existing_cols].copy()
        
        # 補足缺失的 LED 參數欄位供使用者編輯
        if 'led_w' not in df_final.columns: df_final['led_w'] = 18
        if 'led_price' not in df_final.columns: df_final['led_price'] = 300
        
        st.session_state.lighting_data = df_final.to_dict('records')
        st.success("✅ Excel 數據載入成功！請在下方確認並編輯 LED 參數。")
    except Exception as e:
        st.error(f"讀取 Excel 出錯：{e}")

# C. 數據編輯區
st.subheader("📝 步驟二：確認並編輯燈具數據")
# 使用 data_editor，使用者可以直接在這裡貼上 Excel 數據或手動修改
df_lighting = st.data_editor(pd.DataFrame(st.session_state.lighting_data), num_rows="dynamic", use_container_width=True)

# --- 3. 報告生成邏輯 ---
st.divider()
if st.button("🚀 步驟三：生成並下載報告", use_container_width=True):
    try:
        doc = Document()
        # 設置頁面為橫向 (因欄位較多)
        section = doc.sections[0]
        section.orientation = 1 # WD_ORIENT.LANDSCAPE 
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height

        doc.add_heading(f'{unit_name} 照明系統 LED 汰換效益分析表', 0)

        # 建立 10 欄表格 (對應截圖樣式)
        table = doc.add_table(rows=1, cols=10)
        table.style = 'Table Grid'
        set_table_border(table)

        headers = ["區域", "原燈具形式", "數量", "年時數", "原單盞(W)", "改善後(W)", "原耗電(kWh)", "新耗電(kWh)", "節電量(kWh)", "投資額(元)"]
        for i, h in enumerate(headers):
            table.cell(0, i).text = h
            fix_cell_font(table.cell(0, i), is_bold=True)

        total_old_kwh = 0
        total_new_kwh = 0
        total_invest = 0

        for _, row in df_lighting.iterrows():
            # 數值預處理
            qty = float(row.get('qty', 0))
            hr = float(row.get('hr', 0))
            old_w = float(row.get('old_w', 0))
            led_w = float(row.get('led_w', 0))
            led_p = float(row.get('led_price', 0))

            # 計算
            o_kwh = (old_w * qty * hr) / 1000
            n_kwh = (led_w * qty * hr) / 1000
            s_kwh = o_kwh - n_kwh
            inv = (led_p + wage_cost) * qty

            total_old_kwh += o_kwh
            total_new_kwh += n_kwh
            total_invest += inv

            # 填表
            cells = table.add_row().cells
            cells[0].text = str(row.get('area', ''))
            cells[1].text = str(row.get('type', ''))
            cells[2].text = f"{qty:,.0f}"
            cells[3].text = f"{hr:,.0f}"
            cells[4].text = f"{old_w}"
            cells[5].text = f"{led_w}"
            cells[6].text = f"{o_kwh:,.0f}"
            cells[7].text = f"{n_kwh:,.0f}"
            cells[8].text = f"{s_kwh:,.0f}"
            cells[9].text = f"{inv:,.0f}"
            for c in cells: fix_cell_font(c)

        # 合計列
        total_save_kwh = total_old_kwh - total_new_kwh
        last_row = table.add_row().cells
        last_row[0].text = "合計"
        last_row[6].text = f"{total_old_kwh:,.0f}"
        last_row[7].text = f"{total_new_kwh:,.0f}"
        last_row[8].text = f"{total_save_kwh:,.0f}"
        last_row[9].text = f"{total_invest:,.0f}"
        for i in [0, 6, 7, 8, 9]: fix_cell_font(last_row[i], is_bold=True)

        # 效益分析
        save_money = total_save_kwh * elec_price / 10000
        payback = total_invest / (save_money * 10000) if save_money > 0 else 0

        summary = doc.add_paragraph()
        summary.add_run(f"\n【節能效益評估總結】").bold = True
        doc.add_paragraph(f"1. 預估年度總節電量：{total_save_kwh:,.0f} kWh/年")
        doc.add_paragraph(f"2. 預估年度節省電費：{save_money:.2f} 萬元/年 (以 {elec_price} 元/度計算)")
        doc.add_paragraph(f"3. 預估總投資金額：{total_invest:,.0f} 元 (含施工費)")
        doc.add_paragraph(f"4. 投資回收年限：約 {payback:.1f} 年")

        # 儲存並下載
        buf = io.BytesIO()
        doc.save(buf)
        st.success("🎉 報告已成功生成！")
        st.download_button(
            label="📥 下載照明分析 Word 報告",
            data=buf.getvalue(),
            file_name=f"P6_照明節能報告_{unit_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"❌ 生成報告時發生錯誤：{e}")
