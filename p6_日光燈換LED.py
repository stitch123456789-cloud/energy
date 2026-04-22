import streamlit as st
import pandas as pd
from docx import Document
# ... (導入你之前常用的工具函數 set_table_border, fix_cell_font, safe_replace)

st.title("💡 P6. 照明系統汰換 LED 分析")

# 1. 基礎參數設定
col1, col2 = st.columns(2)
with col1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    elec_price = st.number_input("平均電費 (元/度)", value=3.5)
with col2:
    wage_per_lamp = st.number_input("每盞換裝工資 (元)", value=100)

# 2. 動態區域管理 (例如：辦公區、地下室、廠區)
if "lighting_areas" not in st.session_state:
    st.session_state.lighting_areas = [{"name": "辦公區", "type": "T8燈管 20Wx2", "old_w": 46, "qty": 100, "hr": 3000, "led_w": 18, "led_price": 250}]

with st.sidebar:
    st.header("⚙️ 區域管理")
    if st.button("➕ 新增照明區域"):
        st.session_state.lighting_areas.append({"name": f"新區域", "type": "T8燈管", "old_w": 46, "qty": 10, "hr": 2500, "led_w": 18, "led_price": 250})
        st.rerun()
    if st.button("❌ 刪除最後一區"):
        if len(st.session_state.lighting_areas) > 1:
            st.session_state.lighting_areas.pop()
            st.rerun()

# 3. 介面輸入：每一區的詳細參數
area_results = []
for i, area in enumerate(st.session_state.lighting_areas):
    with st.expander(f"📍 區域：{area['name']}", expanded=True):
        c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
        area['name'] = c1.text_input("區域描述", value=area['name'], key=f"an_{i}")
        area['type'] = c2.text_input("原燈具形式", value=area['type'], key=f"at_{i}")
        area['qty'] = c3.number_input("數量 (盞)", value=area['qty'], key=f"aq_{i}")
        area['hr'] = c4.number_input("年時數 (hr)", value=area['hr'], key=f"ah_{i}")
        
        c5, c6, c7 = st.columns(3)
        area['old_w'] = c5.number_input("原單盞總功率 (W)", value=area['old_w'], help="含安定器損耗", key=f"aw_{i}")
        area['led_w'] = c6.number_input("LED單盞功率 (W)", value=area['led_w'], key=f"lw_{i}")
        area['led_price'] = c7.number_input("LED單價 (元)", value=area['led_price'], key=f"lp_{i}")

# 4. 點擊生成報告
if st.button("🚀 生成照明汰換報告", use_container_width=True):
    # 計算邏輯...
    # Word 表格生成 (建議用直式表格，一區一行)...
    st.success("照明分析完成！")
