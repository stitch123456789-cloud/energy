import streamlit as st
import pandas as pd
from docx import Document
import io
import re

# ── 固定參數 ──
LED_SAVE_RATIO = 0.464537  # 節省比例
ELEC_PRICE = 4.6238        # 平均電費
INVEST_PER_KW = 24677     # 投資單價

def run_p6_logic():
    st.subheader("💡 P6. 日光燈汰換 LED 分析")

    # 檢查是否有全域 Excel
    if 'global_excel' not in st.session_state or st.session_state['global_excel'] is None:
        st.error("❌ 請先在側邊欄上傳能源查核 Excel 總表！")
        return

    excel_file = st.session_state['global_excel']

    try:
        # 1. 讀取 Excel (表九之二)
        xl = pd.ExcelFile(excel_file)
        lighting_sheet = next((s for s in xl.sheet_names if "表九之二" in s), None)
        
        if not lighting_sheet:
            st.error("❌ Excel 中找不到「表九之二」工作表")
            return

        # 這裡根據你提供的邏輯讀取數據
        df = pd.read_excel(excel_file, sheet_name=lighting_sheet, skiprows=2) 
        
        # 2. 過濾日光燈數據 (假設第一欄包含 "1.日光燈")
        # 注意：這裡要根據你 Excel 的實際欄位索引調整
        old_lamps = df[df.iloc[:, 1].astype(str).str.contains("1.日光燈", na=False)]

        if old_lamps.empty:
            st.warning("⚠️ 找不到「1.日光燈」的資料，請確認 Excel 內容。")
            return

        # 3. 計算核心數值
        # 假設：[9]數量, [10]kW, [11]時數
        total_qty = old_lamps.iloc[:, 9].sum()
        total_old_kw = old_lamps.iloc[:, 10].sum()
        avg_hours = old_lamps.iloc[:, 11].mean() # 或取加權平均

        total_old_kwh = sum(old_lamps.iloc[:, 10] * old_lamps.iloc[:, 11])
        save_kwh = total_old_kwh * LED_SAVE_RATIO
        save_money = save_kwh * ELEC_PRICE / 10000
        invest = total_old_kw * INVEST_PER_KW / 10000
        payback = invest / save_money if save_money > 0 else 0
        save_kw_peak = total_old_kw * LED_SAVE_RATIO

        # 4. 準備 Word 替換地圖
        # 這些標籤 {{...}} 必須手動打在你的 template_5A03.docx 裡面
        data_map = {
            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
            "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
            "{{SAVE_MONEY}}": f"{save_money:.2f}",
            "{{ENERGY_RATE}}": f"{LED_SAVE_RATIO*100:.2f}",
            "{{INVEST}}": f"{invest:.1f}",
            "{{PAYBACK}}": f"{payback:.1f}",
            "{{SAVE_KW}}": f"{save_kw_peak:.1f}"
        }

        st.success(f"✅ 已計算 {len(old_lamps)} 筆日光燈資料")
        
        # 5. 生成報告按鈕
        if st.button("🚀 生成 P6 LED 改善報告"):
            doc = Document("template_5A03.docx")
            
            # 使用你之前成功的 safe_replace 函數邏輯
            # (此處省略 safe_replace 定義，請引用你 app.py 裡的工具函數)
            # safe_replace(doc, data_map)
            
            buf = io.BytesIO()
            doc.save(buf)
            st.download_button("📥 下載 P6 報告", buf.getvalue(), "P6_LED改善報告.docx")

    except Exception as e:
        st.error(f"分析失敗：{str(e)}")

run_p6_logic()
