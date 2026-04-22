import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import os

# --- 1. 核心工具函數 ---
def fix_cell_font(cell, size=10, is_bold=False):
    for paragraph in cell.paragraphs:
        if not paragraph.runs: paragraph.add_run()
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(size)
            run.font.bold = is_bold

def safe_replace(doc, data_map):
    """暴力替換邏輯：確保碎裂標籤能正確替換並維持格式"""
    for p in doc.paragraphs:
        inline_text = "".join([run.text for run in p.runs])
        for key, val in data_map.items():
            if key in inline_text:
                # 重新組合段落文字
                new_text = inline_text.replace(key, str(val))
                p.clear()
                run = p.add_run(new_text)
                run.font.name = '標楷體'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                run.font.size = Pt(12)

# --- 2. 數據抓取與計算 ---
st.title("💡 P6. 日光燈汰換 LED 分析")

# 定義範本路徑 (使用絕對路徑防止 Package not found)
current_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(current_dir, "template_5A03.docx")

# 常數設定 (依據你的截圖邏輯)
LED_SAVE_RATIO = 0.464537  # 節能率
INVEST_PER_KW = 24677     # 投資單價 (元/kW)
ELEC_PRICE = st.session_state.get('auto_avg_price', 4.62)

global_file = st.session_state.get('global_excel')

if global_file is not None:
    try:
        xl = pd.ExcelFile(global_file)
        target_s = [s for s in xl.sheet_names if "九之二" in s]
        
        if target_s:
            # 讀取 Excel (不設 header，手動抓索引)
            df_raw = pd.read_excel(global_file, sheet_name=target_s[0], header=None)
            
            # 過濾「1. 日光燈」資料列 (通常從第7列開始，索引為6)
            # 根據截圖：B欄(index 1)是種類, I欄(index 8)是瓦數, K欄(index 10)是數量, L欄(index 11)是時數
            data_rows = []
            for i in range(6, len(df_raw)):
                row = df_raw.iloc[i]
                lamp_type = str(row[1])
                if "日光燈" in lamp_type:
                    data_rows.append({
                        "type": lamp_type,
                        "spec": str(row[5]),  # F欄:容量規格
                        "qty": float(row[10]) if pd.notnull(row[10]) else 0,
                        "kw": float(row[11]) if pd.notnull(row[11]) else 0, # L欄其實是耗電瓩(kW)
                        "hr": float(row[12]) if pd.notnull(row[12]) else 0  # M欄是時數
                    })
            
            if data_rows:
                st.success(f"✅ 已自動識別 {len(data_rows)} 筆日光燈資料")
                
                # 計算總計
                total_old_kw = sum(r['kw'] for r in data_rows)
                total_old_kwh = sum(r['kw'] * r['hr'] for r in data_rows)
                save_kwh = total_old_kwh * LED_SAVE_RATIO
                save_money = save_kwh * ELEC_PRICE / 10000 # 萬元
                invest = total_old_kw * INVEST_PER_KW / 10000 # 萬元
                payback = invest / save_money if save_money > 0 else 0
                save_kw_peak = total_old_kw * LED_SAVE_RATIO

                if st.button("🚀 生成 P6 LED 改善報告"):
                    if not os.path.exists(template_path):
                        st.error(f"找不到範本檔案：{template_path}，請確認檔案已上傳至正確資料夾。")
                    else:
                        doc = Document(template_path)
                        
                        # 標籤映射地圖 (對應 Word 裡的紅字標籤)
                        data_map = {
                            "{{OLD_KWH}}": f"{total_old_kwh:,.0f}",
                            "{{SAVE_KWH}}": f"{save_kwh:,.0f}",
                            "{{SAVE_MONEY}}": f"{save_money:.2f}",
                            "{{ENERGY_RATE}}": f"{LED_SAVE_RATIO*100:.2f}",
                            "{{INVEST}}": f"{invest:.1f}",
                            "{{PAYBACK}}": f"{payback:.1f}",
                            "{{SAVE_KW}}": f"{save_kw_peak:.1f}"
                        }
                        
                        safe_replace(doc, data_map)
                        
                        # 處理 Word 內現況表 (尋找標籤 [[LAMP_TABLE]])
                        # 如果需要手動填寫表格，可在此處加入 table 遍歷邏輯

                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state['report_warehouse'][f"P6_LED照明_{unit_name}"] = buf.getvalue()
                        st.success("🎉 報告已生成！請至側邊欄打包下載或直接點擊下方按鈕。")
                        st.download_button("📥 下載單份 P6 報告", buf.getvalue(), "P6_LED報告.docx")
            else:
                st.warning("查無日光燈資料列，請檢查 Excel 表九之二。")
    except Exception as e:
        st.error(f"分析失敗：{str(e)}")
