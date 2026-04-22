import streamlit as st
import pandas as pd

def scan_all_buildings_for_fluorescent(uploaded_file):
    """
    掃描上傳的 Excel 中所有分頁，找出所有建築物裡的「日光燈/傳統燈具」
    """
    try:
        # sheet_name=None 會將所有分頁讀取為一個 Dictionary (字典)
        # key 是分頁名稱，value 是該分頁的 DataFrame
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        target_lights_list = []
        
        # 迴圈檢查每一個分頁
        for sheet_name, df in all_sheets.items():
            # 假設照明資料都在「表九」相關的分頁裡
            if '表九' in sheet_name:
                
                # 確保該分頁裡面有「種類」這個欄位再進行判斷
                if '種類' in df.columns:
                    
                    # 判斷邏輯：找出名稱包含 日光燈、螢光燈、T8、T5，且「不包含」LED 的設備
                    # na=False 是為了略過空白儲存格避免報錯
                    is_fluorescent = df['種類'].str.contains('日光燈|螢光燈|T8|T5|傳統', na=False)
                    is_not_led = ~df['種類'].str.contains('LED', na=False, case=False)
                    
                    # 取交集：是傳統燈具，且不是 LED
                    found_lights = df[is_fluorescent & is_not_led].copy()
                    
                    if not found_lights.empty:
                        # 標註這些燈具是從哪一棟建築物找出來的
                        found_lights['來源建築物'] = sheet_name
                        target_lights_list.append(found_lights)
        
        # 如果有找到任何傳統燈具，就把所有建築物的資料合併成一張大表
        if target_lights_list:
            final_df = pd.concat(target_lights_list, ignore_index=True)
            return final_df
        else:
            return None
            
    except Exception as e:
        st.error(f"讀取 Excel 發生錯誤：{e}")
        return None

# ==========================================
# Streamlit 介面實裝範例
# ==========================================

st.title("💡 LED 燈具汰換自動掃描系統")

# 模擬你左側的檔案上傳區塊
uploaded_file = st.sidebar.file_uploader("上傳完整能源查核 Excel", type=["xlsx"])

if uploaded_file is not None:
    st.info("🔄 正在掃描全廠區各棟建築物資料...")
    
    # 執行掃描引擎
    fluorescent_df = scan_all_buildings_for_fluorescent(uploaded_file)
    
    if fluorescent_df is not None:
        # 計算你在 Word 模板裡需要的數據
        total_non_led_count = fluorescent_df['數量(具)'].sum()
        st.success(f"⚠️ 掃描完畢！在不同建築物中共發現 **{total_non_led_count:,.0f}** 具傳統日光燈/螢光燈。")
        
        # 顯示統整後的清單給你看
        st.write("📊 **待改善燈具清單統整：**")
        st.dataframe(fluorescent_df[['來源建築物', '種類', '瓦數(W/具)', '數量(具)']])
        
        # 這裡可以接續寫計算公式，算出 {{SAVING_KWH}} 等數值
        # ...
        
        if st.button("🚀 生成 LED 汰換提案報告"):
            st.balloons()
            st.write("這裡就會執行 python-docx 去替換你上傳的 Word 模板！")
            
    else:
        st.success("✅ 掃描完畢！所有建築物皆未發現傳統日光燈，或已全數汰換為 LED。")
