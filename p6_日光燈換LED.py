import streamlit as st
import pandas as pd

st.title("⚙️ 表九之二：設備資料提取")

# 從全域抓取已上傳的 Excel 檔案
uploaded_file = st.session_state.get('global_excel')

if uploaded_file is None:
    st.warning("⚠️ 請先在左側邊欄上傳「完整能源查核 Excel」檔案。")
else:
    try:
        # 讀取全部 sheet
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        target_list = []

        # 嚴格只找名稱包含「表九之二」的分頁
        for sheet_name, df_raw in all_sheets.items():
            if '表九之二' in sheet_name:
                
                # 自動尋找表頭 (通常在前 15 列內)
                header_idx = -1
                for idx, row in df_raw.head(15).iterrows():
                    row_str = "".join([str(val) for val in row.values if pd.notna(val)])
                    # 只要這行有「設備」、「種類」或「系統名稱」等關鍵字，就當作表頭
                    if '設備' in row_str or '種類' in row_str or '系統' in row_str:
                        header_idx = idx
                        break
                
                # 如果有找到表頭，將資料切出來
                if header_idx != -1:
                    df = df_raw.iloc[header_idx+1:].copy()
                    df.columns = df_raw.iloc[header_idx]
                    
                    # 移除全部是空值的爛資料
                    df = df.dropna(how='all').dropna(axis=1, how='all')
                    
                    # 標註來源分頁，方便你辨識是哪一棟建築物
                    df['來源建築物'] = sheet_name
                    
                    # 把這張表加入清單
                    target_list.append(df)
        
        # 將所有抓到的表九之二合併顯示
        if target_list:
            final_df = pd.concat(target_list, ignore_index=True)
            
            st.success(f"✅ 成功抓取！共找到 {len(final_df)} 筆「表九之二」的設備資料。")
            st.info("以下是原始提取資料，請確認是否正確：")
            
            # 直接把 DataFrame 顯示出來
            st.dataframe(final_df, use_container_width=True)
            
        else:
            st.info("找不到有效的「表九之二」資料，請確認 Excel 內容與格式。")

    except Exception as e:
        st.error(f"讀取失敗，錯誤訊息: {e}")
