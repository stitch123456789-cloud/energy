import streamlit as st
import pandas as pd
import zipfile
import io

# --- 1. 網頁全域配置 ---
st.set_page_config(page_title="節能診斷工具箱", layout="wide")

# --- 2. 初始化暫存記憶體 (這部分最重要，不能漏) ---
if 'global_excel' not in st.session_state:
    st.session_state['global_excel'] = None

if 'report_warehouse' not in st.session_state:
    st.session_state['report_warehouse'] = {} # 存放所有產出的報告

# --- 3. 側邊欄：功能選單與全域上傳 ---
st.sidebar.title("🛠️ 節能診斷工具箱")

# (1) 全域 Excel 上傳區
st.sidebar.subheader("📂 全域資料庫 (全部工作表)")
uploaded_global = st.sidebar.file_uploader(
    "上傳完整能源查核 Excel", 
    type=["xlsx"], 
    key="global_excel"
)

if uploaded_global:
    st.sidebar.success("✅ 全域檔案已就緒")

st.sidebar.markdown("---")

# (2) 提案選單
mode = st.sidebar.radio(
    "請選擇分析項目：", 
    ["1. 變壓器效益分析", "2. 用戶基本資料"]
)

st.sidebar.markdown("---")

# (3) 報告輸出中心 (打包下載)
st.sidebar.subheader("📦 報告輸出中心")
if st.session_state['report_warehouse']:
    count = len(st.session_state['report_warehouse'])
    st.sidebar.write(f"目前已生成 {count} 份報告")
    
    # 建立 ZIP 壓縮包
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for name, data in st.session_state['report_warehouse'].items():
            zip_file.writestr(f"{name}.docx", data)
    
    st.sidebar.download_button(
        label="📥 一鍵打包全部下載 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="能源診斷全套報告.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    if st.sidebar.button("🗑️ 清空所有產出的報告"):
        st.session_state['report_warehouse'] = {}
        st.rerun()
else:
    st.sidebar.info("尚未生成任何報告")
# --- 3.5 自動計算平均電費 (視覺化除錯版) ---
st.sidebar.markdown("### 🔍 數據監控中心 (測試完後可刪除)")

avg_price_auto = 5.0
current_file = st.session_state.get('global_excel')

if current_file is not None:
    try:
        all_sheets = pd.ExcelFile(current_file).sheet_names
        target_sheet = [s for s in all_sheets if "表五之二" in s]
        
        if target_sheet:
            df_raw = pd.read_excel(current_file, sheet_name=target_sheet[0], header=None)
            
            # 尋找「合計」列
            total_row_index = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains('合計').any(), axis=1)].index
            
            if not total_row_index.empty:
                target_row = df_raw.iloc[total_row_index[0]]
                # 抓取該行所有數字
                numeric_values = []
                for val in target_row:
                    try:
                        c_val = float(str(val).replace(',', '').strip())
                        if c_val > 100: numeric_values.append(c_val)
                    except: continue
                
                if len(numeric_values) >= 2:
                    total_fee = numeric_values[-1]
                    total_kwh = numeric_values[-2]
                    avg_price_auto = round(total_fee / total_kwh, 2)
                    
                    # 💡 直接在側邊欄印出結果，讓你確認
                    st.sidebar.metric("計算出的電費單價", f"{avg_price_auto} 元/度")
                    st.sidebar.write(f"明細：{total_fee}元 / {total_kwh}度")
                else:
                    st.sidebar.error("❌ 找不到合計數字")
            else:
                st.sidebar.error("❌ 找不到『合計』文字")
        else:
            st.sidebar.error("❌ 找不到『表五之二』分頁")
    except Exception as e:
        st.sidebar.error(f"❌ 讀取出錯: {e}")
else:
    st.sidebar.info("💡 請先上傳 Excel 檔案")

# 強制寫入口袋
st.session_state['auto_avg_price'] = avg_price_auto

# --- 4. 轉接器邏輯 (同樣垂直對齊到最左邊或上一層) ---
if mode == "1. 變壓器效益分析":
    try:
        exec(open("p1_變壓器分析.py", encoding="utf-8").read())
    except FileNotFoundError:
        st.error("找不到 p1_變壓器分析.py")

elif mode == "2. 用戶基本資料":
    try:
        exec(open("p2_用戶簡介.py", encoding="utf-8").read())
    except FileNotFoundError:
        st.error("找不到 p2_用戶簡介.py")
