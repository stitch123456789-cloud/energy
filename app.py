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
# --- 3.5 自動計算平均電費 (模糊搜尋分頁版) ---
    avg_price_auto = 5.0
    if uploaded_global is not None:
            try:
                # 1. 先抓出這份 Excel 所有的分頁名稱
                all_sheets = pd.ExcelFile(uploaded_global).sheet_names
                
                # 2. 尋找名字裡包含 "表五之二" 的分頁
                target_sheet = [s for s in all_sheets if "表五之二" in s]
                
                if target_sheet:
                    # 抓到第一個符合的分頁
                    sheet_to_read = target_sheet[0]
                    # skipfooter=1 是為了避開最後一列的「平均」，我們只需要「合計」
                    df_52 = pd.read_excel(uploaded_global, sheet_name=sheet_to_read, skipfooter=1)
                    
                    # 3. 自動辨識度數與金額欄位
                    # 根據你的截圖，欄位名稱應該是 "合計" (度數) 和 "總電費(含稅)(元)"
                    k_cols = [c for c in df_52.columns if '合計' in str(c)]
                    f_cols = [c for c in df_52.columns if '總電費' in str(c)]
                    
                    if k_cols and f_cols:
                        # 抓取最後一列 (合計列) 的數據
                        # 使用 pd.to_numeric 確保避開文字干擾
                        total_kwh = pd.to_numeric(df_52[k_cols[0]].replace({',':''}, regex=True), errors='coerce').iloc[-1]
                        total_fee = pd.to_numeric(df_52[f_cols[0]].replace({',':''}, regex=True), errors='coerce').iloc[-1]
                        
                        if total_kwh > 0:
                            avg_price_auto = round(total_fee / total_kwh, 2)
                            st.sidebar.info(f"✅ 已由 {sheet_to_read} 計算單價：{avg_price_auto}")
                else:
                    st.sidebar.warning("⚠️ Excel 中找不到包含『表五之二』字樣的分頁")
                    
            except Exception as e:
                st.sidebar.error(f"❌ 讀取失敗: {e}")

        # 存入 Session State
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
