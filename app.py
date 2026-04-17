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
# --- 3.5 自動計算平均電費 (精準座標版) ---
st.sidebar.markdown("### 🔍 數據監控中心")

avg_price_auto = 5.0
current_file = st.session_state.get('global_excel')

if current_file is not None:
    try:
        all_sheets = pd.ExcelFile(current_file).sheet_names
        target_sheet = [s for s in all_sheets if "表五之二" in s]
        
        if target_sheet:
            # 💡 關鍵：我們讀取整張表，不設標題
            df_raw = pd.read_excel(current_file, sheet_name=target_sheet[0], header=None)
            
            # 根據你的 Excel 截圖：
            # 列號 22 (在 Python index 是 21)
            # L 欄是第 12 欄 (index 11) -> 合計度數
            # O 欄是第 15 欄 (index 14) -> 合計金額
            
            try:
                # 抓取 L22 和 O22
                val_kwh = df_raw.iloc[21, 11] # L 欄
                val_fee = df_raw.iloc[21, 14] # O 欄
                
                # 清理數據 (轉成字串 -> 刪除逗號 -> 轉數字)
                total_kwh = float(str(val_kwh).replace(',', '').strip())
                total_fee = float(str(val_fee).replace(',', '').strip())
                
                if total_kwh > 0:
                    avg_price_auto = round(total_fee / total_kwh, 2)
                    st.sidebar.metric("✅ 座標抓取成功", f"{avg_price_auto} 元/度")
                    st.sidebar.write(f"度數(L22): {total_kwh}")
                    st.sidebar.write(f"金額(O22): {total_fee}")
                else:
                    st.sidebar.error("❌ L22 數值為 0，無法計算")
            except Exception as coord_err:
                st.sidebar.error(f"❌ 座標抓取失敗: {coord_err}")
                st.sidebar.write("請確認合計是否在第 22 列，L 欄與 O 欄")
        else:
            st.sidebar.error("❌ 找不到『表五之二』分頁")
    except Exception as e:
        st.sidebar.error(f"❌ 系統錯誤: {e}")

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
