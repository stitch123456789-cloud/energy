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
    # --- 3.5 自動計算平均電費 (可視化驗證版) ---
    avg_price_auto = 5.0
        if uploaded_global is not None:
            try:
                # 1. 讀取 Excel (先不設 skipfooter 看看數據長怎樣)
                df_52 = pd.read_excel(uploaded_global, sheet_name="表五之二")
                
                # 2. 💡 在網頁側邊欄顯示一下我抓到了什麼 (這行是驗證關鍵)
                st.sidebar.write("🔍 偵測到表五之二欄位：", list(df_52.columns))

                # 3. 自動辨識度數與金額欄位
                # 我們找包含 '度' 且包含 '年' 的欄位，以及包含 '元' 且包含 '年' 的欄位
                k_cols = [c for c in df_52.columns if '度' in str(c) and '年' in str(c)]
                f_cols = [c for c in df_52.columns if '元' in str(c) and '年' in str(c)]
                
                if k_cols and f_cols:
                    kwh_col = k_cols[0]
                    fee_col = f_cols[0]
                    
                    # 4. 清理數據 (移除逗號並轉數字)
                    total_kwh = pd.to_numeric(df_52[kwh_col].replace({',':''}, regex=True), errors='coerce').sum()
                    total_fee = pd.to_numeric(df_52[fee_col].replace({',':''}, regex=True), errors='coerce').sum()
                    
                    if total_kwh > 0:
                        avg_price_auto = round(total_fee / total_kwh, 2)
                        st.sidebar.info(f"✅ 自動計算成功：{avg_price_auto} 元/度")
                else:
                    st.sidebar.warning("⚠️ 找不到對應的『年用電量』或『年電費』欄位")
                    
            except Exception as e:
                st.sidebar.error(f"❌ 讀取表五之二失敗: {e}")

        # 存入 Session State
        st.session_state['auto_avg_price'] = avg_price_auto

# --- 4. 轉接器邏輯 (根據選單執行對應檔案) ---
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
