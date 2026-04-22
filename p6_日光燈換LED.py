import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re

# --- 1. 格式鎖定工具 (複製你 P1 的標準) ---
def set_font_kai(run, size=12, is_bold=False, color=RGBColor(0, 0, 0)):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = color
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 核心邏輯 ---
st.title("💡 6. 照明系統 (LED汰換) 提案生成")

uploaded_file = st.session_state.get('global_excel')

if uploaded_file is None:
    st.warning("⚠️ 請先在左側邊欄上傳「完整能源查核 Excel」檔案。")
else:
    # 參數設定
    st.sidebar.header("⚙️ 照明參數設定")
    work_hours = st.sidebar.number_input("年照明運轉時數 (hr/年)", value=2500)
    invest_per_unit = st.sidebar.number_input("LED 換裝單價 (元/具)", value=800)
    
    # 電費連動
    val_from_app = st.session_state.get('auto_avg_price', 3.50)
    electricity_price = st.sidebar.number_input("平均電費 (元/度)", value=float(val_from_app))

    # A. 精準提取「表九之二」
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    target_lights_list = []

    for sheet_name, df_raw in all_sheets.items():
        # 模糊匹配名稱中包含「表九之二」的分頁
        clean_name = re.sub(r'\s+', '', sheet_name)
        if bool(re.search(r'表九之二|表9之2', clean_name)):
            
            # 定位表頭：尋找包含「燈具」或「種類」的列
            header_idx = -1
            for idx, row in df_raw.head(15).iterrows():
                row_str = "".join([str(v) for v in row.values if pd.notna(v)])
                if '種類' in row_str or '燈具' in row_str:
                    header_idx = idx
                    break
            
            if header_idx != -1:
                df = df_raw.iloc[header_idx+1:].copy()
                df.columns = df_raw.iloc[header_idx]
                
                # 自動對位欄位 (B欄-種類, C欄-數量, D欄-瓦數)
                type_col = next((c for c in df.columns if '種類' in str(c)), None)
                qty_col = next((c for c in df.columns if '數量' in str(c)), None)
                watt_col = next((c for c in df.columns if '瓦數' in str(c) or '容量' in str(c)), None)

                if type_col and qty_col:
                    df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
                    df[watt_col] = pd.to_numeric(df[watt_col], errors='coerce').fillna(0)
                    
                    # ✨ 關鍵過濾：只抓非 LED 的日光燈/螢光燈/T8/T5
                    mask = (df[type_col].str.contains('日光燈|螢光燈|T8|T5|T9', na=False)) & \
                           (~df[type_col].str.contains('LED', na=False, case=False))
                    
                    found = df[mask].copy()
                    if not found.empty:
                        found['來源建築物'] = sheet_name
                        # 統一名稱方便後續處理
                        found = found.rename(columns={type_col: '種類', qty_col: '數量', watt_col: '瓦數'})
                        target_lights_list.append(found)

    # B. 數據計算與展示
    if not target_lights_list:
        st.info("💡 系統掃描「表九之二」完畢，未發現需汰換之傳統日光燈。")
    else:
        full_df = pd.concat(target_lights_list, ignore_index=True)
        total_count = full_df['數量'].sum()
        total_old_kw = (full_df['瓦數'] * full_df['數量']).sum() / 1000
        
        # 節能估算 (換LED省55%)
        saved_kwh = total_old_kw * 0.55 * work_hours
        saved_money_wan = (saved_kwh * electricity_price) / 10000
        total_invest_wan = (total_count * invest_per_unit) / 10000
        payback = total_invest_wan / saved_money_wan if saved_money_wan > 0 else 0

        st.success(f"✅ 已從表九之二提取 {total_count:,.0f} 具待改善燈具。")
        st.dataframe(full_df[['來源建築物', '種類', '數量', '瓦數']], use_container_width=True)

        # C. 產出報告
        if st.button("🚀 生成報告並同步至打包中心", use_container_width=True):
            doc = Document()
            doc.add_heading('照明系統改善建議報告', 1)
            
            # 正文內容 (帶入紅字)
            p = doc.add_paragraph()
            set_font_kai(p.add_run("1. 經查核貴單位「表九之二」動力及照明清單，發現仍使用傳統燈具共計 "), size=12)
            set_font_kai(p.add_run(f"{total_count:,.0f} 具"), size=12, color=RGBColor(255, 0, 0))
            set_font_kai(p.add_run("，推估年節省電量為 "), size=12)
            set_font_kai(p.add_run(f"{saved_kwh:,.0f} kWh"), size=12, color=RGBColor(255, 0, 0))
            set_font_kai(p.add_run("，預計每年可省下電費 "), size=12)
            set_font_kai(p.add_run(f"{saved_money_wan:.2f} 萬元"), size=12, color=RGBColor(255, 0, 0))
            set_font_kai(p.add_run("。"), size=12)

            # 下載與入庫
            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.session_state['report_warehouse']["6. 照明改善報告"] = doc_io.getvalue()
            st.success("✅ 報告已入庫！")
