import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 通用工具函數 ---
def set_font_kai(run, size=14, is_bold=False, color=RGBColor(0, 0, 0)):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = color
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 數據抓取邏輯 (整合多電號與座標抓取) ---
def fetch_exact_data():
    info = {"comp": "未抓到名稱", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": "115年1月1日"}
    elec_list = [] 
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- A. 抓基本資料 (sheet_b) ---
            sheet_b = next((s for s in xl.sheet_names if "三" in s or "基本資料" in s), None)
            if sheet_b:
                df_b = pd.read_excel(file, sheet_name=sheet_b, header=None)
                
                def get_near_value(items, keyword, min_val=0):
                    import re
                    for i, item in enumerate(items):
                        if keyword in str(item):
                            for target in items[i+1 : i+5]:
                                if target is None or str(target).lower() == "nan": continue
                                clean = str(target).replace(",", "").replace(" ", "")
                                matches = re.findall(r"[-+]?\d*\.\d+|\d+", clean)
                                if matches:
                                    try:
                                        num = int(round(float(matches[0])))
                                        if num > min_val: return f"{num:,d}"
                                    except: continue
                    return None

                for r in range(len(df_b)):
                    row_list = list(df_b.iloc[r, :])
                    row_str = "".join([str(i) for i in row_list])
                    if "員工人數" in row_str:
                        res = get_near_value(row_list, "員工人數")
                        if res: info["emp"] = res
                    if "全年工作時數" in row_str:
                        res = get_near_value(row_list, "全年工作時數")
                        if res: info["hours"] = res
                    if "總樓地板面積" in row_str:
                        res = get_near_value(row_list, "總樓地板面積", min_val=100)
                        if res: info["area"] = res
                    if "總空調使用面積" in row_str:
                        res = get_near_value(row_list, "總空調使用面積", min_val=100)
                        if res: info["air_area"] = res

            # --- B. 遍歷所有五之二 (電力數據) ---
            p_sheets = [s for s in xl.sheet_names if "五之二" in s]
            for i, s_name in enumerate(p_sheets):
                df_p = pd.read_excel(file, sheet_name=s_name, header=None)
                
                # 第一個電號分頁抓取公司名稱
                if i == 0:
                    try:
                        name_val = str(df_p.iloc[5, 4]).strip()
                        if name_val != "nan": info["comp"] = name_val.split('(')[0].split('（')[0]
                    except: pass

                e_data = {
                    "elec_id": str(df_p.iloc[5, 2]).strip(),
                    "contract_cap": "0", "total_kwh": "0", "total_fee": "0", 
                    "avg_pf": "0", "avg_price": "0", "volt": "22.8",
                    "trans_cap": "0", "cap_cap": "0", "peak_max": "0", "offpeak_max": "0"
                }

                # 精確座標抓取
                try:
                    e_data["contract_cap"] = str(int(float(df_p.iloc[9, 2])))
                    kwh = float(df_p.iloc[21, 11])
                    e_data["total_kwh"] = f"{int(kwh):,d}"
                    fee = float(df_p.iloc[21, 14])
                    e_data["total_fee"] = f"{int(fee):,d}"
                    if kwh > 0: e_data["avg_price"] = str(round(fee / kwh, 2))
                    e_data["avg_pf"] = str(int(float(df_p.iloc[22, 13])))
                    
                    # 需量最大值
                    p_vals = [float(df_p.iloc[r, 3]) for r in range(9, 21) if pd.notnull(df_p.iloc[r, 3]) and str(df_p.iloc[r, 3]).strip() not in ["-", "0"]]
                    if p_vals: e_data["peak_max"] = str(int(max(p_vals)))
                    o_vals = [float(df_p.iloc[r, 6]) for r in range(9, 21) if pd.notnull(df_p.iloc[r, 6]) and str(df_p.iloc[r, 6]).strip() not in ["-", "0"]]
                    if o_vals: e_data["offpeak_max"] = str(int(max(o_vals)))
                except: pass

                # 第一個電號處理「表八」設備加總
                if i == 0:
                    sheet_8 = next((s for s in xl.sheet_names if "八" in s), None)
                    if sheet_8:
                        df_8 = pd.read_excel(file, sheet_name=sheet_8, header=None)
                        t_sum = sum([float(df_8.iloc[7, col]) for col in range(5, len(df_8.columns)) if pd.notnull(df_8.iloc[7, col]) and isinstance(df_8.iloc[7, col], (int,float))])
                        e_data["trans_cap"] = f"{int(t_sum):,d}"
                        c_sum = sum([float(df_8.iloc[22, col]) for col in range(5, len(df_8.columns)) if pd.notnull(df_8.iloc[22, col]) and isinstance(df_8.iloc[22, col], (int,float))])
                        e_data["cap_cap"] = str(int(c_sum))

                elec_list.append(e_data)
        except Exception as e:
            st.error(f"解析發生錯誤: {e}")
            
    return info, elec_list

# --- 3. 介面 ---
st.title("📋 節能診斷自動化工具")
info_result, elec_systems = fetch_exact_data()

with st.expander("🔍 檢視/微調自動抓取資料"):
    ec1, ec2 = st.columns(2)
    with ec1:
        v_comp = st.text_input("用戶名稱", info_result["comp"])
        v_area = st.text_input("總面積 (m2)", info_result["area"])
        v_air = st.text_input("空調面積 (m2)", info_result["air_area"])
    with ec2:
        v_emp = st.text_input("員工人數", info_result["emp"])
        v_hours = st.text_input("工作時數 (hr/y)", info_result["hours"])
        v_date = st.text_input("📅 診斷日期", info_result["date"])

st.markdown("### ⚡ 電力系統設備資料")
if elec_systems:
    tabs = st.tabs([f"電號 {e['elec_id']}" for e in elec_systems])
    for i, tab in enumerate(tabs):
        with tab:
            c1, c2, c3 = st.columns(3)
            elec_systems[i]['elec_id'] = c1.text_input("台電電號", elec_systems[i]['elec_id'], key=f"id_{i}")
            elec_systems[i]['total_kwh'] = c1.text_input("年總用電度", elec_systems[i]['total_kwh'], key=f"kwh_{i}")
            elec_systems[i]['total_fee'] = c1.text_input("年總金額", elec_systems[i]['total_fee'], key=f"fee_{i}")
            
            elec_systems[i]['contract_cap'] = c2.text_input("契約容量 (kW)", elec_systems[i]['contract_cap'], key=f"cap_{i}")
            elec_systems[i]['trans_cap'] = c2.text_input("主變壓器容量 (kVA)", elec_systems[i]['trans_cap'], key=f"trans_{i}")
            elec_systems[i]['cap_cap'] = c2.text_input("電容器容量 (kVAR)", elec_systems[i]['cap_cap'], key=f"ccap_{i}")
            
            elec_systems[i]['avg_pf'] = c3.text_input("平均功因 (%)", elec_systems[i]['avg_pf'], key=f"pf_{i}")
            elec_systems[i]['peak_max'] = c3.text_input("尖峰最高需量", elec_systems[i]['peak_max'], key=f"p_max_{i}")
            elec_systems[i]['offpeak_max'] = c3.text_input("離峰最高需量", elec_systems[i]['offpeak_max'], key=f"o_max_{i}")

# --- 4. 封裝 Word 生成邏輯 ---
def generate_docx(comp, area, air, emp, hours, date, elecs):
    doc = Document()
    p_t1 = doc.add_paragraph(); set_font_kai(p_t1.add_run("二、能源用戶概述"), is_bold=True)
    p_t2 = doc.add_paragraph(); set_font_kai(p_t2.add_run("  2-1. 用戶簡介"), is_bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(28)
    set_font_kai(p.add_run(comp), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("總建物面積"))
    set_font_kai(p.add_run(area), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("平方公尺，空調使用面積"))
    set_font_kai(p.add_run(air), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("平方公尺，能源使用主要以電力為主，員工約有"))
    set_font_kai(p.add_run(emp), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("人，全年使用時間約"))
    set_font_kai(p.add_run(hours), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("小時，"))
    set_font_kai(p.add_run(date), color=RGBColor(255, 0, 0)) 
    set_font_kai(p.add_run("經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

    for i, e in enumerate(elecs):
        doc.add_paragraph()
        set_font_kai(doc.add_paragraph().add_run(f"1.{i+1} 電力系統 (電號：{e['elec_id']})："), is_bold=True)
        table = doc.add_table(rows=5, cols=3)
        table.style = 'Table Grid'
        
        # 電號行
        cell_id = table.cell(0, 0); cell_id.merge(table.cell(0, 2))
        set_font_kai(cell_id.paragraphs[0].add_run(f"台電電號：{e['elec_id']}"), size=12, color=RGBColor(255, 0, 0))

        # 內容行
        r1 = table.rows[1].cells
        set_font_kai(r1[0].paragraphs[0].add_run(f"契約型式：高壓 3 段式"), size=12)
        set_font_kai(r1[1].paragraphs[0].add_run(f"契約容量：{e['contract_cap']} [kW]"), size=12)
        set_font_kai(r1[2].paragraphs[0].add_run(f"台電供電電壓：22.8 [kV]"), size=12)

        r2 = table.rows[2].cells
        set_font_kai(r2[0].paragraphs[0].add_run(f"主變壓器總裝置容量：{e['trans_cap']} [kVA]"), size=12)
        set_font_kai(r2[1].paragraphs[0].add_run(f"電容器裝置容量：{e['cap_cap']} [kVAR]"), size=12)
        set_font_kai(r2[2].paragraphs[0].add_run(f"低壓側電壓：380/220 [V]"), size=12)

        r3 = table.rows[3].cells
        set_font_kai(r3[0].paragraphs[0].add_run(f"年總用電度：{e['total_kwh']} [kWh]"), size=12)
        set_font_kai(r3[1].paragraphs[0].add_run(f"年總金額：{e['total_fee']} [元]"), size=12)
        set_font_kai(r3[2].paragraphs[0].add_run(f"平均單價：{e.get('avg_price', '0')} [元/kWh]"), size=12)

        r4 = table.rows[4].cells
        set_font_kai(r4[0].paragraphs[0].add_run(f"平均功因：{e['avg_pf']} [%]"), size=12)
        set_font_kai(r4[1].paragraphs[0].add_run(f"尖峰最高需量：{e['peak_max']} [kW]"), size=12)
        set_font_kai(r4[2].paragraphs[0].add_run(f"離峰最高需量：{e['offpeak_max']} [kW]"), size=12)

        for row in table.rows:
            for cell in row.cells: cell.vertical_alignment = 1

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# --- 5. 下載按鈕 ---
st.markdown("---")
if st.download_button(
    label="💾 生成並下載用戶簡介 Word",
    data=generate_docx(v_comp, v_area, v_air, v_emp, v_hours, v_date, elec_systems),
    file_name=f"能源用戶簡介_{v_comp}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
):
    st.success("報告生成成功！")
