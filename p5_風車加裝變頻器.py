import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# --- 1. 介面設定 ---
st.title("🌀 P5. 冷卻水塔風車變頻分析")

c1, c2, c3 = st.columns(3)
with c1:
    unit_name = st.text_input("單位名稱", value="貴單位")
    ch_info = st.text_input("主機編號", value="CH-1")
with c2:
    motor_hp = st.number_input("單台馬力 (HP)", value=50.0)
    elec_val = st.number_input("平均電費 (元/度)", value=3.5)
with c3:
    rt_info = st.text_input("容量 (RT)", value="1500RT")
    invest_amt = st.number_input("投資金額 (萬元)", value=80.0)
    setup_note = st.text_input("運轉說明", value="僅開啟 2 台")

if "p5_data" not in st.session_state:
    st.session_state.p5_data = pd.DataFrame({
        "季節": ["春秋季", "夏季", "冬季"],
        "時數(hr)": [4380, 2190, 2190],
        "負載率(%)": [70, 85, 60]
    })
current_df = st.data_editor(st.session_state.p5_data, use_container_width=True)

# --- 2. 生成邏輯 ---
if st.button("🚀 生成報告 (文字與表格分段處理)", use_container_width=True):
    try:
        # A. 計算數據
        base_kw = motor_hp * 0.746
        total_old, total_new = 0, 0
        details = []
        for _, r in current_df.iterrows():
            h, l = float(r["時數(hr)"]), float(r["負載率(%)"])/100
            o, n = base_kw * h, base_kw * (l**3) * 1.06 * h
            details.append({"季節": r["季節"], "時數": h, "負載": f"{r['負載率(%)']}%", "舊": o, "新": n, "省": o-n})
            total_old += o
            total_new += n

        # B. 準備替換地圖
        data_map = {
            "{{UN}}": unit_name, "{{COUNT}}": "2", "{{CH_INFO}}": ch_info,
            "{{RT_INFO}}": rt_info, "{{MT}}": f"{int(motor_hp)}hp",
            "{{ON}}": setup_note, "{{OLD_KWH}}": f"{total_old:,.0f}",
            "{{SAVE_KWH}}": f"{(total_old-total_new):,.0f}", "{{MOTOR_SPEC}}": f"{int(motor_hp)}HPx3台",
            "{{SAVE_MONEY}}": f"{((total_old-total_new)*elec_val/10000):.2f}",
            "{{INVEST}}": f"{invest_amt:.1f}", "{{PAYBACK}}": f"{(invest_amt/((total_old-total_new)*elec_val/10000)):.1f}",
            "{{SUPPRESS_KW}}": "13"
        }

        doc = Document("template_p5.docx")

        # C. 第一步：強制替換所有段落中的標籤 (文字優先)
        for p in doc.paragraphs:
            combined_text = "".join(run.text for run in p.runs)
            for k, v in data_map.items():
                if k in combined_text:
                    # 暴力重寫段落，確保標籤消失
                    new_text = combined_text.replace(k, str(v))
                    p.text = "" # 清空
                    run = p.add_run(new_text)
                    run.font.name = '標楷體'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    combined_text = new_text

        # D. 第二步：在對應位置插入表格
        for p in doc.paragraphs:
            if "[[OLD_TABLE]]" in p.text:
                p.text = "" # 移除標籤文字
                # 建立現況表格
                tbl = doc.add_table(rows=1, cols=4)
                tbl.style = 'Table Grid' # 這裡如果報錯請改為 tbl.style = None
                hdr = ["季節", "時數(hr)", "負載(%)", "耗電(kWh)"]
                for i, text in enumerate(hdr):
                    tbl.cell(0, i).text = text
                for d in details:
                    row = tbl.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text = d['季節'], f"{d['時數']:,.0f}", "100%", f"{d['舊']:,.0f}"
                # 合計
                tot = tbl.add_row().cells
                tot[0].text, tot[3].text = "合計", f"{total_old:,.0f}"

            if "[[NEW_TABLE]]" in p.text:
                p.text = "" # 移除標籤文字
                tbl = doc.add_table(rows=1, cols=5)
                tbl.style = 'Table Grid'
                hdr = ["季節", "時數(hr)", "負載(%)", "預期耗電", "節電量"]
                for i, text in enumerate(hdr):
                    tbl.cell(0, i).text = text
                for d in details:
                    row = tbl.add_row().cells
                    row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = d['季節'], f"{d['時數']:,.0f}", d['負載'], f"{d['新']:,.0f}", f"{d['省']:,.0f}"
                # 合計
                tot = tbl.add_row().cells
                tot[0].text, tot[4].text = "合計", f"{(total_old-total_new):,.0f}"

        # E. 處理表格內標籤 (例如看板表格)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for k, v in data_map.items():
                        if k in cell.text:
                            cell.text = cell.text.replace(k, str(v))

        buf = io.BytesIO()
        doc.save(buf)
        st.success("✅ 生成完畢！文字已替換，表格已定位。")
        st.download_button("📥 下載修正版報告", buf.getvalue(), "風車分析報告.docx")

    except Exception as e:
        st.error(f"❌ 出錯: {e}")
