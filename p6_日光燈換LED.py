"""
日光燈換LED節能報告產生器
輸入：Excel能源查核表
輸出：Word改善措施建議表(二)

使用方式：
    python generate_led_report.py <excel路徑> [輸出docx路徑]

模板檔案 template_5A03.docx 需與本程式放在同一目錄。
"""

import sys, os, re, shutil, zipfile, tempfile
import pandas as pd

# ── 固定參數（此報告類型的標準值）──
LED_SAVE_RATIO   = 0.464537   # LED換裝後節省比例（46.45%）
ELEC_PRICE       = 4.6238     # 平均電費（元/kWh）
INVEST_PER_KW    = 24677  # 元/kW（依模板反推）     # 投資費用（元/具）
OLD_LAMP_TYPE    = "1.日光燈"


# ══════════════════════════════════════
# 1. 讀取 Excel
# ══════════════════════════════════════

def read_excel(path):
    xl = pd.ExcelFile(path)

    # 讀取單位名稱
    df_info = pd.read_excel(path, sheet_name="三、能源用戶基本資料", header=None)
    unit_name = "貴單位"
    for _, row in df_info.iterrows():
        vals = [str(v) for v in row if str(v) not in ("nan", "None", "")]
        for i, v in enumerate(vals):
            if "07." in v and "能源用戶名稱" in v:
                if i + 1 < len(vals):
                    unit_name = vals[i + 1].strip()

    # 找照明系統工作表
    lighting_sheet = next(
        (s for s in xl.sheet_names if "表九之二" in s), None
    )
    if not lighting_sheet:
        raise ValueError("找不到「表九之二」照明系統工作表")

    df = pd.read_excel(path, sheet_name=lighting_sheet, header=None)

    # 欄位對照（依Excel實際欄位）：
    #  [1]=燈具種類 [2]=廠牌 [3]=裝設區域(有些欄位合併) [4]=燈管型式
    #  [5]=容量規格 [6]=安定器 [7]=電功率(瓦/具) [8]=製造年份
    #  [9]=數量(具) [10]=設備耗電合計(瓩) [11]=運轉時數(小時/年)
    lamps = []
    for _, row in df.iterrows():
        vals = list(row)
        first = str(vals[1]) if len(vals) > 1 else ""
        if not (first.startswith("1.") or first.startswith("2.")):
            continue

        def g(idx, typ=str):
            v = vals[idx] if idx < len(vals) else ""
            if str(v) in ("nan", "None", ""):
                return "" if typ == str else 0
            try:
                return typ(v)
            except:
                return "" if typ == str else 0

        lamps.append({
            "type":     g(1).strip(),
            "capacity": g(5).strip(),
            "qty":      g(9, int),      # 數量(具)
            "total_kw": g(10, float),   # 設備耗電合計(瓩=kW)
            "hours":    g(11, int),     # 運轉時數(hr/年)
        })

    return unit_name, lamps


# ══════════════════════════════════════
# 2. 計算節能效益
# ══════════════════════════════════════

def calculate(lamps):
    old = [l for l in lamps if l["type"] == OLD_LAMP_TYPE]
    if not old:
        raise ValueError("Excel中找不到日光燈（1.日光燈）資料")

    total_old_kwh = sum(l["total_kw"] * l["hours"] for l in old)
    save_kwh      = round(total_old_kwh * LED_SAVE_RATIO)
    save_money    = round(save_kwh * ELEC_PRICE / 10000, 2)   # 萬元/年
    energy_rate   = round(LED_SAVE_RATIO * 100, 2)             # %
    total_qty     = sum(l["qty"] for l in old)
    total_old_kw  = sum(l["total_kw"] for l in old)
    invest        = round(total_old_kw * INVEST_PER_KW / 10000, 1)  # 萬元
    payback       = round(invest / save_money, 1) if save_money > 0 else 0

    total_old_kw  = sum(l["total_kw"] for l in old)
    save_kw       = round(total_old_kw * LED_SAVE_RATIO)

    return {
        "old_lamps":      old,
        "total_old_kwh":  round(total_old_kwh),
        "save_kwh":       save_kwh,
        "save_money":     save_money,
        "energy_rate":    energy_rate,
        "save_kw":        save_kw,
        "invest":         invest,
        "payback":        payback,
    }


# ══════════════════════════════════════
# 3. 產生燈具資料列 XML
# ══════════════════════════════════════

def make_cell(text, width, left_border="nil", top_border="nil"):
    return f"""              <w:tc>
                <w:tcPr>
                  <w:tcW w:w="{width}" w:type="dxa"/>
                  <w:tcBorders>
                    <w:top w:val="{top_border}"/>
                    <w:left w:val="{left_border}" w:sz="4" w:space="0" w:color="auto"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                  </w:tcBorders>
                  <w:shd w:val="clear" w:color="auto" w:fill="auto"/>
                  <w:noWrap/>
                  <w:vAlign w:val="center"/>
                </w:tcPr>
                <w:p>
                  <w:pPr>
                    <w:widowControl/>
                    <w:jc w:val="center"/>
                    <w:rPr>
                      <w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:cs="新細明體"/>
                      <w:kern w:val="0"/>
                      <w:sz w:val="22"/>
                    </w:rPr>
                  </w:pPr>
                  <w:r>
                    <w:rPr>
                      <w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:cs="新細明體" w:hint="eastAsia"/>
                      <w:kern w:val="0"/>
                      <w:sz w:val="22"/>
                    </w:rPr>
                    <w:t>{text}</w:t>
                  </w:r>
                </w:p>
              </w:tc>"""


def make_lamp_row(lamp):
    qty   = f"{lamp['qty']:,}"   if lamp['qty']   else ""
    hours = f"{lamp['hours']:,}" if lamp['hours']  else ""
    return f"""            <w:tr>
              <w:trPr>
                <w:trHeight w:val="315"/>
                <w:jc w:val="center"/>
              </w:trPr>
{make_cell(lamp["type"],     1680, left_border="single")}
{make_cell(lamp["capacity"], 2216)}
{make_cell(qty,              1144)}
{make_cell(hours,            2100)}
            </w:tr>"""


# ══════════════════════════════════════
# 4. 修改 document.xml
# ══════════════════════════════════════

def patch_xml(xml_path, result):
    with open(xml_path, encoding="utf-8") as f:
        content = f.read()

    c = result
    fmt_kwh   = f"{c['total_old_kwh']:,}"
    fmt_save  = f"{c['save_kwh']:,}"

    # ── 頂部表格數字 ──
    content = content.replace(">298891<",  f">{c['total_old_kwh']}<")
    content = content.replace(">138846<",  f">{c['save_kwh']}<")
    content = content.replace(">64.20<",   f">{c['save_money']}<")
    content = content.replace(">46.45<",   f">{c['energy_rate']}<")
    content = content.replace(">168.3<",   f">{c['invest']}<")
    content = content.replace(">2.6<",     f">{c['payback']}<")

    # ── 現況耗電量合計（逗號格式）──
    content = content.replace(">298,891<", f">{fmt_kwh}<")

    # ── 預期效益段落 ──
    content = re.sub(
        r"降低尖峰用電需量約\d+kW",
        f"降低尖峰用電需量約{c['save_kw']}kW",
        content
    )
    content = re.sub(
        r"減少用電量約[\d,]+kWh/年",
        f"減少用電量約{fmt_save}kWh/年",
        content
    )
    # 節省金額（64.2 萬元）— 出現在預期效益段落
    content = re.sub(
        r"節省</w:t>.*?<w:t[^>]*>64\.2</w:t>",
        lambda m: m.group().replace(">64.2<", f">{c['save_money']}<"),
        content, flags=re.DOTALL
    )
    # 投資費用段落（168.3 × 2）
    # 回收年限段落（168.3 ÷ 64.2 = 2.6）
    content = re.sub(
        r"萬元÷</w:t>.*?<w:t[^>]*>64\.2</w:t>",
        lambda m: m.group().replace(">64.2<", f">{c['save_money']}<"),
        content, flags=re.DOTALL
    )
    content = re.sub(
        r"萬元/年=</w:t>.*?<w:t[^>]*>2\.6</w:t>",
        lambda m: m.group().replace(">2.6<", f">{c['payback']}<"),
        content, flags=re.DOTALL
    )

    # ── 燈具資料列替換 ──
    lamp_rows_xml = "\n".join(make_lamp_row(l) for l in c["old_lamps"])
    tr_pattern = re.compile(
        r'<w:tr\b[^>]*>(?:(?!</w:tr>)[\s\S])*?<w:t>日光燈</w:t>[\s\S]*?</w:tr>',
        re.DOTALL
    )
    matches = list(tr_pattern.finditer(content))
    if matches:
        content = content[:matches[0].start()] + lamp_rows_xml + content[matches[-1].end():]

    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(content)


# ══════════════════════════════════════
# 5. 打包 DOCX
# ══════════════════════════════════════

def build_docx(template_path, output_path, result):
    tmpdir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(template_path, "r") as z:
            z.extractall(tmpdir)
        patch_xml(os.path.join(tmpdir, "word", "document.xml"), result)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    fpath = os.path.join(root, file)
                    zout.write(fpath, os.path.relpath(fpath, tmpdir))
    finally:
        shutil.rmtree(tmpdir)


# ══════════════════════════════════════
# 主程式
# ══════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    excel_path = sys.argv[1]
    if not os.path.exists(excel_path):
        print(f"錯誤：找不到檔案 {excel_path}")
        sys.exit(1)

    output_path = sys.argv[2] if len(sys.argv) >= 3 else \
        f"LED節能報告_{os.path.splitext(os.path.basename(excel_path))[0]}.docx"

    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_5A03.docx")
    if not os.path.exists(template_path):
        print(f"錯誤：找不到模板 {template_path}")
        sys.exit(1)

    print(f"讀取：{excel_path}")
    unit_name, lamps = read_excel(excel_path)
    result = calculate(lamps)

    c = result
    print(f"  單位：{unit_name}")
    print(f"  日光燈資料筆數：{len(c['old_lamps'])}")
    print(f"  現況耗電：{c['total_old_kwh']:,} kWh/年")
    print(f"  節省耗電：{c['save_kwh']:,} kWh/年（{c['energy_rate']}%）")
    print(f"  節省金額：{c['save_money']} 萬元/年")
    print(f"  降低需量：{c['save_kw']} kW")
    print(f"  投資費用：{c['invest']} 萬元")
    print(f"  回收年限：{c['payback']} 年")

    print(f"產生：{output_path}")
    build_docx(template_path, output_path, result)
    print("完成！")

if __name__ == "__main__":
    main()
