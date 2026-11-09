import os
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
from io import BytesIO

# ==========================================
# 1. 核心解析器與輔助函數
# ==========================================
def safe_float(val, default=0.0):
    """ 強制捕捉數字與點，解決小數點遺失問題 """
    if val is None:
        return default
    s = str(val).replace('．', '.').replace('。', '.').replace(',', '.').strip()
    clean_s = "".join(re.findall(r'[0-9.]', s))
    try:
        return float(clean_s) if clean_s else default
    except: 
        return default

def to_m(mm_val):
    """ mm 轉 m """
    return safe_float(mm_val) / 1000.0

def set_cell_background(cell, fill_hex):
    """ 設定 Word 表格單格背景顏色 """
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

# ==========================================
# 2. 資料庫與全域配置
# ==========================================
MATERIAL_DB = {
    "柳安木": {"fb": 12000.0, "fv": 1000.0, "E": 1.1e7}, 
    "美松": {"fb": 13500.0, "fv": 1100.0, "E": 1.2e7}
}

STEEL_DB = {
    "C100x50x20x2":   {"I": 71.4, "Z": 14.28, "A": 4.54},
    "C100x50x20x2.3": {"I": 80.7, "Z": 16.14, "A": 5.17}
}

STUD_SIZES = ["45x90", "90x45", "45x45", "45x60"]
WALER_SIZES = ["45x90", "90x45", "45x45", "45x60", "C100x50x20x2", "C100x50x20x2.3"]

TIE_SPECS = {
    "#3 (9.5mm)": "9.0",
    "#4 (12.7mm)": "11.0",
    "#5 (15.9mm)": "16.0"
}

# ==========================================
# 3. Word 報告產出引擎 (Streamlit 專用內存轉譯版)
# ==========================================
def generate_word_report(beams_queue):
    doc = Document()
    
    # 設定頁邊距
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title = doc.add_heading("模板結構全系統檢核報告 (專業完整版)", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for data in beams_queue:
        wood = MATERIAL_DB[data["樹種"]]
        B, H, live = data["B"], data["H"], data["Live"]
        
        p_bot = 1.3 * (24.0 * H + 0.5) + live
        p_side = 24.0 * H 
        def_limit = int(data["撓度分母"])
        fy_kpa = safe_float(data["fy"], 2500.0) * 100.0

        h1 = doc.add_heading(f"樑編號：{data['ID']} ({int(B*1000)}x{int(H*1000)} mm)", level=1)
        h1.paragraph_format.space_before = Pt(18)
        
        p_info = doc.add_paragraph()
        p_info.add_run(f"【基本設計參數與材料性質】\n").bold = True
        p_info.add_run(f" 樹種木材: {data['樹種']} | 容許撓度分母: L/{def_limit}\n")
        p_info.add_run(f" 活載重 L: {live:.2f} kN/m²\n")
        p_info.add_run(f" 構件計算載重 -> 樑底壓力: {p_bot:.2f} kPa / 樑側壓力: {p_side:.2f} kPa\n")
        p_info.add_run(f" 混凝土設計強度 fc': {data['fc_prime']} kgf/cm²\n")
        p_info.add_run(f" 鋼材屈服強度 fy: {data['fy']} kgf/cm²\n")

        doc.add_heading("【一、構件抗彎與撓度檢核】", level=2)
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        titles = ['檢核構件(規格)', '彎曲應力(kPa)', '剪應力(kPa)', '撓度值(mm)', '力學模型與公式']
        col_widths = [Inches(2.0), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.7)]
        
        for i, title_text in enumerate(titles):
            hdr_cells[i].text = title_text
            hdr_cells[i].width = col_widths[i]
            set_cell_background(hdr_cells[i], "1F4E78")
            for paragraph in hdr_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True

        def add_row_to_doc(table_obj, name, p, spacing_m, span_m, spec_str, is_panel=False):
            is_steel = "c100" in str(spec_str).lower()
            try:
                if is_steel:
                    steel_prop = STEEL_DB.get(spec_str, STEEL_DB["C100x50x20x2"])
                    Z = steel_prop["Z"] * 1e-6  
                    I = steel_prop["I"] * 1e-8  
                    A = steel_prop["A"] * 1e-4  
                    allow_fb = 0.6 * fy_kpa
                    allow_fv = 0.4 * fy_kpa
                    E_modulus = 2.04e8  
                else:
                    clean = str(spec_str).lower().strip()
                    if "x" in clean:
                        wb_m, wh_m = [float(x)/1000.0 for x in clean.split('x')]
                    else: 
                        wb_m, wh_m = 1.0, float(clean)/1000.0
                    Z = (wb_m * wh_m**2 / 6.0)
                    I = (wb_m * wh_m**3 / 12.0)
                    A = (wb_m * wh_m)
                    allow_fb = wood['fb']
                    allow_fv = wood['fv']
                    E_modulus = wood['E']
            except:
                Z, I, A = 1e-5, 1e-5, 1e-5
                allow_fb, allow_fv, E_modulus = wood['fb'], wood['fv'], wood['E']

            if is_panel:
                cm, cd, formula = 8.0, 76.8, "簡支樑模型\nM=wL²/8\nΔ=5wL⁴/384EI"
            else:
                cm, cd, formula = 10.0, 145.0, "連續樑模型\nM=wL²/10\nΔ=wL⁴/145EI"
            
            w = p * spacing_m
            fb = (w * span_m**2 / cm) / Z if Z > 0 else 0
            fv = (1.5 * w * span_m * 0.6) / A if A > 0 else 0
            delta = (w * span_m**4) / (cd * E_modulus * I) if I > 0 else 0
            limit_delta = (span_m / def_limit)
            
            row_cells = table_obj.add_row().cells
            for idx, w_width in enumerate(col_widths):
                row_cells[idx].width = w_width
            
            row_cells[0].text = f"{name}\n規格: {spec_str}\n跨度 L={int(span_m*1000)}mm"
            
            fb_ratio = fb / allow_fb if allow_fb > 0 else 99
            fv_ratio = fv / allow_fv if allow_fv > 0 else 99
            delta_mm = delta * 1000
            limit_mm = limit_delta * 1000
            
            p1 = row_cells[1].paragraphs[0]
            run1 = p1.add_run(f"{fb:.1f}\nRatio: {fb_ratio:.2f}")
            if fb_ratio > 1.0: run1.font.color.rgb = RGBColor(200, 0, 0)
            
            p2 = row_cells[2].paragraphs[0]
            run2 = p2.add_run(f"{fv:.1f}\nRatio: {fv_ratio:.2f}")
            if fv_ratio > 1.0: run2.font.color.rgb = RGBColor(200, 0, 0)
            
            p3 = row_cells[3].paragraphs[0]
            run3 = p3.add_run(f"{delta_mm:.2f} mm\n(限 {limit_mm:.2f})")
            if delta_mm > limit_mm: run3.font.color.rgb = RGBColor(200, 0, 0)
            
            row_cells[4].text = formula
            
            for cell in row_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        add_row_to_doc(table, "樑底襯版(面板)", p_bot, 1.0, data["BotS_m"], data["BotT"], is_panel=True)
        add_row_to_doc(table, "樑底縱向格柵", p_bot, data["BotS_m"], data["BotWalerS_m"], data["BotStud"], is_panel=False)
        add_row_to_doc(table, "樑底橫向小樑", p_bot, data["BotWalerS_m"], B, data["BotWalerStud"], is_panel=False)
        
        add_row_to_doc(table, "樑側襯版(面板)", p_side, 1.0, data["SideS_m"], data["SideT"], is_panel=True)
        add_row_to_doc(table, "樑側縱向格柵", p_side, data["SideS_m"], data["SideStudS_m"], data["SideStud"], is_panel=False)
        add_row_to_doc(table, "樑側橫向垂直小樑", p_side, data["SideStudS_m"], data["TieV_m"], data["SideWalerStud"], is_panel=False)

        doc.add_paragraph()
        doc.add_heading("【二、穩定性支撐檢核 (Tie & Post)】", level=2)
        table2 = doc.add_table(rows=1, cols=4)
        table2.style = 'Table Grid'
        
        hdr2_cells = table2.rows[0].cells
        titles2 = ['檢核項目', '規格與佈設間距', '計算受力 (kN)', '安全檢核比值']
        col_widths2 = [Inches(1.8), Inches(2.0), Inches(1.4), Inches(1.8)]
        
        for i, title_text in enumerate(titles2):
            hdr2_cells[i].text = title_text
            hdr2_cells[i].width = col_widths2[i]
            set_cell_background(hdr2_cells[i], "2E75B6")
            for paragraph in hdr2_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True

        # 1. 樑側拉桿 (Tie)
        tie_f = p_side * data["TieH_m"] * data["TieV_m"]
        tie_ratio = tie_f / data["TieCap"] if data["TieCap"] > 0 else 99
        r1_cells = table2.add_row().cells
        r1_cells[0].text = "樑側拉桿\n(Form Tie)"
        r1_cells[1].text = f"型號 {data['TieDia']}\n間距 {int(data['TieH_m']*1000)}x{int(data['TieV_m']*1000)} mm"
        r1_cells[2].text = f"{tie_f:.2f} kN"
        p_tie = r1_cells[3].paragraphs[0]
        run_tie = p_tie.add_run(f"Ratio: {tie_ratio:.2f}\n(容許 {data['TieCap']} kN)")
        if tie_ratio > 1.0: run_tie.font.color.rgb = RGBColor(200, 0, 0)

        # 2. 樑底撐柱 (Post) - 寬度修正 B/2
        post_f = p_bot * (B / 2.0) * data["BotWalerS_m"]
        post_ratio = post_f / data["PostCap"] if data["PostCap"] > 0 else 99
        r2_cells = table2.add_row().cells
        r2_cells[0].text = "樑底撐柱\n(Post)"
        r2_cells[1].text = f"縱向間距 {int(data['BotWalerS_m'] * 1000)} mm\n受荷寬度(B/2) {int(B * 1000 / 2)} mm"
        r2_cells[2].text = f"{post_f:.2f} kN"
        p_post = r2_cells[3].paragraphs[0]
        run_post = p_post.add_run(f"Ratio: {post_ratio:.2f}\n(容許 {data['PostCap']} kN)")
        if post_ratio > 1.0: run_post.font.color.rgb = RGBColor(200, 0, 0)
        
        for row in table2.rows[1:]:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_page_break()

    # 1. (選填) 設定背景備份資料夾與固定檔名
    backup_path = r"D:\PythonTraining\beam_formwork_app\CMD_BeamFormwork_App"
    if not os.path.exists(backup_path):
        os.makedirs(backup_path)
    doc.save(os.path.join(backup_path, "模板全方位檢核報告_備份.docx"))

    # 2. 依然回傳給網頁下載按鈕使用
    target_stream = BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream
# ==========================================
# 4. Streamlit 網頁介面開發
# ==========================================
st.set_page_config(page_title="專業樑模板結構分析系統 v35.0", layout="wide")

st.title("🏗️ 營建工法型-樑模板全系統結構分析網頁端")
st.markdown("---")

# 初始化階段的 Session 暫存清單
if "BEAMS_QUEUE" not in st.session_state:
    st.session_state.BEAMS_QUEUE = []

# 利用 Layout 區分成左側（輸入區）與右側（待檢核清單預覽）
left_col, right_col = st.columns([7, 3])

with left_col:
    # 區塊一：基礎與材料強度參數設定
    st.subheader("🟢 基礎與材料強度參數設定")
    c1, c2, c3 = st.columns(3)
    with c1:
        proj_name = st.text_input("工程名稱", "新建工程")
        beam_id = st.text_input("樑編號", "B1")
    with c2:
        beam_b = st.number_input("樑寬 B (mm)", min_value=1, value=800)
        beam_h = st.number_input("樑高 H (mm)", min_value=1, value=1200)
    with c3:
        live_load = st.number_input("作業活載 (kN/m2)", min_value=0.0, value=2.5, step=0.1)
        fc_prime = st.selectbox("混凝土強度 fc' (kgf/cm²)", ["210", "245", "280", "315", "350"], index=2)

    c4, c5 = st.columns(2)
    with c4:
        fy_value = st.selectbox("鋼材屈服強度 fy (kgf/cm²)", ["2500", "2800", "4200"], index=0)
        wood_type = st.selectbox("木材樹種", ["柳安木", "美松"], index=0)
    with c5:
        def_limit_select = st.selectbox("容許撓度分母 (L/)", ["270", "360"], index=0)

    # 區塊二：樑底模板支撐系統
    st.subheader("🔵 樑底模板與多層支撐系統")
    c6, c7, c8 = st.columns(3)
    with c6:
        bot_t = st.number_input("襯板厚度 (mm)", min_value=1, value=18, key="bot_t")
        bot_s = st.number_input("格柵間距 (mm)", min_value=1, value=300, key="bot_s")
    with c7:
        bot_stud = st.selectbox("縱向格柵木料尺寸", STUD_SIZES, index=0, key="bot_stud")
        bot_waler_s = st.number_input("格柵跨距 (mm)", min_value=1, value=600, key="bot_waler_s")
    with c8:
        bot_waler_stud = st.selectbox("橫向支撐小樑尺寸", WALER_SIZES, index=4, key="bot_waler_stud")
        post_cap = st.number_input("單支鋼管立支柱容許強度 (kN)", min_value=0.1, value=10.0, step=0.5)

    # 區塊三：樑側模板與拉桿系統
    st.subheader("🟠 樑側模板與拉桿固定系統")
    c9, c10, c11 = st.columns(3)
    with c9:
        side_t = st.number_input("襯板厚度 (mm)", min_value=1, value=18, key="side_t")
        side_s = st.number_input("格柵間距 (mm)", min_value=1, value=300, key="side_s")
        tie_dia = st.selectbox("拉桿規格型號", list(TIE_SPECS.keys()), index=1)
    with c10:
        side_stud = st.selectbox("縱向格柵木料尺寸", STUD_SIZES, index=0, key="side_stud")
        side_stud_s = st.number_input("格柵跨距 (mm)", min_value=1, value=400, key="side_stud_s")
        tie_h = st.number_input("拉桿水平間距 (mm)", min_value=1, value=800)
    with c11:
        side_waler_stud = st.selectbox("橫向垂直小樑尺寸", STUD_SIZES, index=0)
        tie_v = st.number_input("拉桿垂直間距 (mm)", min_value=1, value=600)
        
        # 根據拉桿選單連動容許強度，並保留手動微調空間
        default_tie_cap = float(TIE_SPECS[tie_dia])
        tie_cap = st.number_input("拉桿容許抗拉強度 (kN)", min_value=0.1, value=default_tie_cap, step=1.0)

    # 按鈕控制區
    st.markdown("<br>", unsafe_allow_html=True)
    b_c1, b_c2 = st.columns(2)
    with b_c1:
        if st.button("➕ 加入排程清單", use_container_width=True):
            if not beam_id or not proj_name:
                st.error("工程名稱與樑編號不能為空值！")
            else:
                st.session_state.BEAMS_QUEUE.append({
                    "ID": beam_id, "Name": proj_name, "樹種": wood_type, "撓度分母": def_limit_select,
                    "fc_prime": fc_prime, "fy": fy_value,
                    "B": to_m(beam_b), "H": to_m(beam_h), "Live": live_load,
                    "BotT": bot_t, "BotS_m": to_m(bot_s), "BotStud": bot_stud,
                    "BotWalerS_m": to_m(bot_waler_s), "BotWalerStud": bot_waler_stud,
                    "PostCap": safe_float(post_cap, 10.0),
                    "SideT": side_t, "SideS_m": to_m(side_s), "SideStud": side_stud,
                    "SideStudS_m": to_m(side_stud_s), "SideWalerStud": side_waler_stud,
                    "TieDia": tie_dia, "TieH_m": to_m(tie_h), "TieV_m": to_m(tie_v), "TieCap": safe_float(tie_cap, 11.0)
                })
                st.success(f"已將樑 {beam_id} 成功加入排程清單！")
                st.rerun()

    with b_c2:
        if st.button("📭 快速清空清單", use_container_width=True):
            st.session_state.BEAMS_QUEUE.clear()
            st.warning("排程清單已清空。")
            st.rerun()

with right_col:
    st.subheader("📋 待檢核排程清單")
    if not st.session_state.BEAMS_QUEUE:
        st.info("目前清單內尚無資料。")
    else:
        # 在網頁右側直觀列出排程
        for idx, item in enumerate(st.session_state.BEAMS_QUEUE):
            st.markdown(f"**{idx+1}. 📍 樑編號: {item['ID']}**")
            st.caption(f"尺寸: {int(item['B']*1000)}x{int(item['H']*1000)}mm | 樑底小樑: {item['BotWalerStud']}")
            st.markdown("---")
            
        # 如果清單有資料，即時轉譯內存下載按鈕，使用者一鍵下載！
        st.markdown("<br>", unsafe_allow_html=True)
        word_data = generate_word_report(st.session_state.BEAMS_QUEUE)
        
        st.download_button(
            label="📝 產出並下載高質感 Word 報告",
            data=word_data,
            file_name="樑模板結構系統檢核報告_v35.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

if __name__ == "__main__":
    pass