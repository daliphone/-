import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

# --- 1. 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v12.0", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #0B1C3F; color: #FFFFFF; }
    h1, h2, h3 { color: #FFD700 !important; }
    .stMarkdown p, label { color: #E0E0E0 !important; }
    .stButton>button { background-color: #F39C12; color: white; border-radius: 8px; font-weight: bold; width: 100%; border: none; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; width: 100%; border: none; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化範本數據 (沿用上一版) ---
if 'activity_list' not in st.session_state:
    st.session_state.activity_list = []

TEMPLATES = {
    "🐎 馬年慶：百倍奉還": {
        "name": "2026 馬年慶：百倍奉還抽獎活動",
        "purpose": "迎接 2026 農曆馬年（丙午年），結合春節紅包與「百倍奉還」話題；目的為帶動門市春節人潮及增加會員流量。",
        "core": "對象：所有門市消費者（每人限購 3 包）；範圍：馬尼行動通訊任一門市；產品：100 元新年禮包。",
        "schedule": "115/01/12-01/18 宣傳期\n115/01/19-02/08 販售期\n115/02/11 開獎日\n115/02/12-02/28 兌獎期",
        "prizes": "Sony PS5 | 1名 | 吸睛大獎\n現金 $6,666 | 1名 | 百倍奉還獎\nApple Watch | 2名 | 實用3C獎\n官網購物金 $1,500 | 115名 | 流量轉化獎",
        "sop": "確認限購數量；告知序號保存；限量 66 包管理；引導加入官方 LINE 蒐集個資。",
        "marketing": "FB/IG/脆倒數限動；弱勢分店數位包圍與區域廣告投遞；弱店試賣或加碼。",
        "risk": "稅務：> $1,000 需申報，> $20,000 扣 10%；序號需蓋章防偽；滯銷禮包調度機制。",
        "effect": "預估 2,000+ 人次進店；115 名中獎者帶動官網二次消費；建立長期會員名單。"
    }
}

# --- 3. 側邊欄與系統資訊 ---
with st.sidebar:
    st.title("系統資訊")
    st.info("v12.0 | Logo & 表格化輸出更新\n馬尼行銷規劃提案 © 2025 Money MKT")
    
    st.header("📋 快速範本")
    for t_name, t_data in TEMPLATES.items():
        if st.button(t_name):
            for key in t_data: st.session_state[f"p_{key}"] = t_data[key]
            st.rerun()

    if st.button("🗑️ 清空草稿"):
        for key in list(st.session_state.keys()):
            if key.startswith("p_"): st.session_state[key] = ""
        st.rerun()

# --- 4. 編輯區 ---
st.title("🐎 馬尼通訊 行銷企劃提案系統")
col_info1, col_info2 = st.columns(2)
with col_info1: proposer = st.text_input("提案人", key="p_proposer")
with col_info2: proposal_date = st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()
col_left, col_right = st.columns(2)

with col_left:
    p_name = st.text_input("一、 活動名稱", key="p_name")
    p_purpose = st.text_area("活動時機與目的", key="p_purpose", height=80)
    p_core = st.text_area("二、 活動核心內容", key="p_core", height=80)
    p_schedule = st.text_area("三、 活動時程安排 (建議一行一項)", key="p_schedule", height=100)
    # 提醒用戶使用分隔符號
    st.caption("💡 贈品格式建議：名稱 | 數量 | 備註 (使用 | 分隔可自動轉表格)")
    p_prizes = st.text_area("四、 贈品結構與預算", key="p_prizes", height=100)

with col_right:
    p_sop = st.text_area("五、 門市執行流程 (SOP)", key="p_sop", height=100)
    p_marketing = st.text_area("六、 行銷流程與策略", key="p_marketing", height=100)
    p_risk = st.text_area("七、 風險管理與注意事項", key="p_risk", height=100)
    p_effect = st.text_area("八、 預估成效", key="p_effect", height=100)

# --- 5. Word 核心產出邏輯 (含 Logo 與表格) ---
def generate_advanced_word():
    doc = Document()
    
    # A. 加入 Logo (若本地有 logo.png 則啟用，否則跳過)
    # 請確保腳本同層級有 logo.png，或置換路徑
    try:
        if os.path.exists("logo.png"):
            doc.add_picture("logo.png", width=Inches(1.5))
            last_p = doc.paragraphs[-1]
            last_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    except:
        pass

    # B. 標題與基礎資訊
    title = doc.add_heading('行銷企劃執行提案', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = Inches(6)
    info_table.cell(0,0).text = f"提案人: {st.session_state.get('p_proposer', '')}"
    info_table.cell(0,1).text = f"提案日期: {st.session_state.get('p_date', '')}"

    doc.add_heading(st.session_state.get('p_name', '未命名活動'), level=1)

    # C. 各章節處理
    sections = [
        ("一、 活動時機與目的", st.session_state.p_purpose),
        ("二、 活動核心內容", st.session_state.p_core),
        ("三、 活動時程安排", st.session_state.p_schedule),
        ("四、 贈品結構與預算", st.session_state.p_prizes),
        ("五、 門市執行流程", st.session_state.p_sop),
        ("六、 行銷流程與策略", st.session_state.p_marketing),
        ("七、 風險管理與注意事項", st.session_state.p_risk),
        ("八、 預估成效", st.session_state.p_effect)
    ]

    for title_text, content in sections:
        h = doc.add_heading(title_text, level=2)
        h.runs[0].font.color.rgb = RGBColor(184, 134, 11) # 金色標題
        
        # 特別處理第四點：贈品表格化
        if "贈品結構" in title_text and "|" in content:
            # 建立表格
            lines = [line for line in content.split('\n') if line.strip()]
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '贈品名稱'
            hdr_cells[1].text = '數量'
            hdr_cells[2].text = '備註/預算'
            
            for line in lines:
                parts = line.split('|')
                row_cells = table.add_row().cells
                for i in range(min(len(parts), 3)):
                    row_cells[i].text = parts[i].strip()
        
        # 特別處理第三點：時程清單化
        elif "時程安排" in title_text:
            for line in content.split('\n'):
                if line.strip(): doc.add_paragraph(line.strip(), style='List Bullet')
        
        else:
            doc.add_paragraph(content)

    doc.add_page_break()
    
    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

# --- 6. 執行按鈕 ---
st.divider()
if st.session_state.get('p_name'):
    if st.button("🔥 預覽並準備下載文件"):
        st.balloons()
        data = generate_advanced_word()
        st.download_button(
            label="📄 下載馬尼專用企劃書 (.docx)",
            data=data,
            file_name=f"Money_MKT_{p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
