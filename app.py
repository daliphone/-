import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os

# --- 1. 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v13.0", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; }
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    section[data-testid="stSidebar"] { background-color: #0B1C3F; color: white; }
    section[data-testid="stSidebar"] .stMarkdown h2 { color: #FFD700 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化範本資料 (使用 Session State 確保可動態修改) ---
if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還": {
            "name": "2026 馬尼通訊「馬年慶：百倍奉還」",
            "purpose": "迎接馬年，透過 $100 低門檻吸引新舊客，增加會員與官網流量。",
            "core": "對象：全體消費者；範圍：全台門市；產品：「百倍奉還」禮包 ($100)。",
            "schedule": "01/12-01/18: 宣傳期\n01/19-02/08: 販售期\n02/11: 開獎日",
            "prizes": "Sony PS5 | 1 名 | 吸睛大獎\n現金 $6,666 | 1 名 | 百倍奉還獎",
            "sop": "1.限購3包。 2.引導加入LINE。",
            "marketing": "FB/IG/脆前導；區域廣告投遞。",
            "risk": "稅務申報；序號防偽；滯銷調度。",
            "effect": "預估 2,000+ 人次進店。"
        },
        "📱 範本：新機上市": {"name": "新品發表企劃", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""},
        "🎁 範本：品牌週年": {"name": "十週年盛典", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""},
        "🛍️ 範本：門市振興": {"name": "弱勢門市支援方案", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""}
    }

# --- 3. 側邊欄：範本管理與儲存 ---
with st.sidebar:
    st.header("📋 快速範本區")
    
    # 選擇要載入或儲存的範本目標
    selected_tpl_key = st.selectbox("選擇操作範本", options=list(st.session_state.templates_store.keys()))
    
    col_tpl1, col_tpl2 = st.columns(2)
    with col_tpl1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl_key]
            for key in data: st.session_state[f"p_{key}"] = data[key]
            st.rerun()
    
    with col_tpl2:
        if st.button("💾 儲存至此範本"):
            st.session_state.templates_store[selected_tpl_key] = {
                "name": st.session_state.get("p_name", ""),
                "purpose": st.session_state.get("p_purpose", ""),
                "core": st.session_state.get("p_core", ""),
                "schedule": st.session_state.get("p_schedule", ""),
                "prizes": st.session_state.get("p_prizes", ""),
                "sop": st.session_state.get("p_sop", ""),
                "marketing": st.session_state.get("p_marketing", ""),
                "risk": st.session_state.get("p_risk", ""),
                "effect": st.session_state.get("p_effect", "")
            }
            st.success(f"已更新：{selected_tpl_key}")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for key in list(st.session_state.keys()):
            if key.startswith("p_"): st.session_state[key] = ""
        st.rerun()

    with st.expander("🛠️ 系統資訊", expanded=False):
        st.caption("v13.0 | 修復字體 Bug & 範本雙向編輯\n馬尼行銷規劃提案 © 2025 Money MKT")

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
with c_top1: p_name = st.text_input("一、 活動名稱", key="p_name")
with c_top2: proposer = st.text_input("提案人", key="p_proposer", value="行銷部")
with c_top3: p_date = st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()
c1, c2 = st.columns(2)
with c1:
    st.text_area("活動時機與目的", key="p_purpose", height=100)
    st.text_area("二、 活動核心內容", key="p_core", height=100)
    st.text_area("三、 活動時程安排", key="p_schedule", height=120, help="格式 MM/DD: 內容")
    st.text_area("四、 贈品結構與預算", key="p_prizes", height=120, help="格式 品項 | 數量 | 備註")

with c2:
    st.text_area("五、 門市執行流程 (SOP)", key="p_sop", height=100)
    st.text_area("六、 行銷流程與策略", key="p_marketing", height=100)
    st.text_area("七、 風險管理與注意事項", key="p_risk", height=100)
    st.text_area("八、 預估成效", key="p_effect", height=100)

# --- 5. Word 輸出美化 (修正字體 Bug) ---
def set_msjh_font(run):
    """修正版：設定微軟正黑體"""
    run.font.name = 'Microsoft JhengHei'
    r = run._element
    rFonts = r.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        r.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

def generate_pro_word():
    doc = Document()
    
    # A. Logo
    if os.path.exists("logo.png"):
        doc.add_picture("logo.png", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # B. 標題
    h = doc.add_heading('行銷企劃執行提案書', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_info = info.add_run(f"提案人：{st.session_state.get('p_proposer')}  |  日期：{st.session_state.get('p_date')}")
    set_msjh_font(r_info)

    doc.add_heading(st.session_state.get('p_name', '未命名企劃'), level=1)

    sections = [
        ("一、 活動時機與目的", st.session_state.p_purpose),
        ("二、 活動核心內容", st.session_state.p_core),
        ("三、 活動時程安排 (Timeline)", st.session_state.p_schedule),
        ("四、 贈品結構與預算", st.session_state.p_prizes),
        ("五、 門市執行流程", st.session_state.p_sop),
        ("六、 行銷流程與策略", st.session_state.p_marketing),
        ("七、 風險管理與注意事項", st.session_state.p_risk),
        ("八、 預估成效", st.session_state.p_effect)
    ]

    for title_text, content in sections:
        h2 = doc.add_heading(title_text, level=2)
        h2.runs[0].font.color.rgb = RGBColor(11, 28, 63)
        
        # 時程表格化
        if "時程安排" in title_text and content:
            t = doc.add_table(rows=1, cols=2)
            t.style = 'Light Shading Accent 1'
            t.rows[0].cells[0].text = "階段/日期"
            t.rows[0].cells[1].text = "執行細節"
            for line in content.split('\n'):
                if line.strip():
                    parts = line.split(':') if ':' in line else [line, ""]
                    row = t.add_row().cells
                    row[0].text = parts[0].strip()
                    row[1].text = parts[1].strip() if len(parts)>1 else ""
        
        # 贈品表格化
        elif "贈品結構" in title_text and "|" in content:
            t = doc.add_table(rows=1, cols=3)
            t.style = 'Table Grid'
            hdr = t.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = "品項", "數量", "備註"
            for line in content.split('\n'):
                if "|" in line:
                    parts = line.split('|')
                    row = t.add_row().cells
                    for i in range(min(len(parts), 3)): row[i].text = parts[i].strip()
        else:
            p = doc.add_paragraph()
            r = p.add_run(content)
            set_msjh_font(r)

    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

# --- 6. 執行輸出 ---
st.divider()
if st.session_state.get('p_name'):
    if st.button("✅ 完成企劃並產生文檔"):
        doc_bytes = generate_pro_word()
        st.download_button(
            label="📥 下載馬尼行銷企劃書",
            data=doc_bytes,
            file_name=f"MoneyMKT_{p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
