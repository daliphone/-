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
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.3.2", page_icon="🐎", layout="wide")

# CSS 強制美化側邊欄與 UI
st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    ::placeholder { color: #888888 !important; opacity: 0.7 !important; }
    
    /* 左側側邊欄美化 */
    [data-testid="stSidebar"] { 
        background-color: #003f7e !important; 
        border-right: 2px solid #ef8200;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { 
        color: #ef8200 !important; 
        font-weight: bold;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
        font-size: 16px;
    }
    
    /* 下拉選單 (Selectbox) 美化 - 解決白底白字問題 */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #003f7e !important;
        border-radius: 5px;
    }
    div[data-testid="stSelectbox"] label {
        color: #ef8200 !important;
        margin-bottom: 5px;
    }
    
    /* 按鈕樣式 */
    .stButton>button { border-radius: 8px; font-weight: bold; width: 100%; transition: 0.3s; }
    .ai-btn>div>button { 
        background-color: #6200EA !important; 
        color: white !important; 
        border: 1px solid #ef8200 !important; 
    }
    .ai-btn>div>button:hover { background-color: #3700B3 !important; transform: scale(1.02); }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 分章節 AI 優化邏輯 ---
def section_ai_logic(field_id, text):
    if not text or len(text) < 2: return text
    if field_id == "p_purpose":
        return f"【營運目的優化】本活動核心在於{text}。透過精準時機切入與誘因設計，旨在提升客流並強化品牌高性價比形象。"
    elif field_id == "p_core":
        return f"【核心內容優化】本活動名稱為「{st.session_state.p_name}」，鎖定目標族群需求，透過差異化服務建立優勢。"
    elif field_id == "p_schedule":
        return f"{text}\n\n💡 AI 執行重點：請特別注意宣傳期銜接，確保人員在活動開始前完成所有佈置。"
    elif field_id == "p_prizes":
        return f"{text}\n\n💡 AI 配置建議：大獎創造話題，小額購物金驅動官網二次消費。"
    elif field_id == "p_sop":
        return f"{text}\n\n💡 AI SOP 建議：強調『卸下武裝』話術，先聊需求不推產品，嚴格執行限量管理。"
    elif field_id == "p_marketing":
        return f"🚀【整合行銷】{text}。建議同步佈署區域廣告與官方帳號通知。"
    elif field_id == "p_risk":
        return f"{text}\n\n💡 AI 風險提示：務必注意稅務申報門檻(>$1000)與防偽核對流程。"
    elif field_id == "p_effect":
        return f"【預期效益優化】{text}。預計可累積大量潛在客戶名單作為未來行銷受眾。"
    return text

# --- 3. 初始化 Session State ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
for field in FIELDS:
    if field not in st.session_state:
        st.session_state[field] = ""

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還 (官方範本)": {
            "p_name": "2026 馬尼通訊「馬年慶：百倍奉還」企劃案",
            "p_purpose": "迎接馬年，透過 $100 元低門檻吸引新舊客戶進店，增加官網流量。",
            "p_core": "對象：全門市消費者；核心產品：$100 新年禮包。",
            "p_schedule": "115/01/12 宣傳、01/19 販售。",
            "p_prizes": "PS5 | 1名 | 售價 $100 包裝。",
            "p_sop": "確認限購3包、引導加官方 LINE。",
            "p_marketing": "FB/IG 限動倒數、門市完售海報。",
            "p_risk": "稅金申報規範、序號防偽處理。",
            "p_effect": "預期 2,000+ 人流、官網互動提升。"
        }
    }

# --- 4. 側邊欄設計 ---
with st.sidebar:
    st.header("📋 企劃範本庫")
    # 下拉選單顯示美化
    selected_tpl = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 載入範本"):
            for k, v in st.session_state.templates_store[selected_tpl].items():
                st.session_state[k] = v
            st.rerun()
    with c2:
        if st.button("💾 儲存範本"):
            name_snip = st.session_state.p_name[:5] if st.session_state.p_name else datetime.now().strftime('%H%M')
            st.session_state.templates_store[f"💾 自訂：{name_snip}..."] = {f: st.session_state[f] for f in FIELDS}
            st.success("已存入庫")

    st.divider()
    if st.button("🗑️ 清空所有草稿"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊 v14.3.2", expanded=False):
        st.caption("馬尼門活動企劃系統 © 2025 Money MKT")
        st.caption("修正：AI 寫入衝突解決、左欄高辨識度樣式")

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

t1, t2, t3 = st.columns([2, 1, 1])
with t1:
    # 使用 value=st.session_state[...] 並加上 key 以維持同步，但不直接寫入 key
    p_name = st.text_input("一、 活動名稱", value=st.session_state.p_name, placeholder="例如: 馬年慶：百倍奉還")
    st.session_state.p_name = p_name
with t2:
    p_proposer = st.text_input("提案人", value=st.session_state.p_proposer)
    st.session_state.p_proposer = p_proposer
with t3:
    p_date = st.date_input("提案日期", value=datetime.now())

st.divider()

sections = [
    ("p_purpose", "活動時機與目的", "營運目的邏輯建議", "迎接馬年話題，解決連假後人流痛點。"),
    ("p_core", "二、 活動核心內容", "核心賣點配置建議", "產品具備衝動購買力($100)，適合快速成交。"),
    ("p_schedule", "三、 活動時程安排", "執行重點建議", "宣傳期需於除夕前完成，開獎設定於開工後引流回訪。"),
    ("p_prizes", "四、 贈品結構與預算", "商品配置用意建議", "PS5 創造話題，購物金強制客戶登入官網產生二次消費。"),
    ("p_sop", "五、 門市執行 SOP", "執行環節注意事項", "務必強調『序號正本』為兌獎唯一憑證，先卸下武裝不推產品。"),
    ("p_marketing", "六、 行銷流程與策略", "建議管道與潤稿", "利用紅包色視覺，社群任務可設計分享好運抽購物金。"),
    ("p_risk", "七、 風險管理與注意事項", "規範與注意建議", "每店配額管理避免跨區落空，務必收齊身分證影本報稅。"),
    ("p_effect", "八、 預估成效", "效益面建議", "重點指標：門市進店率、官網註冊數、二次轉化率。")
]

col_a, col_b = st.columns(2)
for i, (fid, title, tip_title, tip_content) in enumerate(sections):
    target_col = col_a if i < 4 else col_b
    with target_col:
        # 使用 value 綁定 session_state，避免直接賦值給 key 導致錯誤
        user_input = st.text_area(title, value=st.session_state[fid], height=120, key=f"input_{fid}")
        st.session_state[fid] = user_input
        
        st.markdown('<div class="ai-btn">', unsafe_allow_html=True)
        if st.button(f"🪄 執行 {title} AI 優化", key=f"btn_{fid}"):
            optimized = section_ai_logic(fid, st.session_state[fid])
            st.session_state[fid] = optimized
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander(f"💡 {tip_title} (馬年慶背景)", expanded=False):
            st.write(tip_content)
        st.write("")

# --- 6. Word 下載 ---
def set_msjh_font(run):
    run.font.name = 'Microsoft JhengHei'
    r = run._element
    rFonts = r.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        r.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

def generate_word():
    doc = Document()
    h = doc.add_heading('行銷企劃執行提案書', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "未命名活動", level=1)
    for fid, title, _, _ in sections:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        data = generate_word()
        st.download_button(label=f"📥 下載 {st.session_state.p_name} 企劃書", data=data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
