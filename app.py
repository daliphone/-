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
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.2.2", page_icon="🐎", layout="wide")

# CSS 強制修正：確保側邊欄文字清晰，Placeholder 顏色正確
st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; }
    ::placeholder { color: #888888 !important; opacity: 0.7 !important; }
    
    /* 側邊欄視覺強制修正 */
    [data-testid="stSidebar"] { background-color: #0B1C3F !important; }
    [data-testid="stSidebar"] .stMarkdown h2, [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] .stMarkdown h2 { color: #FFD700 !important; }
    
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; width: 100%; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化 Session State ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]

for field in FIELDS:
    if field not in st.session_state:
        st.session_state[field] = ""

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還 (官方範本)": {
            "p_name": "2026 馬尼通訊「馬年慶：百倍奉還」抽獎企劃",
            "p_purpose": "迎接馬年，結合春節紅包議題，透過 $100 低門檻吸引新舊客回流門市。",
            "p_core": "對象：全門市消費者；核心產品：$100 新年禮包（含抽獎券）。",
            "p_schedule": "115/01/12-01/18: 宣傳期\\n115/01/19-02/08: 銷售期",
            "p_prizes": "PS5 | 1名 | 大獎\\n現金 $6666 | 1名 | 百倍奉還獎",
            "p_sop": "確認每人限購3包、引導加官方LINE、提醒保留序號至開獎日。",
            "p_marketing": "FB/IG/脆 倒數計時動態、門市張貼限量完售海報。",
            "p_risk": "稅務處理(>$1,000)、序號防偽蓋章、確保退場機制清楚。",
            "p_effect": "預計帶動 2,000+ 人流，官網流量提升 30%。"
        }
    }

# --- 3. 側邊欄：範本管理與系統功能 ---
with st.sidebar:
    st.header("📋 企劃範本庫")
    
    # 範本選擇
    selected_tpl = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl]
            for k, v in data.items():
                st.session_state[k] = v
            st.rerun()
    with c_btn2:
        if st.button("💾 儲存為範本"):
            new_key = f"💾 自訂：{st.session_state.p_name[:10]}..." if st.session_state.p_name else f"💾 自訂：{datetime.now().strftime('%m%d%H%M')}"
            st.session_state.templates_store[new_key] = {f: st.session_state[f] for f in FIELDS}
            st.success("已存入範本庫")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊 v14.2.2", expanded=False):
        st.caption("修正：\n1. 範本儲存與載入功能\n2. 側邊欄對比度視覺修正\n3. 建議框改至輸入框下方\n4. 預設馬年慶範例")

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
with c_top1: st.text_input("一、 活動名稱", key="p_name", placeholder="例如: 2026 馬尼通訊「馬年慶：百倍奉還」")
with c_top2: st.text_input("提案人", key="p_proposer", placeholder="行銷部 / 姓名")
with c_top3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 馬年慶背景的專業建議內容
tips = {
    "purpose": "【馬年慶建議】核心：春節紅包話題，解決「連假後人流下降」痛點。目標：引導消費者在春節前後進入門市消耗紅包財。",
    "core": "【馬年慶建議】機制：購買禮包->獲得序號->線上開獎。定價：$100 元具備衝動性購買力，適合快速成交。",
    "schedule": "【馬年慶建議】時程：1月中旬啟動宣傳，確保除夕前銷售完畢。開獎日定於開工後，吸引二次回流。",
    "prizes": "【馬年慶建議】配置：PS5(話題性)+現金(實用性)。官網購物金能有效將線下人流導向電子商務，建議數量要多。",
    "sop": "【馬年慶建議】話術：先聊新年願望，再推「100元試手氣」。SOP：必須強調序號正本是兌獎唯一憑證。",
    "marketing": "【馬年慶建議】宣傳：利用紅包色系視覺，社群任務可設計「分享好運」抽額外購物金。",
    "risk": "【馬年慶建議】風險：每店配額管理，避免消費者跨區購買落空。法規：務必收齊中獎者身份證影本以便申報。",
    "effect": "【馬年慶建議】指標：門市進店率、官網註冊數、二次消費轉化率。"
}

c1, c2 = st.columns(2)

with c1:
    st.text_area("活動時機與目的", key="p_purpose", height=120, placeholder="迎接馬年，透過 $100 低門檻吸引新舊客...")
    with st.expander("💡 參考建議：營運目的 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["purpose"], height=80)
    
    st.text_area("二、 活動核心內容", key="p_core", height=120, placeholder="對象、執行單位與產品核心...")
    with st.expander("💡 參考建議：核心賣點 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["core"], height=80)
        
    st.text_area("三、 活動時程安排", key="p_schedule", height=120, placeholder="日期: 執行內容...")
    with st.expander("💡 參考建議：時程建議 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["schedule"], height=80)
        
    st.text_area("四、 贈品結構與預算", key="p_prizes", height=120, placeholder="品項 | 數量 | 備註")
    with st.expander("💡 參考建議：獎項配置 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["prizes"], height=80)

with c2:
    st.text_area("五、 門市執行 SOP (含話術)", key="p_sop", height=120, placeholder="銷售話術、限量管理與序號核對...")
    with st.expander("💡 參考建議：實戰話術 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["sop"], height=80)
        
    st.text_area("六、 行銷宣傳與策略", key="p_marketing", height=120, placeholder="線上廣告與標語策略...")
    with st.expander("💡 參考建議：行銷擴散 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["marketing"], height=80)
        
    st.text_area("七、 風險管理與注意事項", key="p_risk", height=120, placeholder="稅務法規、調度與序號防偽...")
    with st.expander("💡 參考建議：風險控管 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["risk"], height=80)
        
    st.text_area("八、 預估成效", key="p_effect", height=120, placeholder="人流、轉化率與 UGC 預期...")
    with st.expander("💡 參考建議：效益預估 (馬年慶背景)", expanded=False):
        st.text_area("建議內容 (可修改後複製)", value=tips["effect"], height=80)

# --- 5. Word 導出邏輯 ---
def set_msjh_font(run):
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
    h = doc.add_heading('行銷企劃執行提案書', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_info = info_p.add_run(f"提案人：{st.session_state.p_proposer if st.session_state.p_proposer else '行銷部'}  |  日期：{st.session_state.p_date}")
    set_msjh_font(r_info)

    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "活動企劃書", level=1)

    sections = [
        ("一、 活動時機與目的", st.session_state.p_purpose),
        ("二、 活動核心內容", st.session_state.p_core),
        ("三、 活動時程安排", st.session_state.p_schedule),
        ("四、 贈品結構與預算", st.session_state.p_prizes),
        ("五、 門市執行流程 (SOP)", st.session_state.p_sop),
        ("六、 行銷流程與策略", st.session_state.p_marketing),
        ("七、 風險管理與注意事項", st.session_state.p_risk),
        ("八、 預估成效", st.session_state.p_effect)
    ]

    for title, content in sections:
        h2 = doc.add_heading(title, level=2)
        h2.runs[0].font.color.rgb = RGBColor(11, 28, 63)
        p = doc.add_paragraph()
        r = p.add_run(content if content else "（尚未填寫）")
        set_msjh_font(r)

    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_pro_word()
        st.download_button(
            label=f"📥 下載 {st.session_state.p_name} 企劃書",
            data=doc_data,
            file_name=f"MoneyMKT_{st.session_state.p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
