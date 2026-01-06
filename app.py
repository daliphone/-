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
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.2.1", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; }
    ::placeholder { color: #888888 !important; opacity: 0.5 !important; }
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; width: 100%; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    /* 側邊欄樣式 */
    section[data-testid="stSidebar"] { background-color: #0B1C3F; color: white; }
    section[data-testid="stSidebar"] .stMarkdown h2 { color: #FFD700 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化 Session State ---
# 欄位清單
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]

for field in FIELDS:
    if field not in st.session_state:
        st.session_state[field] = ""
if "p_proposer" not in st.session_state or not st.session_state.p_proposer:
    st.session_state.p_proposer = "行銷部"

# --- 3. 側邊欄：範本與系統管理 ---
with st.sidebar:
    st.header("📋 範本與草稿管理")
    
    # 範本數據 (整合馬年慶與試戴專案邏輯)
    if st.button("🐎 載入：馬年慶 (百倍奉還)"):
        st.session_state.p_name = "2026 馬尼通訊「馬年慶：百倍奉還」"
        st.session_state.p_purpose = "迎接馬年，透過 $100 低門檻吸引新舊客，增加會員與官網流量。"
        st.session_state.p_core = "對象：全體消費者；核心產品：100元新年禮包。"
        st.session_state.p_schedule = "115/01/12: 宣傳期\n115/01/19: 銷售期"
        st.session_state.p_prizes = "PS5 | 1名 | 大獎\n現金 6666 | 1名 | 獎金"
        st.rerun()

    if st.button("🗑️ 清除所有草稿"):
        for field in FIELDS:
            st.session_state[field] = ""
        st.session_state.p_proposer = "行銷部"
        st.success("草稿已清空")
        st.rerun()

    st.divider()
    st.header("✨ AI 優化風格")
    ai_style = st.radio("主要語氣", ["熱血商務", "創意社群", "專業條列"])

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊", expanded=False):
        st.caption("""
        **版本**: v14.2.1 (Stability)
        - 恢復清除草稿功能
        - 新增可編輯參考範例區
        - 恢復灰色引導文字 (Placeholder)
        
        馬尼門活動企劃系統 © 2025 Money MKT
        """)

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
with c_top1: st.text_input("一、 活動名稱", key="p_name", placeholder="例如: 7日智慧手錶體驗方案｜先體驗，再入手")
with c_top2: st.text_input("提案人", key="p_proposer")
with c_top3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 定義參考建議內容 (依據二、 建議新活動參考的章節與順序)
tips = {
    "purpose": "【建議 1】核心價值：定義活動解決什麼痛點？（如：降低首次購買門檻）。量化目標：預計帶動的人流量或 UGC 素材數量。",
    "core": "【建議 2】機制設計：分階段說明申請/開始、體驗、結束。透明價格表：列出成本、售價、優惠價與押金對應關係。",
    "schedule": "【建議 3】明確時程：包含提案期、整備期、宣傳期、銷售期。建議宣傳期需於銷售期前 7 日啟動。",
    "prizes": "【建議 4】誘因機制：任務化獎勵（如分享即贈小禮）。區分購買與否：即使未成交，只要有回饋也給予小贈品建立信任。",
    "sop": "【建議 7】實戰話術：1. 卸下武裝（先聊需求不推產品）。2. 反向推銷（建議先試戴不要直接買）。3. 禁語清單（避開「今天不買會沒了」）。",
    "marketing": "【建議 4】擴算機制：設計社群任務（標記官方帳號）、FB/IG/Threads 倒數計時增加緊張感。",
    "risk": "【建議 6】控管機制：明確定義損壞界定（如機身傷痕、進水）。稅務規範（> $1,000 需申報）。銷售分佈不均的調度方案。",
    "effect": "【建議 5】數據蒐集：設計問卷詢問「影響購買主要原因」。分析體驗是否幫助決策，作為優化話術依據。"
}

c1, c2 = st.columns(2)

with c1:
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 1 內容", value=tips["purpose"], height=70)
    st.text_area("活動時機與目的", key="p_purpose", height=100, placeholder="(節日活動，透過指定促銷或搭贈，增加成交機率與新客。)")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 2 內容", value=tips["core"], height=70)
    st.text_area("二、 活動核心內容", key="p_core", height=100, placeholder="執行單位、對象、主要商品賣點...")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 3 內容", value=tips["schedule"], height=70)
    st.text_area("三、 活動時程安排", key="p_schedule", height=120, placeholder="提案期、整備期、宣傳期、銷售期...")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 4 內容", value=tips["prizes"], height=70)
    st.text_area("四、 贈品結構與預算", key="p_prizes", height=120, placeholder="品項 | 數量 | 預算配置...")

with c2:
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 7 內容", value=tips["sop"], height=70)
    st.text_area("五、 門市執行 SOP (含話術)", key="p_sop", height=100, placeholder="先幫客人卸下武裝、反向推銷、禁語標籤...")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 4 內容 (擴散)", value=tips["marketing"], height=70)
    st.text_area("六、 行銷宣傳與策略", key="p_marketing", height=100, placeholder="希望曝光的管道、社群回饋任務內容...")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 6 內容", value=tips["risk"], height=70)
    st.text_area("七、 風險管理與退場機制", key="p_risk", height=100, placeholder="損壞判定、稅務規範、退場機制說明...")
    
    with st.expander("💡 參考建議 (可編輯後複製)", expanded=False):
        st.text_area("建議 5 內容", value=tips["effect"], height=100)
    st.text_area("八、 預估成效與數據蒐集", key="p_effect", height=100, placeholder="預期業績、UGC 累積數量、問卷核心指標...")

# --- 5. Word 導出與下載 (穩定邏輯) ---
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
    r_info = info_p.add_run(f"提案人：{st.session_state.p_proposer}  |  日期：{st.session_state.p_date}")
    set_msjh_font(r_info)

    doc.add_heading(st.session_state.p_name, level=1)

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
        r = p.add_run(content)
        set_msjh_font(r)

    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        data = generate_pro_word()
        st.download_button(
            label=f"📥 下載 {st.session_state.p_name} 企劃書",
            data=data,
            file_name=f"MoneyMKT_{st.session_state.p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
