import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
from docx import Document
from io import BytesIO

# --- 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v11.0", page_icon="🐎", layout="wide")

# 強制品牌風格與自訂按鈕
st.markdown("""
    <style>
    .main { background-color: #0B1C3F; }
    h1, h2, h3 { color: #FFD700 !important; }
    .stButton>button { background-color: #F39C12; color: white; border-radius: 8px; font-weight: bold; width: 100%; }
    .event-load-btn > div > button { background-color: #D32F2F !important; border: 2px solid #FFD700 !important; }
    .section-box { padding: 15px; border: 1px solid #FFD700; border-radius: 10px; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("馬尼通訊 行銷企劃提案系統")

# --- 初始化狀態 ---
if 'activity_list' not in st.session_state:
    st.session_state.activity_list = []

# --- 預設載入邏輯：馬年慶：百倍奉還 ---
def load_horse_year_event():
    st.session_state.p_name = "2026 馬年慶：百倍奉還抽獎活動"
    st.session_state.p_purpose = "1. 帶動春節人流量。\n2. 透過 $100 低門檻吸引新舊客，增加會員與官網流量。\n3. 建立品牌高性價比形象。"
    st.session_state.p_core = "活動對象：所有門市消費者（每人限購 3 包）\n參與單位：馬尼行動通訊任一門市\n核心商品：「百倍奉還」新年禮包（售價 $100/包）"
    st.session_state.p_schedule = "宣傳期：01/12-01/18\n販售期：01/19-02/08\n抽獎準備：02/09-02/10\n開獎日：02/11 (三)\n兌獎期：02/12-02/28"
    st.session_state.p_prizes = "總獎值：突破 $130,000 元\n1. Sony PS5 (1名)\n2. 現金 $6,666 (1名)\n3. Apple Watch SE2 (2名)\n4. 官網購物金 $1,500 (115名)"
    st.session_state.p_sop = "1. 每人上限 3 包，主動告知內含序號。\n2. 每店限量 66 包，售罄張貼完售告示。\n3. 引導加入官方 LINE 綁定會員資料。"
    st.session_state.p_marketing = "1. FB/IG/脆製作倒數計時限時動態。\n2. 廣告標語：只要 100 元，PS5 搬回家！\n3. 針對弱勢分店進行 3-5 公里 FB 區域廣告投遞。"
    st.session_state.p_risk = "1. 稅務：> $1,000 需身分證影本；> $20,000 扣繳 10% 稅金。\n2. 爭議：序號需蓋章確認，避免影印冒領。\n3. 調度：第 10 天進行盤點，將剩餘庫存調往熱門門市。"
    st.session_state.p_effect = "1. 預估帶動 2,000+ 人次進店。\n2. 115 名購物金中獎者產生二次消費。\n3. 數據留存：建立潛在行銷名單。"

# --- 側邊欄 ---
with st.sidebar:
    st.header("🧧 企劃範本快捷鍵")
    st.markdown('<div class="event-load-btn">', unsafe_allow_html=True)
    if st.button("🐎 載入【馬年慶：百倍奉還】完整企劃"):
        load_horse_year_event()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
    if st.button("🗑️ 清空目前草稿"):
        for key in st.session_state.keys():
            if key.startswith("p_"): del st.session_state[key]
        st.rerun()

# --- 主要編輯區：依據文件邏輯性排序 ---
st.subheader("📝 企劃提案填寫區")

with st.expander("一、 活動時機與目的", expanded=True):
    p_name = st.text_input("1. 活動名稱", key="p_name")
    p_purpose = st.text_area("2. 活動背景與目的 (時間/檔期/目標)", key="p_purpose", height=100)

with st.expander("二、 活動核心內容", expanded=True):
    p_core = st.text_area("對象、參與單位、核心活動商品", key="p_core", height=100)

with st.expander("三、 活動時程安排", expanded=True):
    p_schedule = st.text_area("提案/製作/宣傳/銷售/抽獎/開獎/兌獎期", key="p_schedule", height=120)

with st.expander("四、 獎項結構與預算", expanded=True):
    p_prizes = st.text_area("獎項配置、總獎值、贈品細節", key="p_prizes", height=120)

with st.expander("五、 門市執行流程 (SOP)", expanded=True):
    p_sop = st.text_area("銷售環節、限量管理、個資蒐集規範", key="p_sop", height=120)

with st.expander("六、 行銷宣傳策略", expanded=True):
    p_marketing = st.text_area("線上管道、廣告標語、弱勢分店加碼策略", key="p_marketing", height=120)

with st.expander("七_風險管理與注意事項", expanded=True):
    p_risk = st.text_area("稅務法規、序號爭議、缺貨調度機制", key="p_risk", height=120)

with st.expander("八、 預估成效", expanded=True):
    p_effect = st.text_area("觸及人數、官網互動、品牌曝光目標", key="p_effect", height=100)

# --- 匯出功能 ---
st.divider()
if st.button("🚀 生成並預覽企劃清單"):
    st.session_state.activity_list.append({
        "名稱": p_name, "內容": f"【目的】\n{p_purpose}\n\n【核心】\n{p_core}\n\n【時程】\n{p_schedule}\n\n【獎項】\n{p_prizes}\n\n【SOP】\n{p_sop}\n\n【行銷】\n{p_marketing}\n\n【風險】\n{p_risk}\n\n【成效】\n{p_effect}"
    })
    st.success("已成功生成企劃草稿！")

if st.session_state.activity_list:
    # Word 生成邏輯
    doc = Document()
    doc.add_heading('馬尼通訊 行銷企劃執行案', 0)
    
    current_p = st.session_state.activity_list[-1] # 抓取最後一筆
    
    sections = [
        ("一、 活動名稱與目的", p_purpose),
        ("二、 活動核心內容", p_core),
        ("三、 活動時程安排", p_schedule),
        ("四、 獎項與贈品結構", p_prizes),
        ("五、 門市執行流程", p_sop),
        ("六、 行銷流程與策略", p_marketing),
        ("七、 風險管理與注意事項", p_risk),
        ("八、 預估成效", p_effect)
    ]
    
    doc.add_heading(current_p['名稱'], level=1)
    for title, content in sections:
        doc.add_heading(title, level=2)
        doc.add_paragraph(content)
        
    word_io = BytesIO()
    doc.save(word_io)
    
    st.download_button(
        label="📄 下載完整企劃書 (.docx)",
        data=word_io.getvalue(),
        file_name=f"馬尼企劃_{p_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
