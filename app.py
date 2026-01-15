import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os
import google.generativeai as genai

# --- 1. 頁面配置與清新感 UI ---
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.3.6", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    /* 全域清新感底色 */
    .main { background-color: #F8FAFC; color: #1E293B; }
    
    /* 左側側邊欄：重新美化為清新白底藍邊 */
    [data-testid="stSidebar"] { 
        background-color: #FFFFFF !important; 
        border-right: 1px solid #E2E8F0 !important;
        box-shadow: 2px 0px 10px rgba(0,0,0,0.02);
    }
    /* 側邊欄標題與文字顏色調教 */
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #003f7e !important; font-weight: 700; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { 
        color: #475569 !important; font-size: 14px; 
    }
    
    /* 下拉選單清新化 */
    div[data-baseweb="select"] > div { background-color: #F1F5F9 !important; border: none !important; color: #0F172A !important; }

    /* 章節標題強化 */
    .section-header { 
        font-size: 18px !important; color: #003f7e !important; font-weight: 700 !important; 
        margin-top: 25px !important; margin-bottom: 8px !important;
        display: flex; align-items: center;
    }
    .section-header::before {
        content: ""; display: inline-block; width: 4px; height: 20px; 
        background-color: #ef8200; margin-right: 10px; border-radius: 2px;
    }
    
    /* AI 按鈕精簡化 */
    .ai-btn-small>div>button { 
        background-color: #F5F3FF !important; color: #6D28D9 !important; 
        border: 1px solid #DDD6FE !important; font-size: 12px !important;
        transition: 0.3s;
    }
    .ai-btn-small>div>button:hover { background-color: #6D28D9 !important; color: white !important; }
    
    /* 修正引導文字顏色 */
    textarea::placeholder { color: #94A3B8 !important; font-style: italic; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化 Session State 與範本庫 ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "請選擇範本": {f: "" for f in FIELDS},
        "🐎 馬年慶：百倍奉還": {
            "p_name": "2026「馬年慶：百倍奉還」",
            "p_purpose": "解決連假後人流痛點，透過 $100 門檻去化新年禮包庫存，達成營運目的邏輯中的數據增長。",
            "p_core": "對象：全門市。賣點：$100 低門檻試手氣。",
            "p_sop": "核心話術：先聊新年願望。SOP：限購3包、引導加官方LINE。",
            "p_marketing": "FB/IG 視覺紅包標語。",
            "p_risk": "每店配額管理。",
            "p_effect": "預計進店人次 +20%。"
        },
        "⌚ 7日智慧手錶試戴": {
            "p_name": "「先體驗再入手」7日試戴方案",
            "p_purpose": "降低高單價智慧手錶購買門檻，解決適配感痛點並增加銷售目的。",
            "p_sop": "核心原則：卸下心理武裝。話術：建議先不要買，戴過再說。",
            "p_risk": "押金全額退還規範。",
            "p_effect": "提升試戴後轉換率。"
        }
    }

for field in FIELDS:
    if field not in st.session_state: st.session_state[field] = ""

# --- 3. 側邊欄：範本與清新管理 ---
with st.sidebar:
    st.header("📋 企劃管理")
    
    # 這裡會隨著儲存動作動態更新下拉選單清單
    selected_tpl_key = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    col_l, col_r = st.columns(2)
    with col_l:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl_key]
            for k, v in data.items(): st.session_state[k] = v
            st.rerun()
    with col_r:
        if st.button("💾 儲存範本"):
            if st.session_state.p_name:
                new_key = f"💾 {st.session_state.p_name[:10]}"
                st.session_state.templates_store[new_key] = {f: st.session_state[f] for f in FIELDS}
                st.success("儲存成功！")
                st.rerun() # 重新整理以更新下拉選單
            else:
                st.error("請先輸入活動名稱")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

# --- 4. 分章節數據配置 ---
# 章節ID, 標題, 營運邏輯(引導文), 實戰建議(展開區)
sections_config = [
    ("p_purpose", "一、 活動時機與目的", "營運目的邏輯：強化解決痛點（如降低購買門檻）與數據增長，增加目標商品銷售或是去化高壓商品。", "核心：春節紅包議題，解決「連假後人流下降」痛點。目標：引導消費者消耗紅包財。"),
    ("p_core", "二、 活動核心內容", "賣點配置建議：依據名稱、對象、執行單位、主要賣點，建立「低門檻、零風險」誘因。", "機制：購買禮包獲得序號。定價：$100 元具備衝動購買力，適合快速成交。"),
    ("p_schedule", "三、 活動時程安排", "執行重點建議：不安排AI潤稿，而是期間需要執行的重點，宣傳、銷售、結案期的資源分配。", "時程：1月中旬啟動，確保除夕前銷售完畢，開工後吸引回流。"),
    ("p_prizes", "四、 贈品結構與預算", "配置用意與賣點：不安排AI潤稿，提供關鍵商品配置用意與賣點，平衡大獎話題與小獎導流。", "配置：PS5 (話題性) + 現金 (實用性)。購物金用於強化官網引流。"),
    ("p_sop", "五、 門市執行流程 (SOP)", "執行環節注意事項建議：不安排AI潤稿，提供執行環節注意事項建議，注入「卸下武裝」策略。", "話術：先聊願望再推「試手氣」。SOP：強調序號正本為兌獎唯一憑證。"),
    ("p_marketing", "六、 行銷流程與策略", "行銷策略：提供建議管道與AI潤稿，自動推薦適合管道並生成行銷標語。", "宣傳：紅包色系視覺，社群任務設計「分享好運」抽額外購物金。"),
    ("p_risk", "七、 風險管理與注意事項", "風險管理建議：活動規範以及注意相關事項建議，針對法務、稅務及產品損壞進行規範。", "風險：每店配額管理避免落空。法規：中獎者身份證影本蒐集以便申報。"),
    ("p_effect", "八、 預估成效", "成效效益面建議：預估活動可以帶來的效益面建議，分析 O2O 轉換、名單累積與問卷數據價值。", "指標：門市進店率、官網註冊數、二次消費轉化率。")
]

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 模組化企劃系統 v14.3.6")

t1, t2, t3 = st.columns([2, 1, 1])
with t1: st.text_input("活動名稱", key="p_name", placeholder="請輸入本案活動名稱")
with t2: st.text_input("提案人", key="p_proposer")
with t3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

col_a, col_b = st.columns(2)
for i, (fid, title, logic_guide, real_tip) in enumerate(sections_config):
    target_col = col_a if i < 4 else col_b
    with target_col:
        # 章節標題
        st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
        
        # 填寫框：提示文帶入「營運邏輯與優化方向」
        st.text_area("", key=fid, height=140, placeholder=logic_guide, label_visibility="collapsed")
        
        # 功能區
        c_ai, c_tip = st.columns([1, 1])
        with c_ai:
            # 只有特定章節需要 AI 潤稿
            if fid in ["p_purpose", "p_core", "p_marketing", "p_risk", "p_effect"]:
                st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
                if st.button(f"🪄 AI 優化文字", key=f"btn_{fid}"):
                    # 模擬或呼叫 AI 邏輯
                    st.session_state[fid] = f"【AI 優化結果】{st.session_state[fid]}"
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.caption("本章節採用實戰建議填充")
        
        with c_tip:
            # 查看建議：帶入「實戰引導內容」
            with st.expander("💡 查看實戰建議", expanded=False):
                st.write(real_tip)

# --- 6. Word 產出邏輯 ---
def generate_word():
    doc = Document()
    doc.add_heading('行銷企劃執行提案書 v14.3.6', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "未命名活動", level=1)
    for fid, title, _, _ in sections_config:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_word()
        st.download_button(label="📥 下載模組化企劃書", data=doc_data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
