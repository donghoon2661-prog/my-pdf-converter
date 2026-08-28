# -*- coding: utf-8 -*-

import streamlit as st


# ==========================================
# 1. 페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="진구",
    page_icon="🧰",
    layout="wide",
)


# ==========================================
# 2. 커스텀 CSS
# ==========================================
st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Apple SD Gothic Neo', -apple-system, sans-serif;
    }

    .main-header {
        padding: 1.8rem 2rem;
        background: linear-gradient(135deg, #4F6EF7 0%, #3D5AFE 100%);
        border-radius: 16px;
        margin-bottom: 1.5rem;
        color: white;
    }

    .main-header h1 {
        font-size: 1.6rem;
        font-weight: 700;
        margin: 0 0 0.3rem 0;
        color: white;
    }

    .main-header p {
        font-size: 0.92rem;
        opacity: 0.9;
        margin: 0;
    }

    .card {
        background: #FFFFFF;
        border: 1px solid #EDEFF3;
        border-radius: 14px;
        padding: 1.5rem 1.6rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 1.2rem;
    }

    .stButton>button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. 헤더
# ==========================================
st.markdown("""
<div class="main-header">
    <h1>🧰 진구 사이트</h1>
    <p>진구 관련 기능을 관리하는 별도 페이지입니다.</p>
</div>
""", unsafe_allow_html=True)


# ==========================================
# 4. 메인 화면
# ==========================================
st.markdown('<div class="card">', unsafe_allow_html=True)

st.markdown("### 🚧 진구 페이지 준비 중")

st.info(
    "이 페이지는 앞으로 진구 관련 기능을 별도로 추가하기 위한 공간입니다."
)

st.markdown('</div>', unsafe_allow_html=True)


# ==========================================
# 5. 기존 메인 페이지로 돌아가기
# ==========================================
if st.button(
    "⬅️ 메인 페이지로 돌아가기",
    use_container_width=True
):
    st.switch_page("app.py")
