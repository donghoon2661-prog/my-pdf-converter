import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import glob
from datetime import datetime

# ==========================================
# 1. 페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================
# 2. 커스텀 CSS (깔끔한 디자인)
# ==========================================
st.markdown("""
<style>
    /* 전체 폰트 & 배경 */
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Apple SD Gothic Neo', -apple-system, sans-serif;
    }

    /* 메인 타이틀 영역 */
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

    /* 카드형 컨테이너 */
    .card {
        background: #FFFFFF;
        border: 1px solid #EDEFF3;
        border-radius: 14px;
        padding: 1.5rem 1.6rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 1.2rem;
    }

    /* 버튼 스타일 */
    .stButton>button, .stDownloadButton>button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
        border: none;
        transition: all 0.15s ease-in-out;
    }
    .stButton>button[kind="primary"] {
        background-color: #3D5AFE;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #2f47d6;
        box-shadow: 0 2px 8px rgba(61, 90, 254, 0.35);
    }
    .stDownloadButton>button {
        background-color: #16A34A;
        color: white;
    }
    .stDownloadButton>button:hover {
        background-color: #128a3e;
    }

    /* 사이드바 */
    section[data-testid="stSidebar"] {
        background-color: #FAFBFD;
        border-right: 1px solid #ECEFF3;
    }
    .sidebar-file-card {
        background: white;
        border: 1px solid #ECEFF3;
        border-radius: 10px;
        padding: 0.8rem 0.9rem;
        margin-bottom: 0.8rem;
    }

    /* 구분선 */
    hr {
        margin: 1.2rem 0;
        border: none;
        border-top: 1px solid #EDEFF3;
    }

    /* 업로더 라벨 여백 정리 */
    .stFileUploader label {
        font-weight: 600;
        font-size: 0.95rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. 헤더
# ==========================================
st.markdown("""
<div class="main-header">
    <h1>📄 COA 성적서 PDF → 엑셀 변환기</h1>
    <p>PDF 파일을 업로드하면 내부의 모든 표를 하나의 통합 엑셀 파일로 정리해 드립니다.</p>
</div>
""", unsafe_allow_html=True)

# 임시 저장용 폴더
TEMP_DIR = "temporary_excel_storage"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# 세션 상태 초기화
if "conversion_done" not in st.session_state:
    st.session_state.conversion_done = False
if "excel_data_bytes" not in st.session_state:
    st.session_state.excel_data_bytes = None
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = ""

# ==========================================
# 사이드바: 최근 파일 관리
# ==========================================
with st.sidebar:
    st.markdown("### 📁 최근 변환 파일")
    st.caption("최대 5개까지 서버에 임시 보관됩니다.")
    st.markdown("---")

    existing_files = glob.glob(os.path.join(TEMP_DIR, "*.xlsx"))
    existing_files.sort(key=os.path.getmtime, reverse=True)

    if len(existing_files) > 5:
        for old_file in existing_files[5:]:
            try:
                os.remove(old_file)
            except Exception:
                pass
        existing_files = existing_files[:5]

    if existing_files:
        for idx, file_path in enumerate(existing_files):
            file_name = os.path.basename(file_path)
            if "_" in file_name:
                timestamp_prefix, pure_name = file_name.split("_", 1)
            else:
                timestamp_prefix, pure_name = "", file_name

            st.markdown(f'<div class="sidebar-file-card">', unsafe_allow_html=True)
            st.markdown(f"**📄 {pure_name}**")

            pure_name_no_ext = pure_name.replace(".xlsx", "")
            new_name_input = st.text_input(
                "이름 수정",
                value=pure_name_no_ext,
                key=f"rename_input_{file_path}",
                label_visibility="collapsed",
            )

            with open(file_path, "rb") as f:
                file_bytes = f.read()

            c1, c2, c3 = st.columns([1.3, 1, 1])
            c1.download_button(
                "💾", data=file_bytes, file_name=pure_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_btn_{file_path}", help="다운로드",
            )

            if new_name_input != pure_name_no_ext:
                if c2.button("✅", key=f"save_btn_{file_path}", help="이름 저장"):
                    new_full_name = f"{timestamp_prefix}_{new_name_input}.xlsx" if timestamp_prefix else f"{new_name_input}.xlsx"
                    new_file_path = os.path.join(TEMP_DIR, new_full_name)
                    try:
                        os.rename(file_path, new_file_path)
                        st.toast("✅ 이름이 변경되었습니다!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

            if c3.button("🗑️", key=f"del_btn_{file_path}", help="삭제"):
                try:
                    os.remove(file_path)
                    st.toast("🗑️ 삭제되었습니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"오류: {e}")

            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("아직 임시 저장된 파일이 없습니다.")

# ==========================================
# 메인: 업로드 & 변환
# ==========================================
st.markdown('<div class="card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("변환할 PDF 파일을 선택하세요", type=["pdf"])
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    pdf_base_name = uploaded_file.name.replace(".pdf", "")

    if st.session_state.current_file_name != uploaded_file.name:
        st.session_state.conversion_done = False
        st.session_state.excel_data_bytes = None
        st.session_state.current_file_name = uploaded_file.name

    st.markdown('<div class="card">', unsafe_allow_html=True)
    col_name, col_btn = st.columns([3, 1])
    with col_name:
        custom_file_name = st.text_input(
            "📥 다운로드할 엑셀 파일 이름",
            value=pdf_base_name,
        )

    if not custom_file_name.endswith(".xlsx"):
        final_excel_name = custom_file_name + ".xlsx"
    else:
        final_excel_name = custom_file_name

    with col_btn:
        st.write("")
        st.write("")
        if not st.session_state.conversion_done:
            start_conversion = st.button("🚀 변환하기", type="primary", use_container_width=True)
        else:
            start_conversion = False
            st.button("✅ 완료됨", disabled=True, use_container_width=True)

    status_text = st.empty()
    st.markdown('</div>', unsafe_allow_html=True)

    if start_conversion and not st.session_state.conversion_done:
        progress_bar = st.progress(0)
        status_text.info("⏳ PDF 파일을 분석하고 있습니다...")
        output_buffer = io.BytesIO()

        try:
            total_tables_found = 0

            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                with pdfplumber.open(uploaded_file) as pdf:
                    total_pages = len(pdf.pages)

                    for page_num in range(total_pages):
                        current_page_idx = page_num + 1
                        status_text.info(
                            f"⚙️ **{current_page_idx} / {total_pages}** 페이지 분석 중 "
                            f"· 표 {total_tables_found}개 발견"
                        )

                        page = pdf.pages[page_num]
                        tables = page.extract_tables()

                        if not tables:
                            continue

                        for i, table in enumerate(tables):
                            df = pd.DataFrame(table)
                            df.dropna(how='all', inplace=True)

                            sheet_name = f"Page{current_page_idx}_Table{i + 1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                            total_tables_found += 1

                        progress_bar.progress(current_page_idx / total_pages)

            st.session_state.excel_data_bytes = output_buffer.getvalue()
            st.session_state.conversion_done = True
            progress_bar.empty()

            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            server_save_path = os.path.join(TEMP_DIR, f"{timestamp}_{final_excel_name}")
            with open(server_save_path, "wb") as f:
                f.write(st.session_state.excel_data_bytes)

            st.rerun()

        except Exception as e:
            status_text.error(f"❌ 변환 도중 오류가 발생했습니다: {e}")
            st.session_state.conversion_done = False

    if st.session_state.conversion_done and st.session_state.excel_data_bytes is not None:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.success("🎉 변환이 완료되었습니다!")
        st.caption("왼쪽 사이드바에도 보관되었습니다. 창을 닫아도 최근 5개 목록에서 다시 받을 수 있어요.")
        st.download_button(
            label=f"📥 {final_excel_name} 다운로드",
            data=st.session_state.excel_data_bytes,
            file_name=final_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown("""
    <div style="text-align:center; padding: 3rem 1rem; color: #9AA1AC;">
        <p style="font-size:0.9rem;">PDF 파일을 업로드하면 변환을 시작할 수 있어요.</p>
    </div>
    """, unsafe_allow_html=True)
