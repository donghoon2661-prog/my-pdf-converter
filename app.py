import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import glob
from datetime import datetime

# 1. 웹사이트 타이틀 및 레이아웃 설정
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")
st.title("📄 COA 성적서 PDF ➡️ 엑셀 변환기 (전체 페이지)")
st.write("PDF 파일을 업로드하고 '변환하기' 버튼을 누르면 내부의 모든 표를 하나의 통합 엑셀 파일로 추출해 드립니다.")

st.markdown("---")

# 임시 저장용 폴더 설정 (Streamlit 서버 내부 폴더)
TEMP_DIR = "temporary_excel_storage"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# 💡 [3번 요청 사항] 무한 변환 방지를 위한 변환 상태(Session State) 기억 장치 세팅
if "conversion_done" not in st.session_state:
    st.session_state.conversion_done = False
if "excel_data_bytes" not in st.session_state:
    st.session_state.excel_data_bytes = None
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = ""

# ==========================================
# [사이드바 기능] 최근 파일 목록 (수정 / 삭제 대시보드)
# ==========================================
st.sidebar.title("📁 최근 변환된 파일 관리 (최대 5개)")
st.sidebar.write("최근 파일의 이름을 수정하거나 서버에서 영구 삭제할 수 있습니다.")

existing_files = glob.glob(os.path.join(TEMP_DIR, "*.xlsx"))
existing_files.sort(key=os.path.getmtime, reverse=True)

if len(existing_files) > 5:
    for old_file in existing_files[5:]:
        try: os.remove(old_file)
        except: pass
    existing_files = existing_files[:5]

if existing_files:
    for idx, file_path in enumerate(existing_files):
        file_name = os.path.basename(file_path)
        if "_" in file_name:
            timestamp_prefix, pure_name = file_name.split("_", 1)
        else:
            timestamp_prefix, pure_name = "", file_name
            
        st.sidebar.markdown(f"**📄 파일 {idx+1}**")
        
        pure_name_no_ext = pure_name.replace(".xlsx", "")
        new_name_input = st.sidebar.text_input(
            "✏️ 이름 수정:", 
            value=pure_name_no_ext, 
            key=f"rename_input_{file_path}"
        )
        
        col1, col2, col3 = st.sidebar.columns([2, 1, 1])
        
        with open(file_path, "rb") as f:
            file_bytes = f.read()
            
        col1.download_button(
            label="💾 다운로드",
            data=file_bytes,
            file_name=pure_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_btn_{file_path}"
        )
        
        if new_name_input != pure_name_no_ext:
            if col2.button("💾 저장", key=f"save_btn_{file_path}"):
                new_full_name = f"{timestamp_prefix}_{new_name_input}.xlsx" if timestamp_prefix else f"{new_name_input}.xlsx"
                new_file_path = os.path.join(TEMP_DIR, new_full_name)
                try:
                    os.rename(file_path, new_file_path)
                    st.toast("✅ 파일 이름이 변경되었습니다!")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"오류: {e}")
        else:
            col2.caption("변경없음")
            
        if col3.button("🗑️ 삭제", key=f"del_btn_{file_path}"):
            try:
                os.remove(file_path)
                st.toast("🗑️ 파일이 삭제되었습니다.")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"오류: {e}")
                
        st.sidebar.markdown("---")
else:
    st.sidebar.info("아직 임시 저장된 파일이 없습니다.")


# ==========================================
# [메인 기능] 파일 업로드 및 변환 흐름 제어
# ==========================================

# 파일 업로드 창
uploaded_file = st.file_uploader("변환할 PDF 파일을 선택하세요", type=["pdf"])

if uploaded_file is not None:
    # 💡 [2번 요청 사항] 업로드된 PDF 파일명을 기본값으로 상속받기 (.pdf 제외)
    pdf_base_name = uploaded_file.name.replace(".pdf", "")
    
    # 만약 새로운 파일을 올렸다면 기존 변환 기억 데이터 리셋
    if st.session_state.current_file_name != uploaded_file.name:
        st.session_state.conversion_done = False
        st.session_state.excel_data_bytes = None
        st.session_state.current_file_name = uploaded_file.name

    # 파일명 지정 입력창 (PDF 파일명이 기본값으로 꽂힙니다)
    custom_file_name = st.text_input(
        "📥 다운로드할 엑셀 파일 이름을 정해주세요:", 
        value=pdf_base_name
    )

    if not custom_file_name.endswith(".xlsx"):
        final_excel_name = custom_file_name + ".xlsx"
    else:
        final_excel_name = custom_file_name

    # 💡 [1번 요청 사항] 변환하기 버튼 생성
    # 아직 변환 전이거나 새로운 파일을 올렸을 때만 변환 버튼을 활성화시켜서 노출합니다.
    if not st.session_state.conversion_done:
        start_conversion = st.button("🚀 엑셀로 변환하기", type="primary")
    else:
        start_conversion = False

    # 작업 상황 알림 영역 설정
    status_text = st.empty()

    # '변환하기' 버튼이 클릭되었을 때만 핵심 로직 진입
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
                        status_text.warning(
                            f"⚙️ 작업 진행 중: **{current_page_idx}** / **{total_pages}** 페이지 분석 중... "
                            f"(현재까지 찾아낸 표: {total_tables_found}개)"
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
            
            # 변환 결과 바이너리를 메모리 세션에 박아두기 (★핵심: 재실행 시 다시 연산 안 함)
            st.session_state.excel_data_bytes = output_buffer.getvalue()
            st.session_state.conversion_done = True # 변환이 끝났음을 메모리에 마킹 (1회 제한)
            progress_bar.empty() # 진행바 지우기
            
            # 서버 임시 폴더에 자동 백업
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            server_save_path = os.path.join(TEMP_DIR, f"{timestamp}_{final_excel_name}")
            with open(server_save_path, "wb") as f:
                f.write(st.session_state.excel_data_bytes)
                
            st.rerun() # 완료된 화면 조성을 위해 페이지 리프레시
            
        except Exception as e:
            status_text.error(f"❌ 변환 도중 오류가 발생했습니다: {e}")
            st.session_state.conversion_done = False

    # 💡 [3번 요청 사항] 변환이 완수된 상태라면, 무한 재연산을 '중단'하고 다운로드 버튼만 고정 노출
    if st.session_state.conversion_done and st.session_state.excel_data_bytes is not None:
        status_text.success("🎉 변환이 성공적으로 완료되었습니다! 아래 버튼을 눌러 저장하세요.")
        st.info("💡 왼쪽 사이드바 목록에도 보관되었습니다. 창을 닫아도 최신 5개 목록에서 다시 받을 수 있습니다.")
        
        st.download_button(
            label=f"📥 {final_excel_name} 다운로드 하기",
            data=st.session_state.excel_data_bytes,
            file_name=final_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
