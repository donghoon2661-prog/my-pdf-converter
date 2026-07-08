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
st.write("PDF 파일을 업로드하면 내부의 모든 표(Table)를 감지하여 하나의 통합 엑셀 파일로 변환해 드립니다.")

st.markdown("---")

# 임시 저장용 폴더 설정 (Streamlit 서버 내부 폴더)
TEMP_DIR = "temporary_excel_storage"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# ==========================================
# [사이드바 기능] 최근 파일 목록 (수정 / 삭제 대시보드)
# ==========================================
st.sidebar.title("📁 최근 변환된 파일 관리 (최대 5개)")
st.sidebar.write("최근 파일의 이름을 수정하거나 서버에서 영구 삭제할 수 있습니다.")

# 폴더 내 xlsx 파일 목록 가져오기 (최신순 정렬)
existing_files = glob.glob(os.path.join(TEMP_DIR, "*.xlsx"))
existing_files.sort(key=os.path.getmtime, reverse=True)

# 5개 유지보수 용량 관리
if len(existing_files) > 5:
    for old_file in existing_files[5:]:
        try: os.remove(old_file)
        except: pass
    existing_files = existing_files[:5]

# 사이드바 리스트 표출 및 개별 제어판 구축
if existing_files:
    for idx, file_path in enumerate(existing_files):
        file_name = os.path.basename(file_path)
        
        # 파일명에서 날짜 표기(타임스탬프)와 본래 파일명 분리
        # 예: 20231024_보고서.xlsx -> 타임스탬프: 20231024_, 이름: 보고서.xlsx
        if "_" in file_name:
            timestamp_prefix, pure_name = file_name.split("_", 1)
        else:
            timestamp_prefix, pure_name = "", file_name
            
        # UI 구분을 위한 실선
        st.sidebar.markdown(f"**📄 파일 {idx+1}**")
        
        # 1. 이름 수정 입력창 (확장자 제외하고 보여줌)
        pure_name_no_ext = pure_name.replace(".xlsx", "")
        new_name_input = st.sidebar.text_input(
            "✏️ 이름 수정:", 
            value=pure_name_no_ext, 
            key=f"rename_input_{file_path}"
        )
        
        # 다운로드 버튼 및 기능 버튼들을 한 줄에 배치하기 위한 가로 레이아웃(컬럼)
        col1, col2, col3 = st.sidebar.columns([2, 1, 1])
        
        # 파일 바이너리 미리 읽기
        with open(file_path, "rb") as f:
            file_bytes = f.read()
            
        # (버튼 1) 다운로드 버튼
        col1.download_button(
            label="💾 다운로드",
            data=file_bytes,
            file_name=pure_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_btn_{file_path}"
        )
        
        # (버튼 2) 이름 변경 저장 버튼
        # 사용자가 입력창에 적은 내용이 기존 파일명과 다를 때만 작동하도록 유도
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
            
        # (버튼 3) 파일 삭제 버튼
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
# [메인 기능] 파일 변환 및 실시간 알림 로직
# ==========================================

# 2. 파일명 지정 입력창 만들기
custom_file_name = st.text_input(
    "📥 다운로드할 엑셀 파일 이름을 정해주세요:", 
    value="COA_Full_Inspection_Report"
)

if not custom_file_name.endswith(".xlsx"):
    final_excel_name = custom_file_name + ".xlsx"
else:
    final_excel_name = custom_file_name

# 3. 파일 업로드 창 만들기
uploaded_file = st.file_uploader("변환할 PDF 파일을 선택하세요", type=["pdf"])

if uploaded_file is not None:
    status_text = st.empty()
    progress_bar = st.progress(0)
    
    status_text.info("⏳ PDF 파일을 읽어오는 중입니다...")
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
        
        excel_data = output_buffer.getvalue()
        
        # 서버에 임시 백업 파일 쓰기 (중복 방지용 타임스탬프 결합)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        server_save_path = os.path.join(TEMP_DIR, f"{timestamp}_{final_excel_name}")
        with open(server_save_path, "wb") as f:
            f.write(excel_data)
            
        status_text.success(
            f"🎉 변환 완료! 총 **{total_pages}페이지**를 스캔하여 **{total_tables_found}개**의 표를 성공적으로 추출했습니다."
        )
        st.info("💡 왼쪽 사이드바 목록에 안전하게 임시 보관되었습니다. 이름 수정 및 삭제가 가능합니다.")
        
        # 4. 다운로드 버튼 생성
        st.download_button(
            label=f"📥 {final_excel_name} 바로 다운로드 하기",
            data=excel_data,
            file_name=final_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 변환 성공 후 사이드바 메뉴 갱신을 위한 리런
        st.rerun()
        
    except Exception as e:
        status_text.error(f"❌ 변환 도중 오류가 발생했습니다: {e}")
