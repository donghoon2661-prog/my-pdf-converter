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
# [사이드바 기능] 최근 생성된 파일 5개 목록 관리
# ==========================================
st.sidebar.title("📁 최근 변환된 파일 (최대 5개)")
st.sidebar.write("최근에 생성된 파일들은 창을 닫았다가 다시 켜도 아래에서 바로 재다운로드할 수 있습니다.")

existing_files = glob.glob(os.path.join(TEMP_DIR, "*.xlsx"))
existing_files.sort(key=os.path.getmtime, reverse=True)

if len(existing_files) > 5:
    for old_file in existing_files[5:]:
        try:
            os.remove(old_file)
        except:
            pass
    existing_files = existing_files[:5]

if existing_files:
    for file_path in existing_files:
        file_name = os.path.basename(file_path)
        display_name = file_name.split("_", 1)[-1] if "_" in file_name else file_name
        
        with open(file_path, "rb") as f:
            st.sidebar.download_button(
                label=f"💾 {display_name}",
                data=f.read(),
                file_name=display_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=file_path
            )
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
    # 실시간 작업 상황 텍스트를 갈아끼우기 위한 임시 플레이스홀더 생성 (★추가)
    status_text = st.empty()
    progress_bar = st.progress(0)
    
    status_text.info("⏳ PDF 파일을 읽어오는 중입니다...")
    
    output_buffer = io.BytesIO()
    
    try:
        total_tables_found = 0 # 총 찾아낸 표 개수 카운트용
        
        # 엑셀 파일 생성을 위한 Writer 선언
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                
                # 전 페이지를 돌며 표 구조 추출
                for page_num in range(total_pages):
                    current_page_idx = page_num + 1
                    
                    # 💡 여기에 사용자가 요청한 실시간 문자열 알림을 계속 업데이트해 줍니다! (★추가)
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
                        total_tables_found += 1 # 표 개수 추가
                    
                    # 진행바 업데이트
                    progress_bar.progress(current_page_idx / total_pages)
        
        excel_data = output_buffer.getvalue()
        
        # 서버에 임시 백업 파일 쓰기
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        server_save_path = os.path.join(TEMP_DIR, f"{timestamp}_{final_excel_name}")
        with open(server_save_path, "wb") as f:
            f.write(excel_data)
            
        # 💡 작업이 끝나면 완료 메시지로 깔끔하게 덮어쓰기! (★추가)
        status_text.success(
            f"🎉 변환 완료! 총 **{total_pages}페이지**를 스캔하여 **{total_tables_found}개**의 표를 성공적으로 추출했습니다."
        )
        st.info("💡 사이드바의 '최근 변환된 파일' 목록에도 안전하게 임시 보관되었습니다.")
        
        # 4. 다운로드 버튼 생성
        st.download_button(
            label=f"📥 {final_excel_name} 바로 다운로드 하기",
            data=excel_data,
            file_name=final_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if st.button("🔄 최근 파일 목록 새로고침 하기"):
            st.rerun()
        
    except Exception as e:
        status_text.error(f"❌ 변환 도중 오류가 발생했습니다: {e}")
