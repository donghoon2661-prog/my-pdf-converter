import streamlit as st
import pdfplumber
import pandas as pd
import io

# 1. 웹사이트 타이틀 및 레이아웃 설정
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="wide")
st.title("📄 COA 성적서 PDF ➡️ 엑셀 변환기 (전체 페이지)")
st.write("PDF 파일을 업로드하면 내부의 모든 표(Table)를 감지하여 하나의 통합 엑셀 파일로 변환해 드립니다.")

st.markdown("---")

# 2. 파일명 지정 입력창 만들기 (★추가된 기능)
# 사용자가 확장자(.xlsx)를 빼고 적어도 자동으로 붙도록 안전장치를 해두었습니다.
custom_file_name = st.text_input(
    "📥 다운로드할 엑셀 파일 이름을 정해주세요:", 
    value="COA_Full_Inspection_Report"
)

# 확장자 정제 (.xlsx가 뒤에 붙어있지 않다면 자동으로 추가)
if not custom_file_name.endswith(".xlsx"):
    final_excel_name = custom_file_name + ".xlsx"
else:
    final_excel_name = custom_file_name

# 3. 파일 업로드 창 만들기 (Drag & Drop 지원)
uploaded_file = st.file_uploader("변환할 PDF 파일을 선택하세요", type=["pdf"])

if uploaded_file is not None:
    st.info("파일 분석 중... 잠시만 기다려 주세요.")
    
    # 메모리 버퍼에 엑셀 데이터를 임시 저장하기 위한 객체
    output_buffer = io.BytesIO()
    
    try:
        # 엑셀 파일 생성을 위한 Writer 선언
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # pdfplumber로 업로드된 PDF 가상 파일 읽기
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                st.write(f"📊 총 **{total_pages}개**의 페이지가 감지되었습니다.")
                
                # 진행률 표시 바
                progress_bar = st.progress(0)
                
                # 전 페이지를 돌며 표 구조 추출
                for page_num in range(total_pages):
                    page = pdf.pages[page_num]
                    tables = page.extract_tables()
                    
                    if not tables:
                        continue
                        
                    for i, table in enumerate(tables):
                        df = pd.DataFrame(table)
                        # 완전히 비어있는 행 제거
                        df.dropna(how='all', inplace=True)
                        
                        # 시트 이름 규칙 설정 (예: Page1_Table1)
                        sheet_name = f"Page{page_num + 1}_Table{i + 1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 진행바 업데이트
                    progress_bar.progress((page_num + 1) / total_pages)
        
        # 버퍼에 쌓인 데이터 정렬 후 준비
        excel_data = output_buffer.getvalue()
        
        st.success("🎉 모든 표를 성공적으로 추출하여 엑셀 구조로 변환 완료했습니다!")
        
        # 4. 다운로드 버튼 생성 (★지정한 파일명이 반영됨)
        st.download_button(
            label=f"📥 {final_excel_name} 다운로드 하기",
            data=excel_data,
            file_name=final_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 변환 도중 오류가 발생했습니다: {e}")
