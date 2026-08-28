# -*- coding: utf-8 -*-

import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import glob
import re
import time
import shutil
import tempfile
from pathlib import Path
from datetime import datetime
from urllib.parse import unquote

import requests

# Selenium은 유안타 자동 수집 탭에서만 사용합니다.
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service


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

    hr {
        margin: 1.2rem 0;
        border: none;
        border-top: 1px solid #EDEFF3;
    }

    .stFileUploader label {
        font-weight: 600;
        font-size: 0.95rem;
    }

    .share-row {
        border: 1px solid #EDEFF3;
        border-radius: 12px;
        padding: 0.9rem 1.1rem;
        margin-bottom: 0.7rem;
        background: #FFFFFF;
    }

    .yuanta-step {
        display: flex;
        align-items: center;
        gap: 0.8rem;
        padding: 0.75rem 0.2rem;
        border-bottom: 1px solid #F0F1F4;
    }
    .yuanta-step:last-child {
        border-bottom: none;
    }
    .yuanta-dot {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        background: #CBD5E1;
        flex: 0 0 auto;
    }
    .yuanta-note {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 10px;
        padding: 0.9rem 1rem;
        color: #166534;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. 헤더
# ==========================================
st.markdown("""
<div class="main-header">
    <h1>📄 COA 성적서 PDF → 엑셀 변환기</h1>
    <p>PDF의 모든 표를 통합 엑셀로 변환하고, 문서 파일을 링크로 공유할 수 있습니다.</p>
</div>
""", unsafe_allow_html=True)


# ==========================================
# 4. 공통 상수 / 유틸
# ==========================================
TEMP_DIR = "temporary_excel_storage"
SHARE_DIR = "shared_file_storage"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(SHARE_DIR, exist_ok=True)

MAX_KEEP = 5
MAX_FILE_MB = 50
MAX_TOTAL_MB = 300

ALLOWED_EXT = [
    "pdf", "hwp", "hwpx", "doc", "docx",
    "ppt", "pptx", "xls", "xlsx", "csv",
    "txt", "zip", "png", "jpg", "jpeg",
]

MIME_MAP = {
    "pdf":  "application/pdf",
    "hwp":  "application/x-hwp",
    "hwpx": "application/hwp+zip",
    "doc":  "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "ppt":  "application/vnd.ms-powerpoint",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xls":  "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "csv":  "text/csv",
    "txt":  "text/plain",
    "zip":  "application/zip",
    "png":  "image/png",
    "jpg":  "image/jpeg",
    "jpeg": "image/jpeg",
}

ICON_MAP = {
    "pdf": "📕", "hwp": "📘", "hwpx": "📘",
    "doc": "📄", "docx": "📄",
    "ppt": "📙", "pptx": "📙",
    "xls": "📗", "xlsx": "📗", "csv": "📗",
    "txt": "📃", "zip": "🗜️",
    "png": "🖼️", "jpg": "🖼️", "jpeg": "🖼️",
}

YUANTA_BASE_URL = (
    "https://www.myasset.com/myasset/research/rs_list/"
    "rs_list.cmd?cd006=&cd007=RE02&cd008="
)
YUANTA_AUTHOR = "황규원"


def human_size(nbytes):
    size = float(nbytes)
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def storage_used(directory):
    total = 0
    for path in glob.glob(os.path.join(directory, "*")):
        if os.path.isfile(path):
            total += os.path.getsize(path)
    return total


def split_stamp(base_name, sep):
    """'20260729123000{sep}파일명.ext' -> (timestamp_str, 순수파일명)"""
    stamp, found, pure = base_name.partition(sep)
    if not found:
        return "", base_name
    return stamp, pure


def parse_dt(stamp, fallback_path):
    try:
        return datetime.strptime(stamp, "%Y%m%d%H%M%S")
    except (ValueError, TypeError):
        return datetime.fromtimestamp(os.path.getmtime(fallback_path))


def safe_filename(filename):
    filename = re.sub(r'[\\/:*?"<>|]', "_", filename)
    filename = re.sub(r"\s+", " ", filename).strip()
    return filename[:180]


# ==========================================
# 5. 유안타 자동 수집 공통 함수
# ==========================================
def create_yuanta_driver(download_dir):
    """
    Windows/로컬 및 Streamlit 배포 환경을 모두 고려한
    Headless Chrome/Chromium 드라이버 생성 함수.
    """
    options = webdriver.ChromeOptions()

    prefs = {
        "download.default_directory": str(Path(download_dir).resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--remote-allow-origins=*")

    # 일반적인 Linux/Streamlit 배포 환경의 Chromium 경로 자동 탐색
    chrome_candidates = [
        os.environ.get("CHROME_BIN"),
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
    ]

    for chrome_path in chrome_candidates:
        if chrome_path and os.path.exists(chrome_path):
            options.binary_location = chrome_path
            break

    driver_candidates = [
        os.environ.get("CHROMEDRIVER_PATH"),
        "/usr/bin/chromedriver",
        "/usr/local/bin/chromedriver",
    ]

    for driver_path in driver_candidates:
        if driver_path and os.path.exists(driver_path):
            return webdriver.Chrome(
                service=Service(driver_path),
                options=options,
            )

    # 로컬 PC의 Selenium Manager 또는 시스템 PATH에 있는 드라이버 사용
    return webdriver.Chrome(options=options)


def wait_for_pdf_download(download_dir, timeout=90):
    start = time.time()
    download_dir = Path(download_dir)

    while time.time() - start < timeout:
        partial = list(download_dir.glob("*.crdownload"))
        pdfs = list(download_dir.glob("*.pdf"))

        if pdfs and not partial:
            latest = max(pdfs, key=lambda p: p.stat().st_mtime)
            if latest.stat().st_size > 0:
                return latest

        time.sleep(1)

    raise TimeoutException("PDF 다운로드 시간이 초과되었습니다.")


def find_board_search_controls(driver):
    """
    상단 통합검색이 아니라 '산업분석 게시판 내부 검색 영역'을 찾습니다.
    '작성자' 옵션을 포함한 select를 기준으로 검색창과 검색 버튼을 찾습니다.
    """
    search_select = None

    for element in driver.find_elements(By.TAG_NAME, "select"):
        try:
            select = Select(element)
            option_texts = [o.text.strip() for o in select.options]
            if any("작성자" in text for text in option_texts):
                search_select = element
                break
        except Exception:
            continue

    if search_select is None:
        raise RuntimeError("'작성자' 옵션이 있는 게시판 검색 조건을 찾지 못했습니다.")

    select = Select(search_select)
    for option in select.options:
        if "작성자" in option.text.strip():
            select.select_by_visible_text(option.text)
            break

    select_y = search_select.location["y"]
    search_input = None

    for element in driver.find_elements(By.CSS_SELECTOR, "input[type='text']"):
        try:
            if element.is_displayed() and abs(element.location["y"] - select_y) < 180:
                search_input = element
                break
        except Exception:
            continue

    if search_input is None:
        raise RuntimeError("게시판 검색창을 찾지 못했습니다.")

    candidates = []

    for button in driver.find_elements(
        By.CSS_SELECTOR,
        "button, input[type='button'], input[type='submit'], a"
    ):
        try:
            if not button.is_displayed():
                continue

            text = (
                button.text
                or button.get_attribute("value")
                or button.get_attribute("title")
                or button.get_attribute("aria-label")
                or ""
            ).strip()

            if "검색" in text and abs(button.location["y"] - select_y) < 180:
                candidates.append(button)
        except Exception:
            continue

    if not candidates:
        raise RuntimeError("게시판 검색 버튼을 찾지 못했습니다.")

    search_button = min(
        candidates,
        key=lambda e: abs(e.location["y"] - select_y)
    )

    return search_input, search_button


def search_yuanta_author(driver, author):
    search_input, search_button = find_board_search_controls(driver)

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        search_input
    )

    search_input.clear()
    search_input.send_keys(author)

    try:
        search_button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", search_button)

    time.sleep(3)


def find_latest_yuanta_report(driver):
    """
    검색 결과 중 최신 목록 순서를 우선 유지하면서
    정유/화학 + 업체별 이슈를 가장 우선적으로 선택합니다.
    """
    candidates = []

    for index, link in enumerate(driver.find_elements(By.TAG_NAME, "a")):
        try:
            if not link.is_displayed():
                continue

            title = link.text.strip()
            if not title:
                continue

            title_compact = title.replace(" ", "")

            if "정유/화학" in title and "업체별 이슈" in title:
                priority = 1
            elif "정유/화학" in title and "업체별이슈" in title_compact:
                priority = 1
            elif "정유/화학" in title:
                priority = 2
            elif "정유" in title and "화학" in title:
                priority = 3
            else:
                continue

            candidates.append((priority, index, link, title))

        except StaleElementReferenceException:
            continue

    if not candidates:
        raise RuntimeError(
            "황규원 연구원의 최신 정유/화학 관련 게시물을 찾지 못했습니다."
        )

    candidates.sort(key=lambda x: (x[0], x[1]))
    _, _, link, title = candidates[0]
    return link, title


def open_yuanta_report(driver, target_link):
    before_handles = set(driver.window_handles)
    before_url = driver.current_url

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        target_link
    )

    try:
        target_link.click()
    except Exception:
        driver.execute_script("arguments[0].click();", target_link)

    time.sleep(2)

    new_handles = set(driver.window_handles) - before_handles

    if new_handles:
        driver.switch_to.window(new_handles.pop())

    try:
        WebDriverWait(driver, 12).until(
            lambda d: d.current_url != before_url or len(d.window_handles) > 1
        )
    except TimeoutException:
        pass

    time.sleep(2)


def click_yuanta_pdf_download(driver):
    """
    상세 화면에서 PDF/첨부 다운로드 버튼 또는 링크를 탐색합니다.
    """
    candidates = []

    for element in driver.find_elements(
        By.CSS_SELECTOR,
        "a, button, input[type='button'], input[type='submit']"
    ):
        try:
            if not element.is_displayed():
                continue

            text = (
                element.text
                or element.get_attribute("value")
                or element.get_attribute("title")
                or ""
            )

            href = element.get_attribute("href") or ""
            onclick = element.get_attribute("onclick") or ""

            data = f"{text} {href} {onclick}".lower()

            score = 99
            if ".pdf" in data:
                score = 1
            elif "pdf" in data:
                score = 2
            elif "첨부" in text:
                score = 3
            elif "download" in data:
                score = 4

            if score < 99:
                candidates.append((score, element))

        except Exception:
            continue

    if not candidates:
        raise RuntimeError("상세 화면에서 PDF 다운로드 버튼을 찾지 못했습니다.")

    candidates.sort(key=lambda x: x[0])
    button = candidates[0][1]

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        button
    )

    try:
        button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", button)


def extract_excel_urls_from_pdf(pdf_path):
    """
    PDF 텍스트와 PDF 내부 링크(annotation)에서
    .xlsx / .xls URL을 모두 찾습니다.
    """
    urls = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # 일반 텍스트에 URL이 인쇄된 경우
            urls.extend(
                re.findall(r'https?://[^\s<>"\']+', text, flags=re.I)
            )

            # PDF 하이퍼링크(annotation)인 경우
            try:
                for annot in (page.annots or []):
                    uri = annot.get("uri")
                    if uri:
                        urls.append(str(uri))
            except Exception:
                pass

    result = []
    seen = set()

    for url in urls:
        url = url.rstrip(".,;:)]}>")
        decoded = unquote(url).lower()

        if ".xlsx" in decoded or ".xls" in decoded:
            if url not in seen:
                seen.add(url)
                result.append(url)

    return result


def download_excel_bytes(excel_url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        ),
        "Referer": "https://www.myasset.com/",
    }

    response = requests.get(
        excel_url,
        headers=headers,
        timeout=90,
        allow_redirects=True,
    )
    response.raise_for_status()

    if not response.content:
        raise RuntimeError("엑셀 파일이 비어 있습니다.")

    return response.content


def download_latest_yuanta_report(status_box=None, progress_bar=None):
    """
    Streamlit용 유안타 자동 수집 메인 함수.

    반환값:
        excel_bytes, excel_filename, report_title
    """
    driver = None
    temp_root = Path(tempfile.mkdtemp(prefix="yuanta_streamlit_"))
    pdf_dir = temp_root / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    def update_status(message, progress=None):
        if status_box is not None:
            status_box.info(message)
        if progress_bar is not None and progress is not None:
            progress_bar.progress(progress)

    try:
        update_status("⏳ STEP 1/5 · 유안타증권 산업분석 게시판에 접속 중입니다...", 10)

        driver = create_yuanta_driver(pdf_dir)
        driver.get(YUANTA_BASE_URL)

        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(2)

        update_status("🔎 STEP 1/5 · 작성자를 '황규원'으로 검색하고 있습니다...", 20)
        search_yuanta_author(driver, YUANTA_AUTHOR)

        update_status("📋 STEP 2/5 · 최신 정유/화학 업체별 이슈 리포트를 찾고 있습니다...", 40)
        target_link, report_title = find_latest_yuanta_report(driver)

        open_yuanta_report(driver, target_link)

        update_status("📕 STEP 3/5 · 상세 화면에서 PDF를 다운로드하고 있습니다...", 60)
        click_yuanta_pdf_download(driver)

        pdf_file = wait_for_pdf_download(pdf_dir, timeout=90)

        update_status("🔗 STEP 4/5 · PDF 내부의 엑셀 파일 링크를 분석하고 있습니다...", 80)
        excel_urls = extract_excel_urls_from_pdf(pdf_file)

        if not excel_urls:
            raise RuntimeError(
                "PDF 안에서 .xlsx 또는 .xls 링크를 찾지 못했습니다."
            )

        update_status("📗 STEP 5/5 · 원본 엑셀 파일을 다운로드하고 있습니다...", 92)
        excel_bytes = download_excel_bytes(excel_urls[0])

        # 사용자 요청 양식: 260724_원재료 시황자료.xlsx
        excel_filename = (
            f"{datetime.now().strftime('%y%m%d')}_원재료 시황자료.xlsx"
        )

        update_status("🎉 모든 작업이 완료되었습니다.", 100)

        return excel_bytes, excel_filename, report_title

    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

        shutil.rmtree(temp_root, ignore_errors=True)


# ==========================================
# 6. 세션 상태 초기화
# ==========================================
if "conversion_done" not in st.session_state:
    st.session_state.conversion_done = False

if "excel_data_bytes" not in st.session_state:
    st.session_state.excel_data_bytes = None

if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = ""

if "yuanta_excel_bytes" not in st.session_state:
    st.session_state.yuanta_excel_bytes = None

if "yuanta_excel_name" not in st.session_state:
    st.session_state.yuanta_excel_name = ""

if "yuanta_report_title" not in st.session_state:
    st.session_state.yuanta_report_title = ""


# ==========================================
# 7. 사이드바 — 최근 변환 파일
# ==========================================
with st.sidebar:
    st.markdown("### 📁 최근 변환 파일")
    st.caption(f"최대 {MAX_KEEP}개까지 보관됩니다. 앱이 재시작되면 사라집니다.")
    st.markdown("---")

    existing_files = glob.glob(os.path.join(TEMP_DIR, "*.xlsx"))
    existing_files.sort(key=os.path.getmtime, reverse=True)

    if len(existing_files) > MAX_KEEP:
        for old_file in existing_files[MAX_KEEP:]:
            try:
                os.remove(old_file)
            except Exception:
                pass
        existing_files = existing_files[:MAX_KEEP]

    if existing_files:
        for file_path in existing_files:
            file_name = os.path.basename(file_path)
            timestamp_prefix, pure_name = split_stamp(file_name, "__")

            st.markdown('<div class="sidebar-file-card">', unsafe_allow_html=True)
            st.markdown(f"**📗 {pure_name}**")

            pure_name_no_ext = (
                pure_name[:-5] if pure_name.endswith(".xlsx") else pure_name
            )

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
                "💾",
                data=file_bytes,
                file_name=pure_name,
                mime=MIME_MAP["xlsx"],
                key=f"dl_btn_{file_path}",
                help="다운로드",
            )

            if new_name_input != pure_name_no_ext and new_name_input.strip():
                if c2.button(
                    "✅",
                    key=f"save_btn_{file_path}",
                    help="이름 저장"
                ):
                    clean = os.path.basename(
                        new_name_input.strip()
                    ).replace("__", "_")

                    new_full_name = (
                        f"{timestamp_prefix}__{clean}.xlsx"
                        if timestamp_prefix
                        else f"{clean}.xlsx"
                    )

                    try:
                        os.rename(
                            file_path,
                            os.path.join(TEMP_DIR, new_full_name)
                        )
                        st.toast("✅ 이름이 변경되었습니다!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

            if c3.button(
                "🗑️",
                key=f"del_btn_{file_path}",
                help="삭제"
            ):
                try:
                    os.remove(file_path)
                    st.toast("🗑️ 삭제되었습니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"오류: {e}")

            st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.info("아직 변환된 파일이 없습니다.")


# ==========================================
# 8. 탭 구성
# ==========================================
tab_convert, tab_share, tab_yuanta = st.tabs([
    "🔄 PDF → 엑셀 변환",
    "🔗 파일 공유함",
    "📊 유안타 시황자료",
])


# ------------------------------------------
# 8-1. 변환 탭 (기존 기능 유지)
# ------------------------------------------
with tab_convert:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "변환할 PDF 파일을 선택하세요",
        type=["pdf"],
        key="pdf_uploader"
    )

    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        pdf_base_name = (
            uploaded_file.name[:-4]
            if uploaded_file.name.lower().endswith(".pdf")
            else uploaded_file.name
        )

        if st.session_state.current_file_name != uploaded_file.name:
            st.session_state.conversion_done = False
            st.session_state.excel_data_bytes = None
            st.session_state.current_file_name = uploaded_file.name

        st.markdown('<div class="card">', unsafe_allow_html=True)

        col_name, col_btn = st.columns([3, 1])

        with col_name:
            custom_file_name = st.text_input(
                "📥 다운로드할 엑셀 파일 이름",
                value=pdf_base_name
            )

        final_excel_name = (
            custom_file_name
            if custom_file_name.endswith(".xlsx")
            else custom_file_name + ".xlsx"
        )

        with col_btn:
            st.write("")
            st.write("")

            if not st.session_state.conversion_done:
                start_conversion = st.button(
                    "🚀 변환하기",
                    type="primary",
                    use_container_width=True
                )
            else:
                start_conversion = False
                st.button(
                    "✅ 완료됨",
                    disabled=True,
                    use_container_width=True
                )

        status_text = st.empty()
        st.markdown('</div>', unsafe_allow_html=True)

        if start_conversion and not st.session_state.conversion_done:
            progress_bar = st.progress(0)
            status_text.info("⏳ PDF 파일을 분석하고 있습니다...")
            output_buffer = io.BytesIO()

            try:
                total_tables_found = 0

                with pd.ExcelWriter(
                    output_buffer,
                    engine="openpyxl"
                ) as writer:
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

                            if tables:
                                for i, table in enumerate(tables):
                                    df = pd.DataFrame(table)
                                    df.dropna(how="all", inplace=True)

                                    sheet_name = (
                                        f"Page{current_page_idx}_Table{i + 1}"
                                    )

                                    df.to_excel(
                                        writer,
                                        sheet_name=sheet_name,
                                        index=False,
                                        header=False
                                    )

                                    total_tables_found += 1

                            progress_bar.progress(
                                current_page_idx / total_pages
                            )

                    if total_tables_found == 0:
                        pd.DataFrame(
                            {"결과": ["추출된 표가 없습니다."]}
                        ).to_excel(
                            writer,
                            sheet_name="NoTable",
                            index=False
                        )

                st.session_state.excel_data_bytes = output_buffer.getvalue()
                st.session_state.conversion_done = True
                progress_bar.empty()

                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                safe_excel_name = os.path.basename(
                    final_excel_name
                ).replace("__", "_")

                server_save_path = os.path.join(
                    TEMP_DIR,
                    f"{timestamp}__{safe_excel_name}"
                )

                with open(server_save_path, "wb") as f:
                    f.write(st.session_state.excel_data_bytes)

                if total_tables_found == 0:
                    st.warning(
                        "⚠️ PDF에서 표를 찾지 못했습니다. "
                        "스캔 이미지 PDF일 수 있습니다."
                    )

                st.rerun()

            except Exception as e:
                status_text.error(
                    f"❌ 변환 도중 오류가 발생했습니다: {e}"
                )
                st.session_state.conversion_done = False

        if (
            st.session_state.conversion_done
            and st.session_state.excel_data_bytes is not None
        ):
            st.markdown('<div class="card">', unsafe_allow_html=True)

            st.success("🎉 변환이 완료되었습니다!")
            st.caption("왼쪽 사이드바에도 보관됩니다. (앱 재시작 시 삭제)")

            st.download_button(
                label=f"📥 {final_excel_name} 다운로드",
                data=st.session_state.excel_data_bytes,
                file_name=final_excel_name,
                mime=MIME_MAP["xlsx"],
                type="primary",
            )

            st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="text-align:center; padding: 3rem 1rem; color: #9AA1AC;">
            <p style="font-size:0.9rem;">
                PDF 파일을 업로드하면 변환을 시작할 수 있어요.
            </p>
        </div>
        """, unsafe_allow_html=True)


# ------------------------------------------
# 8-2. 공유함 탭 (기존 기능 유지)
# ------------------------------------------
with tab_share:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    st.markdown("#### 📤 파일 업로드")
    st.caption(
        f"업로드한 파일은 **이 주소에 접속하는 누구나** 내려받을 수 있습니다. "
        f"파일당 최대 {MAX_FILE_MB}MB · 전체 {MAX_TOTAL_MB}MB · 앱 재시작 시 삭제됩니다."
    )

    used = storage_used(SHARE_DIR)
    limit_bytes = MAX_TOTAL_MB * 1024 * 1024

    st.progress(min(used / limit_bytes, 1.0))
    st.caption(f"사용량 {human_size(used)} / {MAX_TOTAL_MB}.0 MB")

    share_upload = st.file_uploader(
        "공유할 파일 선택 (여러 개 가능)",
        type=ALLOWED_EXT,
        key="share_uploader",
        accept_multiple_files=True,
    )

    if share_upload:
        if st.button("📤 공유함에 올리기", type="primary"):
            success_count = 0

            for f in share_upload:
                size = f.size

                if size > MAX_FILE_MB * 1024 * 1024:
                    st.error(
                        f"❌ {f.name} — {human_size(size)}. "
                        f"파일당 {MAX_FILE_MB}MB를 초과합니다."
                    )
                    continue

                if storage_used(SHARE_DIR) + size > limit_bytes:
                    st.error(
                        f"❌ {f.name} — 저장소 용량이 부족합니다. "
                        "기존 파일을 정리해 주세요."
                    )
                    continue

                safe_name = os.path.basename(
                    f.name
                ).replace("__", "_")

                ts = datetime.now().strftime("%Y%m%d%H%M%S")

                dest = os.path.join(
                    SHARE_DIR,
                    f"{ts}__{safe_name}"
                )

                try:
                    with open(dest, "wb") as out:
                        out.write(f.getbuffer())
                    success_count += 1

                except Exception as e:
                    st.error(f"❌ {f.name} 저장 실패: {e}")

            if success_count:
                st.toast(
                    f"✅ {success_count}개 파일이 업로드되었습니다."
                )

            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("#### 📂 공유 중인 파일")

    shared = [
        p for p in glob.glob(os.path.join(SHARE_DIR, "*"))
        if os.path.isfile(p)
    ]

    shared.sort(key=os.path.getmtime, reverse=True)

    if not shared:
        st.info(
            "아직 공유된 파일이 없습니다. 위에서 파일을 올려보세요."
        )

    else:
        st.caption(f"총 {len(shared)}개")

        for path in shared:
            base = os.path.basename(path)
            stamp, pure = split_stamp(base, "__")
            uploaded_at = parse_dt(stamp, path)

            ext = (
                pure.rsplit(".", 1)[-1].lower()
                if "." in pure
                else ""
            )

            size = os.path.getsize(path)
            icon = ICON_MAP.get(ext, "📁")

            with st.expander(
                f"{icon}  {pure}  ·  {human_size(size)}"
            ):
                st.caption(
                    f"업로드: {uploaded_at.strftime('%Y-%m-%d %H:%M')}"
                )

                c1, c2 = st.columns([1, 1])

                with open(path, "rb") as fh:
                    c1.download_button(
                        "📥 다운로드",
                        data=fh.read(),
                        file_name=pure,
                        mime=MIME_MAP.get(
                            ext,
                            "application/octet-stream"
                        ),
                        key=f"share_dl_{path}",
                        use_container_width=True,
                    )

                if c2.button(
                    "🗑️ 삭제",
                    key=f"share_del_{path}",
                    use_container_width=True
                ):
                    try:
                        os.remove(path)
                        st.toast("🗑️ 삭제되었습니다.")
                        st.rerun()

                    except Exception as e:
                        st.error(f"오류: {e}")


# ------------------------------------------
# 8-3. 유안타 시황자료 자동 수집 (신규 기능)
# ------------------------------------------
with tab_yuanta:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    st.markdown("### 📊 유안타증권 원재료 시황자료 자동 수집")

    st.caption(
        "황규원 연구원의 최신 정유/화학 업체별 이슈 리포트를 자동으로 찾아 "
        "PDF를 분석한 뒤 원본 엑셀 파일을 가져옵니다."
    )

    st.markdown("---")

    metric1, metric2 = st.columns(2)

    with metric1:
        st.metric("👤 작성자", "황규원")

    with metric2:
        st.metric("📋 대상", "최신 정유/화학 시황")

    start_yuanta = st.button(
        "🚀 최신 시황자료 가져오기",
        type="primary",
        use_container_width=True,
        key="start_yuanta_download",
    )

    st.markdown('</div>', unsafe_allow_html=True)

    # 진행 상태 목업 / 실제 상태 표시 영역
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("#### 📋 진행 상태")

    step_placeholder = st.empty()

    step_placeholder.markdown("""
    <div class="yuanta-step">
        <span class="yuanta-dot"></span>
        <b>STEP 1</b>&nbsp;&nbsp;유안타증권 접속 및 작성자 검색
        <span style="margin-left:auto;color:#64748B;">대기 중</span>
    </div>
    <div class="yuanta-step">
        <span class="yuanta-dot"></span>
        <b>STEP 2</b>&nbsp;&nbsp;최신 정유/화학 리포트 선택
        <span style="margin-left:auto;color:#64748B;">대기 중</span>
    </div>
    <div class="yuanta-step">
        <span class="yuanta-dot"></span>
        <b>STEP 3</b>&nbsp;&nbsp;PDF 다운로드
        <span style="margin-left:auto;color:#64748B;">대기 중</span>
    </div>
    <div class="yuanta-step">
        <span class="yuanta-dot"></span>
        <b>STEP 4</b>&nbsp;&nbsp;PDF 내부 엑셀 링크 추출
        <span style="margin-left:auto;color:#64748B;">대기 중</span>
    </div>
    <div class="yuanta-step">
        <span class="yuanta-dot"></span>
        <b>STEP 5</b>&nbsp;&nbsp;원본 엑셀 파일 다운로드
        <span style="margin-left:auto;color:#64748B;">대기 중</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if start_yuanta:
        status_box = st.empty()
        progress_bar = st.progress(0)

        try:
            excel_bytes, excel_name, report_title = (
                download_latest_yuanta_report(
                    status_box=status_box,
                    progress_bar=progress_bar,
                )
            )

            st.session_state.yuanta_excel_bytes = excel_bytes
            st.session_state.yuanta_excel_name = excel_name
            st.session_state.yuanta_report_title = report_title

            st.success("🎉 최신 시황자료를 성공적으로 가져왔습니다.")
            st.caption(f"선택된 리포트: {report_title}")

        except Exception as e:
            progress_bar.empty()
            status_box.error(
                f"❌ 자동 수집 중 오류가 발생했습니다.\n\n"
                f"상세 내용: {e}"
            )

    if st.session_state.yuanta_excel_bytes is not None:
        st.markdown('<div class="card">', unsafe_allow_html=True)

        st.success("📗 원재료 시황자료 다운로드 준비 완료")

        if st.session_state.yuanta_report_title:
            st.caption(
                f"원본 리포트: {st.session_state.yuanta_report_title}"
            )

        st.download_button(
            label=f"📥 {st.session_state.yuanta_excel_name} 다운로드",
            data=st.session_state.yuanta_excel_bytes,
            file_name=st.session_state.yuanta_excel_name,
            mime=MIME_MAP["xlsx"],
            type="primary",
            use_container_width=True,
            key="yuanta_excel_download",
        )

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="yuanta-note">
        🛡️ 이 기능은 유안타증권 게시판에서 공개적으로 접근 가능한 리포트를
        자동으로 검색하고, 해당 리포트 안에 포함된 원본 엑셀 링크를 내려받는 용도입니다.
    </div>
    """, unsafe_allow_html=True)
