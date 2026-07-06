import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime, date

# ==========================================
# ⚙️ 페이지 및 기본 설정 (Wide Layout & 탭 아이콘)
# ==========================================
st.set_page_config(
    page_title="멘소래담 통합 수주업로드", 
    page_icon="https://raw.githubusercontent.com/paak1010/mentholatum_mart_total/main/logo2.png", 
    layout="wide"
)

# ==========================================
# 🎨 B2B SaaS 스타일 커스텀 CSS (스트림릿 느낌 지우기)
# ==========================================
st.markdown("""
<style>
    /* 폰트 적용 (Pretendard 등 모던 폰트 룩) */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"]  {
        font-family: 'Pretendard', sans-serif !important;
    }

    /* 스트림릿 기본 요소 숨기기 (헤더, 푸터, 햄버거 메뉴) */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* 전체 배경색 지정 (아주 연한 회색으로 모던한 느낌) */
    .stApp {
        background-color: #f8fafc;
    }
    
    /* 사이드바 배경색 및 경계선 튜닝 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* 탭 디자인 변경 (SaaS 스타일) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        padding-bottom: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 12px 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
        font-weight: 500;
        color: #64748b;
    }
    .stTabs [aria-selected="true"] {
        background-color: #f0fdf4;
        border: 1px solid #22c55e;
        font-weight: 700;
        color: #15803d;
        box-shadow: 0 2px 4px rgba(34, 197, 94, 0.1);
    }
    
    /* 메트릭 카드(통계 요약) 디자인 */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    }
    
    /* 다운로드 버튼 스타일링 (그라데이션 & 강조) */
    .stDownloadButton button {
        width: 100%;
        border-radius: 8px;
        font-weight: 700;
        letter-spacing: 0.5px;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 12px 0;
        box-shadow: 0 4px 6px rgba(59, 130, 246, 0.25);
        transition: all 0.3s ease;
    }
    .stDownloadButton button:hover {
        background: linear-gradient(135deg, #4338ca 0%, #2563eb 100%);
        box-shadow: 0 6px 8px rgba(59, 130, 246, 0.35);
        transform: translateY(-1px);
        color: white;
    }
    
    /* 파일 업로더 디자인 개선 */
    [data-testid="stFileUploadDropzone"] {
        border-radius: 12px;
        border: 2px dashed #94a3b8;
        background-color: #ffffff;
        padding: 30px;
        transition: all 0.2s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #3b82f6;
        background-color: #f0f9ff;
    }
</style>
""", unsafe_allow_html=True)

# 모든 날짜 형식을 하이픈 없이 YYYYMMDD로 통일
today_str = datetime.today().strftime("%Y%m%d")

# ==========================================
# 🎨 좌측 사이드바 (Sidebar) - 로고 단독 배치
# ==========================================
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except FileNotFoundError:
        pass
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"📅 **시스템 기준일:** `{today_str}`")

# ==========================================
# 📝 메인 화면 타이틀 (깔끔하게 텍스트만)
# ==========================================
st.title("통합 마트 수주 자동 변환 대시보드")
st.markdown("> **Tesco, 이마트 계열(TRD/노브랜드), 롯데마트**의 수주 데이터를 하나의 표준 양식으로 자동 병합·변환합니다.")
st.markdown("<br>", unsafe_allow_html=True)

# ⭐ 최종 통일 양식 컬럼 리스트 (구분 열 추가)
FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

def to_excel_unified(df, sheet_name="통합_수주업로드"):
    """데이터프레임을 엑셀 파일(메모리)로 변환하고 숫자 서식을 지정합니다."""
    numeric_cols = ['수량', '단가', 'Total Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        num_format = workbook.add_format({'num_format': '#,##0'})
        center_format = workbook.add_format({'align': 'center'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#1e293b', 'font_color': 'white', 'border': 1, 'align': 'center'})
        
        for col_num, value in enumerate(df.columns.
