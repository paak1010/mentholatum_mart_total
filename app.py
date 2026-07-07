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

    /* 💡 수정 1: 헤더 전체를 숨기지 않고, 메뉴와 Deploy 버튼만 숨겨서 사이드바 토글 버튼 살리기 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    
    /* 전체 배경색 지정 (아주 연한 회색으로 모던한 느낌) */
    .stApp {
        background-color: #f8fafc;
    }
    
    /* 사이드바 배경색 및 경계선 튜닝 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* 💡 수정 2: 탭 디자인을 자연스러운 하단 밑줄(Underline) 스타일로 변경 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
        border-bottom: 1px solid #e2e8f0; /* 탭 아래 연한 회색 선 */
        padding-bottom: 0px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border: none;
        border-radius: 0;
        padding: 12px 4px;
        font-weight: 500;
        color: #64748b;
        box-shadow: none;
    }
    /* 선택된 탭 자연스럽게 강조 */
    .stTabs [aria-selected="true"] {
        background-color: transparent;
        border: none;
        border-bottom: 3px solid #3b82f6; /* 파란색 밑줄 포인트 */
        font-weight: 700;
        color: #1e293b;
        box-shadow: none;
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
