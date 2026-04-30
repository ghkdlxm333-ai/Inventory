import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="3PL 재고 관리 시스템 (Pro)", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (이미지의 UI를 그대로 재현)
st.markdown("""
    <style>
    /* 헤더 텍스트와 아이콘 배치 최적화 */
    .ag-header-cell-label {
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        font-weight: bold !important;
    }
    
    /* 필터 아이콘 상시 노출 (이미지처럼 목록명 바로 옆) */
    .ag-header-icon.ag-filter-icon {
        opacity: 1 !important;
        visibility: visible !important;
        color: #808e9b !important;
        margin-left: 6px !important;
    }

    /* 메뉴 아이콘(...) 상시 노출 및 배치 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        display: flex !important;
        color: #0984e3 !important;
        margin-left: 4px !important;
    }

    /* 사이드바 설정 기준 문구 디자인 */
    .info-box {
        background-color: #d1e9ff;
        padding: 15px;
        border-radius: 8px;
        color: #0984e3;
        font-weight: bold;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-top: 10px;
    }

    /* 지표 박스 스타일 */
    .metric-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .metric-value { color: #0984e3; font-size: 1.4rem; font-weight: 800; }
    </style>
""", unsafe_allow_html=True)

# 잔여 기간 한글 변환 함수 (이미지 속 "약 1년 6개월 남음" 재현)
def format_period_korean(days):
    if days <= 0: return "만료 임박"
    y = days // 365
    m = (days % 365) // 30
    if y > 0:
        return f"약 {y}년 {m}개월 남음" if m > 0 else f"약 {y}년 남음"
    return f"약 {m}개월 남음"

# 3. 데이터 로드 및 전처리
@st.cache_data
def load_and_validate_data(file):
    try:
        df = pd.read_excel(file, engine='openpyxl', skiprows=1) # 데이터 구조에 따라 조정
        # 이미지에 보이는 컬럼 매핑 (샘플 기반)
        master_df = df.copy()
        # (실제 데이터 컬럼명에 맞춰 수정 필요)
        master_df['유효일자_dt'] = pd.to_datetime(master_df.iloc[:, 3], errors='coerce') 
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("1899-12-29")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['가용재고'] = pd.to_numeric(master_df.iloc[:, 6], errors='coerce').fillna(0).astype(int)
        
        return master_df
    except Exception as e: return f"에러: {e}"

# 4. AgGrid 렌더링 함수
def render_pro_grid(data, threshold):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 필터 및 메뉴 상시 활성화 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=False, # 입력창은 숨기고 아이콘 클릭 시 팝업되게 설정
        suppressMenu=False,   # 메뉴(...) 활성화
        minWidth=120,
        flex=1
    )
    
    # 특정 컬럼 필터 타입 설정 (이미지의 '포함' 필터 팝업 재현)
    gb.configure_column("상품명", filter='agTextColumnFilter', menuTabs=['filterMenuTab', 'generalMenuTab', 'columnsMenuTab'])
    gb.configure_column("상품코드", filter='agTextColumnFilter')
    
    # 그리드 전체 옵션
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        suppressMenuHide=True, # 마우스 오버 전에도 아이콘 상시 표시
        localeText={'filterOoo': '필터...', 'equals': '같음', 'contains': '포함'}
    )

    # 유효일자 임박 색상 (빨간색)
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) return {{'color': '#d63031', 'fontWeight': 'bold'}};
            return null;
        }}
    """))

    return AgGrid(data, gridOptions=gb.build(), height=500, theme='alpine', allow_unsafe_jscode=True)

# 5. 메인 UI 구성
st.title("📦 3PL 재고 관리 시스템 (Pro)")

# 사이드바 관리 설정
st.sidebar.header("⚙️ 관리 설정")
days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 1095, 548)

# [요청사항] 약 0년 0개월 남음 문구 추가
period_info = format_period_korean(days_limit)
st.sidebar.markdown(f"""
    <div class="info-box">
        💡 설정 기준: {period_info}
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("3PL 엑셀 원본(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 상단 지표 영역 (이미지 UI 재현)
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div>✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div>⚠️ 전체 불량재고</div><div class="metric-value">499,724 EA</div></div>', unsafe_allow_html=True) # 예시
        with c3: st.markdown(f'<div class="metric-container"><div>🚨 임박(가용기준)</div><div class="metric-value">{len(master_df[master_df["잔여일수"] <= days_limit])}건</div></div>', unsafe_allow_html=True)

        st.markdown("🔍 **통합 검색 (품목명/코드/LOT)**")
        search_term = st.text_input("", placeholder="검색어를 입력하세요", label_visibility="collapsed")

        # 탭 구성
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1:
            render_pro_grid(master_df, days_limit)
