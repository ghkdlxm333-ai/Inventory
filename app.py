import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="3PL 재고 관리 시스템 (Pro)", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (이미지의 UI 디테일 재현)
st.markdown("""
    <style>
    /* 헤더 텍스트와 필터 아이콘 상시 고정 */
    .ag-header-cell-label { 
        display: flex !important;
        align-items: center !important;
        font-weight: bold !important;
        font-size: 13px !important;
    }
    
    /* 1. 필터 아이콘 (깔때기) */
    .ag-header-icon.ag-filter-icon {
        opacity: 1 !important;
        visibility: visible !important;
        color: #808e9b !important; /* 차분한 회색 */
        margin-left: 8px !important;
    }

    /* 2. 메뉴 아이콘 (점점점 ...) */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        display: flex !important;
        color: #0984e3 !important; /* 강조 파란색 */
        margin-left: 4px !important;
    }

    /* 헤더 열 간 구분선 */
    .ag-header-cell::after {
        content: "|";
        position: absolute;
        right: 0;
        color: #dfe6e9;
        font-weight: normal;
    }

    /* 사이드바 파란색 안내 박스 */
    .info-box {
        background-color: #d1e9ff;
        padding: 12px;
        border-radius: 8px;
        color: #0984e3;
        font-weight: bold;
        font-size: 0.9rem;
        margin-top: 10px;
        border: 1px solid #74b9ff;
    }

    /* 상단 지표 스타일 */
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

# 한글 기간 변환 함수
def get_korean_period(days):
    if days <= 0: return "당일 만료"
    y, m = days // 365, (days % 365) // 30
    if y > 0: return f"약 {y}년 {m}개월 남음" if m > 0 else f"약 {y}년 남음"
    return f"약 {m}개월 남음"

# 3. 데이터 로드 및 전처리
@st.cache_data
def load_and_validate_data(file):
    try:
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        for col in ['가용재고', '불량재고']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류: {e}"

# 4. AgGrid 렌더링
def render_pro_grid(data, threshold, use_filter):
    gb = GridOptionsBuilder.from_dataframe(data)
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=use_filter, 
        menuTabs=['filterMenuTab', 'columnsMenuTab'], # 필터와 컬럼숨김 메뉴 활성화
        minWidth=120, flex=1
    )

    # 특정 컬럼 설정
    gb.configure_column("상품명", minWidth=250, flex=2)
    gb.configure_column("유효일자", cellStyle=JsCode(f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"))

    gb.configure_grid_options(
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        suppressMenuHide=True, # 아이콘 상시 노출
        headerHeight=40,
        enableRangeSelection=True
    )
    return AgGrid(data, gridOptions=gb.build(), height=550, theme='alpine', allow_unsafe_jscode=True)

# 5. 메인 UI
st.title("📦 3PL 재고 관리 시스템 (Pro)")

uploaded_file = st.file_uploader("3PL 엑셀 원본(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바 설정
        st.sidebar.title("⚙️ 관리 설정")
        use_filter = st.sidebar.checkbox("🔍 열별 필터 검색창 표시", value=False)
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 1095, 548)
        
        # [이미지 재현] 설정 기준 안내 박스
        st.sidebar.markdown(f"""
            <div class="info-box">
                💡 설정 기준: {get_korean_period(days_limit)}
            </div>
        """, unsafe_allow_html=True)
        
        # 상단 요약 지표
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 메인 탭
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1: render_pro_grid(master_df[master_df['가용재고']>0], days_limit, use_filter)
        with tab2: render_pro_grid(master_df[master_df['불량재고']>0], days_limit, use_filter)
        with tab3: render_pro_grid(slow_df, days_limit, use_filter)
