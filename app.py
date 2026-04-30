import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (아이콘 가시성 확보 및 하단 상태바)
st.markdown("""
    <style>
    .metric-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 12px 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; margin-bottom: 3px; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    
    /* [핵심] 필터 아이콘(☰)을 항상 파란색으로 표시하고 강조 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: inline-block !important; 
        visibility: visible !important; 
        color: #0984e3 !important;
        margin-right: 5px !important;
    }
    .ag-header-icon { color: #0984e3 !important; }
    
    /* 하단 상태바 강조 */
    .ag-status-bar {
        background-color: #f8f9fa !important;
        font-weight: bold;
        color: #0984e3;
        min-height: 35px !important;
        border-top: 1px solid #dee2e6 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 한국어 언어팩
AG_GRID_LOCALE_KR = {
    'filterOoo': '조회...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함', 
    'startsWith': '시작값', 'endsWith': '끝값', 'sum': '선택 합계', 'avg': '평균', 'count': '개수',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제', 'blank': '비어있음', 'notBlank': '내용있음'
}

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
        master_df['잔여비율'] = (master_df['잔여일수'] / 730 * 100).clip(0, 100).fillna(0).astype(int)
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        master_df['가용_Box환산'] = master_df.apply(lambda r: f"{r['가용재고']//r['입수량(BOX)']}B+{r['가용재고']%r['입수량(BOX)']}E" if r['입수량(BOX)']>0 else f"{r['가용재고']}E", axis=1)
        
        search_target_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT']
        master_df['_search_idx'] = master_df[search_target_cols].apply(lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1)
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류: {e}"

def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)

    # [핵심] 개별 목록 필터 아이콘 설정
    gb.configure_default_column(
        resizable=True, 
        sortable=True, 
        filterable=True,      # 개별 필터 활성화
        floatingFilter=False,  # 입력창 상시 노출 끔 (아이콘 클릭 방식)
        suppressMenuHide=False, # 메뉴 아이콘(☰)을 항상 노출하도록 강제
        menuTabs=['filterMenuTab'], # 아이콘 클릭 시 즉시 필터 탭 노출
        minWidth=110,
        flex=1
    )

    # 수량 컬럼 합계 및 필터 타입 지정
    if "가용재고" in data.columns:
        gb.configure_column("가용재고", aggFunc='sum', filter='agNumberColumnFilter', type=["numericColumn"])
    if "불량재고" in data.columns:
        gb.configure_column("불량재고", aggFunc='sum', filter='agNumberColumnFilter', type=["numericColumn"])

    # 텍스트 필터 명시 (상품코드, 상품명, LOT 등)
    text_filter_cols = ['상품코드', '상품명', '화주LOT', '웰로스코드']
    for col in text_filter_cols:
        if col in data.columns:
            gb.configure_column(col, filter='agTextColumnFilter')

    # 그리드 옵션 설정 (드래그 합계 및 페이지네이션 제거)
    gb.configure_grid_options(
        enableRangeSelection=True,
        pagination=False, 
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        localeText=AG_GRID_LOCALE_KR,
        allow_unsafe_jscode=True
    )

    return AgGrid(
        data, 
        gridOptions=gb.build(), 
        height=600, 
        theme='alpine', 
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True, 
        update_mode=GridUpdateMode.MODEL_CHANGED
    )

# --- 메인 로직 ---
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 730, 548)
        
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(master_df[(master_df["가용재고"] > 0) & (master_df["잔여일수"] <= days_limit)])}건</div></div>', unsafe_allow_html=True)

        search_input = st.text_input("🔍 빠른 통합 검색", placeholder="검색어 입력...").strip()
        filtered_df = master_df.copy()
        if search_input:
            filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(search_input.lower(), na=False)]

        tab1, tab2 = st.tabs(["✅ 가용재고 현황", "⚠️ 불량재고 현황"])
        with tab1:
            render_styled_aggrid(filtered_df[filtered_df['가용재고']>0], days_limit, "avail")
        with tab2:
            render_styled_aggrid(filtered_df[filtered_df['불량재고']>0], days_limit, "normal")
