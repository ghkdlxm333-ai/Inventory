import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 디자인 커스텀 CSS (필터 가시성 극대화)
st.markdown("""
    <style>
    /* 필터 입력창 배경색 및 글자색 강조 */
    .ag-floating-filter-input {
        background-color: #ffffff !important;
        border: 1px solid #0984e3 !important;
        border-radius: 4px !important;
    }
    /* 필터 아이콘(☰) 강제 표시 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: inline-block !important; 
        visibility: visible !important; 
        color: #0984e3 !important;
    }
    /* 상태바 합계창 강조 */
    .ag-status-bar {
        background-color: #f8f9fa !important;
        font-weight: bold;
        color: #0984e3;
        border-top: 1px solid #dee2e6 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 데이터 로딩 함수 (기존 로직 유지)
@st.cache_data
def load_data(file):
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
        master_df['잔여일수'] = (master_df['유효일자_dt'] - datetime.now()).dt.days.fillna(0).astype(int)
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        return master_df
    except Exception as e: return str(e)

def render_grid(data):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심 변경] 모든 컬럼에 개별 필터 및 상단 입력창(Floating Filter) 적용
    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filterable=True,      # 필터 활성화
        floatingFilter=True,   # 컬럼 헤더 바로 아래에 필터 입력창 상시 노출
        suppressMenuHide=False, # 메뉴 아이콘 숨김 방지
        minWidth=120
    )
    
    # 수량 컬럼은 숫자형 필터 적용
    gb.configure_column("가용재고", type=["numericColumn"], filter='agNumberColumnFilter', aggFunc='sum')
    gb.configure_column("불량재고", type=["numericColumn"], filter='agNumberColumnFilter', aggFunc='sum')
    
    # 그리드 설정 (페이지네이션 제거, 드래그 합계 활성화)
    gb.configure_grid_options(
        enableRangeSelection=True,
        pagination=False,      # 하단 페이지창 제거
        statusBar={
            "statusPanels": [
                { "statusPanel": "agAggregationComponent", "align": "right" } # 드래그 합계
            ]
        },
        localeText={
            'filterOoo': '조회...', 
            'equals': '같음', 
            'contains': '포함',
            'sum': '선택 합계',
            'avg': '평균',
            'count': '개수'
        }
    )
    
    return AgGrid(
        data,
        gridOptions=gb.build(),
        height=600,
        theme='alpine',
        enable_enterprise_modules=True, # 합계 기능을 위해 필요
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.MODEL_CHANGED
    )

# 메인 UI
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])
if uploaded_file:
    df = load_data(uploaded_file)
    if isinstance(df, str):
        st.error(df)
    else:
        tab1, tab2 = st.tabs(["✅ 재고 목록", "🚨 임박 분석"])
        with tab1:
            render_grid(df)
