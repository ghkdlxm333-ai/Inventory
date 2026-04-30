import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (필터 아이콘 및 레이아웃 유지)
st.markdown("""
    <style>
    .ag-floating-filter { display: none !important; }
    .ag-header-cell-label {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        width: 100% !important;
    }
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        visibility: visible !important;
        color: #0984e3 !important;
    }
    .exp-setting-container {
        background-color: #f1f3f9;
        padding: 15px;
        border-radius: 12px;
        margin-bottom: 15px;
        border-left: 5px solid #ff4d4f;
    }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 설정
AG_GRID_LOCALE_KR = {
    'filterOoo': '필터링 검색...', 'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제',
    'columns': '컬럼 관리', 'sum': '합계', 'count': '개수'
}

# 3. 데이터 로드 및 전처리
@st.cache_data(show_spinner="데이터 분석 중...")
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
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류: {e}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [필터 로직 유지]
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=False, menuTabs=['filterMenuTab'], suppressMenu=False
    )

    # [컬럼 구성 및 순서 조정] - 요청사항 반영
    if tab_type == "avail":
        # 보일 컬럼: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 가용재고
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        hide_cols = ['잔여일수', '셀', '입수량(BOX)']
    
    elif tab_type == "bad":
        # 보일 컬럼: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 불량재고
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        hide_cols = ['잔여일수', '셀', '입수량(BOX)']
        
    else: # exp (임박재고)
        # 보일 컬럼: 상품코드, 상품명, 화주LOT, 유효일자, 가용재고, 잔여일수, 잔여비율(%)
        # 숨길 컬럼: 웰로스코드, 셀, 입수량(BOX)
        show_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)']
        hide_cols = ['웰로스코드', '셀', '입수량(BOX)']

    # 컬럼 가시성 및 순서 적용
    for col in data.columns:
        if col in show_cols:
            # 순서 보장을 위해 index 추출 후 설정
            gb.configure_column(col, hide=False, pinned='left' if col in ['상품코드', '상품명'] else None)
        else:
            gb.configure_column(col, hide=True)

    # 상태바 및 사이드바 (기존 유지)
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        sideBar={"toolPanels": [{"id": "columns", "labelDefault": "컬럼 관리", "toolPanel": "agColumnsToolPanel"}], "defaultToolPanel": ""},
        localeText=AG_GRID_LOCALE_KR,
        suppressMenuHide=True
    )

    # 유효일자 색상 강조
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    return AgGrid(
        data[show_cols + hide_cols], # 정의된 순서대로 데이터 전달
        gridOptions=gb.build(),
        height=600,
        theme='alpine',
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        key=f"grid_{tab_type}"
    )

# 5. 메인 실행부
st.title("📦 3PL 통합 재고관리 Pro")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, pd.DataFrame):
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            render_styled_aggrid(master_df[master_df['가용재고']>0], 548, "avail")
            
        with tab2:
            render_styled_aggrid(master_df[master_df['불량재고']>0], 548, "bad")
            
        with tab3:
            st.markdown('<div class="exp-setting-container">', unsafe_allow_html=True)
            col_l, col_r = st.columns([3, 1])
            with col_l:
                days_limit = st.slider("🚨 임박 기준 설정", 30, 1095, 548, key="slider_exp")
            with col_r:
                st.metric("기준", f"{days_limit}일")
            st.markdown('</div>', unsafe_allow_html=True)
            
            slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
            render_styled_aggrid(slow_df, days_limit, "exp")

        # 수주 분석 섹션
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_upload")
        
        if order_file:
            try:
                order_df = pd.read_excel(order_file)
                if '상품코드' in order_df.columns and '수량' in order_df.columns:
                    order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                    stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                    analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                    analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                    analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                    st.dataframe(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']], use_container_width=True)
            except Exception as e:
                st.warning(f"분석 오류: {e}")
