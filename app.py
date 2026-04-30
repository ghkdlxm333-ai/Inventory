import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS
st.markdown("""
    <style>
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        display: inline-block !important;
        color: #0984e3 !important;
    }
    .ag-floating-filter { display: none !important; }
    .ag-header-cell-label {
        display: flex !important;
        justify-content: space-between !important;
        width: 100% !important;
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
    'filterOoo': '필터링 검색...', 'applyFilter': '적용', 'resetFilter': '초기화', 
    'clearFilter': '해제', 'columns': '컬럼 관리'
}

# 3. 데이터 로드 및 전처리 함수 (생략되었던 부분 복구)
@st.cache_data(show_spinner="데이터 분석 중...")
def load_and_validate_data(file):
    try:
        # 데이터 구조 파악을 위한 샘플링 로드
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        
        # 실제 데이터 로드
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        
        # 필요한 컬럼 인덱스 추출 (데이터 위치가 고정적이라는 가정)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 날짜 변환 및 잔여일 계산
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"데이터 로드 중 오류 발생: {e}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, tab_type):
    # 탭별 노출 컬럼 고정
    if tab_type == "avail":
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
    elif tab_type == "bad":
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
    else: # exp (임박재고)
        show_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)', '웰로스코드']
    
    display_df = data[show_cols].copy()
    gb = GridOptionsBuilder.from_dataframe(display_df)
    
    # 기본 컬럼 설정 (필터 아이콘 강제 노출 포함)
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=False, suppressMenu=False, menuTabs=['filterMenuTab']
    )

    # 상품명 등 고정
    gb.configure_column("상품코드", pinned='left')
    gb.configure_column("상품명", pinned='left')

    # 유효일자 임박 시 빨간색 강조 (JS)
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    grid_options = gb.build()
    grid_options['suppressMenuHide'] = True  # 상시 노출
    grid_options['ensureDomOrder'] = True
    grid_options['headerHeight'] = 48
    grid_options['localeText'] = AG_GRID_LOCALE_KR

    return AgGrid(
        display_df,
        gridOptions=grid_options,
        height=500,
        theme='alpine',
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        key=f"fixed_grid_{tab_type}"
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
                days_limit = st.slider("🚨 임박 기준 설정(일)", 30, 1095, 548, key="slider_exp")
            with col_r:
                st.metric("기준일수", f"{days_limit}일")
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
                else:
                    st.warning("수주서 양식이 잘못되었습니다. '상품코드'와 '수량' 컬럼이 필요합니다.")
            except Exception as e:
                st.error(f"수주 분석 중 오류: {e}")
    else:
        st.error(master_df) # load_and_validate_data에서 반환된 에러 메시지 출력
