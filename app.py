import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (아이콘 가시성 최우선 순위 강제)
st.markdown("""
    <style>
    /* CSS 강제성 극대화: 어떤 상황에서도 필터 아이콘 노출 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        display: inline-block !important;
        color: #0984e3 !important;
    }
    /* 필터 팝업 위치 및 디자인 유지 */
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

# 3. AgGrid 렌더링 함수 (아이콘 고정 로직 강화)
def render_styled_aggrid(data, threshold, tab_type):
    # 컬럼 정의 및 순서 고정
    if tab_type == "avail":
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
    elif tab_type == "bad":
        show_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
    else: # exp
        show_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)', '웰로스코드']
    
    display_df = data[show_cols].copy()
    gb = GridOptionsBuilder.from_dataframe(display_df)
    
    # [핵심 1] 모든 컬럼에 대해 메뉴 아이콘 노출 옵션 강제 주입
    gb.configure_default_column(
        resizable=True, 
        sortable=True, 
        filterable=True, 
        floatingFilter=False,
        suppressMenu=False,         # 메뉴 숨김 기능 자체를 끔 (항상 노출)
        menuTabs=['filterMenuTab']  # 클릭 시 필터가 바로 나오도록 고정
    )

    # 유효일자 강조 (JS)
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    grid_options = gb.build()

    # [핵심 2] AgGrid 내부 최적화 기능을 역으로 이용 (아이콘 상시 렌더링)
    grid_options['suppressMenuHide'] = True      # 마우스 오버 전에도 아이콘 노출 (필수)
    grid_options['ensureDomOrder'] = True        # DOM 순서 보장 (아이콘 사라짐 방지)
    grid_options['headerHeight'] = 48            # 헤더 높이를 충분히 주어 아이콘 잘림 방지
    
    grid_options['sideBar'] = {
        'toolPanels': [{
            'id': 'columns', 'labelDefault': '컬럼 관리', 'labelKey': 'columns',
            'iconKey': 'columns', 'toolPanel': 'agColumnsToolPanel'
        }],
        'defaultToolPanel': ''
    }

    return AgGrid(
        display_df,
        gridOptions=grid_options,
        height=550,
        theme='alpine',
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        key=f"fixed_grid_{tab_type}" # 고유 키 유지
    )

# 4. 메인 로직 (기존 전처리 함수 및 실행부 유지)
# ... [load_and_validate_data 함수는 이전과 동일하게 유지] ...

# [실행부 요약]
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])
if uploaded_file:
    # 엑셀 로드 부분은 생략 (이전 코드와 동일하게 사용하세요)
    master_df = load_and_validate_data(uploaded_file) 
    
    if isinstance(master_df, pd.DataFrame):
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1: render_styled_aggrid(master_df[master_df['가용재고']>0], 548, "avail")
        with tab2: render_styled_aggrid(master_df[master_df['불량재고']>0], 548, "bad")
        with tab3:
            st.markdown('<div class="exp-setting-container">', unsafe_allow_html=True)
            days_limit = st.slider("🚨 임박 기준 설정", 30, 1095, 548, key="slider_exp")
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
