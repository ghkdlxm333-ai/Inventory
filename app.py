import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (필터 아이콘 및 사이드바 시각화 강제 고정)
st.markdown("""
    <style>
    /* 필터 아이콘 상시 노출 및 색상 고정 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        color: #0984e3 !important;
        display: block !important;
    }
    /* 사이드바 버튼 영역 확보 */
    .ag-side-bar {
        border-left: 1px solid #ddd !important;
    }
    /* 임박재고 탭 내부 설정 박스 */
    .exp-setting-container {
        background-color: #f1f3f9;
        padding: 15px;
        border-radius: 12px;
        margin-bottom: 15px;
        border-left: 5px solid #ff4d4f;
    }
    </style>
""", unsafe_allow_html=True)

# 3. 데이터 로드 및 전처리 함수
@st.cache_data(show_spinner="데이터 분석 중...")
def load_and_validate_data(file):
    try:
        # 데이터 로드 (헤더 행 자동 찾기)
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        
        # 원본 인덱스 기준 필수 열 추출
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 날짜 및 숫자 전처리
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        # 잔여 기간 계산
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"오류 발생: {e}"

# 4. AgGrid 렌더링 함수 (사용자 선호 로직 유지)
def render_scm_grid(data, threshold, tab_name):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # 기본 컬럼 설정 (필터 및 메뉴 강제 활성화)
    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filterable=True,
        menuTabs=['filterMenuTab', 'generalMenuTab', 'columnsMenuTab'],
        suppressMenu=False
    )

    # 잔여비율 프로그레스 바 (JS)
    percent_renderer = JsCode("""
    function(params) {
        let val = params.value;
        let color = val <= 20 ? '#ff4d4f' : (val <= 50 ? '#ffa940' : '#40a9ff');
        return `
            <div style="width:100%; background-color:#e0e0e0; border-radius:10px; height:18px; position:relative; overflow:hidden; margin-top:6px;">
                <div style="width:${val}%; background-color:${color}; height:100%;"></div>
                <span style="position:absolute; width:100%; text-align:center; top:0; left:0; font-size:11px; font-weight:800; color:#222; line-height:18px;">${val}%</span>
            </div>`;
    }
    """)

    # 탭별 컬럼 구성
    if tab_name == "exp":
        active_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)']
        gb.configure_column("잔여비율(%)", cellRenderer=percent_renderer, minWidth=150)
    else:
        active_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고', '불량재고']

    for col in data.columns:
        if col in active_cols:
            gb.configure_column(col, hide=False, pinned='left' if col in ['상품코드', '상품명'] else None)
        else:
            gb.configure_column(col, hide=True)

    # 유효일자 강조 (JS)
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    # 사이드바 및 엔터프라이즈 옵션 설정
    grid_options = gb.build()
    grid_options['sideBar'] = {
        'toolPanels': [
            {'id': 'columns', 'labelDefault': '컬럼 설정', 'labelKey': 'columns', 'iconKey': 'columns', 'toolPanel': 'agColumnsToolPanel'},
            {'id': 'filters', 'labelDefault': '전체 필터', 'labelKey': 'filters', 'iconKey': 'filter', 'toolPanel': 'agFiltersToolPanel'}
        ],
        'defaultToolPanel': ''
    }

    return AgGrid(
        data,
        gridOptions=grid_options,
        height=550,
        theme='alpine',
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        key=f"grid_{tab_name}",
        reload_data=False
    )

# 5. 메인 로직 실행
st.title("📦 3PL 통합 재고관리 Pro")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    
    if isinstance(master_df, str):
        st.error(master_df)
    else:
        # 상단 요약 지표 삭제됨 (요청 사항)
        
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            render_scm_grid(master_df[master_df['가용재고'] > 0], 548, "avail")
            
        with tab2:
            render_scm_grid(master_df[master_df['불량재고'] > 0], 548, "bad")
            
        with tab3:
            # 임박재고 탭 전용 설정 섹션
            st.markdown('<div class="exp-setting-container">', unsafe_allow_html=True)
            col_l, col_r = st.columns([3, 1])
            with col_l:
                days_limit = st.slider("🚨 임박 기준(일) 설정", 30, 1095, 548, key="exp_slider_unique")
            with col_r:
                st.metric("설정 기준", f"{days_limit}일")
            st.markdown('</div>', unsafe_allow_html=True)
            
            filtered_exp = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
            render_scm_grid(filtered_exp, days_limit, "exp")

        # 6. 수주 가용성 분석 섹션 (하단 유지)
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
                    
                    # 분석 결과 출력
                    st.dataframe(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']], use_container_width=True)
                else:
                    st.warning("수주서 엑셀에 '상품코드'와 '수량' 컬럼이 반드시 포함되어야 합니다.")
            except Exception as e:
                st.error(f"분석 중 오류 발생: {e}")
