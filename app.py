import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (아이콘을 목록명 바로 옆에 고정하고 강조)
st.markdown("""
    <style>
    /* 1. 상시 검색창 영역 제거 (공간 확보) */
    .ag-floating-filter { display: none !important; }
    
    /* 2. 헤더 목록명과 아이콘 정렬 */
    .ag-header-cell-label {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        width: 100% !important;
    }

    /* 3. 필터 아이콘(깔때기) 상시 노출 및 파란색 강조 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        visibility: visible !important;
        color: #0984e3 !important;
        cursor: pointer !important;
    }
    
    /* 4. 필터 팝업창 내부 디자인 살짝 수정 */
    .ag-filter-wrapper { padding: 10px !important; }

    /* 지표 카드 디자인 */
    .metric-container {
        background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px;
        padding: 12px 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
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
        
        # 1899년 날짜 오류 방지 처리
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        master_df['가용_Box환산'] = master_df.apply(lambda x: round(x['가용재고']/x['입수량(BOX)'], 1) if x['입수량(BOX)'] > 0 else 0, axis=1)
        master_df['불량_Box환산'] = master_df.apply(lambda x: round(x['불량재고']/x['입수량(BOX)'], 1) if x['입수량(BOX)'] > 0 else 0, axis=1)

        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류: {e}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [가장 중요한 설정] 상시 검색창은 끄고, 아이콘 클릭 시 필터 팝업만 활성화
    gb.configure_default_column(
        resizable=True, 
        sortable=True, 
        filterable=True, 
        floatingFilter=False, # 상시 검색창 OFF
        menuTabs=['filterMenuTab'], # 아이콘 클릭 시 즉시 필터 탭 노출
        suppressMenu=False
    )

    # 개별 컬럼 필터 타입 지정
    for col in data.columns:
        if col in ['가용재고', '불량재고', '가용_Box환산', '불량_Box환산', '입수량(BOX)', '잔여일수']:
            gb.configure_column(col, filter='agNumberColumnFilter')
        else:
            gb.configure_column(col, filter='agTextColumnFilter')

    # 그리드 기능 (드래그 합계 + 컬럼 숨김 사이드바)
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        sideBar={
            "toolPanels": [{
                "id": "columns", "labelDefault": "컬럼 관리", "labelKey": "columns",
                "iconKey": "columns", "toolPanel": "agColumnsToolPanel"
            }],
            "defaultToolPanel": ""
        },
        localeText=AG_GRID_LOCALE_KR,
        suppressMenuHide=True, # 아이콘 상시 노출
        headerHeight=45
    )
    
    # 탭별 컬럼 구성 및 순서
    if tab_type == "avail":
        active = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
    elif tab_type == "bad":
        active = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
    else: # exp
        active = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)']

    for col in data.columns:
        if col in active:
            gb.configure_column(col, hide=False, pinned='left' if col in ['상품코드', '상품명'] else None)
        else:
            gb.configure_column(col, hide=True)

    # 유효일자 강조 및 프로그레스 바
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    return AgGrid(
        data, 
        gridOptions=gb.build(), 
        height=600, 
        theme='alpine', 
        allow_unsafe_jscode=True, 
        enable_enterprise_modules=True
    )

# 5. 메인 실행부
st.title("📦 3PL 통합 재고관리 Pro")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 상단 요약 지표
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= 548)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 탭 출력
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1: render_styled_aggrid(master_df[master_df['가용재고']>0], 548, "avail")
        with tab2: render_styled_aggrid(master_df[master_df['불량재고']>0], 548, "bad")
        with tab3: render_styled_aggrid(slow_df, 548, "exp")

        # 수주 가용성 실시간 분석
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'])
        
        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                
                analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                
                # 분석 결과 출력
                st.table(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']])
            except Exception as e:
                st.warning(f"분석 오류: {e}")
