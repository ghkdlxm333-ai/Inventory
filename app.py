import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 및 레이아웃 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (헤더 아이콘 상시 노출 및 엑셀 스타일 UI)
st.markdown("""
    <style>
    /* 헤더 텍스트와 필터 아이콘 배치 */
    .ag-header-cell-label { 
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        font-weight: bold !important;
        font-size: 13px !important;
    }
    /* 필터 아이콘 스타일: 제목 바로 옆에 파란색 강조 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        margin-left: 6px !important;
        color: #0984e3 !important;
        visibility: visible !important;
    }
    .ag-header-icon { color: #0984e3 !important; }
    
    /* 상단 지표(Metric) 디자인 */
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
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 언어 팩
AG_GRID_LOCALE_KR = {
    'filterOoo': '검색...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제',
    'noRowsToShow': '표시할 데이터가 없습니다', 'pinColumn': '열 고정', 'export': '내보내기',
    'sum': '합계', 'min': '최소', 'max': '최대', 'avg': '평균', 'count': '개수'
}

# 3. 데이터 로드 및 전처리 로직 (기능 유지)
@st.cache_data(show_spinner="데이터를 분석하고 있습니다...")
def load_and_validate_data(file):
    try:
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        
        # 원본 데이터 매핑 (기존 인덱스 유지)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 수치형 데이터 정제 및 Box 환산
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        safe_box = master_df['입수량(BOX)'].replace(0, 1)
        master_df['가용_Box환산'] = (master_df['가용재고'] / safe_box).round(2)
        master_df['불량_Box환산'] = (master_df['불량재고'] / safe_box).round(2)
        
        # 날짜 및 유효기한 계산
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"오류 발생: {e}"

# 4. AgGrid 렌더링 함수 (기능 통합: 탭별 컬럼 제어 + 드래그 합계 + 상시 필터)
def render_tab_grid(data, tab_type, threshold):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심 설정 1] 필터/검색창 상시 활성화 및 엑셀 스타일 메뉴
    gb.configure_default_column(
        resizable=True, 
        sortable=True, 
        filterable=True, 
        floatingFilter=True,  # 검색창 상시 노출 (버튼 관계없이 적용)
        menuTabs=['filterMenuTab'], 
        suppressMenuHide=True, # 아이콘 항상 표시
        minWidth=110,
        flex=1
    )
    
    # [핵심 설정 2] 드래그 합계 기능 (StatusBar 복구)
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={
            "statusPanels": [
                {"statusPanel": "agTotalAndFilteredRowCountComponent", "align": "left"},
                {"statusPanel": "agAggregationComponent", "align": "right"}
            ]
        },
        localeText=AG_GRID_LOCALE_KR
    )

    # [핵심 설정 3] 탭별 컬럼 가시성 및 순서 요구사항 적용
    all_cols = data.columns.tolist()
    
    if tab_type == "avail": # 1. 가용재고 탭
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        hide_cols = ['가용_Box환산', '잔여일수', '셀', '입수량(BOX)']
        remove_cols = ['잔여비율(%)', '불량재고', '불량_Box환산']
        
    elif tab_type == "bad": # 2. 불량재고 탭
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        hide_cols = ['불량_Box환산', '잔여일수', '셀', '입수량(BOX)']
        remove_cols = ['잔여비율(%)', '가용재고', '가용_Box환산']
        
    elif tab_type == "exp": # 3. 임박재고 탭
        display_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)', '웰로스코드']
        hide_cols = ['웰로스코드', '셀', '입수량(BOX)']
        remove_cols = ['가용_Box환산', '불량_Box환산', '불량재고']

    # 컬럼 설정 루프
    for col in all_cols:
        if col in display_cols:
            gb.configure_column(col, hide=False)
            # 수치 데이터 필터 및 집계 설정
            if col in ['가용재고', '불량재고', '잔여일수']:
                gb.configure_column(col, type=["numericColumn", "numberColumnFilter"], aggFunc='sum')
            else:
                gb.configure_column(col, filter='agTextColumnFilter')
        elif col in hide_cols:
            gb.configure_column(col, hide=True)
        else:
            gb.configure_column(col, hide=True) # 불필요 컬럼 완전 숨김

    # 고정 열 설정
    gb.configure_column("상품코드", pinned='left', width=130, cellStyle={'textAlign': 'center'})
    gb.configure_column("상품명", pinned='left', width=250, flex=2)

    # 임박재고 탭 전용 잔여비율 프로그레스 바
    if tab_type == "exp":
        percent_renderer = JsCode("""
        class PercentBarRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let val = params.value;
                let color = val <= 20 ? '#ff4d4f' : (val <= 50 ? '#faad14' : '#52c41a');
                this.eGui.innerHTML = `<div style="width:100%; background:#f0f0f0; border-radius:4px; height:16px; margin-top:8px; border:1px solid #ddd;">
                    <div style="width:${val}%; background:${color}; height:100%; border-radius:4px;"></div>
                </div>`;
            }
            getGui() { return this.eGui; }
        }
        """)
        gb.configure_column("잔여비율(%)", cellRenderer=percent_renderer, width=150)

    # 유효일자 조건부 서식 (기능 유지)
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) return {{'color': 'red', 'fontWeight': 'bold'}};
            return null;
        }}
    """))

    return AgGrid(
        data, 
        gridOptions=gb.build(), 
        height=600, 
        theme='alpine', 
        allow_unsafe_jscode=True,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.VALUE_CHANGED
    )

# 5. 메인 UI 및 실행부
st.title("📦 3PL 통합 재고 관리 시스템")
uploaded_file = st.file_uploader("3PL 엑셀 원본(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str):
        st.error(master_df)
    else:
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 1095, 548)
        
        # 상단 지표 영역
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 탭 구성 및 요구사항별 그리드 출력
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            render_tab_grid(master_df[master_df['가용재고'] > 0], "avail", days_limit)
        with tab2:
            render_tab_grid(master_df[master_df['불량재고'] > 0], "bad", days_limit)
        with tab3:
            render_tab_grid(slow_df, "exp", days_limit)

        # 6. 수주 가용성 분석 (기능 유지)
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_analysis")
        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                
                # 가독성을 위한 상품명 결합
                names = master_df[['상품코드', '상품명']].drop_duplicates('상품코드')
                analysis = pd.merge(analysis, names, on='상품코드', how='left')
                
                st.table(analysis[['출고판단', '상품코드', '상품명', '수주요청량', '현재고', '부족수량']])
            except Exception as e:
                st.warning(f"분석 오류: 수주서 형식을 확인하세요. ({e})")
