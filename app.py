import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 및 레이아웃 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (엑셀 스타일 UI 강화)
st.markdown("""
    <style>
    /* 헤더 텍스트와 필터 아이콘 배치 */
    .ag-header-cell-label { 
        display: flex !important;
        align-items: center !important;
        font-weight: bold !important;
    }
    /* 필터 아이콘 상시 노출 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        margin-left: 6px !important;
        color: #0984e3 !important;
        visibility: visible !important;
    }
    /* 상단 지표 박스 */
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

# 3. 데이터 로드 및 전처리 로직
@st.cache_data(show_spinner="데이터를 분석하고 있습니다...")
def load_and_validate_data(file):
    try:
        # 헤더 위치 자동 찾기
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        
        # 컬럼 매핑 (인덱스 기준 추출)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 수치형 변환 및 박스 환산 계산
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        # 입수량이 0인 경우를 방지하여 계산
        safe_box = master_df['입수량(BOX)'].replace(0, 1)
        master_df['가용_Box환산'] = (master_df['가용재고'] / safe_box).round(2)
        master_df['불량_Box환산'] = (master_df['불량재고'] / safe_box).round(2)
        
        # 날짜 및 잔여일수 계산
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"오류 발생: {e}"

# 4. AgGrid 렌더링 엔진 (요구사항 반영)
def render_tab_grid(data, tab_type, threshold):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [공통] 필터 아이콘/검색창 상시 활성화 및 드래그 합계 기능
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=True, # 모든 열 아래 검색창 상시 노출
        menuTabs=['filterMenuTab'], suppressMenuHide=True,
        minWidth=100
    )
    
    # [드래그 합계 기능] 상태바 활성화
    gb.configure_grid_options(
        enableRangeSelection=True, # 마우스 드래그 활성화
        statusBar={
            "statusPanels": [
                {"statusPanel": "agTotalAndFilteredRowCountComponent", "align": "left"},
                {"statusPanel": "agAggregationComponent", "align": "right"} # 드래그 시 합계 표시
            ]
        },
        localeText={
            'filterOoo': '검색...', 'applyFilter': '적용', 'resetFilter': '초기화',
            'sum': '합계', 'min': '최소', 'max': '최대', 'avg': '평균', 'count': '개수'
        }
    )

    # 5. 탭별 컬럼 가시성 및 순서 설정
    all_cols = data.columns.tolist()
    
    if tab_type == "avail": # 1. 가용재고 탭
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        hide_cols = ['가용_Box환산', '잔여일수', '셀', '입수량(BOX)']
        # 그 외 불량 관련 및 잔여비율은 아예 제거(숨김)
        
    elif tab_type == "bad": # 2. 불량재고 탭
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        hide_cols = ['불량_Box환산', '잔여일수', '셀', '입수량(BOX)']
        
    elif tab_type == "exp": # 3. 임박재고 탭
        display_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)', '웰로스코드']
        hide_cols = ['웰로스코드', '셀', '입수량(BOX)']

    # 컬럼 설정 적용
    for col in all_cols:
        if col in display_cols:
            gb.configure_column(col, hide=False)
            # 숫자 열은 필터 타입을 숫자로 변경하고 합계 기능 부여
            if col in ['가용재고', '불량재고', '잔여일수', '입수량(BOX)']:
                gb.configure_column(col, type=["numericColumn", "numberColumnFilter"], aggFunc='sum')
        elif col in hide_cols:
            gb.configure_column(col, hide=True) # 우측 컬럼 메뉴에서 다시 꺼낼 수 있음
        else:
            gb.configure_column(col, hide=True) # 목록에서 아예 제외

    # 임박재고 탭 전용 잔여비율 프로그레스 바 (가시성)
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

    # 유효일자 조건부 서식 (임박 시 빨간색)
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold} && params.data.잔여일수 > 0) return {{'color': 'red', 'fontWeight': 'bold'}};
            if (params.data.잔여일수 <= 0) return {{'color': 'white', 'backgroundColor': '#d63031'}};
            return null;
        }}
    """))

    return AgGrid(
        data, 
        gridOptions=gb.build(), 
        height=600, 
        theme='alpine', 
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED
    )

# 6. 메인 화면 구성
st.title("📦 3PL 통합 재고 관리 시스템")
uploaded_file = st.file_uploader("재고 원본 파일(.xlsx)을 업로드하세요", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    
    if isinstance(master_df, str):
        st.error(f"데이터를 읽는 중 오류가 발생했습니다: {master_df}")
    else:
        # 사이드바 설정 (기준일만 조절)
        st.sidebar.header("⚙️ 관리 기준 설정")
        days_limit = st.sidebar.slider("🚨 유효기한 임박 기준(일)", 30, 1095, 548)
        
        # 상단 요약 지표
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박/경과 품목</div><div class="metric-value">{len(master_df[master_df["잔여일수"] <= days_limit])}건</div></div>', unsafe_allow_html=True)

        # 탭별 화면 출력
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고 현황", "⚠️ 불량재고 현황", "🚨 임박재고 관리"])
        
        with tab1:
            render_tab_grid(master_df[master_df['가용재고'] > 0], "avail", days_limit)
            
        with tab2:
            render_tab_grid(master_df[master_df['불량재고'] > 0], "bad", days_limit)
            
        with tab3:
            # 가용재고가 있으면서 임박한 데이터만 추출
            exp_data = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
            render_tab_grid(exp_data, "exp", days_limit)
