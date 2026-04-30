import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# [GitHub 배포 필수] 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 디자인 커스텀 CSS (필터 아이콘 시인성 강화)
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
    
    /* AgGrid 헤더 및 필터 아이콘 스타일링 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: block !important; 
        visibility: visible !important;
        color: #0984e3 !important; 
    }
    .ag-header-icon { color: #0984e3 !important; }
    .ag-floating-bottom { background-color: #f8f9fa !important; font-weight: bold !important; }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 언어 팩
AG_GRID_LOCALE_KR = {
    'pivotMode': '피벗 모드', 'columns': '열', 'filters': '필터', 'valueColumns': '값 열',
    'pivotColumns': '피벗 열', 'groups': '그룹', 'filterOoo': '필터...', 'equals': '같음',
    'notEqual': '같지 않음', 'empty': '비어 있음', 'lessThan': '작음', 'greaterThan': '큼',
    'lessThanOrEqual': '작거나 같음', 'greaterThanOrEqual': '크거나 같음', 'inRange': '범위 내',
    'contains': '포함', 'notContains': '포함하지 않음', 'startsWith': '시작값', 'endsWith': '끝값',
    'andCondition': '그리고', 'orCondition': '또는', 'applyFilter': '필터 적용', 'resetFilter': '필터 초기화',
    'clearFilter': '필터 해제', 'noRowsToShow': '표시할 데이터가 없습니다', 'pinColumn': '열 고정',
    'export': '내보내기', 'csvExport': 'CSV로 내보내기', 'excelExport': 'Excel로 내보내기',
}

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
        
        # 원본 인덱스 기반 컬럼 추출
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 전처리
        master_df['웰로스코드'] = master_df['웰로스코드'].astype(str).str.strip().replace('nan', '')
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 730 * 100).clip(0, 100).fillna(0).astype(int)
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        master_df['_search_idx'] = master_df[['상품코드', '상품명', '웰로스코드', '화주LOT']].apply(
            lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1
        )
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류 발생: {e}"

def render_styled_aggrid(data, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 모든 컬럼에 필터 아이콘 상시 노출 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        minWidth=100, flex=1, suppressMenu=False,
        menuTabs=['filterMenuTab']
    )
    
    # 틀 고정
    gb.configure_column("상품코드", pinned='left', width=130)
    gb.configure_column("상품명", pinned='left', width=250, flex=2)

    # 탭별 컬럼 정의 (요청하신 규칙 적용)
    all_cols = data.columns.tolist()
    
    if tab_type == "avail":
        # 표시: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 가용재고
        visible = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        for col in all_cols:
            if col not in visible: gb.configure_column(col, hide=True)
            
    elif tab_type == "bad":
        # 표시: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 불량재고
        visible = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        for col in all_cols:
            if col not in visible: gb.configure_column(col, hide=True)
            
    elif tab_type == "exp":
        # 표시: 상품코드, 상품명, 화주LOT, 유효일자, 가용재고, 잔여일수, 잔여비율(%), 웰로스코드
        # 특이사항: 웰로스코드는 '숨겨야 할 컬럼'에 포함되어 있어 hide 처리
        visible = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율']
        for col in all_cols:
            if col not in visible: gb.configure_column(col, hide=True)
        gb.configure_column("웰로스코드", hide=True) # 보이지만 숨김(데이터는 포함)
        
        # 잔여비율 시각화
        percent_renderer = JsCode("""
        class PercentBarRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let val = params.value;
                let color = val <= 20 ? '#ff9f43' : (val <= 50 ? '#feca57' : '#d1e9ff');
                this.eGui.innerHTML = `<div style="width:100%; background:#f1f1f1; border-radius:10px; height:18px; border:1px solid #ddd; margin-top:6px;">
                    <div style="width:${val}%; background:${color}; text-align:center; font-size:11px; font-weight:bold; height:100%; line-height:18px;">${val}%</div>
                </div>`;
            }
            getGui() { return this.eGui; }
        }
        """)
        gb.configure_column("잔여비율", headerName="잔여비율(%)", cellRenderer=percent_renderer, minWidth=140)

    # 유효일자 강조 (JsCode 내 변수 전달을 위해 f-string 사용)
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) {{
                return {{'color': 'white', 'backgroundColor': '#d63031', 'fontWeight': 'bold'}};
            }}
            return null;
        }}
    """))

    # 그리드 옵션 (상태바 합계 기능 활성화)
    go = gb.build()
    go['enableRangeSelection'] = True
    go['statusBar'] = {"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]}
    go['localeText'] = AG_GRID_LOCALE_KR
    
    return AgGrid(data, gridOptions=go, height=550, theme='alpine', allow_unsafe_jscode=True)

# --- 메인 영역 ---
st.title("📦 3PL 통합 재고 관리 시스템")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바: 임박 기준 및 시간 환산
        st.sidebar.title("⚙️ 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준일", 30, 730, 548)
        years = days_limit // 365
        months = (days_limit % 365) // 30
        st.sidebar.write(f"⏱️ 기준: **{years}년 {months}개월** ({days_limit}일)")

        # 상단 지표
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 가용재고 합계</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 불량재고 합계</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박 대상</div><div class="metric-value">{len(slow_df)} 건</div></div>', unsafe_allow_html=True)

        search_q = st.text_input("🔍 품목명 또는 코드로 검색", placeholder="검색어를 입력하세요...")
        view_df = master_df.copy()
        if search_q:
            view_df = view_df[view_df['_search_idx'].str.contains(search_q.lower())]

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            st.caption("※ 상품코드, 명, 웰로스, LOT, 유효일자, 가용재고만 표시됩니다.")
            render_styled_aggrid(view_df[view_df['가용재고']>0], days_limit, "avail")
            
        with tab2:
            st.caption("※ 상품코드, 명, 웰로스, LOT, 유효일자, 불량재고만 표시됩니다.")
            render_styled_aggrid(view_df[view_df['불량재고']>0], days_limit, "bad")
            
        with tab3:
            st.caption("※ 잔여비율 시각화가 포함된 임박 재고 목록입니다.")
            render_styled_aggrid(view_df[(view_df['가용재고']>0) & (view_df['잔여일수']<=days_limit)], days_limit, "exp")
