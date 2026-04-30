import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

# [GitHub 배포 필수] 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 디자인 커스텀 CSS
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
    /* AgGrid 필터 아이콘 상시 노출 및 스타일링 */
    .ag-header-cell-menu-button { opacity: 1 !important; display: block !important; color: #0984e3 !important; }
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
    'autosizeAllColumns': '모든 열 자동 크기 조절', 'export': '내보내기', 'csvExport': 'CSV로 내보내기',
}

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
        
        # 원본 데이터 매핑
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
        
        for col in ['가용재고', '불량재고']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        master_df['_search_idx'] = master_df[['상품코드', '상품명', '웰로스코드', '화주LOT']].apply(
            lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1
        )
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류 발생: {e}"

def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # 기본 컬럼 설정 (필터 활성화)
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        minWidth=100, flex=1, suppressMenu=False
    )
    
    # 틀 고정
    gb.configure_column("상품코드", pinned='left', width=120)
    gb.configure_column("상품명", pinned='left', width=200, flex=2)

    # 잔여비율 커스텀 렌더러
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

    # 탭별 컬럼 제어 로직
    if tab_type == "avail":
        # 표시: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 가용재고
        cols_to_hide = ['잔여일수', '잔여비율', '불량재고', '셀', '입수량(BOX)', '상품바코드']
        for col in cols_to_hide: 
            if col in data.columns: gb.configure_column(col, hide=True)
            
    elif tab_type == "bad":
        # 표시: 상품코드, 상품명, 웰로스코드, 화주LOT, 유효일자, 불량재고
        cols_to_hide = ['잔여일수', '잔여비율', '가용재고', '셀', '입수량(BOX)', '상품바코드']
        for col in cols_to_hide:
            if col in data.columns: gb.configure_column(col, hide=True)
            
    elif tab_type == "exp":
        # 표시: 상품코드, 상품명, 화주LOT, 유효일자, 가용재고, 잔여일수, 잔여비율(%), 웰로스코드
        gb.configure_column("잔여비율", headerName="잔여비율(%)", cellRenderer=percent_renderer, minWidth=140)
        gb.configure_column("웰로스코드", hide=True) # 요청대로 숨김 처리 (데이터는 존재)
        cols_to_hide = ['불량재고', '셀', '입수량(BOX)', '상품바코드']
        for col in cols_to_hide:
            if col in data.columns: gb.configure_column(col, hide=True)

    # 유효일자 경고 색상
    gb.configure_column("유효일자", cellStyle=JsCode(f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"))
    
    # 상태바 설정 (드래그 시 합계 표시)
    gridOptions = gb.build()
    gridOptions['enableRangeSelection'] = True  # 영역 선택 활성화
    gridOptions['statusBar'] = {
        "statusPanels": [
            {"statusPanel": "agAggregationComponent", "align": "right"} # 합계, 평균 등 표시
        ]
    }
    gridOptions['localeText'] = AG_GRID_LOCALE_KR

    return AgGrid(
        data, 
        gridOptions=gridOptions, 
        height=500, 
        theme='alpine', 
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True # 상태바 기능을 위해 필요 (Community 버전도 일부 지원)
    )

# --- 메인 실행부 ---
st.title("📦 3PL 재고 관리 시스템 (Pro)")
uploaded_file = st.file_uploader("3PL 엑셀 원본(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바 설정
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 730, 548)
        
        # n년 n개월 환산 표시
        years = days_limit // 365
        months = (days_limit % 365) // 30
        time_desc = f"설정 기준: "
        if years > 0: time_desc += f"**{years}년 **"
        if months > 0: time_desc += f"**{months}개월**"
        time_desc += " 남음"
        st.sidebar.info(f"💡 {time_desc}")

        # 메인 메트릭
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 검색 필터
        search_input = st.text_input("🔍 통합 검색", placeholder="품목명/코드/LOT 검색").strip()
        filtered_df = master_df.copy()
        if search_input:
            for word in search_input.split():
                filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(word.lower(), na=False)]

        # 탭 구성
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            st.info("💡 마우스로 수량 영역을 드래그하면 하단에 합계가 표시됩니다.")
            render_styled_aggrid(filtered_df[filtered_df['가용재고']>0], days_limit, "avail")
            
        with tab2:
            render_styled_aggrid(filtered_df[filtered_df['불량재고']>0], days_limit, "bad")
            
        with tab3:
            render_styled_aggrid(filtered_df[(filtered_df['가용재고']>0) & (filtered_df['잔여일수']<=days_limit)], days_limit, "exp")
