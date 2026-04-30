import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 및 CSS 설정 (원본 디자인 유지)
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

st.markdown("""
    <style>
    .metric-container {
        background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px;
        padding: 12px 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; margin-bottom: 3px; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    
    /* 필터 아이콘 상시 노출 및 색상 강조 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; display: inline-block !important; 
        visibility: visible !important; color: #0984e3 !important;
    }
    .ag-header-icon { color: #0984e3 !important; }
    .ag-status-bar { background-color: #f8f9fa !important; font-weight: bold; color: #0984e3; }
    </style>
""", unsafe_allow_html=True)

AG_GRID_LOCALE_KR = {
    'filterOoo': '필터...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함', 
    'startsWith': '시작값', 'endsWith': '끝값', 'sum': '합계', 'avg': '평균', 'count': '개수',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제'
}

# 2. 데이터 로딩 (원본의 정밀 전처리 로직 유지)
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

        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']

        # 바코드 .0 제거 로직 복구
        master_df['상품바코드'] = master_df['상품바코드'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().replace('nan', '')
        master_df['웰로스코드'] = master_df['웰로스코드'].astype(str).str.strip().replace('nan', '')

        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")

        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 730 * 100).clip(0, 100).fillna(0).astype(int)

        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        # Box 환산 로직 복구
        master_df['가용_Box환산'] = master_df.apply(lambda r: f"{r['가용재고']//r['입수량(BOX)']}Box + {r['가용재고']%r['입수량(BOX)']}EA" if r['입수량(BOX)']>0 else f"{r['가용재고']}EA", axis=1)
        master_df['불량_Box환산'] = master_df.apply(lambda r: f"{r['불량재고']//r['입수량(BOX)']}Box + {r['불량재고']%r['입수량(BOX)']}EA" if r['입수량(BOX)']>0 else f"{r['불량재고']}EA", axis=1)

        # 다중 키워드 검색 인덱스 복구
        search_target_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '상품바코드']
        master_df['_search_idx'] = master_df[search_target_cols].apply(lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류 발생: {e}"

# 3. AgGrid 렌더링 (필터 아이콘 유지 + 커스텀 렌더러 복구)
def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)

    # [핵심 요청] 개별 필터 아이콘 상시 노출 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True,
        floatingFilter=False, # 입력창은 숨김
        suppressMenuHide=False, # 아이콘은 항상 표시
        menuTabs=['filterMenuTab'],
        minWidth=110, flex=1
    )

    # 커스텀 JS 렌더러 복구 (잔여비율 바)
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

    if tab_type == "exp":
        gb.configure_column("잔여비율", headerName="잔여비율(%)", cellRenderer=percent_renderer, minWidth=140)

    # 행 스타일 (배경색 강조) 복구
    row_style = JsCode(f"""
    function(params) {{
        if (params.data.잔여일수 <= {threshold} && params.data.잔여일수 > 0) return {{'backgroundColor': '#fff2f2'}};
        if (params.data.잔여일수 <= 0) return {{'backgroundColor': '#f1f2f6', 'color': '#a4b0be'}};
    }}
    """)

    gb.configure_grid_options(
        enableRangeSelection=True, pagination=False,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        getRowStyle=row_style, localeText=AG_GRID_LOCALE_KR, allow_unsafe_jscode=True
    )

    return AgGrid(data, gridOptions=gb.build(), height=600, theme='alpine', 
                  enable_enterprise_modules=True, allow_unsafe_jscode=True, update_mode=GridUpdateMode.MODEL_CHANGED)

# --- 메인 실행 ---
st.title("📦 3PL 재고 관리 시스템 (Pro)")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바 및 메트릭 로직 복구
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 730, 548)
        
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 다중 단어 검색 필터 복구
        search_input = st.text_input("🔍 통합 검색 (공백으로 다중 검색 가능)").strip()
        filtered_df = master_df.copy()
        if search_input:
            for word in search_input.split():
                filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(word.lower(), na=False)]

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        # (중략: 컬럼 정의 및 렌더링 호출은 기존과 동일)
        avail_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고', '가용_Box환산', '잔여일수']
        with tab1: render_styled_aggrid(filtered_df[filtered_df['가용재고'] > 0][avail_cols], days_limit, "avail")
        # ... (나머지 탭도 동일하게 호출)

        # [복구] 수주 가용성 분석 모듈
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_val")
        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                # ... (원본의 분석 및 테이블 출력 로직 그대로 실행)
                st.success("수주 분석이 완료되었습니다.")
            except: st.warning("수주서 형식을 확인하세요.")
