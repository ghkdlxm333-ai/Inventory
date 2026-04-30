import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (아이콘 상시 노출 및 배치)
st.markdown("""
    <style>
    /* 헤더 레이아웃: 목록명 + 필터 + 메뉴(...) 순서 */
    .ag-header-cell-label { 
        display: flex !important; 
        align-items: center !important; 
        font-weight: bold !important;
        width: 100%;
    }
    
    /* 필터 아이콘: 파란색으로 상시 노출 */
    .ag-header-icon.ag-filter-icon {
        opacity: 1 !important;
        visibility: visible !important;
        color: #0984e3 !important;
        margin-left: 4px !important;
    }

    /* 메뉴 아이콘(...): 회색으로 상시 노출 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        visibility: visible !important;
        display: flex !important;
        margin-left: auto !important; /* 오른쪽 끝으로 밀기 */
        color: #636e72 !important;
    }

    /* 지표 및 텍스트 스타일 */
    .metric-container { background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 12px; text-align: center; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    .period-info { color: #eb4d4b; font-size: 0.9rem; font-weight: bold; margin-top: -15px; }
    </style>
""", unsafe_allow_html=True)

def get_period_text(days):
    if days <= 0: return "당일 만료"
    y, m, d = days // 365, (days % 365) // 30, (days % 365) % 30
    res = []
    if y > 0: res.append(f"{y}년")
    if m > 0: res.append(f"{m}개월")
    if d > 0: res.append(f"{d}일")
    return " ".join(res) + " 이내 품목 표시"

# 3. 데이터 로드 (기존 로직 유지)
@st.cache_data
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
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"에러: {e}"

# 4. AgGrid 렌더링
def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 필터창은 끄고, 아이콘 메뉴 및 상시 노출 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=False, 
        menuTabs=['filterMenuTab', 'columnsMenuTab'], # 필터 아이콘 클릭 시 메뉴
        suppressMenu=False,
        minWidth=120
    )
    
    # 요약바 (합계/평균/개수)
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        suppressMenuHide=True, # 아이콘 상시 노출 핵심 설정
        localeText={'sum': '합계', 'avg': '평균', 'count': '개수', 'filterOoo': '필터링...'}
    )

    # 컬럼 설정 (요청하신 보임/숨김 구성)
    all_cols = data.columns.tolist()
    if tab_type == "avail":
        display, hidden = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고'], ['잔여일수', '셀']
    elif tab_type == "bad":
        display, hidden = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고'], ['잔여일수', '셀']
    else: # exp
        display, hidden = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)'], ['웰로스코드', '셀']

    for col in all_cols:
        if col in display:
            gb.configure_column(col, hide=False)
            if col in ['가용재고', '불량재고', '잔여일수']: gb.configure_column(col, aggFunc='sum')
        elif col in hidden:
            gb.configure_column(col, hide=True) # ... 메뉴에서 꺼낼 수 있음
        else:
            gb.configure_column(col, hide=True)

    # 유효일자 경고 색상
    gb.configure_column("유효일자", cellStyle=JsCode(f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"))

    return AgGrid(data, gridOptions=gb.build(), height=600, theme='alpine', allow_unsafe_jscode=True)

# 5. 실행
st.title("📦 스마트 재고 관리 (엑셀 피벗형)")

uploaded_file = st.file_uploader("재고 데이터 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바: 임박 기준 설정 및 한글 변환 문구
        st.sidebar.title("⚙️ 기준 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 1095, 548)
        st.sidebar.markdown(f'<p class="period-info">{get_period_text(days_limit)}</p>', unsafe_allow_html=True)
        
        # 상단 요약
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div>✅ 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div>⚠️ 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div>🚨 임박(가용)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1: render_styled_aggrid(master_df[master_df['가용재고']>0], days_limit, "avail")
        with tab2: render_styled_aggrid(master_df[master_df['불량재고']>0], days_limit, "bad")
        with tab3: render_styled_aggrid(slow_df, days_limit, "exp")
