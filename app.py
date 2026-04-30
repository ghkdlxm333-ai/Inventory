import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀
st.markdown("""
    <style>
    .ag-header-cell-label { display: flex !important; align-items: center !important; font-weight: bold !important; }
    .ag-header-cell-menu-button { opacity: 1 !important; visibility: visible !important; color: #0984e3 !important; }
    .metric-container { background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 12px 15px; text-align: center; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    /* 슬라이더 밑 설명 문구 스타일 */
    .period-info { color: #eb4d4b; font-size: 0.9rem; font-weight: bold; margin-top: -15px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

# 숫자를 년/월/일 문구로 바꾸는 헬퍼 함수 (슬라이더 밑 표시용)
def get_period_text(days):
    if days <= 0: return "당일 만료"
    y = days // 365
    m = (days % 365) // 30
    d = (days % 365) % 30
    res = []
    if y > 0: res.append(f"{y}년")
    if m > 0: res.append(f"{m}개월")
    if d > 0: res.append(f"{d}일")
    return " ".join(res) + " 이내 품목 표시"

# 3. 데이터 로드 및 전처리
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

# 4. AgGrid 렌더링 엔진
def render_tab_grid(data, tab_type, threshold, show_filter):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # 엑셀식 상시 메뉴(...) 및 필터 토글 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=show_filter,
        menuTabs=['filterMenuTab', 'generalMenuTab', 'columnsMenuTab'],
        minWidth=110
    )
    
    # 드래그 합계(StatusBar)
    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={
            "statusPanels": [
                {"statusPanel": "agTotalAndFilteredRowCountComponent", "align": "left"},
                {"statusPanel": "agAggregationComponent", "align": "right"}
            ]
        },
        localeText={'sum': '합계', 'avg': '평균', 'count': '개수', 'filterOoo': '검색...'}
    )

    # 탭별 컬럼 노출 규칙 적용
    all_cols = data.columns.tolist()
    if tab_type == "avail":
        display = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        hidden = ['가용_Box환산', '잔여일수', '셀', '입수량(BOX)']
    elif tab_type == "bad":
        display = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        hidden = ['불량_Box환산', '잔여일수', '셀', '입수량(BOX)']
    elif tab_type == "exp":
        display = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)', '웰로스코드']
        hidden = ['셀', '입수량(BOX)']

    for col in all_cols:
        if col in display:
            gb.configure_column(col, hide=False)
            if col in ['가용재고', '불량재고', '잔여일수']: gb.configure_column(col, aggFunc='sum')
        elif col in hidden:
            gb.configure_column(col, hide=True) 
        else:
            gb.configure_column(col, hide=True)

    # 유효일자 경고 조건부 서식
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) return {{'color': 'red', 'fontWeight': 'bold'}};
            return null;
        }}
    """))

    return AgGrid(data, gridOptions=gb.build(), height=550, theme='alpine', allow_unsafe_jscode=True)

# 5. 실행부
st.title("📦 스마트 재고 관리 시스템")

# 필터 검색창 토글 세션 관리
if 'filter_toggle' not in st.session_state:
    st.session_state.filter_toggle = False

def toggle_filter():
    st.session_state.filter_toggle = not st.session_state.filter_toggle

col_btn1, col_btn2 = st.columns([2, 8])
with col_btn1:
    st.button(f"🔍 필터 검색창 {'닫기' if st.session_state.filter_toggle else '열기'}", on_click=toggle_filter)

uploaded_file = st.file_uploader("재고 파일 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바 설정
        st.sidebar.title("⚙️ 관리 기준 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준 설정 (일)", 30, 1095, 548)
        
        # 슬라이더 바로 밑에 요청하신 '몇년 몇개월 몇일' 표시
        period_text = get_period_text(days_limit)
        st.sidebar.markdown(f'<p class="period-info">{period_text}</p>', unsafe_allow_html=True)
        
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1:
            render_tab_grid(master_df[master_df['가용재고']>0], "avail", days_limit, st.session_state.filter_toggle)
        with tab2:
            render_tab_grid(master_df[master_df['불량재고']>0], "bad", days_limit, st.session_state.filter_toggle)
        with tab3:
            # 임박재고 데이터 필터링
            exp_data = master_df[master_df['잔여일수'] <= days_limit]
            render_tab_grid(exp_data, "exp", days_limit, st.session_state.filter_toggle)
