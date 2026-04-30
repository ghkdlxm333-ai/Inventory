import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS
st.markdown("""
    <style>
    .ag-header-cell-label { display: flex !important; align-items: center !important; font-weight: bold !important; }
    .ag-header-cell-menu-button { opacity: 1 !important; visibility: visible !important; color: #0984e3 !important; }
    .metric-container { background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 12px 15px; text-align: center; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    /* 임박 기준 설정 밑의 한글 기간 문구 스타일 */
    .period-info { color: #eb4d4b; font-size: 0.9rem; font-weight: bold; margin-top: -15px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

# 숫자를 년/월/일 문구로 바꾸는 헬퍼 함수 (사이드바 표시용)
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
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"에러: {e}"

# 4. AgGrid 렌더링 함수
def render_tab_grid(data, tab_type, threshold, show_filter):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 필터창 토글 연동 및 상시 아이콘(...) 노출
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=show_filter, # 토글 상태에 따라 입력창 표시
        menuTabs=['filterMenuTab', 'generalMenuTab', 'columnsMenuTab'], # 숨긴 컬럼 꺼내기 아이콘 활성화
        minWidth=110,
        flex=1
    )
    
    # [기능] 드래그 합계 기능 (StatusBar) 다시 활성화
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

    # 탭별 컬럼 노출/숨김/제외 규칙 적용
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
            if col in ['가용재고', '불량재고', '잔여일수']:
                gb.configure_column(col, type=["numericColumn"], aggFunc='sum')
        elif col in hidden:
            gb.configure_column(col, hide=True) # ... 메뉴에서 다시 활성화 가능
        else:
            gb.configure_column(col, hide=True) # 아예 제외

    # 유효일자 경고 색상
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) return {{'color': 'red', 'fontWeight': 'bold'}};
            return null;
        }}
    """))

    return AgGrid(data, gridOptions=gb.build(), height=550, theme='alpine', allow_unsafe_jscode=True)

# 5. 실행부
st.title("📦 스마트 재고 관리 시스템")

# 필터 검색창 토글 세션 관리 (버튼 클릭 시 펼치기/닫기)
if 'filter_toggle' not in st.session_state:
    st.session_state.filter_toggle = False

def toggle_filter():
    st.session_state.filter_toggle = not st.session_state.filter_toggle

col_btn, _ = st.columns([2, 8])
with col_btn:
    st.button(f"🔍 필터 검색창 {'닫기' if st.session_state.filter_toggle else '열기'}", on_click=toggle_filter, use_container_width=True)

uploaded_file = st.file_uploader("재고 파일(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # 사이드바 설정 영역
        st.sidebar.title("⚙️ 관리 기준 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준 설정 (일)", 30, 1095, 548)
        
        # [수정사항] 슬라이더 밑에 한글 기간 표시 (잔여일수 열은 숫자 유지)
        period_text = get_period_text(days_limit)
        st.sidebar.markdown(f'<p class="period-info">{period_text}</p>', unsafe_allow_html=True)
        
        # 지표 요약
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 탭 구성 및 렌더링
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1:
            render_tab_grid(master_df[master_df['가용재고']>0], "avail", days_limit, st.session_state.filter_toggle)
        with tab2:
            render_tab_grid(master_df[master_df['불량재고']>0], "bad", days_limit, st.session_state.filter_toggle)
        with tab3:
            render_tab_grid(slow_df, "exp", days_limit, st.session_state.filter_toggle)

        # 6. 수주 가용성 분석
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_up")
        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                st.table(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']])
            except Exception as e: st.warning(f"분석 오류: {e}")
