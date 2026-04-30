import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정 (최상단 배치)
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (필터 아이콘 및 UI 개선)
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
    
    /* 필터 아이콘 상시 노출 강제 설정 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: inline-block !important; 
        visibility: visible !important; 
        color: #0984e3 !important;
    }
    .ag-header-icon { color: #0984e3 !important; }
    
    /* 하단 상태바(드래그 합계) 강조 */
    .ag-status-bar {
        background-color: #f8f9fa !important;
        font-weight: bold;
        color: #0984e3;
        min-height: 35px !important;
        border-top: 1px solid #dee2e6 !important;
    }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 언어 팩
AG_GRID_LOCALE_KR = {
    'filterOoo': '필터...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함', 
    'startsWith': '시작값', 'endsWith': '끝값', 'sum': '합계', 'avg': '평균', 'count': '개수',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제'
}

# 3. 데이터 로딩 및 전처리
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

        master_df['상품바코드'] = master_df['상품바코드'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().replace('nan', '')
        master_df['웰로스코드'] = master_df['웰로스코드'].astype(str).str.strip().replace('nan', '')

        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")

        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 730 * 100).clip(0, 100).fillna(0).astype(int)

        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        master_df['가용_Box환산'] = master_df.apply(lambda r: f"{r['가용재고']//r['입수량(BOX)']}Box + {r['가용재고']%r['입수량(BOX)']}EA" if r['입수량(BOX)']>0 else f"{r['가용재고']}EA", axis=1)
        master_df['불량_Box환산'] = master_df.apply(lambda r: f"{r['불량재고']//r['입수량(BOX)']}Box + {r['불량재고']%r['입수량(BOX)']}EA" if r['입수량(BOX)']>0 else f"{r['불량재고']}EA", axis=1)

        search_target_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '상품바코드']
        master_df['_search_idx'] = master_df[search_target_cols].apply(
            lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1
        )
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류 발생: {e}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)

    # [중요] 개별 목록 필터 유지 및 아이콘 고정 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True,
        suppressMenuHide=False, # 메뉴 아이콘 항상 노출
        menuTabs=['filterMenuTab'], # 클릭 시 필터가 바로 나오도록
        minWidth=110, flex=1
    )

    gb.configure_column("상품코드", pinned='left', width=130)
    gb.configure_column("상품명", pinned='left', width=250, flex=2)
    
    # 합계 설정 (aggFunc 적용)
    if "가용재고" in data.columns: gb.configure_column("가용재고", aggFunc='sum', type=["numericColumn","numberColumnFilter"])
    if "불량재고" in data.columns: gb.configure_column("불량재고", aggFunc='sum', type=["numericColumn","numberColumnFilter"])

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
    
    # [중요] 페이지네이션 제거 및 드래그 합계(상태바) 활성화
    gb.configure_grid_options(
        enableRangeSelection=True,
        pagination=False,  # 페이지 번호창 제거
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        localeText=AG_GRID_LOCALE_KR,
        suppressMenuHide=False,
        allow_unsafe_jscode=True
    )

    return AgGrid(
        data, 
        gridOptions=gb.build(), 
        height=600, 
        theme='alpine', 
        enable_enterprise_modules=True, # 드래그 합계 기능을 위해 필수
        allow_unsafe_jscode=True, 
        update_mode=GridUpdateMode.MODEL_CHANGED
    )

# --- 메인 실행부 ---
st.title("📦 3PL 재고 관리 시스템 (Pro)")
uploaded_file = st.file_uploader("3PL 엑셀 원본(.xlsx) 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 730, 548)
        
        # 상단 메트릭
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,} EA</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        search_input = st.text_input("🔍 통합 검색 (품목명/코드/LOT)", placeholder="검색어를 입력하세요").strip()
        filtered_df = master_df.copy()
        if search_input:
            for word in search_input.split():
                filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(word.lower(), na=False)]

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])

        # 열 구성 정의
        avail_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고', '가용_Box환산', '잔여일수']
        bad_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고', '불량_Box환산', '잔여일수']
        exp_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '가용_Box환산', '잔여일수', '잔여비율', '웰로스코드']

        with tab1: render_styled_aggrid(filtered_df[filtered_df['가용재고'] > 0][avail_cols], days_limit, "avail")
        with tab2: render_styled_aggrid(filtered_df[filtered_df['불량재고'] > 0][bad_cols], days_limit, "normal")
        with tab3: render_styled_aggrid(filtered_df[(filtered_df['가용재고'] > 0) & (filtered_df['잔여일수'] <= days_limit)][exp_cols], days_limit, "exp")

        # 수주 가용성 체크 모듈
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx)를 올리면 메인 재고와 즉시 비교합니다.", type=['xlsx'], key="order_checker")

        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                
                analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                
                names = master_df[['상품코드', '상품명']].drop_duplicates('상품코드')
                analysis = pd.merge(analysis, names, on='상품코드', how='left')
                
                st.table(analysis[['출고판단', '상품코드', '상품명', '수주요청량', '현재고', '부족수량']].style.apply(
                    lambda row: ['background-color: #ffcccc' if row['출고판단'] == "❌ 부족" else ''] * len(row), axis=1))
            except:
                st.warning("분석 중 오류 발생: 시트 명칭이나 컬럼을 확인하세요.")

        st.download_button(label="📊 결과 CSV 다운로드", data=filtered_df.to_csv(index=False).encode('utf-8-sig'), file_name="재고현황.csv", mime='text/csv')
