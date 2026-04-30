import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

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
    .stMetric { display: none; }
    .ag-header-cell-label { font-weight: bold !important; font-size: 13px !important; color: #2d3436; }
    .ag-header-cell-menu-button { opacity: 1 !important; display: block !important; color: #0984e3 !important; visibility: visible !important; }
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
    'clearFilter': '필터 해제', 'noRowsToShow': '표시할 데이터가 없습니다', 'export': '내보내기',
}

@st.cache_data(show_spinner="데이터를 분석하고 있습니다...")
def load_and_validate_data(file):
    try:
        # 헤더 자동 찾기 (상품코드 키워드 기준)
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in str(row.values):
                header_row = i + 1
                break
        
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row)
        
        # 필요한 컬럼만 추출 (인덱스 기준)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        # 데이터 클렌징
        master_df['상품바코드'] = master_df['상품바코드'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().replace('nan', '')
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        # [수정] 화장품 유통기한 3년(1095일) 기준 잔여비율 계산
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        master_df['_search_idx'] = master_df[['상품코드', '상품명', '웰로스코드', '화주LOT']].apply(
            lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1
        )
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"오류 발생: {e}"

def render_styled_aggrid(data, threshold, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)
    gb.configure_default_column(resizable=True, sortable=True, filterable=True, minWidth=110, flex=1)
    
    gb.configure_column("상품코드", pinned='left', width=130)
    gb.configure_column("상품명", pinned='left', width=250, flex=2)
    
    # 잔여비율 시각화 렌더러
    percent_renderer = JsCode("""
    class PercentBarRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            let val = params.value;
            // 3년 기준: 20%(약 7개월 미만)는 위험군으로 주황색 표시
            let color = val <= 20 ? '#ff7675' : (val <= 50 ? '#fdcb6e' : '#74b9ff');
            this.eGui.innerHTML = `<div style="width:100%; background:#f1f1f1; border-radius:10px; height:18px; border:1px solid #ddd; margin-top:6px;">
                <div style="width:${val}%; background:${color}; text-align:center; font-size:11px; font-weight:bold; height:100%; line-height:18px; color:white; border-radius:10px;">${val}%</div>
            </div>`;
        }
        getGui() { return this.eGui; }
    }
    """)

    if tab_type == "exp":
        gb.configure_column("잔여비율", headerName="유통기한 잔여도(3년 기준)", cellRenderer=percent_renderer, minWidth=180)
    else:
        gb.configure_column("잔여일수", hide=True)
        gb.configure_column("잔여비율", hide=True)

    # 유효일자 임박 시 텍스트 빨간색 강조
    gb.configure_column("유효일자", cellStyle=JsCode(f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"))
    
    gb.configure_grid_options(
        enableRangeSelection=True,
        pagination=True,
        paginationPageSize=20,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        localeText=AG_GRID_LOCALE_KR
    )
    return AgGrid(data, gridOptions=gb.build(), height=500, theme='alpine', allow_unsafe_jscode=True)

# 메인 UI
st.title("📦 3PL 화장품 통합 재고관리 시스템")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str):
        st.error(master_df)
    else:
        # 사이드바 설정
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 유효일자 알림 기준(일)", 30, 1095, 365) # 기본 1년
        
        # 대시보드 카드
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 검색창
        search_input = st.text_input("🔍 통합 검색", placeholder="상품명, 코드, LOT 등으로 검색하세요").strip()
        filtered_df = master_df.copy()
        if search_input:
            for word in search_input.split():
                filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(word.lower(), na=False)]

        # 탭 구성
        tab1, tab2, tab3, tab4 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고 분석", "📑 수주 가용성 체크"])
        
        with tab1:
            render_styled_aggrid(filtered_df[filtered_df['가용재고']>0], days_limit, "avail")
        with tab2:
            render_styled_aggrid(filtered_df[filtered_df['불량재고']>0], days_limit, "bad")
        with tab3:
            render_styled_aggrid(filtered_df[(filtered_df['가용재고']>0) & (filtered_df['잔여일수']<=days_limit)], days_limit, "exp")
        
        with tab4:
            order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_val")
            if order_file:
                try:
                    order_df = pd.read_excel(order_file, sheet_name='서식')
                    order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                    stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                    
                    analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                    analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                    analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                    
                    # 상품명 매칭
                    names = master_df[['상품코드', '상품명']].drop_duplicates('상품코드')
                    analysis = pd.merge(analysis, names, on='상품코드', how='left')
                    
                    st.dataframe(
                        analysis[['출고판단', '상품코드', '상품명', '수주요청량', '현재고', '부족수량']],
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "출고판단": st.column_config.TextColumn("판단"),
                            "부족수량": st.column_config.NumberColumn("부족분", format="%d 📦")
                        }
                    )
                except Exception as e:
                    st.warning(f"수주서 서식을 확인해주세요: {e}")
