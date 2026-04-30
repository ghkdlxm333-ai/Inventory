import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 커스텀 CSS (아이콘 위치 및 슬라이더 색상 조정)
st.markdown("""
    <style>
    /* 상시 검색창 영역 제거 */
    .ag-floating-filter { display: none !important; }
    
    /* 헤더 목록명과 아이콘 정렬 */
    .ag-header-cell-label {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        width: 100% !important;
    }

    /* 필터 아이콘 파란색 강조 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        visibility: visible !important;
        color: #0984e3 !important;
    }

    /* 지표 카드 디자인 */
    .metric-container {
        background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px;
        padding: 12px 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    
    /* Streamlit 슬라이더 색상 커스텀 (사진의 빨간색 테마 반영) */
    .stSlider [data-baseweb="slider"] { margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

# 3. 데이터 로드 및 전처리
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
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17] 
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량(BOX)']
        
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        for col in ['가용재고', '불량재고', '입수량(BOX)']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        # 비율 계산 기준: 3년(1095일)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        master_df['가용_Box환산'] = master_df.apply(lambda x: round(x['가용재고']/x['입수량(BOX)'], 1) if x['입수량(BOX)'] > 0 else 0, axis=1)

        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"오류: {e}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True, 
        floatingFilter=False, 
        menuTabs=['filterMenuTab'], 
        suppressMenu=False
    )

    # 잔여비율 프로그레스 바 디자인 (JsCode)
    percent_renderer = JsCode("""
    class PercentBarRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            let val = params.value;
            let color = val <= 20 ? '#ff4d4f' : (val <= 50 ? '#ffa940' : '#40a9ff');
            this.eGui.innerHTML = `
                <div style="width:100%; background-color:#f0f0f0; border-radius:12px; height:18px; position:relative; overflow:hidden; border:0.5px solid #d9d9d9; margin-top:6px;">
                    <div style="width:${val}%; background-color:${color}; height:100%; transition: width 0.5s ease-in-out;"></div>
                    <span style="position:absolute; width:100%; text-align:center; top:0; left:0; font-size:11px; font-weight:800; color:#262626; line-height:18px;">${val}%</span>
                </div>
            `;
        }
        getGui() { return this.eGui; }
    }
    """)

    # 탭별 컬럼 구성
    if tab_type == "avail":
        active = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
    elif tab_type == "bad":
        active = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
    else: # exp (임박재고)
        active = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율(%)']
        gb.configure_column("잔여비율(%)", cellRenderer=percent_renderer, minWidth=150)

    for col in data.columns:
        if col in active:
            gb.configure_column(col, hide=False, pinned='left' if col in ['상품코드', '상품명'] else None)
        else:
            gb.configure_column(col, hide=True)

    # 유효일자 강조
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    gb.configure_grid_options(
        enableRangeSelection=True,
        statusBar={"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]},
        localeText={'filterOoo': '검색...', 'applyFilter': '적용', 'resetFilter': '초기화'},
        suppressMenuHide=True
    )

    return AgGrid(data, gridOptions=gb.build(), height=500, theme='alpine', allow_unsafe_jscode=True, enable_enterprise_modules=True)

# 5. 메인 실행부
st.title("📦 3PL 통합 재고관리 Pro")
uploaded_file = st.file_uploader("3PL 엑셀 원본 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        # --- 사이드바 설정 영역 ---
        with st.sidebar:
            st.markdown("### ⚙️ 관리 설정")
            # 사진의 디자인 반영: 임박 기준 슬라이더
            days_limit = st.slider("🚨 임박 기준(일)", 30, 1095, 548)
            
            # 기준일 읽어주기 로직 (년/개월 계산)
            y = days_limit // 365
            m = (days_limit % 365) // 30
            period_text = f"약 {f'{y}년 ' if y > 0 else ''}{f'{m}개월 ' if m > 0 else ''}남음"
            
            st.info(f"💡 설정 기준: {period_text}")
            st.markdown("---")
            st.write("각 열 제목 우측 아이콘을 클릭하여 필터링하세요.")

        # 데이터 필터링
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        
        # 상단 요약 지표
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 탭 출력
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1: render_styled_aggrid(master_df[master_df['가용재고']>0], days_limit, "avail")
        with tab2: render_styled_aggrid(master_df[master_df['불량재고']>0], days_limit, "bad")
        with tab3: render_styled_aggrid(slow_df, days_limit, "exp")

        # 수주 가용성 분석 섹션
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'], key="order_upload")
        
        if order_file:
            try:
                order_df = pd.read_excel(order_file) # 서식 시트가 없다면 기본 로드
                # 상품코드/수량 컬럼이 있는지 확인 후 분석
                if '상품코드' in order_df.columns and '수량' in order_df.columns:
                    order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                    stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                    analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                    analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                    analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                    st.table(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']])
                else:
                    st.warning("수주서에 '상품코드'와 '수량' 컬럼이 필요합니다.")
            except Exception as e:
                st.warning(f"분석 오류: {e}")
