import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 커스텀 CSS
st.markdown("""
    <style>
    .ag-header-cell-label { display: flex !important; align-items: center !important; font-weight: bold !important; }
    .ag-header-cell-menu-button { opacity: 1 !important; color: #0984e3 !important; visibility: visible !important; }
    
    /* 정보 창(Info Box) 스타일 */
    .info-box {
        background-color: #f0f7ff;
        border-left: 5px solid #0984e3;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
        font-size: 0.95rem;
    }
    
    .metric-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 12px 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 설정
AG_GRID_LOCALE_KR = {
    'filterOoo': '검색...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제',
    'noRowsToShow': '표시할 데이터가 없습니다', 'pinColumn': '열 고정', 'export': '내보내기',
    'autosizeThiscolumn': '이 열 자동맞춤', 'autosizeAllColumns': '모든 열 자동맞춤'
}

# 3. 데이터 로드 및 전처리 함수
@st.cache_data(show_spinner="데이터를 분석 중입니다...")
def load_and_validate_data(file):
    try:
        # 헤더 위치 찾기 (상품코드 키워드 기준)
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in [str(v) for v in row.values]:
                header_row = i + 1
                break
        
        # 전체 데이터 로드 (dtype=object로 소수점 자동변환 방지)
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row, dtype=object)
        
        # 필수 열 선택 및 이름 정의
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17]
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량']
        
        # 데이터 클렌징
        master_df['상품바코드'] = master_df['상품바코드'].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        # 숫자 형변환
        for col in ['가용재고', '불량재고', '입수량']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        # 잔여일수 계산
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        # 잔여비율 (기준 3년=1095일)
        master_df['잔여비율'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e:
        return f"오류 발생: {str(e)}"

# 4. AgGrid 렌더링 함수
def render_styled_aggrid(data, threshold, use_filter, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # 기본 컬럼 설정
    gb.configure_default_column(
        resizable=True, sortable=True, filterable=True,
        floatingFilter=use_filter,
        menuTabs=['filterMenuTab'],
        minWidth=100
    )
    
    # 숫자 포맷팅 (천 단위 콤마)
    num_format = JsCode("function(params) { return params.value ? params.value.toLocaleString() : '0'; }")
    for col in ['가용재고', '불량재고', '입수량']:
        gb.configure_column(col, type=["numericColumn"], filter='agNumberColumnFilter', valueFormatter=num_format)

    # 고정 및 스타일링
    gb.configure_column("상품코드", pinned='left', width=130)
    gb.configure_column("상품명", pinned='left', width=250)
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{ 
            return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; 
        }}
    """))

    # 잔여비율 프로그레스 바
    if tab_type == "exp":
        percent_renderer = JsCode("""
        class PercentBarRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let val = params.value;
                let color = val <= 20 ? '#ff4d4f' : (val <= 50 ? '#faad14' : '#52c41a');
                this.eGui.innerHTML = `<div style="width:100%; background:#f0f0f0; border-radius:4px; height:16px; position:relative; border:0.5px solid #ccc; margin-top:8px;">
                    <div style="width:${val}%; background:${color}; height:100%; border-radius:4px; transition: width 0.3s;"></div>
                    <span style="position:absolute; width:100%; text-align:center; font-size:10px; top:0; font-weight:bold; color:#000;">${val}%</span>
                </div>`;
            }
            getGui() { return this.eGui; }
        }
        """)
        gb.configure_column("잔여비율", headerName="잔여비율(%)", cellRenderer=percent_renderer, width=150)
    
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
    gb.configure_grid_options(
        localeText=AG_GRID_LOCALE_KR,
        suppressMenuHide=True,
        headerHeight=40
    )
    
    return AgGrid(data, gridOptions=gb.build(), height=550, theme='alpine', allow_unsafe_jscode=True)

# 5. 메인 UI 및 로직
def main():
    st.title("📦 SCM Pro: 통합 재고 관리 시스템")
    
    uploaded_file = st.file_uploader("3PL 엑셀 원본 파일을 업로드하세요.", type=['xlsx'])
    
    if uploaded_file:
        master_df = load_and_validate_data(uploaded_file)
        
        if isinstance(master_df, str):
            st.error(master_df)
        else:
            # 사이드바 설정
            st.sidebar.header("⚙️ 관리 설정")
            use_filter = st.sidebar.checkbox("🔍 열별 상세 검색창 표시", value=True)
            days_limit = st.sidebar.slider("🚨 임박 기준 설정(일)", 30, 1095, 365)
            
            # [요청 사항] 임박 기준 정보창 계산
            years = days_limit // 365
            months = (days_limit % 365) // 30
            rem_days = (days_limit % 365) % 30
            
            period_str = ""
            if years > 0: period_str += f" **{years}년**"
            if months > 0: period_str += f" **{months}개월**"
            if rem_days > 0: period_str += f" **{rem_days}일**"

            st.sidebar.markdown(f"""
            <div class="info-box">
                현재 유효기간이 <b>{period_str}</b> 이하로 남은 재고를 임박 대상으로 분류하고 있습니다.
            </div>
            """, unsafe_allow_html=True)

            # 대시보드 지표
            slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
            with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
            with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박 관리대상</div><div class="metric-value">{len(slow_df):,} 건</div></div>', unsafe_allow_html=True)

            # 메인 탭
            tab1, tab2, tab3 = st.tabs(["📊 가용재고 현황", "❌ 불량재고 현황", "⏰ 유효기간 임박분"])
            
            with tab1:
                render_styled_aggrid(master_df[master_df['가용재고']>0], days_limit, use_filter, "avail")
            with tab2:
                render_styled_aggrid(master_df[master_df['불량재고']>0], days_limit, use_filter, "bad")
            with tab3:
                st.info(f"유효기간이 {period_str} 이내인 가용재고 리스트입니다.")
                render_styled_aggrid(slow_df, days_limit, use_filter, "exp")

            # 수주 가용성 분석
            st.markdown("---")
            st.subheader("📑 실시간 수주 가용성 분석")
            order_file = st.file_uploader("수주서(서식 시트 포함)를 업로드하세요.", type=['xlsx'], key="order")
            
            if order_file:
                try:
                    order_df = pd.read_excel(order_file, sheet_name='서식')
                    order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '요청량'})
                    stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                    
                    analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                    analysis['부족량'] = (analysis['요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                    analysis['결과'] = analysis['부족량'].apply(lambda x: "✅ 출고가능" if x == 0 else "❌ 재고부족")
                    
                    st.dataframe(analysis[['결과', '상품코드', '요청량', '현재고', '부족량']], use_container_width=True)
                except Exception as e:
                    st.warning(f"수주서 양식을 확인해주세요. (에러: {e})")

if __name__ == "__main__":
    main()
