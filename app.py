import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 커스텀 CSS (UI 개선)
st.markdown("""
    <style>
    /* 필터 아이콘 및 헤더 스타일 */
    .ag-header-cell-label { display: flex !important; align-items: center !important; font-weight: bold !important; }
    .ag-header-cell-menu-button { opacity: 1 !important; color: #0984e3 !important; visibility: visible !important; }
    
    /* 정보 창(Info Box) 스타일 */
    .info-box {
        background-color: #e3f2fd;
        border-left: 5px solid #2196f3;
        padding: 15px;
        border-radius: 8px;
        margin: 10px 0px;
        font-size: 0.9rem;
        line-height: 1.6;
    }
    
    /* 메트릭 스타일 */
    .metric-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #0984e3; font-size: 1.3rem; font-weight: 700; }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 설정
AG_GRID_LOCALE_KR = {
    'filterOoo': '조회...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함',
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제',
    'noRowsToShow': '데이터가 없습니다', 'pinColumn': '열 고정', 'export': '내보내기'
}

# 3. 데이터 로드 함수
@st.cache_data(show_spinner="엑셀 분석 중...")
def load_data(file):
    try:
        df_temp = pd.read_excel(file, engine='openpyxl', nrows=15)
        header_row = 0
        for i, row in df_temp.iterrows():
            if '상품코드' in [str(v) for v in row.values]:
                header_row = i + 1
                break
        
        df = pd.read_excel(file, engine='openpyxl', skiprows=header_row, dtype=object)
        raw_cols = [3, 4, 6, 13, 2, 5, 27, 30, 33, 17]
        master_df = df.iloc[:, raw_cols].copy()
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드', '입수량']
        
        # 클렌징
        master_df['상품바코드'] = master_df['상품바코드'].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        for col in ['가용재고', '불량재고', '입수량']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).astype(int)
        
        return master_df.sort_values(by='유효일자_dt')
    except Exception as e: return f"에러: {e}"

# 4. AgGrid 렌더링 (필터 복구 버전)
def render_aggrid(data, threshold, show_filter, tab_type="normal"):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 모든 컬럼에 필터와 개별 검색창(floatingFilter) 적용
    gb.configure_default_column(
        filterable=True,
        sortable=True,
        resizable=True,
        floatingFilter=show_filter,  # 사이드바 설정에 따라 검색창 노출
        suppressMenu=False           # 필터 메뉴 아이콘 활성화
    )
    
    # 컬럼별 세부 설정
    gb.configure_column("상품코드", pinned='left', width=140)
    gb.configure_column("상품명", pinned='left', width=280)
    
    # 숫자 포맷 (천 단위 콤마)
    v_format = JsCode("function(params) { return params.value ? params.value.toLocaleString() : '0'; }")
    for c in ['가용재고', '불량재고', '입수량']:
        gb.configure_column(c, valueFormatter=v_format, type=["numericColumn"], filter='agNumberColumnFilter')

    # 유효기간 강조 스타일
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{ 
            if (params.data.잔여일수 <= {threshold}) return {{'color': 'white', 'backgroundColor': '#e74c3c', 'fontWeight': 'bold'}};
            return null;
        }}
    """))

    if tab_type == "exp":
        bar_js = JsCode("""
        class BarRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let v = params.value;
                let c = v <= 20 ? '#ff4d4f' : (v <= 50 ? '#faad14' : '#52c41a');
                this.eGui.innerHTML = `<div style="width:100%; background:#eee; border-radius:10px; height:14px; margin-top:10px; border:1px solid #ddd">
                    <div style="width:${v}%; background:${c}; height:100%; border-radius:10px; text-align:center; font-size:10px; line-height:14px; color:black; font-weight:bold">${v}%</div>
                </div>`;
            }
            getGui() { return this.eGui; }
        }
        """)
        gb.configure_column("잔여비율", cellRenderer=bar_js, width=150)

    gb.configure_pagination(paginationPageSize=20)
    gb.configure_grid_options(localeText=AG_GRID_LOCALE_KR, suppressMenuHide=True)
    
    return AgGrid(data, gridOptions=gb.build(), height=550, theme='alpine', allow_unsafe_jscode=True)

# 5. 메인 실행
def main():
    st.title("📦 SCM 통합 재고관리 시스템 Pro")
    file = st.file_uploader("3PL 재고 엑셀 업로드", type=['xlsx'])
    
    if file:
        df = load_data(file)
        if isinstance(df, str): st.error(df)
        else:
            # --- 사이드바 및 정보창 ---
            st.sidebar.header("⚙️ 검색 및 필터 설정")
            show_filter = st.sidebar.checkbox("🔍 열별 개별 검색창 표시", value=True)
            days = st.sidebar.slider("🚨 임박 기준 설정(일)", 30, 1095, 365)
            
            # [요청] 년/월/일 변환 로직
            y, m, d = days // 365, (days % 365) // 30, (days % 365) % 30
            period_text = f"{f'**{y}년** ' if y else ''}{f'**{m}개월** ' if m else ''}{f'**{d}일**' if d else ''}"
            
            st.sidebar.markdown(f"""
            <div class="info-box">
                현재 <b>{period_text}</b> 이하로 남은<br>재고를 임박 관리대상으로 분류합니다.
            </div>
            """, unsafe_allow_html=True)
            # ------------------------

            # 지표 계산
            slow = df[(df['가용재고'] > 0) & (df['잔여일수'] <= days)]
            cols = st.columns(3)
            cols[0].markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
            cols[1].markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
            cols[2].markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용)</div><div class="metric-value">{len(slow):,} 건</div></div>', unsafe_allow_html=True)

            tab1, tab2, tab3 = st.tabs(["📊 가용재고", "❌ 불량재고", "⏰ 임박재고"])
            with tab1: render_aggrid(df[df['가용재고']>0], days, show_filter, "avail")
            with tab2: render_aggrid(df[df['불량재고']>0], days, show_filter, "bad")
            with tab3: 
                st.warning(f"유효기간이 {period_text} 이내인 품목입니다.")
                render_aggrid(slow, days, show_filter, "exp")

if __name__ == "__main__":
    main()
