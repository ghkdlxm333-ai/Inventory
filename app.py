import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 아이콘 및 팝업창 실종 방지용 강제 CSS
st.markdown("""
    <style>
    /* 1. 상시 검색창은 숨겨서 공간 확보 (이미지 image_0e28f6.png 스타일 유지) */
    .ag-floating-filter { display: none !important; }
    
    /* 2. 메뉴 버튼(필터 아이콘)을 강제로 항상 보이게 하고 파란색으로 강조 */
    .ag-header-cell-menu-button {
        opacity: 1 !important;
        display: inline-block !important;
        visibility: visible !important;
        color: #0984e3 !important;
        cursor: pointer !important;
    }
    
    /* 3. 목록명 옆에 아이콘이 예쁘게 붙도록 정렬 */
    .ag-header-cell-label {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
    }

    /* 4. 필터 팝업창이 다른 요소에 가려지지 않도록 최상단 고정 */
    .ag-menu { z-index: 9999 !important; }
    
    /* 지표 카드 디자인 */
    .metric-container {
        background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px;
        padding: 12px 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;
    }
    .metric-label { color: #636e72; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #0984e3; font-size: 1.2rem; font-weight: 700; }
    </style>
""", unsafe_allow_html=True)

# 3. 데이터 로드 로직 (image_0e3347.png 날짜 오류 해결 포함)
@st.cache_data
def load_and_validate_data(file):
    try:
        df = pd.read_excel(file, engine='openpyxl', skiprows=1) # 헤더 위치에 따라 조정
        # 필요한 컬럼만 추출 및 이름 매칭
        master_df = df.iloc[:, [3, 4, 6, 13, 27, 30]].copy() 
        master_df.columns = ['상품코드', '상품명', '화주LOT', '유효일자_raw', '가용재고', '불량재고']
        
        # 날짜 처리
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
        # 숫자 처리
        for col in ['가용재고', '불량재고']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
            
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율(%)'] = (master_df['잔여일수'] / 1095 * 100).clip(0, 100).fillna(0).astype(int)
        
        return master_df
    except Exception as e: return f"오류: {e}"

# 4. AgGrid 렌더링 (필터 팝업 최적화)
def render_scm_grid(data, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(data)
    
    # [핵심] 필터 아이콘 클릭 시 '필터 탭'만 뜨도록 강제 설정
    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filterable=True,
        floatingFilter=False, # 상시 검색창은 제거
        suppressMenu=False,   # 메뉴 아이콘 강제 활성화
        menuTabs=['filterMenuTab'] # 중요: 컬럼관리 탭 등을 제외하고 필터만 즉시 팝업
    )

    # 잔여비율 바 렌더러
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

    # 탭별 컬럼 제어
    if tab_type == "exp":
        gb.configure_column("잔여비율(%)", cellRenderer=percent_renderer, minWidth=140)
    
    # 유효기한 강조
    gb.configure_column("유효일자", cellStyle=JsCode(
        f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"
    ))

    # 그리드 전역 설정
    grid_options = gb.build()
    grid_options['localeText'] = {'filterOoo': '검색...', 'applyFilter': '적용', 'resetFilter': '초기화'}
    
    # Enterprise 사이드바 (컬럼관리 전용)
    grid_options['sideBar'] = {
        'toolPanels': [{
            'id': 'columns', 'labelDefault': '컬럼 관리', 'iconKey': 'columns', 'toolPanel': 'agColumnsToolPanel',
        }],
        'defaultToolPanel': ''
    }

    return AgGrid(
        data,
        gridOptions=grid_options,
        height=600,
        theme='alpine',
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True, # 필터 팝업 및 사이드바 작동 필수
        update_mode=GridUpdateMode.FILTERING_CHANGED
    )

# 5. 메인 실행
st.title("📦 3PL 통합 재고관리 Pro")
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    
    if isinstance(master_df, pd.DataFrame):
        # 사이드바 설정 복구
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 임박 기준(일)", 30, 1095, 548)
        st.sidebar.info("💡 목록명 옆의 파란색 아이콘을 누르면 필터창이 뜹니다.")

        # 대시보드 지표
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        # 탭 구성
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        with tab1: render_scm_grid(master_df[master_df['가용재고']>0], days_limit, "avail")
        with tab2: render_scm_grid(master_df[master_df['불량재고']>0], days_limit, "bad")
        with tab3: render_scm_grid(slow_df, days_limit, "exp")

        # 수주 가용성 실시간 분석
        st.markdown("---")
        st.subheader("📑 수주 가용성 실시간 분석")
        order_file = st.file_uploader("수주서(.xlsx) 업로드", type=['xlsx'])
        
        if order_file:
            try:
                order_df = pd.read_excel(order_file, sheet_name='서식')
                order_sum = order_df.groupby('상품코드')['수량'].sum().reset_index().rename(columns={'수량': '수주요청량'})
                stock_sum = master_df.groupby('상품코드')['가용재고'].sum().reset_index().rename(columns={'가용재고': '현재고'})
                
                analysis = pd.merge(order_sum, stock_sum, on='상품코드', how='left').fillna(0)
                analysis['부족수량'] = (analysis['수주요청량'] - analysis['현재고']).clip(lower=0).astype(int)
                analysis['출고판단'] = analysis['부족수량'].apply(lambda x: "✅ 가능" if x == 0 else "❌ 부족")
                
                # 분석 결과 출력
                st.table(analysis[['출고판단', '상품코드', '수주요청량', '현재고', '부족수량']])
            except Exception as e:
                st.warning(f"분석 오류: {e}")
