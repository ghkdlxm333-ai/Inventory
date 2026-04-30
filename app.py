import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode

# 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 디자인 커스텀 CSS (필터 아이콘 및 스타일 보강)
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
    
    /* AgGrid 필터 아이콘 강제 노출 설정 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: block !important; 
        visibility: visible !important; 
        color: #0984e3 !important;
    }
    .ag-header-icon { color: #0984e3 !important; }
    .ag-header-cell-label { font-weight: bold !important; }
    
    .sidebar-info { font-size: 0.9rem; color: #0984e3; font-weight: bold; background: #e1f5fe; padding: 10px; border-radius: 8px; margin-top: -15px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

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
        
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        
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

def render_styled_aggrid(data, threshold, tab_type):
    # 탭별 순서 정의
    if tab_type == "avail":
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
    elif tab_type == "bad":
        display_cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
    elif tab_type == "exp":
        display_cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율', '웰로스코드']
    
    # 데이터 순서 재정렬 및 필요한 컬럼만 추출
    target_df = data[display_cols].copy()
    
    gb = GridOptionsBuilder.from_dataframe(target_df)
    
    # 공통 설정: 필터 메뉴 항상 노출(suppressMenuHide) 및 필터 활성화
    gb.configure_default_column(
        resizable=True, 
        sortable=True, 
        filterable=True, 
        suppressMenuHide=False,  # 메뉴 버튼 숨김 방지 (아이콘 상시 노출 핵심)
        minWidth=100
    )
    
    # 컬럼별 개별 설정
    gb.configure_column("상품코드", pinned='left', width=130)
    gb.configure_column("상품명", pinned='left', width=250)
    
    if "잔여비율" in display_cols:
        percent_renderer = JsCode("""
        class PercentBarRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let val = params.value;
                let color = val <= 20 ? '#ff7675' : (val <= 50 ? '#fdcb6e' : '#74b9ff');
                this.eGui.innerHTML = `<div style="width:100%; background:#f1f1f1; border-radius:10px; height:18px; border:1px solid #ddd; margin-top:6px;">
                    <div style="width:${val}%; background:${color}; text-align:center; font-size:11px; font-weight:bold; height:100%; line-height:18px; color:white; border-radius:10px;">${val}%</div>
                </div>`;
            }
            getGui() { return this.eGui; }
        }
        """)
        gb.configure_column("잔여비율", headerName="잔여비율(%)", cellRenderer=percent_renderer, minWidth=150)

    # 유효일자 강조 색상
    gb.configure_column("유효일자", cellStyle=JsCode(f"function(params) {{ return params.data.잔여일수 <= {threshold} ? {{'color': '#d63031', 'fontWeight': 'bold'}} : null; }}"))
    
    gb.configure_grid_options(
        enableRangeSelection=True,
        pagination=True,
        paginationPageSize=20,
        localeText={'filterOoo': '필터...', 'equals': '같음', 'contains': '포함', 'applyFilter': '적용', 'resetFilter': '초기화'}
    )
    
    return AgGrid(
        target_df, 
        gridOptions=gb.build(), 
        height=550, 
        theme='alpine', 
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.VALUE_CHANGED
    )

# 실행부
uploaded_file = st.file_uploader("📦 3PL 재고 엑셀 업로드", type=['xlsx'])

if uploaded_file:
    master_df = load_and_validate_data(uploaded_file)
    if isinstance(master_df, str):
        st.error(master_df)
    else:
        st.sidebar.title("⚙️ 관리 설정")
        days_limit = st.sidebar.slider("🚨 유효일자 알림 기준(일)", 30, 1095, 365)
        
        years = days_limit // 365
        months = (days_limit % 365) // 30
        duration_text = f"현재 설정: {f'{years}년 ' if years > 0 else ''}{f'{months}개월 ' if months > 0 else ''}이하 재고 검색"
        st.sidebar.markdown(f'<div class="sidebar-info">{duration_text}</div>', unsafe_allow_html=True)
        
        slow_df = master_df[(master_df['가용재고'] > 0) & (master_df['잔여일수'] <= days_limit)]
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-container"><div class="metric-label">✅ 전체 가용재고</div><div class="metric-value">{master_df["가용재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-container"><div class="metric-label">⚠️ 전체 불량재고</div><div class="metric-value">{master_df["불량재고"].sum():,}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-container"><div class="metric-label">🚨 임박(가용기준)</div><div class="metric-value">{len(slow_df)}건</div></div>', unsafe_allow_html=True)

        search_input = st.text_input("🔍 통합 검색", placeholder="상품명, 코드, LOT 검색").strip()
        filtered_df = master_df.copy()
        if search_input:
            for word in search_input.split():
                filtered_df = filtered_df[filtered_df['_search_idx'].str.contains(word.lower(), na=False)]

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            render_styled_aggrid(filtered_df[filtered_df['가용재고']>0], days_limit, "avail")
        with tab2:
            render_styled_aggrid(filtered_df[filtered_df['불량재고']>0], days_limit, "bad")
        with tab3:
            render_styled_aggrid(filtered_df[(filtered_df['가용재고']>0) & (filtered_df['잔여일수']<=days_limit)], days_limit, "exp")
