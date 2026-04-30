import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# 1. 페이지 설정
st.set_page_config(page_title="SCM 통합 재고관리 Pro", layout="wide", page_icon="📦")

# 2. 디자인 CSS (배경색 및 메트릭 스타일)
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
    
    /* 필터 아이콘 강제 노출 및 색상 설정 */
    .ag-header-cell-menu-button { 
        opacity: 1 !important; 
        display: block !important; 
        color: #0984e3 !important; 
    }
    </style>
""", unsafe_allow_html=True)

# AgGrid 한글 설정
AG_GRID_LOCALE_KR = {
    'filterOoo': '필터...', 'equals': '같음', 'notEqual': '같지 않음', 'contains': '포함', 
    'applyFilter': '적용', 'resetFilter': '초기화', 'clearFilter': '해제', 'csvExport': 'CSV 내보내기'
}

@st.cache_data
def load_data(file):
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
        master_df['웰로스코드'] = master_df['웰로스코드'].astype(str).str.strip().replace('nan', '')
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자_raw'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.strftime('%Y-%m-%d').fillna("미기입")
        today = datetime.now()
        master_df['잔여일수'] = (master_df['유효일자_dt'] - today).dt.days.fillna(0).astype(int)
        master_df['잔여비율'] = (master_df['잔여일수'] / 730 * 100).clip(0, 100).fillna(0).astype(int)
        for col in ['가용재고', '불량재고']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)
        master_df['_search_idx'] = master_df[['상품코드', '상품명', '웰로스코드', '화주LOT']].apply(lambda x: ' '.join([str(val).lower() for val in x if pd.notna(val)]), axis=1)
        return master_df
    except Exception as e: return f"Error: {e}"

def render_grid(df, threshold, tab_type):
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # [필터 아이콘 고정 설정]
    gb.configure_default_column(
        filterable=True, sortable=True, resizable=True,
        suppressMenuHide=False, # 메뉴(필터) 아이콘 항상 노출
        menuTabs=['filterMenuTab']
    )

    # 탭별 순서 및 컬럼 제어 (지정해주신 목록 엄수)
    if tab_type == "avail":
        cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '가용재고']
        df_display = df[cols].copy()
    elif tab_type == "bad":
        cols = ['상품코드', '상품명', '웰로스코드', '화주LOT', '유효일자', '불량재고']
        df_display = df[cols].copy()
    elif tab_type == "exp":
        cols = ['상품코드', '상품명', '화주LOT', '유효일자', '가용재고', '잔여일수', '잔여비율', '웰로스코드']
        df_display = df[cols].copy()
        # 잔여비율 시각화 바
        gb.configure_column("잔여비율", cellRenderer=JsCode("""
            class PercentBarRenderer {
                init(params) {
                    this.eGui = document.createElement('div');
                    let val = params.value;
                    let color = val <= 20 ? '#ff9f43' : '#d1e9ff';
                    this.eGui.innerHTML = `<div style="width:100%; background:#f1f1f1; border-radius:5px; height:14px; margin-top:8px;">
                        <div style="width:${val}%; background:${color}; text-align:center; font-size:10px; height:100%; line-height:14px;">${val}%</div>
                    </div>`;
                }
                getGui() { return this.eGui; }
            }
        """))
        gb.configure_column("웰로스코드", hide=True) # 목록에는 포함되나 화면에선 숨김

    # [수정] 유효일자 임박 배경색: 적당한 연분홍색(#ffe3e3)으로 변경
    gb.configure_column("유효일자", cellStyle=JsCode(f"""
        function(params) {{
            if (params.data.잔여일수 <= {threshold}) {{
                return {{'backgroundColor': '#ffe3e3', 'color': '#d63031', 'fontWeight': 'bold'}};
            }}
            return null;
        }}
    """))

    # [복구] 드래그 합계 기능 (상태바)
    go = gb.build()
    go['enableRangeSelection'] = True
    go['statusBar'] = {"statusPanels": [{"statusPanel": "agAggregationComponent", "align": "right"}]}
    go['localeText'] = AG_GRID_LOCALE_KR

    return AgGrid(df_display, gridOptions=go, height=500, theme='alpine', allow_unsafe_jscode=True, enable_enterprise_modules=True)

# --- 실행부 ---
st.title("📦 3PL 재고관리 시스템")
file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

if file:
    master_df = load_data(file)
    if isinstance(master_df, str): st.error(master_df)
    else:
        st.sidebar.title("⚙️ 설정")
        days = st.sidebar.slider("🚨 임박 기준일", 30, 730, 548)
        st.sidebar.info(f"💡 기준: **{days // 365}년 {(days % 365) // 30}개월** 남음")

        search = st.text_input("🔍 검색 (품목명/코드/LOT)")
        filtered = master_df.copy()
        if search:
            filtered = filtered[filtered['_search_idx'].str.contains(search.lower())]

        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "🚨 임박재고"])
        
        with tab1:
            render_grid(filtered[filtered['가용재고']>0], days, "avail")
        with tab2:
            render_grid(filtered[filtered['불량재고']>0], days, "bad")
        with tab3:
            render_grid(filtered[(filtered['가용재고']>0) & (filtered['잔여일수']<=days)], days, "exp")
