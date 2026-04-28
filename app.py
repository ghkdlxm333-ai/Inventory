import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="SCM 재고관리 시스템", layout="wide")

def highlight_expiry(val):
    try:
        if pd.isna(val) or val == "" or val == -1:
            return ''
        expiry_date = pd.to_datetime(val)
        if expiry_date <= datetime.now() + timedelta(days=548):
            return 'color: red; font-weight: bold'
    except:
        pass
    return ''

st.title("📦 3PL 통합 재고 관리 시스템")

uploaded_file = st.file_uploader("3PL 엑셀 파일(.xlsx)을 업로드하세요", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        raw_data = df.drop(0).reset_index(drop=True)

        # 컬럼 추출 및 명칭 고정
        master_df = raw_data.iloc[:, [3, 4, 6, 13, 2, 5, 27, 30, 34]].copy()
        master_df.columns = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드']

        # [데이터 클렌징 강화]
        # 1. 유효일자: 시간 제거
        master_df['유효일자'] = pd.to_datetime(master_df['유효일자'], errors='coerce').dt.date

        # 2. 바코드 검색 최적화 (가장 중요)
        # 지수 형태(E+) 방지를 위해 숫자를 문자열로 변환 후 소수점 제거
        def clean_barcode(val):
            if pd.isna(val): return ""
            try:
                # 숫자형태인 경우 소수점 제거 후 문자열화
                return str(int(float(val))).strip()
            except:
                return str(val).strip()

        master_df['상품바코드'] = master_df['상품바코드'].apply(clean_barcode)
        master_df['상품코드(D)'] = master_df['상품코드(D)'].astype(str).str.strip()
        master_df['웰로스코드'] = master_df['웰로스코드'].astype(str).str.strip()
        master_df['상품명'] = master_df['상품명'].astype(str).str.strip()

        # 3. 재고수량 숫자화
        master_df['가용재고'] = pd.to_numeric(master_df['가용재고'], errors='coerce').fillna(0).astype(int)
        master_df['불량재고'] = pd.to_numeric(master_df['불량재고'], errors='coerce').fillna(0).astype(int)

        tab1, tab2 = st.tabs(["✅ 가용재고 조회", "⚠️ 불량/출고불가 조회"])

        with tab1:
            st.subheader("정상 판매 가능 재고 (AB열 기준)")
            search_input = st.text_input("🔍 검색 (ME코드 / 웰로스코드 / 상품명 / 바코드)", key="search_normal").strip()
            
            if search_input:
                # 검색어가 포함된 모든 행 찾기
                mask = (
                    master_df['상품코드(D)'].str.contains(search_input, case=False, na=False) |
                    master_df['웰로스코드'].str.contains(search_input, case=False, na=False) |
                    master_df['상품명'].str.contains(search_input, case=False, na=False) |
                    master_df['상품바코드'].str.contains(search_input, na=False)
                )
                filtered = master_df[mask]
                available_only = filtered[filtered['가용재고'] > 0]
                
                if not available_only.empty:
                    st.metric("총 가용재고 합계", f"{available_only['가용재고'].sum():,} EA")
                    display_cols = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '가용재고']
                    st.dataframe(
                        available_only[display_cols].style.map(highlight_expiry, subset=['유효일자']),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.warning("일치하는 가용재고 내역이 없습니다.")

        with tab2:
            st.subheader("불량 및 출고 불가 재고 (AE열 기준)")
            search_input_bad = st.text_input("🔍 검색 (ME코드 / 웰로스코드 / 상품명 / 바코드)", key="search_bad").strip()
            
            if search_input_bad:
                mask_bad = (
                    master_df['상품코드(D)'].str.contains(search_input_bad, case=False, na=False) |
                    master_df['웰로스코드'].str.contains(search_input_bad, case=False, na=False) |
                    master_df['상품명'].str.contains(search_input_bad, case=False, na=False) |
                    master_df['상품바코드'].str.contains(search_input_bad, na=False)
                )
                filtered_bad = master_df[mask_bad]
                bad_only = filtered_bad[filtered_bad['불량재고'] > 0]
                
                if not bad_only.empty:
                    st.metric("총 불량재고 합계", f"{bad_only['불량재고'].sum():,} EA", delta_color="inverse")
                    display_cols_bad = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '불량재고']
                    st.dataframe(
                        bad_only[display_cols_bad].style.map(highlight_expiry, subset=['유효일자']),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.warning("일치하는 불량재고 내역이 없습니다.")

    except Exception as e:
        st.error(f"프로그램 실행 중 오류가 발생했습니다: {e}")
else:
    st.info("3PL 엑셀 원본 파일(.xlsx)을 업로드해주세요.")
