import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="SCM 통합 재고관리", layout="wide")

# 유효일자 하이라이트 (1.5년 미만 빨간색)
def highlight_expiry(val):
    try:
        if pd.isna(val) or val == "" or val == -1: return ''
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

        # 컬럼 인덱스: D(3), E(4), G(6), N(13), C(2), F(5), AB(27), AE(30), AI(34)
        master_df = raw_data.iloc[:, [3, 4, 6, 13, 2, 5, 27, 30, 34]].copy()
        master_df.columns = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '가용재고', '불량재고', '상품바코드']

        # [데이터 정밀 세척]
        # 1. 유효일자: 시간 제거 및 날짜형 변환
        master_df['유효일자_dt'] = pd.to_datetime(master_df['유효일자'], errors='coerce')
        master_df['유효일자'] = master_df['유효일자_dt'].dt.date

        # 2. 바코드 복원 (지수 형태 8.809E+12 대응)
        def fix_barcode(val):
            if pd.isna(val): return ""
            try:
                # 숫자인 경우 지수 형태를 제거하고 정수 문자열로 변환
                return "{:.0f}".format(float(val)).strip()
            except:
                return str(val).strip()

        master_df['상품바코드'] = master_df['상품바코드'].apply(fix_barcode)
        
        # 3. 기타 검색어 필드 문자열화
        for col in ['상품코드(D)', '웰로스코드', '상품명', '화주LOT']:
            master_df[col] = master_df[col].astype(str).str.strip()

        # 4. 재고수량 숫자화
        master_df['가용재고'] = pd.to_numeric(master_df['가용재고'], errors='coerce').fillna(0).astype(int)
        master_df['불량재고'] = pd.to_numeric(master_df['불량재고'], errors='coerce').fillna(0).astype(int)

        # 탭 구성 (부진조회 추가)
        tab1, tab2, tab3 = st.tabs(["✅ 가용재고", "⚠️ 불량재고", "📉 부진재고조회"])

        # 검색 공통 함수
        def search_data(df_input, query):
            if not query: return df_input
            mask = (
                df_input['상품코드(D)'].str.contains(query, case=False) |
                df_input['웰로스코드'].str.contains(query, case=False) |
                df_input['상품명'].str.contains(query, case=False) |
                df_input['상품바코드'].str.contains(query) |
                df_input['화주LOT'].str.contains(query, case=False)
            )
            return df_input[mask]

        # --- 탭 1: 가용재고 ---
        with tab1:
            st.subheader("정상 가용 재고 조회")
            q1 = st.text_input("🔍 검색 (ME코드/웰로스/명칭/바코드/LOT)", key="q1").strip()
            res1 = search_data(master_df[master_df['가용재고'] > 0], q1)
            if not res1.empty:
                st.metric("가용 합계", f"{res1['가용재고'].sum():,} EA")
                cols = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '가용재고']
                st.dataframe(res1[cols].style.map(highlight_expiry, subset=['유효일자']), use_container_width=True, hide_index=True)
            else: st.warning("검색 결과가 없습니다.")

        # --- 탭 2: 불량재고 ---
        with tab2:
            st.subheader("불량 및 출고불가 재고")
            q2 = st.text_input("🔍 검색 (ME코드/웰로스/명칭/바코드/LOT)", key="q2").strip()
            res2 = search_data(master_df[master_df['불량재고'] > 0], q2)
            if not res2.empty:
                st.metric("불량 합계", f"{res2['불량재고'].sum():,} EA")
                cols_bad = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '불량재고']
                st.dataframe(res2[cols_bad].style.map(highlight_expiry, subset=['유효일자']), use_container_width=True, hide_index=True)
            else: st.warning("검색 결과가 없습니다.")

        # --- 탭 3: 부진조회 ---
        with tab3:
            st.subheader("부진 재고 (유효일자 6개월 이하)")
            # 오늘 기준 6개월(183일) 이하 남은 가용재고 추출
            six_months_later = datetime.now() + timedelta(days=548)
            slow_moving = master_df[(master_df['가용재고'] > 0) & (master_df['유효일자_dt'] <= six_months_later)]
            
            if not slow_moving.empty:
                st.error(f"⚠️ 유효일자가 {six_months_later.date()} 이전인 재고가 {len(slow_moving)}건 발견되었습니다.")
                st.metric("부진재고 총합", f"{slow_moving['가용재고'].sum():,} EA")
                cols_slow = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '웰로스코드', '가용재고']
                st.dataframe(slow_moving[cols_slow].style.map(highlight_expiry, subset=['유효일자']), use_container_width=True, hide_index=True)
            else:
                st.success("유효일자 6개월 이내의 부진 재고가 없습니다.")

    except Exception as e:
        st.error(f"오류 발생: {e}")
else:
    st.info("3PL 엑셀 파일을 업로드해주세요.")
