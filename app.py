import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="SCM 재고관리 시스템", layout="wide")

# 유효일자 하이라이트 함수 (Pandas 최신 버전 대응: map 사용)
def highlight_expiry(val):
    try:
        if pd.isna(val) or val == "" or val == -1:
            return ''
        expiry_date = pd.to_datetime(val)
        # 오늘 기준 1년 6개월(548일) 이하인 경우
        if expiry_date <= datetime.now() + timedelta(days=548):
            return 'color: red; font-weight: bold'
    except:
        pass
    return ''

st.title("📦 3PL 통합 재고 관리 시스템")

uploaded_file = st.file_uploader("3PL 엑셀 파일(.xlsx)을 업로드하세요", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 엑셀 로드 (헤더가 복잡하므로 일단 전체 로드 후 데이터 가공)
        df = pd.read_excel(uploaded_file)
        
        # 1번 행(단위행) 제외 및 실제 데이터 시작
        raw_data = df.drop(0).reset_index(drop=True)

        # [필수 고정 순서 및 검색용 컬럼 추출]
        # D(3):상품코드, E(4):상품명, G(6):화주LOT, N(13):유효일자, C(2):셀, F(5):상품(3PL코드), AB(27):가용(환산), AE(30):불량(환산), AI(34):상품바코드
        master_df = raw_data.iloc[:, [3, 4, 6, 13, 2, 5, 27, 30, 34]].copy()
        master_df.columns = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '가용재고', '불량재고', '상품바코드']

        # 숫자 형변환
        master_df['가용재고'] = pd.to_numeric(master_df['가용재고'], errors='coerce').fillna(0).astype(int)
        master_df['불량재고'] = pd.to_numeric(master_df['불량재고'], errors='coerce').fillna(0).astype(int)

        tab1, tab2 = st.tabs(["✅ 가용재고 조회", "⚠️ 불량/출고불가 조회"])

        with tab1:
            st.subheader("정상 판매 가능 재고 (AB열 기준)")
            search_val = st.text_input("🔍 검색 (ME코드 / 3PL코드 / 상품명 / 바코드)", key="search_normal")
            
            if search_val:
                filtered = master_df[
                    (master_df['상품코드(D)'].astype(str).str.contains(search_term := search_val, case=False)) |
                    (master_df['상품코드(F)'].astype(str).str.contains(search_term, case=False)) |
                    (master_df['상품명'].astype(str).str.contains(search_term, case=False)) |
                    (master_df['상품바코드'].astype(str).str.contains(search_term, case=False))
                ]
                
                available_only = filtered[filtered['가용재고'] > 0]
                
                if not available_only.empty:
                    st.metric("총 가용재고 합계", f"{available_only['가용재고'].sum():,} EA")
                    
                    # [순서 고정] 상품코드(D), 상품명, 화주LOT, 유효일자, 셀, 상품(F), 합계(가용재고)
                    display_cols = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '가용재고']
                    
                    # applymap 대신 map 사용 (Pandas 최신버전 대응)
                    st.dataframe(
                        available_only[display_cols].style.map(highlight_expiry, subset=['유효일자']),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.warning("일치하는 가용재고 내역이 없습니다.")

        with tab2:
            st.subheader("불량 및 출고 불가 재고 (AE열 기준)")
            search_val_bad = st.text_input("🔍 검색 (ME코드 / 3PL코드 / 상품명 / 바코드)", key="search_bad")
            
            if search_val_bad:
                filtered_bad = master_df[
                    (master_df['상품코드(D)'].astype(str).str.contains(search_term_bad := search_val_bad, case=False)) |
                    (master_df['상품코드(F)'].astype(str).str.contains(search_term_bad, case=False)) |
                    (master_df['상품명'].astype(str).str.contains(search_term_bad, case=False)) |
                    (master_df['상품바코드'].astype(str).str.contains(search_term_bad, case=False))
                ]
                
                bad_only = filtered_bad[filtered_bad['불량재고'] > 0]
                
                if not bad_only.empty:
                    st.metric("총 불량재고 합계", f"{bad_only['불량재고'].sum():,} EA", delta_color="inverse")
                    
                    # [순서 고정] 가용재고 대신 불량재고 수량 표시
                    display_cols_bad = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '불량재고']
                    
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
