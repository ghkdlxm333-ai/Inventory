import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="SCM 재고관리 시스템", layout="wide")

# 스타일 설정 (빨간색 글자 등)
def highlight_expiry(val):
    try:
        expiry_date = pd.to_datetime(val)
        if expiry_date <= datetime.now() + timedelta(days=548): # 1.5년(약 548일)
            return 'color: red; font-weight: bold'
    except:
        pass
    return ''

st.title("📦 3PL 통합 재고 관리 시스템")

# 1. 파일 업로드
uploaded_file = st.file_uploader("3PL 엑셀 파일(.xlsx)을 업로드하세요", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 엑셀 로드 및 전처리
        df = pd.read_excel(uploaded_file)
        # 1번 행(단위행) 제외 및 필요한 열 인덱스로 추출
        # D(3), E(4), F(5), G(6), C(2), N(13), AB(27), AE(30), AF(31), AI(34)
        raw_data = df.drop(0).reset_index(drop=True)

        # 데이터 재구성 (요청하신 고정 순서 및 검색용 데이터 포함)
        # 고정 순서: D열(ME코드), 상품명, 화주LOT(G), 유효일자(N), 셀(C), F열(3PL코드), 가용재고환산(AB)
        master_df = raw_data.iloc[:, [3, 4, 6, 13, 2, 5, 27, 30, 31, 34]].copy()
        master_df.columns = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '가용재고', '불량재고', '총합계', '상품바코드']

        # 숫자 형변환
        for col in ['가용재고', '불량재고', '총합계']:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).astype(int)

        # 탭 구성: 가용재고 vs 불량재고
        tab1, tab2 = st.tabs(["✅ 가용재고 조회", "⚠️ 불량/출고불가 조회"])

        with tab1:
            st.subheader("정상 판매 가능 재고")
            search_val = st.text_input("🔍 검색 (ME코드 / 3PL코드 / 상품명 / 바코드)", key="search_normal")
            
            if search_val:
                # 통합 검색 로직 (D열, F열, 상품명, 바코드 모두 검색 가능)
                filtered = master_df[
                    master_df['상품코드(D)'].astype(str).str.contains(search_val, case=False) |
                    master_df['상품코드(F)'].astype(str).str.contains(search_val, case=False) |
                    master_df['상품명'].str.contains(search_val) |
                    master_df['상품바코드'].astype(str).str.contains(search_val)
                ]
                
                if not filtered.empty:
                    # 가용재고가 있는 데이터만 표시 (AB열 > 0)
                    available_only = filtered[filtered['가용재고'] > 0]
                    
                    st.metric("총 가용재고 합계", f"{available_only['가용재고'].sum():,} EA")
                    
                    # 요청하신 순서 고정 출력
                    display_cols = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '가용재고']
                    st.dataframe(
                        available_only[display_cols].style.applymap(highlight_expiry, subset=['유효일자']),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.warning("일치하는 가용재고 정보가 없습니다.")

        with tab2:
            st.subheader("불량 및 출고 불가 재고")
            search_val_bad = st.text_input("🔍 검색 (ME코드 / 3PL코드 / 상품명 / 바코드)", key="search_bad")
            
            if search_val_bad:
                filtered_bad = master_df[
                    master_df['상품코드(D)'].astype(str).str.contains(search_val_bad, case=False) |
                    master_df['상품코드(F)'].astype(str).str.contains(search_val_bad, case=False) |
                    master_df['상품명'].str.contains(search_val_bad) |
                    master_df['상품바코드'].astype(str).str.contains(search_val_bad)
                ]
                
                if not filtered_bad.empty:
                    # 불량재고가 있는 데이터만 표시 (AE열 > 0)
                    bad_only = filtered_bad[filtered_bad['불량재고'] > 0]
                    
                    st.metric("총 불량재고 합계", f"{bad_only['불량재고'].sum():,} EA", delta_color="inverse")
                    
                    # 불량 탭에서도 동일한 순서 고정 (가용재고 대신 불량재고 표시)
                    display_cols_bad = ['상품코드(D)', '상품명', '화주LOT', '유효일자', '셀', '상품코드(F)', '불량재고']
                    st.dataframe(
                        bad_only[display_cols_bad].style.applymap(highlight_expiry, subset=['유효일자']),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.warning("일치하는 불량재고 정보가 없습니다.")

    except Exception as e:
        st.error(f"파일을 분석하는 중 오류가 발생했습니다. 양식을 확인해 주세요. ({e})")
else:
    st.info("3PL 엑셀 파일을 업로드하면 탭별 조회가 가능합니다.")
