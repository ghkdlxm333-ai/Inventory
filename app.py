import streamlit as st
import pandas as pd

st.set_page_config(page_title="SCM 재고관리 마스터", layout="wide")

st.title("🚀 SCM팀 가용재고 조회 시스템")
st.markdown("업로드하신 엑셀 데이터를 바탕으로 **가용/불량/합계 재고**를 자동으로 분류합니다.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("3PL 엑셀(CSV) 파일을 업로드하세요", type=['csv'])

if uploaded_file is not None:
    try:
        # 데이터 로드: 0번 행은 컬럼명, 1번 행은 단위행이므로 drop(0) 처리
        df = pd.read_csv(uploaded_file)
        df_data = df.drop(0).reset_index(drop=True)

        # 컬럼 매핑 (0부터 시작하는 인덱스 기준)
        # D(3): 상품코드, E(4): 상품명
        # AB(27): 가용재고(정상수량-환산)
        # AE(30): 불량/출고불가(불량수량-환산)
        # AF(31): 총 재고 합계
        
        inventory_df = df_data.iloc[:, [3, 4, 27, 30, 31]].copy()
        inventory_df.columns = ['상품코드', '상품명', '가용재고', '불량재고', '총합계']
        
        # 숫자 형변환 (에러 발생 시 0으로 대체)
        cols_to_fix = ['가용재고', '불량재고', '총합계']
        for col in cols_to_fix:
            inventory_df[col] = pd.to_numeric(inventory_df[col], errors='coerce').fillna(0).astype(int)

        # 2. 검색 인터페이스
        search_term = st.text_input("🔍 상품코드 또는 상품명을 입력하세요", "")

        if search_term:
            # 검색 필터링
            filtered = inventory_df[
                inventory_df['상품코드'].astype(str).str.contains(search_term, case=False) |
                inventory_df['상품명'].str.contains(search_term)
            ]
            
            if not filtered.empty:
                # 동일 상품코드 합계 계산 (여러 로케이션 분산 재고 합산)
                summary = filtered.groupby(['상품코드', '상품명']).sum().reset_index()
                
                # 상단에 요약 카드 표시
                st.divider()
                item = summary.iloc[0]
                st.subheader(f"📦 {item['상품명']}")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("✅ 가용재고 (AB열)", f"{item['가용재고']:,} EA")
                m2.metric("⚠️ 불량/불가 (AE열)", f"{item['불량재고']:,} EA", delta_color="inverse")
                m3.metric("📊 총 합계 (AF열)", f"{item['총합계']:,} EA")
                
                # 상세 데이터 (로케이션 정보 등이 필요할 경우)
                with st.expander("상세 로케이션별 내역 보기"):
                    st.dataframe(filtered, use_container_width=True, hide_index=True)
            else:
                st.warning("조회된 상품이 없습니다.")

        # 3. 하단 전체 현황판
        st.divider()
        st.subheader("전체 품목 재고 현황")
        full_summary = inventory_df.groupby(['상품코드', '상품명']).sum().reset_index()
        st.dataframe(full_summary, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다. 컬럼 위치를 다시 확인해주세요. ({e})")
else:
    st.info("왼쪽 상단의 파일 업로더를 통해 3PL 엑셀 파일을 넣어주세요.")
