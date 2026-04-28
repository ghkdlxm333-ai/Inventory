import streamlit as st
import pandas as pd

st.set_page_config(page_title="SCM 재고관리 마스터", layout="wide")

st.title("📦 3PL 엑셀 전용 재고 조회 시스템")
st.markdown("3PL에서 다운로드한 **.xlsx** 파일을 수정 없이 그대로 업로드하세요.")

# 1. 파일 업로드 (엑셀 파일 전용으로 설정)
uploaded_file = st.file_uploader("3PL 엑셀 파일(.xlsx)을 선택하세요", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 엑셀 파일 읽기 (헤더가 2줄이므로 header=[0,1]로 설정하여 구조 파악)
        # 만약 컬럼 위치가 고정적이라면 속도를 위해 정밀하게 인덱스로 접근합니다.
        df = pd.read_excel(uploaded_file)

        # 엑셀 시트의 컬럼 위치(Index)를 기준으로 데이터 추출
        # D(3):상품코드, E(4):상품명, AB(27):가용(정상-환산), AE(30):불량(환산), AF(31):합계
        inventory_df = df.iloc[:, [3, 4, 27, 30, 31]].copy()
        
        # 1번 행(단위 행)은 데이터가 아니므로 제외
        inventory_df = inventory_df.drop(0).reset_index(drop=True)
        
        # 컬럼명 명확히 지정
        inventory_df.columns = ['상품코드', '상품명', '가용재고', '불량재고', '총합계']
        
        # 숫자 데이터 형식 변환 (문자열 섞임 방지)
        for col in ['가용재고', '불량재고', '총합계']:
            inventory_df[col] = pd.to_numeric(inventory_df[col], errors='coerce').fillna(0).astype(int)

        # 2. 검색창
        search_code = st.text_input("🔍 조회할 상품코드(D열)를 입력하세요", "")

        if search_code:
            # 대소문자 구분 없이 코드 검색
            result = inventory_df[inventory_df['상품코드'].astype(str).str.contains(search_code, case=False)]
            
            if not result.empty:
                # 동일 코드 합산 처리
                total_data = result.groupby(['상품코드', '상품명']).sum().reset_index()
                item = total_data.iloc[0]

                st.divider()
                st.subheader(f"📍 {item['상품명']}")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ 가용재고 (AB열)", f"{item['가용재고']:,} EA")
                col2.metric("⚠️ 불량/불가 (AE열)", f"{item['불량재고']:,} EA")
                col3.metric("📊 총 재고합계 (AF열)", f"{item['총합계']:,} EA")
                
                # 상세 내역
                with st.expander("세부 데이터 보기"):
                    st.dataframe(result, use_container_width=True, hide_index=True)
            else:
                st.error("입력하신 상품코드를 찾을 수 없습니다.")

        # 3. 전체 현황 요약 리스트
        st.divider()
        st.subheader("📋 전체 품목 재고 현황 (합산)")
        summary_all = inventory_df.groupby(['상품코드', '상품명']).sum().reset_index()
        st.dataframe(summary_all, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다. 파일이 열려있는지 또는 형식이 맞는지 확인해 주세요. \n 오류 내용: {e}")
else:
    st.info("3PL 전산에서 받은 엑셀 파일을 업로드하면 조회가 시작됩니다.")
