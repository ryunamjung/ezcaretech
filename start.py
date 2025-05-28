import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document

st.title("공지문 대상별 Word 생성기")

st.warning("""
📢 **공지사항**

업로드하시는 파일은 한 칸에는 꼭 **모두병원** 데이터를 **한 개씩** 따로 찍어서 입력해 주세요.

데이터 형식이 맞지 않으면 처리가 어려울 수 있습니다.  
정확한 업로드 부탁드립니다!
""")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    required_columns = ['배포홀드', '홀드사유', '공지문', '대상']
    if not all(col in df.columns for col in required_columns):
        st.error(f"필수 컬럼이 모두 포함되어야 합니다: {required_columns}")
    else:
        def is_blank_or_dash(val):
            return pd.isna(val) or str(val).strip() in ['', '-']

        # 유효 데이터만 필터링
        df_filtered = df[~(
            df['배포홀드'].apply(is_blank_or_dash) &
            df['홀드사유'].apply(is_blank_or_dash) &
            df['공지문'].apply(is_blank_or_dash) &
            df['대상'].apply(is_blank_or_dash)
        )]

        # 'Y' 포함된 행 제거
        df_filtered = df_filtered[
            ~df_filtered['배포홀드'].astype(str).str.contains('Y', na=False) &
            ~df_filtered['홀드사유'].astype(str).str.contains('Y', na=False)
        ]

        전체병원_공지문 = []
        rows = []

        for _, row in df_filtered.iterrows():
            대상_raw = row['대상']
            공지 = str(row['공지문']).strip()

            if pd.isna(대상_raw):
                대상_raw = ''

            targets = [t.strip().replace('\xa0', '') for t in str(대상_raw).split(',') if t.strip()]

            if '전체병원' in targets:
                전체병원_공지문.append(공지)
                targets = [t for t in targets if t != '전체병원']

            for target in targets:
                rows.append({'대상': target, '공지문': 공지})

        # DataFrame 생성 및 전체병원 공지 병합
        grouped_df = pd.DataFrame(rows)
        모든_대상 = grouped_df['대상'].unique().tolist()

        if grouped_df.empty and not 전체병원_공지문:
            st.warning("유효한 데이터가 없습니다.")
        else:
            if not grouped_df.empty:
                grouped = grouped_df.groupby('대상')['공지문'].apply(list).reset_index()
                # 전체병원 공지문 추가
                grouped['공지문'] = grouped['공지문'].apply(lambda lst: lst + 전체병원_공지문)
            else:
                grouped = pd.DataFrame({'대상': ['전체병원'], '공지문': [전체병원_공지문]})

            # 공지문 리스트를 문자열로 변환해서 동일한 공지문끼리 그룹화
            grouped['공지문_str'] = grouped['공지문'].apply(lambda lst: '||'.join(lst))  # 구분자 사용
            merged = grouped.groupby('공지문_str')['대상'].apply(list).reset_index()
            merged['공지문'] = merged['공지문_str'].apply(lambda s: s.split('||'))

            st.success("Word 파일 생성 완료! 아래에서 다운로드하세요:")

            for _, row in merged.iterrows():
                대상들 = sorted(row['대상'])
                공지문_list = row['공지문']

                대상_str = ', '.join(대상들)
                doc = Document()
                doc.add_heading(f"{대상_str} 공지문", level=1)

                for idx, notice in enumerate(공지문_list, start=1):
                    doc.add_paragraph(f"{idx}. {notice}")

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label=f"{대상_str}.docx 다운로드",
                    data=buffer,
                    file_name=f"{대상_str}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_{hash(대상_str)}"
                )
