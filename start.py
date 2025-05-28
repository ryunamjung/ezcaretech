import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document

st.title("공지문 대상별 Word 생성기")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    required_columns = ['배포홀드', '홀드사유', '공지문', '대상']
    if not all(col in df.columns for col in required_columns):
        st.error(f"필수 컬럼이 모두 포함되어야 합니다: {required_columns}")
    else:
        def is_blank_or_dash(val):
            return pd.isna(val) or str(val).strip() in ['', '-']

        df_filtered = df[~(
            df['배포홀드'].apply(is_blank_or_dash) &
            df['홀드사유'].apply(is_blank_or_dash) &
            df['공지문'].apply(is_blank_or_dash) &
            df['대상'].apply(is_blank_or_dash)
        )]

        df_filtered = df_filtered[~df_filtered['홀드사유'].astype(str).str.contains('Y', na=False)]

        전체병원_공지문 = []

        rows = []
        for _, row in df_filtered.iterrows():
            대상_raw = row['대상']
            # NaN이면 빈 문자열로 치환
            if pd.isna(대상_raw):
                대상_raw = ''

            # 쉼표로 분리 후, 각 항목 strip + 보이지 않는 공백 제거
            targets = [t.strip().replace('\xa0', '') for t in str(대상_raw).split(',')]
            # 빈 문자열 제거
            targets = [t for t in targets if t]

            # '전체병원' 포함 여부 체크 (소문자 변환해서 비교할 수도 있음)
            if any(t == '전체병원' for t in targets):
                전체병원_공지문.append(str(row['공지문']))
                # '전체병원' 공지문은 개별 대상에 넣지 않음
                # 만약 '전체병원'외 다른 대상도 있으면 개별 대상에도 넣고 싶으면 아래로 분리 처리 필요
                other_targets = [t for t in targets if t != '전체병원']
                for target in other_targets:
                    rows.append({'대상': target, '공지문': str(row['공지문'])})
            else:
                for target in targets:
                    rows.append({'대상': target, '공지문': str(row['공지문'])})

        grouped_df = pd.DataFrame(rows)

        if grouped_df.empty and not 전체병원_공지문:
            st.warning("유효한 데이터가 없습니다.")
        else:
            if not grouped_df.empty:
                grouped = grouped_df.groupby('대상')['공지문'].apply(list).reset_index()
                # 전체병원 공지문을 모든 대상 공지문 리스트 뒤에 추가
                grouped['공지문'] = grouped['공지문'].apply(lambda lst: lst + 전체병원_공지문)
            else:
                grouped = pd.DataFrame({'대상': ['전체병원'], '공지문': [전체병원_공지문]})

            st.success("Word 파일 생성 완료! 아래에서 다운로드하세요:")

            for _, row in grouped.iterrows():
                대상 = row['대상']
                공지문_list = row['공지문']

                doc = Document()
                doc.add_heading(f"{대상} 공지문", level=1)

                for idx, notice in enumerate(공지문_list, start=1):
                    doc.add_paragraph(f"{idx}. {notice}")

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label=f"{대상}.docx 다운로드",
                    data=buffer,
                    file_name=f"{대상}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
