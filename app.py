# requirements:
# pip install streamlit python-docx

import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
from datetime import date

st.title("결석 신고서 자동 생성기")

st.subheader("학생 정보")
grade = st.text_input("학년")
cls = st.text_input("반")
num = st.text_input("번호")
name = st.text_input("성명")

st.subheader("결석 기간")
start_date = st.date_input("결석 시작일", value=date.today())
end_date = st.date_input("결석 종료일", value=date.today())
days = st.number_input("결석 일수(공휴일 제외)", min_value=1, value=1)

st.subheader("결석 사유")
reason = st.text_area("사유를 구체적으로 입력하세요")

st.subheader("붙임 서류")
has_doc = st.checkbox("진단서/진료 확인서 있음 (3일 이상 시 필수)")
has_rx = st.checkbox("병원 처방전 또는 약봉투")
has_parent = st.checkbox("보건결석 학부모 의견서")
has_etc = st.checkbox("기타 증빙 서류")

parent_name = st.text_input("보호자 성명")
today = st.date_input("작성일", value=date.today())

if st.button("결석 신고서 만들기 (docx)"):
    doc = Document()

    # 제목
    title = doc.add_paragraph()
    run = title.add_run("결 석 신 고 서")
    run.bold = True
    run.font.size = Pt(18)
    title.alignment = 1  # 가운데 정렬

    doc.add_paragraph("※ 「결석신고서」는 결석한 날로부터 3일 이내에 제출하여 학교의 승인을 받아야 합니다.")
    doc.add_paragraph("    [  ]에는 해당되는 곳에 √표를 합니다. 「담임교사 확인서」는 결석신고서를 바탕으로 담임교사가 작성합니다.")

    # 학생 정보
    p = doc.add_paragraph()
    p.add_run("\n학 생  ").bold = True
    p.add_run(f"{grade}학년 {cls}반 {num}번    성명: {name}")

    # 기간
    p = doc.add_paragraph()
    p.add_run("\n기 간  ").bold = True
    p.add_run(f"{start_date.year}년 {start_date.month}월 {start_date.day}일 부터\n")
    p.add_run(f"        ({days}일간)\n")
    p.add_run(f"        {end_date.year}년 {end_date.month}월 {end_date.day}일 까지")

    doc.add_paragraph(" ※ 결석 기간 중 공휴일 또는 학교 휴무일은 결석일 수에 포함하지 않습니다.")

    # 사유
    doc.add_paragraph("\n사  유").bold = True
    doc.add_paragraph(reason)

    # 붙임 서류 체크
    doc.add_paragraph("\n붙  임").bold = True
    def check_box(checked): return "[√]" if checked else "[   ]"

    doc.add_paragraph(f"{check_box(has_doc)} 진단서 또는 진료 확인서(3일 이상인 경우 꼭 첨부)    {check_box(not has_doc)} 없음")
    doc.add_paragraph(f"{check_box(has_rx)} 병원처방전 또는 약봉투")
    doc.add_paragraph(f"{check_box(has_parent)} 보건결석 학부모의견서")
    doc.add_paragraph(f"{check_box(has_etc)} 기타(                )")

    doc.add_paragraph(" ※ 규정된 증빙서류를 첨부하지 않으면 ‘미인정(무단)’ 결석 처리됩니다.")

    # 마무리 문구
    doc.add_paragraph("\n  위와 같이 결석하고자/하였기에 보호자 연서로 신고합니다.")

    doc.add_paragraph(
        f"\n{today.year}년    {today.month}월    {today.day}일"
    )

    doc.add_paragraph(f"\n학  생 성명      {name}           (서명 또는 인)")
    doc.add_paragraph(f"보호자 성명      {parent_name}     (서명 또는 인)")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="결석 신고서 다운로드 (docx)",
        data=buffer,
        file_name=f"결석신고서_{name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
