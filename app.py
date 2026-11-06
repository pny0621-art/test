# -*- coding: utf-8 -*-
"""
Streamlit 배포 실습용 예제 앱 (연습용 · 가상 데이터)
- 기능 1: 물질명 + ICH M7 Class 입력 → 자동 분류 + json 사례 라이브러리 매칭
- 기능 2: PDF 업로드 → 페이지 수/첫 페이지 텍스트 읽기 (pdfplumber)
- 기능 3: 분류 결과를 Word 보고서(python-docx)로 다운로드
※ API 키, 비밀번호, 실제 기밀자료는 전혀 사용하지 않습니다.
"""
import io
import json
from datetime import date
from pathlib import Path

import streamlit as st
from docx import Document

st.set_page_config(page_title="실습용 Web App", page_icon="🧪", layout="centered")

# ---------- 1. json 사례 라이브러리 불러오기 ----------
CASES_PATH = Path(__file__).parent / "cases.json"


@st.cache_data
def load_cases():
    with open(CASES_PATH, encoding="utf-8") as f:
        return json.load(f)


# ---------- 2. 분류 규칙 (ICH M7 Class → 분류) ----------
def classify(m7_class: str) -> str:
    if m7_class in ("Class 1", "Class 2", "Class 3", "Inconclusive"):
        return "유전독성의심물질"
    return "일반 유연물질"  # Class 4, Class 5


# ---------- 3. Word 보고서 만들기 ----------
def make_report(name: str, m7_class: str, group: str, matched: list) -> bytes:
    doc = Document()
    doc.add_heading("실습용 평가 보고서 (가상 예시)", level=1)
    doc.add_paragraph(f"작성일: {date.today():%Y-%m-%d}")
    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    for i, (k, v) in enumerate([("물질명", name), ("ICH M7 Class", m7_class), ("분류 결과", group)]):
        table.cell(i, 0).text = k
        table.cell(i, 1).text = v
    doc.add_heading("매칭된 사례", level=2)
    if matched:
        for c in matched:
            doc.add_paragraph(f"[{c['출처']}] {c['제목']} - {c['요약']}", style="List Bullet")
    else:
        doc.add_paragraph("매칭된 사례 없음 (일반 유연물질)")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------- 화면 ----------
st.title("🧪 Streamlit 배포 실습 앱")
st.caption("연습용 가상 데이터입니다. 실제 기밀자료는 업로드하지 마세요.")

tab1, tab2, tab3 = st.tabs(["① 자동 분류", "② PDF 읽기", "③ 보고서 다운로드"])

with tab1:
    name = st.text_input("물질명", value="Sample-Nitrosamine-A")
    m7 = st.selectbox("ICH M7 Class", ["Class 1", "Class 2", "Class 3", "Class 4", "Class 5", "Inconclusive"], index=2)
    if st.button("분류하기"):
        group = classify(m7)
        cases = load_cases()
        matched = [c for c in cases if c["대상"] == group]
        st.session_state["result"] = {"name": name, "m7": m7, "group": group, "matched": matched}

    res = st.session_state.get("result")
    if res:
        st.success(f"분류 결과: **{res['group']}**")
        if res["matched"]:
            st.write("매칭된 사례 (cases.json):")
            for c in res["matched"]:
                st.markdown(f"- **[{c['출처']}]** {c['제목']} — {c['요약']}")
        else:
            st.info("매칭할 사례가 없습니다.")

with tab2:
    st.write("PDF를 올리면 페이지 수와 첫 페이지 텍스트를 보여줍니다.")
    pdf = st.file_uploader("PDF 파일 (연습용 문서만)", type="pdf")
    if pdf is not None:
        try:
            import pdfplumber

            with pdfplumber.open(pdf) as doc_pdf:
                st.write(f"총 페이지 수: {len(doc_pdf.pages)}")
                first = doc_pdf.pages[0].extract_text() or "(추출된 텍스트가 없습니다. 스캔본일 수 있어요)"
                st.text_area("첫 페이지 텍스트", first, height=250)
        except Exception as e:
            st.error(f"PDF를 읽는 중 오류가 발생했습니다: {e}")

with tab3:
    res = st.session_state.get("result")
    if not res:
        st.info("먼저 ① 탭에서 '분류하기'를 눌러주세요.")
    else:
        data = make_report(res["name"], res["m7"], res["group"], res["matched"])
        st.download_button(
            "📄 Word 보고서 다운로드",
            data=data,
            file_name="실습_보고서.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
