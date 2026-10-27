"""DOCX → HWPX 변환 웹앱 (Streamlit, 드래그 앤 드롭)  v12

v2: 굵게·기울임·밑줄·취소선·글자 색·형광펜·글자 크기, 표 셀 병합·셀 배경색
v3: 글꼴 이름, 문단 정렬, 들여쓰기
v4: 표 안의 표, 머리글·바닥글(쪽 번호), 각주
v5: 줄 간격, 문단 앞/뒤 간격, 제목 크기·표 행 높이를 원본에 맞춤
v6: 미주, 첫 쪽 다른 머리글(표지), 홀수/짝수 쪽 다른 머리글, 쪽 나누기
v7: 각주·미주 안 글자 서식(굵게·색·형광펜 등)과 여러 문단 유지
v8: 여러 구역 문서의 구역별 머리글·바닥글, 쪽 번호 새로 시작
v9: 각주·미주 안 표·이미지
v10: 구역별 용지 크기·방향·여백, 다단
v11: 세로쓰기, 줄 번호, 쪽 테두리
v12: 문단별 줄 번호 건너뛰기, 그림 쪽 테두리(선으로 근사), 쪽 배경색

실행 (Windows 명령 프롬프트):
    python -m streamlit run app.py
"""
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from docx2hwpx import convert

st.set_page_config(page_title="DOCX → 한글(HWPX) 변환", page_icon="📄", layout="centered")
st.title("DOCX → 한글(HWPX) 변환기")
st.caption("구글 문서는 파일 → 다운로드 → Microsoft Word(.docx)로 받은 뒤 아래에 끌어다 놓으세요.")

with st.sidebar:
    st.header("변환 옵션")
    keep_format = st.toggle(
        "글자 서식·셀 배경색 유지", value=True,
        help="글자 서식(굵게·색·형광펜·크기·글꼴), 문단 서식(정렬·들여쓰기·줄 간격·문단 간격), "
             "표 셀 배경색을 옮깁니다. 끄면 글자만 옮깁니다(제목·표 셀 병합은 항상 유지).",
    )
    header_footer = st.toggle("머리글·바닥글 옮기기", value=True,
                              help="머리글·바닥글의 글자와 쪽 번호(현재 쪽 / 전체 쪽 수)를 옮깁니다. "
                                   "첫 쪽 다른 머리글, 홀수/짝수 쪽 다른 머리글도 옮깁니다.")
    footnote_label = st.radio(
        "각주·미주 처리 방식", ["한글 각주·미주로 넣기", "본문 [1] 표시 + 문서 끝 목록"], index=0,
        help="한글 각주·미주: 한글에서 각주/미주로 편집할 수 있습니다(미주 번호는 i, ii, iii). "
             "문서 끝 목록: 각주([1])와 미주([i])를 일반 글로 문서 마지막에 모읍니다(어떤 프로그램에서도 똑같이 보임).",
    )
    st.markdown(
        "**지원 서식**\n"
        "- 굵게 · 기울임 · 밑줄 · 취소선 · 글자 색 · 형광펜 · 글자 크기 · 글꼴\n"
        "- 정렬 · 들여쓰기(첫 줄/내어쓰기)\n"
        "- 줄 간격(배수·고정·최소) · 문단 앞/뒤 간격\n"
        "- 표 셀 병합 · 셀 배경색 · 표 안의 표\n"
        "- 머리글 · 바닥글 · 쪽 번호 · 각주 · 미주(안의 글자 서식·표·이미지 포함)\n"
        "- 첫 쪽(표지) 다른 머리글 · 홀수/짝수 쪽 다른 머리글 · 쪽 나누기\n"
        "- 여러 구역 문서의 구역별 머리글·바닥글 · 쪽 번호 새로 시작\n"
        "- 구역별 용지 크기·방향(세로/가로)·여백 · 다단(구분선·단별 폭)\n"
        "- 세로쓰기 · 줄 번호(문단별 건너뛰기 포함) · 쪽 테두리(선 종류·굵기·색) · 쪽 배경색\n"
        "- 제목 · 글머리 목록 · 이미지\n\n"
        "**근사 변환**: 그림 쪽 테두리(사과·별 등) → 비슷한 선, 쪽 배경 그림·그라데이션 → 기본 색"
    )

FOOTNOTE_MODE = {"한글 각주·미주로 넣기": "note", "본문 [1] 표시 + 문서 끝 목록": "text"}

files = st.file_uploader(
    "여기에 .docx 파일을 끌어다 놓거나 클릭해서 선택하세요",
    type=["docx"],
    accept_multiple_files=True,
)


def convert_one(data: bytes, keep_format: bool, footnote_mode: str = "note", header_footer: bool = True):
    """업로드된 docx 바이트 → (hwpx 바이트, 통계). 임시 폴더에서만 처리."""
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = Path(tmp) / "in.docx", Path(tmp) / "out.hwpx"
        src.write_bytes(data)
        if not zipfile.is_zipfile(src):
            raise ValueError("올바른 DOCX 파일이 아닙니다.")
        stats = convert(str(src), str(dst), keep_format=keep_format,
                        footnote_mode=footnote_mode, header_footer=header_footer)
        return dst.read_bytes(), stats


def summary(stats: dict) -> str:
    parts = [f"제목 {stats['heading']}", f"문단 {stats['para']}", f"표 {stats['table']}",
             f"이미지 {stats['image']}"]
    if stats.get("nested_table"):
        parts.append(f"표 안의 표 {stats['nested_table']}")
    if stats.get("merged"):
        parts.append(f"병합 셀 {stats['merged']}")
    if stats.get("footnote"):
        parts.append(f"각주 {stats['footnote']}")
    if stats.get("endnote"):
        parts.append(f"미주 {stats['endnote']}")
    if stats.get("note_table") or stats.get("note_image"):
        parts.append(f"각주·미주 안 표 {stats.get('note_table', 0)} · 이미지 {stats.get('note_image', 0)}")
    if stats.get("header") or stats.get("footer"):
        hf = "머리글·바닥글"
        extra = []
        if stats.get("first_page"):
            extra.append(f"첫 쪽 {stats['first_page']}")
        if stats.get("odd_even"):
            extra.append("홀짝 쪽 구분")
        if stats.get("sections"):
            extra.append(f"구역 {stats['sections']}개")
        if stats.get("page_restart"):
            extra.append(f"쪽 번호 새로 시작 {stats['page_restart']}곳")
        parts.append(hf + (f"({', '.join(extra)})" if extra else ""))
    layout = []
    if stats.get("landscape"):
        layout.append(f"가로 용지 구역 {stats['landscape']}")
    if stats.get("columns"):
        layout.append(f"다단(최대 {stats['columns']}단)")
    if stats.get("page_setup"):
        layout.append(f"용지·여백 설정 {stats['page_setup']}종")
    if stats.get("vertical"):
        layout.append(f"세로쓰기 구역 {stats['vertical']}")
    if stats.get("line_number"):
        layout.append(f"줄 번호 구역 {stats['line_number']}")
    if stats.get("page_border"):
        layout.append(f"쪽 테두리 구역 {stats['page_border']}")
    if stats.get("lnnum_skip"):
        layout.append(f"줄 번호 건너뛴 문단 {stats['lnnum_skip']}")
    if stats.get("page_fill"):
        layout.append(f"쪽 배경색 {stats['page_fill']}")
    if stats.get("art_border"):
        layout.append(f"그림 테두리 → 선으로 근사({stats['art_border']})")
    if layout:
        parts.append(" · ".join(layout))
    if stats.get("styled_runs"):
        parts.append(f"서식 적용 {stats['styled_runs']}곳")
    if stats.get("note_styled"):
        parts.append(f"각주·미주 서식 {stats['note_styled']}곳")
    return " · ".join(parts)


if files and st.button("변환하기", type="primary"):
    results = []
    progress = st.progress(0.0)
    for i, f in enumerate(files, 1):
        out_name = Path(f.name).stem + ".hwpx"
        try:
            hwpx, stats = convert_one(f.getvalue(), keep_format,
                                      FOOTNOTE_MODE[footnote_label], header_footer)
            results.append((out_name, hwpx))
            st.success(f"{f.name} → {out_name}  ({summary(stats)})")
        except Exception as e:
            st.error(f"{f.name}: 변환 실패 — {type(e).__name__}: {e}")
        progress.progress(i / len(files))
    st.session_state["results"] = results

results = st.session_state.get("results", [])
if len(results) == 1:
    name, data = results[0]
    st.download_button(f"⬇ {name} 내려받기", data, file_name=name, mime="application/hwp+zip")
elif len(results) > 1:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in results:
            z.writestr(name, data)
    st.download_button(f"⬇ 변환 파일 {len(results)}개 ZIP으로 내려받기", buf.getvalue(),
                       file_name="hwpx_변환결과.zip", mime="application/zip")

st.divider()
st.caption("결과는 한컴 '한글'에서 한 번 열어 확인하세요. 올린 파일은 변환하는 동안만 임시로 처리하고 서버에 저장하지 않습니다.")
