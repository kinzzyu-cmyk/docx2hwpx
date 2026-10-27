---
title: DOCX to HWPX
emoji: 📄
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 7860
pinned: false
---

# DOCX → 한글(HWPX) 변환기

구글 문서·Word(.docx) 파일을 한글(.hwpx)로 바꾸는 웹앱입니다. 파일을 끌어다 놓고 [변환하기]를 누르면 됩니다.

## 사용 방법
1. 구글 문서: 파일 → 다운로드 → Microsoft Word(.docx)
2. 변환기 화면에 .docx 파일을 끌어다 놓기(여러 개 가능)
3. [변환하기] → .hwpx 내려받기(여러 개면 zip)

## 지원하는 서식
- 글자: 굵게·기울임·밑줄·취소선·색·형광펜·크기·글꼴
- 문단: 정렬·들여쓰기·줄 간격·문단 간격
- 표: 셀 병합·배경색·표 안의 표 / 이미지 / 글머리 목록
- 머리글·바닥글·쪽 번호(첫 쪽·홀짝 쪽 다르게, 구역별)
- 각주·미주(안의 서식·표·이미지 포함)
- 구역별 용지 크기·방향·여백·다단, 세로쓰기, 줄 번호, 쪽 테두리, 쪽 배경색

## 직접 실행 (Windows)
- 내 PC에서만: `run.bat` 더블클릭
- 같은 사무실 사람들과 함께: `run_server.bat` 더블클릭 → 화면에 나온 주소를 공유
- 필요: Python 3.10~3.13 (https://www.python.org, 설치 때 "Add python.exe to PATH" 체크)

## 개인정보
올린 파일은 변환하는 동안 임시 폴더에서만 처리하고 변환이 끝나면 지웁니다. 서버에 저장하지 않습니다.

## 라이선스
- 이 변환기 코드: 배포하는 분이 정하세요(예: MIT).
- 사용 라이브러리: python-hwpx(Apache-2.0), python-docx(MIT), Streamlit(Apache-2.0)
