@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [1/2] 필요한 패키지 확인/설치 중...
python -m pip install --quiet --no-warn-script-location -r requirements.txt
if errorlevel 1 (
  echo 패키지 설치에 실패했습니다. 인터넷 연결과 Python 설치를 확인하세요.
  pause
  exit /b 1
)
echo [2/2] 변환기 실행 중... 브라우저가 자동으로 열립니다. 끝내려면 이 창에서 Ctrl+C
python -m streamlit run app.py
pause
