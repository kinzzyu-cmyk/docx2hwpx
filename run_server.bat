@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo 같은 사무실(같은 네트워크) 사람들이 이 PC의 변환기를 쓰도록 서버로 실행합니다.
python -m pip install --quiet --no-warn-script-location -r requirements.txt
if errorlevel 1 (
  echo 패키지 설치에 실패했습니다.
  pause
  exit /b 1
)
echo.
echo 아래 "IPv4 주소"를 확인해 다른 사람에게 http://IPv4주소:8501 을 알려 주세요.
ipconfig | findstr /i "IPv4"
echo.
echo 처음 실행할 때 Windows 방화벽 창이 뜨면 [허용]을 누르세요. 끝내려면 Ctrl+C
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
