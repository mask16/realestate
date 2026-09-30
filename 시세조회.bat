@echo off
chcp 65001 > nul
cd /d "%~dp0"
python 시세조회.py
if errorlevel 1 (
  echo.
  echo 실행에 실패했습니다. Python과 아래 라이브러리를 확인하세요.
  echo python -m pip install -r requirements.txt
  pause
)
