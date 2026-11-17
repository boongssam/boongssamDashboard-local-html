@echo off
chcp 65001 >nul
title 학급 대시보드 교실 서버
cd /d "%~dp0"

where node >/dev/null 2>nul
if errorlevel 1 (
  echo Node.js 가 설치되어 있지 않아요. https://nodejs.org 에서 LTS 버전을 설치한 뒤 다시 실행해 주세요.
  pause
  exit /b 1
)

if not exist "node_modules\react\umd\react.production.min.js" (
  echo 처음 한 번 준비하는 중이에요. 잠시만 기다려 주세요...
  call npm install --no-audit --no-fund
)

rem 서버가 켜질 시간을 2초 준 뒤 선생님 화면을 엶
start "" cmd /c "timeout /t 2 /nobreak >/dev/null & start http://localhost:8080"

echo 이 창을 닫으면 서버가 꺼져요. 수업 중에는 창을 최소화만 해 두세요.
echo.
rem 데이터 폴더: 선생님 앱에서 고른 폴더와 같아야 학생 화면에 데이터가 나옴
node server.js "C:\Users\master\Desktop\개인\개인웹앱\data-dashboard"
echo.
echo 서버가 꺼졌어요. 이미 다른 창에서 켜져 있다면 그쪽을 쓰면 돼요.
pause


\





