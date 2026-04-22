@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [폴더] %CD%
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js가 설치되어 있지 않거나 PATH에 없습니다.
  echo https://nodejs.org 에서 LTS 버전을 설치한 뒤, PC를 다시 시작하거나 터미널을 새로 여세요.
  pause
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo npm을 찾을 수 없습니다. Node.js를 다시 설치해 주세요.
  pause
  exit /b 1
)

echo [버전 확인]
node -v
npm -v
echo.
echo [npm install 실행 중...]
call npm install
if errorlevel 1 (
  echo.
  echo 설치에 실패했습니다. 위에 나온 빨간색/에러 줄을 복사해 두면 원인 파악에 도움이 됩니다.
  pause
  exit /b 1
)

echo.
echo 설치가 끝났습니다. 다음으로 서버 실행:
echo   npm run serve
echo   또는 브라우저 주소: http://localhost:5173/
pause
