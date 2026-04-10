@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================================
echo   시험 문제집 HWPX 생성기 (hwpx-bookmaker) 설치
echo ============================================================
echo.

set "INSTALL_DIR=%~dp0"
set "VENV_DIR=%INSTALL_DIR%.venv"

:: Python 확인
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo        https://www.python.org/downloads/ 에서 Python 3.10 이상을 설치하세요.
    pause
    exit /b 1
)

:: Python 버전 확인
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [정보] Python %PYVER% 확인됨

:: 가상환경 생성
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo [1/3] 가상환경 생성 중...
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 (
        echo [오류] 가상환경 생성 실패
        pause
        exit /b 1
    )
) else (
    echo [1/3] 기존 가상환경 사용
)

:: 의존성 설치
echo [2/3] 의존성 설치 중...
"%VENV_DIR%\Scripts\pip.exe" install -q -r "%INSTALL_DIR%requirements.txt"
if %errorlevel% neq 0 (
    echo [오류] 의존성 설치 실패
    pause
    exit /b 1
)

echo [3/3] 설치 완료!
echo.
echo ============================================================
echo   Claude Desktop 설정 (claude_desktop_config.json)
echo ============================================================
echo.
echo 아래 내용을 Claude Desktop 설정에 추가하세요:
echo.
echo {
echo   "mcpServers": {
echo     "hwpx-bookmaker": {
echo       "command": "%VENV_DIR:\=\\%\\Scripts\\python.exe",
echo       "args": ["%INSTALL_DIR:\=\\%server.py"]
echo     }
echo   }
echo }
echo.
echo ============================================================
echo.
pause
