@echo off
REM build_win.bat — 在 Windows 上建立 瑞肯COA.exe
REM
REM 前置需求：
REM   1. 安裝 Python 3.13（與 Mac 版相同版本）
REM   2. 建立 venv：python -m venv .venv
REM   3. 安裝依賴：.venv\Scripts\python.exe -m pip install -r requirements.txt
REM
REM 執行方式（從專案根目錄或 scripts\ 目錄皆可）：
REM   scripts\build_win.bat

setlocal enabledelayedexpansion

REM 無論從哪裡執行，都切換到專案根目錄
cd /d "%~dp0.."

set APP_NAME=瑞肯COA
set SPEC_FILE=RaychemCOA_windows.spec
set VENV_PYTHON=.venv\Scripts\python.exe

echo ==========================================
echo  瑞肯COA Windows Build Script
echo ==========================================

REM 確認 venv 存在
if not exist "%VENV_PYTHON%" (
    echo 錯誤：找不到 %VENV_PYTHON%
    echo 請先建立 venv 並安裝依賴：
    echo   python -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    exit /b 1
)

REM 確認 PyInstaller 可用（使用 python -m，避免 venv 搬移後 pyinstaller.exe 路徑失效）
"%VENV_PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤：.venv 中未安裝 PyInstaller，請執行：
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    exit /b 1
)

REM 確認 Raychem.ico 存在
if not exist "icons\Raychem.ico" (
    echo 錯誤：找不到 icons\Raychem.ico
    echo 請先在 Mac 上執行 scripts\convert_icon.py 產生圖示。
    exit /b 1
)

REM 清除舊的 build / dist
echo.
echo [1/2] 清除舊的 build 產物...
if exist build\ rmdir /s /q build
if exist dist\ rmdir /s /q dist

REM 執行 PyInstaller
echo.
echo [2/2] 執行 PyInstaller...
"%VENV_PYTHON%" -m PyInstaller "%SPEC_FILE%" --clean

if errorlevel 1 (
    echo 錯誤：PyInstaller 執行失敗
    exit /b 1
)

echo.
echo ==========================================
echo  完成！
echo  輸出目錄 : dist\%APP_NAME%\
echo  執行檔   : dist\%APP_NAME%\%APP_NAME%.exe
echo ==========================================
