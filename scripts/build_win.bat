@echo off
REM build_win.bat — 在 Windows 上建立 瑞肯COA.exe
REM
REM 前置需求：
REM   1. 安裝 Python 3.10（與 Mac 版相同版本）
REM   2. 建立 venv：python -m venv venv
REM   3. 安裝依賴：venv\Scripts\pip install -r requirements.txt pyinstaller
REM
REM 執行方式（從專案根目錄或 scripts\ 目錄皆可）：
REM   scripts\build_win.bat

setlocal enabledelayedexpansion

REM 無論從哪裡執行，都切換到專案根目錄
cd /d "%~dp0.."

set APP_NAME=瑞肯COA
set SPEC_FILE=RaychemCOA_windows.spec
set VENV_PYINSTALLER=venv\Scripts\pyinstaller.exe

echo ==========================================
echo  瑞肯COA Windows Build Script
echo ==========================================

REM 確認 venv 存在
if not exist "%VENV_PYINSTALLER%" (
    echo 錯誤：找不到 %VENV_PYINSTALLER%
    echo 請先建立 venv 並安裝依賴：
    echo   python -m venv venv
    echo   venv\Scripts\pip install -r requirements.txt pyinstaller
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
"%VENV_PYINSTALLER%" "%SPEC_FILE%" --clean

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
