#!/bin/bash
# build_mac.sh — 在 macOS 上建立 瑞肯COA.app 並打包為 DMG
#
# 前置需求：
#   brew install create-dmg
#
# 執行方式（從專案根目錄或 scripts/ 目錄皆可）：
#   ./scripts/build_mac.sh

set -e  # 任何指令失敗就停止

# 無論從哪裡執行，都切換到專案根目錄
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT"

APP_NAME="瑞肯COA"
SPEC_FILE="RaychemReport.spec"
VENV_PYTHON=".venv/bin/python"
DIST_DIR="dist"
PYINSTALLER_CONFIG_DIR="${PROJECT_ROOT}/.pyinstaller"

# 版本：優先用環境變數，否則從 git tag 自動取得
APP_VERSION="${APP_VERSION:-$(git describe --tags --abbrev=0 2>/dev/null || echo 'dev')}"
DMG_OUTPUT="release/${APP_NAME}_${APP_VERSION}_Installer.dmg"

echo "=========================================="
echo " 瑞肯COA macOS Build Script"
echo "=========================================="

# 確認 venv python 存在
if [ ! -x "$VENV_PYTHON" ]; then
    echo "錯誤：找不到可執行的 $VENV_PYTHON，請先建立 venv 並安裝依賴。"
    exit 1
fi

# 確認 PyInstaller 可用（使用 python -m，避免 venv 搬移後 shebang 路徑失效）
if ! "$VENV_PYTHON" -m PyInstaller --version > /dev/null 2>&1; then
    echo "錯誤：.venv 中未安裝 PyInstaller，請執行：.venv/bin/pip install pyinstaller"
    exit 1
fi

# 確認 create-dmg 存在
if ! command -v create-dmg &> /dev/null; then
    echo "錯誤：找不到 create-dmg，請執行：brew install create-dmg"
    exit 1
fi

# 清除舊的 build / dist
echo ""
echo "[1/3] 清除舊的 build 產物..."
rm -rf "build/" "${DIST_DIR}/" || true

# 執行 PyInstaller
echo ""
echo "[2/3] 執行 PyInstaller..."
mkdir -p "$PYINSTALLER_CONFIG_DIR"
PYINSTALLER_CONFIG_DIR="$PYINSTALLER_CONFIG_DIR" "$VENV_PYTHON" -m PyInstaller "$SPEC_FILE" --clean

APP_PATH="${DIST_DIR}/${APP_NAME}.app"
if [ ! -d "$APP_PATH" ]; then
    echo "錯誤：PyInstaller 完成但找不到 ${APP_PATH}"
    exit 1
fi
echo "  App bundle 建立成功：${APP_PATH}"

# 建立 DMG
echo ""
echo "[3/3] 建立 DMG 安裝檔..."
mkdir -p release
rm -f "$DMG_OUTPUT"

# 建立暫存目錄，只包含 .app，避免 dist/ 中的資料夾也被放入 DMG
STAGING_DIR=$(mktemp -d)
cp -r "${APP_PATH}" "${STAGING_DIR}/"

codesign --force --deep --sign - "${APP_PATH}"

create-dmg \
    --volname "${APP_NAME}" \
    --volicon "icons/Raychem.icns" \
    --window-pos 200 120 \
    --window-size 600 400 \
    --icon-size 128 \
    --icon "${APP_NAME}.app" 175 190 \
    --hide-extension "${APP_NAME}.app" \
    --app-drop-link 425 190 \
    "$DMG_OUTPUT" \
    "${STAGING_DIR}/"

rm -rf "${STAGING_DIR}"

echo ""
echo "=========================================="
echo " 完成！"
echo " App bundle : ${APP_PATH}"
echo " DMG 安裝檔 : ${DMG_OUTPUT}"
echo "=========================================="
