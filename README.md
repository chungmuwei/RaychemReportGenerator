# 瑞肯COA

瑞肯材料科技有限公司 Certificate of Analysis (COA) 報告產生器。

---

## 專案結構

```
├── app.py                    # 主程式（GUI 入口）
├── generator.py              # 報告產生邏輯（docxtpl）
├── product_specs.json        # 各產品規格資料
├── config.json               # 儲存上次匯出路徑
├── requirements.txt
├── Makefile                  # 建置指令
├── RaychemReport.spec        # PyInstaller macOS 設定
├── RaychemCOA_windows.spec   # PyInstaller Windows 設定
├── templates/                # COA Word 模板
├── fonts/                    # Noto Sans TC 字型
├── icons/                    # 應用程式圖示（.icns / .ico / .svg）
├── scripts/
│   ├── build_mac.sh          # macOS 建置腳本
│   ├── build_win.bat         # Windows 建置腳本
│   └── convert_icon.py       # 圖示轉換工具（.icns → .ico）
└── .github/workflows/
    └── build.yml             # GitHub Actions 自動建置
```

---

## 開發環境設定

### 前置需求

- Python 3.10（建議使用 pyenv）
- macOS（本機開發）

### 建立虛擬環境並安裝依賴

```bash
python3.10 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 執行程式

```bash
python app.py
```

---

## 套件與工具

### Python 套件

| 套件 | 版本 | 用途 |
|---|---|---|
| [dearpygui](https://github.com/hoffstadt/DearPyGui) | 1.8.0 | GUI 框架 |
| [docxtpl](https://docxtpl.readthedocs.io/) | 0.16.4 | Word 模板渲染（Jinja2 語法） |
| [python-docx](https://python-docx.readthedocs.io/) | 0.8.11 | Word 文件操作 |
| [docxcompose](https://github.com/4teamwork/docxcompose) | 1.4.0 | 合併 Word 文件 |
| [python-dateutil](https://dateutil.readthedocs.io/) | 2.9.0 | 日期計算（有效期限推算）|
| Jinja2 | 3.1.2 | 模板引擎（docxtpl 依賴）|
| lxml | 4.9.2 | XML 解析（python-docx 依賴）|

### 建置工具

| 工具 | 用途 |
|---|---|
| [PyInstaller](https://pyinstaller.org/) | 打包成獨立執行檔（.app / .exe）|
| [create-dmg](https://github.com/create-dmg/create-dmg) | 建立 macOS DMG 安裝檔 |
| [GitHub Actions](https://docs.github.com/actions) | CI/CD 自動建置與發佈 |
| Make | 本機建置指令管理 |
| sips（macOS 內建）| 圖示格式轉換（.icns → PNG）|
| Pillow | 圖示尺寸調整並輸出 .ico |

---

## 打包成應用程式

### macOS（.app + DMG）

**前置需求：**

```bash
brew install create-dmg
```

**建置：**

```bash
make build
# 或直接執行
./scripts/build_mac.sh
```

產出：
- `dist/瑞肯COA.app`
- `release/瑞肯COA_<版本>_Installer.dmg`

### Windows（.exe）

在 Windows 機器上執行：

```bat
scripts\build_win.bat
```

產出：`dist\瑞肯COA.exe`

---

## 版本管理與發佈

本專案使用[語意化版本](https://semver.org/lang/zh-TW/)（例如 `v1.0.0`）。

### 發佈新版本

```bash
make tag VERSION_ARG=v1.0.0
```

這個指令會：
1. 在當前 commit 建立 git tag
2. 將 tag push 到 GitHub
3. 自動觸發 GitHub Actions 建置 macOS DMG
4. 在 GitHub Releases 頁面發佈可下載的安裝檔

### 版本號規則

| 版本 | 情境 |
|---|---|
| `v1.0.0` | 正式版 |
| `v1.1.0` | 新增功能 |
| `v1.1.1` | Bug 修正 |
| `v1.2.0-beta` | 預覽版（自動標為 prerelease）|

### 手動觸發建置

在 GitHub → Actions → **Build 瑞肯COA** → **Run workflow**，可輸入版本號並選擇是否發佈 Release。

---

## Makefile 指令

| 指令 | 說明 |
|---|---|
| `make` 或 `make build` | 本機建置 macOS .app + DMG |
| `make clean` | 清除 build/、dist/、release/ |
| `make tag VERSION_ARG=vX.Y.Z` | 打版本 tag 並推送，觸發 CI |
