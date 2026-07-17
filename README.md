# 瑞肯COA

瑞肯材料科技有限公司 Certificate of Analysis (COA) 報告產生器。

---

## 專案結構

```
├── src/
│   └── raychem_report_generator/
│       ├── __main__.py       # GUI 入口
│       ├── app.py            # Tkinter 介面
│       ├── coa_config.py     # 設定檔與產品資料載入
│       ├── coa_context.py    # 報告資料驗證與組裝
│       ├── coa_utils.py      # 輸入與格式化工具
│       └── generator.py      # docxtpl 報告產生邏輯
├── tests/                    # 單元測試
├── product_specs.json        # 各產品規格資料
├── requirements.txt
├── Makefile                  # 建置指令
├── RaychemReport.spec        # PyInstaller macOS 設定
├── templates/                # COA Word 模板
├── fonts/                    # Noto Sans TC 字型
├── icons/                    # 應用程式圖示（.icns / .svg）
├── scripts/
│   └── build_mac.sh          # macOS 建置腳本
```

---

## 匯出路徑設定檔（macOS）

程式會將「上次匯出路徑」儲存在：

`~/Library/Application Support/com.raychemmaterial.coa/config.json`

---

## 開發環境設定

### 前置需求

- Python 3.13（建議使用 pyenv）
- macOS（本機開發）

### 建立虛擬環境並安裝依賴

```bash
python3.13 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m pip install -e .
```

### 執行程式

```bash
python -m raychem_report_generator
```

### 執行測試

```bash
make test
```

---

## 套件與工具

### Python 套件

| 套件 | 版本 | 用途 |
|---|---|---|
| tkinter | Python standard library | GUI 框架與原生檔案對話框 |
| [docxtpl](https://docxtpl.readthedocs.io/) | 0.20.2 | Word 模板渲染（Jinja2 語法） |
| [python-docx](https://python-docx.readthedocs.io/) | 1.2.0 | Word 文件操作 |
| [docxcompose](https://github.com/4teamwork/docxcompose) | 2.2.0 | 合併 Word 文件 |
| [python-dateutil](https://dateutil.readthedocs.io/) | 2.9.0 | 日期計算（有效期限推算）|
| [PyInstaller](https://pyinstaller.org/) | 6.21.0 | 打包成 macOS .app |
| Jinja2 | 3.1.6 | 模板引擎（docxtpl 依賴）|
| lxml | 6.1.1 | XML 解析（python-docx 依賴）|

### 建置工具

| 工具 | 用途 |
|---|---|
| [PyInstaller](https://pyinstaller.org/) | 打包成 macOS .app |
| [create-dmg](https://github.com/create-dmg/create-dmg) | 建立 macOS DMG 安裝檔 |
| Make | 本機建置指令管理 |

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

## 版本管理與發佈

本專案使用[語意化版本](https://semver.org/lang/zh-TW/)（例如 `v1.0.0`）。

### 發佈新版本

```bash
git checkout main
git pull origin main
make test
make tag VERSION_ARG=v1.0.0
```

這個指令會：
1. 在當前 commit 建立 git tag
2. 將 tag push 到 GitHub
3. 之後可執行 `make build` 在本機建立 macOS DMG

> Release tag 必須指向已合併到 `main` 的 commit；`make tag` 會檢查本機 `main` 是否與 `origin/main` 同步。

### 版本號規則

| 版本 | 情境 |
|---|---|
| `v1.0.0` | 正式版 |
| `v1.1.0` | 新增功能 |
| `v1.1.1` | Bug 修正 |
| `v1.2.0-beta` | 預覽版（自動標為 prerelease）|

## Makefile 指令

| 指令 | 說明 |
|---|---|
| `make` 或 `make build` | 本機建置 macOS .app + DMG |
| `make clean` | 清除 build/、dist/、release/ |
| `make tag VERSION_ARG=vX.Y.Z` | 打版本 tag 並推送 |

## TODO

目前無已知待辦事項。
