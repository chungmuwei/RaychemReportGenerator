.PHONY: build clean tag

# 從 git tag 取得目前版本（沒有 tag 時顯示 dev）
VERSION := $(shell git describe --tags --abbrev=0 2>/dev/null || echo "dev")

build:
	./scripts/build_mac.sh

# 打版本 tag 並 push → 自動觸發 GitHub Actions build
# 用法：make tag VERSION=v1.0.0
tag:
	@if [ "$(VERSION_ARG)" = "" ]; then \
		echo "用法：make tag VERSION_ARG=v1.0.0"; exit 1; \
	fi
	@if ! printf '%s\n' "$(VERSION_ARG)" | grep -Eq '^v[0-9]+\.[0-9]+\.[0-9]+([-.][0-9A-Za-z.-]+)?$$'; then \
		echo "錯誤：版本號格式錯誤：$(VERSION_ARG)"; \
		echo "請使用 vMAJOR.MINOR.PATCH，例如 v1.1.0 或 v1.1.0-beta.1"; \
		exit 1; \
	fi
	@if [ "$$(git branch --show-current)" != "main" ]; then \
		echo "錯誤：請先切到 main 並 pull 最新版本後再建立 release tag。"; exit 1; \
	fi
	@if ! git diff --quiet || ! git diff --cached --quiet; then \
		echo "錯誤：工作目錄尚有未提交變更，請先 commit 或 stash。"; exit 1; \
	fi
	git tag -a $(VERSION_ARG) -m "Release $(VERSION_ARG)"
	git push origin $(VERSION_ARG)
	@echo "✓ 已推送 tag $(VERSION_ARG)，GitHub Actions 正在建立 release..."

clean:
	rm -rf build/ dist/ release/
