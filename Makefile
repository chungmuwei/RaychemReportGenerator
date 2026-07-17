.PHONY: build clean tag test

# 從 git tag 取得目前版本（沒有 tag 時顯示 dev）
VERSION := $(shell git describe --tags --abbrev=0 2>/dev/null || echo "dev")

build:
	./scripts/build_mac.sh

test:
	PYTHONPATH=src .venv/bin/python -m unittest discover -v

# 打版本 tag 並 push
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
	git fetch origin main:refs/remotes/origin/main
	@if [ "$$(git rev-parse HEAD)" != "$$(git rev-parse origin/main)" ]; then \
		echo "錯誤：本機 main 與 origin/main 不同步，不能建立 release tag。"; \
		echo "請先執行 git pull --ff-only origin main，或先 git push origin main 後再打 tag。"; \
		exit 1; \
	fi
	@if git show-ref --verify --quiet "refs/tags/$(VERSION_ARG)"; then \
		echo "錯誤：本機 tag $(VERSION_ARG) 已存在。"; \
		echo "請改用下一個版本，例如 $(VERSION_ARG)-next，或先刪除本機與遠端舊 tag 後再重建。"; \
		exit 1; \
	fi
	@if git ls-remote --exit-code --tags origin "refs/tags/$(VERSION_ARG)" >/dev/null 2>&1; then \
		echo "錯誤：遠端 tag $(VERSION_ARG) 已存在於 origin。"; \
		echo "若要重新建立同一個 tag，請先刪除本機與遠端舊 tag。"; \
		exit 1; \
	fi
	git tag -a $(VERSION_ARG) -m "Release $(VERSION_ARG)"
	git push origin $(VERSION_ARG)
	@echo "✓ 已推送 tag $(VERSION_ARG)。請執行 make build 建立 macOS DMG。"

clean:
	rm -rf build/ dist/ release/
