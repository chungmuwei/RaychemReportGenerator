"""
convert_icon.py — 將 Raychem.icns 轉換為 Raychem.ico（供 Windows 使用）

使用 macOS 內建的 sips 指令，不需額外安裝系統函式庫。
執行前需安裝 Python 依賴：
    pip install Pillow

執行方式（在 Mac 上）：
    python3 convert_icon.py
"""

import sys
import os
import subprocess

def main():
    try:
        from PIL import Image
    except ImportError:
        print("缺少依賴，請先執行：pip install Pillow")
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    icns_path = os.path.join(project_root, "icons", "Raychem.icns")
    ico_path = os.path.join(project_root, "icons", "Raychem.ico")

    if not os.path.exists(icns_path):
        print(f"找不到 {icns_path}")
        sys.exit(1)

    sizes = [16, 32, 48, 256]
    images = []

    print("步驟 1：用 sips 將 .icns 轉為 PNG...")
    base_png = os.path.join(project_root, "icons", "_icon_base.png")
    result = subprocess.run(
        ["sips", "-s", "format", "png", icns_path, "--out", base_png],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"sips 失敗：{result.stderr.strip()}")
        sys.exit(1)
    print("  轉換成功 ✓")

    print("步驟 2：用 Pillow 縮放並合併為 ICO...")
    try:
        base_img = Image.open(base_png).convert("RGBA")
        for size in sizes:
            img = base_img.resize((size, size), Image.LANCZOS)
            images.append(img)
            print(f"  {size}x{size} ✓")

        print(f"步驟 3：儲存 ICO：{ico_path}")
        images[0].save(
            ico_path,
            format="ICO",
            sizes=[(s, s) for s in sizes],
            append_images=images[1:],
        )
    finally:
        if os.path.exists(base_png):
            os.remove(base_png)

    print("完成！Raychem.ico 已建立。")

if __name__ == "__main__":
    main()
