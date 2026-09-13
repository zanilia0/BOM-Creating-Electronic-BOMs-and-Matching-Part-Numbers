"""Build the open-source desktop app. Author: 桀桀 / WeChat JJ-Linnnnn."""
from pathlib import Path
import os
import subprocess
import sys
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent


def main():
    os.chdir(ROOT)
    target = ROOT / "build"
    target.mkdir(exist_ok=True)
    icon = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(icon)
    draw.rounded_rectangle((0, 0, 255, 255), radius=48, fill="#1677D2")
    for y in (48, 112, 176):
        draw.rounded_rectangle((40, y, 80, y + 32), radius=6, fill="white")
        draw.rounded_rectangle((104, y, 216, y + 32), radius=6, fill="white")
    icon_path = target / "bom.ico"
    icon.save(icon_path, sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])
    subprocess.run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
        "--onefile", "--windowed", "--name", "BOM物料库匹配工具_开源版",
        "--icon", str(icon_path), *sys.argv[1:], "ad_bom_matcher.py",
    ], check=True)


if __name__ == "__main__":
    main()
