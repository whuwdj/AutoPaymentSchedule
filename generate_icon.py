# -*- coding: utf-8 -*-
"""生成应用图标 app.ico（需 Pillow）。运行：python generate_icon.py"""

from __future__ import annotations

import os

from PIL import Image, ImageDraw, ImageFont


def main() -> None:
    root = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(root, "app.ico")
    size = 256
    img = Image.new("RGBA", (size, size), (37, 99, 235, 255))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((8, 8, size - 8, size - 8), radius=48, fill=(255, 255, 255, 38))
    draw.rounded_rectangle((24, 24, size - 24, size - 24), radius=40, fill=(255, 255, 255, 255))
    try:
        font = ImageFont.truetype("arial.ttf", 72)
    except OSError:
        font = ImageFont.load_default()
    text = "付"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(((size - tw) // 2, (size - th) // 2 - 8), text, fill=(37, 99, 235, 255), font=font)
    img.save(out, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
    print("written:", out)


if __name__ == "__main__":
    main()
