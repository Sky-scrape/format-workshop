"""生成 static/favicon.ico（多尺寸），并顺带输出一张预览 PNG 到 build/。

为什么需要真 ICO
----------------
base.html 里原来挂的是内联 SVG data-URI，浏览器标签页够用；但 app_ctl.py 改成
Edge/Chrome --app 独立窗口后，窗口标题栏与任务栏图标取自站点 favicon，SVG 在
Windows 的窗口壳里支持不稳。ICO 是唯一全场景兜住的格式。

设计完全照搬现有 data-URI SVG（viewBox 24）：#143b33 圆角方块（rx=5）打底，
白色折角文档描边（stroke 1.4），两条白色正文线。先在 60 倍画布上画再 LANCZOS
缩小，小尺寸的锯齿才压得住。

用法
----
    py -3 tools/make_icon.py        # 依赖 Pillow（仅开发机需要，包里用不到）
"""

import sys
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "static" / "favicon.ico"
PREVIEW = ROOT / "build" / "icon-preview.png"

BG = (20, 59, 51, 255)        # #143b33，与页头 brand-mark 同色
FG = (255, 255, 255, 255)
SIZES = (16, 24, 32, 48, 64, 128, 256)
S = 60                        # 超采样倍数：viewBox 24 × 60 = 1440px 画布
STROKE = round(1.4 * S)       # SVG stroke-width 1.4
R = 5 * S                     # 圆角 rx=5


def canvas():
    img = Image.new("RGBA", (24 * S, 24 * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((0, 0, 24 * S - 1, 24 * S - 1), radius=R, fill=BG)

    # 文档主体：白描边矩形，再拿背景色三角形切掉右上角、补上折线，
    # 复现 SVG 里 M7 5h7l3 3v11H7z 的「折角即边线」画法
    box = (7 * S, 5 * S, 17 * S, 19 * S)
    d.rounded_rectangle(box, radius=1 * S, outline=FG, width=STROKE)
    d.polygon([(14 * S, 5 * S), (17 * S, 5 * S), (17 * S, 8 * S)], fill=BG)
    d.line([(14 * S, 5 * S), (17 * S, 8 * S)], fill=FG, width=STROKE)

    # 正文两行：M10 11h4 与 M10 14h4
    for y in (11, 14):
        d.line([(10 * S, y * S), (14 * S, y * S)], fill=FG, width=STROKE)
    return img


def main():
    big = canvas()
    frames = [big.resize((n, n), Image.LANCZOS) for n in SIZES]
    OUT.parent.mkdir(exist_ok=True)
    # 逐张喂现成的小图而不是给一张让 PIL 自己缩：从 256 直接缩到 16 锯齿偏重
    frames[-1].save(OUT, format="ICO", append_images=frames[:-1],
                    sizes=[(n, n) for n in SIZES])
    frames[-1].save(PREVIEW, format="PNG")
    print(f"已生成 {OUT}（{'/'.join(map(str, SIZES))}）")
    print(f"预览 {PREVIEW}")


if __name__ == "__main__":
    sys.exit(main())
