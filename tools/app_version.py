"""软件版本号的唯一读取口：从 app.py 源码正则解析 APP_VERSION。

真源是 app.py 的 APP_VERSION 常量（状态栏/帮助弹窗经 context_processor
直接用它）。构建脚本要同一份号，但不为取版本 import 整个 app——build 跑在
构建机的解释器上，app.py 的依赖（flask、python-docx 等）未必装着，import
会当场炸；正则解析源码一行就够（D-10）。

曾经的三处漂移：build_portable 的 zip 名用的是 --python 参数（3.12.10，
解释器版本），build_installer 用自己的 --version 默认值（1.0.0）——dist 里
「格式工坊-3.12.10.zip」与「格式工坊-setup-1.0.0.exe」并存，谁也说不清
哪个对应对软件的哪个版本。
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def app_version(root=ROOT):
    """读 app.py 源码里的 APP_VERSION 常量；找不到就当场失败，不猜。"""
    src = (Path(root) / "app.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', src, re.M)
    if not m:
        raise SystemExit("app.py 里找不到 APP_VERSION 常量（版本真源）")
    return m.group(1)
