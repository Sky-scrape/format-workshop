r"""把 tools/launcher.py 冻成包根目录的 格式工坊.exe（软件门面）。

为什么只冻一个十几 MB 的「门面」而不冻整个应用：app.py 用 sys.executable 起
子进程跑 tools/*.py（三层进程模型，见 build_portable.py 头注），冻结后没有
真 Python 可用，进程模型要整体重写。门面只做一件事：拉起旁边
python\pythonw.exe 跑 desktop.py——用户双击的是带图标的 exe，体感即软件，
底下机制零改动。

用法
----
    py -3 tools/build_launcher.py          # 产物 build\格式工坊.exe

build_portable.py 会把 build\格式工坊.exe 复制进便携包根目录（没有就跳过并
提示）。冻结件是未签名二进制，个别杀软可能误报，加白或改发 onedir 形态即可。
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main():
    out = ROOT / "build" / "launcher"
    out.mkdir(parents=True, exist_ok=True)
    print("PyInstaller 冻结门面 …")
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--onefile", "--noconsole",
         "--clean", "--icon", str(ROOT / "static" / "favicon.ico"),
         "--name", "格式工坊",
         "--distpath", str(ROOT / "build"),
         "--workpath", str(out / "work"),
         "--specpath", str(out),
         str(ROOT / "tools" / "launcher.py")],
        check=True)
    exe = ROOT / "build" / "格式工坊.exe"
    print(f"完成：{exe}（{exe.stat().st_size / 1024 / 1024:.1f} MB）")


if __name__ == "__main__":
    main()
