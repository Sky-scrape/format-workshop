"""格式工坊 的 exe 门面：双击它 = 打开软件。

真正的应用 unfrozen 跑在旁边的 python\\pythonw.exe + desktop.py 上（为什么不
把整个应用冻成 exe，见 tools/build_launcher.py 头注与 build_portable.py 头注）。
本程序只负责找到它们并拉起：冻结后 sys.executable 就是本 exe 自身，以其所在
目录为根。

找不到运行件时不弹控制台，用 MessageBox 说清——门面没有第二种输出渠道。
"""

import ctypes
import subprocess
import sys
from pathlib import Path

DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200


def main():
    root = Path(sys.executable).resolve().parent
    pyw = root / "python" / "pythonw.exe"
    target = root / "desktop.py"
    if not (pyw.is_file() and target.is_file()):
        # 源码目录形态：exe 与 .venv 并排时同样可用（开发自测用）
        venv = root / ".venv" / "Scripts" / "pythonw.exe"
        if venv.is_file():
            pyw, target = venv, root / "desktop.py"
    if not (pyw.is_file() and target.is_file()):
        ctypes.windll.user32.MessageBoxW(
            0, "没有找到 python\\pythonw.exe 与 desktop.py。\n"
               "请把本程序放在格式工坊便携包根目录（与 python\\ 文件夹并排）。",
            "格式工坊", 0x10)
        return 1
    subprocess.Popen(
        [str(pyw), str(target)], cwd=str(root),
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP)
    return 0


if __name__ == "__main__":
    sys.exit(main())
