"""程序目录只读时的可写数据目录回退。

绿色包/开发版：程序目录可写，任务、日志与配置仍留在程序目录——保持
“解压即用、数据跟着软件走”的直觉。安装到 Program Files 等只读位置时，
这些写入改落 ``%LOCALAPPDATA%\\格式工坊``，普通用户不需要管理员权限。

为什么是“实际写一次”而不是 ``os.access``：Windows 上后者对 ACL 不可靠，
管理员账户在 Program Files 里也常返回可写；只有真正建一个临时文件才
能判断。探测文件立即删除，失败则回退。

环境变量 ``PAPER_DATA_DIR`` 可强制指定数据目录（测试与多份并存时用）。
"""

from __future__ import annotations

import os
import secrets
from pathlib import Path

APP_DIR_NAME = "格式工坊"
DATA_ENV = "PAPER_DATA_DIR"
_CACHE: dict[tuple[str, str, str], Path] = {}


def is_writable(path: Path) -> bool:
    """建目录并试建一个子目录，能建才算数；探测目录立即删除。

    不用 ``tempfile.mkstemp``：在 Windows 的拒绝写 ACL 下，CreateFile 可能
    卡在安全软件扫描里（实测挂住 6 秒以上）；``os.mkdir`` 会立刻回
    PermissionError，判断还更直接。
    """
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError:
        return False

    for _ in range(3):
        probe = path / f".write-probe-{os.getpid()}-{secrets.token_hex(3)}"
        try:
            os.mkdir(probe)
        except FileExistsError:
            continue
        except OSError:
            return False
        try:
            os.rmdir(probe)
        except OSError:
            pass
        return True
    return False


def data_root(root: Path) -> Path:
    """返回可写的数据根目录。

    优先程序目录；不可写时回退到 ``%LOCALAPPDATA%\\格式工坊``。
    两处都不可写才报错——这时没有地方能安全落盘，继续跑只会丢数据。
    结果按（程序目录、覆盖变量、LOCALAPPDATA）缓存：探测要真写文件，
    日志等热路径不能每次都探一遍。
    """
    key = (
        str(Path(root).resolve()),
        os.environ.get(DATA_ENV, ""),
        os.environ.get("LOCALAPPDATA", ""),
    )
    if key in _CACHE:
        return _CACHE[key]

    override = key[1]
    if override:
        candidates = [Path(override).expanduser()]
    else:
        base = os.environ.get("LOCALAPPDATA")
        local = (Path(base) if base else Path.home() / "AppData" / "Local") / APP_DIR_NAME
        candidates = [root, local]

    for target in candidates:
        if is_writable(target):
            # is_writable 在真实环境里已经建了目录；这里再确保一次，
            # 调用方拿到的路径一定可写（测试会替换 is_writable）。
            try:
                target.mkdir(parents=True, exist_ok=True)
            except OSError:
                continue
            _CACHE[key] = target
            return target

    tried = "、".join(str(p) for p in candidates)
    raise OSError(
        f"没有可写的数据目录（已尝试：{tried}）。"
        f"请检查磁盘权限，或用环境变量 {DATA_ENV} 指定一个可写目录。"
    )
