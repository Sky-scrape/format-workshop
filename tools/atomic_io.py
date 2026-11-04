"""原子写小配置文件：先写临时文件，再 os.replace 原子替换。

为什么要有这一层：进程被硬杀是预期场景（app.py read_json_safe 的注释原话
「半截 JSON 会留下」），读侧早就做了容错，写侧却是裸 write_text——硬杀
正好落在写入中段，close.json/window.json/ai.json 留半截 JSON，用户看到的
是「设置静默回默认」。os.replace 在同一目录内是原子的：替换成功前旧文件
完好，替换失败也不碰旧文件。

调用方：app.py（close.json）、desktop.py（window.json）、ai_assist.py
（ai.json）。写 JSON 的 dumps 仍在调用方——这里只管「怎么写」，不管
「写什么」。
"""

import os
import time
import uuid
from pathlib import Path


def _replace_with_retry(src, dst):
    """os.replace 的 Windows 有限退避：并发的两个替换者在同一目标上短暂
    互斥，MoveFileEx 偶发 PermissionError（FX1 实测 0.247s 内退得过去）。
    参数与 app.py 的 replace_with_retry 同款——同一结论不抄成两样。"""
    delays = (0.02, 0.04, 0.08, 0.1, 0.2)
    for attempt in range(5):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(delays[attempt])


def write_text(path, text, encoding="utf-8"):
    """把 text 原子写到 path（自动建父目录），返回写入的 Path。

    临时名带 pid + 随机后缀：pid 区分两个进程（比如使用版与开发版各开一个
    壳），随机后缀区分同进程的并发写者——同进程两个线程写同一目标（如
    threaded 服务下两个 /settings/shell 并发、closing 与 reap 兜底同时
    save_window_state）共用一个固定临时名会互相踩，先落位的 os.replace 把
    临时文件挪走，另一个写者对已不存在的路径替换，当场 FileNotFoundError。
    app.py 的 write_json_atomic 为同一类并发把临时名改成随机后缀在先（评审⑤），
    这层 helper 必须同一结论。失败时 finally 把临时文件收走，不留垃圾。
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(f"{p.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp")
    try:
        tmp.write_text(text, encoding=encoding)
        _replace_with_retry(tmp, p)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    return p
