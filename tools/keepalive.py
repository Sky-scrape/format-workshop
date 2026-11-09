"""让 5000 端口一直活着：外部守护进程，只在端口掉了时把服务拉起来。

为什么不用 bat 里的 while 循环：bat 自己也活在被回收的进程树里，一起没了。
这个脚本由 Windows 计划任务（Task Scheduler 服务）启动，父进程不是任何终端或
编辑器，因此不受会话回收影响。

用法：
    pythonw tools/keepalive.py            # 后台常驻（计划任务用这个）
    python tools/keepalive.py              # 前台，能看见日志
    python tools/keepalive.py --interval 3 --port 5000
"""

import argparse
import datetime as dt
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import app_paths        # noqa: E402  只读安装时日志落用户目录

APP = ROOT / "app.py"
LOG = app_paths.data_root(ROOT) / "outputs" / "keepalive.log"


def log(msg):
    LOG.parent.mkdir(parents=True, exist_ok=True)
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    try:
        with LOG.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def port_alive(port, timeout=0.6):
    with socket.socket() as s:
        s.settimeout(timeout)
        return s.connect_ex(("127.0.0.1", port)) == 0


def spawn(port):
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    target = str(pyw if pyw.exists() else exe)
    flags = 0
    if sys.platform == "win32":
        flags = 0x00000008 | 0x00000200        # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(
        [target, str(APP), "--no-browser", "--port", str(port)],
        cwd=str(ROOT), close_fds=True,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=flags,
    )


def parent_info():
    """记录父进程是谁。父进程若是终端/编辑器，进程会在该终端关闭时被一起收走；
    只有父进程是 explorer.exe（即用户从桌面双击启动）才能真正长期驻留。"""
    try:
        import subprocess
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"$p=Get-CimInstance Win32_Process -Filter 'ProcessId={os.getpid()}';"
             "$q=Get-CimInstance Win32_Process -Filter \"ProcessId=$($p.ParentProcessId)\";"
             "Write-Output \"$($q.Name)|$($q.ProcessId)\""],
            capture_output=True, text=True, timeout=10,
            creationflags=0x08000000 if os.name == "nt" else 0).stdout.strip()
        return out or "未知|?"
    except Exception:  # noqa: BLE001
        return "未知|?"


def main():
    ap = argparse.ArgumentParser(description="5000 端口守护")
    ap.add_argument("--port", type=int, default=5000)
    ap.add_argument("--interval", type=float, default=3.0)
    args = ap.parse_args()

    par = parent_info()
    log(f"=== keepalive 启动 pid={os.getpid()} 父进程={par} 端口 {args.port} 每 {args.interval:g} 秒检查 ===")
    if par.split("|")[0].lower() not in ("explorer.exe", "svchost.exe", "taskeng.exe"):
        log("提示：父进程不是 explorer/svchost，用户关掉那个终端窗口时本进程可能被一起结束")
    down_at = None
    restarts = 0
    while True:
        if port_alive(args.port):
            if down_at is not None:
                log(f"端口已恢复（中断 {(dt.datetime.now() - down_at).total_seconds():.0f} 秒）")
                down_at = None
        else:
            if down_at is None:
                down_at = dt.datetime.now()
                log("端口无响应，准备拉起服务")
            time.sleep(1.0)                     # 给刚启动的实例一点时间绑定端口
            if not port_alive(args.port):
                restarts += 1
                log(f"第 {restarts} 次拉起服务")
                try:
                    spawn(args.port)
                except Exception as e:          # noqa: BLE001
                    log(f"拉起失败：{e}")
        time.sleep(args.interval)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("=== keepalive 停止 ===")
