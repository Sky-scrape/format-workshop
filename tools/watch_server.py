"""盯住本地服务：记录它每次上线、消失的时间点，用来判断"断开"到底发生在什么时候。

为什么需要它：服务被外部杀掉时不会留下任何日志（无 traceback、无系统崩溃记录），
事后只能看到"日志停在某条请求"，无法区分是被人 taskkill、被沙箱回收，还是自己退出。
这个脚本从外部轮询端口，把"还活着 / 已消失"的变化写进 outputs/watch.log。

用法（另开一个窗口）：
    python tools/watch_server.py
    python tools/watch_server.py --port 5000 --interval 5
"""

import argparse
import datetime as dt
import socket
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import app_paths        # noqa: E402  只读安装时日志落用户目录

LOG = app_paths.data_root(ROOT) / "outputs" / "watch.log"


def probe(port, timeout=0.5):
    with socket.socket() as s:
        s.settimeout(timeout)
        return s.connect_ex(("127.0.0.1", port)) == 0


def log(msg):
    LOG.parent.mkdir(parents=True, exist_ok=True)
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line, flush=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    ap = argparse.ArgumentParser(description="轮询本地服务并记录上下线")
    ap.add_argument("--port", type=int, default=5000)
    ap.add_argument("--interval", type=float, default=5.0, help="轮询间隔秒数")
    args = ap.parse_args()

    alive = None
    down_since = None
    log(f"=== watchdog 启动，监视 127.0.0.1:{args.port}，每 {args.interval:g} 秒一次 ===")
    try:
        while True:
            now_alive = probe(args.port)
            if alive is None:
                log("服务在线" if now_alive else "服务当前不在线")
            elif now_alive != alive:
                if now_alive:
                    gap = (dt.datetime.now() - down_since).total_seconds() if down_since else 0
                    log(f"服务重新上线（此前中断 {gap:.0f} 秒）")
                    down_since = None
                else:
                    down_since = dt.datetime.now()
                    log("服务已消失 —— 若 app.py 没有留下退出记录，说明进程被外部终止")
            alive = now_alive
            time.sleep(args.interval)
    except KeyboardInterrupt:
        log("=== watchdog 停止 ===")


if __name__ == "__main__":
    main()
