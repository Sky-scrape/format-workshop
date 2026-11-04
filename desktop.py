"""格式工坊 · 源码版桌面壳：原生窗口替代浏览器页。

为什么有这个文件：Flask 应用本体不变，但双击 开发版.bat 弹出的不该是浏览器
——pywebview 借 WebView2（Win10/11 自带）画一个真正的原生窗口：没有地址栏、
没有浏览器边框，体感是桌面软件。使用版继续走 Edge --app 独立窗口。

进程模型刻意沿用既有链路：窗口只负责「看」，服务仍由 keepalive 常驻守护，
开发版.bat stop 照常能停；窗口随关随走，不留进程。

用法：
    pythonw desktop.py                  # 原生窗口（开发版.bat 双击即此）
    PAPER_PORT=5002 python desktop.py   # 指定端口、控制台日志可见
"""

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent

# 开发版默认 5001（与 开发版.bat 固定值一致）。必须先定端口再 import app_ctl：
# app_ctl 在导入时读 PAPER_PORT 算 PORT，spawn_keepalive 把这个值传给 keepalive，
# 顺序反了守护就会去拉 5000。
os.environ.setdefault("PAPER_PORT", "5001")
PORT = int(os.environ["PAPER_PORT"])

sys.path.insert(0, str(ROOT / "tools"))
import app_ctl  # noqa: E402


def service_ready():
    """HTTP 200 才算就绪：端口被无关进程占着时，不能把别人的页面装进窗口。"""
    import http.client
    try:
        c = http.client.HTTPConnection("127.0.0.1", PORT, timeout=1.5)
        c.request("GET", "/")
        return c.getresponse().status == 200
    except Exception:                       # noqa: BLE001
        return False


def ensure_service():
    if service_ready():
        return True
    app_ctl.spawn_keepalive()
    for _ in range(40):                     # 最多等 20 秒
        time.sleep(0.5)
        if service_ready():
            return True
    return False


def main():
    import webview

    if not ensure_service():
        # pythonw 下 print 无人看见，失败也要用窗口说出来
        webview.create_window(
            "格式工坊 · 启动失败",
            html=f"<div style='font-family:sans-serif;padding:24px'>"
                 f"<h2>服务启动失败</h2>"
                 f"<p>端口 {PORT} 无响应，请查看 outputs\\keepalive.log 与 "
                 f"outputs\\server.log</p></div>",
            width=520, height=240)
        webview.start()
        return 1

    webview.create_window("格式工坊", f"http://127.0.0.1:{PORT}",
                          width=1280, height=860, min_size=(980, 640))
    icon = ROOT / "static" / "favicon.ico"
    webview.start(icon=str(icon) if icon.is_file() else None)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:                       # noqa: BLE001
        # pythonw 没有控制台，WebView2 缺失之类的初始化异常只有落盘才能查
        import traceback
        (ROOT / "outputs").mkdir(parents=True, exist_ok=True)
        with (ROOT / "outputs" / "desktop.log").open("a", encoding="utf-8") as f:
            f.write(traceback.format_exc() + "\n")
        raise
