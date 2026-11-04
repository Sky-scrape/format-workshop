"""格式工坊 · 桌面壳：原生窗口替代浏览器页（使用版与开发版同一入口）。

为什么有这个文件：Flask 应用本体不变，但双击 bat/exe 弹出的不该是浏览器
——pywebview 借 WebView2（Win10/11 自带）画一个真正的原生窗口：没有地址栏、
没有浏览器边框，体感是桌面软件。

进程模型刻意沿用既有链路：窗口只负责「看」，服务仍由 keepalive 常驻守护，
bat 的 stop 照常能停；窗口随关随走，不留进程。

用法：
    pythonw desktop.py                  # 原生窗口（格式工坊.bat 双击即此）
    PAPER_PORT=5002 python desktop.py   # 指定端口、控制台日志可见
"""

import json
import os
import sys
import threading
import time
from http.client import HTTPConnection
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parent

# 端口默认值按环境判：便携包（无 .venv）服务常驻 5000；源码目录（有 .venv）
# 开发版在 5001（开发版.bat 也会显式设 PAPER_PORT）。explorer 拉起快捷方式时
# 提供不了环境变量，只能靠这个判据。必须先定端口再 import app_ctl：app_ctl
# 在导入时读 PAPER_PORT 算 PORT，spawn_keepalive 把它传给 keepalive，顺序反了
# 便携包的守护就会去拉 5001。
os.environ.setdefault(
    "PAPER_PORT", "5001" if (ROOT / ".venv").is_dir() else "5000")
PORT = int(os.environ["PAPER_PORT"])

sys.path.insert(0, str(ROOT / "tools"))
import app_ctl  # noqa: E402

CFG = ROOT / "config" / "window.json"


# ---------------------------------------------------------------- 单实例
# 窗口标题按版本区分：使用版与开发版可能同时开在桌面上，单实例检查按标题
# 精确匹配——同名会互相把对方认成自己（dist 壳启动时撞上源码壳的窗口就
# 退出，还以为已经聚焦了自己的），这是实测踩过的坑。
def window_title():
    suffix = " · 开发版" if (ROOT / ".venv").is_dir() else ""
    return f"格式工坊{suffix}"


def focus_existing():
    """已有本版壳窗口就带到前台并返回 True——exe 连点两下不该开两个窗。

    用窗口标题精确匹配（FindWindowW），不查端口：服务由 keepalive 常驻，
    端口活着不代表窗口开着（用户可能只关了窗、留着服务）。
    """
    if os.name != "nt":
        return False
    import ctypes
    title = window_title()
    hwnd = ctypes.windll.user32.FindWindowW(None, title)
    if not hwnd:
        return False
    user32 = ctypes.windll.user32
    user32.ShowWindow(hwnd, 9)          # SW_RESTORE：最小化时还原
    user32.SetForegroundWindow(hwnd)
    return True


# ---------------------------------------------------------------- 窗口记忆
def load_window_state():
    try:
        d = json.loads(CFG.read_text(encoding="utf-8"))
        w, h = int(d.get("w", 0)), int(d.get("h", 0))
        # 明显损坏的尺寸回默认；坐标不校（副屏拔掉后 OS 层能兜住）
        if 500 <= w <= 8000 and 400 <= h <= 4000:
            return {"w": w, "h": h,
                    "x": int(d.get("x") or 0), "y": int(d.get("y") or 0),
                    "maximized": bool(d.get("maximized"))}
    except Exception:                    # noqa: BLE001
        pass
    return {}


def save_window_state(window):
    try:
        CFG.parent.mkdir(parents=True, exist_ok=True)
        CFG.write_text(json.dumps({
            "w": window.width, "h": window.height,
            "x": window.x, "y": window.y, "maximized": window.maximized,
        }, ensure_ascii=False), encoding="utf-8")
    except Exception:                    # noqa: BLE001
        pass


# ---------------------------------------------------------------- 另存为桥
class SaveBridge:
    """给页面提供「原生另存为」：页面拦截下载点击 → 这里选位置 → 落盘。

    浏览器式的「下到下载文件夹、重名加 (1)」是最露馅的网页习惯。桥复用
    服务端原有的下载 URL（附件语义不变），只是把落盘位置交还给用户。
    """

    def _fetch(self, url_path):
        """GET 服务端下载路由，返回 (字节, 文件名)。文件名取自
        Content-Disposition（attachment() 写的 RFC 5987 中文名）。"""
        conn = HTTPConnection("127.0.0.1", PORT, timeout=120)
        conn.request("GET", url_path)
        resp = conn.getresponse()
        if resp.status != 200:
            raise RuntimeError(f"服务端返回 HTTP {resp.status}")
        data = resp.read()
        if not data:
            raise RuntimeError("没有可保存的内容")
        name = None
        disp = resp.getheader("Content-Disposition") or ""
        for part in disp.split(";"):
            part = part.strip()
            if part.startswith("filename*="):
                name = unquote(part.split("''", 1)[1])
                break
            if part.startswith('filename="'):
                name = part[len('filename="'):-1]
        return data, name or "未命名"

    def save_as(self, url_path):
        import webview
        try:
            if not url_path.startswith("/"):
                return {"ok": False, "message": "路径不合法"}
            data, name = self._fetch(url_path)
            windows = webview.windows
            if not windows:
                return {"ok": False, "message": "窗口不存在"}
            target = windows[0].create_file_dialog(
                webview.SAVE_DIALOG, save_filename=name)
            if not target:
                return {"ok": False}      # 用户取消：静默
            Path(target).write_bytes(data)
            return {"ok": True, "path": str(target)}
        except Exception as e:            # noqa: BLE001
            return {"ok": False, "message": str(e)}


# ---------------------------------------------------------------- 服务拉起
def service_ready():
    try:
        c = HTTPConnection("127.0.0.1", PORT, timeout=1.5)
        c.request("GET", "/")
        return c.getresponse().status == 200
    except Exception:                    # noqa: BLE001
        return False


def ensure_service():
    """就绪返回 True；首轮未就绪先拉起 keepalive（幂等，start 同口径）。"""
    for attempt in range(40):            # 最多 20 秒
        if service_ready():
            return True
        if attempt == 0:
            try:
                app_ctl.spawn_keepalive()
            except Exception:            # noqa: BLE001
                pass
        time.sleep(0.5)
    return False


SPLASH = ("<!DOCTYPE html><html><head><meta charset='utf-8'></head>"
          "<body style='margin:0'><div style='font-family:Inter,\"PingFang SC\","
          "\"Microsoft YaHei\",sans-serif;display:grid;place-items:center;"
          "height:100vh;background:#f4f6f3'><div style='text-align:center'>"
          "<div style='width:44px;height:44px;margin:0 auto 18px;border-radius:12px;"
          "background:#143b33'></div>"
          "<div style='color:#52655f;font-size:14px'>正在启动格式工坊…</div>"
          "</div></div></body></html>")

FAIL = ("<!DOCTYPE html><html><head><meta charset='utf-8'></head>"
        "<body style='margin:0'><div style='font-family:sans-serif;padding:28px'>"
        "<h2>服务启动失败</h2><p>端口 {port} 无响应，请查看 outputs\\keepalive.log "
        "与 outputs\\server.log，或运行 格式工坊.bat stop 后重试。</p></div>"
        "</body></html>")


# ---------------------------------------------------------------- 托盘
class Tray:
    """关窗不退场：窗口收到 closing 时隐藏进托盘，托盘菜单提供退出。

    为什么关窗=隐藏：服务由 keepalive 常驻，窗口一关用户就失去了「软件
    还在跑」的可见性，任务栏里也没了落点。托盘把控制权还回来：点开主窗、
    直接停服、或彻底退出（窗+托盘；服务留给下次秒开）。
    """

    def __init__(self, window):
        self.window = window
        self.icon = None
        self.quitting = False

    def _image(self):
        from PIL import Image
        ico = ROOT / "static" / "favicon.ico"
        if ico.is_file():
            return Image.open(ico)
        return Image.new("RGB", (64, 64), "#143b33")

    def show_window(self, icon=None, item=None):
        w = self.window
        try:
            w.show()
            w.restore()
        except Exception:            # noqa: BLE001
            pass

    def stop_service(self, icon=None, item=None):
        try:
            app_ctl.stop()
        except Exception:            # noqa: BLE001
            pass

    def quit(self, icon=None, item=None):
        self.quitting = True
        self.stop_service()
        try:
            self.window.destroy()
        except Exception:            # noqa: BLE001
            pass
        if self.icon:
            self.icon.stop()

    def on_closing(self):
        """closing 事件返回 False 取消关闭 → 隐藏进托盘；真退出放行。"""
        if self.quitting:
            return True
        try:
            self.window.hide()
        except Exception:            # noqa: BLE001
            return True              # 隐藏不了就让它关
        return False

    def run(self):
        import pystray
        menu = pystray.Menu(
            pystray.MenuItem("打开 格式工坊", self.show_window, default=True),
            pystray.MenuItem("停止后台服务", self.stop_service),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出（并停止服务）", self.quit),
        )
        self.icon = pystray.Icon(window_title(), self._image(), window_title(), menu)
        self.window.events.closing += self.on_closing
        threading.Thread(target=self.icon.run, daemon=True).start()

    def stop(self):
        if self.icon:
            try:
                self.icon.stop()
            except Exception:        # noqa: BLE001
                pass


_TRAY = None


def start_tray(window):
    global _TRAY
    try:
        _TRAY = Tray(window)
        _TRAY.run()
    except Exception:                # noqa: BLE001
        # 托盘起不来（缺 pystray/Pillow 等）不挡主流程：关窗即退出
        _TRAY = None


def stop_tray():
    if _TRAY:
        _TRAY.stop()


def main():
    import webview

    if focus_existing():
        return 0

    # 秒开：先出窗（启动页盖住服务冷启的 2~4s），就绪后原地换页
    st = load_window_state()
    win = webview.create_window(
        window_title(), html=SPLASH, js_api=SaveBridge(),
        width=st.get("w", 1280), height=st.get("h", 860),
        x=st.get("x") or None, y=st.get("y") or None,
        min_size=(980, 640), text_select=True)

    win.events.closed += lambda: save_window_state(win)

    def boot():
        if ensure_service():
            if st.get("maximized"):
                try:
                    win.maximize()
                except Exception:        # noqa: BLE001
                    pass
            # ?from=<port> 只是页面初载的壳信号；站内跳转丢失后由
            # window.pywebview 轮询兜底（base.html enableShellUI）
            win.load_url(f"http://127.0.0.1:{PORT}/?from={PORT}")
        else:
            win.load_html(FAIL.format(port=PORT))

    threading.Thread(target=boot, daemon=True).start()
    start_tray(win)

    icon = ROOT / "static" / "favicon.ico"
    webview.start(icon=str(icon) if icon.is_file() else None)
    stop_tray()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:                    # noqa: BLE001
        # pythonw 没有控制台，WebView2 缺失之类的初始化异常只有落盘才能查
        import traceback
        (ROOT / "outputs").mkdir(parents=True, exist_ok=True)
        with (ROOT / "outputs" / "desktop.log").open("a", encoding="utf-8") as f:
            f.write(traceback.format_exc() + "\n")
        raise
