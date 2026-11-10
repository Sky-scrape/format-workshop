"""格式工坊 · 桌面壳：原生窗口 + 进程内引擎，不监听端口。

为什么有这个文件：Flask 应用本体不变，但双击 bat/exe 弹出的不该是浏览器
——pywebview 借 WebView2（Win10/11 自带）画一个真正的原生窗口：没有地址栏、
没有浏览器边框，体感是桌面软件。

引擎形态（2026-11 起）：页面由本进程内的 Flask 直接应答，不再起 keepalive
和 app.py 的 HTTP 服务。WebView2 的 WebResourceRequested 把
http://127.0.0.1/* 拦下来，转成一次 WSGI 调用——没有监听 socket、没有端口、
没有后台服务进程；关窗即退，和其他桌面软件一致。响应在后台线程生成，
WebView2 的 deferral 负责等它完成，界面线程不被慢请求（解析、转换）阻塞。

用法：
    pythonw desktop.py                  # 原生窗口（格式工坊.bat 双击即此）
    python desktop.py                   # 同前，但控制台日志可见
"""

import io
import json
import os
import sys
import threading
import time
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parent

# 便携包用 embeddable Python（_pth 模式）：脚本所在目录不会自动进 sys.path，
# 下面的 import app 会报 ModuleNotFoundError。显式补上，与 tools/*.py 的
# 处理一致；源码运行时 ROOT 本来就在 path 里，重复插入无副作用。
# tools/ 也要进 path：atomic_io（原子写小配置）在那里。
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import app_paths        # noqa: E402  只读安装时把窗口状态/日志落到用户目录

# 站内页面的虚拟主机：请求不出进程，由下面的 WSGI 桥直接应答。不带端口，
# 页面里的相对链接、重定向与 request.host 都按 http://127.0.0.1 走。
LOCAL_URL = "http://127.0.0.1"
LOCAL_HOST = "127.0.0.1"

def data_root():
    """跟随 ROOT 动态解析：测试会 monkeypatch ROOT，不能钉死在 import 时。"""
    return app_paths.data_root(ROOT)


CFG = None


def cfg_file():
    """窗口几何配置文件：每次使用时动态解析（同 data_root() 口径）。

    import 时解析 data_root() 可能两处候选都不可写、当场抛 OSError——模块
    顶层一旦抛，__main__ 的 log_exc 兜底还没进 try，pythonw 下双击即静默
    消失，desktop.log 零痕迹（连日志都落不了，正是 D-02 要消灭的形态）。
    测试显式替换 desktop.CFG 的口径不变：非 None 直接用，None 才动态解析。
    """
    if CFG is not None:
        return CFG
    return data_root() / "config" / "window.json"


# ---------------------------------------------------------------- 壳侧日志
def log_line(text):
    """往 outputs/desktop.log 追加一行（带时间戳）；写不进就放弃，不抛。

    壳进程跑在 pythonw 下没有控制台，任何异常不落盘就是零痕迹；启动链路
    曾是全零日志——install_engine 失败、请求处理失败、窗口几何保存失败都
    静默（D-02）。FAIL 页让用户来查这个文件，它必须真的有内容。
    """
    try:
        (data_root() / "outputs").mkdir(parents=True, exist_ok=True)
        with (data_root() / "outputs" / "desktop.log").open("a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")
    except Exception:                    # noqa: BLE001
        # 数据目录两处候选都不可写时 desktop.log 本身就落不了——主兜底再
        # 兜一层：%TEMP%（用户目录总可写）留同款时间戳一行，异常链至少有
        # 一处可查，绝不回到零痕迹形态（D-02）。
        try:
            import tempfile
            with (Path(tempfile.gettempdir()) / "格式工坊-诊断.log").open(
                    "a", encoding="utf-8", errors="replace") as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")
        except Exception:                # noqa: BLE001
            pass


def log_exc(context, extra=""):
    """异常 + 上下文落 desktop.log（统一取证落点）。必须在 except 块内调用。"""
    import traceback
    log_line(f"{context}{('：' + extra) if extra else ''}\n"
             f"{traceback.format_exc()}")


# ---------------------------------------------------------------- 关窗语义
def close_action():
    """× 按钮的行为，读 config/close.json：exit（默认）| tray。

    收敛后的两种语义（D-05，settings.html 同口径）：exit 下点 × 即退出，
    转换进行中会先弹确认（见 confirm_exit_on_close）；tray 下点 × 驻留
    托盘（图标不就绪时不驻留，见 Tray.on_closing）。曾经承诺的「三选
    对话框（最小化/直接关闭/取消）」从未可达——ask_on_close 只被
    on_closing 调，而 on_closing 只在 tray 模式注册，默认路径没有订阅者，
    已随死代码一并删除。
    """
    try:
        d = json.loads((data_root() / "config" / "close.json").read_text(encoding="utf-8"))
        return "tray" if d.get("action") == "tray" else "exit"
    except Exception:                # noqa: BLE001
        return "exit"


def conversion_busy():
    """引擎里有没有任务正在处理（stage_parse/stage_render 持着任务锁）。

    转换同步跑在请求线程（壳里即 WSGI 桥的 work 线程）：默认路径关窗后
    reap 会 os._exit(0) 硬杀整进程，等几十秒的成果只剩半截 job 目录——
    关窗前必须先问一句引擎忙不忙（D-05）。引擎没起来（启动页阶段）不算忙。
    """
    eng = _ENGINE
    if eng is None:
        return False
    try:
        return bool(eng.conversion_running())
    except Exception:                # noqa: BLE001
        return False


def _confirm_interrupt_box():
    """原生确认框：转换进行中退出会中断转换。返回 True 才放行关窗。

    用 Win32 MessageBoxW 而不是网页 dialog：× 是原生窗口消息，弹原生框
    才同层（网页弹窗盖不住标题栏）。MB_TOPMOST 防被主窗压住点不到。
    """
    if os.name != "nt":
        return True
    import ctypes
    MB_OKCANCEL = 0x1
    MB_ICONWARNING = 0x30
    MB_TOPMOST = 0x40000
    MB_SETFOREGROUND = 0x10000
    text = ("还有转换正在进行中。现在退出会中断转换，\n"
            "这个任务的成果可能只剩半截。\n\n仍要退出吗？")
    r = ctypes.windll.user32.MessageBoxW(
        0, text, window_title(),
        MB_OKCANCEL | MB_ICONWARNING | MB_TOPMOST | MB_SETFOREGROUND)
    return r == 1                    # IDOK=1 放行；IDCANCEL=2 拦下


def confirm_exit_on_close():
    """exit 模式的 closing 处理：忙时先确认，闲时即关即退（D-05）。"""
    if conversion_busy():
        return _confirm_interrupt_box()
    return True


def warn_tray_dead():
    """托盘不可用时的原生告知：程序将退出，而不是藏进空气里（D-11）。"""
    if os.name != "nt":
        return
    import ctypes
    MB_OK = 0x0
    MB_ICONWARNING = 0x30
    ctypes.windll.user32.MessageBoxW(
        0, "托盘图标不可用，程序将退出。\n"
           "可在「设置」里把关窗行为改回「关闭窗口即退出」。",
        window_title(), MB_OK | MB_ICONWARNING | MB_TOPMOST)


# ---------------------------------------------------------------- 单实例
# 窗口标题按版本区分：使用版与开发版可能同时开在桌面上，单实例检查按标题
# 精确匹配——同名会互相把对方认成自己（dist 壳启动时撞上源码壳的窗口就
# 退出，还以为已经聚焦了自己的），这是实测踩过的坑。
def window_title():
    suffix = " · 开发版" if (ROOT / ".venv").is_dir() else ""
    return f"格式工坊{suffix}"


# ---------------------------------------------------------------- 标题栏融色
def tint_title_bar(hwnd, timeout=5.0):
    """把系统标题栏刷成页面顶栏色（DWMWA_CAPTION_COLOR，Win10 1806+/11）。

    为什么不 frameless 自绘：拖动、双击最大化、Win+方向键、窗口阴影、系统
    动画全是系统白送的，自绘要逐个兜回来。DWM 刷色零交互代价——只是把那条
    灰黑系统栏换成与页面同源的底色，视觉上「融」进网页顶栏。

    颜色取 base.html :root 的 --surface(#fff) 与 --ink(#18332e)：标题栏白底、
    墨绿字，与顶栏浑然一体。Windows 老于 1806 时 DWM 调用静默无效，不影响
    启动（那台机器看到原色标题栏，仅此而已）。
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            import ctypes
            user32 = ctypes.windll.user32
            # 窗口刚创建时句柄有效但 DWM 属性可能不生效，重试几次
            if not user32.IsWindow(hwnd):
                time.sleep(0.2)
                continue

            def color_int(rgb_hex):
                r, g, b = (int(rgb_hex[i:i + 2], 16) for i in (1, 3, 5))
                return r | (g << 8) | (b << 16)     # COLORREF 是 0x00BBGGRR

            surface = color_int("#ffffff")          # --surface
            ink = color_int("#18332e")              # --ink
            dwm = ctypes.windll.dwmapi.DwmSetWindowAttribute
            for attr, val in ((35, surface),       # DWMWA_CAPTION_COLOR
                              (36, ink),           # DWMWA_TEXT_COLOR
                              (34, surface)):      # DWMWA_BORDER_COLOR
                dwm(int(hwnd), attr, ctypes.byref(ctypes.c_int(val)), 4)
            return True
        except Exception:            # noqa: BLE001
            time.sleep(0.2)
    return False


def _find_window(title):
    """FindWindowW 按标题精确找顶层窗口，找到返回 True。"""
    import ctypes
    return bool(ctypes.windll.user32.FindWindowW(None, title))


def focus_existing():
    """已有本版壳窗口就带到前台并返回 True——exe 连点两下不该开两个窗。

    用窗口标题精确匹配（FindWindowW），不查进程/端口：同目录两个版本
    可能同时在跑，只有标题能区分是谁的窗口。
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


def _mutex_name():
    """互斥量名与窗口标题同源（带「 · 开发版」后缀）：使用版与开发版并存
    时互不锁——与标题检查同一个坑（见 window_title 注释）。Local\\ 前缀把
    名字圈在本登录会话里，同机多用户各开各的互不干扰。"""
    return "Local\\" + window_title()


def claim_single_instance(name=None):
    """命名互斥量做单实例闸：新建成功返回 True，已存在返回 False。

    标题查询挡不住竞态：FindWindowW 发生在原生窗口创建之前（要等
    webview.start），连点两下两个进程都查不到对方窗口就都继续跑。互斥量
    创建即判定，没有竞态窗口；句柄不关，进程存活期由内核持有，退出（含
    os._exit）即释放，不会留下死锁。name 参数给测试传专名验证机制，不碰
    真名。互斥量建不出来（权限等罕见情形，last error 不是 0/183）退回
    标题检查：宁可有旧竞态，也不把第二个实例挡在门外还聚焦不到旧窗。
    """
    if os.name != "nt":
        return True                     # 非 Windows 没有互斥量，标题检查兜底
    import ctypes
    ERROR_ALREADY_EXISTS = 183
    ctypes.windll.kernel32.CreateMutexW(None, False, name or _mutex_name())
    err = ctypes.GetLastError()
    if err == 0:
        return True                     # 新建：本进程是第一个实例
    if err == ERROR_ALREADY_EXISTS:
        return False                    # 同名互斥量已在：另一个实例在跑
    return not _find_window(window_title())


# ---------------------------------------------------------------- 窗口记忆
_GEOMETRY_SAVED = False


def load_window_state():
    """只记尺寸，不记位置：位置每次交给 Windows 居中级联（SHELLPOS）。

    「记住位置」在小屏单屏机上弊大于利——恢复出来的旧位置永远不如居中
    舒服，还出过「副屏拔掉后窗口落在屏幕外」的事故（已修过一轮越界校验，
    但校验放行的合法旧位置照样贴角落）。Windows 的级联规则自己会把连续
    打开的窗口排得整整齐齐，比记忆可靠。
    """
    try:
        d = json.loads(cfg_file().read_text(encoding="utf-8"))
        w, h = int(d.get("w", 0)), int(d.get("h", 0))
        if 500 <= w <= 8000 and 400 <= h <= 4000:
            return {"w": w, "h": h, "maximized": bool(d.get("maximized"))}
    except Exception:                    # noqa: BLE001
        pass
    return {}


def fit_to_work_area(w, h):
    """窗口尺寸钳制到主屏工作区：默认 1280x860 在 1280x800 的屏上底/右被
    裁掉（复测实测）；来自更大屏幕的记忆尺寸（拔副屏、换分辨率）同样越界。
    工作区不含任务栏；留 32px 余量是给 Windows 居中级联的——贴满工作区的
    窗口被级联一推就出屏。取不到工作区（SPI 失败/异常）按原值放行，宁可
    维持旧行为也不瞎猜。实际生效尺寸还受 min_size 托底，这里不必复刻。
    """
    if os.name != "nt":
        return w, h
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        # SPI_GETWORKAREA：主屏工作区。开窗不传 x/y、由系统居中级联，
        # 窗口落在主屏，主屏工作区就是唯一需要迁就的参照。
        ok = ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
        work_w = int(rect.right) - int(rect.left)
        work_h = int(rect.bottom) - int(rect.top)
        if not ok or work_w <= 0 or work_h <= 0:
            return w, h
        return max(1, min(int(w), work_w - 32)), max(1, min(int(h), work_h - 32))
    except Exception:                    # noqa: BLE001
        return w, h


def current_geometry(window):
    """closing 时刻的窗口几何：宽高趁窗口活着取，最大化用 IsZoomed 实测。

    pywebview 6.2.1 的 winforms on_close 先 del BrowserView.instances[uid]
    再 set closed——在 closed/reap 里读 window.width 是对已销毁实例实时
    解包 get_size 的 None，必抛 TypeError（D-01）；而 maximized 是构造参数
    不是实时状态，存它只会记到启动值，改从原生窗口的 IsZoomed 取。
    """
    try:
        w, h = int(window.width), int(window.height)
    except Exception:                    # noqa: BLE001
        return None
    maximized = False
    if os.name == "nt":
        try:
            import ctypes
            native = getattr(window, "native", None)
            hwnd = int(native.Handle.ToInt64()) if native is not None else 0
            if hwnd:
                maximized = bool(ctypes.windll.user32.IsZoomed(hwnd))
        except Exception:                # noqa: BLE001
            pass
    return {"w": w, "h": h, "maximized": maximized}


def save_window_state(window):
    """窗口尺寸记忆落盘。必须在窗口存活期调用（closing 时实例未注销）：
    旧调用点 closed/reap 都在 FormClosed 之后，width 抛 TypeError 被裸
    except 吞掉，尺寸记忆因此从未落盘过（D-01）。原子写（D-14），失败落
    desktop.log——静默丢的话「下次还是这个尺寸」悄悄失效，查都没处查。
    """
    global _GEOMETRY_SAVED
    geo = current_geometry(window)
    if geo is None:
        log_line("save_window_state：窗口已销毁，宽高取不到，几何不保存"
                 "（几何保存应挂在 closing 上）")
        return
    try:
        import atomic_io
        atomic_io.write_text(cfg_file(), json.dumps(geo, ensure_ascii=False))
        _GEOMETRY_SAVED = True
    except Exception as e:               # noqa: BLE001
        log_line(f"save_window_state failed: {e!r}")


# ---------------------------------------------------------------- 进程内引擎
_ENGINE = None
_ENGINE_LOCK = threading.Lock()
_ENGINE_INSTALLED = set()
_INSTALL_LOCK = threading.Lock()


def engine():
    """导入并初始化 Flask 应用；只做一次。

    应用以 WSGI 可调用对象跑在壳进程里——没有监听 socket，也就没有端口。
    初始化走 app.init_app() 的同一份引导（建目录、清过期任务、装日志、装
    生命周期钩子）：壳曾手抄前三步、漏掉 install_lifecycle_hooks——壳进程
    同住 pythonnet/WebView2 COM、WSGI 与编排 Office 的线程，是原生崩溃
    高发处，faulthandler/crash.log/心跳缺了就是 pythonw 静默消失（D-07）。
    日志落 outputs/desktop.log，server.log 留给监听端口的服务模式——两个
    进程各持 RotatingFileHandler 写同一文件，轮转 rename 在 Windows 上互踩。
    信号处理只在主线程装得上（引导线程里 ValueError 被钩子吞掉），壳拿到
    的是 faulthandler/atexit/心跳三样。
    """
    global _ENGINE
    with _ENGINE_LOCK:
        if _ENGINE is None:
            import app as paper_app
            paper_app.init_app(install_hooks=True, log_name="desktop.log")
            _ENGINE = paper_app
        return _ENGINE


def wsgi_call(method, uri, headers=None, body=b""):
    """把一次浏览器请求交给进程内 Flask，返回 (code, reason, headers, payload)。

    入参与返回值都是纯 Python 类型，可在任意线程调用；只有把结果交给
    WebView2 的那一步必须回 UI 线程（CoreWebView2 的 COM 约束）。
    """
    from urllib.parse import unquote_to_bytes, urlsplit

    parts = urlsplit(uri)
    environ = {
        "REQUEST_METHOD": method,
        "SCRIPT_NAME": "",
        "PATH_INFO": unquote_to_bytes(parts.path).decode("latin-1"),
        "QUERY_STRING": parts.query,
        "SERVER_NAME": LOCAL_HOST,
        "SERVER_PORT": "80",
        "SERVER_PROTOCOL": "HTTP/1.1",
        "REMOTE_ADDR": LOCAL_HOST,
        "wsgi.version": (1, 0),
        "wsgi.url_scheme": "http",
        "wsgi.input": io.BytesIO(body),
        "wsgi.errors": sys.stderr or io.StringIO(),
        "wsgi.multithread": True,
        "wsgi.multiprocess": False,
        "wsgi.run_once": False,
    }
    for key, value in (headers or {}).items():
        name = key.upper().replace("-", "_")
        if name in ("CONTENT_TYPE", "CONTENT_LENGTH"):
            environ[name] = value
        elif name == "HOST":
            environ["HTTP_HOST"] = value
        else:
            environ["HTTP_" + name] = value
    environ.setdefault("HTTP_HOST", LOCAL_HOST)
    if "CONTENT_LENGTH" not in environ and body:
        environ["CONTENT_LENGTH"] = str(len(body))

    captured = {}

    def start_response(status, resp_headers, exc_info=None):
        captured["status"] = status
        captured["headers"] = resp_headers

    chunks = engine().app(environ, start_response)
    payload = b"".join(chunks)
    if hasattr(chunks, "close"):
        chunks.close()
    status = captured.get("status", "500 Internal Server Error")
    code_text, _, reason = status.partition(" ")
    return int(code_text), reason or "OK", captured.get("headers", []), payload


def response_header_text(headers, length, method):
    """WSGI 响应头 → WebView2 的 CRLF 文本；长度按实际字节重算。

    HEAD 没有 body，但 Content-Length 必须保留 Flask 声明的真实长度，
    否则 WebView2 会按 0 处理（静态文件的 HEAD 探测会失真）。
    """
    lines, declared = [], None
    for key, value in headers:
        name = key.lower()
        if name in ("transfer-encoding", "connection"):
            continue
        if name == "content-length":
            declared = value
            continue
        lines.append(f"{key}: {value}")
    if method == "HEAD" and declared is not None:
        lines.append(f"Content-Length: {declared}")
    else:
        lines.append(f"Content-Length: {length}")
    return "\r\n".join(lines) + "\r\n"


def install_engine(window, timeout=25):
    """等 WebView2 就绪，把 LOCAL_URL 下的请求接到进程内 Flask。

    CoreWebView2 只能在 UI 线程访问，所以安装走 native.Invoke；请求本身
    在后台线程跑，跑完再回 UI 线程设置响应并 Complete deferral——解析、
    转换这类慢请求只占一个工作线程，界面不会被冻住。

    重试必须幂等（D-02）：Invoke 是投递到 UI 线程的异步动作，UI 线程繁忙
    时下一个 install() 会在上一个生效前排入队列，不查「已装过」就无条件
    AddWebResourceRequestedFilter + WebResourceRequested += 会装出双处理器，
    站内请求被处理两遍——install() 开头先看 state["ok"]，上一个成功了就
    直接退出。失败各路（install/请求处理/放弃）都落 desktop.log，FAIL 页
    指过来的文件不能是空的。
    """
    key = getattr(window, "uid", id(window))
    with _INSTALL_LOCK:
        if key in _ENGINE_INSTALLED:
            return True

    native = None
    deadline = time.time() + timeout
    while time.time() < deadline:
        native = getattr(window, "native", None)
        if native is not None:
            break
        time.sleep(0.1)
    if native is None:
        log_line(f"install_engine：{timeout}s 内等不到 WebView2 native，放弃")
        return False

    from System import Action
    state = {}

    def install():
        if state.get("ok"):
            return                       # 重试幂等：上一个已装好（D-02）
        try:
            core = native.webview.CoreWebView2
            if core is None:
                return
            from Microsoft.Web.WebView2.Core import CoreWebView2WebResourceContext
            core.AddWebResourceRequestedFilter(
                f"{LOCAL_URL}/*", CoreWebView2WebResourceContext.All)

            def on_request(sender, args):
                req = args.Request
                uri, method = str(req.Uri), str(req.Method)
                headers = {str(h.Key): str(h.Value)
                           for h in req.Headers.GetEnumerator()}
                body = b""
                try:
                    if req.Content is not None:
                        import System
                        ms = System.IO.MemoryStream()
                        req.Content.CopyTo(ms)
                        body = bytes(ms.ToArray())
                except Exception:        # noqa: BLE001
                    body = b""
                deferral = args.GetDeferral()

                def work():
                    try:
                        code, reason, resp_headers, payload = wsgi_call(
                            method, uri, headers, body)
                        header_text = response_header_text(
                            resp_headers, len(payload), method)

                        def apply():
                            try:
                                import System
                                stream = System.IO.MemoryStream(payload)
                                args.Response = core.Environment.CreateWebResourceResponse(
                                    stream, code, reason, header_text)
                            finally:
                                deferral.Complete()

                        native.Invoke(Action(apply))
                    except Exception:    # noqa: BLE001
                        # 只补 Complete 用户看到的是 WebView2 原生错误页，
                        # desktop.log 必须留下一笔才能回溯（D-02）
                        log_exc("请求处理失败", uri)
                        try:
                            native.Invoke(Action(lambda: deferral.Complete()))
                        except Exception:    # noqa: BLE001
                            pass

                threading.Thread(target=work, daemon=True).start()

            core.WebResourceRequested += on_request
            state["ok"] = True
        except Exception as e:           # noqa: BLE001
            state["error"] = e
            log_exc("install：装引擎拦截器失败")

    while time.time() < deadline:
        try:
            native.Invoke(Action(install))
        except Exception:                # noqa: BLE001
            pass
        if state.get("ok"):
            with _INSTALL_LOCK:
                _ENGINE_INSTALLED.add(key)
            return True
        time.sleep(0.2)
    if state.get("error") is not None:
        log_line(f"install_engine 放弃：error={state['error']!r}")
    else:
        log_line("install_engine 放弃：重试期内 WebView2 未就绪")
    return False


def load_app_url(window, url):
    """装好进程内引擎的拦截器，再加载站内页面；失败落错误页。"""
    def run():
        try:
            engine()
            if install_engine(window):
                window.load_url(url)
                return
        except Exception:                # noqa: BLE001
            log_exc("引擎装载失败")
        window.load_html(FAIL)

    threading.Thread(target=run, daemon=True).start()


# ---------------------------------------------------------------- 站内新窗口
def child_url(url_path):
    """把站内相对路径补成本机 URL；外站地址返回 None。

    只认 "/" 开头的路径（"//host" 是协议相对地址，也算外站）。壳里的
    target="_blank" 由页面 JS 拦截后走这里，不让它落到系统浏览器。
    """
    if not isinstance(url_path, str):
        return None
    if not url_path.startswith("/") or url_path.startswith("//"):
        return None
    return f"{LOCAL_URL}{url_path}"


CHILD_TITLES = (
    ("/spec", "格式规格"),
    ("triage.html", "投稿体检"),
    ("audit.html", "审计报告"),
)


def child_title(url_path):
    """子窗口标题要带后缀，绝不能与主窗标题同名。

    主窗的收尾进程靠 FindWindowW 精确匹配标题判断「窗口还在不在」，
    子窗口同名会被误认成主窗，关主窗后进程退不掉。
    """
    for key, label in CHILD_TITLES:
        if key in url_path:
            return f"格式工坊 · {label}"
    return "格式工坊 · 详情"


# ---------------------------------------------------------------- JS 桥
class ShellBridge:
    """给页面提供的原生能力：另存为 + 站内新窗口。

    浏览器式的「下到下载文件夹、重名加 (1)」「新标签页里露出地址栏」都是
    最露馅的网页习惯。桥复用引擎原有的 URL（附件语义不变），只把落盘
    位置和窗口形态换成桌面软件该有的样子。
    """

    def __init__(self, window=None):
        # 桥挂在窗口上；file dialog 要以发起调用的窗口为父窗，不能固定用
        # windows[0]——子窗口里点保存时，主窗可能在最小化或刚被关掉。
        # 必须用下划线属性：pywebview 注入 JS API 时会递归暴露公开属性，
        # 公共名 window 会让它顺着 window.native → WinForms 无限下钻而卡死。
        self._window = window

    def _owner(self):
        # 已知发起窗口就不必 import webview（失败页重试时它总在，但惰性
        # 导入让这条路径在任何环境都可调）
        if self._window:
            return self._window
        import webview
        return webview.windows[0] if webview.windows else None

    def open_in_app(self, url_path):
        """站内 target="_blank" 链接：开壳自己的子窗口，不交给系统浏览器。

        pywebview 默认 OPEN_EXTERNAL_LINKS_IN_BROWSER=True，会把新窗口请求
        丢给 webbrowser.open——用户看到的是一个带 127.0.0.1:端口 的浏览器
        窗口，桌面壳藏了半天的地址栏又露出来了。
        """
        url = child_url(url_path)
        if url is None:
            return {"ok": False, "message": "路径不合法"}
        import webview
        try:
            bridge = ShellBridge()
            # 子窗同样钳到工作区：1000x780 在 1280x800 屏上底边也被裁
            w0, h0 = fit_to_work_area(1000, 780)
            child = webview.create_window(
                child_title(url_path), html=SPLASH, js_api=bridge,
                width=w0, height=h0, min_size=(760, 520),
                text_select=True)
            bridge._window = child
            load_app_url(child, url)
            return {"ok": True}
        except Exception as e:        # noqa: BLE001
            return {"ok": False, "message": str(e)}

    def _fetch(self, url_path):
        """GET 站内下载路由，返回 (字节, 文件名)。文件名取自
        Content-Disposition（attachment() 写的 RFC 5987 中文名）。

        请求直接交给进程内 Flask，不再经 127.0.0.1:端口 的 HTTP 回环。
        """
        code, _, headers, data = wsgi_call(
            "GET", f"{LOCAL_URL}{url_path}", {"Host": LOCAL_HOST})
        if code != 200:
            raise RuntimeError(f"引擎返回 HTTP {code}")
        if not data:
            raise RuntimeError("没有可保存的内容")
        name = None
        disp = next((v for k, v in headers
                     if k.lower() == "content-disposition"), "")
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
            owner = self._owner()
            if owner is None:
                return {"ok": False, "message": "窗口不存在"}
            target = owner.create_file_dialog(
                webview.SAVE_DIALOG, save_filename=name)
            # 返回形状按版本分叉：6.x 是 1-元组（winforms.py「A tuple of
            # selected files」），5.x 是 str，取消统一 falsy（None）。requirements
            # 曾对 pywebview 只写了 >=5 无上限，便携包实际捆的就是 6.x——两种
            # 契约都接住，否则每次另存为都在 Path(元组) 上抛 TypeError。
            if isinstance(target, (tuple, list)) and target:
                target = target[0]
            if not target:
                return {"ok": False}      # 用户取消：静默
            Path(target).write_bytes(data)
            return {"ok": True, "path": str(target)}
        except Exception as e:            # noqa: BLE001
            return {"ok": False, "message": str(e)}

    def retry(self):
        """启动失败页的「重试」：重新装载引擎并回到站内首页。

        失败页曾是死胡同——文案让用户「查看日志后重试」，却没有任何重试
        动作，只能关窗重开（D-13）。重试就是重跑 load_app_url：engine()
        失败不会缓存，import 成功即原地换页。
        """
        win = self._owner()
        if win is None:
            return {"ok": False, "message": "窗口不存在"}
        load_app_url(win, f"{LOCAL_URL}/")
        return {"ok": True}

    def open_outputs(self):
        """启动失败页的「打开日志文件夹」：资源管理器定位 outputs。

        找日志不该要求用户懂目录结构；server.log 只在 import 成功后才创建
        （setup_logging 在 engine() 里），import 阶段失败时按旧文案去找
        server.log 是扑空——干脆打开整个 outputs 文件夹（D-13）。
        """
        out = data_root() / "outputs"
        try:
            out.mkdir(parents=True, exist_ok=True)
            if os.name != "nt":
                return {"ok": False, "message": "仅支持 Windows"}
            os.startfile(out)             # noqa: S606  用户自己点的按钮
            return {"ok": True}
        except Exception as e:            # noqa: BLE001
            return {"ok": False, "message": str(e)}


# ---------------------------------------------------------------- 启动页
SPLASH = ("<!DOCTYPE html><html><head><meta charset='utf-8'></head>"
          "<body style='margin:0'><div style='font-family:Inter,\"PingFang SC\","
          "\"Microsoft YaHei\",sans-serif;display:grid;place-items:center;"
          "height:100vh;background:#f4f6f3'><div style='text-align:center'>"
          "<div style='width:44px;height:44px;margin:0 auto 18px;border-radius:12px;"
          "background:#143b33'></div>"
          "<div style='color:#52655f;font-size:14px'>正在启动格式工坊…</div>"
          "</div></div></body></html>")

FAIL = ("<!DOCTYPE html><html><head><meta charset='utf-8'></head>"
        "<body style='margin:0'><div style='font-family:Inter,\"PingFang SC\","
        "\"Microsoft YaHei\",sans-serif;max-width:560px;margin:84px auto;"
        "padding:0 24px'>"
        "<div style='width:44px;height:44px;margin-bottom:18px;border-radius:12px;"
        "background:#143b33'></div>"
        "<h2 style='color:#18332e;margin:0 0 10px'>软件启动失败</h2>"
        "<p style='color:#52655f;font-size:14px;line-height:1.9;margin:0 0 22px'>"
        "内置引擎没能启动。请点击重试；若反复出现，点下面第二个按钮打开 "
        "outputs 文件夹，把里面的日志发到 GitHub Issues（地址见使用说明.md）。</p>"
        "<div style='display:flex;gap:10px'>"
        "<button id='retry' style='border:0;background:#176b57;color:#fff;"
        "border-radius:9px;padding:10px 22px;font-size:14px;cursor:pointer'>重试</button>"
        "<button id='openLogs' style='border:1px solid #c9d4cd;background:#fff;"
        "color:#52655f;border-radius:9px;padding:10px 18px;font-size:14px;"
        "cursor:pointer'>打开日志文件夹</button></div>"
        "<p id='note' style='color:#9aa69f;font-size:12px;display:none;margin-top:14px'>"
        "</p></div>"
        "<script>"
        "function api(name){var a=(window.pywebview&&pywebview.api)?pywebview.api:null;"
        "if(!a)return Promise.reject(new Error('api not ready'));"
        "return Promise.resolve(a).then(function(x){return x[name]()})}"
        "var busy=false;"
        "document.getElementById('retry').addEventListener('click',function(){"
        "if(busy)return;busy=true;var b=this;b.disabled=true;b.textContent='正在重试…';"
        "api('retry').catch(function(){b.disabled=false;b.textContent='重试';busy=false;"
        "var n=document.getElementById('note');n.textContent='还没就绪，稍候片刻再点一次';"
        "n.style.display='block'})});"
        "document.getElementById('openLogs').addEventListener('click',function(){"
        "api('open_outputs').catch(function(){var n=document.getElementById('note');"
        "n.textContent='打不开文件夹：请手动到数据目录找 outputs（只读安装时在 %LocalAppData%\\\\格式工坊）';"
        "n.style.display='block'})});"
        "</script>"
        "</body></html>")


# ---------------------------------------------------------------- 托盘
class Tray:
    """关窗不退场：窗口收到 closing 时隐藏进托盘，托盘菜单提供退出。

    为什么关窗=隐藏：转换可能还在跑，窗口一关就把进程收掉会中断任务；
    托盘给用户一个「程序还在」的落点。托盘菜单可点开主窗或彻底退出。
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

    def hide_to_tray(self, icon=None, item=None):
        try:
            self.window.hide()
        except Exception:            # noqa: BLE001
            pass

    def quit(self, icon=None, item=None):
        """真退出：窗与托盘一并收掉；引擎跟着进程走，没有后台服务可停。

        转换进行中先弹与 × 同款的原生确认（D-05）：托盘菜单退出同样会经
        reap 的 os._exit 带走整进程，× 有确认而这里不确认，等于留了一条
        无确认硬杀在途转换的路。拒绝则什么都不动，托盘原样驻留。
        """
        if conversion_busy() and not _confirm_interrupt_box():
            return
        self.quitting = True
        save_window_state(self.window)   # destroy 前窗口还活着（D-01）
        try:
            self.window.destroy()
        except Exception:            # noqa: BLE001
            pass
        if self.icon:
            try:
                self.icon.stop()
            except Exception:        # noqa: BLE001
                pass

    def on_closing(self):
        """tray 模式的 closing：驻留托盘；图标不就绪就不驻留（D-11）。

        本处理器只在托盘模式注册（Tray.run ← start_tray ← main 的 tray
        分支）——exit 模式的 × 走 main() 里注册的 confirm_exit_on_close，
        与这里无关。曾经写在里面的「三选对话框」分支因此不可达，已删
        （D-05）。icon.run 在 daemon 线程里失败主流程无感知，不验证就
        hide，窗口隐藏、无图标、无气泡，进程还活着——用户唯一的出路是
        任务管理器。不就绪就不藏，告知后放行关闭。

        放行关闭的每一条路（图标不就绪、配置中途改回 exit）都要趁窗口
        活着先存几何（D-01；评审④）——return True 之后窗口随即销毁，再取
        宽高就是销毁后解包。tray 模式的 main() 没有注册 closing 几何回调，
        这两处就是几何记忆在托盘模式的全部落点。
        """
        if self.quitting:
            return True
        if close_action() == "tray":
            if not self.tray_ready():
                warn_tray_dead()
                # 放行关闭前窗口还活着，几何趁此刻存（评审④补的洞）
                save_window_state(self.window)
                return True               # 放行关闭：退出好过无窗僵尸
            # 驻留前窗口还活着：几何趁此刻存（hide 后再取就是销毁后解包，
            # D-01 的教训——保存必须贴着窗口存活期）
            save_window_state(self.window)
            try:
                self.window.hide()
            except Exception:            # noqa: BLE001
                return True
            self.notify_once()
            return False
        # 配置中途改回 exit（壳还挂着本处理器）：与 × 同样放行，几何同场保存
        save_window_state(self.window)
        return True

    def tray_ready(self):
        """托盘图标真的在跑：icon 已构造且 visible。

        icon.run 在 daemon 线程里失败主流程无感知（explorer 崩溃重启也会丢
        图标）；pystray 的 visible 在托盘线程成功展示后才为 True。
        """
        return self.icon is not None and bool(getattr(self.icon, "visible", False))

    def notify_once(self):
        """首次藏进托盘时气泡告知，点 × 找不到窗口的用户才知道软件还在。

        只提示一次（记 config/tray-hint.flag）——每次都弹就成了骚扰。flag
        在 notify 成功之后才落盘：写反了的话 notify 一抛错 flag 已在盘上，
        此后永不补发气泡（D-11 实核的失败面）。没有 icon 时不写 flag——
        下次藏进托盘（托盘就绪后）再补。
        """
        flag = data_root() / "config" / "tray-hint.flag"
        if flag.exists() or not self.icon:
            return
        try:
            self.icon.notify("程序最小化到了屏幕右下角的托盘"
                             "（可能收在 ^ 里）。点这里或双击图标即可回来。",
                             "格式工坊")
        except Exception:                # noqa: BLE001
            return                       # notify 失败不写 flag，下次再试
        try:
            import atomic_io
            atomic_io.write_text(flag, "ok")
        except Exception:                # noqa: BLE001
            pass

    def run(self):
        import pystray
        menu = pystray.Menu(
            pystray.MenuItem("打开 格式工坊", self.show_window, default=True),
            pystray.MenuItem("隐藏窗口", self.hide_to_tray),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出", self.quit),
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


def webview_udf():
    """WebView2 用户数据目录的固定落点：%LOCALAPPDATA%\\<窗口标题>\\webview。

    pywebview 默认 private_mode=True，WinForms 后端此时把用户数据目录设成
    tempfile.TemporaryDirectory().name——每次启动冷建一套完整浏览器
    profile（实测一个 ~70MB），而 exit 收尾的 os._exit 抢在
    clear_user_data 之前，临时目录基本必留；%TEMP% 残留不属于任何卸载器
    管辖，GPU/代码缓存也无法跨次复用，启动偏慢（D-03）。固定目录后跨次
    复用、二次启动更快。稿件内容不进这里：app.py 对非 /static 响应一律
    Cache-Control: no-store，缓存只落 css/js/图标这类静态资产。使用版与
    开发版按窗口标题判据分目录，两壳并存互不锁。
    """
    base = os.environ.get("LOCALAPPDATA") or (ROOT / "outputs")
    return Path(base) / window_title() / "webview"


def main():
    import webview

    # 单实例两道：互斥量创建即判定（上面的 claim_single_instance），标题
    # 查询只剩「把已有实例带到前台」的职责——它挡不住连点两下（TOCTOU），
    # 但首窗还在启动中时短暂重试等它出现。
    if not claim_single_instance():
        deadline = time.time() + 2.0
        while not focus_existing() and time.time() < deadline:
            time.sleep(0.2)
        return 0

    # 页面里的普通下载（历史页等）交给 WebView2 的原生「另存为」；
    # 结果页的下载卡片仍由 ShellBridge 选位置，两者并存。
    webview.settings["ALLOW_DOWNLOADS"] = True

    # 秒开：先出窗（启动页盖住引擎冷启的 1~3s），就绪后原地换页。
    # 不传 x/y：Windows 居中级联——每次打开都在舒服的位置（SHELLPOS）
    st = load_window_state()
    # 尺寸先过工作区钳制：默认值与来自大屏的旧记忆都可能比当前屏高/宽
    # （复测：1280x860 默认值在 1280x800 屏上底边被任务栏裁掉）
    w0, h0 = fit_to_work_area(st.get("w", 1280), st.get("h", 860))
    bridge = ShellBridge()
    win = webview.create_window(
        window_title(), html=SPLASH, js_api=bridge,
        width=w0, height=h0,
        min_size=(980, 640), text_select=True)
    bridge._window = win

    # ×=退出（默认）路径的收尾：closing 同步触发、窗口实例未注销——几何
    # 保存挂在这里才拿得到宽高（closed/reap 都在 FormClosed 之后，曾让尺寸
    # 记忆 100% 失效，D-01）；转换防护同一时机（D-05）。订阅有前科（注释见
    # 下）：不触发的事件回调与 main() 尾部都可能不跑，所以 reap 线程保留
    # 作兜底，两道都在。
    if close_action() != "tray":
        win.events.closing += confirm_exit_on_close
        win.events.closing += lambda: save_window_state(win)

    def boot():
        if st.get("maximized"):
            try:
                win.maximize()
            except Exception:        # noqa: BLE001
                pass
        # 引擎在本进程内加载，装好请求拦截后再导航；没有端口可连。
        load_app_url(win, f"{LOCAL_URL}/")

    threading.Thread(target=boot, daemon=True).start()

    # 收尾保障（exit 与 tray 模式都要）：不依赖事件回调——实测订阅 closing
    # 后，closed 事件与 main() 尾部代码都可能不触发，壳进程会以无窗状态空转。
    # 唯一可靠的「窗口真没了」信号是 FindWindowW（与单实例同款；native.Visible
    # 在关窗后仍返回 True，是坑）。窗口消失 → 存几何、带走进程。
    # 引擎就在本进程里，没有服务需要另起进程去停。tray 模式曾被排除在外，
    # 但放行关闭（图标不就绪等）同样可能留下无窗进程（评审④）。
    def reap_when_window_gone():
        # 启动页→页面阶段窗口可能重建，先等它出现，出现 3 秒后才开始盯
        title = window_title()
        first_seen = None
        deadline = time.time() + 60
        while time.time() < deadline:
            if _find_window(title):
                if first_seen is None:
                    first_seen = time.time()
                if time.time() - first_seen >= 3:
                    break
            time.sleep(0.5)
        else:
            return               # 窗口迟迟没出现，放弃（别误收）
        while True:
            if not _find_window(title):
                # 连续两次确认（1 秒），防切主题/闪屏误判
                time.sleep(1)
                if not _find_window(title):
                    break
            time.sleep(0.5)
        # tray 模式窗口合法地藏着（WinForms Hide 不销毁 HWND，FindWindow
        # 仍找得到；这里再拦一道是双保险）：托盘还活着且不是在退出，就还有
        # 回得来的路，进程绝不能收——收了等于把驻留中的转换硬杀掉
        tray = _TRAY
        if tray is not None and tray.tray_ready() and not tray.quitting:
            return
        if not _GEOMETRY_SAVED:
            save_window_state(win)   # closing 没跑成时的兜底（D-01）
        stop_tray()
        # os._exit 跳过 atexit（生命周期钩子的正常退出留痕不会跑），
        # 退出这一笔得自己写：事后能确定壳活到了窗口关闭这一刻（D-07）
        log_line(f"窗口已关闭，壳进程退出（os._exit）pid={os.getpid()}")
        os._exit(0)              # 事件循环自己退不干净，直接带走进程
    threading.Thread(target=reap_when_window_gone, daemon=True).start()

    if close_action() == "tray":
        # 托盘只在「关窗驻留」模式下存在；默认语义（×=退出）下它是噪音——
        # 每次启动都多一个图标，菜单里的「退出」和点 × 完全等价
        start_tray(win)

    # 标题栏融色要等原生句柄就绪（webview.start 起事件循环后才有），放线程里刷。
    # WinForms 后端：win.native 是 BrowserForm，句柄在 .Handle——pythonnet 的
    # IntPtr 不能 int() 强转（TypeError），必须走 .ToInt64()（探针实测）
    def tint_when_ready():
        deadline = time.time() + 10
        while time.time() < deadline:
            hwnd = 0
            try:
                native = getattr(win, "native", None)
                if native is not None:
                    hwnd = int(native.Handle.ToInt64())
            except Exception:            # noqa: BLE001
                hwnd = 0
            if hwnd:
                tint_title_bar(hwnd)
                return
            time.sleep(0.2)

    threading.Thread(target=tint_when_ready, daemon=True).start()

    icon = ROOT / "static" / "favicon.ico"
    # 用户数据目录落固定位置（D-03）：private_mode 默认 True 会让 WebView2
    # 每次在 %TEMP% 一次性建全套 profile，os._exit 又抢在清理前杀进程——
    # 残留没人回收、缓存不复用。固定目录跨次复用，二次启动更快。
    udf = webview_udf()
    udf.mkdir(parents=True, exist_ok=True)
    webview.start(icon=str(icon) if icon.is_file() else None,
                  private_mode=False, storage_path=str(udf))
    stop_tray()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:                    # noqa: BLE001
        # pythonw 没有控制台，WebView2 缺失之类的初始化异常只有落盘才能查
        log_exc("壳进程初始化异常")
        raise
