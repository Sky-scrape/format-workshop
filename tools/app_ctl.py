"""格式工坊的单一入口：启动 / 停止 / 状态。

为什么逻辑放在 Python 而不是 .bat：cmd 按当前控制台代码页逐字节解析批处理，
.bat 里写中文极不可靠（同一份文件在 chcp 936 与 65001 下都会出现把一行拆成
两条命令、乱码报错等情况）。批处理只保留一行转发，其余全在这里。

用法：
    格式工坊.bat            # 双击 = 桌面壳；bat 直接 start desktop.py，不经这里
    python tools/app_ctl.py # 无参默认同样是拉起桌面壳（launch_shell）
    格式工坊.bat start      # 浏览器服务模式（keepalive + app.py）并打开页面
    格式工坊.bat quiet      # 浏览器服务模式：只启动，不开浏览器
    格式工坊.bat stop       # 停止守护、服务与桌面壳（按启动目录圈定；
                            # 壳先收 WM_CLOSE 走自己的退出语义，超时再强杀）
    格式工坊.bat status     # 状态自检
    格式工坊.bat shortcut   # 桌面建带图标的快捷方式（pythonw 直启，无控制台）

两个 bat 的默认（无参数）都是根目录 desktop.py——pywebview 原生窗口，
页面由进程内引擎应答，不占端口、不起后台服务；带参数（start/stop/status/
shortcut/quiet）转发到这里。下面的端口与守护只服务**可选的浏览器模式**
（python app.py / start、quiet 子命令），桌面壳不使用——但 stop/status 仍
按启动目录圈定并报告桌面壳进程（壳命令行是 pythonw desktop.py，曾圈不住，
覆盖安装/卸载删不动 {app}\\python，D-04）。使用版（便携包/安装版）与开发版
（源码目录 .venv）并存时，浏览器模式用 5000/5001 区分，启动、停止、状态互
不越界（stop 按启动目录圈定进程，见 stop() 与 kill_by_cmdline.ps1）。

打开页面的形态：默认优先用 Edge/Chrome 的 --app 模式弹独立窗口（无地址栏、
任务栏有自己的图标，体感是桌面软件）；找不到这两家浏览器、或设了
PAPER_OPEN=browser 时，回落为系统默认浏览器的标签页。
"""

import http.client
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import app_paths        # noqa: E402  只读安装时日志落用户目录

def data_root():
    """跟随 ROOT 动态解析：测试会 monkeypatch ROOT，不能钉死在 import 时。"""
    return app_paths.data_root(ROOT)


TOOLS = ROOT / "tools"

# 旧版快捷方式（开发版 lnk 曾把端口写进 Arguments：「pythonw app_ctl.py
# --port 5001」——explorer 拉起 lnk 时拿不到创建时的环境变量）还躺在一批
# 桌面上。兼容起见先把 --port 摘出来写回 PAPER_PORT，下面的 PORT/URL 与模式
# 分发就都按它来；再从 sys.argv 删掉，免得「5001」被当成启动模式。新建的
# shortcut 已不写端口（开发版判据见 is_dev_checkout）。
_args = sys.argv[1:]
if "--port" in _args:
    _i = _args.index("--port")
    if _i + 1 < len(_args):
        os.environ["PAPER_PORT"] = _args[_i + 1]
        sys.argv = sys.argv[:1] + _args[:_i] + _args[_i + 2:]

PORT = int(os.environ.get("PAPER_PORT", "5000"))
URL = f"http://127.0.0.1:{PORT}"

# 子进程一律不弹控制台：桌面壳的收尾（停服）经无窗进程调到这里，powershell/
# netstat/taskkill 都是 console 程序，不加 CREATE_NO_WINDOW 每次都闪黑窗。
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def py_exe(name="python.exe"):
    exe = Path(sys.executable)
    p = exe.with_name(name)
    return str(p if p.exists() else exe)


def port_alive(port=PORT, timeout=0.6):
    with socket.socket() as s:
        s.settimeout(timeout)
        return s.connect_ex(("127.0.0.1", port)) == 0


def http_ok(port=PORT):
    try:
        c = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
        c.request("GET", "/")
        return c.getresponse().status
    except Exception:                       # noqa: BLE001
        return None


# app.py /healthz 应答里的身份标记：端口上的监听者报出它，才算自家服务
HEALTHZ_MARK = "geshi-gongfang"


def healthz_says_ours(port=PORT):
    """问端口上跑的是不是本软件的服务：app.py 的 /healthz 带身份标记。

    映像名核对认得出「不是 python」，认不出「5000 上的另一个 python 应用」
    （D-06 的盲区）；/healthz 只有本软件的服务会应答出标记，是最可靠的
    身份证明。探到明确答案返回 True/False；探不到（连接失败、超时——
    占用者不说 HTTP 等）返回 None，调用方退回映像名核对。
    """
    try:
        c = http.client.HTTPConnection("127.0.0.1", port, timeout=3)
        c.request("GET", "/healthz")
        r = c.getresponse()
        body = r.read(256).decode("utf-8", "replace")
        return r.status == 200 and HEALTHZ_MARK in body
    except Exception:                       # noqa: BLE001
        return None


def listener_pids(port=PORT, wildcard=False):
    """按列解析 netstat，返回监听本端口回环可达地址的 PID 集合。

    本软件的服务绑定 127.0.0.1（app.py 的 app.run host 固定），所以本地地址
    做整列精确比对就够。旧兜底是对整行做无边界子串匹配（拿「127.0.0.1:5000」
    在行里找）——监听相邻端口（如 127.0.0.1:50001）的行同样命中，随后不核对
    映像名就 taskkill /F，会强杀无关进程（D-06）；解析不出标准五列的行一律跳过。

    wildcard=True 额外认 0.0.0.0:port 与 [::]:port：绑了通配地址的进程同样
    回应 127.0.0.1 的连接（port_alive 探得到它们），只认 127.0.0.1 会漏掉——
    start() 把陌生占用者报成「服务已在运行」并代开页面、status() 报「无进程
    监听」，正是 D-06 要消灭的冒名相（评审②）。stop 的按端口兜底保持默认
    （wildcard=False）：通配绑定者绝不可能是本软件的服务（host 写死回环），
    不给它被「按端口结束」的机会。绑在具体外网 IP 或 [::1] 的监听不回应
    127.0.0.1，两种口径都不在其列。
    """
    local = {f"127.0.0.1:{port}"}
    if wildcard:
        local |= {f"0.0.0.0:{port}", f"[::]:{port}"}
    out = subprocess.run(["netstat", "-ano"], capture_output=True, text=True,
                         encoding="utf-8", errors="replace",
                         creationflags=CREATE_NO_WINDOW).stdout or ""
    pids = set()
    for line in out.splitlines():
        # 典型行：TCP    127.0.0.1:5000    0.0.0.0:0    LISTENING    4321
        cols = line.split()
        if len(cols) < 5 or cols[3] != "LISTENING":
            continue
        if cols[1] in local:
            pids.add(cols[-1])
    return pids


def _image_of(pid):
    """tasklist 查 PID 的映像名；查不到（刚退出/权限不足）返回 None。

    CSV 单行形如 "python.exe","4321","Console","1"，首个字段即映像名；
    输出编码随控制台代码页，映像名是 ASCII，errors=replace 兜底。
    """
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, errors="replace",
            timeout=15, creationflags=CREATE_NO_WINDOW)
    except Exception:                       # noqa: BLE001
        return None
    m = re.match(r'"([^"]+)"', (r.stdout or "").strip())
    return m.group(1) if m else None


def pid_is_python(pid):
    """动杀之前核对映像名：本软件的进程只会以 python/pythonw 形态在跑。

    F48 给 stop 的 -Root 圈定曾被端口兜底绕掉——兜底按端口抓到的 PID 可能
    属于任何程序，不核对映像就 taskkill /F 会强杀无辜进程（D-06）。拿不准
    （查不到映像）一律不动手：漏杀只是让 stop 如实报「仍被占用」，误杀
    却没法挽回。
    """
    img = _image_of(pid)
    return img is not None and img.lower() in ("python.exe", "pythonw.exe")


def spawn_keepalive():
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    target = str(pyw if pyw.exists() else exe)
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0
    subprocess.Popen(
        [target, str(TOOLS / "keepalive.py"), "--port", str(PORT)],
        cwd=str(ROOT), close_fds=True,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=flags,
    )


def run_ps1_full(name, *args):
    """跑 tools/ 下的 PowerShell 脚本，返回 (stdout, 退出码)。

    kill_by_cmdline.ps1 的约定：无匹配进程时 exit 2 且无输出。stop() 需要
    真实退出码才能如实报「未找到匹配进程」，不能拿空输出当成功（F48）。
    脚本没能执行（powershell 缺失、超时、进程被杀）返回 -1；脚本自身崩溃
    （-File 下抛终止错误）实测退出码是 1、stdout 常为空——旧约定「无匹配
    exit 1」与之同形，崩溃会被报成「未找到匹配进程」，所以无匹配改走 2
    （APPCTLHON）：执行失败与确实无进程是两回事。
    """
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", str(TOOLS / name), *args]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=60,
                           creationflags=CREATE_NO_WINDOW)
        return r.stdout, r.returncode
    except Exception:                       # noqa: BLE001
        return "", -1


def run_ps1(name, *args):
    return run_ps1_full(name, *args)[0]


def find_browser_exe():
    """按常见安装位置找 Edge / Chrome，找不到返回 None。

    Win10/11 必带 Edge，所以这条路几乎总走得通；which 兜住加了 PATH 的场景。
    """
    candidates = []
    for env in ("ProgramFiles(x86)", "ProgramFiles", "LocalAppData"):
        base = os.environ.get(env)
        if not base:
            continue
        candidates.append(Path(base) / r"Microsoft\Edge\Application\msedge.exe")
        candidates.append(Path(base) / r"Google\Chrome\Application\chrome.exe")
    for p in candidates:
        if p.is_file():
            return str(p)
    for name in ("msedge", "chrome"):
        found = shutil.which(name)
        if found:
            return found
    return None


def open_page(url=URL):
    """优先 --app 独立窗口，回落系统默认浏览器。

    --app 是无地址栏、无标签页的独立窗口，任务栏图标取自站点 favicon
    （app.py 的 /favicon.ico → static/favicon.ico），整个像桌面软件；
    webbrowser.open 则在现有浏览器窗口里塞一个标签页，容易被淹没。
    PAPER_OPEN=browser 强制走标签页（调试时看地址栏、开 devtools 方便）。
    """
    if os.environ.get("PAPER_OPEN", "").lower() != "browser":
        exe = find_browser_exe()
        if exe:
            subprocess.Popen([exe, f"--app={url}"],
                             stdin=subprocess.DEVNULL,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
    webbrowser.open(url)


def _ps_quote(s):
    """PowerShell 单引号字符串转义：内部一个 ' 写成两个。"""
    return s.replace("'", "''")


def is_dev_checkout(root=ROOT):
    """开发版判据：源码目录带 .venv——与 desktop.py 的窗口标题后缀同源。

    不再看「PAPER_PORT 是否为 5000」：那是 开发版.bat 顺手 export 的隐性信号，
    开发用户直接调 shortcut 时根本没有，会建出「格式工坊.lnk」盖掉使用版
    的同名快捷方式（D-12）。端口只属于可选的浏览器模式，与壳的形态无关。
    """
    return (root / ".venv").is_dir()


def shortcut():
    """在桌面建「格式工坊」快捷方式（开发环境是「格式工坊·开发版」）。

    直指 pythonw + desktop.py（桌面壳，原生窗口），双击无控制台闪屏；
    图标用 static/favicon.ico。桌面路径让 PowerShell 解析——桌面被 OneDrive
    重定向的用户（%USERPROFILE%\\Desktop 不存在）只有这条路找得对。

    开发环境（源码目录带 .venv，同 desktop.py 的窗口标题判据）lnk 名加
    「·开发版」后缀，不覆盖使用版的同名快捷方式。桌面壳不占端口，
    Arguments 只指向 desktop.py，explorer 拉起 lnk 时拿不到环境变量也不怕。
    """
    dev = is_dev_checkout()
    name = "格式工坊·开发版.lnk" if dev else "格式工坊.lnk"
    exe = ROOT / "格式工坊.exe"
    pyw = ROOT / "python" / "pythonw.exe"
    if not pyw.is_file():
        sibling = Path(sys.executable).with_name("pythonw.exe")
        pyw = sibling if sibling.is_file() else None
    if exe.is_file():
        target, args = str(exe), ""           # 软件门面：真 exe，图标即自身
    elif pyw is not None:
        target = str(pyw)                     # 没打门面时退回直拉桌面壳
        args = f'"{ROOT / "desktop.py"}"'
    else:
        target = str(ROOT / "格式工坊.bat")   # 连 pythonw 都没有：带控制台但能用
        args = ""
    icon = ROOT / "static" / "favicon.ico"
    icon_loc = target if target.endswith(".exe") else (
        f"{icon},0" if icon.is_file() else target)
    ps = (
        "$ws = New-Object -ComObject WScript.Shell;"
        "$d = [Environment]::GetFolderPath('Desktop');"
        f"$s = $ws.CreateShortcut((Join-Path $d '{_ps_quote(name)}'));"
        f"$s.TargetPath = '{_ps_quote(target)}';"
        f"$s.Arguments = '{_ps_quote(args)}';"
        f"$s.WorkingDirectory = '{_ps_quote(str(ROOT))}';"
        f"$s.IconLocation = '{_ps_quote(icon_loc)}';"
        "$s.Save();"
        "Write-Output ok"
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=30, creationflags=CREATE_NO_WINDOW)
    if r.returncode == 0 and r.stdout.strip() == "ok":
        print(f"已在桌面创建快捷方式：{name}")
        return 0
    print(f"快捷方式创建失败（exit {r.returncode}）：{r.stderr.strip()[:200]}")
    return 1


def start(open_browser=True):
    if port_alive():
        # 不能只看 TCP connect 成功就认自家服务：别的程序占了 5000 时，把用户
        # 领去看陌生页面等于冒名。身份核验两道（D-06）：先问 /healthz 的身份
        # 标记——最可靠，映像名核对区分不了「5000 上的另一个 python 应用」；
        # healthz 探不到答案（占用者不说 HTTP 等）才退回映像名核对，至少认出
        # 非 python 形态的占用者。监听者按通配口径看（wildcard=True）：只认
        # 127.0.0.1 会漏掉绑 0.0.0.0/[::] 的占用者（评审②）。
        ours = healthz_says_ours()
        if ours is None:
            foreign = [p for p in sorted(listener_pids(wildcard=True))
                       if not pid_is_python(p)]
            if not foreign:
                ours = True
        if not ours:
            who = "、".join(f"pid={p}（{_image_of(p) or '映像未知'}）"
                            for p in sorted(listener_pids(wildcard=True)))
            detail = f"（{who}）" if who else ""
            print(f"端口 {PORT} 被别的程序占用{detail}，"
                  f"不是格式工坊的服务，不代开页面。")
            return 1
        print(f"服务已在运行：{URL}")
        if open_browser:
            open_page()
        return 0

    print("正在启动服务，请稍候……")
    spawn_keepalive()
    for _ in range(40):                     # 最多等 20 秒
        time.sleep(0.5)
        if port_alive():
            print(f"已启动：{URL}")
            if open_browser:
                open_page()
            return 0
    print("启动失败或超时，请查看 outputs\\keepalive.log")
    return 1


def close_shell_window(timeout=15):
    """按窗口标题请桌面壳自行关闭（WM_CLOSE），窗口不在了返回 True。

    stop 对壳曾是 Stop-Process -Force 直杀：转换进行中弹都不弹就把任务拦腰
    斩断，与 × 的确认语义（D-05）不一致（评审③）。WM_CLOSE 与点 × 是同一条
    消息路径，壳的 closing 处理（忙时原生确认框、几何落盘）照常生效；托盘
    模式的 × 语义是藏进托盘，WM_CLOSE 同样只藏不退——窗口隐藏（或超时、
    脚本没跑成）都按 False 处理，由调用方回落强杀，不改变 stop 的收口能力。

    标题与 desktop.window_title() 同源（is_dev_checkout）：使用版与开发版
    的壳同时开在桌面上时互不误伤。本步静默 best-effort：成败都由
    kill_by_cmdline 的报话收口，这里不插一行，免得报话口径变乱。
    """
    title = f"格式工坊{' · 开发版' if is_dev_checkout() else ''}"
    _, rc = run_ps1_full("close_window_by_title.ps1",
                         "-Title", title, "-TimeoutSec", str(timeout))
    if rc != 0:
        return False
    time.sleep(1)                # 窗口刚关，给进程一点退场时间
    return True


def stop():
    print("正在停止格式工坊……")
    # 先礼后兵：桌面壳有窗口语义（× 的确认框、几何保存，D-05/D-01），直接
    # 强杀会把进行中的转换无确认斩断（评审③）。先按标题发 WM_CLOSE 请它走
    # 自己的退出路径，窗口不在了（或托盘模式把窗藏了起来）再交给下面的
    # pattern 强杀收尾。
    close_shell_window()
    # 三个 pattern 圈定三类进程：守护（keepalive）、浏览器服务（app.py）、
    # 桌面壳（desktop.py）。壳既不监听端口、命令行也不叫 app.py，此前对它
    # 完全失明——status 报「服务未运行」、stop 报「已停止」而壳原地不动，
    # 覆盖安装/卸载因 {app}\python 被运行中的壳占用而删不动（D-04）。
    # -Root 圈定不变（F48）：命令行里带同款字样的别的项目不受影响。
    for label, pattern in (("守护进程", "keepalive"),
                           ("服务进程", r"app\.py"),
                           ("桌面壳进程", r"desktop\.py")):
        # ps1 无匹配 exit 2 且无输出，如实报「未找到」，不再像旧版那样
        # 无论成败都打印「已处理」。
        out, rc = run_ps1_full("kill_by_cmdline.ps1", "-Pattern", pattern,
                               "-Root", str(ROOT))
        killed = [ln.split()[1] for ln in out.splitlines()
                  if ln.strip().startswith("killed ")]
        if killed:
            print(f"  {label}已停止（pid {'、'.join(killed)}）")
        elif rc == 0:
            # 找到了进程但结束失败，ps1 的 failed 行带着原因
            print(f"  {label}结束失败：{out.strip()}")
        elif rc == 2 and not out.strip():
            # ps1 约定：无匹配 exit 2 且无输出——只有这一种情形才算「无进程」。
            # 崩溃的 ps1 在 -File 下退出码是 1、stdout 为空（APPCTLHON 实测），
            # 不能再拿「exit 1 + 空输出」当无匹配的记号
            print(f"  {label}未找到匹配进程")
        else:
            # 脚本本身没跑成（powershell 缺失、超时、崩溃，exit<0 或异常退出
            # 码）——与「无匹配」不是一回事，混报等于把执行失败说成没有进程。
            detail = f"：{out.strip()}" if out.strip() else ""
            print(f"  {label}停止脚本执行失败（exit {rc}）{detail}")
    time.sleep(1)
    if port_alive():
        # 兜底：还有监听者就按端口结束。本地地址按列精确比对（无边界子串会
        # 误中监听相邻端口如 127.0.0.1:50001 的行），动手前再核映像名——
        # 非 python 的占用者不是本软件，不代杀（D-06）。
        for pid in sorted(listener_pids()):
            if pid_is_python(pid):
                subprocess.run(["taskkill", "/PID", pid, "/F"],
                               capture_output=True,
                               creationflags=CREATE_NO_WINDOW)
            else:
                print(f"  端口 {PORT} 的占用者 pid={pid} 不是 python 进程，未代杀。")
    time.sleep(0.5)
    print(f"\n已停止，端口 {PORT} " + ("已释放。" if not port_alive() else "仍被占用。"))
    return 0


def status():
    print("=" * 46)
    print("  格式工坊 · 状态")
    print("=" * 46)
    pids = sorted(listener_pids(wildcard=True))
    if not pids:
        print(f"\n[1] 端口 {PORT}：无进程监听（浏览器服务未运行）")
    else:
        # 端口活着不等于自家服务：报出占用者身份再下结论（D-06）
        note = "有进程在监听"
        foreign = [f"pid={p}（{_image_of(p) or '映像未知'}）"
                   for p in pids if not pid_is_python(p)]
        if foreign:
            note += f"；非 python 占用者：{'、'.join(foreign)}——不是本软件的形态"
        print(f"\n[1] 端口 {PORT}：{note}")
    print("\n[2] 守护、服务与桌面壳进程：")
    # ps1 只输出 pid|role|parent，中文在这里组装，避免 PowerShell 输出编码不一致
    rows = [ln.strip() for ln in run_ps1("proc_report.ps1").splitlines() if ln.count("|") == 2]
    if not rows:
        print("     （无）")
    for row in rows:
        pid, role, parent = row.split("|")
        if role == "keepalive":
            if parent.lower() == "explorer.exe":
                note = "父进程 explorer.exe，可长期驻留"
            elif parent == "-":
                note = "已脱离启动窗口，独立运行"
            else:
                note = f"父进程 {parent}，那个窗口关掉可能一起结束"
            label = "守护"
        elif role == "shell":
            # 桌面壳不监听端口，此前 status 对它完全不显示——壳在跑时整页
            # 输出的是「服务未运行」的故障相（D-04）
            note = "桌面窗口（点 × 或托盘退出；升级/卸载由 stop 圈定结束）"
            label = "壳"
        else:
            note = f"由守护拉起（父进程 {parent}）"
            label = "服务"
        print(f"     pid={pid}  {label}  {note}")
    code = http_ok()
    print("\n[3] HTTP 探测：", f"HTTP {code}  OK" if code else "FAILED - 无法连接")
    print("\n[4] 最近日志：")
    # 壳与守护、服务同为 status 的一等公民（D-04），它的故障留痕
    # （outputs/desktop.log）也要进状态自检，不然壳的异常无处可见（评审⑥）
    for name in ("keepalive.log", "server.log", "desktop.log"):
        p = data_root() / "outputs" / name
        if not p.exists():
            continue
        tail = p.read_text(encoding="utf-8", errors="replace").splitlines()[-3:]
        print(f"  {name}:")
        for line in tail:
            print("     " + line[:100])
    return 0


def launch_shell():
    """无参默认：拉起桌面壳（pythonw desktop.py）。

    与 bat 的无参分支同一形态——壳是常驻窗口进程，detached 拉起后本进程
    立刻返回，控制台不跟着挂住。浏览器服务模式要显式给 start/quiet：
    以前无参默认是 start()（浏览器服务），裸跑 python tools/app_ctl.py 会
    悄悄拉起 keepalive 常驻，与「桌面壳不使用端口/守护」的自述相悖（D-12）。
    """
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0
    subprocess.Popen(
        [py_exe("pythonw.exe"), str(ROOT / "desktop.py")],
        cwd=str(ROOT), close_fds=True,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=flags,
    )
    print("已启动桌面窗口。")
    return 0


def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else "auto").lower()
    if mode in ("stop", "停止"):
        return stop()
    if mode in ("status", "状态"):
        return status()
    if mode in ("quiet", "silent"):
        return start(open_browser=False)
    if mode in ("start", "浏览器"):
        # 浏览器服务模式改为显式子命令：无参默认是桌面壳（launch_shell），
        # 不再静默起 keepalive 常驻服务
        return start(open_browser=True)
    if mode in ("shortcut", "快捷方式"):
        return shortcut()
    return launch_shell()


if __name__ == "__main__":
    sys.exit(main())
