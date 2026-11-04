"""格式工坊的单一入口：启动 / 停止 / 状态。

为什么逻辑放在 Python 而不是 .bat：cmd 按当前控制台代码页逐字节解析批处理，
.bat 里写中文极不可靠（同一份文件在 chcp 936 与 65001 下都会出现把一行拆成
两条命令、乱码报错等情况）。批处理只保留一行转发，其余全在这里。

用法（都由 格式工坊.bat 转发）：
    格式工坊.bat            # 没在跑就静默启动，已在跑就直接打开页面
    格式工坊.bat quiet      # 只启动，不开浏览器（开机自启用）
    格式工坊.bat stop       # 停止守护与服务
    格式工坊.bat status     # 状态自检
    格式工坊.bat shortcut   # 桌面建带图标的快捷方式（pythonw 直启，无控制台）

两个 bat 的默认（无参数）都是根目录 desktop.py——pywebview 原生窗口；
带参数（stop/status/shortcut/quiet）转发到这里。使用版（便携包/安装版，
5000 端口）与开发版（源码目录 .venv，5001 端口）并存，启动、停止、状态
互不越界（stop 按启动目录圈定进程，见 stop() 与 kill_by_cmdline.ps1）。

打开页面的形态：默认优先用 Edge/Chrome 的 --app 模式弹独立窗口（无地址栏、
任务栏有自己的图标，体感是桌面软件）；找不到这两家浏览器、或设了
PAPER_OPEN=browser 时，回落为系统默认浏览器的标签页。
"""

import http.client
import os
import shutil
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"

# 快捷方式的形态是「pythonw app_ctl.py --port 5001」（开发版 shortcut 把端口写
# 进 Arguments——explorer 拉起 lnk 时拿不到创建时的环境变量）。先把 --port 摘
# 出来写回 PAPER_PORT，下面的 PORT/URL 与模式分发就都按它来；再从 sys.argv 删
# 掉，免得「5001」被当成启动模式。
_args = sys.argv[1:]
if "--port" in _args:
    _i = _args.index("--port")
    if _i + 1 < len(_args):
        os.environ["PAPER_PORT"] = _args[_i + 1]
        sys.argv = sys.argv[:1] + _args[:_i] + _args[_i + 2:]

PORT = int(os.environ.get("PAPER_PORT", "5000"))
URL = f"http://127.0.0.1:{PORT}"


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
                           encoding="utf-8", errors="replace", timeout=60)
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


def shortcut():
    """在桌面建「格式工坊」快捷方式（开发环境是「格式工坊·开发版」）。

    直指 pythonw + desktop.py（桌面壳，原生窗口），双击无控制台闪屏；
    图标用 static/favicon.ico。桌面路径让 PowerShell 解析——桌面被 OneDrive
    重定向的用户（%USERPROFILE%\\Desktop 不存在）只有这条路找得对。

    开发环境（PAPER_PORT 不是 5000，比如经 开发版.bat 启动）lnk 名加
    「·开发版」后缀，不覆盖使用版的同名快捷方式。端口不写进 Arguments——
    desktop.py 自带默认（便携包 5000、源码 5001），explorer 拉起 lnk 时
    拿不到环境变量也不怕。
    """
    dev = PORT != 5000
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
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
    if r.returncode == 0 and r.stdout.strip() == "ok":
        print(f"已在桌面创建快捷方式：{name}")
        return 0
    print(f"快捷方式创建失败（exit {r.returncode}）：{r.stderr.strip()[:200]}")
    return 1


def start(open_browser=True):
    if port_alive():
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


def stop():
    print("正在停止格式工坊……")
    for label, pattern in (("守护进程", "keepalive"), ("服务进程", r"app\.py")):
        # -Root 把匹配限定在本仓库路径：命令行里带 app.py 字样的**别的项目**
        # python 进程不再被误杀（F48）。ps1 无匹配 exit 1，如实报「未找到」，
        # 不再像旧版那样无论成败都打印「已处理」。
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
        # 兜底：还有监听者就直接按端口结束
        out = subprocess.run(["netstat", "-ano"], capture_output=True, text=True,
                             encoding="utf-8", errors="replace").stdout
        for line in out.splitlines():
            if f"127.0.0.1:{PORT}" in line and "LISTENING" in line:
                pid = line.split()[-1]
                subprocess.run(["taskkill", "/PID", pid, "/F"],
                               capture_output=True)
    time.sleep(0.5)
    print(f"\n已停止，端口 {PORT} " + ("已释放。" if not port_alive() else "仍被占用。"))
    return 0


def status():
    print("=" * 46)
    print("  格式工坊 · 状态")
    print("=" * 46)
    print(f"\n[1] 端口 {PORT}：", "有进程在监听" if port_alive() else "无（服务未运行）")
    print("\n[2] 守护与服务进程：")
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
        else:
            note = f"由守护拉起（父进程 {parent}）"
            label = "服务"
        print(f"     pid={pid}  {label}  {note}")
    code = http_ok()
    print("\n[3] HTTP 探测：", f"HTTP {code}  OK" if code else "FAILED - 无法连接")
    print("\n[4] 最近日志：")
    for name in ("keepalive.log", "server.log"):
        p = ROOT / "outputs" / name
        if not p.exists():
            continue
        tail = p.read_text(encoding="utf-8", errors="replace").splitlines()[-3:]
        print(f"  {name}:")
        for line in tail:
            print("     " + line[:100])
    return 0


def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else "auto").lower()
    if mode in ("stop", "停止"):
        return stop()
    if mode in ("status", "状态"):
        return status()
    if mode in ("quiet", "silent"):
        return start(open_browser=False)
    if mode in ("shortcut", "快捷方式"):
        return shortcut()
    return start(open_browser=True)


if __name__ == "__main__":
    sys.exit(main())
