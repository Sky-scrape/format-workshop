"""格式工坊的单一入口：启动 / 停止 / 状态。

为什么逻辑放在 Python 而不是 .bat：cmd 按当前控制台代码页逐字节解析批处理，
.bat 里写中文极不可靠（同一份文件在 chcp 936 与 65001 下都会出现把一行拆成
两条命令、乱码报错等情况）。批处理只保留一行转发，其余全在这里。

用法（都由 格式工坊.bat 转发）：
    格式工坊.bat            # 没在跑就静默启动，已在跑就直接打开页面
    格式工坊.bat quiet      # 只启动，不开浏览器（开机自启用）
    格式工坊.bat stop       # 停止守护与服务
    格式工坊.bat status     # 状态自检
"""

import http.client
import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
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
        [target, str(TOOLS / "keepalive.py")],
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


def start(open_browser=True):
    if port_alive():
        print(f"服务已在运行：{URL}")
        if open_browser:
            webbrowser.open(URL)
        return 0

    print("正在启动服务，请稍候……")
    spawn_keepalive()
    for _ in range(40):                     # 最多等 20 秒
        time.sleep(0.5)
        if port_alive():
            print(f"已启动：{URL}")
            if open_browser:
                webbrowser.open(URL)
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
    return start(open_browser=True)


if __name__ == "__main__":
    sys.exit(main())
