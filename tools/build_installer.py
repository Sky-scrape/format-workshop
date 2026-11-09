r"""把 dist/格式工坊 打成 setup.exe 安装包（Inno Setup 6 + tools/installer.iss）。

前置
----
    python tools/build_portable.py     # 先有绿色包，安装包就是把它铺进用户目录

ISCC 从哪找：--iscc 参数 → 环境变量 INNO_SETUP → %LocalAppData%\Programs\Inno Setup 6
（官方安装器 /CURRENTUSER 的默认落点，不需要管理员）→ Program Files 两处。
机器上没装时，官方下载入口 https://jrsoftware.org/isdl.php，静默装到用户目录：
    innosetup-6.7.4.exe /VERYSILENT /CURRENTUSER

用法
----
    py -3 tools/build_installer.py                # 版本默认取 app.py 的 APP_VERSION
    py -3 tools/build_installer.py --version 1.1  # 也可手动指定（发版脚本用）

产物
----
    dist/格式工坊-setup-<版本>.exe
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

import app_version as appver
from package_docs import strip_private_docs

ROOT = Path(__file__).resolve().parent.parent
ISS = ROOT / "tools" / "installer" / "installer.iss"


def iscc_ready_source() -> Path:
    r"""把 .iss 规整成 ISCC 吃的形态（UTF-8 BOM + CRLF）再交给它。

    ISCC 按 CRLF 切指令：LF 结尾的 .iss 里，指令值末尾会带进一个不可见
    字符——LicenseFile=..\..\LICENSE 看着路径全对，却报「找不到文件」，
    排查极绕，索性在构建时统一转，仓库里存什么行尾都无所谓（git 的
    autocrlf 也伤不到）。

    副本必须和 ChineseSimplified.isl 同目录，且 [Setup] 里的 ..\..\ 是按
    「仓库根下两层」写死的（tools/installer），所以副本就落在那儿：
    tools/installer/installer-gen.iss（已进 .gitignore）。
    """
    work = ISS.parent
    isl = work / "ChineseSimplified.isl"
    if not isl.is_file():
        raise SystemExit(f"缺少 {isl}——中文界面语言包应随仓库走")
    text = ISS.read_text("utf-8-sig")
    p = work / "installer-gen.iss"
    p.write_bytes(b"\xef\xbb\xbf" + text.replace("\r\n", "\n").replace("\n", "\r\n")
                  .encode("utf-8"))
    return p


def find_iscc(explicit=None):
    if explicit:
        p = Path(explicit)
        if not p.is_file():
            raise SystemExit(f"--iscc 指定的文件不存在：{p}")
        return str(p)
    candidates = []
    if os.environ.get("INNO_SETUP"):
        candidates.append(Path(os.environ["INNO_SETUP"]) / "ISCC.exe")
    # 自选目录安装（本机是 D:\Inno Setup 7）：扫盘符根；新版优先
    for drive in "CDEFGH":
        for ver in (7, 6):
            candidates.append(Path(f"{drive}:\\") / f"Inno Setup {ver}" / "ISCC.exe")
    local = os.environ.get("LocalAppData")
    if local:
        candidates.append(Path(local) / "Programs" / "Inno Setup 6" / "ISCC.exe")
    for env in ("ProgramFiles(x86)", "ProgramFiles"):
        base = os.environ.get(env)
        if base:
            candidates.append(Path(base) / "Inno Setup 6" / "ISCC.exe")
    for p in candidates:
        if p.is_file():
            return str(p)
    raise SystemExit(
        "找不到 ISCC.exe（Inno Setup 编译器）。\n"
        "  官方下载：https://jrsoftware.org/isdl.php\n"
        "  或指位置：--iscc \"D:\\Inno Setup 7\\ISCC.exe\"")


def main():
    ap = argparse.ArgumentParser()
    # 版本真源在 app.py 的 APP_VERSION（经 tools/app_version.py 解析）；
    # 曾经默认值在这里另写一份 1.0.0，两处一漂移产物就对不上号（D-10）
    ap.add_argument("--version", default=appver.app_version())
    ap.add_argument("--iscc", default=None, help="ISCC.exe 路径（默认自动探测）")
    a = ap.parse_args()

    portable = ROOT / "dist" / "格式工坊"
    if not (portable / "app.py").is_file():
        raise SystemExit(f"缺少 {portable}——先跑 python tools/build_portable.py")

    # 分发包若被拿去做过冒烟测试，outputs 里会留着运行日志；安装包不带这些
    stale = portable / "outputs"
    if stale.is_dir():
        shutil.rmtree(stale)
    (portable / "outputs").mkdir()

    # 旧版 build_portable 可能已把根目录内部文档拷进包里（D-08）：进
    # Program Files 前清掉，白名单见 tools/package_docs.py
    strip_private_docs(portable)

    iscc = find_iscc(a.iscc)
    out = ROOT / "dist" / f"格式工坊-setup-{a.version}.exe"
    gen = iscc_ready_source()
    print(f"ISCC {iscc}")
    print(f"版本 {a.version} → {out.name}")
    r = subprocess.run(
        [iscc, f"/DMyAppVersion={a.version}", str(gen)],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    # ISCC 的进度/告警在 stdout，出错时原样带出来好排查
    if r.returncode != 0:
        print(r.stdout[-2000:])
        raise SystemExit(f"ISCC 失败（exit {r.returncode}）：{r.stderr.strip()[-500:]}")
    print(f"完成：{out}（{out.stat().st_size / 1024 / 1024:.0f} MB）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
