"""把项目打成可直接双击的绿色分发包（embeddable Python + 依赖 + 项目文件）。

为什么不用 PyInstaller
----------------------
现有进程模型有三层，且 app.py 用 sys.executable 起子进程跑 tools/*.py：

    bat → pythonw app_ctl.py → pythonw keepalive.py → pythonw app.py
                                                      └→ pythonw tools/*.py

打包成单个 exe 后，这三处全部失效：没有 pythonw.exe、tools/*.py 不是磁盘实体、
停服靠 PowerShell 按命令行里的 app.py 字样匹配也匹配不到。改造成本是一次进程模型
重构，还要动 test_fix_job_lifecycle.py 那 19 项。

embeddable 绿色包保留完整目录结构，上面每一条都照常工作——bat 里只多了一句
「优先用自带的 pythonw」。代价是「一个文件夹」而不是「一个文件」，收益是
零架构风险、规则库可直接替换更新、用户也能看见源码（AGPL 项目本就该如此）。

用法
----
    python tools/build_portable.py                # 完整构建
    python tools/build_portable.py --skip-deps    # 复用上次装好的依赖
    python tools/build_portable.py --zip          # 额外打一个 zip

产物
----
    dist/格式工坊/                        ← 双击 格式工坊.bat
    dist/格式工坊-<软件版本>.zip          ← --zip 时生成；包内附 BUILD-INFO.txt
"""

import argparse
import base64
import binascii
import hashlib
import ipaddress
import json
import re
import shutil
import socket
import subprocess
import sys
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

import app_version as appver
from package_docs import is_private_doc

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PY = "3.12.10"
OUT_NAME = "格式工坊"
# embeddable 包与摘要 sidecar 只认这两家官方源（护栏白名单，见 _safe_official_url）
_OFFICIAL_HOSTS = {"www.python.org", "python.org"}

# 不进分发包的顶层条目。三类：别人的稿子、期刊原件（版权）、含密钥的配置，
# 其余是构建产物与开发用文件。
EXCLUDE_TOP = {
    "rules/sources", "训练论文", "outputs", "旧脚本", "tests", "docs",
    ".git", ".github", ".workbuddy", ".pytest_cache", ".venv", "venv",
    "dist", "build", "shell", "__pycache__",
}
# 任意层级都不带的目录名
EXCLUDE_DIRS = {"__pycache__", ".git", ".pytest_cache", ".idea", ".vscode"}
# 单个文件级别排除：ai.json 里可能有 API key，开发依赖与 CI 配置也用不上；
# 开发版.bat 只在源码目录有意义（便携包没有 .venv，开发版判据也判不成）
EXCLUDE_FILES = {"config/ai.json", "config/window.json", "config/close.json",
                 "config/tray-hint.flag", "pytest.ini", "requirements-dev.txt",
                 ".gitignore", ".gitattributes", "开发版.bat"}
EXCLUDE_SUFFIX = {".pyc", ".pyo", ".log", ".tmp", ".bak", ".swp"}


def rel_posix(p: Path) -> str:
    return p.as_posix()


def is_excluded(rel: Path) -> bool:
    """一条路径是否不进分发包。

    两个坑都踩过：
      1. 嵌套路径要按前缀判。`"rules/sources"` 精确匹配排不掉
         `rules/sources/xxx.pdf`，53MB 的期刊原件会原样进包。
      2. 隐藏目录整体排除。`.git` / `.workbuddy` / `.mimosa` 这类不改名的
         工具产物没人想到要列，靠规则兜住，别靠清单追。
    """
    s = rel_posix(rel)
    parts = rel.parts
    if any(p.startswith(".") for p in parts):
        return True
    for e in EXCLUDE_TOP:
        if s == e or s.startswith(e + "/"):
            return True
    if parts and parts[0].startswith("bt_"):
        return True
    if is_private_doc(parts, rel.name):
        # 根目录散落的内部文档（实测报告等）不随 AGPL 包分发（D-08）
        return True
    if s in EXCLUDE_FILES:
        return True
    if any(part in EXCLUDE_DIRS for part in parts):
        return True
    return rel.suffix in EXCLUDE_SUFFIX


def _safe_official_url(url: str) -> str:
    """官方下载 URL 的护栏（URL 只许 http/https、host 白名单、拒绝环回/
    私网/链路本地/保留地址）。host 固定两家官方域名，ver/name 等动态段
    白名单校验——异常值进不了任何请求。

    例外：198.18.0.0/15 是 IANA 保留给基准测试的段，永不承载真实内网服务；
    Fake-IP 模式的本机代理（实测环境）把所有域名都解析到该段再由代理转发，
    拒绝它会让代理环境下的构建连官方源都下不了。其余非公网地址照拒。
    """
    parts = urllib.parse.urlsplit(url)
    if parts.scheme not in ("https", "http") or not parts.hostname:
        raise SystemExit(f"URL 协议/host 非法，拒绝请求：{url!r}")
    host = parts.hostname.lower()
    if host not in _OFFICIAL_HOSTS:
        raise SystemExit(f"host 不在官方域名白名单，拒绝请求：{host!r}")
    try:
        for info in socket.getaddrinfo(parts.hostname, None):
            ip = ipaddress.ip_address(info[4][0])
            if ip == ipaddress.ip_address("198.18.1.145"):
                continue
            if ip in ipaddress.ip_network("198.18.0.0/15"):
                continue                      # Fake-IP 代理的标记段，见 docstring
            if (ip.is_loopback or ip.is_private or ip.is_link_local
                    or ip.is_reserved or ip.is_multicast or ip.is_unspecified):
                raise SystemExit(f"官方域名解析到非公网地址，拒绝请求：{ip}")
    except socket.gaierror as e:
        raise SystemExit(f"官方域名解析失败，拒绝请求：{e}")
    return url


def _fetch_official(url: str, timeout: int = 60):
    """带护栏的官方源 GET：返回已校验的响应对象（调用方负责 close/read）。"""
    return urllib.request.urlopen(_safe_official_url(url), timeout=timeout)


def _official_digest(ver: str, name: str) -> str | None:
    """取官方公布的 zip 摘要（hex 小写字符串），拿不到返回 None。

    两个来源（都实测闭环：3.12.10/3.11.9 的 sha256、3.9.13 的 md5 均与
    zip 全量实算一致）：
      1. 同目录 .sigstore（3.10.11 起才有）：bundle 的 messageDigest 是
         SHA2_256，base64 → hex；
      2. 发布页 checksum（更老版本只有这个）：每文件 32 hex 的 md5，
         拆成两个 16 hex 半串展示，拼接即 md5。
    摘要拿不到（网络抖动、未来页面改版）不拦构建——退回体积兜底，只是
    失去完整性校验，与旧行为同。护栏拒绝（非法协议/劫持到内网）则照抛。
    """
    try:
        with _fetch_official(
                f"https://www.python.org/ftp/python/{ver}/{name}.sigstore") as r:
            b = json.loads(r.read())
        digest = b["messageSignature"]["messageDigest"]["digest"]
        return binascii.hexlify(base64.b64decode(digest)).decode("ascii")
    except Exception:                        # noqa: BLE001
        pass
    try:
        with _fetch_official(
                f"https://www.python.org/downloads/release/python-{ver.replace('.', '')}/") as r:
            page = r.read().decode("utf-8", "replace")
        row = next(m.group(0) for m in re.finditer(r"<tr>.*?</tr>", page, re.S)
                   if f"{name}\"" in m.group(0))
        halves = re.findall(r"checksum-half\">([0-9a-f]+)<", row)
        if halves and all(len(h) == 16 for h in halves):
            return "".join(halves)           # 32 hex 的 md5
    except Exception:                        # noqa: BLE001
        pass
    return None


def _hash_file(path: Path, algo: str) -> str:
    h = hashlib.new(algo)
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def fetch_python(cache: Path, ver: str) -> Path:
    """下载官方 embeddable 包并缓存。带超时；下载物按官方摘要校验。

    曾经的两处裸奔：urlretrieve 无 timeout，网络停滞时构建无限挂起；
    缓存有效性只看体积 >5MB，损坏或被篡改的 zip 解压时才暴露。现在下载
    走 urlopen(timeout)（经 _safe_official_url 护栏），摘要按官方
    .sigstore（sha256）或发布页（老版本 md5）核对——算不出官方摘要时退
    回体积兜底，不把构建锁死。
    """
    if not re.fullmatch(r"3\.\d+\.\d+", ver):
        raise SystemExit(f"--python 版本号形态异常，拒绝下载：{ver!r}")
    name = f"python-{ver}-embed-amd64.zip"
    z = cache / name
    # 缓存也要验：入缓存时校验过一次，但磁盘不会永远忠诚（坏块、被工具
    # 动过、上一轮构建中断残缺）。摘要重新拉一次官方值很便宜（几 KB），
    # 换来坏缓存当场识破而不是解压时炸——比发现里「只看体积」的口径强。
    if z.is_file() and z.stat().st_size > 5_000_000:
        expected = _official_digest(ver, name)
        if expected:
            algo = "sha256" if len(expected) == 64 else "md5"
            actual = _hash_file(z, algo)
            if actual != expected:
                z.unlink(missing_ok=True)
                print(f"  缓存 {algo} 不符（实际 {actual[:16]}…），丢弃重下")
            else:
                print(f"  复用缓存 {z.name}（{algo} 校验通过）")
                return z
        else:
            # 官方摘要拿不到（离线/页面改版）：保留旧口径按体积兜底
            print(f"  复用缓存 {z.name}（官方摘要拿不到，仅按体积兜底）")
            return z
    url = f"https://www.python.org/ftp/python/{ver}/{name}"
    cache.mkdir(parents=True, exist_ok=True)
    print(f"  下载 {url}")
    tmp = z.with_suffix(".part")
    with _fetch_official(url) as r, tmp.open("wb") as f:
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
    expected = _official_digest(ver, name)
    if expected:
        algo = "sha256" if len(expected) == 64 else "md5"
        actual = _hash_file(tmp, algo)
        if actual != expected:
            tmp.unlink(missing_ok=True)
            raise SystemExit(
                f"{name} 的 {algo} 不符：实际 {actual}，官方 {expected}——"
                "下载已丢弃，不进构建。")
        print(f"  {algo} 校验通过 {actual[:16]}…")
    else:
        print("  （官方摘要拿不到，仅按体积兜底——完整性未校验）")
    tmp.replace(z)
    return z


def installed_pyver(py_dir: Path) -> str | None:
    """读包内解释器的实际版本（python3XX._pth / python3XX.dll 的 XXX）。

    embeddable 包的目录名带版本：python312._pth、python312.dll。--skip-deps
    复用上次构建的 python/ 时，目录里躺着的是哪个解释器以这里为准——
    BUILD-INFO 写的是 --python 参数，两者不一致时记的就是假版本（实测：
    包内 python311 残留、参数 3.12.10，BUILD-INFO 照记 3.12.10）。
    """
    for pth in sorted(py_dir.glob("python3*._pth")):
        digits = pth.stem[len("python"):]
        if digits.isdigit():
            return f"{digits[0]}.{digits[1:]}"
    return None


def configure_pth(py_dir: Path, ver: str) -> None:
    """改 _pth：默认清单里没有 site-packages，也不 import site。

    embeddable 包设计成「完全隔离」，不改这两行的话装进去的依赖一个都 import
    不到——报错会是找不到 flask，看着像没装，实际是路径没进 sys.path。
    """
    major, minor = ver.split(".")[:2]
    pth = py_dir / f"python{major}{minor}._pth"
    if not pth.is_file():
        raise SystemExit(f"找不到 {pth.name}，embeddable 包结构可能变了")
    pth.write_text(
        f"python{major}{minor}.zip\n.\nLib\\site-packages\nimport site\n",
        encoding="ascii")


# 只有 sdist、没有轮子的纯 Python 依赖。install_deps 用 --only-binary=:all:
# 跨版本装依赖（防 cp313 轮子混进 cp312 包），会把这些误杀——先在本机把
# sdist 转成 none-any 通用轮子放进 build/wheels，再用 --find-links 取用：
# 通用轮子对任何目标环境都合法。
SDIST_ONLY = ("proxy-tools",)


def ensure_pure_wheels() -> None:
    out = ROOT / "build" / "wheels"
    out.mkdir(parents=True, exist_ok=True)
    for pkg in SDIST_ONLY:
        print(f"  本机补轮子 {pkg} → build/wheels")
        subprocess.run([sys.executable, "-m", "pip", "wheel", pkg,
                        "--no-deps", "-w", str(out)], check=True)


def install_deps(py_dir: Path, ver: str) -> None:
    """用本机 pip 把依赖装进目标目录。

    必须显式给 --platform/--python-version/--abi：本机解释器版本可能和目标包
    不同（本机 3.13、包里 3.12），不指定就会装成 cp313 的 wheel，进包后
    import 直接失败——而且是运行时才炸，构建时看不出来。
    """
    major, minor = ver.split(".")[:2]
    dest = py_dir / "Lib" / "site-packages"
    dest.mkdir(parents=True, exist_ok=True)
    cmd = [sys.executable, "-m", "pip", "install", "--quiet", "--upgrade",
           "--target", str(dest),
           "--platform", "win_amd64",
           "--python-version", f"{major}.{minor}",
           "--implementation", "cp",
           "--abi", f"cp{major}{minor}",
           "--only-binary=:all:",
           "--find-links", str(ROOT / "build" / "wheels"),
           "-r", str(ROOT / "requirements.txt")]
    print("  安装依赖：", " ".join(cmd[3:6]), f"→ {dest}")
    subprocess.run(cmd, check=True)


def strip_bytecode(site_packages: Path) -> int:
    """清掉 site-packages 里的 __pycache__ 与 .pyc。

    pip 用**本机**解释器装包时会顺手编译字节码，产出的 cpython-313.pyc 进了
    cp312 的目标目录：目标解释器版本对不上，根本不会加载，纯占体积；而文件名
    里那个 313 还会让人以为装错版本。标准库的 python312.zip 自带 .pyc，不在此列。
    """
    removed = 0
    for cache in site_packages.rglob("__pycache__"):
        if cache.is_dir():
            shutil.rmtree(cache, ignore_errors=True)
            removed += 1
    for pyc in site_packages.rglob("*.py[co]"):
        pyc.unlink(missing_ok=True)
    # pip 生成的 bin/ 是给类 Unix 系统用的启动脚本，Windows 下没有任何入口引用它
    bin_dir = site_packages / "bin"
    if bin_dir.is_dir():
        shutil.rmtree(bin_dir, ignore_errors=True)
    return removed


def copy_project(out: Path) -> int:
    """按白名单思路复制项目文件（排除表见 EXCLUDE_*）。"""
    n = 0
    for src in sorted(ROOT.rglob("*")):
        if not src.is_file():
            continue
        rel = src.relative_to(ROOT)
        if is_excluded(rel):
            continue
        dst = out / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        n += 1
    return n


def dir_size(p: Path) -> int:
    return sum(f.stat().st_size for f in p.rglob("*") if f.is_file())


def write_build_info(out: Path, ver: str, py_ver: str) -> None:
    """包根目录写 BUILD-INFO.txt：这包是谁、从哪份源码、什么时候打的。

    zip 名只放得下软件版本；构建来源（git rev）与解释器版本放包内，用户
    报问题时能对上号（D-10）。非仓库构建（无 git）如实标注。
    """
    try:
        r = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=str(ROOT),
                           capture_output=True, text=True, timeout=15)
        rev = r.stdout.strip() if r.returncode == 0 else ""
    except Exception:                        # noqa: BLE001
        rev = ""
    lines = [
        f"软件版本 {ver}",
        f"Python {py_ver}（embeddable）",
        f"git {rev or '（非仓库构建，无版本记录）'}",
        f"构建时间 {datetime.now():%Y-%m-%d %H:%M}",
    ]
    (out / "BUILD-INFO.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")


def working_tree_dirty() -> bool:
    """发布纪律的机器闸门：工作区不干净就不许打包。

    便携包是日常版的唯一来源，BUILD-INFO 里的 git rev 是它唯一的出处锚点
    ——树上有未提交改动时打包，包内容就领先/偏离任何提交，锚点失真（发生过：
    包内容对应 da78114，BUILD-INFO 却记 6261d5d）。顺序只能是：改完 → 提交 →
    再打包。--allow-dirty 是显式逃生口，每次都要亲手打出来才会走。
    """
    try:
        r = subprocess.run(["git", "status", "--porcelain"], cwd=str(ROOT),
                           capture_output=True, text=True, timeout=30)
    except Exception:                    # noqa: BLE001
        return False                     # 非 git 环境（罕见）不拦，维持旧行为
    return r.returncode == 0 and bool(r.stdout.strip())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--python", default=DEFAULT_PY, help=f"embeddable 版本，默认 {DEFAULT_PY}")
    ap.add_argument("--out", default=str(ROOT / "dist"))
    ap.add_argument("--skip-deps", action="store_true", help="跳过 pip（复用已有依赖）")
    ap.add_argument("--zip", action="store_true", help="构建后打包成 zip")
    ap.add_argument("--allow-dirty", action="store_true",
                    help="工作区不干净时仍打包（不推荐：BUILD-INFO 的 git rev 会失真）")
    a = ap.parse_args()

    # 脏树闸门放在一切破坏性操作（rmtree 旧包）之前：发布必须锚到一个提交
    if working_tree_dirty() and not a.allow_dirty:
        raise SystemExit(
            "工作区有未提交改动，拒绝打包——BUILD-INFO 的 git rev 会失真。\n"
            "先提交再构建；确要带脏树打包，用 --allow-dirty 显式放行。")

    # zip 名用软件版本（app.py 的 APP_VERSION），不是 --python——3.12.10 是
    # 解释器版本，放在文件名里没人知道对应软件的哪一版（D-10）
    ver = appver.app_version()
    py_ver = a.python
    out = Path(a.out) / OUT_NAME
    py_dir = out / "python"

    print(f"构建 {OUT_NAME} {ver}（Python {py_ver}）→ {out}")

    # 全量清理再重建：copy_project 只覆盖不删除，上一版留在 rules/ 里的
    # 文件（比如误复制的期刊原件）会一直跟着走。--skip-deps 时保留 python/，
    # 那是唯一值得复用的部分（重装一趟要几分钟）。
    if out.exists():
        if a.skip_deps and py_dir.exists():
            for child in out.iterdir():
                if child.name == "python":
                    continue
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink()
        else:
            shutil.rmtree(out)
    out.mkdir(parents=True, exist_ok=True)

    if not py_dir.exists():
        z = fetch_python(ROOT / "build" / "cache", py_ver)
        print("  解压解释器")
        with zipfile.ZipFile(z) as zf:
            zf.extractall(py_dir)
        configure_pth(py_dir, py_ver)
    else:
        # 复用上次构建的 python/（--skip-deps 路径）：先核对包内解释器
        # 与 --python 一致。不核对会静默产出不符包——依赖按旧解释器装好、
        # 内部自洽能跑，但版本与参数不符，BUILD-INFO 记假版本（发布溯源
        # 锚点失真，与脏树闸门拦的是同一类事故）。不一致即拒绝，删掉
        # dist/<包名>/python（或整个 out）重跑完整构建即可换版本。
        actual = installed_pyver(py_dir)
        if actual is not None and actual != py_ver:
            raise SystemExit(
                f"python/ 里是 Python {actual}，与 --python {py_ver} 不一致，"
                f"拒绝复用——换版本请删掉 {out} 后完整构建（去掉 --skip-deps）。")

    if not a.skip_deps:
        ensure_pure_wheels()
        install_deps(py_dir, py_ver)
    # 每次构建都清：--skip-deps 复用旧目录时，上一轮留下的缓存也在里面
    site_packages = py_dir / "Lib" / "site-packages"
    if site_packages.is_dir():
        n_cache = strip_bytecode(site_packages)
        if n_cache:
            print(f"  清理本机解释器编译的字节码缓存 {n_cache} 个目录")

    n = copy_project(out)
    print(f"  复制项目文件 {n} 个")

    write_build_info(out, ver, py_ver)

    # 软件门面：tools/build_launcher.py 的产物进包根目录——用户双击的是
    # 带图标的 exe 而不是 bat，体感即软件；没打门面就提示一句
    facade = ROOT / "build" / "格式工坊.exe"
    if facade.is_file():
        shutil.copy2(facade, out / "格式工坊.exe")
        print("  已放入门面 格式工坊.exe")
    else:
        print("  提示：未找到 build\\格式工坊.exe（跑一次 tools/build_launcher.py），"
              "包内将没有 exe 门面")

    # 产物目录：清空重建——冒烟测试留下的运行日志不该进分发包（--skip-deps
    # 复用目录时它不在清除范围内，尤其会带）。程序会自建，先建好第一屏更干净
    outputs = out / "outputs"
    if outputs.exists():
        shutil.rmtree(outputs)
    outputs.mkdir()

    size = dir_size(out)
    print(f"\n完成：{out}")
    print(f"  体积 {size / 1024 / 1024:.0f} MB")
    print(f"  入口 {out / '格式工坊.bat'}")

    if a.zip:
        zp = Path(a.out) / f"{OUT_NAME}-{ver}.zip"   # ver=软件版本（D-10）
        print(f"\n打包 {zp}")
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
            for f in sorted(out.rglob("*")):
                if f.is_file():
                    zf.write(f, Path(OUT_NAME) / f.relative_to(out))
        print(f"  zip {zp.stat().st_size / 1024 / 1024:.0f} MB")


if __name__ == "__main__":
    main()
