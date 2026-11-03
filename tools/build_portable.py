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
    dist/格式工坊/            ← 双击 格式工坊.bat
    dist/格式工坊.zip         ← --zip 时生成
"""

import argparse
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PY = "3.12.10"
OUT_NAME = "格式工坊"

# 不进分发包的顶层条目。三类：别人的稿子、期刊原件（版权）、含密钥的配置，
# 其余是构建产物与开发用文件。
EXCLUDE_TOP = {
    "rules/sources", "训练论文", "outputs", "旧脚本", "tests", "docs",
    ".git", ".github", ".workbuddy", ".pytest_cache", ".venv", "venv",
    "dist", "build", "shell", "__pycache__",
}
# 任意层级都不带的目录名
EXCLUDE_DIRS = {"__pycache__", ".git", ".pytest_cache", ".idea", ".vscode"}
# 单个文件级别排除：ai.json 里可能有 API key，开发依赖与 CI 配置也用不上
EXCLUDE_FILES = {"config/ai.json", "pytest.ini", "requirements-dev.txt",
                 ".gitignore", ".gitattributes"}
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
    if s in EXCLUDE_FILES:
        return True
    if any(part in EXCLUDE_DIRS for part in parts):
        return True
    return rel.suffix in EXCLUDE_SUFFIX


def fetch_python(cache: Path, ver: str) -> Path:
    """下载官方 embeddable 包并缓存。已存在且体积合理就跳过。"""
    name = f"python-{ver}-embed-amd64.zip"
    z = cache / name
    if z.is_file() and z.stat().st_size > 5_000_000:
        print(f"  复用缓存 {z.name}")
        return z
    url = f"https://www.python.org/ftp/python/{ver}/{name}"
    cache.mkdir(parents=True, exist_ok=True)
    print(f"  下载 {url}")
    tmp = z.with_suffix(".part")
    urllib.request.urlretrieve(url, tmp)
    tmp.replace(z)
    return z


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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--python", default=DEFAULT_PY, help=f"embeddable 版本，默认 {DEFAULT_PY}")
    ap.add_argument("--out", default=str(ROOT / "dist"))
    ap.add_argument("--skip-deps", action="store_true", help="跳过 pip（复用已有依赖）")
    ap.add_argument("--zip", action="store_true", help="构建后打包成 zip")
    a = ap.parse_args()

    ver = a.python
    out = Path(a.out) / OUT_NAME
    py_dir = out / "python"

    print(f"构建 {OUT_NAME}（Python {ver}）→ {out}")

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
        z = fetch_python(ROOT / "build" / "cache", ver)
        print("  解压解释器")
        with zipfile.ZipFile(z) as zf:
            zf.extractall(py_dir)
        configure_pth(py_dir, ver)

    if not a.skip_deps:
        install_deps(py_dir, ver)
    # 每次构建都清：--skip-deps 复用旧目录时，上一轮留下的缓存也在里面
    site_packages = py_dir / "Lib" / "site-packages"
    if site_packages.is_dir():
        n_cache = strip_bytecode(site_packages)
        if n_cache:
            print(f"  清理本机解释器编译的字节码缓存 {n_cache} 个目录")

    n = copy_project(out)
    print(f"  复制项目文件 {n} 个")

    # 产物目录与运行日志目录：程序会自建，但先建好能让双击后的第一屏更干净
    (out / "outputs").mkdir(exist_ok=True)

    size = dir_size(out)
    print(f"\n完成：{out}")
    print(f"  体积 {size / 1024 / 1024:.0f} MB")
    print(f"  入口 {out / '格式工坊.bat'}")

    if a.zip:
        zp = Path(a.out) / f"{OUT_NAME}-{ver}.zip"
        print(f"\n打包 {zp}")
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
            for f in sorted(out.rglob("*")):
                if f.is_file():
                    zf.write(f, Path(OUT_NAME) / f.relative_to(out))
        print(f"  zip {zp.stat().st_size / 1024 / 1024:.0f} MB")


if __name__ == "__main__":
    main()
