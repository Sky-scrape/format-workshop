r"""发布链：构建产物 →（显式）上传 GitHub Releases。正式版只从 Release 下发。

发布纪律（README「发布」节同款）：
    1. 开发版改完 → 全量测试 → 提交 → 推送
    2. py -3 tools/release.py            # 构建 zip + setup 安装包
    3. py -3 tools/release.py --upload   # 构建并上传（tag v<版本> + GitHub Releases）
    4. 新版本 = 从 GitHub Releases 下载安装——开发者自己的机器也走这条路径

保证：
    - 构建锚定已提交状态（build_portable 的脏树闸门在这里提前拦，双保险）；
    - 上传前要求 HEAD 与上游一致——Release 的 tag 永远钉在远端已有的提交上，
      没推送就没有发布；
    - 不给 --upload 绝不外发，构建完只打印产物与上传命令。
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

import app_version as appver
from build_portable import working_tree_dirty

ROOT = Path(__file__).resolve().parent.parent

# 版本号只允许字母数字与点/连字符：它要进 tag（v<版本>）、文件名与子命令参数，
# 异常值（空格、shell 元字符）在这里就地拒绝，不让它流进任何进程参数
_VER_RE = re.compile(r"[0-9][0-9A-Za-z.\-]*")


def last_tag() -> str:
    """最近一个 tag（语义化倒推 git describe）；首次发布返回空串。"""
    r = subprocess.run(["git", "describe", "--tags", "--abbrev=0"], cwd=str(ROOT),
                       capture_output=True, text=True, timeout=30, shell=False)
    return r.stdout.strip() if r.returncode == 0 else ""


def release_notes(tag: str) -> str:
    """发布说明：上个 tag 以来的提交一行一条；首次发布退化为固定文案。"""
    if tag:
        r = subprocess.run(["git", "log", "--pretty=- %s", tag + "..HEAD"],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=30,
                           shell=False)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    return "首次发布。"


def upstream_synced() -> tuple:
    """Release 必须锚在远端已有提交上：HEAD 不等于 upstream（没推/落后）就不许传。"""
    r = subprocess.run(["git", "rev-parse", "HEAD"], cwd=str(ROOT),
                       capture_output=True, text=True, timeout=30, shell=False)
    if r.returncode != 0:
        return False, ""
    head = r.stdout.strip()
    r2 = subprocess.run(["git", "rev-parse", "@{upstream}"], cwd=str(ROOT),
                        capture_output=True, text=True, timeout=30, shell=False)
    if r2.returncode != 0:
        return False, head
    return head == r2.stdout.strip(), head


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--upload", action="store_true",
                    help="构建后上传 GitHub Release（tag v<版本>）；不给就绝不外发")
    a = ap.parse_args()

    if working_tree_dirty():
        raise SystemExit(
            "工作区有未提交改动，拒绝发布——正式版必须锚定已提交状态。\n"
            "先提交推送，再回来跑发布。")

    ver = appver.app_version()
    if not _VER_RE.fullmatch(ver):
        raise SystemExit("APP_VERSION 形态异常，拒绝发布：" + repr(ver))

    # 构建两条腿：绿色包 zip → setup 安装包（顺序有依赖，不能并行）。
    # 命令列表就地内联（ Mimosa 闸门要求：不把整条命令列表攒在变量里传）
    print("→ build_portable.py --zip")
    r = subprocess.run([sys.executable, "tools/build_portable.py", "--zip"],
                       cwd=str(ROOT), shell=False)
    if r.returncode != 0:
        raise SystemExit("构建步骤失败（build_portable --zip）")

    print("→ build_installer.py --version " + ver)
    r = subprocess.run([sys.executable, "tools/build_installer.py", "--version", ver],
                       cwd=str(ROOT), shell=False)
    if r.returncode != 0:
        raise SystemExit("构建步骤失败（build_installer）")

    zip_path = ROOT / "dist" / ("格式工坊-" + ver + ".zip")
    setup_path = ROOT / "dist" / ("格式工坊-setup-" + ver + ".exe")
    missing = [p.name for p in (zip_path, setup_path) if not p.is_file()]
    if missing:
        raise SystemExit("构建完成但产物缺失：" + repr(missing))

    print("")
    print("产物就绪（版本 " + ver + "）：")
    print("  " + str(zip_path) + "（" + str(zip_path.stat().st_size // 1048576) + " MB）")
    print("  " + str(setup_path) + "（" + str(setup_path.stat().st_size // 1048576) + " MB）")

    if not a.upload:
        print("")
        print("上传（显式动作，不会自动发）：py -3 tools/release.py --upload")
        return 0

    ok, head_sha = upstream_synced()
    if not ok:
        raise SystemExit(
            "本地 HEAD 与上游不一致（没推送？）——Release 的 tag 必须钉在远端已有"
            "的提交上，先 push 再 --upload。")
    tag = "v" + ver
    notes = release_notes(last_tag())
    print("→ 上传 " + tag + "（github.com/Sky-scrape/format-workshop）")
    r = subprocess.run(["gh", "release", "create", tag, str(zip_path), str(setup_path),
                        "--title", "格式工坊 " + ver, "--notes", notes,
                        "--target", head_sha],
                       cwd=str(ROOT), shell=False)
    if r.returncode != 0:
        raise SystemExit("gh release create 失败（exit " + str(r.returncode) + "）")
    print("已发布：https://github.com/Sky-scrape/format-workshop/releases/tag/" + tag)
    return 0


if __name__ == "__main__":
    sys.exit(main())
