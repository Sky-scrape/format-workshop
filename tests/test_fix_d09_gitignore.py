"""D-09 仓库卫生的回归钉：门面 exe 与壳运行期状态不得挂进 git status。

两条线，各自钉一件事：

- **根目录门面 exe 已删除且永不再现**：根目录 格式工坊.exe 与
  build\\格式工坊.exe 曾是同字节的 8,346,631 字节复制（md5 一致），无
  ignore 规则时 git status 永远挂着 8MB 未跟踪噪音，一次 `git add .`
  就把二进制带进公开 AGPL 历史。门面属构建产物，只该存在于 build/
  （tools/build_launcher.py:37 产物落 build\\，tools/build_portable.py:250
  从那里复制进便携包）；快捷方式没有门面时本就回落 pythonw 直拉
  desktop.py（tools/app_ctl.py:175-183），删根目录副本不改变行为。
  正因为产物路径在 build/ 而不在根，断言「根目录不存在」是稳定的。
- **壳的运行期状态文件必须被 ignore**：在仓库目录跑过壳就会生成
  config/window.json（desktop.py:39，窗口尺寸记忆）、config/close.json
  （app.py:1235，用户在设置页写的关窗偏好）、config/tray-hint.flag
  （desktop.py:611，托盘气泡只提示一次的标记）。此前 .gitignore 只挡
  config/ai.json，这三个文件一旦生成，提交即入库。

用例以 `git check-ignore` 的退出码为准，而不是对 .gitignore 做子串
匹配——pattern 写错（多了空格、斜杠方向反了）时子串照样命中，真值
只有 git 自己知道。
"""

import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent

# D-09 新增的四条：门面 exe + 壳运行期状态。路径写成仓库相对形式，
# 与 .gitignore 里的条目一一对应
MUST_BE_IGNORED = [
    "格式工坊.exe",
    "config/window.json",
    "config/close.json",
    "config/tray-hint.flag",
    # 本机作者档案：姓名、单位、邮箱等身份信息，2026-11-13 起跨会话保存
    "config/author_profile.json",
]


@pytest.mark.parametrize("relpath", MUST_BE_IGNORED)
def test_gitignore_covers_shell_state_and_facade(relpath):
    """四条路径在 git 眼里都必须是 ignored（check-ignore 退出码 0）。"""
    r = subprocess.run(["git", "check-ignore", "--", relpath],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", cwd=str(ROOT))
    assert r.returncode == 0, \
        f"{relpath} 未被 .gitignore 覆盖（check-ignore 退出码 {r.returncode}）"


def test_root_facade_exe_is_gone():
    """根目录不再有门面 exe；门面只从 build/ 出货。"""
    assert not (ROOT / "格式工坊.exe").exists(), \
        "根目录 门面 exe 又出现了——它是 build/ 产物，不该躺在仓库根"
