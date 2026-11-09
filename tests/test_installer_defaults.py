"""安装器默认装到用户目录。

程序把任务、日志与配置写在安装目录；默认装进 Program Files 会让普通用户
双击启动时写不进去（历史 bug）。默认必须是 ``{localappdata}`` + 非提权。
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ISS = ROOT / "tools" / "installer" / "installer.iss"


def _src():
    return ISS.read_text(encoding="utf-8")


def test_default_dir_is_user_writable():
    """默认非提权安装：{autopf} 落 %LocalAppData%\\Programs，登录用户可写。"""
    src = _src()
    assert "DefaultDirName={autopf}\\格式工坊" in src
    assert "PrivilegesRequired=lowest" in src


def test_default_install_does_not_require_admin():
    src = _src()
    assert "PrivilegesRequired=lowest" in src
    assert "PrivilegesRequiredOverridesAllowed=commandline dialog" not in src


def test_all_users_is_still_reachable_by_command_line():
    """需要给所有用户装时仍可 /ALLUSERS；数据目录回退兜住只读安装。"""
    assert "PrivilegesRequiredOverridesAllowed=commandline" in _src()
