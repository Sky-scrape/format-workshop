"""本机作者档案：跨任务复用姓名、单位、邮箱等身份信息（默认关闭）。

DESIGN §10.7 的约束：
  - 默认仅会话级保存；跨会话复用需用户显式开启，且可一键清除；
  - 档案只存「与具体论文无关」的身份字段，不存基金项目、摘要、标题；
  - 不进入规则集、不进审计报告，明文只在本机配置目录。

隐私取舍：不保存身份证号（id_number）——它属于高敏感证件信息，补录页
照常可填，但不落盘到跨会话档案。档案文件与 ai.json 同级，卸载/清除即删。
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import app_paths
import atomic_io

ROOT = Path(__file__).resolve().parent.parent
# 测试与 ai_assist.CONFIG_PATH 同样处理：monkeypatch 这个模块级路径。
PROFILE_PATH = app_paths.data_root(ROOT) / "config" / "author_profile.json"

# 可跨论文复用的身份字段。键与补录页/作者页字段同名；基金、摘要、标题、
# 分类号等与论文绑定的字段不在列。身份证号也排除。
REUSABLE_KEYS = {
    "author_name", "english_authors", "english_affiliation", "affiliation_zh",
    "affiliation", "email", "author_email", "corresponding", "phone",
    "contact", "contact_address", "mailing_address", "postcode", "orcid",
    "author_birth", "author_gender", "author_title", "author_degree",
    "author_city", "author_hometown", "organization_and_position",
    "professional_title", "highest_degree", "research_direction",
    "author_bio",
}

MAX_VALUE = 500


def load():
    """读档案。坏文件/不存在按「未启用、无字段」兜底，绝不让配置坏了挡出稿。"""
    if not PROFILE_PATH.exists():
        return {"enabled": False, "fields": {}}
    try:
        raw = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {"enabled": False, "fields": {}}
    if not isinstance(raw, dict):
        return {"enabled": False, "fields": {}}
    fields = raw.get("fields")
    fields = {k: str(v) for k, v in fields.items()
              if k in REUSABLE_KEYS and isinstance(v, str) and v.strip()} \
        if isinstance(fields, dict) else {}
    return {"enabled": bool(raw.get("enabled")), "fields": fields}


def _write(data):
    atomic_io.write_text(PROFILE_PATH,
                         json.dumps(data, ensure_ascii=False, indent=2) + "\n")


def set_enabled(flag):
    data = load()
    data["enabled"] = bool(flag)
    _write(data)
    return data


def remember(fields):
    """把表单里可复用的身份字段并入档案（只增改非空值，不清空已有项）。"""
    data = load()
    changed = 0
    for key, value in (fields or {}).items():
        if key not in REUSABLE_KEYS or not isinstance(value, str):
            continue
        value = value.strip()
        if not value or len(value) > MAX_VALUE:
            continue
        if data["fields"].get(key) != value:
            data["fields"][key] = value
            changed += 1
    if changed:
        data["enabled"] = True
        _write(data)
    return changed


def prefill(existing=None):
    """给补录页的预填值：启用时用档案补空位，已有值优先（§10.3 不覆盖）。"""
    data = load()
    if not data["enabled"]:
        return {}, 0
    out = {}
    for key, value in data["fields"].items():
        if not (existing or {}).get(key):
            out[key] = value
    return out, len(out)


def clear():
    """一键清除：删除档案文件（含启用标记）。"""
    try:
        PROFILE_PATH.unlink()
        return True
    except FileNotFoundError:
        return False
