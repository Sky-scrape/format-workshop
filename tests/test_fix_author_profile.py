"""本机作者档案：默认关闭、白名单字段、一键清除（DESIGN §10.7）。

隐私边界：
  - 不勾「记住这些信息」不落盘；启用是显式动作；
  - 只存与具体论文无关的身份字段（摘要、标题、基金、身份证号都不存）；
  - 预填只补空位，不覆盖本次已填值；清除后文件消失。
"""

import sys
from pathlib import Path

import pytest

from conftest import RULES  # noqa: F401  （确保 tools 在 sys.path）

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

import author_profile as ap  # noqa: E402


@pytest.fixture()
def profile(tmp_path, monkeypatch):
    monkeypatch.setattr(ap, "PROFILE_PATH", tmp_path / "author_profile.json")
    return tmp_path / "author_profile.json"


def test_remember_only_whitelisted_identity_fields(profile):
    changed = ap.remember({
        "author_name": "张三", "email": "z@example.com", "corresponding": "张三（z@x）",
        "english_abstract": "This paper studies...", "funding": "国家自然科学基金",
        "id_number": "110101199001011234", "english_title": "Digital Finance",
    })
    data = ap.load()
    assert changed == 3
    assert data["enabled"] is True
    assert data["fields"] == {"author_name": "张三", "email": "z@example.com",
                              "corresponding": "张三（z@x）"}
    # 论文专属与高敏感字段一个都不许进档案
    for key in ("english_abstract", "funding", "id_number", "english_title"):
        assert key not in data["fields"]


def test_prefill_fills_gaps_without_overwriting(profile):
    ap.remember({"author_name": "张三", "email": "z@example.com"})
    vals, n = ap.prefill({"author_name": "李四"})
    assert vals == {"email": "z@example.com"} and n == 1
    assert vals.get("author_name") is None, "预填覆盖了本次已填值"


def test_disable_stops_prefill_but_keeps_fields(profile):
    ap.remember({"author_name": "张三"})
    ap.set_enabled(False)
    vals, _n = ap.prefill({})
    assert vals == {}
    assert ap.load()["fields"]["author_name"] == "张三", "暂停预填不该删数据"


def test_clear_removes_file(profile):
    ap.remember({"author_name": "张三"})
    assert profile.exists()
    assert ap.clear() is True
    assert not profile.exists()
    assert ap.load() == {"enabled": False, "fields": {}}


def test_broken_file_is_ignored(profile):
    profile.write_text("{not json", encoding="utf-8")
    assert ap.load() == {"enabled": False, "fields": {}}
    # 坏文件不影响再次保存
    assert ap.remember({"author_name": "张三"}) == 1


def test_settings_route_can_enable_and_clear(acceptance_client, profile):
    client = acceptance_client
    page = client.get("/settings").get_data(as_text=True)
    assert "本机作者档案" in page

    r = client.post("/settings/profile", data={"action": "enable"})
    assert r.status_code == 302
    assert ap.load()["enabled"] is True
    page = client.get("/settings").get_data(as_text=True)
    assert "已启用" in page

    ap.remember({"author_name": "张三"})
    r = client.post("/settings/profile", data={"action": "clear"})
    assert r.status_code == 302
    assert not profile.exists()

    assert client.post("/settings/profile", data={"action": "nope"}).status_code == 400
