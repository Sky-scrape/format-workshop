"""第二轮低风险优化的回归钉：期刊列表缓存与安全响应头。"""

import json
import shutil
import time

import pytest

import app
from conftest import RULES


@pytest.fixture()
def client():
    app.app.config["TESTING"] = True
    return app.app.test_client()


def _one_journal_rules(tmp_path, monkeypatch):
    rules_dir = tmp_path / "journal"
    rules_dir.mkdir()
    target = rules_dir / "tjyj.json"
    shutil.copy(RULES / "tjyj.json", target)
    (rules_dir / "tjyj.spec.html").write_text("spec", encoding="utf-8")
    monkeypatch.setattr(app, "RULES_DIR", rules_dir)
    return target


def test_journal_cache_hits_until_rule_or_spec_changes(tmp_path, monkeypatch):
    target = _one_journal_rules(tmp_path, monkeypatch)
    real_load = app._load_journals_uncached
    calls = []

    def counted():
        calls.append(1)
        return real_load()

    monkeypatch.setattr(app, "_load_journals_uncached", counted)
    first = app.load_journals()
    second = app.load_journals()
    assert len(calls) == 1, "无变化的连续请求仍在重解析规则集"
    assert first == second and len(first) == 1
    assert first[0]["spec"] is True

    # 仅规格页变化也要让缓存失效：工作台链接是否可用取决于它。
    (tmp_path / "journal" / "tjyj.spec.html").unlink()
    assert app.load_journals()[0]["spec"] is False
    assert len(calls) == 2

    # 规则内容变化后名称必须刷新，不能继续展示旧刊名。
    data = json.loads(target.read_text(encoding="utf-8"))
    data["journal"]["name_zh"] = "缓存失效测试刊"
    app.write_json_atomic(target, data)
    stamp = time.time_ns() + 1_000_000
    import os
    os.utime(target, ns=(stamp, stamp))
    assert app.load_journals()[0]["name"] == "缓存失效测试刊"
    assert len(calls) == 3


def test_journal_cache_returns_independent_list(tmp_path, monkeypatch):
    _one_journal_rules(tmp_path, monkeypatch)
    first = app.load_journals()
    first.clear()
    assert len(app.load_journals()) == 1, "调用方清空返回值污染了共享缓存"


def test_security_headers_are_present(client):
    r = client.get("/")
    assert r.status_code == 200
    assert r.headers["X-Content-Type-Options"] == "nosniff"
    assert r.headers["X-Frame-Options"] == "SAMEORIGIN"
    assert r.headers["Referrer-Policy"] == "no-referrer"
    csp = r.headers["Content-Security-Policy"]
    assert "frame-ancestors 'self'" in csp
    assert "object-src 'none'" in csp
    assert r.headers["Cache-Control"] == "no-store", "论文相关页面不得被浏览器持久缓存"
