"""路径与页面安全五条修复的回归钉（FX1 / F44 / F45 / F46 / F47）。

- FX1：docx_meta.sanitize_docx 的 os.replace 套 PermissionError 重试——
  Windows 瞬时文件占用（WinError 5）曾让走真实渲染的路由测试偶发红。
- F44：失败页技术细节脱敏——traceback/stderr 里的本机绝对路径（仓库根、
  用户主目录）不得进 error.json 与失败页。
- F45：规格页 source_url 只认 http(s)，其余降级纯文本（html.escape 不拦
  javascript: 协议）。
- F46：审计报告 pos_label 对 idx 先 int() 归一、非数字转义（三处插值
  共用此函数，修一处全覆盖）。
- F47：期刊 id 统一先过 JID_RE 再参与路径拼接/写 meta（..%5C 复现）。

夹具姿势照 tests/test_fix_route_guards.py：WORK 指临时目录；render_spec /
render_audit 走 run_tool 子进程，与真实命令行同一入口。
"""

import io
import json
import re
from pathlib import Path

import pytest
from docx import Document

from conftest import RULES, run_tool

import app
import docx_meta


@pytest.fixture()
def app_module():
    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


@pytest.fixture()
def work(app_module, tmp_path, monkeypatch):
    """把任务目录指到临时位置，测试不碰仓库里的 outputs/web。"""
    w = tmp_path / "work"
    w.mkdir(parents=True)
    monkeypatch.setattr(app_module, "WORK", w)
    return w


# ---------------------------------------------------------------- FX1

def test_sanitize_survives_transient_replace_denial(tmp_path, monkeypatch):
    """FX1：replace 首次抛 PermissionError（WinError 5）→ 重试后落位。

    修复前 sanitize_docx 的 os.replace 一撞 Windows 瞬时占用就直接失败，
    交付写全军覆没；修法与 app.py details.json 原子写同一姿势（5 次递增退避）。
    """
    path = tmp_path / "private.docx"
    doc = Document()
    doc.add_paragraph("正文段落，用于确认重试后内容完好。")
    doc.core_properties.author = "张三"
    doc.save(str(path))

    real_replace = docx_meta.os.replace
    calls = {"n": 0}

    def flaky_replace(src, dst):
        if str(dst) == str(path):
            calls["n"] += 1
            if calls["n"] == 1:
                raise PermissionError(5, "拒绝访问。")
        return real_replace(src, dst)

    monkeypatch.setattr(docx_meta.os, "replace", flaky_replace)
    removed = docx_meta.sanitize_docx(path)

    assert calls["n"] == 2, "首次 PermissionError 后必须重试，而不是放弃落位"
    assert "core:creator" in removed, "重试成功后清理项必须如实报告"
    assert docx_meta.scan_docx(path) == [], "重试后成品属性必须干净"
    assert Document(str(path)).paragraphs[0].text.startswith("正文段落"), \
        "清元数据不能动正文"
    assert not list(tmp_path.glob("private-*.docx")), "临时文件必须清掉"


# ---------------------------------------------------------------- F44

def test_fail_page_detail_has_no_local_paths(app_module, work, client,
                                             monkeypatch):
    """F44：带绝对路径的 detail → error.json 与失败页均无盘符路径。

    复现四处 detail 来源（traceback/format_exc/stderr）里最常见的形态：
    工具仓库根 + 带用户名的解释器目录。只抹路径，异常信息本身要保留。
    """
    d = work / "a1b2c3d4e5f6"
    d.mkdir()
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "s.docx", "stage": "details"})
    tool_frame = str(app_module.ROOT / "tools" / "render_from_ir.py")
    home_frame = str(Path.home() / "AppData" / "py") + "\\zipfile.py"
    detail = ("Traceback (most recent call last):\n"
              f'  File "{tool_frame}", line 1600, in main\n'
              f'  File "{home_frame}", line 9, in read\n'
              "BadZipFile: File is not a zip file")

    def boom(*a, **k):
        raise app_module.ConvertError("转换没有完成", "请重试一次", detail)

    monkeypatch.setattr(app_module, "stage_render", boom)
    r = client.post(f"/job/{d.name}/details", data={"english_abstract": "x"})
    assert r.status_code == 200, "渲染失败也要落到可读的失败页，不能裸 500"

    saved = json.loads((d / "error.json").read_text(encoding="utf-8"))
    page = r.get_data(as_text=True)
    root, home = app_module.ROOT, Path.home()
    for label, text in (("error.json", saved["detail"]), ("失败页", page)):
        assert str(root) not in text, f"{label} 泄露仓库根路径（F44）"
        assert root.as_posix() not in text, f"{label} 泄露仓库根路径（正斜杠写法）"
        assert str(home) not in text, f"{label} 泄露用户主目录（含用户名）"
        assert home.name not in text, f"{label} 泄露用户名"
    assert saved["title"] == "转换没有完成" and saved["hint"] == "请重试一次"
    assert "BadZipFile" in saved["detail"], "只抹路径，异常信息本身要保留"
    assert not re.search(r"[A-Za-z]:[\\/]", saved["detail"]), \
        "error.json 的 detail 里不应再有任何盘符路径"


# ---------------------------------------------------------------- F45

def _spec_html(tmp_path, rules, name="spec.json"):
    p = tmp_path / name
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / (name + ".html")
    run_tool("render_spec.py", p, "-o", out)
    return out.read_text(encoding="utf-8")


def test_spec_source_url_javascript_scheme_degrades_to_text(tmp_path):
    """F45：非 http(s) 的 source_url 不得渲染成链接，出处仍如实展示。"""
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["verification"]["source_url"] = "javascript:fetch('/job/x/download/docx')"
    html = _spec_html(tmp_path, rules)
    assert 'href="javascript' not in html, "javascript: 协议不得进 href（html.escape 拦不住）"
    assert "javascript:fetch(" in html, "降级为纯文本时出处本身仍要如实展示"


def test_spec_source_url_https_stays_link(tmp_path):
    """F45 正常值不受影响：现库的 http(s) 出处照常渲染成链接（href 内仍转义）。"""
    from html import escape
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    url = rules["verification"]["source_url"]
    assert url.startswith("https://"), "夹具前提：tjyj 的出处是 https"
    assert f'<a href="{escape(url)}">' in _spec_html(tmp_path, rules)


# ---------------------------------------------------------------- F46

def test_audit_pos_label_escapes_non_integer_idx(tmp_path):
    """F46：idx 为字符串值时不得原样插值进报告（int() 归一 + 转义兜底）。

    三处插值（改动表 / 版式分组 / 说明节）共用 pos_label，这里各钉一条：
    字符串 idx 转义，整数 idx 的正常显示不受影响。
    """
    audit = {
        "target": "统计研究",
        "source": "manuscript.docx",
        "items": [
            {"id": "t1", "action": "标签归一", "idx": "<script>alert(1)</script>",
             "before": "内容摘要：旧", "after": "摘　要：新", "note": "标签归一"},
            {"id": "t2", "action": "题注归一", "idx": 5,
             "before": "表1 收入分配", "after": "表1  收入分配", "note": "间隔归一"},
            {"id": "l1", "action": "正文套版", "idx": "<i>grp</i>",
             "before": "", "after": "", "note": "正文套版"},
        ],
    }
    p = tmp_path / "audit.json"
    p.write_text(json.dumps(audit, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "audit.html"
    run_tool("render_audit.py", p, "-o", out)
    html = out.read_text(encoding="utf-8")

    assert "<script>alert(1)" not in html, "改动表位置列不得原样插值"
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in html, "应转义为纯文本"
    assert "<i>grp</i>" not in html, "版式分组位置列不得原样插值"
    assert "&lt;i&gt;grp&lt;/i&gt;" in html
    assert '<td class="pos">5' in html, "整数 idx 的正常显示不得受影响"


# ---------------------------------------------------------------- F47

def test_traversal_journal_ids_are_rejected(app_module, client, work):
    """F47：..%5C 类期刊 id 不得参与路径拼接；正常 id 不受影响。"""
    # 证据复现：修复前 /journal/..%5Cjournal%5Ctjyj/spec → 200
    r = client.get("/journal/..%5Cjournal%5Ctjyj/spec")
    assert r.status_code == 404, "spec 页的期刊 id 必须先过 JID_RE"

    # workspace 预选参数同型守卫
    r = client.get("/workspace?journal=..%5Ctjyj")
    assert r.status_code == 400

    # diff 的 to 参数经 rules_of 拒收 → 404
    d = work / "a1b2c3d4e5f6"
    d.mkdir()
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj"})
    assert client.get(f"/job/{d.name}/diff?to=..%5Ccjyj").status_code == 404

    # 正常 id 不受影响
    assert client.get("/journal/tjyj/spec").status_code == 200
    assert client.get("/workspace?journal=tjyj").status_code == 200


def test_convert_and_retry_reject_traversal_journal(app_module, client, work):
    """F47：convert 的 keep 判断与 retry 的 jid 同样先过 JID_RE，
    retry 不再把原始非法 jid 写进 meta。"""
    r = client.post("/convert", data={
        "journal": "..\\tjyj", "file": (io.BytesIO(b"PK"), "稿.docx")})
    assert r.status_code == 400
    assert "请选择一本目标期刊" in r.get_data(as_text=True)
    assert not any(work.iterdir()), "非法期刊 id 不该创建任务目录"

    d = work / "a1b2c3d4e5f6"
    d.mkdir()
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "docx": "manuscript.tjyj.docx"})
    r = client.post(f"/job/{d.name}/retry", data={"journal": "..\\cjyj"})
    assert r.status_code == 400
    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta["journal"] == "tjyj", "非法 jid 不得写进 meta"
    # 正常 id 不受守卫影响：放行后走到后续的稿源检查（夹具无 manuscript.docx → 404），
    # 与非法 id 在守卫处就被 400 拦下形成区分
    assert client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"}) \
        .status_code == 404
