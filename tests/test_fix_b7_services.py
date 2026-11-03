"""批七「可见性与安全小项」七问的回归钉子（WEBLOG1 / WEBLOG3 / KNCSS /
F45UP / APPCTLHON / PDFRACE / COMPNOND）。

- WEBLOG1  web 建库链路不传 log：F52 截断降级重试、F51 警告在 Flask 路径
  全部静默。→ extract 调用传 app.logger.warning，事件进服务日志。
- WEBLOG3  ocr_low_text_pages 在草稿复核页无任何展示。→ 黄色 flash warn
  横幅，与识别失败的红档同区、分级别（低文字层≠失败，不替人下结论）。
- KNCSS    审计页折叠徽标 class="badge kN" 但 CSS 只定义 .kA/.kB。→ 补
  中性灰系 .kN。
- F45UP    scheme 白名单区分大小写："HTTPS://example.com" 被降级纯文本
  （浏览器按不区分大小写解析）。→ 判断前 url.lower()。
- APPCTLHON ps1 崩溃/超时与「无匹配」同样报「未找到匹配进程」——执行失败
  与确实无进程混为一谈。→ run_ps1_full 失败返回 -1，stop() 分开报。
  （后续补钉：-File 下 ps1 崩溃退出码实测是 1 而非负数、stdout 为空，与旧
  「无匹配 exit 1」仍同形——无匹配改走 exit 2，见 test_fix_b7_review_appctl.py。）
- PDFRACE  remove_old_pdf 的 exists() 与 unlink() 之间文件被外部删掉会以
  FileNotFoundError 裸 traceback 退出。→ 捕获后按已删除继续。
- COMPNOND blocks 元素为非 dict 时 _typed 盖了缺 type 键、盖不了元素本身，
  取值处仍冒 AttributeError。→ 非 dict 跳过并记 warning 点名。

夹具姿势：web 路由直接 import app（同 test_fix_web_robust / b6_services）；
工具走 run_tool 子进程（同 test_fix_f16noise_spec_audit_layers）。
"""

import io
import json
import logging
import pathlib
import subprocess
import types

import pytest

from conftest import RULES, run_tool

import app as app_module
import app_ctl
import build_journal
import compliance_report


# ================================================================ WEBLOG1

class _CaptureHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.messages = []

    def emit(self, record):
        self.messages.append(record.getMessage())


def test_web_extract_logs_degrade_retry(tmp_path, monkeypatch):
    """复现 WEBLOG1：web 建库链路必须把抽取过程事件记进日志——mock 截断
    场景（首轮输出截成半截 JSON），降级重试的提示要出现在服务日志里，
    不能像旧版那样整条链路静默。"""
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    monkeypatch.setattr(app_module, "DRAFT_DIR", tmp_path)
    monkeypatch.setattr(app_module.ai_assist, "available", lambda cfg=None: True)
    monkeypatch.setattr(app_module.ai_assist, "load_config", lambda: {})

    calls = []

    def fake_chat(cfg, messages, max_tokens=2048, **kw):
        calls.append(messages)
        if len(calls) == 1:
            # 首轮：截成半截的 JSON（_json_blob 找得到花括号，json.loads 失败）
            return '{"facts": {"body_size_pt":'
        # 降级重试：最简格式，只要取值
        return ('{"facts": {"body_size_pt": "10.5"}, '
                '"evidence": {"body_size_pt": "正文用小五号宋体"}, "notes": ""}')

    monkeypatch.setattr(app_module.ai_assist, "_chat", fake_chat)
    monkeypatch.setattr(
        app_module.extract_journal, "pdf_text",
        lambda path, max_chars=24000, log=None:
            ("稿约正文：正文用小五号宋体。",
             {"method": "text_layer", "truncated": False,
              "ocr_failed_pages": [], "ocr_low_text_pages": []}))

    cap = _CaptureHandler()
    app_module.app.logger.addHandler(cap)
    try:
        resp = client.post("/journals/extract", data={
            "jid": "b7weblog", "name": "演示刊",
            "file": (io.BytesIO(b"%PDF-1.4 mock"), "guide.pdf"),
        })
    finally:
        app_module.app.logger.removeHandler(cap)

    assert resp.status_code == 302, "抽取应成功并跳去复核页"
    assert len(calls) == 2, "截断后应真的降级重试了一次"
    assert any("降级重试" in m for m in cap.messages), cap.messages
    assert any("事实抽取" in m for m in cap.messages), cap.messages
    draft = json.loads((tmp_path / "b7weblog" / "draft.json")
                       .read_text(encoding="utf-8"))
    assert draft["verification"]["extraction"]["used"], "带引文的值应正常落盘"


# ================================================================ WEBLOG3

def test_draft_page_warns_low_text_pages(tmp_path, monkeypatch):
    """复现 WEBLOG3：带 ocr_low_text_pages 的草稿必须在复核页出黄色横幅，
    与识别失败的红档同区分级别。"""
    rules = {"journal": {"name_zh": "演示刊"},
             "verification": {"extraction": {
                 "source_file": "稿约.pdf", "text_method": "text_layer",
                 "unverified": [], "ocr_failed_pages": [],
                 "ocr_low_text_pages": [{"page": 2, "chars": 30},
                                        {"page": 5, "chars": 12}]}}}
    report = {"stats": {"extracted": 1, "unverified": 0, "inherited": 0,
                        "dropped": 0}, "errors": [], "warnings": []}
    d = tmp_path / "b7lowtx"
    d.mkdir()
    (d / "draft.json").write_text(json.dumps(rules, ensure_ascii=False),
                                  encoding="utf-8")
    (d / "report.json").write_text(json.dumps(report, ensure_ascii=False),
                                   encoding="utf-8")
    monkeypatch.setattr(app_module, "DRAFT_DIR", tmp_path)

    client = app_module.app.test_client()
    page = client.get("/journals/draft/b7lowtx").data.decode("utf-8")

    assert "文字层异常少" in page, "低文字层页在复核页完全没有提示"
    assert "第 2、5 页" in page, "横幅要点名具体页码"
    assert "flash warn" in page, "低文字层是提示不是失败，用黄档"
    assert "识别失败" not in page, "没发生失败就不该出现红档文案"


# ================================================================ KNCSS

def test_audit_kn_badge_has_css(tmp_path):
    """复现 KNCSS：N 档条目页面含 kN 徽标，且页面 CSS 里真的定义了 .kN。"""
    audit = {"target": "统计研究", "source": "manuscript.docx", "items": [],
             "assumptions": [f"推定字段{i} = 5（未从稿约抽到，沿用基线值，"
                             f"须人工核对）" for i in range(15)]}
    p = tmp_path / "audit.json"
    p.write_text(json.dumps(audit, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "audit.html"
    run_tool("render_audit.py", p, "-o", out)
    html = out.read_text(encoding="utf-8")

    assert '<span class="badge kN">' in html, "折叠分组用的是 N 档徽标"
    assert ".kN{" in html, "CSS 里没有 .kN——徽标可见但无配色"


# ================================================================ F45UP

def _rules_with_source_url(tmp_path, url, name="rules_f45.json"):
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["verification"]["source_url"] = url
    p = tmp_path / name
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return p


def test_spec_https_uppercase_renders_link(tmp_path):
    """复现 F45UP："HTTPS://" 大写 scheme 浏览器按 http(s) 解析，规格页
    必须照样渲染成链接，而不是降级纯文本。"""
    p = _rules_with_source_url(tmp_path, "HTTPS://journal.example.com/guide")
    out = tmp_path / "spec.html"
    run_tool("render_spec.py", p, "-o", out)
    html = out.read_text(encoding="utf-8")

    assert '<a href="HTTPS://journal.example.com/guide">' in html, \
        "大写 scheme 被错误降级成纯文本"
    # 防编造原闸不放松：非 http(s) 的 scheme 照旧降级（fail-closed 不变）
    p2 = _rules_with_source_url(tmp_path, "javascript:alert(1)",
                                name="rules_f45_js.json")
    out2 = tmp_path / "spec_js.html"
    run_tool("render_spec.py", p2, "-o", out2)
    assert '<a href="javascript:alert(1)">' not in out2.read_text(encoding="utf-8")


# ================================================================ APPCTLHON

def _quiet_stop(monkeypatch, ps1_result):
    monkeypatch.setattr(app_ctl, "run_ps1_full",
                        lambda name, *args: ps1_result)
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: False)
    monkeypatch.setattr(app_ctl, "time",
                        types.SimpleNamespace(sleep=lambda *_: None))


def test_stop_ps1_crash_reported_as_failure(monkeypatch, capsys):
    """复现 APPCTLHON：ps1 崩溃/超时（exit<0）必须报「执行失败」，
    不能与「无匹配」混为一谈。"""
    _quiet_stop(monkeypatch, ("", -1))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("停止脚本执行失败") == 2, "守护与服务两条都要如实报告"
    assert "未找到匹配进程" not in out, "执行失败不许说成没有进程"


def test_stop_no_match_still_reported(monkeypatch, capsys):
    """对照组：exit 2 且空输出（ps1 约定的无匹配）照旧报「未找到」——
    修 APPCTLHON 不能把 F48 的如实报告改回去。"""
    _quiet_stop(monkeypatch, ("", 2))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("未找到匹配进程") == 2


def test_run_ps1_full_exception_returns_negative(monkeypatch):
    """run_ps1_full 自身：powershell 缺失/超时等异常返回 -1，不再伪装成
    「无匹配」的 exit 1。"""
    def boom(*a, **kw):
        raise subprocess.TimeoutExpired(cmd="powershell", timeout=60)
    monkeypatch.setattr(app_ctl.subprocess, "run", boom)
    assert app_ctl.run_ps1_full("kill_by_cmdline.ps1", "-Pattern", "x",
                                "-Root", ".") == ("", -1)


# ================================================================ PDFRACE

def test_remove_old_pdf_survives_external_delete(tmp_path, monkeypatch):
    """复现 PDFRACE：exists() 与 unlink() 之间文件被外部删掉——unlink 抛
    FileNotFoundError 时按已删除继续，不再裸 traceback。"""
    pdf = tmp_path / "old.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    real_unlink = pathlib.Path.unlink

    def racing_unlink(self):
        real_unlink(self)                       # 外部先删掉了
        raise FileNotFoundError(2, "No such file or directory", str(self))

    monkeypatch.setattr(pathlib.Path, "unlink", racing_unlink)
    build_journal.remove_old_pdf(pdf)           # 不抛即通过
    assert not pdf.exists()


def test_remove_old_pdf_permission_error_still_fails_loudly(tmp_path,
                                                            capsys):
    """对照组：被占用（PermissionError）仍要点名文件并 SystemExit(1)——
    修竞态不能把这条出路一起吞掉。"""
    class BusyPDF:
        def __str__(self):
            return str(tmp_path / "busy.pdf")

        def unlink(self):
            raise PermissionError(13, "Permission denied")

    with pytest.raises(SystemExit) as ei:
        build_journal.remove_old_pdf(BusyPDF())
    assert ei.value.code == 1
    assert "被占用" in capsys.readouterr().out


# ================================================================ COMPNOND

def test_check_skips_nondict_block_with_warning():
    """复现 COMPNOND：blocks 元素为字符串（非 dict）时工具不得崩，
    warning 点名是哪份数据的第几块、判定照常给出。"""
    ir = {"blocks": ["这一块整个不是对象",
                     {"idx": 1, "type": "body", "text": "正文段落"}]}
    rules = {"journal": {"name_zh": "演示刊"},
             "citation": {"system": "numbered"},
             "blocks": [{"type": "body", "required": True}]}
    res = compliance_report.check(ir, rules)    # 修复前此处 AttributeError
    assert any("稿件 IR 第 1 块不是对象" in w and "str" in w
               for w in res["warnings"]), res["warnings"]
    assert res["rows"], "跳过坏块后其余判定照常给出"
    # 字数统计只数没被跳过的正文块
    assert res["words"][0][1] == "4", res["words"]


def test_cli_nondict_block_no_traceback(tmp_path):
    """端到端：blocks 首元素为字符串走 CLI 不崩（退出码 0），warning 在
    stdout 点名、报告里照常渲染。"""
    ir = tmp_path / "ir_nondict.json"
    rules = tmp_path / "rules_nondict.json"
    ir.write_text(json.dumps({"blocks": ["不是对象",
                                         {"type": "body", "text": "正文"}]},
                             ensure_ascii=False), encoding="utf-8")
    rules.write_text(json.dumps({"journal": {"name_zh": "演示刊"},
                                 "citation": {"system": "numbered"},
                                 "blocks": [{"type": "body",
                                             "required": True}]},
                                ensure_ascii=False), encoding="utf-8")
    out = run_tool("compliance_report.py", ir, rules, "-o", tmp_path / "r.html")
    assert "不是对象" in out, "CLI 必须把 warning 喊出来"
    assert (tmp_path / "r.html").exists()
    assert "不是对象" in (tmp_path / "r.html").read_text(encoding="utf-8")
