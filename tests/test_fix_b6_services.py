"""第六批「服务与工具」十个已确认问题的回归钉子。

覆盖（id 与问题单一一对应）：
- F48  stop() 死分支 + ps1 按子串误杀别项目进程 → -Root 双条件 + 如实报「未找到」
- F49  ENUM_HINT 把字数口径 B 写成「中文汉字数」→ 改为 B=中文字符数+英文单词数
- F50  候选超 max_blocks 说「AI 跳过」实为截断 → 日志如实 + trace 记 truncated/dropped
- F51  20 字阈值旁路（页眉凑数的扫描页静默跳过 OCR）→ 低文字层页进 meta + CLI 警告
- F52  _chat 硬编码 2048 且事实抽取无重试 → max_tokens 形参 + 截断降级重试走 unverified
- F57  历史页无单条删除、遮罩假进度 → /job/<id>/delete + 确认按钮 + 静态遮罩
- COMPSUB  IR/规则集块缺 type 键冒 KeyError → 按未知类型处理 + warnings 点名
- F16NOISE 全量继承后复核页噪音 → 折叠分组 + 计数徽标，落盘条数不减少
- SCRUBEXE 脱敏漏解释器目录 → 补 sys.prefix / sys.executable 目录
- BKB024   退避注释与实现不符 → 注释改「4 次退避共约 0.24s，最后一次直接抛出」

纪律：WORK 指到临时目录（不碰仓库 outputs/web）；AI 一律 monkeypatch，不联网；
F48 的 dry-run 用例会真实拉起一个带标记的 python 子进程并在 finally 里杀掉。
"""

import json
import shutil
import subprocess
import sys
import time
import types
from pathlib import Path

import pytest

from conftest import ROOT, run_tool

import app

TOOLS = ROOT / "tools"

sys.path.insert(0, str(TOOLS))

import ai_assist  # noqa: E402
import compliance_report  # noqa: E402
import extract_journal as ej  # noqa: E402

fitz = pytest.importorskip("fitz")
import ocr  # noqa: E402


# ================================================================ F48

PS1 = TOOLS / "kill_by_cmdline.ps1"
MARKER = "b6f48marker"


def run_ps1(*args):
    return subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(PS1), *args],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120)


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_ps1_no_match_exits_1_silently():
    """ps1 约定：无匹配 exit 1 零输出——stop() 靠它如实报「未找到」。"""
    r = run_ps1("-Pattern", "zzz_no_such_process_b6", "-Root", str(ROOT))
    assert r.returncode == 1, f"无匹配应 exit 1，实际 {r.returncode}"
    assert not r.stdout.strip(), "无匹配不该有任何输出（旧版把空输出当成功）"


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_ps1_scoped_to_project_root():
    """复现 F48 误杀：过滤必须是「命令行含模式 **且** 含本项目路径」双条件。

    拉起一个命令行带标记 + app.py 字样 + 本仓库路径的 python 子进程：
    -Root 指向本仓库 → 恰好命中它（且是 DryRun，不杀）；
    -Root 指向别的目录 → 即使模式命中也一个不杀（别项目的 app.py 不再被扫射）。
    """
    child = subprocess.Popen(
        [sys.executable, "-c", "import time; time.sleep(60)",
         MARKER, "app.py", str(ROOT)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        # 新进程进 CIM 有毫秒级延迟，轮询几次避免 flake
        scoped = None
        for _ in range(10):
            scoped = run_ps1("-Pattern", MARKER, "-Root", str(ROOT), "-DryRun")
            if scoped.returncode == 0 and scoped.stdout.strip():
                break
            time.sleep(0.8)
        assert scoped is not None and scoped.returncode == 0, scoped
        pids = {ln.split()[1] for ln in scoped.stdout.splitlines()
                if ln.startswith("match ")}
        assert str(child.pid) in pids, \
            f"本项目进程应命中：child={child.pid}，实际 {pids}"
        assert child.poll() is None, "DryRun 不得杀进程"

        foreign = run_ps1("-Pattern", MARKER, "-Root", r"D:\b6_other_project")
        assert foreign.returncode == 1 and not foreign.stdout.strip(), \
            "模式命中但路径不属于本项目时必须不匹配（F48 的误杀就在这）"
        assert child.poll() is None
    finally:
        child.kill()
        child.wait()


def test_stop_reports_not_found_honestly(monkeypatch, capsys):
    """复现 F48 死分支：旧版无论成败都打「已处理」。无匹配要如实报「未找到」。"""
    import app_ctl

    monkeypatch.setattr(app_ctl, "run_ps1_full", lambda name, *args: ("", 1))
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: False)
    monkeypatch.setattr(app_ctl, "time", types.SimpleNamespace(sleep=lambda *_: None))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("未找到匹配进程") == 2, "守护与服务两条都要如实报告"
    assert "已处理" not in out, "说反话的「已处理」不许再出现"


def test_stop_reports_killed_pids(monkeypatch, capsys):
    """有匹配且杀掉时，报真实 pid，而不是含糊的「已处理」。"""
    import app_ctl

    monkeypatch.setattr(app_ctl, "run_ps1_full",
                        lambda name, *args: ("killed 4321\n", 0))
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: False)
    monkeypatch.setattr(app_ctl, "time", types.SimpleNamespace(sleep=lambda *_: None))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert "已停止（pid 4321）" in out


# ================================================================ F49

def test_enum_hint_b_metric_wording():
    """复现 F49：B 口径不是「中文汉字数」，是中文字符数+英文单词数（DESIGN §5.1）。

    措辞与 compliance_report.measure 的实现锚在同一句上，抽出的 metric
    不会再把英文占比大的稿高估成假超限。
    """
    src = (TOOLS / "extract_journal.py").read_text(encoding="utf-8")
    assert "中文汉字数" not in src, "误导性措辞必须删干净"
    line = next(l for l in src.splitlines()
                if l.strip().startswith("-") and "limit_total_metric" in l)
    assert "中文字符数+英文单词数" in line and "Word" in line, line
    # 实现锚点：B = 中文字符数 + 英文单词数
    assert compliance_report.measure("中文 english words", "B") == 2 + 2


# ================================================================ F50

def test_propose_types_truncation_log_and_trace(monkeypatch):
    """复现 F50：候选超上限不是「AI 跳过」，是截断续跑——日志如实，
    trace 记 truncated/dropped，被截块没有提案。"""
    def fake_chat(cfg, messages, max_tokens=2048):
        return '{"labels":[{"idx":0,"type":"heading1","conf":"high","why":"编号"}]}'

    monkeypatch.setattr(ai_assist, "_chat", fake_chat)
    cfg = {**ai_assist.DEFAULTS, "enabled": True,
           "base_url": "https://api.example.com/v1", "api_key": "k",
           "model": "m", "max_blocks": 1}
    blocks = [{"idx": i, "type": "body", "text": f"第{i}段内容"} for i in range(2)]
    logs, trace = [], {}
    prop = ai_assist.propose_types(blocks, blocks, cfg=cfg, log=logs.append,
                                   trace=trace)
    assert trace == {"truncated": True, "dropped": 1, "kept": 1}
    assert list(prop) == [0], "只应标注截断后的前 max_blocks 个候选"
    assert any("仅标注前 1 块" in m and "其余按规则判定" in m for m in logs)
    assert not any("AI 跳过" in m for m in logs), "说反话的日志不许再出现"

    trace2 = {}
    ai_assist.propose_types(blocks, blocks, cfg={**cfg, "max_blocks": 10},
                            trace=trace2)
    assert trace2 == {"truncated": False, "dropped": 0, "kept": 2}, \
        "未截断时 trace 也必须如实"


# ================================================================ F51

LOW_TEXT = "abcdefghijklmnopqrstuvwxyzabcd"          # 恰 30 个有效字符
# 每行 62 个非空白字符、分四行写入——单行太长会被 PyMuPDF 按页面宽度裁剪，
# 提取出来的字符数就不足 200 了
FULL_LINES = ["abcdefghijklmnopqrstuvwxyz0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ"
              ] * 4


def _make_pdf(path, page1_text, page2_lines):
    doc = fitz.open()
    p1 = doc.new_page()
    p1.insert_text((72, 72), page1_text)
    p2 = doc.new_page()
    for k, line in enumerate(page2_lines):
        p2.insert_text((72, 72 + 14 * k), line)
    doc.save(str(path))
    doc.close()


def test_low_text_pages_recorded(tmp_path):
    """复现 F51 旁路：30 字文字层 ≥ 20 → 不触发 OCR，旧版无任何留痕。
    现在必须进 meta.ocr_low_text_pages，正常页不陪榜。"""
    f = tmp_path / "thin.pdf"
    _make_pdf(f, LOW_TEXT, FULL_LINES)
    text, meta = ocr.extract_pdf_text(str(f))
    assert meta["ocr_low_text_pages"] == [{"page": 1, "chars": 30}], \
        meta["ocr_low_text_pages"]
    assert meta["ocr_failed_pages"] == []      # 没发生失败，不算失败


def test_low_text_cli_warns(tmp_path, monkeypatch, capsys):
    """CLI：低文字层页必须在 stderr 打「若为扫描页请人工核对」。"""
    f = tmp_path / "thin.pdf"
    _make_pdf(f, LOW_TEXT, FULL_LINES)
    out = tmp_path / "out.txt"
    monkeypatch.setattr(sys, "argv", ["ocr.py", str(f), "-o", str(out)])
    assert ocr.main() == 0
    captured = capsys.readouterr()
    assert "第 1 页文字层仅 30 字" in captured.err
    assert "若为扫描页请人工核对" in captured.err
    assert "人工核对" not in captured.out, "警告不进正文输出"


def test_no_low_text_warning_when_plenty(tmp_path):
    """两页都正常（>200 字）时不得产生假警告。"""
    f = tmp_path / "full.pdf"
    _make_pdf(f, FULL_LINES, FULL_LINES)
    _, meta = ocr.extract_pdf_text(str(f))
    assert meta["ocr_low_text_pages"] == []
    assert ocr.low_text_warnings(meta) == []


def test_draft_carries_low_text_pages():
    """extract_journal.build 把低文字层页写进草稿留痕，复核者看得到。"""
    meta = {"ocr_low_text_pages": [{"page": 2, "chars": 30}]}
    rules, _ = ej.build({}, {}, "无关正文", jid="demob6", name_zh="演示乙",
                        meta=meta)
    assert rules["verification"]["extraction"]["ocr_low_text_pages"] == \
        meta["ocr_low_text_pages"]


# ================================================================ F52

def test_chat_max_tokens_default_and_override(monkeypatch):
    """_chat 默认 2048（标签批不变），调用方可传大值。"""
    seen = {}

    def fake_post(cfg, payload, timeout):
        seen.update(payload)
        return {"choices": [{"message": {"content": "ok"}}]}

    monkeypatch.setattr(ai_assist, "_post", fake_post)
    cfg = {"model": "m", "timeout": 3}
    assert ai_assist._chat(cfg, [{"role": "user", "content": "x"}]) == "ok"
    assert seen["max_tokens"] == 2048, "默认值必须保持标签批原状"
    ai_assist._chat(cfg, [{"role": "user", "content": "x"}], max_tokens=8192)
    assert seen["max_tokens"] == 8192


def test_label_batch_keeps_default_tokens(monkeypatch):
    """标签批链路（propose_types）不传大值——F52 只给事实抽取放大额度。"""
    captured = {}

    def fake_chat(cfg, messages, max_tokens=2048):
        captured["max_tokens"] = max_tokens
        return '{"labels":[{"idx":0,"type":"heading1","conf":"high","why":"编号"}]}'

    monkeypatch.setattr(ai_assist, "_chat", fake_chat)
    cfg = {**ai_assist.DEFAULTS, "enabled": True,
           "base_url": "https://api.example.com/v1", "api_key": "k", "model": "m"}
    blocks = [{"idx": 0, "type": "body", "text": "一、引言"}]
    ai_assist.propose_types(blocks, blocks, cfg=cfg)
    assert captured["max_tokens"] == 2048


def test_ai_facts_truncated_falls_back_to_unverified(monkeypatch):
    """复现 F52：密稿截断 JSON 原本一次即失败。现在降级重试一次（只要值
    不要引文），无引文的值走 F15 通道进 unverified，而不是整链报废。"""
    calls = []

    def fake_chat(cfg, messages, max_tokens=2048):
        calls.append({"max_tokens": max_tokens, "system": messages[0]["content"]})
        if len(calls) == 1:
            return '{"facts":{"body_size_pt":10.5,"limit_total":40'   # 截断
        return '{"facts":{"body_size_pt":10.5,"limit_total":40}}'     # 降级：无引文

    monkeypatch.setattr(ai_assist, "_chat", fake_chat)
    logs = []
    facts, ev, notes, _raw = ej.ai_facts("稿约正文", {}, log=logs.append)
    assert facts == {"body_size_pt": 10.5, "limit_total": 40}
    assert ev == {}, "降级模式不带引文"
    assert len(calls) == 2 and all(c["max_tokens"] == 8192 for c in calls)
    assert "不要 evidence" in calls[1]["system"], "重试必须换成不要引文的提示"
    assert any("降级重试" in m for m in logs)

    # 值仍走 F15 无引文通道：进 unverified，不进规则集
    rules, report = ej.build(facts, ev, "稿约正文", jid="demob6c",
                             name_zh="演示丙", meta={})
    assert report["stats"]["unverified"] == 2
    assert report["stats"]["extracted"] == 0
    unv = {u["key"] for u in rules["verification"]["extraction"]["unverified"]}
    assert unv == {"body_size_pt", "limit_total"}


# ================================================================ F57

@pytest.fixture()
def app_module():
    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


@pytest.fixture()
def work(app_module, tmp_path, monkeypatch):
    w = tmp_path / "work"
    w.mkdir(parents=True)
    monkeypatch.setattr(app_module, "WORK", w)
    return w


@pytest.fixture()
def light_job(work):
    d = work / "a1b2c3d4e5f6"
    d.mkdir(parents=True)
    (d / "meta.json").write_text(json.dumps(
        {"job": d.name, "journal": "tjyj", "journal_name": "统计研究",
         "original": "论文.docx", "stem": "论文", "stage": "details"},
        ensure_ascii=False), encoding="utf-8")
    (d / "成品.txt").write_text("占位", encoding="utf-8")
    return d


def test_history_row_has_delete_with_confirm(client, light_job):
    """F57：历史页每行要有「删除」入口，且带确认（隐私承诺要能就地兑现）。"""
    html = client.get("/history").get_data(as_text=True)
    assert f"/job/{light_job.name}/delete" in html
    assert "删除" in html
    assert "confirm(" in html


def test_delete_route_removes_job_dir(client, work, light_job):
    r = client.post(f"/job/{light_job.name}/delete")
    assert r.status_code == 302 and "/history" in r.headers["Location"]
    assert not light_job.exists(), "任务目录没有删掉"
    page = client.get("/history").get_data(as_text=True)
    assert "论文.docx" not in page, "删除后历史页不得再列出该任务"


def test_delete_rejects_illegal_id_with_400(client):
    assert client.post("/job/zzz/delete").status_code == 400, "非法 id 是 400 不是 404"
    assert client.post("/job/abc123/delete").status_code == 400, "位数不足也是 400"


def test_delete_missing_job_is_404(client, work):
    assert client.post("/job/0123456789ab/delete").status_code == 404, \
        "格式合法但目录已不在（TTL 清掉）按 404 兜底"


def test_workspace_mask_is_static(client):
    """F57：遮罩去掉轮播假文案，改静态「正在处理，可在历史页查看进度」。"""
    html = client.get("/workspace").get_data(as_text=True)
    assert "正在处理，可在历史页查看进度" in html
    assert "processingText" not in html, "旧轮播的挂点必须移除"
    assert "已用" not in html, "秒表轮播文案不许残留"


# ================================================================ COMPSUB

def test_check_survives_missing_type_keys():
    """复现 COMPSUB：IR 块 / 规则集块缺 type 键不得再冒 KeyError，
    warning 必须点名是哪份数据的第几块，报告要渲染出来。"""
    ir = {"blocks": [{"idx": 0, "type": "body", "text": "正文段落"},
                     {"idx": 1, "text": "这一块没有类型"}]}
    rules = {"journal": {"name_zh": "演示刊"}, "citation": {"system": "numbered"},
             "blocks": [{"type": "body", "required": True}, {"required": True}]}
    res = compliance_report.check(ir, rules)
    assert any("稿件 IR 第 2 块缺 type 键" in w for w in res["warnings"])
    assert any("规则集 第 2 块缺 type 键" in w for w in res["warnings"])

    page = compliance_report.render(ir, rules, res)
    assert "缺 type 键" in page and "按未知类型处理" in page


def test_count_words_without_type_key():
    """count_words 是公开函数，缺 type 的块不得崩；未知类型不算 body。"""
    res = compliance_report.count_words([{"text": "正文"}], scope="body")
    assert res["A"] == 0, "缺 type 的块不是 body，不该计入正文口径"
    res2 = compliance_report.count_words([{"text": "正文", "type": "body"}],
                                         scope="body")
    assert res2["A"] == 2 and res2["B"] == 2


def test_cli_missing_type_no_traceback(tmp_path):
    """端到端：缺 type 的输入走 CLI 不崩（退出码 0），提示语在 stdout。"""
    ir = tmp_path / "bad.ir.json"
    rules = tmp_path / "bad.rules.json"
    ir.write_text(json.dumps({"blocks": [{"idx": 0, "type": "body", "text": "正文"},
                                         {"idx": 1, "text": "无类型"}]},
                             ensure_ascii=False), encoding="utf-8")
    rules.write_text(json.dumps({"journal": {"name_zh": "演示刊"},
                                 "citation": {"system": "numbered"},
                                 "blocks": [{"type": "body", "required": True},
                                            {"required": True}]},
                                ensure_ascii=False), encoding="utf-8")
    out = run_tool("compliance_report.py", ir, rules, "-o", tmp_path / "r.html")
    assert "缺 type 键" in out, "CLI 必须把 warning 喊出来"
    assert (tmp_path / "r.html").exists()


# ================================================================ F16NOISE

def _big_inherited(n=48):
    items = []
    for i in range(30):
        items.append({"key": f"typography.body.f{i}", "label": f"正文字段{i}",
                      "value": i, "section": "typography", "fact": True})
    for i in range(n - 30):
        items.append({"key": f"limits.f{i}", "label": f"限制字段{i}",
                      "value": i, "section": "limits", "fact": False})
    return items


def test_draft_page_layers_inherited_groups(app_module):
    """复现 F16NOISE：48+ 条继承披露按区块分组折叠、默认收起、计数徽标、
    KPI 保留总数——复核页重点不再被淹没。"""
    from flask import render_template

    inherited = _big_inherited()
    ctx = dict(jid="demob6d", rules={"assumptions": ["占位"] * len(inherited)},
               report={"stats": {"extracted": 3, "unverified": 0,
                                 "inherited": len(inherited), "dropped": 0},
                       "errors": [], "warnings": [], "ok_to_publish": True},
               ext={"source_file": "稿约.pdf", "baseline": "tjyj",
                    "text_method": "text_layer", "unverified": [],
                    "ocr_failed_pages": [], "ocr_low_text_pages": []},
               journal={"name_zh": "演示丁"}, evidence={}, inherited=inherited,
               dropped=[], ai_on=False)
    with app_module.app.test_request_context("/journals/draft/demob6d"):
        page = render_template("journal_draft.html", **ctx)
    assert "按区块折叠" in page and "逐条可达" in page
    assert f'<span class="badge">{len(inherited)}</span>' in page, "总数徽标"
    assert "事实字段 30 项" in page and "结构字段 18 项" in page, "分层计数"
    assert page.count('class="badge"') == 1 + 2, "总数 + 每个区块各一枚徽标"
    assert "<details open" not in page, "48+ 条时所有分组默认收起"
    assert page.count("<details") == 1 + 2, "外层折叠 + typography/limits 两组"
    for it in inherited[:3]:
        assert it["label"] in page, "逐条仍可达（渲染进各组表格）"


def test_disclosure_count_not_reduced():
    """展示分层不许以删披露为代价：全量 inherited 与 assumptions 照旧落盘。"""
    rules, report = ej.build({}, {}, "无关正文", jid="demob6e",
                             name_zh="演示戊", meta={})
    inherited = rules["verification"]["extraction"]["inherited"]
    assert inherited, "F16 全量继承披露不能回退"
    assert rules["assumptions"], "assumptions 必须全量落盘"
    assert report["stats"]["inherited"] == len(inherited)
    assert any("逐字段清单" in a for a in rules["assumptions"]), \
        "结构字段汇总条仍在"


# ================================================================ SCRUBEXE

def test_scrub_covers_interpreter_dirs():
    """复现 SCRUBEXE：标准库帧残留解释器安装目录（盘符结构）必须一并抹掉。"""
    prefix, posix_prefix = str(Path(sys.prefix)), Path(sys.prefix).as_posix()
    exe_dir, posix_exe = str(Path(sys.executable).parent), Path(sys.executable).as_posix()

    detail = (f'Traceback (most recent call last):\n'
              f'  File "{prefix}\\Lib\\traceback.py", line 1, in boom\n'
              f'ValueError: 数据有问题')
    out = app._scrub_local_paths(detail)
    assert prefix not in out, f"解释器目录残留：{out}"
    assert "\\Lib\\traceback.py" in out and "ValueError: 数据有问题" in out, \
        "异常信息本身要保留，只抹路径"

    out2 = app._scrub_local_paths(f"ran {exe_dir}\\python.exe and {posix_prefix}/Lib/x")
    assert exe_dir not in out2 and posix_prefix not in out2, "正斜杠写法也要抹"
    assert posix_exe not in app._scrub_local_paths(f"at {posix_exe}")

    # 既有行为不回退：仓库根与用户主目录照旧替换
    out3 = app._scrub_local_paths(f"see {app.ROOT}\\app.py and {Path.home()}\\x")
    assert str(app.ROOT) not in out3 and str(Path.home()) not in out3
    assert "app.py" in out3


# ================================================================ BKB024

def test_backoff_comments_match_implementation():
    """复现 BKB024：注释说「5 次共约 0.44s」，实际 4 次退避 ≈0.24s（FX1 实测
    0.247s），第 5 次直接 raise。两处注释都要与实现一致；行为本身不变。

    app.py 的退避循环后来被并行修复收拢进 replace_with_retry（BKB024 的
    注释随之落在该函数的 docstring 上），所以这里对两份文件分别核对：
    虚假说法一处不留，如实说法两处都在。
    """
    for name in ("app.py", "tools/docx_meta.py"):
        src = (ROOT / name).read_text(encoding="utf-8")
        assert "0.44s" not in src, f"{name} 不得残留与实现不符的 0.44s 说法"
        assert "共约 0.24s" in src and "最后一次直接抛出" in src, name
    # 实现锚点：attempt==4 直接 raise，只退避 0.02+0.04+0.08+0.1
    assert abs(sum((0.02, 0.04, 0.08, 0.1, 0.2)[:4]) - 0.24) < 1e-9
    app_src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "def replace_with_retry" in app_src and "time.sleep(delays[attempt])" in app_src
