"""审读修复回归钉：解析 stdout 在网页流水线的落点（b7 审读⑤，WEBLOG2 收尾）。

WEBLOG2 让 parse() 把 AI 补漏事件（调用失败/采纳数）log 进解析 stdout
（log=print），但网页流水线以子进程跑 parse_docx.py 且从不接收 stdout——
那些日志行在 server.log 仍无处可看，持久留痕只有截断数值（ir["ai"]，F50）。
→ app._stage_parse 把解析 stdout 逐行转进 app.logger（server.log），AI 补漏
事件与识别汇总（块统计、存疑标题）随之可在后台关窗后回溯；纯日志落点，
不改变流水线行为（IR 结构、解析参数均不动）。

夹具姿势：app.logger 的级别只有 setup_logging()（main 路径）才抬到 INFO，
测试里用 caplog.set_level 抬上去；run() 整体打桩，不碰真子进程。
"""

import json
import logging

import app as app_module

PARSE_STDOUT = (
    "=== manuscript.docx ===\n"
    "识别 6 个块\n"
    "\n"
    "-- AI 补漏（stub-model） --\n"
    "  采纳 2 条\n"
    "已写入 X:\\job\\manuscript.ir.json\n"
)


def _make_job(tmp_path, name="a1b2c3d4e5f6"):
    d = tmp_path / name
    d.mkdir()
    # 至少一个内容块：_stage_parse 现在会拒收 0 块空稿（避免交付空成品），
    # 本文件测的是 stdout 日志落点，不该被空稿守卫挡在日志逻辑之前。
    (d / "manuscript.ir.json").write_text(
        json.dumps({"source": "论文.docx",
                    "blocks": [{"idx": 0, "type": "body", "text": "正文"}]}),
        encoding="utf-8")
    return d


def test_parse_stdout_lands_in_server_log(monkeypatch, tmp_path, caplog):
    """复现审读⑤：_stage_parse 必须把解析 stdout 转进服务日志。

    修复前 run() 的返回值被丢弃，AI 补漏的「调用失败/采纳数」在网页流水线
    任何地方都看不见。"""
    app_module.app.config["TESTING"] = True
    d = _make_job(tmp_path)

    calls = []

    def fake_run(args):
        calls.append([str(a) for a in args])
        if "parse_docx.py" in str(args[0]):
            return PARSE_STDOUT
        return ""                                  # compliance_report 一步

    monkeypatch.setattr(app_module, "run", fake_run)

    logger_name = app_module.app.logger.name
    with caplog.at_level(logging.INFO, logger=logger_name):
        ir = app_module._stage_parse(d, "tjyj", None)

    joined = "\n".join(r.getMessage() for r in caplog.records)
    assert "[parse]" in joined, f"解析 stdout 没进日志：{caplog.records}"
    assert "[parse] -- AI 补漏（stub-model） --" in joined, joined
    assert "[parse] 采纳 2 条" in joined, joined
    assert "[parse] 识别 6 个块" in joined, "识别汇总一并可回溯"
    # 流水线行为不变：两步子进程照常调用，IR 原样返回
    assert any("parse_docx.py" in " ".join(c) for c in calls)
    assert any("compliance_report.py" in " ".join(c) for c in calls)
    assert ir["source"] == "论文.docx"
    assert ir["blocks"][0]["text"] == "正文"


def test_parse_stdout_blank_lines_produce_no_log_records(monkeypatch, tmp_path,
                                                         caplog):
    """空行/纯空白行不产生日志记录；stdout 为空串也不出错。"""
    app_module.app.config["TESTING"] = True
    d = _make_job(tmp_path, name="b2c3d4e5f6a1")
    monkeypatch.setattr(app_module, "run", lambda args: "\n  \n")

    logger_name = app_module.app.logger.name
    with caplog.at_level(logging.INFO, logger=logger_name):
        app_module._stage_parse(d, "tjyj", None)
    parse_lines = [r.getMessage() for r in caplog.records
                   if r.getMessage().startswith("[parse]")]
    assert parse_lines == [], parse_lines
