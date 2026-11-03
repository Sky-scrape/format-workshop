"""F50 收口回归钉：parse_docx.apply_ai 把截断 trace 接进 IR 的 ai 记录。

审读发现：ai_assist.propose_types 的 trace 出参（truncated/dropped/kept）
已实现并有测试（test_fix_b6_services.py::test_propose_types_truncation_log_and_trace），
但唯一调用方 tools/parse_docx.py 的 apply_ai 不传 trace，被截断的候选块
数量只存在于 stdout 日志行「仅标注前 M 块」，审计 JSON 无留痕。

修复：apply_ai 调 propose_types 时传入 trace dict，并把结果作为
ai_trace["truncation"] 并进返回值——parse() 把返回值整体放进 ir["ai"]，
截断信息随之落库。审计不变量：kept + dropped == candidates（候选总数）。

测试姿势与 test_fix_page_setup_sections.py 相同：子进程直传参数列表、
不经 shell；对 ai_assist._chat 的补丁必须打在 parse 所在的子进程内，
父进程只读落盘的 IR。
"""

import json
import subprocess
import sys

from conftest import ROOT, TOOLS, minimal_doc

# 子进程内执行的驱动脚本：
#   argv = [tools目录, docx路径, IR输出路径, max_blocks, chat模式]
# chat 模式：
#   reply  —— 对消息里第一个候选块回一条 heading1 提案（走「提案应用」路径）
#   raise  —— 抛异常（retries=0 时 AI 调用失败，propose_types 返回空 dict，
#             走「无有效提案」短路路径）
PATCH_SCRIPT = """\
import json
import sys
from pathlib import Path

sys.path.insert(0, sys.argv[1])
import ai_assist
import parse_docx

max_blocks = int(sys.argv[4])
mode = sys.argv[5]

def fake_chat(cfg, messages, max_tokens=2048):
    if mode == "raise":
        raise RuntimeError("stub 网络失败")
    payload = json.loads(messages[1]["content"])
    idx = payload["blocks"][0]["idx"]
    return json.dumps({"labels": [{"idx": idx, "type": "heading1",
                                   "conf": "high", "why": "编号"}]},
                      ensure_ascii=False)

ai_assist._chat = fake_chat

cfg = {**ai_assist.DEFAULTS, "enabled": True,
       "base_url": "https://api.example.com/v1", "api_key": "k",
       "model": "stub-model", "max_blocks": max_blocks,
       "batch_size": 2, "retries": 0}

docx, out = Path(sys.argv[2]), Path(sys.argv[3])
ir = parse_docx.parse(str(docx), ai_cfg=cfg)
out.write_text(json.dumps(ir, ensure_ascii=False, indent=2), encoding="utf-8")
print("ok")
"""


def _run_parse(tmp_path, n_paragraphs, max_blocks, mode="reply"):
    """造一份骨架＋n_paragraphs 段正文的稿子，在子进程里跑 parse，
    返回落盘 IR。骨架的摘要段（front_matter）与正文段（body）都落在
    AI_UPGRADE_FROM 候选集里，标题是规则高置信、不进候选。"""
    docx = tmp_path / "manuscript.docx"
    minimal_doc(docx,
                extra_paragraphs=[f"第{i}段正文内容用于补漏候选。" for i in range(n_paragraphs)])
    ir_path = tmp_path / "manuscript.ir.json"
    script = tmp_path / "driver.py"
    script.write_text(PATCH_SCRIPT, encoding="utf-8")
    r = subprocess.run(
        [sys.executable, str(script), str(TOOLS), str(docx), str(ir_path),
         str(max_blocks), mode],
        capture_output=True, text=True, cwd=str(ROOT),
    )
    assert r.returncode == 0, f"子进程失败：\n{r.stdout[-1500:]}\n{r.stderr[-1500:]}"
    return json.loads(ir_path.read_text(encoding="utf-8"))


def test_truncation_lands_in_ir_ai_record(tmp_path):
    """候选超过上限：ir["ai"]["truncation"] 必须如实记录 dropped/kept。

    修复前：apply_ai 不传 trace，ir["ai"] 里没有任何截断字段，
    dropped 数只存在于 stdout 日志行。审计不变量：kept+dropped==candidates。
    """
    ir = _run_parse(tmp_path, n_paragraphs=3, max_blocks=2)
    ai = ir["ai"]
    assert ai["used"] is True
    tr = ai["truncation"]
    assert tr["truncated"] is True, tr
    assert tr["kept"] == 2, tr
    assert tr["kept"] + tr["dropped"] == ai["candidates"], (tr, ai["candidates"])
    # 截断与提案应用共存：提案、采纳数照常落库
    assert ai["proposed"] >= 1 and ai["applied"] >= 1


def test_no_truncation_reports_full_counts(tmp_path):
    """未截断时也必须如实：truncated=False、dropped=0、kept=候选总数。"""
    ir = _run_parse(tmp_path, n_paragraphs=3, max_blocks=10)
    ai = ir["ai"]
    assert ai["used"] is True
    tr = ai["truncation"]
    assert tr == {"truncated": False, "dropped": 0, "kept": ai["candidates"]}, (tr, ai["candidates"])
    assert tr["kept"] >= 3, "骨架摘要段与三段正文至少 4 个候选"


def test_ai_failure_path_still_leaves_trace(tmp_path):
    """AI 调用失败走「无有效提案」短路时，trace 照样落库：
    used=False 的 ai 记录里也有 truncation——审计不留黑窗。"""
    ir = _run_parse(tmp_path, n_paragraphs=2, max_blocks=2, mode="raise")
    ai = ir["ai"]
    assert ai["used"] is False
    assert ai["reason"] == "无有效提案（未配置、调用失败或返回不合规）"
    tr = ai["truncation"]
    # 短路返回无 candidates 键（既有形状不变）；骨架摘要段+两段正文共 3 个
    # 候选 > max_blocks 2，即失败发生在截断之后——trace 照如实记录
    assert tr == {"truncated": True, "dropped": 1, "kept": 2}, \
        f"3 个候选对 max_blocks=2 应记截断，实际 {tr}"
