"""批七「一致性与解析日志」三问回归钉（WEBLOG2 / SIZEKEYS / PLACEMENT）。

- WEBLOG2  生产链路拿不到 AI 截断提示：批六 F50 已把 trace 接进 propose_types
  （截断数值落 ir["ai"]，test_fix_f50_trace_wiring.py 钉住），但 parse() 调
  apply_ai 没传 log——截断、调用失败、采纳数这些日志行全部落到 no-op，网页
  流水线（app.py 以子进程跑 parse_docx.py）与 CLI 都看不见。→ parse() 传
  log=print，AI 日志进解析 stdout；回归同时钉 trace 侧（ir["ai"] 含
  truncated）与 log 侧（stdout 出现截断提示）。
- SIZEKEYS  SIZE_KEYS 在 tools 侧两份且键集不一致：normalize_rulesets 只认
  四把核心键，references_size_pt / note_size_pt 的字符串值被静默放过（audit
  侧六键会报警，normalize 侧漏检）。→ 两处统一为六键并集并互指；回归断言
  两份名单相等、字符串型 references/note 字号现在会报警、统一对 23 本在库
  规则集零漂移（在库数据没有任何命中，实测扫描为空）。
- PLACEMENT  数据在用的 title_page / separate_file / submission_system 三个
  placement 值未入 journal-schema.md §9 枚举且无消费方（23 本实查：
  title_page 6 处在 blocks.corresponding_author，separate_file 12 处、
  submission_system 8 处在 personal_fields）。→ 枚举补三值并各配一句含义、
  注明消费方现状；回归断言枚举罩住数据在用的每一个值（blocks +
  personal_fields 口径）。

夹具姿势：parse_docx 走子进程补丁 _chat（同 test_fix_f50_trace_wiring.py）；
normalize/audit 照 test_fix_toolchain_isolation.py 在进程内 import，
nr.main 只写 tmp 目录，不触在库规则集。
"""

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import ROOT, RULES, TOOLS, minimal_doc

SCHEMA_MD = ROOT / "rules" / "journal-schema.md"
sys.path.insert(0, str(TOOLS))


# ---------------------------------------------------------------- WEBLOG2

# 子进程内执行的驱动脚本（同 test_fix_f50_trace_wiring.py 姿势）：
#   argv = [tools目录, docx路径, IR输出路径]
# 对 ai_assist._chat 的补丁打在 parse 所在的子进程内；对第一个候选块回一条
# heading1 提案，让截断与提案应用两条路都走到。
PATCH_SCRIPT = """\
import json
import sys
from pathlib import Path

sys.path.insert(0, sys.argv[1])
import ai_assist
import parse_docx


def fake_chat(cfg, messages, max_tokens=2048):
    payload = json.loads(messages[1]["content"])
    idx = payload["blocks"][0]["idx"]
    return json.dumps({"labels": [{"idx": idx, "type": "heading1",
                                   "conf": "high", "why": "编号"}]},
                      ensure_ascii=False)


ai_assist._chat = fake_chat

cfg = {**ai_assist.DEFAULTS, "enabled": True,
       "base_url": "https://api.example.com/v1", "api_key": "k",
       "model": "stub-model", "max_blocks": 2,
       "batch_size": 2, "retries": 0}

docx, out = Path(sys.argv[2]), Path(sys.argv[3])
ir = parse_docx.parse(str(docx), ai_cfg=cfg)
out.write_text(json.dumps(ir, ensure_ascii=False, indent=2), encoding="utf-8")
print("ok")
"""


def _run_parse_with_truncation(tmp_path, n_paragraphs=3):
    """骨架＋n 段正文，max_blocks=2：候选必超上限，触发截断。

    骨架的摘要段（front_matter）与正文段（body）都落在候选集里（标题规则
    高置信、不进候选，同 F50 用例口径），n=3 时候选 ≥4 > 2。
    """
    docx = tmp_path / "manuscript.docx"
    minimal_doc(docx, extra_paragraphs=[
        f"第{i}段正文内容用于补漏候选。" for i in range(n_paragraphs)])
    ir_path = tmp_path / "manuscript.ir.json"
    script = tmp_path / "driver.py"
    script.write_text(PATCH_SCRIPT, encoding="utf-8")
    r = subprocess.run(
        [sys.executable, str(script), str(TOOLS), str(docx), str(ir_path)],
        capture_output=True, text=True, cwd=str(ROOT),
    )
    assert r.returncode == 0, f"子进程失败：\n{r.stdout[-1500:]}\n{r.stderr[-1500:]}"
    return json.loads(ir_path.read_text(encoding="utf-8")), r.stdout


def test_weblog2_truncation_lands_in_ir_ai_record(tmp_path):
    """trace 侧（F50 已修，防回退）：截断场景下 ir["ai"] 含 truncated 信息。

    审计不变量：kept + dropped == candidates（候选总数）。
    """
    ir, _ = _run_parse_with_truncation(tmp_path)
    ai = ir["ai"]
    assert ai["used"] is True
    tr = ai["truncation"]
    assert tr["truncated"] is True, tr
    assert tr["kept"] == 2, tr
    assert tr["kept"] + tr["dropped"] == ai["candidates"], (tr, ai["candidates"])
    assert ai["proposed"] >= 1 and ai["applied"] >= 1, "截断与提案应用共存"


def test_weblog2_truncation_log_reaches_parse_stdout(tmp_path):
    """log 侧（本批修复）：截断提示必须出现在解析 stdout。

    修复前：parse() 调 apply_ai 不传 log，apply_ai 与 propose_types 里的
    log 全部落到 no-op——候选超上限这件事在整条生产链路上没有任何一句
    人话提示。修复后（log=print）：ai_assist 的截断行与 apply_ai 的采纳
    汇总都进解析 stdout。
    """
    _, stdout = _run_parse_with_truncation(tmp_path)
    assert "超过上限" in stdout and "仅标注前" in stdout, \
        f"截断提示没进解析 stdout：\n{stdout[-1500:]}"
    assert "AI 补漏" in stdout, f"采纳汇总没进解析 stdout：\n{stdout[-1500:]}"


# ---------------------------------------------------------------- SIZEKEYS

def test_sizekeys_two_copies_are_same_six_keys():
    """两份 SIZE_KEYS 必须是同一个六键并集（含 references/note 两键）。

    字符串字号能否被点名报警，完全取决于所在工具的这份名单——名单不一致，
    同一个值就会在体检报警、在归整漏检（SIZEKEYS 的原始病灶）。
    """
    import audit_rulesets as ar
    import normalize_rulesets as nr
    assert nr.SIZE_KEYS == ar.SIZE_KEYS, \
        (f"两份名单漂移：normalize={sorted(nr.SIZE_KEYS)} "
         f"audit={sorted(ar.SIZE_KEYS)}")
    assert {"references_size_pt", "note_size_pt"} <= nr.SIZE_KEYS
    assert len(ar.SIZE_KEYS) == 6


def test_sizekeys_normalize_warns_on_string_references_and_note_size(tmp_path, capsys):
    """normalize 遇到字符串型 references_size_pt / note_size_pt 必须报警。

    修复前：这两键不在 normalize 的 SIZE_KEYS 里，字符串值被静默放过——
    与 audit 侧同一份数据两种结论。字符串字号没法归整，但「没法归整」要
    说出来，不能装没看见。
    """
    import normalize_rulesets as nr
    rs = {
        "journal": {"id": "sizekeys", "name_zh": "字号键检查刊", "language": "zh"},
        "typography": {
            "body": {"font_cn": "宋体", "size_pt": 10.5},
            "references": {"references_size_pt": "九号"},
            "footnote": {"note_size_pt": "小五"},
        },
        "verification": {"status": "verified", "source_url": "https://example.com"},
        "blocks": [{"type": "title", "required": True}],
    }
    (tmp_path / "sizekeys.json").write_text(
        json.dumps(rs, ensure_ascii=False), encoding="utf-8")

    nr.TYPE_WARNINGS.clear()
    try:
        assert nr.main(tmp_path) == 0, "字符串字号是警告，不是失败"
        warns = list(nr.TYPE_WARNINGS)
    finally:
        nr.TYPE_WARNINGS.clear()
    capsys.readouterr()  # main 的汇总打印与本用例无关，吞掉
    paths = {p for p, _ in warns}
    assert "typography.references.references_size_pt" in paths, warns
    assert "typography.footnote.note_size_pt" in paths, warns


def test_sizekeys_unify_flags_nothing_in_shipped_rulesets():
    """零漂移守卫：统一键集后，23 本在库规则集不新增任何字符串字号报警。

    在库数据目前没有任何六键字符串值（修前实测扫描为空），统一名单纯属
    「今后不再漏检」，不得突然对存量数据发出新警告。
    """
    import normalize_rulesets as nr
    nr.TYPE_WARNINGS.clear()
    try:
        for p in sorted(RULES.glob("*.json")):
            rules = json.loads(p.read_text(encoding="utf-8"))
            for sec in ("typography", "headings", "tables", "figures"):
                nr.walk(rules.get(sec, {}), (sec,), lambda *a: None)
        assert nr.TYPE_WARNINGS == [], \
            f"统一键集给存量规则集带来了新报警：{nr.TYPE_WARNINGS}"
    finally:
        nr.TYPE_WARNINGS.clear()


# ---------------------------------------------------------------- PLACEMENT

def _placement_enum_row():
    text = SCHEMA_MD.read_text(encoding="utf-8")
    sec9 = text.split("## 9. personal_fields", 1)[1].split("## 10. ", 1)[0]
    row = next(ln for ln in sec9.splitlines() if "`placement` |" in ln)
    # 只从第二列（取值枚举）解析——第三列说明里的 `none`（代码防御值）
    # 不算枚举成员；含义断言仍对整行做
    cells = [c.strip() for c in row.split("|")]
    enum = set(re.findall(r"`([a-z_]+)`", cells[2]))
    return enum, row


def test_placement_enum_has_three_data_values_with_meanings():
    """枚举补三值并各配一句含义，且注明消费方现状（PLACEMENT）。"""
    enum, row = _placement_enum_row()
    for val, meaning in (("title_page", "独立标题页"),
                         ("separate_file", "单独成文件"),
                         ("submission_system", "投稿系统")):
        assert f"`{val}`" in row, f"placement 枚举缺 {val}：{row}"
        assert meaning in row, f"{val} 缺一句含义（应含「{meaning}」）：{row}"
    assert "消费方现状" in row, f"未注明消费方现状：{row}"
    assert enum >= {"byline", "footnote", "section", "author_page",
                    "back_matter", "system", "title_page", "separate_file",
                    "submission_system"}, f"枚举缺值：{sorted(enum)}"


# 反向覆盖（新枚举值各有在库使用）以 23 本全库为口径；公开检出 3 本示例
# 凑不齐这些值，跳过而不是报假回归。
@pytest.mark.skipif(len(list(RULES.glob("*.json"))) < 10,
                    reason="枚举反向覆盖以完整规则库为口径，公开检出跳过")
def test_placement_enum_covers_every_value_in_data():
    """枚举必须罩住数据在用的每一个值（F62 残余的复现口径：23 本实查）。

    口径与证据一致：扫 blocks 与 personal_fields 两处的 placement 值
    （submission.ai_declaration 的 before_references 是另一属性的枚举，
    不在本行管辖）。同时复现证据本身：三个新值确实各有在库使用。
    """
    enum, _ = _placement_enum_row()
    observed = {}

    def collect(node):
        if isinstance(node, dict):
            v = node.get("placement")
            if isinstance(v, str):
                observed[v] = observed.get(v, 0) + 1
            for x in node.values():
                collect(x)
        elif isinstance(node, list):
            for x in node:
                collect(x)

    for p in sorted(RULES.glob("*.json")):
        rules = json.loads(p.read_text(encoding="utf-8"))
        for sec in ("blocks", "personal_fields"):
            collect(rules.get(sec))

    missing = {v: n for v, n in observed.items() if v not in enum}
    assert not missing, f"数据在用但 §9 枚举没有：{missing}"
    for val in ("title_page", "separate_file", "submission_system"):
        assert observed.get(val, 0) > 0, f"{val} 在 23 本实查中应有在库使用"
