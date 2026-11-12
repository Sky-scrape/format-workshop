"""输入稿版式体检（只读）：把稿件与目标刊规则集对撞，列出排版前就该知道的偏差。

与「审计报告」的分工：
  - 审计报告回答「这一轮改了什么」，读的是渲染器写出的 *.audit.json；
  - 版式体检回答「输入稿哪里不合本刊要求」，读的是原稿本身，一个字都不改。
投稿人先看体检、再决定要不要排版；编辑也能拿它核对退回稿。

**为什么只用「显式设置」判偏差**：真实稿件大量依赖 Word 的样式继承——正文
run 没有 w:sz、段落没有 w:jc 是常态，靠 Normal 样式渲染出来可能完全合规。
把这些一律判红，报告会满屏假问题，用户很快就不看了。所以：
  - run 的字体/字号/加粗没有显式设置 → 记 inherited（随样式），不计偏差；
  - 段落对齐/行距/缩进没有显式设置 → 同样记 inherited；
  - 只有「显式写成别的值」才报 auto（渲染器会改）或 manual（本工具不修）。
inherited 不是「没问题」，是「本工具无法凭直接格式断言」——报告里单列计数，
排版时渲染器仍会按本刊要求显式写入。

**为什么期望值不另写一份**：体检与 render_from_ir 必须共用同一套口径，否则
「体检说没问题、排版却改了」或者反过来。本模块的期望值全部按
render_from_ir.main 的分支逐行取值；表格边框更进一步——在同一份 XML 的
深拷贝上直接跑 apply_table，再与 table_snapshot 对账。两边共用一段代码，
就不可能出现两套期望值（tests/test_preflight.py 的幂等用例守着这条）。

用法：
    py -3 tools/preflight.py <docx> <ruleset> -o preflight.json --html preflight.html
"""

import argparse
import copy
import html
import json
import re
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

import parse_docx
import xmlsafe
from render_from_ir import (
    DEFAULTS, _sdt_blocks, apply_table, block_text, elem_of, fmt_snapshot,
    head_line_pt, para_runs, run_style, strip_number_prefix, table_snapshot,
    text_locked, tpl_prefix_rx,
)
from report_theme import BASE
from rules_ref import resolvable

ALIGN_DESC = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}
SIDE_DESC = {"top": "上边距", "bottom": "下边距", "left": "左边距", "right": "右边距"}

# 行距规则里这几种都算「倍数行距」。python-docx 对 1.0/1.5/2.0 会分别报
# SINGLE / ONE_POINT_FIVE / DOUBLE，其余倍数报 MULTIPLE；按「是否倍数」判，
# 不按具体枚举值判，否则 1.5 倍会被误当成固定值。
MULTIPLE_RULES = {WD_LINE_SPACING.MULTIPLE, WD_LINE_SPACING.ONE_POINT_FIVE,
                  WD_LINE_SPACING.DOUBLE, WD_LINE_SPACING.SINGLE}

# 渲染器真正读取过的 typography 键。没在这个集合里的样式键，规则集里声明了
# 也不会被 main 套用——体检必须把它们列进 not_handled，不能假装检查过。
# abstract_en / keywords_en 在 abstract/keywords 分支里按 lang=="en" 读到，
# 算「渲染器会读」，不列入未消费（虽然当前解析器不会产出带 lang 的块）。
CONSUMED_TYPO = {
    "body", "title", "subtitle", "authors", "abstract", "keywords",
    "abstract_en", "keywords_en", "references", "references_heading",
    "author_note", "footnote", "title_en",
}

# SEMANTIC（render_from_ir）里补录块 → (typography 键, 对齐, 加粗)。
SEMANTIC = {
    "affiliations": ("authors", "center", False),
    "english_authors": ("authors", "center", False),
    "english_affiliations": ("abstract", "center", False),
    "classification": ("abstract", "left", False),
    "english_title": ("title", "center", False),
    "english_abstract": ("abstract", "both", False),
    "english_keywords": ("abstract", "both", False),
}

FIX_AUTO = "auto"
FIX_MANUAL = "manual"
FIX_OK = "ok"
FIX_INHERITED = "inherited"


# ---------------------------------------------------------------- 基础读值

def _ref(rules, *cands):
    """挑一个能被 rules_ref 解析的规则引用。

    体检条目的 rule 必须真的能落到规则集里的一个值上（DESIGN §6/§8.5 的
    同一条要求）。候选路径可能因刊而异（有的刊没有 typography.references），
    逐个试，全落空时退回 journal 根——报告里有个能点开的引用，好过坏引用。
    """
    for ref in cands:
        if ref and resolvable(rules, ref):
            return ref
    return "journal"


def _align_name(value):
    if value is None:
        return None
    name = getattr(value, "name", "")
    return {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
            "JUSTIFY": "both"}.get(name, name.lower())


def _body_line(rules):
    """正文行距目标：与 render_from_ir.main 的开头逐字同口径。

    返回 (倍数, 固定磅值)。倍数为真时固定磅值为 None；固定值缺省时落
    DEFAULTS["body_line_pt"]（代码约定，报告里要标缺省）。
    """
    body = (rules.get("typography") or {}).get("body") or {}
    ls = body.get("line_spacing") or {}
    mult = ls.get("value") if ls.get("mode") == "multiple" else None
    pt = None if mult else (ls.get("value_pt") or ls.get("equivalent_exact_pt")
                            or DEFAULTS["body_line_pt"][0])
    return mult, pt


def _line_ref(rules, owner, mult, pt, own):
    """行距条目的规则引用；own=False 表示落到了正文行距/代码缺省上。"""
    if mult:
        return _ref(rules, f"{owner}.line_spacing.value", owner)
    if own:
        return _ref(rules, f"{owner}.line_spacing.value_pt",
                    f"{owner}.line_spacing.equivalent_exact_pt", owner)
    body = "typography.body"
    body_ls = (rules.get("typography") or {}).get("body", {}).get("line_spacing") or {}
    if body_ls.get("mode") == "multiple":
        return _ref(rules, f"{body}.line_spacing.value", body)
    if body_ls.get("value_pt"):
        return _ref(rules, f"{body}.line_spacing.value_pt", body)
    if body_ls.get("equivalent_exact_pt"):
        return _ref(rules, f"{body}.line_spacing.equivalent_exact_pt", body)
    return _ref(rules, f"{body}.line_spacing.value_pt~", body)


def _line_def_note(rules):
    """正文行距落到代码缺省时的说明；有明文/实测值时返回空串。"""
    body_ls = (rules.get("typography") or {}).get("body", {}).get("line_spacing") or {}
    if body_ls.get("mode") == "multiple":
        return ""
    if body_ls.get("value_pt") or body_ls.get("equivalent_exact_pt"):
        return ""
    return f"本刊未声明固定行距，取代码缺省值 {DEFAULTS['body_line_pt'][0]:g}pt"


# ---------------------------------------------------------------- 结果收集

class Report:
    """收集每一条检查的原始结果，最后按「偏差逐块、合规/随样式聚合」出条目。"""

    def __init__(self):
        self.raw = []
        self.unchecked = []

    def add(self, idx, section, block, prop, rule, expected, actual, fix, note=""):
        self.raw.append({"idx": idx, "section": section, "block": block,
                         "property": prop, "rule": rule, "expected": expected,
                         "actual": actual, "fix": fix, "note": note})

    def skip(self, idx, block, reason):
        self.unchecked.append({"idx": idx, "block": block, "reason": reason})


def _run_item(rep, idx, section, block, prop, rule, target, target_desc,
              values, partial, fmt, note_ok="", note_bad="", fix_bad=FIX_AUTO,
              inherited_note=None):
    """一条 run 级检查（字体/字号/加粗）。

    values 是显式设置的不同取值；partial 表示还有 run 没显式设置。
    没有显式值 → inherited；显式值恰为目标但部分 run 随样式 → 也记
    inherited（不能断言整段合规）；显式值里有别的 → 偏差。
    """
    if not values:
        rep.add(idx, section, block, prop, rule, target_desc,
                "随样式（未显式设置）", FIX_INHERITED,
                inherited_note or "源 run 未显式设置该项，Word 按样式渲染；"
                                  "排版时会按本刊要求显式写入")
    elif values == {target}:
        if partial:
            rep.add(idx, section, block, prop, rule, target_desc,
                    "部分 run 随样式", FIX_INHERITED,
                    "部分 run 未显式设置该项，无法按直接格式断言整段合规；"
                    "排版时会按本刊要求显式写入")
        else:
            rep.add(idx, section, block, prop, rule, target_desc, target_desc,
                    FIX_OK, note_ok)
    else:
        actual = "、".join(fmt(v) for v in sorted(values, key=str))
        rep.add(idx, section, block, prop, rule, target_desc, actual, fix_bad,
                note_bad or "与本刊要求不符")


def _para_item(rep, idx, section, block, prop, rule, expected, actual, status,
               note_ok="", note_bad="", fix_bad=FIX_AUTO, inherited_note=None):
    """一条段落级检查（对齐/行距/缩进）。status: ok / bad / inherited。"""
    if status == FIX_OK:
        rep.add(idx, section, block, prop, rule, expected, actual, FIX_OK, note_ok)
    elif status == FIX_INHERITED:
        rep.add(idx, section, block, prop, rule, expected, actual, FIX_INHERITED,
                inherited_note or "源段落未显式设置该项，Word 按样式渲染；"
                                  "排版时会按本刊要求显式写入")
    else:
        rep.add(idx, section, block, prop, rule, expected, actual, fix_bad,
                note_bad or "与本刊要求不符")


def _line_status(snap):
    """（值, 类型）— 类型是 multiple / exact / at_least；未设置为 None。"""
    line, rule = snap.get("line"), snap.get("rule")
    if line is None:
        return None
    if rule in MULTIPLE_RULES:
        return (float(line), "multiple")
    if rule == WD_LINE_SPACING.AT_LEAST:
        return (float(line), "at_least")
    return (float(line), "exact")


def _line_ok(actual, target):
    if actual is None:
        return FIX_INHERITED
    if target.get("mult"):
        if actual[1] == "multiple" and abs(actual[0] - float(target["mult"])) < 0.02:
            return FIX_OK
        return "bad"
    want_rule = target.get("rule", "exact")
    if actual[1] == want_rule and abs(actual[0] - float(target["pt"])) < 0.02:
        return FIX_OK
    return "bad"


def _line_actual_desc(actual):
    if actual is None:
        return "未设置"
    if actual[1] == "multiple":
        return f"{actual[0]:g} 倍行距"
    if actual[1] == "at_least":
        return f"{actual[0]:g}pt 最小值"
    return f"{actual[0]:g}pt 固定值"


def _line_target_desc(target):
    if target.get("mult"):
        return f"{float(target['mult']):g} 倍行距"
    suffix = "最小值" if target.get("rule") == "at_least" else "固定值"
    return f"{float(target['pt']):g}pt {suffix}"


def _indent_info(snap):
    """读段落的缩进状态（字符数）。firstLine/hanging/left/right 都可能只有
    twips 写法，一律折算成字符（1 字符 ≈ 210 twips，与 set_chars_indent 的
    刻度一致）。左右缩进不能忽略：渲染器 clear_para_formatting 会删除整个
    w:ind，只看到 firstLine/hanging 就会把「有左缩进」误判成「无缩进、已合规」。
    """
    attrs = dict(snap["ind"]) if snap.get("ind") else {}

    def val(tag):
        raw = attrs.get(qn(tag))
        try:
            return int(raw)
        except (TypeError, ValueError):
            return None

    def chars(char_tag, twip_tag):
        c, t = val(char_tag), val(twip_tag)
        if c is not None:
            return c / 100.0
        return t / 210.0 if t is not None else 0.0

    fc, fl = val("w:firstLineChars"), val("w:firstLine")
    hc, hl = val("w:hangingChars"), val("w:hanging")
    left = chars("w:leftChars", "w:left")
    right = chars("w:rightChars", "w:right")
    # 首行/悬挂缩进之外还要带着 left/right：渲染器 clear_para_formatting 会删掉
    # 整个 w:ind，只比 firstLine/hanging 会让「首行 2 字符 + 左缩进 2 字符」判成
    # 合规，实际排版却把左缩进清掉。悬挂缩进的 left==hanging 是渲染器成对写入
    # 的契约，是否合规由 _indent_ok 按目标类型判断。
    if fc is not None:
        return ("first_line", fc / 100.0, left, right)
    if fl is not None:
        return ("first_line", fl / 210.0, left, right)
    if hc is not None:
        return ("hanging", hc / 100.0, left, right)
    if hl is not None:
        return ("hanging", hl / 210.0, left, right)
    if left or right:
        return ("block", left or right, left, right)
    return ("none", 0.0, 0.0, 0.0)


def _indent_ok(actual, target):
    kind, chars = target
    akind, achars, aleft, aright = actual
    if kind == "none":
        # firstLineChars=0 等价于没有缩进；但 left/right 非零时渲染器会清掉，
        # 不能因为首行缩进是 0 就报 ok。
        no_first = akind == "none" or abs(achars) < 0.01
        return FIX_OK if no_first and not aleft and not aright else "bad"
    if akind == "none":
        # 没有显式 first/hanging，Word 可能按样式继承；但显式 left/right 会
        # 被渲染器删掉，仍算会改。
        return "bad" if (aleft or aright) else FIX_INHERITED
    if akind != kind or abs(achars - chars) >= 0.05:
        return "bad"
    if kind == "hanging":
        # set_hanging_indent 同时写 left=hanging；left 对不上或 right 非零
        # 都意味着渲染器会改动源格式。
        if abs(aleft - chars) >= 0.05 or aright:
            return "bad"
    elif aleft or aright:
        return "bad"
    return FIX_OK


def _indent_actual_desc(actual):
    akind, achars, aleft, aright = actual
    if akind == "none":
        desc = "无"
    elif akind == "block":
        desc = "无首行/悬挂缩进"
    else:
        name = "首行缩进" if akind == "first_line" else "悬挂缩进"
        desc = f"{name} {achars:g} 字符"
    extra = []
    if aleft:
        extra.append(f"左缩进 {aleft:g} 字符")
    if aright:
        extra.append(f"右缩进 {aright:g} 字符")
    return desc + (f"（另有{'；'.join(extra)}）" if extra else "")


def _indent_target_desc(target):
    kind, chars = target
    if kind == "none":
        return "无缩进"
    name = "首行缩进" if kind == "first_line" else "悬挂缩进"
    return f"{name} {chars:g} 字符"


# ---------------------------------------------------------------- 段落检查

def _run_sets(p):
    """段落的显式 run 取值：中文/西文字体、字号(pt)、加粗。"""
    runs = [run_style(r) for r in para_runs(p)]
    return {
        "cn": {r[0] for r in runs if r[0] is not None},
        "latin": {r[1] for r in runs if r[1] is not None},
        "size": {r[2] / 2 for r in runs if r[2] is not None},
        "bold": {r[3] for r in runs if r[3] is not None},
        "partial": {k: any(r[i] is None for r in runs)
                    for k, i in (("cn", 0), ("latin", 1), ("size", 2), ("bold", 3))},
        "has_runs": bool(runs),
    }


def _check_spacing(rep, idx, section, block, p, target, rule, fallback="",
                   fix_bad=FIX_AUTO):
    """字间距（run 级 w:spacing）检查。target 是 twips 整数；None 表示目标为 0。

    set_run 对「规则集声明了 0」是清除已有 w:spacing，不是不处理；所以这里
    必须把 target=0 当成一个真目标，否则带字间距的模板切到 0pt 的刊时，
    体检会说「随样式/合规」，排版却悄悄清了字间距。
    """
    vals = []
    for r in para_runs(p):
        rpr = r.find(qn("w:rPr"))
        sp = rpr.find(qn("w:spacing")) if rpr is not None else None
        raw = sp.get(qn("w:val")) if sp is not None else None
        try:
            vals.append(int(raw) if raw is not None else None)
        except (TypeError, ValueError):
            vals.append(None)
    explicit = {v for v in vals if v is not None}
    partial = any(v is None for v in vals)

    if target is None:
        expected = "无字间距"
        if explicit:
            actual = "、".join(f"{v / 20:g}pt" for v in sorted(explicit))
            _para_item(rep, idx, section, block, "字间距", rule,
                       expected, actual, "bad",
                       note_bad="源 run 带有显式字间距，本刊要求无字间距；"
                                "排版时会清除" + fallback, fix_bad=fix_bad)
        else:
            _para_item(rep, idx, section, block, "字间距", rule,
                       expected, expected, FIX_OK)
        return

    expected = f"{target / 20:g}pt"
    if not explicit:
        _para_item(rep, idx, section, block, "字间距", rule,
                   expected, "随样式（未显式设置）", FIX_INHERITED,
                   inherited_note="源 run 未显式设置字间距，Word 按样式渲染；"
                                  "排版时会按本刊要求显式写入")
    elif explicit == {target}:
        if partial:
            _para_item(rep, idx, section, block, "字间距", rule,
                       expected, "部分 run 随样式", FIX_INHERITED,
                       inherited_note="部分 run 未显式设置字间距，无法按直接格式"
                                      "断言整段合规；排版时会按本刊要求显式写入")
        else:
            _para_item(rep, idx, section, block, "字间距", rule,
                       expected, expected, FIX_OK)
    else:
        actual = "、".join(f"{v / 20:g}pt" for v in sorted(explicit))
        _para_item(rep, idx, section, block, "字间距", rule,
                   expected, actual, "bad",
                   note_bad="显式字间距与本刊要求不符" + fallback,
                   fix_bad=fix_bad)


def _check_paragraph(rep, blk, p, section, spec):
    """按 spec 检查一个段落；spec 缺某个键就跳过该项（该刊没声明/渲染器不读）。"""
    idx = blk["idx"]
    block = blk["type"]
    snap = fmt_snapshot(p)
    rs = _run_sets(p)

    if rs["has_runs"]:
        if spec.get("cn"):
            target, rule, fallback = spec["cn"]
            _run_item(rep, idx, section, block, "中文字体", rule, target,
                      f"{target}", rs["cn"], rs["partial"]["cn"], str,
                      note_bad="显式设置的字体与本刊要求不符" + fallback,
                      inherited_note="源 run 未显式设置中文字体，Word 按样式渲染；"
                                     "排版时会按本刊要求显式写入")
        if spec.get("latin"):
            target, rule, fallback = spec["latin"]
            _run_item(rep, idx, section, block, "西文字体", rule, target,
                      f"{target}", rs["latin"], rs["partial"]["latin"], str,
                      note_bad="显式设置的西文字体与本刊要求不符" + fallback,
                      inherited_note="源 run 未显式设置西文字体，Word 按样式渲染；"
                                     "排版时会按本刊要求显式写入")
        if spec.get("size"):
            target, rule, fallback = spec["size"]
            target = round(target * 2) / 2
            _run_item(rep, idx, section, block, "字号", rule, target,
                      f"{target:g}pt", rs["size"], rs["partial"]["size"],
                      lambda v: f"{v:g}pt",
                      note_bad="显式设置的字号与本刊要求不符" + fallback)
        if spec.get("bold") is not None:
            target, rule, fallback, fix_bad = spec["bold"]
            _run_item(rep, idx, section, block, "加粗", rule, target,
                      "加粗" if target else "不加粗", rs["bold"],
                      rs["partial"]["bold"], lambda v: "加粗" if v else "不加粗",
                      note_bad="与本刊要求不符" + fallback, fix_bad=fix_bad)
        if "spacing" in spec:
            target, rule, fallback, fix_bad = spec["spacing"]
            _check_spacing(rep, idx, section, block, p, target, rule, fallback, fix_bad)

    if spec.get("align"):
        target, rule, fallback, fix_bad = spec["align"]
        actual = _align_name(snap.get("jc"))
        if actual is None:
            _para_item(rep, idx, section, block, "对齐", rule,
                       ALIGN_DESC[target], "随样式（未显式设置）", FIX_INHERITED,
                       inherited_note="源段落未显式设置对齐，Word 按样式渲染；"
                                      "排版时会按本刊要求显式写入")
        else:
            _para_item(rep, idx, section, block, "对齐", rule,
                       ALIGN_DESC[target], ALIGN_DESC.get(actual, actual),
                       FIX_OK if actual == target else "bad",
                       note_bad="显式对齐与本刊要求不符" + fallback, fix_bad=fix_bad)

    if spec.get("line"):
        target, rule, fallback, fix_bad = spec["line"]
        actual = _line_status(snap)
        status = _line_ok(actual, target)
        _para_item(rep, idx, section, block, "行距", rule,
                   _line_target_desc(target), _line_actual_desc(actual), status,
                   note_bad="显式行距与本刊要求不符" + fallback, fix_bad=fix_bad,
                   inherited_note="源段落未显式设置行距，Word 按样式渲染；"
                                  "排版时会按本刊要求显式写入")

    if spec.get("indent"):
        target, rule, fallback, fix_bad = spec["indent"]
        actual = _indent_info(snap)
        status = _indent_ok(actual, target)
        _para_item(rep, idx, section, block, "缩进", rule,
                   _indent_target_desc(target), _indent_actual_desc(actual), status,
                   note_bad="显式缩进与本刊要求不符" + fallback, fix_bad=fix_bad,
                   inherited_note="源段落未显式设置缩进，Word 按样式渲染；"
                                  "排版时会按本刊要求显式写入")


def _check_prefix(rep, blk, p, section, level, hd, fix_bad=FIX_AUTO):
    """标题编号前缀是否存在。模板由 rules.headings.levels 给出。"""
    levels = hd.get("levels") or []
    if level < 1 or len(levels) < level:
        return
    text = block_text(p).strip()
    if not text:
        return
    # 未编号节（「引言」「结论」）：渲染器把整段换成 intro_title_text，没有编号。
    # 判据与渲染器同一函数（strip_number_prefix）：稿件可能写成「一、引言」，
    # 只要 core 命中 unnumbered_sections 就按未编号处理。
    if level == 1:
        core, _stripped = strip_number_prefix(text, level, hd)
        norm = "".join(core.split())
        if norm in {"".join(str(s).split()) for s in hd.get("unnumbered_sections") or []}:
            return
    tpl = levels[level - 1]
    rx = tpl_prefix_rx(tpl)
    if rx is None:
        return
    ok = bool(rx.match(text))
    expected = f"「{str(tpl).strip()}」编号前缀"
    actual = "「" + text[:10] + ("…」" if len(text) > 10 else "」")
    rep.add(blk["idx"], section, blk["type"], "编号前缀", "headings.levels",
            expected, actual, FIX_OK if ok else fix_bad,
            "" if ok else "标题没有本刊要求的编号前缀（排版时会按规则集重新编号）")


# ---------------------------------------------------------------- 期望规格

def _body_spec(rules):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    cn = body.get("font_cn")
    size = body.get("size_pt")
    latin = body.get("font_latin") or "Times New Roman"
    mult, pt = _body_line(rules)
    line = {"mult": mult} if mult else {"pt": pt}
    indent = body.get("first_line_indent_chars", 2)
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, "typography.body.font_cn", "typography.body"), "")
    spec["latin"] = (latin, _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    if size:
        spec["size"] = (size, _ref(rules, "typography.body.size_pt", "typography.body"), "")
    spec["align"] = ("both", _ref(rules, "typography.body.align~", "typography.body"),
                     "", FIX_AUTO)
    body_ls = body.get("line_spacing") or {}
    own_line = bool(mult or body_ls.get("value_pt") or body_ls.get("equivalent_exact_pt"))
    spec["line"] = (line, _line_ref(rules, "typography.body", mult, pt, own_line),
                    "（" + _line_def_note(rules) + "）" if _line_def_note(rules) else "",
                    FIX_AUTO)
    if isinstance(indent, (int, float)):
        spec["indent"] = (("first_line", float(indent)),
                          _ref(rules, "typography.body.first_line_indent_chars", "typography.body"),
                          "", FIX_AUTO)
    return spec


def _heading_spec(rules, level):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    hd = rules.get("headings") or {}
    lv = hd.get("level1") if level == 1 else (hd.get(f"level{level}") or {})
    lv = lv or {}
    inline = level >= 4 and bool(lv.get("inline_with_body"))
    cn = lv.get("font_cn") or body.get("font_cn")
    size = lv.get("size_pt") or body.get("size_pt")
    align = lv.get("align") or ("center" if level == 1 else "left")
    if inline:
        # 行内四级标题按正文套版：字体字号行距缩进都取正文，只有编号前缀
        # 按 level4 模板重写。spec 必须与渲染器同一口径，否则幂等闸会假红。
        cn = body.get("font_cn")
        size = body.get("size_pt")
    mult, pt = _body_line(rules)
    if mult:
        line = {"mult": mult}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)
    elif not inline and lv.get("line_multiple") and size:
        line = {"pt": max(pt or 0, round(size * lv["line_multiple"]))}
        line_ref = _ref(rules, f"headings.level{level}.line_multiple", f"headings.level{level}")
    else:
        line = {"pt": pt}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)

    spec = {}
    cn_ref = (_ref(rules, "typography.body.font_cn", "typography.body") if inline
              else _ref(rules, f"headings.level{level}.font_cn",
                        "headings.level1" if level == 1 else f"headings.level{level}~",
                        "typography.body.font_cn"))
    if cn:
        spec["cn"] = (cn, cn_ref, "")
    size_ref = (_ref(rules, "typography.body.size_pt", "typography.body") if inline
                else _ref(rules, f"headings.level{level}.size_pt",
                          "headings.level1" if level == 1 else f"headings.level{level}~",
                          "typography.body.size_pt"))
    if size:
        spec["size"] = (size, size_ref, "")
    spec["align"] = (align, _ref(rules, f"headings.level{level}.align",
                                "headings.level1" if level == 1 else f"headings.level{level}~",
                                "headings.level1"), "", FIX_AUTO)
    spec["line"] = (line, line_ref, "", FIX_AUTO)
    bold = False if inline else lv.get("bold", False)
    if bold is not None:
        spec["bold"] = (bool(bold),
                        _ref(rules, "typography.body", "typography.body") if inline
                        else _ref(rules, f"headings.level{level}.bold",
                                  "headings.level1" if level == 1 else f"headings.level{level}~"),
                        "", FIX_AUTO)
    # 一级标题不缩进；二至四级按 indent_chars，未声明/为 0 时渲染器会清掉缩进。
    # 行内四级跟正文首行缩进，不是标题自己的 indent_chars。
    if level == 1:
        spec["indent"] = (("none", 0.0), _ref(rules, "headings.level1"), "", FIX_AUTO)
    elif inline:
        chars = body.get("first_line_indent_chars", 2)
        spec["indent"] = ((("first_line", float(chars)) if chars else ("none", 0.0)),
                          _ref(rules, "typography.body.first_line_indent_chars",
                               "typography.body"), "", FIX_AUTO)
    else:
        chars = lv.get("indent_chars")
        spec["indent"] = ((("first_line", float(chars)) if chars else ("none", 0.0)),
                          _ref(rules, f"headings.level{level}.indent_chars",
                              f"headings.level{level}~", "typography.body.first_line_indent_chars"),
                          "", FIX_AUTO)
    return spec


def _title_spec(rules, which):
    """主标题/副标题。渲染器固定居中、不加粗；规则集若声明了别的值，
    偏差只能记 manual——那是渲染器的已知缺口，不能许诺自动修正。"""
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    tp = typo.get(which)
    fallback = []
    if tp is None:
        tp = {}
        fallback = [f"{which}_font_cn", f"{which}_size_pt"]
    size = tp["size_pt"] if "size_pt" in tp else \
        (DEFAULTS[f"{which}_size_pt"][0] if which == "subtitle" else
         DEFAULTS["title_size_pt"][0])
    cn = tp["font_cn"] if "font_cn" in tp else \
        (DEFAULTS[f"{which}_font_cn"][0] if which == "subtitle" else
         DEFAULTS["title_font_cn"][0])
    decl_align = tp.get("align")
    decl_bold = tp.get("bold")
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, f"typography.{which}.font_cn",
                              f"typography.{which}~", "typography.body.font_cn"),
                      "（本刊未声明，取代码缺省值）" if "font_cn" in fallback else "")
    if size:
        spec["size"] = (size, _ref(rules, f"typography.{which}.size_pt",
                                  f"typography.{which}~", "typography.body.size_pt"),
                        "（本刊未声明，取代码缺省值）" if "size_pt" in fallback else "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    if decl_align and decl_align != "center":
        spec["align"] = (decl_align, _ref(rules, f"typography.{which}.align", f"typography.{which}"),
                         "；渲染器目前固定居中，排版后需手工核对", FIX_MANUAL)
    else:
        spec["align"] = ("center", _ref(rules, f"typography.{which}"), "", FIX_AUTO)
    if size:
        line_pt = round(size * DEFAULTS[f"{which}_line_mult"][0])
        spec["line"] = ({"pt": line_pt}, _ref(rules, f"typography.{which}"),
                        "（行距按字号×1.5 折算，代码约定）", FIX_AUTO)
    if decl_bold is True:
        spec["bold"] = (True, _ref(rules, f"typography.{which}.bold", f"typography.{which}"),
                        "；渲染器目前固定不加粗，排版后需手工核对", FIX_MANUAL)
    else:
        spec["bold"] = (False, _ref(rules, f"typography.{which}"), "", FIX_AUTO)
    spec["indent"] = (("none", 0.0), _ref(rules, f"typography.{which}"), "", FIX_AUTO)
    # 主标题字间距：渲染器在 front_matter 的首段读 letter_spacing_pt（副标题
    # 不读，所以只给 title 建检查项）。0 也是真目标——set_run 会清除已有字间距。
    if which == "title" and tp.get("letter_spacing_pt") is not None:
        pts = tp["letter_spacing_pt"]
        target = int(round(float(pts) * 20)) if pts else None
        spec["spacing"] = (target,
                           _ref(rules, "typography.title.letter_spacing_pt",
                                "typography.title"), "", FIX_AUTO)
    return spec


def _authors_spec(rules):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    au = typo.get("authors") or {}
    cn = au.get("font_cn", body.get("font_cn"))
    size = au.get("size_pt") or body.get("size_pt")
    mult, pt = _body_line(rules)
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, "typography.authors.font_cn", "typography.body.font_cn"), "")
    if size:
        spec["size"] = (size, _ref(rules, "typography.authors.size_pt", "typography.body.size_pt"), "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    if "align" in au and au.get("align"):
        spec["align"] = (au["align"], _ref(rules, "typography.authors.align",
                                          "typography.authors~"), "", FIX_AUTO)
    elif "align" not in au:
        spec["align"] = ("center", _ref(rules, "typography.authors~", "typography.body"), "", FIX_AUTO)
    # 渲染器作者行按 1.5 倍字号折算固定行距（代码约定）
    spec["line"] = ({"pt": max(pt or 0, round(size * DEFAULTS["authors_line_mult"][0]))},
                    _ref(rules, "typography.authors~", "typography.body"),
                    "（行距按字号×1.5 折算，代码约定）", FIX_AUTO)
    spec["bold"] = (False, _ref(rules, "typography.authors~", "typography.body"), "", FIX_AUTO)
    spec["indent"] = (("none", 0.0), _ref(rules, "typography.authors~", "typography.body"), "", FIX_AUTO)
    return spec


def _semantic_spec(rules, btype):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    key, align, bold = SEMANTIC[btype]
    tp = typo.get(key) or body
    size = tp.get("size_pt", body.get("size_pt"))
    latin = body.get("font_latin") or "Times New Roman"
    if btype == "english_title":
        ten = typo.get("title_en")
        if ten:
            tp = ten
            size = ten.get("size_pt", size)
            align = ten.get("align", align)
            bold = ten.get("bold", bold)
            latin = ten.get("font_latin", latin)
        else:
            # 没单列英文标题样式的刊：渲染器取「标题与副标题中较小的字号」，
            # 免得英文标题压过中文标题（render_from_ir 的 english_title 分支）
            sub = typo.get("subtitle") or {}
            sub_sz = sub.get("size_pt", size)
            if isinstance(sub_sz, (int, float)) and isinstance(size, (int, float)):
                size = min(size, sub_sz)
    mult, pt = _body_line(rules)
    line = {"mult": mult} if mult else {"pt": pt}
    rule = ("typography.title_en" if (btype == "english_title" and typo.get("title_en"))
            else f"typography.{key}" if typo.get(key) else "typography.body")
    cn = tp.get("font_cn", body.get("font_cn"))
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, f"{rule}.font_cn", rule), "")
    if size:
        spec["size"] = (size, _ref(rules, f"{rule}.size_pt", rule), "")
    spec["latin"] = (latin, _ref(rules, f"{rule}.font_latin~", "typography.body.font_latin~",
                                 "typography.body"),
                     "" if latin else "")
    spec["align"] = (align, _ref(rules, f"{rule}.align~", rule), "", FIX_AUTO)
    spec["line"] = (line, _line_ref(rules, "typography.body", mult, pt, False), "", FIX_AUTO)
    spec["bold"] = (bool(bold), _ref(rules, f"{rule}.bold~", rule), "", FIX_AUTO)
    spec["indent"] = (("none", 0.0), _ref(rules, rule), "", FIX_AUTO)
    return spec


def _abstract_spec(rules, btype):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    key = btype
    absp = typo.get(key) or typo.get("abstract") or {}
    cn = absp.get("font_cn", body.get("font_cn"))
    size = absp.get("size_pt", body.get("size_pt"))
    mult, pt = _body_line(rules)
    own_ls = (absp.get("line_spacing") or {})
    own_pt = own_ls.get("value_pt") or own_ls.get("equivalent_exact_pt")
    if mult:
        line = {"mult": mult}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)
    elif own_pt:
        line = {"pt": own_pt}
        line_ref = _ref(rules, f"typography.{key}.line_spacing.value_pt",
                        f"typography.{key}.line_spacing.equivalent_exact_pt", f"typography.{key}")
    else:
        line = {"pt": pt}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)
    base_rule = (f"typography.{key}" if typo.get(key)
                 else "typography.abstract" if typo.get("abstract") else "typography.body")
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, f"{base_rule}.font_cn", base_rule), "")
    if size:
        spec["size"] = (size, _ref(rules, f"{base_rule}.size_pt", base_rule), "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    spec["align"] = ("both", _ref(rules, base_rule), "", FIX_AUTO)
    spec["line"] = (line, line_ref, "", FIX_AUTO)
    indent = body.get("first_line_indent_chars", 2)
    if isinstance(indent, (int, float)):
        spec["indent"] = (("first_line", float(indent)),
                          _ref(rules, "typography.body.first_line_indent_chars", "typography.body"),
                          "", FIX_AUTO)
    return spec


def _author_note_spec(rules, p, label):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    an = typo.get("author_note") or {}
    size = an.get("size_pt") or body.get("size_pt")
    is_label = block_text(p).strip() == label
    own_ls = an.get("line_spacing") or {}
    own_pt = own_ls.get("value_pt") or own_ls.get("equivalent_exact_pt")
    line_pt = own_pt or round((size or 0) * DEFAULTS["author_note_line_mult"][0], 1)
    rule = _ref(rules, "typography.author_note", "typography.body")
    spec = {}
    if is_label:
        cn = an.get("label_font_cn", DEFAULTS["author_note_label_font_cn"][0])
        if cn:
            spec["cn"] = (cn, _ref(rules, "typography.author_note.label_font_cn~",
                                  "typography.author_note~", "typography.body.font_cn"),
                          "（本刊未声明，取代码缺省值）" if "label_font_cn" not in an else "")
        align = an.get("heading_align", "left") if "heading_align" in an else "left"
        if align:
            spec["align"] = (align, _ref(rules, "typography.author_note.heading_align~", rule),
                             "", FIX_AUTO)
        indent_chars = an.get("heading_indent_chars") or body.get("first_line_indent_chars", 2)
    else:
        cn = an.get("font_cn", body.get("font_cn"))
        spec["cn"] = (cn, _ref(rules, "typography.author_note.font_cn~", rule, "typography.body"),
                      "")
        spec["align"] = ("both", rule, "", FIX_AUTO)
        indent_chars = body.get("first_line_indent_chars", 2)
    if size:
        spec["size"] = (size, _ref(rules, "typography.author_note.size_pt~", rule), "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    if line_pt:
        spec["line"] = ({"pt": line_pt},
                        _ref(rules, "typography.author_note.line_spacing.value_pt~", rule),
                        "" if own_pt else "（行距按字号×1.6 折算，代码约定）", FIX_AUTO)
    if isinstance(indent_chars, (int, float)) and indent_chars:
        spec["indent"] = (("first_line", float(indent_chars)),
                          _ref(rules, "typography.author_note.heading_indent_chars~",
                              "typography.body.first_line_indent_chars"), "", FIX_AUTO)
    else:
        spec["indent"] = (("none", 0.0), rule, "", FIX_AUTO)
    return spec


def _reference_spec(rules):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    ref_tp = typo.get("references") or {}
    rl = (rules.get("citation") or {}).get("reference_list") or {}
    cn = ref_tp.get("font_cn", body.get("font_cn"))
    size = ref_tp.get("size_pt", body.get("size_pt"))
    mult, pt = _body_line(rules)
    own_ls = ref_tp.get("line_spacing") or {}
    own_pt = own_ls.get("value_pt") or own_ls.get("equivalent_exact_pt")
    if ref_tp.get("line_single"):
        line = {"mult": 1.0}
        line_ref = _ref(rules, "typography.references.line_single", "typography.references")
    elif mult:
        line = {"mult": mult}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)
    elif own_pt:
        line = {"pt": own_pt}
        line_ref = _ref(rules, "typography.references.line_spacing.value_pt",
                        "typography.references.line_spacing.equivalent_exact_pt",
                        "typography.references")
    else:
        line = {"pt": pt}
        line_ref = _line_ref(rules, "typography.body", mult, pt, False)
    spec = {}
    if cn:
        spec["cn"] = (cn, _ref(rules, "typography.references.font_cn", "typography.body.font_cn"), "")
    if size:
        spec["size"] = (size, _ref(rules, "typography.references.size_pt",
                                  "typography.body.size_pt"), "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    spec["align"] = ("both", _ref(rules, "typography.references", "typography.body"), "", FIX_AUTO)
    spec["line"] = (line, line_ref, "", FIX_AUTO)
    spec["bold"] = (False, _ref(rules, "typography.references", "typography.body"), "", FIX_AUTO)
    hang = rl.get("hanging_indent_chars")
    indent_chars = rl.get("indent_chars")
    if hang:
        spec["indent"] = (("hanging", float(hang)),
                          _ref(rules, "citation.reference_list.hanging_indent_chars",
                              "citation.reference_list"), "", FIX_AUTO)
    elif indent_chars:
        spec["indent"] = (("first_line", float(indent_chars)),
                          _ref(rules, "citation.reference_list.indent_chars",
                              "citation.reference_list"), "", FIX_AUTO)
    else:
        spec["indent"] = (("none", 0.0), _ref(rules, "citation.reference_list", "citation"),
                          "（本刊未要求条目缩进，排版时会清除）", FIX_AUTO)
    return spec


def _caption_spec(rules):
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    tables = rules.get("tables") or {}
    size = tables["caption_size_pt"] if "caption_size_pt" in tables else body.get("size_pt")
    font = tables["caption_font_cn"] if "caption_font_cn" in tables \
        else DEFAULTS["caption_font_cn"][0]
    mult, pt = _body_line(rules)
    line = {"mult": mult} if mult else {"pt": pt}
    spec = {}
    if font:
        spec["cn"] = (font, _ref(rules, "tables.caption_font_cn~", "tables.caption_size_pt~",
                                "tables"),
                      "（本刊未声明题注字体，取代码缺省值）" if "caption_font_cn" not in tables else "")
    if size:
        spec["size"] = (size, _ref(rules, "tables.caption_size_pt~", "tables"), "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    spec["align"] = ("center", _ref(rules, "tables.caption_align~", "tables"), "", FIX_AUTO)
    spec["line"] = (line, _line_ref(rules, "typography.body", mult, pt, False), "", FIX_AUTO)
    spec["bold"] = (False, _ref(rules, "tables.caption_font_cn~", "tables"), "", FIX_AUTO)
    spec["indent"] = (("none", 0.0), _ref(rules, "tables"), "", FIX_AUTO)
    return spec


def _table_note_spec(rules):
    tables = rules.get("tables") or {}
    typo = rules.get("typography") or {}
    body = typo.get("body") or {}
    layers = tables.get("note_layers") or [{}]
    first = layers[0] if isinstance(layers[0], dict) else {}
    size = first["size_pt"] if "size_pt" in first else DEFAULTS["note_size_pt"][0]
    mult, pt = _body_line(rules)
    line = {"mult": mult} if mult else {"pt": pt}
    spec = {}
    if body.get("font_cn"):
        spec["cn"] = (body.get("font_cn"),
                      _ref(rules, "typography.body.font_cn", "typography.body"), "")
    if size:
        spec["size"] = (size, _ref(rules, "tables.note_layers", "tables"),
                        "（本刊未声明表注字号，取代码缺省值）" if "size_pt" not in first else "")
    spec["latin"] = (body.get("font_latin") or "Times New Roman",
                     _ref(rules, "typography.body.font_latin~", "typography.body"),
                     "" if body.get("font_latin") else "（本刊未声明西文字体，按 Times New Roman）")
    spec["align"] = ("left", _ref(rules, "tables.note_layers", "tables"), "", FIX_AUTO)
    spec["line"] = (line, _line_ref(rules, "typography.body", mult, pt, False), "", FIX_AUTO)
    spec["indent"] = (("none", 0.0), _ref(rules, "tables.note_layers", "tables"), "", FIX_AUTO)
    return spec


def _figure_spec(rules):
    mult, pt = _body_line(rules)
    spec = {
        "align": ("center", _ref(rules, "figures"), "", FIX_AUTO),
        "line": ({"mult": 1.0}, _ref(rules, "figures"), "", FIX_AUTO),
        "indent": (("none", 0.0), _ref(rules, "figures"), "", FIX_AUTO),
    }
    return spec


# ---------------------------------------------------------------- 表格

def _cb_shape(cb):
    """单元格边框只比形状（哪条边、线型、线宽）；space/color 的变化不判偏差，
    避免「黑线但没写 color」这种视觉无差的稿子被满屏标红。"""
    return tuple(sorted((name, val, sz) for name, val, sz, _space, _color in cb))


def _border_desc(snap):
    borders = snap.get("borders") or {}
    parts = []
    for edge, zh in (("top", "顶线"), ("bottom", "底线")):
        item = borders.get(edge)
        if not item or item[0] in (None, "none", "nil"):
            parts.append(f"{zh}无")
        else:
            try:
                parts.append(f"{zh} {int(item[1]) / 8:g} 磅")
            except (TypeError, ValueError):
                parts.append(f"{zh}线型 {item[0]}")
    extra = [e for e in ("left", "right", "insideH", "insideV")
             if borders.get(e) and borders[e][0] not in (None, "none", "nil")]
    if extra:
        parts.append(f"{len(extra)} 条左右/内部线")
    cell_borders = sum(1 for cell in snap.get("cells") or [] if cell[0])
    if cell_borders:
        parts.append(f"{cell_borders} 处单元格边框")
    return "，".join(parts) or "无边框"


def _expected_table_snapshot(table, bp, size, cn, latin):
    """在 XML 深拷贝上跑一遍 apply_table，返回渲染器眼中的目标快照。

    这是「体检与渲染器共用同一段代码」的关键：expected 不是另写一份三线表
    规则，而是让渲染器自己算。拷贝只在内存里，源文档一个字节都不动。
    """
    clone = copy.deepcopy(table._tbl)
    wrapper = Table(clone, table._parent)
    apply_table(wrapper, bp, size, cn, latin)
    return table_snapshot(wrapper)


def _table_runs(table):
    """表内 run 元素，范围与 apply_table 一致（每个 tc 的直接 run）。"""
    out = []
    for tc in table._tbl.iter(qn("w:tc")):
        for p_el in tc.findall(qn("w:p")):
            out.extend(r._element for r in Paragraph(p_el, table).runs)
    return out


def _check_table(rep, blk, table, rules, body):
    idx = blk["idx"]
    tables = rules.get("tables") or {}
    bp = tables.get("border_pt")
    if not isinstance(bp, dict) or not all(k in bp for k in ("top", "bottom", "header_bottom")):
        rep.skip(idx, "table", "规则集未声明 tables.border_pt，无法判定三线表线宽")
        return
    size = tables["font_size_pt"] if "font_size_pt" in tables \
        else DEFAULTS["table_font_size_pt"][0]
    cn = body.get("font_cn")
    latin = body.get("font_latin") or "Times New Roman"

    try:
        actual = table_snapshot(table)
        expected = _expected_table_snapshot(table, bp, size, cn, latin)
    except Exception as exc:  # noqa: BLE001 —— 单表结构异常不该拖垮整份报告
        rep.skip(idx, "table", f"表格结构无法套版核对：{type(exc).__name__}: {exc}")
        return

    act_b = {k: v[:2] for k, v in (actual.get("borders") or {}).items()}
    exp_b = {k: v[:2] for k, v in (expected.get("borders") or {}).items()}
    act_cb = [_cb_shape(cell[0]) for cell in actual.get("cells") or []]
    exp_cb = [_cb_shape(cell[0]) for cell in expected.get("cells") or []]
    target = (f"顶/底 {bp['top']:g}/{bp['bottom']:g} 磅，"
              f"栏目线 {bp['header_bottom']:g} 磅，左右与内部竖线无线")
    if act_b == exp_b and act_cb == exp_cb:
        rep.add(idx, "表格", "table", "三线表边框", "tables.border_pt",
                target, "与本刊三线表一致", FIX_OK, "")
    else:
        rep.add(idx, "表格", "table", "三线表边框", "tables.border_pt",
                target, _border_desc(actual), FIX_AUTO,
                "表格边框会重设为三线表（左右与内部竖线一并清除）")

    # 表内文字：与 apply_table 同一范围（直接 run），字号/中西文字体逐项对账
    runs = [run_style(r) for r in _table_runs(table)]
    if not runs:
        return
    sizes = {r[2] / 2 for r in runs if r[2] is not None}
    partial_size = any(r[2] is None for r in runs)
    if size:
        target_size = round(size * 2) / 2
        _run_item(rep, idx, "表格", "table", "表内字号",
                  _ref(rules, "tables.font_size_pt~", "tables"), target_size,
                  f"{target_size:g}pt", sizes, partial_size, lambda v: f"{v:g}pt",
                  note_bad="表内文字字号与本刊要求不符",
                  inherited_note="表内 run 未显式设置字号，Word 按样式渲染；"
                                 "排版时会按本刊要求显式写入")
    if cn:
        values = {r[0] for r in runs if r[0] is not None}
        _run_item(rep, idx, "表格", "table", "表内中文字体",
                  _ref(rules, "typography.body.font_cn", "typography.body"), cn, cn,
                  values, any(r[0] is None for r in runs), str,
                  note_bad="表内中文字体与本刊要求不符")
    if latin:
        values = {r[1] for r in runs if r[1] is not None}
        _run_item(rep, idx, "表格", "table", "表内西文字体",
                  _ref(rules, "typography.body.font_latin~", "typography.body"), latin, latin,
                  values, any(r[1] is None for r in runs), str,
                  note_bad="表内西文字体与本刊要求不符")


# ---------------------------------------------------------------- 版面/脚注

def _check_page(rep, doc, rules):
    page = rules.get("page") or {}
    size_mm = page.get("size_mm") or {}
    margins = page.get("margin_mm") or {}
    secs = list(doc.sections)
    tw, th = size_mm.get("width"), size_mm.get("height")

    if tw and th:
        bad = []
        for i, sec in enumerate(secs, 1):
            w, h = sec.page_width, sec.page_height
            if (w is None or h is None or abs(w.mm - tw) > 0.05 or abs(h.mm - th) > 0.05):
                actual = ("未设置" if w is None or h is None
                          else f"{w.mm:.1f}×{h.mm:.1f} mm")
                bad.append(f"第 {i} 节 {actual}")
        expected = f"{tw:g}×{th:g} mm"
        actual = bad[0] if len(bad) == 1 else "；".join(bad)
        rep.add(-1, "页面", "page", "纸张", _ref(rules, "page.size_mm", "page"),
                expected, expected if not bad else actual,
                FIX_OK if not bad else FIX_AUTO,
                "" if not bad else f"排版时会按本刊纸张设置全部 {len(secs)} 节")

    for side in ("top", "bottom", "left", "right"):
        want = margins.get(side)
        if want is None:
            continue
        bad = []
        for i, sec in enumerate(secs, 1):
            cur = getattr(sec, f"{side}_margin")
            if cur is None or abs(cur.mm - want) > 0.05:
                bad.append(f"第 {i} 节 {'未设置' if cur is None else f'{cur.mm:.1f} mm'}")
        expected = f"{want:g} mm"
        actual = bad[0] if len(bad) == 1 else "；".join(bad)
        rep.add(-1, "页面", "page", SIDE_DESC[side],
                _ref(rules, f"page.margin_mm.{side}", "page.margin_mm"),
                expected, expected if not bad else actual,
                FIX_OK if not bad else FIX_AUTO,
                "" if not bad else f"排版时会按本刊页边距设置全部 {len(secs)} 节")


def _check_footnotes(rep, doc, rules):
    typo = rules.get("typography") or {}
    tp = typo.get("footnote") or {}
    size, cn = tp.get("size_pt"), tp.get("font_cn")
    latin = tp.get("font_latin") or "Times New Roman"
    if not (size or cn):
        return
    runs = []
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/footnotes"):
            continue
        if rel.is_external:
            rep.skip(-2, "footnote",
                     "外部脚注关系（TargetMode=External）没有可解析部件，"
                     "套版与体检都够不着，已原样保留")
            continue
        try:
            # 走项目统一的解析入口（拒 DOCTYPE/ENTITY、不解析外部实体），
            # 与 style_footnotes 同一姿势；裸 lxml/ET 入口在这里不出现。
            root = xmlsafe.fromstring(rel.target_part.blob)
        except Exception as exc:  # noqa: BLE001 —— 脚注部件坏了不该拖垮正文报告
            rep.skip(-2, "footnote", f"脚注部件无法解析：{type(exc).__name__}: {exc}")
            continue
        for fn in root.findall(qn("w:footnote")):
            # 分隔符/续行符没有正文；判据与 style_footnotes 同口径
            if (fn.get(qn("w:type"))
                    or fn.find(".//" + qn("w:separator")) is not None
                    or fn.find(".//" + qn("w:continuationSeparator")) is not None):
                continue
            runs.extend(fn.iter(qn("w:r")))
    if not runs:
        return
    styles = [run_style(r) for r in runs]
    if cn:
        values = {r[0] for r in styles if r[0] is not None}
        _run_item(rep, -2, "脚注", "footnote", "中文字体",
                  _ref(rules, "typography.footnote.font_cn", "typography.footnote"),
                  cn, cn, values, any(r[0] is None for r in styles), str,
                  note_bad="脚注中文字体与本刊要求不符")
    if size:
        target = round(size * 2) / 2
        values = {r[2] / 2 for r in styles if r[2] is not None}
        _run_item(rep, -2, "脚注", "footnote", "字号",
                  _ref(rules, "typography.footnote.size_pt", "typography.footnote"),
                  target, f"{target:g}pt", values,
                  any(r[2] is None for r in styles), lambda v: f"{v:g}pt",
                  note_bad="脚注字号与本刊要求不符")
    if latin:
        values = {r[1] for r in styles if r[1] is not None}
        _run_item(rep, -2, "脚注", "footnote", "西文字体",
                  _ref(rules, "typography.footnote.font_latin~", "typography.footnote"),
                  latin, latin, values, any(r[1] is None for r in styles), str,
                  note_bad="脚注西文字体与本刊要求不符")


# ---------------------------------------------------------------- 不处理清单

# ---------------------------------------------------------------- 脚注使用方式

# 手敲脚注的常见开头：①②③ 起头，或「注：①」「注1：」。Word 真脚注的
# 编号是 footnoteReference 对象，不是这些字符——认出来只提示「需你处理」，
# 不替用户改成真脚注（那会动正文结构）。
FAKE_FOOTNOTE = re.compile(r"^(?:[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|注\s*\d*\s*[：:])")
_ENCLOSED = "①②③④⑤⑥⑦⑧⑨⑩"


def _footnote_fmt_name(text):
    """规则集 numbering.footnotes 的写法 → Word numFmt 值。"""
    return "decimalEnclosedCircle" if any(c in (text or "") for c in _ENCLOSED) \
        else "decimal"


def _footnote_settings(doc):
    """读 Word 文档的脚注编号格式与重编号范围（未设置时按 Word 默认）。"""
    fmt, restart = "decimal", "continuous"
    try:
        settings = doc.settings.element
    except Exception:  # noqa: BLE001 —— 设置部件坏了不该拖垮整份体检
        return fmt, restart
    pr = settings.find(qn("w:footnotePr"))
    if pr is None:
        return fmt, restart
    nf = pr.find(qn("w:numFmt"))
    if nf is not None and nf.get(qn("w:val")):
        fmt = nf.get(qn("w:val"))
    nr = pr.find(qn("w:numRestart"))
    if nr is not None and nr.get(qn("w:val")):
        restart = nr.get(qn("w:val"))
    return fmt, restart


def _check_footnote_usage(rep, doc, rules, blocks):
    """脚注/尾注的只读体检（DESIGN §12.1 的已知缺口）。

    本工具只改脚注字体字号，不搬位置、不重编号。那几件事仍然不做，但
    「稿件用了尾注，目标刊要页脚脚注」「Word 设的是全文连续编号，本刊要
    每页重排」「脚注编号是 1 而本刊要 ①②③」这些偏差必须让作者看见——
    报告归入「需你处理」，不伪装成已检查或已修正。
    """
    num = rules.get("numbering") or {}
    want_fmt_txt = num.get("footnotes")
    want_scope = num.get("footnotes_scope")
    body = doc.element.body
    fn_refs = len(list(body.iter(qn("w:footnoteReference"))))
    en_refs = len(list(body.iter(qn("w:endnoteReference"))))
    rule = _ref(rules, "numbering.footnotes_scope", "numbering.footnotes",
                "typography.footnote", "journal")

    if en_refs:
        rep.add(-2, "脚注", "endnote", "位置", rule,
                "页脚脚注", f"尾注 {en_refs} 处",
                FIX_MANUAL,
                note="稿件把注释放在文末尾注，本刊要求页脚脚注。本工具不搬动"
                     "脚注/尾注位置，请在 Word 的「引用→脚注」里转换后再上传；"
                     "字体字号套版仍会应用到脚注部件")

    if fn_refs:
        fmt, restart = _footnote_settings(doc)
        if want_fmt_txt:
            want_fmt = _footnote_fmt_name(want_fmt_txt)
            if fmt != want_fmt:
                rep.add(-2, "脚注", "footnote", "编号样式",
                        _ref(rules, "numbering.footnotes", "journal"),
                        want_fmt_txt,
                        "①②③（圈号）" if fmt == "decimalEnclosedCircle"
                        else "1.2.3.（阿拉伯数字）",
                        FIX_MANUAL,
                        note="Word 的脚注编号格式与本刊要求不一致。本工具只调"
                             "字体字号，不改编号格式——请在 Word 里手工设置")
        if want_scope:
            want_restart = "eachPage" if want_scope == "per_page" else "continuous"
            if restart != want_restart:
                rep.add(-2, "脚注", "footnote", "编号范围",
                        _ref(rules, "numbering.footnotes_scope", "journal"),
                        "每页重新编号" if want_restart == "eachPage" else "全文连续编号",
                        "每页重新编号" if restart == "eachPage" else "全文连续编号",
                        FIX_MANUAL,
                        note="Word 的脚注重编号范围与本刊要求不一致。本工具不重"
                             "编号——请在 Word 的脚注设置里手工改")

    # 手敲的「①②③」段落：形似脚注但根本不是 Word 脚注对象，套版碰不到它，
    # 投稿系统也不会把它当脚注处理。
    fakes = [(b.get("idx"), b.get("type") or "body", (b.get("text") or "").strip())
             for b in blocks
             if b.get("type") in ("body", "front_matter", "table_note")
             and FAKE_FOOTNOTE.match((b.get("text") or "").strip())]
    if fakes:
        frag = "；".join(t[:16] + ("…" if len(t) > 16 else "") for _i, _t, t in fakes[:3])
        rep.add(fakes[0][0], "脚注", fakes[0][1], "手工脚注", rule,
                "Word 脚注功能", f"{len(fakes)} 段以圈号/「注：」开头",
                FIX_MANUAL,
                note=f"这些段落看起来是手敲的脚注（{frag}），不是 Word 脚注对象："
                     f"版式套版不会把它排成页脚，投稿系统也不识别。需要真脚注请在"
                     f"Word 里用「引用→插入脚注」重建")


def _not_handled(rules):
    """规则集声明了、渲染器不消费的项。逐条写明「本工具不修」，不给假承诺。"""
    typo = rules.get("typography") or {}
    tables = rules.get("tables") or {}
    figures = rules.get("figures") or {}
    page = rules.get("page") or {}
    out = []

    def add(prop, *refs, note):
        out.append({"property": prop, "rule": _ref(rules, *refs), "note": note,
                    "fix": FIX_MANUAL})

    for owner, conf in (("tables", tables), ("figures", figures)):
        if "caption_position" in conf:
            add(f"{'表' if owner == 'tables' else '图'}题位置",
                f"{owner}.caption_position",
                note="渲染器按原段落顺序排，不搬动题注与表/图的先后位置")
        if "caption_align" in conf:
            add(f"{'表' if owner == 'tables' else '图'}题对齐",
                f"{owner}.caption_align",
                note="渲染器对题注统一居中，规则集的 caption_align 不消费")
        if "caption_numbering" in conf:
            add(f"{'表' if owner == 'tables' else '图'}题编号体例",
                f"{owner}.caption_numbering",
                note="只归一题号与题名之间的间隔，不把 Table 1 / 表1 改写成别的编号体系")
    if "border" in figures:
        add("图边框", "figures.border",
            note=("规则集要求图加边框，但渲染器不设置图边框，需手工添加"
                  if figures.get("border") else
                  "规则集声明图不加边框，渲染器不会去除原图已有的边框"))
    if "repeat_header_on_break" in tables:
        add("跨页表头重复", "tables.repeat_header_on_break",
            note="渲染器不设置 Word 的 tblHeader 行属性；长表跨页时请手工勾选「重复标题行」")
    cols = page.get("columns")
    if isinstance(cols, int) and cols > 1:
        add("分栏", "page.columns",
            note=f"本刊为 {cols} 栏；渲染器只排单栏投稿稿，分栏由编辑部排版")
    add("页眉页脚与行号", "page",
        note="渲染器与体检都不处理页眉、页脚、页码与行号，原样保留")
    add("图分辨率", "figures.resolution_dpi_min" if "resolution_dpi_min" in figures else "figures",
        note="图内像素与分辨率不在 docx 版式层，渲染器只保留原图，不重采样")

    # 规则集里的 typography 样式键：渲染器 main 没有分支读取的，逐条点名。
    for key in sorted(typo):
        if key in CONSUMED_TYPO:
            continue
        add(f"样式 typography.{key}", f"typography.{key}",
            note="规则集声明了该样式，但渲染器没有分支读取它，排版不会套用")

    # 声明值与渲染器硬编码值冲突的几处。不点名的话，用户会以为排版后就是
    # 规则集写的样子。
    for which, label in (("title", "主标题"), ("subtitle", "副标题")):
        tp = typo.get(which) or {}
        if tp.get("align") and tp.get("align") != "center":
            add(f"{label}对齐", f"typography.{which}.align",
                note=f"规则集要求 {ALIGN_DESC.get(tp['align'], tp['align'])}，"
                     f"渲染器目前固定居中")
        if tp.get("bold") is True:
            add(f"{label}加粗", f"typography.{which}.bold",
                note="规则集要求加粗，渲染器目前固定不加粗")
        if tp.get("line_spacing"):
            add(f"{label}行距", f"typography.{which}.line_spacing",
                note="规则集声明了行距，渲染器目前按字号×1.5 折算，不读取该值")
    body = typo.get("body") or {}
    if body.get("align") and body.get("align") != "both":
        add("正文对齐", "typography.body.align",
            note=f"规则集要求 {ALIGN_DESC.get(body['align'], body['align'])}，"
                 f"渲染器目前固定两端对齐")
    if tables.get("font_cn") and tables.get("font_cn") != body.get("font_cn"):
        add("表内字体", "tables.font_cn",
            note="渲染器表内文字统一用正文字体，tables.font_cn 不消费")
    notes = tables.get("note_layers") or [{}]
    first = notes[0] if isinstance(notes[0], dict) else {}
    if first.get("font_cn") and first.get("font_cn") != body.get("font_cn"):
        add("表注字体", "tables.note_layers",
            note="渲染器表注文字统一用正文字体，note_layers[0].font_cn 不消费")
    # 图注：渲染器对表题/图题统一按 tables.caption_* 取值。本刊若给图注单列了
    # 不同的字号/字体，那个值不会被用上。
    if "caption_size_pt" in figures and figures.get("caption_size_pt") != tables.get("caption_size_pt"):
        add("图注字号", "figures.caption_size_pt",
            note="渲染器对表题/图题统一取 tables.caption_size_pt，"
                 "figures.caption_size_pt 不单独消费")
    if "caption_font_cn" in figures and figures.get("caption_font_cn") != tables.get("caption_font_cn"):
        add("图注字体", "figures.caption_font_cn",
            note="渲染器对表题/图题统一取 tables.caption_font_cn，"
                 "figures.caption_font_cn 不单独消费")
    return out


# ---------------------------------------------------------------- 主流程

def analyze(docx_path, rules):
    """对一份 docx 做只读版式体检，返回报告字典。"""
    path = Path(docx_path)
    doc = Document(str(path))
    # ai_cfg 显式关掉：体检不该读本地 AI 配置、更不该联网；解析结果只用来
    # 认块类型，与渲染器的 IR 同源。
    ir = parse_docx.parse(str(path), ai_cfg={"enabled": False})
    blocks = ir.get("blocks") or []

    elems = {}
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            elems[i] = Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            elems[i] = Table(child, doc)
        elif child.tag == qn("w:sdt"):
            # 与 render_from_ir.main 同一套下钻：sdt 槽位里挂段落/表格列表，
            # 由 elem_of 按 sdt_pos 取。两边不一致就会取错块。
            sdt_els, _unknown = _sdt_blocks(child.find(qn("w:sdtContent")))
            inner = [Paragraph(e, doc) if e.tag == qn("w:p") else Table(e, doc)
                     for e in sdt_els]
            if inner:
                elems[i] = inner

    rep = Report()
    body = (rules.get("typography") or {}).get("body") or {}
    hd = rules.get("headings") or {}
    note_label = next((b.get("label") for b in rules.get("blocks") or []
                       if b.get("type") == "author_note" and b.get("label")), None) \
        or "作者简介"
    fm_index = 0

    _check_page(rep, doc, rules)
    _check_footnote_usage(rep, doc, rules, blocks)

    for blk in blocks:
        el = elem_of(blk, elems)
        if el is None:
            rep.skip(blk.get("idx"), blk.get("type", ""),
                     "取不到对应元素（IR 的块序号与文档不配套），未检查")
            continue
        btype = blk.get("type") or ""
        idx = blk.get("idx")

        if btype == "equation_pseudo_table":
            rep.skip(idx, btype, "伪表格公式（1 行表格排的公式与编号），不拆不判")
            continue
        if btype == "equation":
            rep.skip(idx, btype, "OMML 原生公式对象，版式由公式编辑器掌管，未检查")
            continue
        if text_locked(blk):
            kinds = "/".join(sorted({f.get("kind", "") for f in blk.get("fields") or []
                                     if f.get("kind") in ("citation", "seq", "toc")}))
            rep.skip(idx, btype,
                     f"含 {kinds} 类 Word 域，文字归域所有；本体检不据此判偏差"
                     f"（排版只调版式、不改域文字）")
            continue

        if isinstance(el, Table):
            if btype != "table":
                rep.skip(idx, btype, "块元素是表格但类型不是 table，未按表格检查")
                continue
            _check_table(rep, blk, el, rules, body)
            continue

        p = el
        if btype == "table":
            rep.skip(idx, btype, "表格块取到的不是表格元素，未检查")
            continue

        if btype == "front_matter":
            fm_index += 1
            txt = block_text(p).strip()
            if fm_index == 1:
                spec = _title_spec(rules, "title")
            elif fm_index == 2 and txt.startswith(("——", "—", "-")):
                spec = _title_spec(rules, "subtitle")
            else:
                spec = _authors_spec(rules)
            _check_paragraph(rep, blk, p, "前置", spec)
            continue

        if btype == "figure":
            _check_paragraph(rep, blk, p, "图注", _figure_spec(rules))
            continue

        if btype == "references_heading":
            typo = rules.get("typography") or {}
            rh = typo.get("references_heading") or {}
            lv1 = hd.get("level1") or {}
            size = rh.get("size_pt") or lv1.get("size_pt") or body.get("size_pt")
            cn = rh.get("font_cn") or lv1.get("font_cn") or body.get("font_cn")
            align = rh.get("align") or lv1.get("align") or "center"
            mult, pt = _body_line(rules)
            if mult:
                line = {"mult": mult}
            else:
                line = {"pt": head_line_pt(lv1, size, pt)}
            rule = _ref(rules, "typography.references_heading", "headings.level1")
            spec = {}
            if cn:
                spec["cn"] = (cn, _ref(rules, "typography.references_heading.font_cn",
                                      "headings.level1.font_cn"), "")
            if size:
                spec["size"] = (size, _ref(rules, "typography.references_heading.size_pt",
                                          "headings.level1.size_pt"), "")
            spec["latin"] = (body.get("font_latin") or "Times New Roman",
                             _ref(rules, "typography.body.font_latin~", "typography.body"),
                             "" if body.get("font_latin") else
                             "（本刊未声明西文字体，按 Times New Roman）")
            spec["align"] = (align, _ref(rules, "typography.references_heading.align",
                                        "headings.level1.align"), "", FIX_AUTO)
            spec["line"] = (line, _line_ref(rules, "typography.body", mult, pt, False), "",
                            FIX_AUTO)
            spec["bold"] = (False, rule, "", FIX_AUTO)
            spec["indent"] = (("none", 0.0), rule, "", FIX_AUTO)
            _check_paragraph(rep, blk, p, "参考文献", spec)
            continue

        if btype in SEMANTIC:
            _check_paragraph(rep, blk, p, "前置", _semantic_spec(rules, btype))
            continue

        if btype in ("heading1", "heading2", "heading3", "heading4"):
            level = blk.get("level") or int(btype[-1])
            max_level = hd.get("max_level") or len(hd.get("levels") or [])
            if level > max_level:
                # 该刊只排到 X 级：渲染器会把这一段按正文套版，编号不重写。
                # 体检按正文样式核，另出一条「层级超出」的说明，不假装按
                # headings.level4 检查过。
                rep.add(idx, "标题", btype, "标题层级",
                        _ref(rules, "headings.max_level", "headings"),
                        f"该刊只排到 {max_level} 级",
                        f"稿件用了 {level} 级标题", FIX_MANUAL,
                        "四级标题会按正文套版，编号不重写；需要四级层级请换刊或手工处理")
                _check_paragraph(rep, blk, p, "正文", _body_spec(rules))
                continue
            spec = _heading_spec(rules, level)
            _check_paragraph(rep, blk, p, "标题", spec)
            _check_prefix(rep, blk, p, "标题", level, hd, FIX_AUTO)
            continue

        if btype == "title":
            # 外部直接构造 IR 时的防御分支（parse 不产出 title），与渲染器一致
            _check_paragraph(rep, blk, p, "前置", _title_spec(rules, "title"))
            continue

        if btype in ("abstract", "keywords"):
            _check_paragraph(rep, blk, p, "前置", _abstract_spec(rules, btype))
            continue

        if btype == "author_note":
            _check_paragraph(rep, blk, p, "参考文献",
                             _author_note_spec(rules, p, note_label))
            continue

        if btype == "reference_item":
            _check_paragraph(rep, blk, p, "参考文献", _reference_spec(rules))
            continue

        if btype in ("table_caption", "figure_caption"):
            _check_paragraph(rep, blk, p, "表格" if btype == "table_caption" else "图注",
                             _caption_spec(rules))
            continue

        if btype == "table_note":
            _check_paragraph(rep, blk, p, "表格", _table_note_spec(rules))
            continue

        # 其余（body / appendix_heading / 未识别类型）：渲染器最后都会落到
        # 正文套版，体检按同一口径判，避免「体检说没事、排版却改了」。
        _check_paragraph(rep, blk, p, "正文", _body_spec(rules))

    _check_footnotes(rep, doc, rules)

    # 解析端记下的未识别内容控件结构也要进「未检查」：它不是 blocks 里的块，
    # 不列出来就等于静默跳过（与 render_from_ir 的 unsupported_content_blocks
    # 提示同一口径）。
    if ir.get("unsupported_content_blocks"):
        rep.unchecked.append({"idx": None, "block": "sdt",
                              "reason": f"内容控件内另有 {ir['unsupported_content_blocks']} "
                                        f"个未识别结构，已原样保留"})

    items = _fold(rep.raw)
    stats = Counter(item["fix"] for item in rep.raw)
    stats["unchecked"] = len(rep.unchecked)
    for key in (FIX_AUTO, FIX_MANUAL, FIX_OK, FIX_INHERITED, "unchecked"):
        stats.setdefault(key, 0)

    block_total = len(blocks)
    # 页面/脚注条目的 idx 是负数，不是块号；数进覆盖率会让「核对块数」超过
    # 总块数（如 34/33），报告自己就不可信了。
    checked_ids = {item["idx"] for item in rep.raw
                   if isinstance(item["idx"], int) and item["idx"] >= 0}
    unchecked_ids = {u["idx"] for u in rep.unchecked
                     if isinstance(u.get("idx"), int) and u["idx"] >= 0}
    blocks_checked = len(checked_ids)
    coverage = {
        "blocks_total": block_total,
        "blocks_checked": blocks_checked,
        "blocks_unchecked": len(unchecked_ids),
        "checks_total": len(rep.raw),
        "rate": round(blocks_checked / block_total * 100, 1) if block_total else 100.0,
        "text": (f"逐块核对 {blocks_checked}/{block_total} 个块"
                 + (f"，{len(unchecked_ids)} 个块未检查" if unchecked_ids else "")),
    }

    journal = rules.get("journal") or {}
    return {
        "source": path.name,
        "journal": journal.get("name_zh") or journal.get("name_en") or "未知期刊",
        "journal_id": journal.get("id") or "",
        "stats": {k: stats[k] for k in (FIX_AUTO, FIX_MANUAL, FIX_OK,
                                        FIX_INHERITED, "unchecked")},
        "items": items,
        "not_handled": _not_handled(rules),
        "unchecked": rep.unchecked,
        "coverage": coverage,
    }


def _fold(raw):
    """偏差逐块列出；ok/inherited 按「块类型+属性+目标」聚合成一条带 count 的
    汇总。真实稿件里 90% 的检查都会落在这两类，逐块列会把偏差淹没，也会让
    preflight.json 膨胀到几百 KB。单条汇总不带 count，条目字段与约定一致。"""
    out = []
    groups = {}
    for row in raw:
        if row["fix"] not in (FIX_OK, FIX_INHERITED):
            out.append(dict(row))
            continue
        key = (row["section"], row["block"], row["property"], row["fix"],
               row["expected"], row["rule"])
        grp = groups.get(key)
        if grp is None:
            grp = dict(row)
            groups[key] = grp
            out.append(grp)
        else:
            grp["count"] = grp.get("count", 1) + 1
    for row in out:
        if row.get("count", 1) > 1 and row["fix"] in (FIX_OK, FIX_INHERITED):
            row["actual"] = (f"{row['count']} 处均一致" if row["fix"] == FIX_OK
                             else f"{row['count']} 处随样式（未显式设置）")
    return out


# ---------------------------------------------------------------- HTML

def _fix_label(fix):
    return {FIX_AUTO: "自动修正", FIX_MANUAL: "需你处理", FIX_OK: "已合规",
            FIX_INHERITED: "随样式"}.get(fix, fix)


def _idx_label(idx):
    if idx == -1:
        return "页面"
    if idx == -2:
        return "脚注"
    return str(idx)


def render_html(result):
    """把体检结果渲染成静态 HTML；样式与另两份报告共用 report_theme.BASE。"""
    st = result["stats"]
    cov = result["coverage"]
    auto = [it for it in result["items"] if it["fix"] == FIX_AUTO]
    manual = [it for it in result["items"] if it["fix"] == FIX_MANUAL]

    def rows(items):
        return "".join(
            f'<tr><td class="pos">{html.escape(_idx_label(it["idx"]))}'
            f'<span class="blk">{html.escape(it.get("block", ""))}</span></td>'
            f'<td>{html.escape(it["property"])}</td>'
            f'<td class="exp">{html.escape(str(it["expected"]))}</td>'
            f'<td class="act">{html.escape(str(it["actual"]))}</td>'
            f'<td class="note">{html.escape(it.get("note", ""))}'
            f'<code class="ref">{html.escape(it.get("rule", ""))}</code></td></tr>'
            for it in items)

    dev_html = ""
    if auto:
        dev_html += ('<section><h2>自动修正（排版时会按本刊要求改动）</h2>'
                     '<table><tr><th>位置</th><th>属性</th><th>本刊要求</th>'
                     '<th>输入稿现状</th><th>说明</th></tr>'
                     + rows(auto) + '</table></section>')
    if manual:
        dev_html += ('<section class="warn"><h2>需你处理（本工具不修）</h2>'
                     '<p class="meta">这些项渲染器没有对应分支，或规则集声明了但'
                     '渲染器不消费。排版不会替你改，请按「本刊要求」手工核对。</p>'
                     '<table><tr><th>位置</th><th>属性</th><th>本刊要求</th>'
                     '<th>输入稿现状</th><th>说明</th></tr>'
                     + rows(manual) + '</table></section>')
    if not (auto or manual):
        dev_html = ('<section><h2>偏差</h2><p class="empty">'
                    '没有发现需要修正的显式版式偏差。</p></section>')

    nh_html = ""
    if result["not_handled"]:
        nh_html = ('<section><h2>本工具不处理的项（不写成「已检查」）</h2>'
                   '<div class="glist">'
                   + "".join(
                       f'<div class="grow"><span class="gname">{html.escape(x["property"])}</span>'
                       f'<span class="gnote">{html.escape(x["note"])}</span>'
                       f'<code class="ref">{html.escape(x["rule"])}</code></div>'
                       for x in result["not_handled"])
                   + '</div></section>')

    uc_html = ""
    if result["unchecked"]:
        uc_html = ('<section><h2>未检查（说明原因，不假装看过）</h2>'
                   '<table><tr><th>位置</th><th>块</th><th>原因</th></tr>'
                   + "".join(
                       f'<tr><td class="pos">{html.escape(_idx_label(u.get("idx")))}</td>'
                       f'<td>{html.escape(u.get("block", ""))}</td>'
                       f'<td class="note">{html.escape(u.get("reason", ""))}</td></tr>'
                       for u in result["unchecked"])
                   + '</table></section>')

    # 合规/随样式按属性汇总，让用户看得出「哪些项真的核对过」
    agg = {}
    for it in result["items"]:
        if it["fix"] in (FIX_OK, FIX_INHERITED):
            key = (it["section"], it["property"], _fix_label(it["fix"]))
            agg[key] = agg.get(key, 0) + it.get("count", 1)
    agg_html = "".join(
        f'<tr><td>{html.escape(sec)}</td><td>{html.escape(prop)}</td>'
        f'<td>{html.escape(label)}</td><td class="num">{n}</td></tr>'
        for (sec, prop, label), n in sorted(agg.items()))

    verdict = ("输入稿没有需要自动修正的显式偏差。"
               if not auto else
               f"有 {len(auto)} 项会在排版时自动修正。")
    if manual:
        verdict += f"另有 {len(manual)} 项本工具不修，需手工核对。"

    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>版式体检 · {html.escape(result['source'])} → {html.escape(result['journal'])}</title>
<style>{BASE}
.kpi{{display:flex;margin:0 0 20px;background:var(--surface);border:1px solid var(--line);
border-radius:14px;overflow:hidden}}
.kpi>div{{flex:1;padding:15px 18px;border-left:1px solid var(--line2);min-width:0}}
.kpi>div:first-child{{border-left:none}}
.kpi b{{display:block;font-size:22px;font-weight:600;color:var(--dark);line-height:1.3;
font-variant-numeric:tabular-nums;letter-spacing:.5px}}
.kpi span{{display:block;font-size:12px;color:var(--ink3);margin-top:3px}}
.kpi .hl-auto b{{color:var(--accent)}} .kpi .hl-manual b{{color:var(--gold)}}
.kpi .hl-unchecked b{{color:var(--ink3)}}
.cov{{margin:9px 0 0;font-size:12.5px;color:var(--ink3);line-height:1.7}}
.meta{{color:var(--ink2);font-size:12.5px;margin:0 0 10px;line-height:1.8}}
.pos{{color:var(--ink3);font-variant-numeric:tabular-nums;white-space:nowrap;width:88px;
font-size:12.5px}}
.pos .blk{{display:block;font-size:10.5px;color:#8b9a94;margin-top:1px}}
.exp{{color:var(--accent);font-size:12.5px}}
.act{{color:var(--danger);font-size:12.5px}}
.note{{color:var(--ink2);font-size:12.5px;line-height:1.75}}
.ref{{display:block;margin-top:3px;font-size:11px;color:var(--ink3);letter-spacing:.2px}}
.num{{font-variant-numeric:tabular-nums;white-space:nowrap}}
.empty{{color:var(--ink3);font-size:13px;padding:4px 0}}
.glist{{font-size:13px}}
.grow{{padding:8px 0;border-top:1px solid var(--line2);line-height:1.75}}
.grow:first-child{{border-top:none}}
.gname{{font-weight:500;margin-right:9px}}
.gnote{{color:var(--ink2);font-size:12.5px}}
@media(max-width:680px){{
  .kpi{{flex-wrap:wrap}}.kpi>div{{flex:1 1 48%;border-left:none;border-top:1px solid var(--line2)}}
  .kpi>div:nth-child(-n+2){{border-top:none}}
}}
</style></head><body><div class="wrap">
<header>
  <h1>版式体检 · {html.escape(result['source'].replace('.docx', ''))}</h1>
  <p class="sub">目标期刊：<strong>{html.escape(result['journal'])}</strong>
  · 只读检查，未修改稿件任何字节</p>
  <p class="meta">{html.escape(cov['text'])}</p>
</header>
<div class="kpi">
  <div class="hl-auto"><b>{st[FIX_AUTO]}</b><span>自动修正（项）</span></div>
  <div class="hl-manual"><b>{st[FIX_MANUAL]}</b><span>需你处理（项）</span></div>
  <div><b>{st[FIX_OK]}</b><span>已合规（项）</span></div>
  <div class="hl-unchecked"><b>{st['unchecked']}</b><span>未检查（块）</span></div>
</div>
<p class="cov">{html.escape(verdict)}　随样式（未显式设置、无法凭直接格式断言）{st[FIX_INHERITED]} 项。</p>
{dev_html}
{nh_html}
{uc_html}
<section><h2>已核对且合规 / 随样式的项</h2>
<table><tr><th>分区</th><th>属性</th><th>结论</th><th>数量</th></tr>
{agg_html or '<tr><td colspan="4" class="empty">没有可汇总的项</td></tr>'}
</table></section>
<footer>
  本报告只读取版式，不修改稿件；偏差判定只针对「显式设置且与本刊要求不符」的项<br>
  未显式设置的字体字号记为「随样式」，排版时渲染器仍会按本刊要求显式写入
</footer>
</div></body></html>"""


# 与另外两份报告（compliance_report / render_audit）的 render 同名，
# app 侧统一按 `render(result)` 调用即可；详细名字保留给直接 import 的场景。
render = render_html


# ---------------------------------------------------------------- CLI

def main(argv=None):
    ap = argparse.ArgumentParser(description="输入稿版式体检（只读）")
    ap.add_argument("docx", help="待检稿件 .docx")
    ap.add_argument("ruleset", help="目标期刊规则集 .json")
    ap.add_argument("-o", "--out", required=True, help="输出 preflight.json")
    ap.add_argument("--html", help="另存 HTML 报告")
    a = ap.parse_args(argv)

    try:
        rules = json.loads(Path(a.ruleset).read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001 —— 坏文件点名报错，不甩 traceback
        print(f"失败：规则集 {a.ruleset} 读取失败：{type(exc).__name__}: {exc}",
              file=sys.stderr)
        return 1
    if not Path(a.docx).is_file():
        print(f"失败：找不到稿件 {a.docx}", file=sys.stderr)
        return 1

    try:
        result = analyze(a.docx, rules)
    except Exception as exc:  # noqa: BLE001 —— 用户拿到的应是结论，不是 traceback
        print(f"失败：体检未完成：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    out = Path(a.out)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已生成: {out}")
    if a.html:
        Path(a.html).write_text(render_html(result), encoding="utf-8")
        print(f"已生成: {a.html}")
    st = result["stats"]
    print(f"  自动修正 {st[FIX_AUTO]} · 需你处理 {st[FIX_MANUAL]} · "
          f"已合规 {st[FIX_OK]} · 随样式 {st[FIX_INHERITED]} · 未检查 {st['unchecked']}")
    print(f"  {result['coverage']['text']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
