"""把 Paper IR 渲染成目标期刊格式，并输出审计报告。

实现选择：就地重排（打开原稿 → 逐段套目标格式 → 另存），而不是新建文档重新拼装。
原因：原稿里的图片、公式对象、脚注都带关系 ID，跨文档复制必然断链，
重建一个文档会静默丢失这些内容。

红线：正文与表格内的文字零改动。只改版式、编号、标签，且全部进审计报告。

用法：python tools/render_from_ir.py <ir.json> <ruleset.json> -o out.docx
"""

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.table import Table
from lxml import etree
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_meta import sanitize_docx
# 版式写入原语与缺省值只有一份（tools/docx_fmt.py）。这里 import 后转出同名
# API：preflight 与既有测试仍从 render_from_ir 取 DEFAULTS/set_run/set_hanging，
# 换实现不换入口。set_hanging 是 docx_fmt.set_hanging_indent 的别名（旧名）。
from docx_fmt import (  # noqa: F401  （公开名，供 preflight/测试 import）
    ALIGN, DEFAULTS, set_chars_indent, set_hanging, set_hanging_indent, set_run,
)
import xmlsafe
import assume_groups

# 对齐方式的中文说法，只用于审计条目的说明文字，不是格式值
ALIGN_DESC = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}

# 规则集用中文描述题注编号与标题的间隔，这里映射成实际空格
CAPTION_GAP = {"一字空": "  ", "半字空": " ", "两字空": "    ", "无间隔": ""}

CN_SEQ = "一二三四五六七八九十"

# ---- F13：规则集必填键 ----------------------------------------------------
# 渲染器会裸下标消费的键全在这里列清单：入口一次报全，缺任何一个都在动稿
# 之前拦下，而不是渲染中途 KeyError 崩掉（崩溃输出还会被上层一律翻译成
# 「稿件结构超出识别范围」，把锅甩给稿件——失败且误诊）。
RULESET_ERR_PREFIX = "[RULESET-INCOMPLETE]"

SECTION_NAMES = {
    "journal": "期刊", "typography": "版式", "headings": "标题",
    "blocks": "块声明", "page": "页面", "tables": "表格",
    "figures": "图", "citation": "引用",
}

# 字符串项查「键存在」；元组项是深一层断言（F13DEEP）——批五只查键存在性，
# line_spacing={"mode":"multiple"} 缺 value 子键照样在渲染中途裸 KeyError，
# note_layers 条目缺 kind 也是 reorder_notes 里的裸下标。断言形态：
#   "路径"                             键存在即可（不限形态）
#   (路径, "dict")                     键存在且为 dict（渲染器要下钻消费它）
#   (路径, "num")                      键存在且为数值（渲染器要做算术与 :g 格式化）
#   (路径, "str")                      键存在且为字符串
#   (路径, "list")                     键存在且为 list
#   (路径, "keys_when", 条件, *子键)    是 dict，且条件全满足时必须含这些子键
#   (路径, "each_keys", *子键)          是 list，每项都是 dict 且含这些子键；
#                                      键缺失或不是 list 时按路径本身报缺
# 形态断言（F13STR）：该是 dict 的键给了字符串（typography.body.line_spacing
# = "1.5倍"）在「只查存在」的校验下放行，渲染中途才以 AttributeError 崩——
# 形态不符必须同样在动稿之前拦下，报「字段 X 形态不符：期望 Y」。
REQUIRED_RULE_KEYS = [
    ("journal.name_zh", "str"),
    ("typography.body.font_cn", "str"),
    ("typography.body.size_pt", "num"),
    ("typography.body.line_spacing", "dict"),
    # value 只在倍数行距声明下消费（固定值行距走 value_pt，可落到代码缺省），
    # 条件断言避免误杀 exact 模式的规则集；value:num 断言其是数值（给了
    # 字符串，渲染中途 f"{value:g}" 会裸崩）。存在性由 value:num 一并管，
    # 不再单列 value——两列会把「缺失」报成两条
    ("typography.body.line_spacing", "keys_when", {"mode": "multiple"}, "value:num"),
    ("typography.body.first_line_indent_chars", "num"),
    ("typography.title", "dict"),
    ("headings.level1", "dict"),
    # levels 只够「条目里每个编号模板各是字符串」还不够：少于 max_level 时
    # 高层标题在渲染中途 IndexError（[None]*4 只兜整条缺失）——条目形态由
    # each_keys 断言（含非 list），条数由 levels_len 断言，都在入口拦下。
    # max_level 未声明时 levels_len 只要求至少 1 项（level1 分支无守卫下标）
    ("headings.levels", "each_keys", "str"),
    ("headings.levels", "levels_len"),
    ("blocks", "each_keys", "type"),
    ("page.size", "str"),
    ("page.size_mm.width", "num"),
    ("page.size_mm.height", "num"),
    ("page.margin_mm.top", "num"),
    ("page.margin_mm.bottom", "num"),
    ("page.margin_mm.left", "num"),
    ("page.margin_mm.right", "num"),
    ("tables.border_pt.top", "num"),
    ("tables.border_pt.header_bottom", "num"),
    ("tables.border_pt.bottom", "num"),
    ("tables.note_layers", "each_keys", "kind"),
    ("figures.note_layers", "each_keys", "kind"),
    ("citation", "dict"),
]

# 形态判定与报错用名。num 排除 bool（True 是 int 的子类，但 size_pt=True
# 只能是规则集写错了）
_SHAPE_CHECK = {
    "dict": lambda v: isinstance(v, dict),
    "num": lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
    "str": lambda v: isinstance(v, str),
    "list": lambda v: isinstance(v, list),
}
_SHAPE_NAME = {"dict": "dict", "num": "数值", "str": "字符串", "list": "list"}


def _required_walk(rules, path):
    """沿点路径取节点；中途断掉或落成显式 null 返回 None。"""
    node = rules
    for key in path.split("."):
        if not isinstance(node, dict) or key not in node:
            return None
        node = node[key]
    return node


def missing_required_keys(rules):
    """返回规则集缺失的必填键（点路径列表），一次报全。

    元组项是深一层断言（形态见 REQUIRED_RULE_KEYS）：缺子键报成
    「路径.子键」，list 条目缺子键报成「路径[].子键」；子键名带 `:num`
    后缀时断言该子键是数值（渲染中途会当数字消费），形态不符报成
    「路径.子键 形态不符：期望 数值」。形态不符（该是 dict/数值/字符串
    的键给了别的类型，F13STR）报成「路径 形态不符：期望 X」——与存在性
    缺口在同一份清单里一次报全。
    """
    missing = []
    for item in REQUIRED_RULE_KEYS:
        if isinstance(item, str):
            path, kind, args = item, "exists", ()
        else:
            path, kind, args = item[0], item[1], item[2:]
        node = _required_walk(rules, path)
        if node is None:
            missing.append(path)
            continue
        if kind == "exists":
            continue
        if kind == "keys_when":
            cond, subkeys = args[0], args[1:]
            if isinstance(node, dict) and all(node.get(k) == v for k, v in cond.items()):
                for s in subkeys:
                    if s.endswith(":num"):
                        key, shape = s[:-4], "num"
                        if key not in node:
                            missing.append(f"{path}.{key}")
                        elif not _SHAPE_CHECK[shape](node[key]):
                            missing.append(f"{path}.{key} 形态不符：期望 {_SHAPE_NAME[shape]}")
                    elif s not in node:
                        missing.append(f"{path}.{s}")
        elif kind == "each_keys":
            if not isinstance(node, list):
                missing.append(f"{path} 形态不符：期望 {_SHAPE_NAME['list']}")
                continue
            for e in node:
                if isinstance(e, dict):
                    absent = tuple(s for s in args if s not in e)
                elif "str" in args:
                    absent = () if isinstance(e, str) else ("str",)
                else:
                    absent = args
                missing.extend(f"{path}[].{s}" for s in absent)
        elif kind == "levels_len":
            # levels 条数少于标题层级时高层标题会裸 IndexError（F13 拦渲染
            # 中途崩）：声明了 max_level 就要求 levels 覆盖到它；未声明时
            # 只要求至少 1 项（level1 分支同样无守卫地下标 levels[0]）。
            # 单层编号刊（levels 仅 1 项、未声明 max_level）是合法形态，不误杀。
            if not isinstance(node, list):
                continue          # 形态不符已由 "list" 断言报，这里不重复
            max_level = _required_walk(rules, "headings.max_level")
            need = max_level if isinstance(max_level, int) and not isinstance(max_level, bool) \
                else 1
            if len(node) < need:
                missing.append(f"headings.levels 条目不足：{len(node)} 项，"
                               f"需要 {need} 项（headings.max_level "
                               f"{'=' + str(max_level) if isinstance(max_level, int) and not isinstance(max_level, bool) else '未声明'}）")
        elif not _SHAPE_CHECK[kind](node):
            missing.append(f"{path} 形态不符：期望 {_SHAPE_NAME[kind]}")
    return missing


# F30 的缺省值表在 tools/docx_fmt.py：样板与交付稿必须共用同一源，
# 否则同一规则集会渲出两个字号/行距。这里只保留审计用的措辞函数。


def _fmt_default(v):
    return f"{v:g}" if isinstance(v, (int, float)) else f"「{v}」"


def presumed(keys):
    """缺省值标注：本条审计用到了哪些代码缺省，在 note 里自报家门（F30）。

    措辞是「缺省值」不是「推定值」：推定值专指规则集标了 measured:false 的
    取值（assumed() 负责），代码缺省不冒称推定（test_audit_traceable）。
    """
    if not keys:
        return ""
    why = "；".join(f"{DEFAULTS[k][1]} {_fmt_default(DEFAULTS[k][0])}" for k in keys)
    return f"，缺省值（{why}）"


def cn_numeral(n):
    """1→一、10→十、11→十一、21→二十一。

    `CN_SEQ` 只有十位，直接下标会在第 11 个一级标题上 IndexError——
    把整篇转换崩成"转换失败"。真实稿件十几个小节并不罕见。
    """
    if n <= 0:
        return str(n)
    d = "零一二三四五六七八九"
    if n < 10:
        return d[n]
    if n < 20:
        return "十" + (d[n % 10] if n % 10 else "")
    if n < 100:
        return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")
    return str(n)


def fill_level_template(tpl, n):
    """把编号模板里的序号换成第 n 个，中文数字与阿拉伯数字两种写法都认。

    模板是规则集 `headings.levels` 里的一项：「一、」「（一）」「1.」「（1）」。
    只按一种写法替换，另一种报刊的编号会全部冻结在模板原样——《经济管理》
    《南开管理评论》二级标题用「1.」，旧代码拿中文数字正则去替换，替不上，
    结果整篇二级标题都印成「1.」。
    """
    if re.search(r"[一二三四五六七八九十]", tpl):
        return re.sub(r"[一二三四五六七八九十]+", cn_numeral(n), tpl, count=1)
    return re.sub(r"\d+", str(n), tpl, count=1)


def tpl_prefix_rx(tpl):
    """把目标刊的编号模板变成「剥除已有前缀」的正则。

    `1.` → `^\\d{1,3}\\.(?!\\d)`；`（1）` → `^[（(]\\d{1,3}[）)](?!\\d)`；
    `一、` → `^[一二三四五六七八九十]{1,3}、`。

    为什么需要这条：NUM_PREFIX / NUM_PREFIX_NUM 覆盖的是「稿件原本可能是什么
    编号」，不含目标刊自己的写法。于是同一份成品重跑时——比如《经济管理》
    二级标题已经是「1.标题」——通用正则只认「1.1」不认「1.」，剥不掉，
    新编号叠加上去就成了「1.1.标题」。一次重跑叠一层，成品被彻底毁掉。
    渲染器写出的形式必须能被自己读回来，这是幂等的前提。
    """
    if not tpl:
        return None
    # 先把「数字」「中文数字」两类占位符抽出来，其余字符一律转义，
    # 免得模板里的括号、顿号被当成正则元字符
    marks = re.sub(r"[一二三四五六七八九十]+", "\x01", re.sub(r"\d+", "\x00", tpl))
    pat = re.escape(marks)
    pat = pat.replace(re.escape("\x00"), r"\d{1,3}")
    pat = pat.replace(re.escape("\x01"), r"[一二三四五六七八九十]{1,3}")
    # 中文刊里全角/半角括号混用很常见，模板写全角时也要认半角
    pat = pat.replace("（", "[（(]").replace("）", "[）)]")
    # (?!\d) 防「1.5亿元」这类数字开头的正文被剥成「5亿元」
    return re.compile("^" + pat + r"\s*(?!\d)")


def strip_number_prefix(core, level, hd):
    """剥掉标题里已有的编号前缀，返回（正文, 是否剥掉了）。

    候选顺序：通用中文编号 → 通用数字编号 → 目标刊自己的模板形状。
    前两条处理「稿件原本的体系」，第三条处理「本工具自己上一轮的输出」。
    """
    for rx in (NUM_PREFIX.get(level), NUM_PREFIX_NUM.get(level),
               tpl_prefix_rx((hd.get("levels") or [None] * 4)[level - 1])):
        if rx is None:
            continue
        stripped = rx.sub("", core).strip()
        if stripped != core:
            return stripped, True
    return core, False


NUM_PREFIX = {
    1: re.compile(r"^[一二三四五六七八九十]{1,3}、\s*"),
    2: re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]\s*"),
    # 三级的单段编号（「1.」）后不得再跟数字（(?!\d)）：三级标题写成「2.1.1 标题」
    # 时，裸的「N.」只会吃掉第一段编号、把「1.1 」残留在正文里，重编号后
    # 变成「2.1.1 1.1 标题」越跑越长（IDMPOT 双渲染实测）。链式形态由
    # NUM_PREFIX_NUM[3] 整段剥掉
    3: re.compile(r"^\d{1,2}\s*[.．]\s*(?!\d)"),
    4: re.compile(r"^[（(]\d{1,2}[）)]\s*"),
}
# 数字编号体系（1 / 1.1 / 1.1.1）的前缀，用于剥掉稿件原有的编号。
# 一级要同时认「1. 引言」和「1 引言」两种写法——后者是 Word 多级列表的
# 常见输出，漏了就会套成「一、1 引言」这种双编号。结尾的 (?!\d) 防止
# 「2.1」被一级规则吃掉前半个编号。
NUM_PREFIX_NUM = {
    1: re.compile(r"^\d{1,2}(?:\s*[.．]\s*|\s+)(?!\d)"),
    2: re.compile(r"^\d{1,2}\.\d{1,2}\s*"),
    3: re.compile(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\s*"),
    4: re.compile(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}\s*"),
}


# 用户补录后注入的语义块 → (规则集 typography 键, 对齐, 加粗)
# 只做"类型到语义"的映射，不含任何期刊判断；字号字体一律查规则集
SEMANTIC = {
    "affiliations": ("authors", "center", False),
    "english_authors": ("authors", "center", False),
    "english_affiliations": ("abstract", "center", False),
    "classification": ("abstract", "left", False),
    "english_title": ("title", "center", False),
    "english_abstract": ("abstract", "both", False),
    "english_keywords": ("abstract", "both", False),
}


def reject_sig(text):
    """驳回键用的块文本指纹（改前文本的短哈希）。

    键若用「块序号|动作」，补录注入（apply_details 重排 body 子元素）会让
    全部序号平移，上一轮保存的 reject.json 在注入后的下一轮会驳到错误的块
    ——用户要保留的原文改动照常执行且无提示。指纹与块在文档里的位置无关；
    各轮渲染都从原稿重出，同一块的改前文本不变，键才真正跨轮恒定。
    """
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12] if text else ""


class Audit:
    def __init__(self, rejects=None):
        self.items = []
        self.sec = "前置"
        # 驳回键 = 「块文本指纹|动作」（见 reject_sig）。指纹为空（取不到
        # 文本的块）退回「块序号|动作」；旧版保存的「块序号|动作」清单在
        # 序号未漂移时也继续被认（见 blocked），升级前的驳回决定不失效。
        self.rejects = set(rejects or [])
        # 逐块核对过的块数。报告要能回答「你没改的那些块，是看过还是漏了」——
        # 只列改动项，用户无从判断覆盖率；把核对总数报出来，才谈得上可审计。
        self.blocks_seen = 0

    @staticmethod
    def key(idx, action, sig=""):
        """审计条目 id／驳回键：有指纹用指纹，没有退回块序号。"""
        return f"{sig or idx}|{action}"

    def blocked(self, idx, action, sig=""):
        # 新键「块指纹|动作」与位置无关；旧键「块序号|动作」只在序号恰好
        # 没漂移时才命中——它本身不含文本信息，漂移场景无从迁移，认出
        # 多少算多少（好过整份作废）
        return (self.key(idx, action, sig) in self.rejects
                or f"{idx}|{action}" in self.rejects)

    def add(self, idx, btype, action, before="", after="", note="", rejected=False, rule="",
            changed=True, sig=""):
        """rule 是本条改动的依据，指向规则集里的具体字段（见 rules_ref.py）。

        结尾带 `~` 表示该键在部分规则集里可缺省，缺省时按代码内约定取值。

        changed=False 表示本条是「说明」而非「改动」（跳过、原样保留、域锁定）：
        它必须出现在报告里——「这里看过、但没动」和「这里漏了」是两件事——
        但不能计入改动数，否则报告自己就在虚报。

        sig 是驳回键的块文本指纹：可驳回动作的条目必须与 blocked() 用同一
        指纹，勾选框提交回来的 id 下一轮才命中同一块（见 reject_sig）。
        """
        self.items.append({"id": self.key(idx, action, sig), "idx": idx, "sec": self.sec,
                           "block": btype,
                           "action": action, "before": before, "after": after, "note": note,
                           "rule": rule, "rejected": bool(rejected), "changed": bool(changed)})


def fmt_snapshot(p):
    """取段落当前的版式状态，用于判断这次套版到底改没改东西。

    restyle 会动的每一项都要在快照里（字号、中西字体、加粗、字间距、行距值
    与规则、对齐、缩进、段前后距），否则「仅对齐变了」「仅缩进变了」的段落
    会被误判成没变化，审计漏记（F25；字间距同病，PB2）。
    run 范围与 restyle 同口径（段落自身的全部 run）：只看 p.runs 时，
    超链接里的 run 改了字号快照却看不见，审计会漏记这笔版式改动。
    hAnsi 也要进快照：set_run 写 ascii 时一并重写 hAnsi，不看它的话
    「仅 hAnsi 不同」的真改动就成了静默改动（同 table_snapshot 的做法）。
    """
    sizes, cns, latins, bolds, spacings, hansis = set(), set(), set(), set(), set(), set()
    for r in para_runs(p):
        cn_f, latin_f, sz, bold = run_style(r)
        if sz is not None:
            sizes.add(round(int(sz) / 2, 2))
        if cn_f:
            cns.add(cn_f)
        if latin_f:
            latins.add(latin_f)
        bolds.add(bold)
        hansi_f = run_hansi(r)
        if hansi_f:
            hansis.add(hansi_f)
        rpr = r.find(qn("w:rPr"))
        sp = rpr.find(qn("w:spacing")) if rpr is not None else None
        if sp is not None:
            # 原样存 twips 字符串（set_run 写的就是这个刻度），往返零误差
            spacings.add(sp.get(qn("w:val")))
    pf = p.paragraph_format
    ls = pf.line_spacing
    line = round(ls.pt, 2) if hasattr(ls, "pt") else (ls if isinstance(ls, (int, float)) else None)
    ppr = p._p.pPr
    ind_el = ppr.find(qn("w:ind")) if ppr is not None else None
    ind = tuple(sorted(ind_el.attrib.items())) if ind_el is not None else None
    sb, sa = pf.space_before, pf.space_after
    return {"sizes": sizes, "cn": cns, "latin": latins, "bold": bolds, "hansi": hansis,
            "spacing": spacings,
            "line": line, "rule": pf.line_spacing_rule, "jc": p.alignment,
            "ind": ind,
            "sb": round(sb.pt, 2) if sb is not None else None,
            "sa": round(sa.pt, 2) if sa is not None else None}


def fmt_delta(before, after, target):
    """把「改了什么」写成一句话。没变化就只报目标值。

    行距的「值」与「规则」（固定值／最小值／倍数）算同一件事：只换规则
    不换值也是行距改动，不能拆成两条两个名字。
    """
    names = {"sizes": "字号", "cn": "中文字体", "latin": "西文字体", "bold": "加粗",
             "spacing": "字间距", "hansi": "西文字体(hAnsi)",
             "jc": "对齐", "ind": "缩进", "sb": "段前距", "sa": "段后距"}
    changed = [label for key, label in sorted(names.items()) if before[key] != after[key]]
    if (before["line"], before["rule"]) != (after["line"], after["rule"]):
        changed.append("行距")
    return (f"改{'、'.join(changed)} → {target}" if changed else target)


def restyle_audited(p, audit, idx, btype, action, target, rule, tp=None, **kw):
    """套版并按前后快照出审计条目：真的变了才记，与正文套版同一口径。

    之前 keywords、表注、标题、作者简介、文献条目这些分支的 restyle 不出
    审计条目——改了字体字号报告里一条没有（漏报）。统一走这里后，变了才
    出条目，本来就近规的对象不再虚报，同刊重跑审计不会重出一批假改动。
    note 写目标值（改动名由 fmt_delta 给出），rule 指向本次取值的规则集字段。
    """
    before = fmt_snapshot(p)
    restyle(p, **kw)
    after = fmt_snapshot(p)
    if before != after:
        audit.add(idx, btype, action, note=fmt_delta(before, after, target) + assumed(tp),
                  rule=rule)


def assumed(tp):
    """规则集标了 measured: false 的属于推定值，报告里必须说明。

    用户有权知道这个字号是期刊明文写的，还是我们照着别的刊推出来的。
    """
    return "，推定值（样本未实测）" if isinstance(tp, dict) and tp.get("measured") is False else ""


def clear_para_formatting(p, keep_numbering=True):
    """清掉段落上的直接版式（行距、缩进、对齐）。

    编号定义 w:numPr 默认保留：它挂着 Word 的自动编号，是看得见的内容，
    删了编号就无声消失，正文分支没有任何补偿。唯一的例外是标题分支——
    那里的编号会被重写成字面文字，留着自动编号反而双重编号。
    """
    ppr = p._p.get_or_add_pPr()
    for tag in ("w:spacing", "w:ind", "w:jc"):
        el = ppr.find(qn(tag))
        if el is not None:
            ppr.remove(el)
    if not keep_numbering:
        el = ppr.find(qn("w:numPr"))
        if el is not None:
            ppr.remove(el)


def head_line_pt(lv, size, body_line):
    """标题行距：规则集可用 line_multiple 指定倍数（如一级标题两倍行距）。

    倍数只在大于正文字号时生效——否则「20 磅固定值 + 22pt 标题」会重叠。
    正文用倍数行距（body_line 为 None）时这里也返回 None，由 restyle 按
    倍数设置，标题就不会被单独钉成固定值。
    """
    mult = lv.get("line_multiple")
    if mult:
        return max(body_line or 0, round(size * mult)) or None
    return body_line


def restyle(p, *, cn, latin, size, align, line_pt=None, indent=None, bold=None,
            before_pt=None, after_pt=None, letter_spacing=None, line_rule="exact",
            line_multiple=None, drop_numbering=False):
    clear_para_formatting(p, keep_numbering=not drop_numbering)
    if align:
        p.alignment = ALIGN[align]
    if line_multiple:
        # 规则集写「1.5 倍行距」时必须真的写成倍数（lineRule=auto）。
        # 折算成固定磅值看着差不多，但编辑在 Word 里一眼就看出不是「1.5 倍」。
        p.paragraph_format.line_spacing = line_multiple
    elif line_pt:
        p.paragraph_format.line_spacing = Pt(line_pt)
        if line_rule == "at_least":
            # 「固定值」会裁掉超出行高的内容（分式、积分、上下标、内嵌图），
            # 「最小值」只保底不封顶，视觉上一样，但不会被裁
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    # 前置区的段间距：标题、作者、摘要之间靠它撑开，否则 26pt 标题会与下一行贴死
    if before_pt is not None:
        p.paragraph_format.space_before = Pt(before_pt)
    if after_pt is not None:
        p.paragraph_format.space_after = Pt(after_pt)
    if indent:
        set_chars_indent(p, indent)
    # p.runs 看不见超链接、修订插入（w:ins）里的 run：漏套的后果是链接内
    # 文字保留旧字号，成品里一眼就能看出这两种字号。所以按段落自身的全部
    # 后代 run 套（范围与 fmt_snapshot 同口径）
    for r in para_runs(p):
        set_run(Run(r, p), cn=cn, latin=latin, size=size, bold=bold,
                letter_spacing=letter_spacing)


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# ---- 段落文字的物理视图 -------------------------------------------------
# python-docx 1.2.0 的 p.text 只聚合「直接子级的 w:r 与 w:hyperlink」，
# p.runs 只取直接子级的 w:r：超链接里的文字前者看得见、后者改不到；
# 修订插入（w:ins）与域结果里的文字两者都看不见。体检、改写、审计若
# 各用各的口径，就会出现「闸门放行、改写空转」的漏改。下面这组函数把
# 段落自身的全部后代 run（不下钻文本框与已删除内容）拆成物理单元，
# 供体检、改写与审计共用同一份文本。

_SHAPE_GUARD = ("not(ancestor::w:drawing) and not(ancestor::w:pict)"
                " and not(ancestor::w:object) and not(ancestor::w:del)")


def para_runs(p):
    """段落自身的全部 w:r：含超链接、域结果、修订插入里的 run。"""
    return p._p.xpath(f".//w:r[{_SHAPE_GUARD}]")


def text_units(p):
    """段落文字的物理单元 [(元素, 贡献字符)]，顺序即阅读顺序。

    w:t 贡献原文；w:tab 与 w:ptab 贡献制表符；换行型 w:br/w:cr 贡献
    换行符（列/页分隔的 br 不产生文字）——与 p.text 同一套字符映射，
    但覆盖它看不见的角落。这是后文所有替换的字符计数口径。
    """
    units = []
    for r in para_runs(p):
        for child in r:
            tag = child.tag
            if tag == qn("w:t"):
                units.append((child, child.text or ""))
            elif tag in (qn("w:tab"), qn("w:ptab")):
                units.append((child, "\t"))
            elif tag == qn("w:cr"):
                units.append((child, "\n"))
            elif tag == qn("w:br"):
                kind = child.get(qn("w:type")) or "textWrapping"
                units.append((child, "\n" if kind == "textWrapping" else ""))
    return units


def para_full_text(p):
    """段落聚合文本：text_units 的字符拼接。体检与改写都以它为准。"""
    return "".join(chars for _el, chars in text_units(p))


def block_text(el):
    """块的聚合文本；表格等非段落元素返回空串。"""
    if getattr(el, "_p", None) is None:
        return ""
    return para_full_text(el)


# 文字归域所有、本工具不动的域类型：
#   citation —— EndNote/Zotero 等插入的引文与文献表，文字归文献库
#   seq      —— Word 自动编号（图表序号、章节号），改了就与编号指令脱钩
#   toc      —— 目录域，条目由 Word 生成
# page / ref 类不在此列：页码与交叉引用不参与本工具的文本改写分支。
TEXT_OWNED_KINDS = {"citation", "seq", "toc"}


def text_locked(blk):
    """该块的文字是否归 Word 域所有，本工具只改版式、不改文字。

    这些块里的可见文字是域的「结果」，不是作者敲的 run。改了不会报错，
    但后果比格式不对严重：作者返修时在 EndNote 里加一篇文献，按 F9 更新，
    工具改过的文字被域重新生成覆盖回去，而条目与文献库的对应关系已经断了。
    """
    return any(f.get("kind") in TEXT_OWNED_KINDS for f in blk.get("fields") or ())


def locked_note(audit, blk, action, rule=""):
    """含域块的「没改」说明：把边界写进报告，别让用户以为工具漏了。

    这是 N 档条目（未改动任何内容）：不计入改动数，但必须出现——
    「这里我看见了、但我不能动」和「这里我没看见」是两件事。
    """
    kinds = sorted({f["kind"] for f in blk.get("fields") or ()
                    if f.get("kind") in TEXT_OWNED_KINDS})
    audit.add(blk["idx"], blk["type"], "含域，文字未改",
              note=f"原计划「{action}」；本段文字是 {'/'.join(kinds)} 类 Word 域的"
                   f"显示结果，改动会与域指令脱钩（Word 里按 F9 更新即回滚），"
                   f"故保留原样。版式仍按本刊要求套用",
              rule=rule, changed=False)


def _write_t(node, text):
    """写 w:t 文字；带首尾空白时补 xml:space，否则 Word 渲染时会修掉它。"""
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set(XML_SPACE, "preserve")


def replace_text(p, new_text):
    """替换段落文字，但保留公式、脚注引用这类非文字对象。

    不能用「删光 run 再塞一个新 run」的做法：段落里若有 OMML 公式或脚注引用，
    会被连对象一起删掉，而且删得悄无声息——文字看着对，公式没了。
    这里只改 w:t 文字节点，其余节点原样留着。

    范围取段落自身的全部文字节点：整段文字都在超链接里时 p.runs 是空的，
    按 p.runs 找不到节点就会另起一个 run，段落文字翻倍。被替换区里的
    制表符／换行一并删除，成品才与审计记的 after 逐字一致。
    """
    nodes, drop = [], []
    for el, _chars in text_units(p):
        if el.tag == qn("w:t"):
            nodes.append(el)
        else:
            drop.append(el)
    if not nodes:
        return p.add_run(new_text)

    _write_t(nodes[0], new_text)
    for node in nodes[1:]:
        node.text = ""
    for el in drop:
        el.getparent().remove(el)
    return None


def set_run_xml(r, cn=None, latin=None, size=None):
    """在原始 XML 层给一个 w:r 套字体字号。

    脚注部件不在 python-docx 的对象模型里，拿不到 Run 对象，只能直接改元素。
    """
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    if cn:
        rf.set(qn("w:eastAsia"), cn)
    if latin:
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)
    if size:
        for tag in ("w:sz", "w:szCs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rpr.append(el)
            el.set(qn("w:val"), str(int(round(size * 2))))


def run_style(r):
    """读一个 w:r 当前的中文/西文字体、字号（半磅）与加粗，用于判断是否真的需要改。"""
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return (None, None, None, None)
    rf = rpr.find(qn("w:rFonts"))
    sz = rpr.find(qn("w:sz"))
    b = rpr.find(qn("w:b"))
    if b is None:
        bold = None
    else:
        val = b.get(qn("w:val"))
        bold = False if val in ("0", "false") else True
    return (rf.get(qn("w:eastAsia")) if rf is not None else None,
            rf.get(qn("w:ascii")) if rf is not None else None,
            int(sz.get(qn("w:val"))) if sz is not None and sz.get(qn("w:val")) else None,
            bold)


def run_hansi(r):
    """读一个 w:r 的 hAnsi 字体。set_run 写 ascii 时会一并重写 w:hAnsi，
    快照不看它的话，「仅 hAnsi 不同」的真改动就成了静默改动。"""
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return None
    rf = rpr.find(qn("w:rFonts"))
    return rf.get(qn("w:hAnsi")) if rf is not None else None


def style_footnotes(doc, rules):
    """按规则集给脚注正文套字体字号，返回 (实际改动的条数, 目标描述)。

    期刊对脚注字号要求差异很大（这本 7.5pt、那本 9pt），而脚注部件独立于正文，
    正文套版碰不到它。不处理的话，成品正文合规、脚注字号还是稿件原来的——
    编辑一眼就看出来，报告却写着"已按规则集处理"。

    只统计**真的变了**的脚注：本来就合规的不进报告，否则"改为 7.5pt"这种话
    会出现在什么都没变的稿子上。

    返回 (改动条数, 目标描述, 未处理的外部脚注关系数)：TargetMode="External"
    的脚注关系没有可解析部件（访问 target_part 抛 ValueError，EXTFOOT），
    跳过并把数量交给调用方写进汇总。
    """
    tp = (rules.get("typography") or {}).get("footnote") or {}
    size, cn = tp.get("size_pt"), tp.get("font_cn")
    latin = tp.get("font_latin") or "Times New Roman"
    if not (size or cn):
        # 没声明脚注样式时也数一下外部关系，汇总照样要报（渲染不崩、有说明）
        return 0, "", sum(1 for rel in doc.part.rels.values()
                          if rel.reltype.endswith("/footnotes") and rel.is_external)

    changed = 0
    external = 0
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/footnotes"):
            continue
        if rel.is_external:
            external += 1
            continue
        try:
            blob = rel.target_part.blob
            xmlsafe.reject_dtd(blob)
            parser = etree.XMLParser(resolve_entities=False, no_network=True)
            root = etree.fromstring(blob, parser)
        except (etree.XMLSyntaxError, xmlsafe.UnsafeXml):
            continue

        touched = False
        for fn in root.findall(qn("w:footnote")):
            # 分隔符与续行符没有正文，不必动。判断依据是 w:type 与分隔符元素，
            # 不是 w:id：Word 只要求这两条带 type，id 用多少由写入它的软件决定。
            # WPS 存出来的稿子里分隔符是 id=4/5，而 id=0 是**正文脚注**
            # （作者信息那一条）。旧代码按 id<1 排除，正好把作者脚注跳过、
            # 反倒把分隔符改了字号。老式写法（只有 id 没有 type）靠 separator
            # 元素兜住。
            if (fn.get(qn("w:type"))
                    or fn.find(".//" + qn("w:separator")) is not None
                    or fn.find(".//" + qn("w:continuationSeparator")) is not None):
                continue

            before = [run_style(r) for r in fn.iter(qn("w:r"))]
            for r in fn.iter(qn("w:r")):
                set_run_xml(r, cn=cn, latin=latin, size=size)
            after = [run_style(r) for r in fn.iter(qn("w:r"))]

            if before != after:
                changed += 1
                touched = True

        if touched:
            rel.target_part._blob = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True)

    desc = f"{cn or ''} {size:g}pt".strip() if size else (cn or "")
    return changed, desc, external


def replace_prefix(p, old_len, new_prefix):
    """只改段落开头的 old_len 个字符，其余文字与格式原样保留。

    参考文献条目常带斜体刊名，整段重写会把这些格式一起抹平，所以这里
    逐个文字节点扣减，不重建段落。

    old_len 按 p.text 同一口径的聚合文本数出来（w:tab 占 1 字符、换行占
    1 字符），扣减也必须走同一口径：tab/br 是独立的物理节点、不在任何
    w:t 文本里，只按 w:t 扣字符会把多出来的账全扣到正文头上——
    「关键词\\t：数字金融」能被改成「关键词：\\t字金融」，凭空吃掉一个
    「数」字。现在落在前缀区里的 tab/br 随前缀一起删除，审计的 after
    与成品逐字一致。

    新前缀写回「原前缀所在的第一个文字节点」，因此继承的是前缀自己的
    字体与加粗：稿件里「摘要」两字常单独加粗，改成「内容摘要：」后仍是
    粗体标签，而正文那一段依旧不加粗。若把整段塞进第一个 run（旧做法），
    标签的粗体会传染给整段摘要——这正是转换后摘要整段变粗的原因。
    """
    remaining = old_len
    placed = False
    for el, chars in text_units(p):
        if remaining <= 0:
            break
        take = min(remaining, len(chars))
        if el.tag == qn("w:t"):
            if not placed:
                _write_t(el, new_prefix + chars[take:])
                placed = True
            else:
                _write_t(el, chars[take:])
        elif take:
            el.getparent().remove(el)   # 前缀区里的 tab/br 随前缀一起删
        remaining -= take


def replace_after(p, keep_len, new_text):
    """保留开头 keep_len 个字符（通常就是「关键词：」这类标签），其余整体替换。

    和 replace_text 的区别：新文字落在「正文所在的 run」里，而不是塞进第一个
    run。标签若是加粗的，用 replace_text 会把整段都变成粗体；这里只动标签
    之后的文字，标签与正文各自保有自己的格式。

    字符口径与 replace_prefix 相同：按聚合文本计数（w:tab 占 1、换行占 1），
    落在替换区里的 tab/br 一并删除，保留区里的原样留着。

    keep_len 恰好落在物理单元边界上时（标签独占一个加粗 run，或前缀归一刚把
    「关键词\\t：」合并成一个节点），新文字必须写进替换区第一个文字节点——
    只在跨界节点里写的话，边界对齐的段落没有任何节点进入「写」分支，
    整段正文会被静默清空。替换区没有文字节点时兜底另起一个 run。
    """
    remaining = keep_len
    placed = False
    for el, chars in text_units(p):
        if remaining > 0:
            take = min(remaining, len(chars))
            if take < len(chars):
                # tab/br 只占 0/1 个字符，跨界分支必是文字节点：
                # 保留属于标签的半截，新文字接在它后面
                _write_t(el, chars[:take] + new_text)
                placed = True
            remaining -= take             # 边界恰好落在单元末尾时整单元保留
            continue
        if el.tag == qn("w:t"):
            if placed:
                _write_t(el, "")
            else:
                _write_t(el, new_text)    # 新文字落在正文自己的 run，继承正文格式
                placed = True
        else:
            el.getparent().remove(el)     # 替换区里的 tab/br 随正文一起删
    if not placed:
        p.add_run(new_text)


INDENT_WS = re.compile(r"^[\t \u3000]+")


def strip_leading_indent(p):
    """删掉段首手工敲入的缩进空白（全角空格 U+3000、半角空格、制表符）。

    返回删除的字符数；段首没有缩进空白时返回 0。口径与 replace_prefix 相同
    （按聚合文本计数，前缀区里的 tab 物理节点随前缀一起删）；换行不算缩进
    空白，段首的 w:br 原样保留。这就是白名单 N-01 的段首版与 N-08 的落地。
    """
    m = INDENT_WS.match(para_full_text(p))
    if not m:
        return 0
    replace_prefix(p, m.end(), "")
    return m.end()


def _frag(text, limit=24):
    """审计条目用的片段：只动段首时，头部片段足以定位与核对。"""
    return text if len(text) <= limit else text[:limit] + "…"


def strip_leading_indent_audited(p, audit, idx, btype, indent):
    """段首缩进空格删除（N-01 段首＋N-08）：逐块进审计，可驳回。

    手工敲的段首空格与渲染器写入的首行缩进设置叠加就是双重缩进——手敲缩进
    的稿子每份都带着这个观感缺陷。删掉空格、缩进改由段落设置接管；驳回则
    保留原文，缩进设置照写（文字归文字，版式归版式），条目如实记「已驳回」。
    """
    before_txt = block_text(p)
    m = INDENT_WS.match(before_txt)
    if not m:
        return
    rule = "typography.body.first_line_indent_chars"
    sig = reject_sig(before_txt)
    if audit.blocked(idx, "段首缩进空格删除", sig):
        audit.add(idx, btype, "段首缩进空格删除", _frag(before_txt), _frag(before_txt),
                  f"已驳回：保留段首空格，首行缩进仍按 {indent:g} 字符设置",
                  rejected=True, rule=rule, sig=sig)
        return
    strip_leading_indent(p)
    audit.add(idx, btype, "段首缩进空格删除", _frag(before_txt), _frag(block_text(p)),
              f"删除段首 {m.end()} 个缩进空白（全角/半角空格、制表符），"
              f"首行缩进改由段落设置 {indent:g} 字符接管",
              rule=rule, sig=sig)


def split_keywords(text, sep):
    """把标签后的关键词串切成词条。

    逗号、顿号、分号是明确的分隔符，总是切。空白分两等：目标分隔符本身
    就是空白时全部当分隔符；否则只有「连续两个以上的空白」或全角空格才
    算分隔符——英文词组「digital finance」里的单个空格是词间空格，把它
    当分隔符会把一个关键词凭空拆成两个，词条数说变就变。
    """
    pattern = (r"[\s,，;；、]+" if not (sep or "").strip()
               else r"(?:\s{2,}|\u3000|[,，;；、])+")
    return [x for x in re.split(pattern, text or "") if x]


REF_NUM = re.compile(
    r"^(?P<open>[\[［〔【(（])?\s*(?P<num>\d{1,3})\s*(?P<close>[\]］〕】)）])?(?P<tail>[.．、]?)")

def ref_number_prefix(text):
    """识别参考文献条目开头的编号前缀，返回 (原文前缀, 序号) 或 (None, None)。

    只认真正像编号的开头：带括号，或数字后跟句点/顿号/空白。否则
    「2016 年……」这种以年份开头的条目会被当成 201 号文献。
    """
    m = REF_NUM.match(text)
    if not m or not m.group("num"):
        return None, None
    bracketed = bool(m.group("open") and m.group("close"))
    tail = m.group("tail")
    # 看数字紧后面那个字符：空白说明它是「123 张三」这种编号，
    # 而「2020：」的 202 后面是数字 0，不该被当成编号
    after = text[m.end("num"):m.end("num") + 1]
    if not (bracketed or tail or after.isspace()):
        return None, None
    # 编号与其后正文之间的空白也算前缀区：目标刊的 space_after 会重新给出
    # 这段空白。不计入的话，`replace_prefix` 只删掉 `[1]` 三个字符再写回
    # `[1] `，而原文本里那个空格还在——每重跑一次多叠一个空格，条目越跑
    # 越松。这是幂等的前提，不是格式洁癖。
    trailing = re.match(r"\s*", text[m.end():]).group(0)
    return m.group(0) + trailing, int(m.group("num"))


def format_ref_number(fmt, num, space_after=None):
    """按模板把序号渲染成目标刊的编号写法，如 [1] / 〔1〕 / 1．"""
    if not fmt:
        return None
    open_ = close = tail = ""
    m = re.search(r"([\[［〔【(（])\s*1\s*([\]］〕】)）])", fmt)
    if m:
        open_, close = m.group(1), m.group(2)
    elif fmt.strip() in ("1．", "1.", "1、"):
        tail = fmt.strip()[1:]
    out = f"{open_}{num}{close}{tail}"
    if space_after:
        out += " "
    return out


def _grid_span(tc):
    """横向合并占几列（w:gridSpan 缺省为 1）。"""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return 1
    el = tcPr.find(qn("w:gridSpan"))
    if el is None:
        return 1
    try:
        return max(1, int(el.get(qn("w:val"))))
    except (TypeError, ValueError):
        return 1


def _vmerge_state(tc):
    """纵向合并状态：restart（起始）/ continue（续行空壳）/ None（不合并）。

    `w:vMerge` 没有 val 时按 OOXML 规范就是 continue——直接写 `<w:vMerge/>`
    的稿子很常见，漏了这条会把续行当成一个独立单元格。
    """
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    el = tcPr.find(qn("w:vMerge"))
    if el is None:
        return None
    return "restart" if (el.get(qn("w:val")) or "").lower() == "restart" else "continue"


def _table_grid(tbl):
    """把表格铺成 grid[row][col] = tc，纵向合并的续行回填成起始 tc。

    为什么必须铺开：三线表的栏目线是「表头区与表体区的分界」，而经管论文的
    表头经常跨两行（「变量类型」跨行、分组标题跨列）。按 row.cells 逐行判断
    `ri == 0` 时，跨行单元格的 tc 在第 0 行，而它的底边落在合并区底部——
    于是同一张表里一部分列的线画在第 0 行下、一部分画在合并区底部，线是断开的。
    """
    grid, starts = [], {}
    for r, tr in enumerate(tbl.findall(qn("w:tr"))):
        row, col = [], 0
        for tc in tr.findall(qn("w:tc")):
            span = _grid_span(tc)
            if _vmerge_state(tc) == "continue":
                # 回溯上一行同列的起始 tc；拿不到就用自己的（异常稿兜底）
                tc = starts.get((r - 1, col), tc)
            for k in range(span):
                starts[(r, col + k)] = tc
            row.extend([tc] * span)
            col += span
        grid.append(row)
    return grid


def _header_rows(grid):
    """表头区占几行。

    vMerge 的续行已在 _table_grid 里回填成起始 tc，所以「相邻两行出现同一个
    元素」就等价于「这里有一个跨行单元格」——这是判断表头跨几行最干净的依据，
    不必去猜字号或空行。

    上限取 len-1：栏目线是「表头与表体的分界」，全表都算表头的话这条线没有
    分隔对象，画出来只会和表格底线重叠。
    """
    if not grid:
        return 0
    n = 1
    for r in range(1, len(grid)):
        if set(grid[r - 1]) & set(grid[r]):
            n = r + 1
        else:
            break
    return min(n, max(1, len(grid) - 1))


def apply_table(table, bp, size_pt, cn, latin):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(el)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(int(bp[edge] * 8)))
        else:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)

    # 栏目线按「逻辑单元格的底边落在哪一行」来画，而不是按 ri == 0。
    # 合并单元格的边框只有起始 tc 上的算数（续行空壳的边框 Word 直接忽略），
    # 所以必须先铺成 grid、把续行回填成起始 tc，再判断底边位置。
    grid = _table_grid(tbl)
    n_header = _header_rows(grid)
    # 每个逻辑单元格占的行区间
    extent = {}
    for r, row in enumerate(grid):
        for tc in row:
            lo, hi = extent.get(tc, (r, r))
            extent[tc] = (min(lo, r), max(hi, r))

    seen = set()
    for row in grid:
        for tc in row:
            if tc in seen:
                continue
            seen.add(tc)
            tcPr = tc.get_or_add_tcPr()
            for el in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(el)
            # 底边正好落在表头区最后一行的，才画栏目线：跨行表头的单元格
            # 底边在合并区底部，正是"表头与表体之间"那条线该在的位置
            if extent[tc][1] == n_header - 1:
                tcb = OxmlElement("w:tcBorders")
                b = OxmlElement("w:bottom")
                b.set(qn("w:val"), "single")
                b.set(qn("w:sz"), str(int(bp["header_bottom"] * 8)))
                b.set(qn("w:space"), "0")
                b.set(qn("w:color"), "000000")
                tcb.append(b)
                tcPr.append(tcb)
            for p_el in tc.findall(qn("w:p")):
                for run in Paragraph(p_el, table).runs:
                    set_run(run, cn=cn, latin=latin, size=size_pt)


def table_snapshot(table):
    """三线表套版前后的状态：表级边框、单元格级边框、表内字体字号。

    apply_table 无条件重写的每一项都必须在快照里：边框除 val/sz 外还有
    w:space 与 w:color（统一重写为 0 与黑色），run 字体之外还有 w:hAnsi
    （set_run 写 ascii 时一并重写）——少看任何一样，「形状合规但颜色
    不同」这类真改动前后快照相同，成了静默改动，审计漏记。
    快照口径齐全后，本来就完全合规的三线表前后快照相同，不会在重跑时
    再报一次「表格重设为三线表」（F28）。
    """
    borders = {}
    el = table._tbl.tblPr.find(qn("w:tblBorders"))
    if el is not None:
        for child in el:
            borders[etree.QName(child).localname] = (child.get(qn("w:val")),
                                                     child.get(qn("w:sz")),
                                                     child.get(qn("w:space")),
                                                     child.get(qn("w:color")))
    # 与 apply_table 同一套 grid 口径：合并单元格只按起始 tc 记一次。
    # 两边口径不一致的话，快照会记到 apply_table 根本没碰的续行空壳上，
    # 「已完全合规」的表格每次重跑都被判成改动（F28 的老毛病换了个位置复发）。
    grid = _table_grid(table._tbl)
    cells, seen = [], set()
    for row in grid:
        for tc in row:
            if tc in seen:
                continue
            seen.add(tc)
            tcpr = tc.find(qn("w:tcPr"))
            tcb = tcpr.find(qn("w:tcBorders")) if tcpr is not None else None
            cb = tuple(sorted((etree.QName(c).localname, c.get(qn("w:val")), c.get(qn("w:sz")),
                               c.get(qn("w:space")), c.get(qn("w:color")))
                              for c in tcb)) if tcb is not None else ()
            fonts = tuple((run_style(r), run_hansi(r))
                          for p_el in tc.findall(qn("w:p"))
                          for r in para_runs(Paragraph(p_el, table)))
            cells.append((cb, fonts))
    return {"borders": borders, "cells": tuple(cells)}


# sdt 嵌套下钻的深度上限（SDTCAP）：与 parse_docx.SDT_MAX_DEPTH 同一口径，
# 两边必须一致——解析端到顶不下钻的部分计进 unsupported_content_blocks，
# 渲染端若继续下钻，sdt_pos 就会与 IR 错位、套版错段
SDT_MAX_DEPTH = 8


def _sdt_blocks(content, depth=1):
    """sdtContent 递归下钻：返回（可处理块元素列表, 未识别叶节点数）。

    depth 是当前内容控件的嵌套层数（顶层 sdt 为 1，SDTCAP）：超过
    SDT_MAX_DEPTH 的更深 w:sdt 不再下钻，整体计 1 个未识别结构。

    w:p / w:tbl 直接收进块列表；内嵌 w:sdt 继续下钻，其块按文档序并入同一
    列表（SDTNEST：批五只下钻一层，外层有直接段落时内嵌 sdt 的段落既不进
    IR 也不计数，无声丢失）；其余子节点是不认识的叶节点，逐个计数。一个块
    都没取到、也没有未识别节点（空内容控件）时计 1——整块没有可处理段落/
    表格的对象要报出来（F24 口径，内容原样保留）。

    渲染端必须与 parse_docx._sdt_blocks 保持同一展开顺序与同一深度上限：
    IR 的 sdt_pos 是槽位内序，两边顺序不一致就套错段。
    """
    blocks, unknown = [], 0
    for e in (list(content) if content is not None else []):
        if e.tag in (qn("w:p"), qn("w:tbl")):
            blocks.append(e)
        elif e.tag == qn("w:sdt"):
            if depth >= SDT_MAX_DEPTH:
                unknown += 1
                continue
            sub_blocks, sub_unknown = _sdt_blocks(e.find(qn("w:sdtContent")),
                                                  depth + 1)
            blocks.extend(sub_blocks)
            unknown += sub_unknown
        else:
            unknown += 1
    if not blocks and not unknown:
        unknown = 1
    return blocks, unknown


def elem_of(blk, elems):
    """块 → 文档元素。

    sdt 内容控件里的块（parse_docx 下钻产出，带 sdt_pos）落在同槽位的元素
    列表里，按 sdt_pos 取（F24）。槽位编号仍与 body 子元素位置同源，消费端
    按 idx 定位的契约不破。列表里取不到（IR 与文档不配套）就返回 None，
    不猜位置——猜错就是套版错段。
    """
    el = elems.get(blk.get("idx"))
    if isinstance(el, list):
        pos = blk.get("sdt_pos")
        return el[pos] if isinstance(pos, int) and 0 <= pos < len(el) else None
    return el


def footnote_texts(doc):
    """脚注部件里的正文文本（F23）：脚注不在 body，正文体检看不见它。

    解析姿势与 style_footnotes 同一套（拒 DOCTYPE/ENTITY、不解析外部实体）。
    分隔符/续行符没有正文，跳过。
    """
    texts = []
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/footnotes"):
            continue
        if rel.is_external:
            # TargetMode="External" 的脚注关系（外部超链接式）没有可解析部件，
            # 访问 target_part 抛 ValueError（EXTFOOT 实测崩溃路径）——它也不
            # 可能藏编号引用，直接跳过
            continue
        try:
            blob = rel.target_part.blob
            xmlsafe.reject_dtd(blob)
            parser = etree.XMLParser(resolve_entities=False, no_network=True)
            root = etree.fromstring(blob, parser)
        except (etree.XMLSyntaxError, xmlsafe.UnsafeXml):
            continue
        for fn in root.findall(qn("w:footnote")):
            if (fn.get(qn("w:type"))
                    or fn.find(".//" + qn("w:separator")) is not None
                    or fn.find(".//" + qn("w:continuationSeparator")) is not None):
                continue
            t = "".join(n.text or "" for n in fn.iter(qn("w:t")))
            if t.strip():
                texts.append(t)
    return texts


def reorder_notes(elems, ir, rules, audit):
    """按目标刊的层序重排表注／图注。

    真实差异：国际金融研究是「资料来源 → 注」，金融研究是「注 → 资料来源」。
    稿件通常沿用其中一种，投另一本时必须换过来，否则读者看到的解释顺序是错的。
    """
    layers = {
        "table": [l["kind"] for l in rules["tables"]["note_layers"]],
        "figure": [l["kind"] for l in rules["figures"]["note_layers"]],
    }
    blocks = ir["blocks"]

    def kind_of(txt):
        if txt.startswith("数据来源"):
            return "data_source"
        if txt.startswith("资料来源"):
            return "source"
        if txt.startswith("注"):
            return "general"
        return "cell"

    for i, b in enumerate(blocks):
        if b["type"] not in ("table", "figure"):
            continue
        j = i + 1
        cap = None
        # 图题夹在插图和图注之间，收集图注时要先跳过它
        if b["type"] == "figure" and j < len(blocks) and blocks[j]["type"] == "figure_caption":
            cap = blocks[j]
            j += 1
        notes = []
        while j < len(blocks) and blocks[j]["type"] == "table_note":
            notes.append(blocks[j])
            j += 1
        if len(notes) < 2:
            continue
        rank = {k: n for n, k in enumerate(layers[b["type"]])}
        seq = [rank.get(kind_of(n["text"]), 99) for n in notes]
        if seq == sorted(seq):
            continue
        if audit.blocked(b["idx"], "表注／图注层序调整",
                         reject_sig(" / ".join(kind_of(n["text"]) for n in notes))):
            audit.add(b["idx"], b["type"], "表注／图注层序调整",
                      " / ".join(kind_of(n["text"]) for n in notes),
                      " / ".join(kind_of(n["text"]) for n in notes),
                      "已驳回：保持原层序", rejected=True,
                      rule=f"{b['type']}s.note_layers",
                      sig=reject_sig(" / ".join(kind_of(n["text"]) for n in notes)))
            continue
        if b["type"] == "table":
            anchor = elem_of(b, elems)._tbl
        else:
            anchor = elem_of(cap or b, elems)._p
        ordered = [n for _, n in sorted(zip(seq, notes), key=lambda x: x[0])]
        for n in notes:
            el = elem_of(n, elems)._p
            el.getparent().remove(el)
        cur = anchor
        for n in ordered:
            el = elem_of(n, elems)._p      # sdt 内的注也要能搬回（F24，同上）
            cur.addnext(el)
            cur = el
        audit.add(b["idx"], b["type"], "表注／图注层序调整",
                  " / ".join(kind_of(n["text"]) for n in notes),
                  " / ".join(kind_of(n["text"]) for n in ordered),
                  rule=f"{b['type']}s.note_layers",
                  sig=reject_sig(" / ".join(kind_of(n["text"]) for n in notes)))


# 文献表排序体例（F21）：行为完全由规则集 citation.reference_list.sort 声明
# 决定。渲染器只实现 SORT_NOTE 里列出的声明；appearance 类保持条目原序不排；
# 声明不认识（或没声明）就宁可不排，note 如实说明——不再内置「中文在前、
# 英文在后」却冒用 citation.reference_list.sort 之名出审计。
APPEARANCE_SCHEMES = ("by_appearance_order", "by-appearance")
# 两种「中文优先」声明的中文键统一用 GB 序（一级字库即拼音序）实现：声明里的
# 「拼音序/字母序」对中文的可行实现就是它——码位序会把「周一」排到「张三」
# 前面，与声明的字面意思相反（JRYJNOTE：note 描述的排序必须真的发生，F21/F22
# 措辞与实现一致的原则不变，这里改的是实现而不是措辞迁就实现）
SORT_NOTE = {
    "cn-first-by-pinyin-then-en-by-alphabet":
        "中文按 GB 序（近似拼音）在前、英文按字母序在后",
    "cn-first-by-alphabet-then-en-by-alphabet":
        "中文按字母序（GB 序，近似拼音）在前、英文按字母序在后",
    "alphabetical": "条目按字母序排列",
}


def sort_key(entry, scheme):
    """按目标刊声明的排序体例给条目一个比较键（F21/F22/JRYJNOTE）。

    排序键完全由 scheme 决定，这里不内置「中文在前」的默认。中文键统一用
    GB2312 序近似（citation_ref.gb_sort_key）——环境里没有拼音库，码位序
    会把「周一」排到「张三」前面；GB 序的一级字库本身就是拼音序，两种
    「中文优先」声明的「拼音序/字母序」对中文都由它实现。英文键是姓氏的
    str.lower 键序。审计措辞见 SORT_NOTE，必须描述实际发生的排序。
    """
    name = entry["authors"][0] if entry["authors"] else ""
    if scheme == "alphabetical":
        # 全文键名：剥掉条目开头的 [n] 编号，整条按字母排（英文刊的声明）
        import citation_ref as cr
        return (cr.NUMBERED_PREFIX.sub("", entry["raw"]).lower(), entry["year"])
    if entry["lang"] == "cn":
        import citation_ref as cr
        return (0, cr.gb_sort_key(name), entry["year"])
    return (1, re.split(r"[\s,，]", name)[0].lower(), entry["year"])


def rewrite_citations(p, index, rules):
    """把段落里的编号引用改写成作者-年份，返回改写处数。

    命中在段落聚合文本上找（与 cr.check 完全同一口径），再按偏移量落回
    物理文字节点——引用横跨多个 run、藏在超链接或修订插入（w:ins）里时
    也改得到。改写落在命中起点所在的节点（继承该节点的字体），命中区里
    其余文字节点清空、tab/br 删除；各节点只按自己命中的区间重写自己的
    文本，节点之间互不覆盖。
    """
    import citation_ref as cr

    nodes, parts, pos = [], [], 0
    for el, chars in text_units(p):
        nodes.append((el, pos, chars))
        parts.append(chars)
        pos += len(chars)
    text = "".join(parts)

    # 多合一括号与分隔符来自规则集 citation.in_text（只消费作者-年份制刊物的
    # 声明，见 cr.in_text_style；编号制刊物的 bracket「[]」描述编号标注，不取），
    # 缺省回退全角——中文刊行为与硬编码时期逐字节一致（PB1）
    (open_b, close_b), sep = cr.in_text_style(rules)

    spans = []
    for m in cr.citation_spans(text):
        labels = []
        for n in cr.expand_numbers(m.group(1)):
            e = index.get(n)
            if not e or not e["year"]:
                labels = None
                break
            labels.append(cr.label_body(e, rules))
        if labels:
            spans.append((m.start(), m.end(), open_b + sep.join(labels) + close_b))
    if not spans:
        return 0

    # 每个命中的替换文字都落在命中起点所在的节点上
    owners = [next(el for el, off, chars in nodes if off <= s < off + len(chars))
              for s, _e, _r in spans]
    hits = 0
    for el, off, chars in nodes:
        mine = [(s - off, e - off, repl, owner is el)
                for (s, e, repl), owner in zip(spans, owners)
                if off < e and s < off + len(chars)]
        if not mine:
            continue
        if el.tag != qn("w:t"):
            el.getparent().remove(el)    # 命中区里的 tab/br 属于被替换的文字
            continue
        buf, cur = [], 0
        for local_s, local_e, repl, is_owner in mine:
            # 命中可以跨节点：本节点只认自己覆盖的那段（起点可能在本节点之前）
            local_s, local_e = max(local_s, 0), min(local_e, len(chars))
            buf.append(chars[cur:local_s])
            buf.append(repl if is_owner else "")
            cur = local_e
        buf.append(chars[cur:])
        _write_t(el, "".join(buf))
        owners_here = sum(1 for *_x, is_owner in mine if is_owner)
        hits += owners_here
        if owners_here:
            # 原稿若是上标引用，转成作者-年份后必须转回正常字号
            rpr = el.getparent().find(qn("w:rPr"))
            if rpr is not None:
                for va in rpr.findall(qn("w:vertAlign")):
                    rpr.remove(va)
    return hits


def apply_citation(doc, elems, ir, rules, audit):
    """正文编号引用 → 作者-年份，文献条目按目标刊体例重写并重排。

    只在用户勾选时调用（--citation）。这里会改写文字，触及产品红线，
    所以先跑双向校验：
      - 正文每个编号都要能定位到文献表某一条
      - 每条文献都要被正文引用过
    任一不满足就整批不转，把原因写进审计——引错文献比不转严重得多。
    """
    import citation_ref as cr

    refs = [(b["idx"], elem_of(b, elems)) for b in ir["blocks"]
            if b["type"] == "reference_item"]
    # 条目文字按聚合口径取：整条藏在超链接或修订插入里的条目 p.text 看不见
    refs = [(i, p) for i, p in refs if p is not None and block_text(p).strip()]
    if not refs:
        return

    # EndNote 插入的文献表是一整个 EN.REFLIST 域，条目的作者、年份、编号全在
    # 域结果里。重排它等于绕过文献库另写一份：作者返修时按 F9 更新，整表
    # 会按库里的旧数据重新生成，重排白做，还可能把已改的条目变成死文本。
    by_idx = {b["idx"]: b for b in ir["blocks"]}
    locked_refs = [i for i, _p in refs if text_locked(by_idx.get(i) or {})]
    if locked_refs:
        audit.add(-3, "citation", "条目重排（含域，已跳过）",
                  note=f"文献表 {len(locked_refs)} 条是文献管理软件的域结果，"
                       f"重排会与文献库脱钩（Word 里按 F9 更新即整表回滚），"
                       f"故保持原样；如需按本刊体例重排，请先在 EndNote/Zotero 里"
                       f"改输出样式后重新插入",
                  rule="citation.reference_list.templates", changed=False)
        return

    # 体例硬闸（P4-1）：引用重排只服务于声明了作者-年份体的目标刊。编号实引的
    # 刊（cjyj 等，citation.system=numbered）勾了也不转——编号映射再正确，交付
    # 稿的引用体例也是错的。判别只读规则集字段，代码不内置任何刊名；未声明
    # system 时同样不转（越线开关宁可不做，见 journal-schema §4）。
    system = ((rules or {}).get("citation") or {}).get("system") or ""
    if system != "author-year":
        why = ("该刊声明编号实引，不转" if system == "numbered"
               else f"该刊引用体例声明为「{system or '未声明'}」，仅 author-year 才转")
        audit.add(-3, "citation", "引用体系转换（已跳过）",
                  note=f"{why}（citation.system，正文编号引用与文献表保持原样）",
                  rule="citation.system", changed=False)
        return

    index = cr.build_index([block_text(p) for _, p in refs])
    body_paras = [(b["idx"], b["type"], elem_of(b, elems)) for b in ir["blocks"]
                  if b["type"] not in ("reference_item", "references_heading")
                  and hasattr(elem_of(b, elems), "text")]
    body_paras = [x for x in body_paras if x[2] is not None]

    # 体检与改写必须同一份文本：体检用 p.text、改写走直接 run 时，藏在
    # 超链接、修订插入里或横跨多个 run 的引用能过闸门却永远改不到，
    # 条目重排后残留编号指向错位的文献，审计里还没有漏改记录
    cited, problems = cr.check([block_text(p) for _, _, p in body_paras], index, rules)

    # F23 fail-closed 两道补充闸。这三类位置两头（体检、改写）都够不着，
    # 转换后必然出现指向错位/已改写条目的残留——整批不转，点名让人看得见。
    strays = cr.unnumbered_entries([block_text(p) for _, p in refs])
    if strays:
        problems.append("文献表有 {} 条不带 [n] 编号的条目：{}".format(
            len(strays), "；".join(s[:20] for s in strays[:3])))
    fn_hits = cr.footnote_citation_hits(footnote_texts(doc))
    if fn_hits:
        problems.append("脚注里出现编号引用 {}，条目重排会让它指向错位的条目".format(
            "、".join(m.group(0) for m in fn_hits[:3])))

    if problems:
        audit.add(-3, "citation", "引用体系转换（已跳过）",
                  note="；".join(problems[:4]) + ("…" if len(problems) > 4 else ""),
                  rule="citation.reference_list.must_require_bidirectional_match",
                  changed=False)
        return

    # ---- 正文引用点 ----
    for idx, btype, p in body_paras:
        old_text = block_text(p)
        if rewrite_citations(p, index, rules):
            audit.add(idx, btype, "正文引用改作者-年份", old_text, block_text(p),
                      "编号引用换写法，正文其余文字未动",
                      rule="citation.in_text")

    # ---- 数量语境的可疑样文本：不改写，但逐处报给人看 ----
    # 形似引用却紧邻量词/区间词的方括号（「调研对象为[90, 100]家企业」），
    # 转了是语义事故；不转也得让人知道这里有个原样保留的编号样文本。
    for idx, btype, p in body_paras:
        text = block_text(p)
        spans = cr.ambiguous_spans(text)
        if not spans:
            continue
        frags = []
        for m in spans[:3]:
            s = max(0, m.start() - 8)
            frags.append("…" + text[s:m.end() + 8].strip() + "…")
        audit.add(idx, btype, "编号样文本未转换",
                  note="形似引用但紧邻数量/区间说法，无法安全判定，原样保留："
                       + "；".join(frags),
                  rule="citation.in_text", changed=False)

    # ---- 文献条目重排体例并按目标刊排序 ----
    rewritten = []
    for idx, p in refs:
        num = next((n for n, e in index.items() if e["raw"] == block_text(p)), None)
        if num is None:
            continue
        entry = index[num]
        new = cr.format_entry(entry, rules)
        old = block_text(p)
        if not new or new == old:
            continue
        replace_text(p, new)
        rewritten.append((idx, p, entry))
        audit.add(idx, "reference_item", "条目按目标刊体例重排", old, new,
                  "类型、卷期页字段机器提取，请复核作者名与刊名",
                  rule="citation.reference_list.templates")

    if rewritten:
        # 排序体例只由规则集声明决定（F21）：声明了渲染器实现的 scheme 才排；
        # appearance 类保持条目原序不排；声明不认识（或没声明）就宁可不排、
        # note 如实说明，绝不内置「中文在前、英文在后」冒充已排序。
        scheme = (((rules or {}).get("citation") or {}).get("reference_list")
                  or {}).get("sort") or ""
        if scheme in SORT_NOTE:
            ordered = sorted(rewritten, key=lambda x: sort_key(x[2], scheme))
            sort_note = SORT_NOTE[scheme]
            sort_rule = "citation.reference_list.sort"
        else:
            # appearance 类声明与未知/未声明声明都不排
            ordered = list(rewritten)
            sort_note = ("按声明保持正文引用出现序" if scheme in APPEARANCE_SCHEMES
                         else f"未支持的排序声明「{scheme or '未声明'}」，保持原序")
            sort_rule = ("citation.reference_list.sort" if scheme
                         else "citation.reference_list~")
        # 锚点必须选一个不会跟着搬家的元素：「参考文献」标题。若拿条目自己的
        # 前一个兄弟当锚点，那个兄弟可能也在待搬列表里，摘掉后整段条目就丢了。
        anchor = None
        for b in ir["blocks"]:
            if b["type"] == "references_heading":
                el = elem_of(b, elems)
                if el is None:
                    # IR 与文档失配时 elem_of 返回 None（按设计），换下一个
                    # 标题候选；在 None 上 break 会让 anchor 落进兜底
                    continue
                anchor = el._p
                break
        moving = {p._p for _, p, _e in ordered}
        if anchor is None:
            # 没有标题时的兜底：条目的文档前驱。ordered 是排序序而非文档序，
            # ordered[0] 的前驱完全可能是另一条待搬条目——不校验「不在待搬
            # 集合」就把锚点自己摘掉，整段文献表会以脱挂碎片的形式静默丢失。
            cand = ordered[0][1]._p.getprevious()
            if cand is not None and cand not in moving:
                anchor = cand
        if anchor is None:
            # 找不到安全的锚点：宁可不搬（条目文字已按目标刊体例重写，只是
            # 保持原文档序），审计如实说明——静默丢整段文献表比不排序严重
            audit.add(-4, "reference_item", "文献表排序",
                      note="找不到不会跟着搬家的锚点（IR 无可解析的文献表标题，"
                           "条目的文档前驱也在待搬集合），条目保持原文档序",
                      rule=sort_rule, changed=False)
            return
        for _, p, _e in ordered:
            p._p.getparent().remove(p._p)
        for _, p, _e in ordered:
            anchor.addnext(p._p)
            anchor = p._p
        changed = [x[0] for x in ordered] != [x[0] for x in rewritten]
        # 顺序没变（或声明不支持本就不排）时条目是零改动的说明——
        # changed=False，否则说明型条目虚报进改动数（见 Audit.add 约定）
        if changed or (scheme not in SORT_NOTE
                       and scheme not in APPEARANCE_SCHEMES):
            audit.add(-4, "reference_item", "文献表排序",
                      note=sort_note, rule=sort_rule, changed=changed)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ir")
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--reject", help="驳回清单 JSON：{\"ids\": [...]} 或 [...]，键为「块序号|动作」")
    ap.add_argument("--citation", action="store_true",
                    help="编号引用改为目标刊的作者-年份体例（会改写文献条目文字，默认不做）")
    a = ap.parse_args()

    ir = json.loads(Path(a.ir).read_text(encoding="utf-8"))
    rules = json.loads(Path(a.ruleset).read_text(encoding="utf-8"))

    # F13：规则集缺字段不在渲染中途裸 KeyError（那会被上层一律翻译成
    # 「稿件结构超出当前识别范围」，把锅甩给稿件——失败且误诊）。入口按
    # 清单一次报全，带专属前缀输出到 stderr 后干净退出，traceback 不外泄；
    # 上层识别前缀即可翻译成「规则集不完整」的友好页，不再误诊稿件。
    missing = missing_required_keys(rules)
    if missing:
        rid = (rules.get("journal") or {}).get("id") or Path(a.ruleset).stem
        for ref in missing:
            # 形态不符的条目自带「路径 形态不符：期望 X」（F13STR），与
            # 「缺少字段」分开措辞——键在但类型错了，不是键不在
            path, sep, expect = ref.partition(" 形态不符：期望 ")
            sec = SECTION_NAMES.get(path.split(".")[0], path.split(".")[0])
            if sep:
                print(f"{RULESET_ERR_PREFIX} 规则集 {rid} 字段 {path} "
                      f"形态不符：期望 {expect}（章节 {sec}）", file=sys.stderr)
            else:
                print(f"{RULESET_ERR_PREFIX} 规则集 {rid} 缺少字段 {ref}"
                      f"（章节 {sec}）", file=sys.stderr)
        raise SystemExit(1)

    src = Path(a.ir).with_suffix("")
    src_docx = Path(a.ir.replace(".ir.json", ".docx"))
    doc = Document(str(src_docx))

    rejects = set()
    if a.reject and Path(a.reject).is_file():
        raw = json.loads(Path(a.reject).read_text(encoding="utf-8"))
        rejects = set(raw.get("ids", []) if isinstance(raw, dict) else raw)

    audit = Audit(rejects)
    body = rules["typography"]["body"]
    cn, latin = body["font_cn"], body.get("font_latin", "Times New Roman")
    size = body["size_pt"]
    # 行距：倍数行距（「1.5 倍」）直接写成 Word 倍数；固定值用 value_pt；
    # 其余倍数若规则集已给出折算依据（equivalent_exact_pt）才折算成固定值，
    # 代码不猜系数
    ls = body["line_spacing"]
    line_mult = ls["value"] if ls.get("mode") == "multiple" else None
    line_pt = None if line_mult else (ls.get("value_pt") or ls.get("equivalent_exact_pt")
                                      or DEFAULTS["body_line_pt"][0])
    # 行距兜底值是代码约定：用到的审计条目要标缺省（F30）
    line_def = [] if (line_mult or ls.get("value_pt") or ls.get("equivalent_exact_pt")) \
        else ["body_line_pt"]
    line_desc = (f"{line_mult:g} 倍行距" if line_mult
                 else f"行距 {line_pt:g}pt" + presumed(line_def))
    indent = body["first_line_indent_chars"]
    hd = rules["headings"]
    lv1 = hd["level1"]

    # 页面设置：与规则集真有差异才出条目。第二遍渲染时版式已合规，
    # 审计不该有一条「无差异」记录——每条目都必须对应一次真实改动。
    # 容差 0.05mm：mm↔twips 往返的舍入误差实测 ≤0.009mm（半 twip 刻度），
    # 而真实设置差异至少 0.1mm，两者差一个数量级以上
    # 页面设置遍历全部节（F32）：只设 sections[0] 的话，分节符造出的其余节
    # 保持 Word 默认版面。任一节不合规即整组设置（幂等：全合规后不再出条目）。
    secs = list(doc.sections)

    def page_diffs():
        def off(cur, want):
            return cur is None or abs(cur.mm - want) > 0.05
        diffs = []
        if any(off(s.page_width, rules["page"]["size_mm"]["width"])
               or off(s.page_height, rules["page"]["size_mm"]["height"])
               for s in secs):
            diffs.append("纸张")
        for side in ("top", "bottom", "left", "right"):
            if any(off(getattr(s, f"{side}_margin"), rules["page"]["margin_mm"][side])
                   for s in secs):
                diffs.append(side)
        return diffs

    if page_diffs():
        for sec in secs:
            sec.page_width = Mm(rules["page"]["size_mm"]["width"])
            sec.page_height = Mm(rules["page"]["size_mm"]["height"])
            for side in ("top", "bottom", "left", "right"):
                setattr(sec, f"{side}_margin", Mm(rules["page"]["margin_mm"][side]))
        mg = rules["page"]["margin_mm"]
        mg_txt = "，".join(f"{cn}{mg[side]:g}" for side, cn in
                           (("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右")))
        audit.add(-1, "page", "页面设置",
                  note=f"纸张 {rules['page']['size']}，页边距 {mg_txt} mm"
                       f"（应用于全部 {len(secs)} 节）",
                  rule="page.margin_mm")

    # 建立 idx → 元素映射。sdt 内容控件整块占一个槽位，其内的段落/表格下钻
    # 展开成列表挂在同一槽位，由块的 sdt_pos 取用（F24，见 elem_of）——槽位
    # 编号与 IR 的 idx 仍同源，不会让按 body 子元素定位的其它环节错位。
    elems = {}
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            elems[i] = Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            elems[i] = Table(child, doc)
        elif child.tag == qn("w:sdt"):
            # 与 parse_docx._sdt_blocks 同一套递归展开（SDTNEST）：内嵌 sdt
            # 继续展开直到取尽段落/表格，未识别叶节点已由解析端计数进 IR
            sdt_els, _unknown = _sdt_blocks(child.find(qn("w:sdtContent")))
            inner = [Paragraph(e, doc) if e.tag == qn("w:p") else Table(e, doc)
                     for e in sdt_els]
            if inner:
                elems[i] = inner

    # 一级标题起始号：introduction 编 1（多数刊）或 0（管理科学学报）
    sec_no = {1: hd.get("intro_number", 1) - 1, 2: 0, 3: 0, 4: 0}
    table_no = 0
    figure_no = 0
    fm_index = 0
    for blk in ir["blocks"]:
        el = elem_of(blk, elems)
        if el is None:
            continue
        # 只数真正取到元素、进过分支判断的块：取不到元素的块没被核对，
        # 把它算进「已核对」就是虚增覆盖率
        audit.blocks_seen += 1
        if blk.get("sdt_pos") is not None:
            # sdt 内块的审计键带「槽.序」：同槽多块的条目 id 仍然唯一、可驳回
            blk = {**blk, "idx": f'{blk["idx"]}.{blk["sdt_pos"]}'}
        p = el
        t = blk["type"]

        if t == "front_matter":
            fm_index += 1
            txt = p.text.strip()
            title_tp = rules["typography"]["title"]
            sub_tp = rules["typography"].get("subtitle")
            s_def = []
            if sub_tp is None:
                # 副标题样式整条未声明：字号字体取代码缺省（F30），要标缺省
                sub_tp = {"font_cn": DEFAULTS["subtitle_font_cn"][0],
                          "size_pt": DEFAULTS["subtitle_size_pt"][0]}
                s_def = ["subtitle_font_cn", "subtitle_size_pt"]
            s_def.append("subtitle_line_mult")   # 副标题行距折算恒为代码约定
            if fm_index == 1:
                t_def = [k for k, f in (("title_size_pt", "size_pt"),
                                        ("title_font_cn", "font_cn")) if f not in title_tp]
                t_def.append("title_line_mult")  # 标题行距折算恒为代码约定
                sz = title_tp.get("size_pt", DEFAULTS["title_size_pt"][0])
                restyle_audited(p, audit, blk["idx"], t, "文章标题套版",
                                f"{sz:g}pt，行距 {round(sz * DEFAULTS['title_line_mult'][0])}pt"
                                + presumed(t_def),
                                "typography.title" + ("~" if t_def else ""), tp=title_tp,
                                cn=title_tp.get("font_cn", DEFAULTS["title_font_cn"][0]),
                                latin=latin, size=sz,
                                align="center", line_pt=round(sz * DEFAULTS["title_line_mult"][0]),
                                bold=False,
                                before_pt=round(sz * 0.2), after_pt=round(sz * 0.25),
                                letter_spacing=title_tp.get("letter_spacing_pt"))
            elif fm_index == 2 and txt.startswith(("——", "—", "-")):
                sz = sub_tp.get("size_pt", DEFAULTS["subtitle_size_pt"][0])
                restyle_audited(p, audit, blk["idx"], t, "副标题套版",
                                f"{sz:g}pt" + presumed(s_def),
                                "typography.subtitle~", tp=sub_tp,
                                cn=sub_tp.get("font_cn", DEFAULTS["subtitle_font_cn"][0]),
                                latin=latin, size=sz,
                                align="center", line_pt=round(sz * DEFAULTS["subtitle_line_mult"][0]),
                                bold=False,
                                after_pt=round(sz * 0.2))
            else:
                # 作者署名：期刊可单独定字体字号（如四号楷体），未定则随正文
                au = rules["typography"].get("authors") or {}
                asz = au.get("size_pt", size)
                al = au.get("align", "center")
                restyle_audited(p, audit, blk["idx"], t, "作者署名套版",
                                f"{au.get('font_cn', cn)} {asz:g}pt {ALIGN_DESC[al]}"
                                + presumed(["authors_line_mult"]),
                                # 该刊没单列 authors 样式时随正文，依据就是正文样式
                                "typography.authors" if au else "typography.body~", tp=au,
                                cn=au.get("font_cn", cn), latin=latin, size=asz,
                                align=al,
                                # 正文用倍数行距时 line_pt 为 None，作者行按自身字号取 1.5 倍
                                line_pt=max(line_pt or 0,
                                            round(asz * DEFAULTS["authors_line_mult"][0])),
                                bold=False,
                                before_pt=round(asz * 0.15), after_pt=round(asz * 0.9))
            continue

        if t == "table":
            table_no += 1
            bp = rules["tables"]["border_pt"]
            tbl_def = "font_size_pt" not in rules["tables"]
            tbl_size = rules["tables"].get("font_size_pt", DEFAULTS["table_font_size_pt"][0])
            before_tbl = table_snapshot(p)
            apply_table(p, bp, tbl_size, cn, latin)
            after_tbl = table_snapshot(p)
            # 本来就合规的三线表不出条目：同刊重跑审计里不再重出这批「改动」
            if before_tbl != after_tbl:
                # 边框分两层比：val/sz 是三线表的「形状」，space/color 是套版时
                # 按规范统一重写的「上色」（间距 0、黑色）；表内 run 除中西文字体、
                # 字号外，set_run 写 ascii 时会一并重写 w:hAnsi。分开比，说明才
                # 能写清这次到底改了什么——只换了颜色的表不该说成「重设为三线表」
                def b_shape(tb):
                    return {k: v[:2] for k, v in tb["borders"].items()}

                def cb_shape(tb):
                    return [tuple(c[:3] for c in cell[0]) for cell in tb["cells"]]

                def fonts_of(tb, part):
                    return [[f[part] for f in cell[1]] for cell in tb["cells"]]

                bits = []
                if (b_shape(before_tbl) != b_shape(after_tbl)
                        or cb_shape(before_tbl) != cb_shape(after_tbl)):
                    n_old = sum(1 for cb, _f in before_tbl["cells"] if cb)
                    bits.append(f"边框重设为三线表：顶/底 {bp['top']:g}/{bp['bottom']:g} 磅、"
                                f"栏目线 {bp['header_bottom']:g} 磅"
                                + (f"，清除原有 {n_old} 处单元格级边框" if n_old else ""))
                elif (before_tbl["borders"] != after_tbl["borders"]
                        or [cell[0] for cell in before_tbl["cells"]]
                        != [cell[0] for cell in after_tbl["cells"]]):
                    bits.append("边框按三线表规范重写：颜色统一为黑色、间距归 0")
                border_bits = bool(bits)
                if fonts_of(before_tbl, 0) != fonts_of(after_tbl, 0):
                    bits.append(f"表内文字改为 {tbl_size:g}pt"
                                + presumed(["table_font_size_pt"] if tbl_def else []))
                elif fonts_of(before_tbl, 1) != fonts_of(after_tbl, 1):
                    bits.append(f"表内西文字体（hAnsi）统一为 {latin}")
                # 表内字号走缺省值时不再挂在 border_pt 的依据下（F30）
                rule = ("tables.border_pt" if border_bits else
                        "tables.font_size_pt" + ("~" if tbl_def else ""))
                audit.add(blk["idx"], t, "表格重设为三线表", note="；".join(bits),
                          rule=rule)
            continue

        if isinstance(p, Table):
            # 公式常被排成「1 行 2 列」的伪表格：左边公式、右边编号。这种结构一改
            # 就散（拆表会把公式对象与编号断开），所以原样保留——但必须在报告里
            # 说明"这里没动"，否则用户看到公式没变，会以为工具漏了。
            audit.add(blk["idx"], t, "公式（伪表格）原样保留",
                      note="1 行表格排版的公式与编号不做改动，避免拆散公式对象与编号的对应关系",
                      changed=False)
            continue

        if t == "figure":
            # 图片段落必须解除固定行距：固定 18 磅的段落会把 12cm 高的内嵌图裁成一条。
            # 末尾的正文套版会给所有剩余块设固定行距，所以这里必须提前拦下、直接跳过。
            before_fmt = fmt_snapshot(p)
            clear_para_formatting(p)
            p.alignment = ALIGN["center"]
            p.paragraph_format.line_spacing = 1.0
            after_fmt = fmt_snapshot(p)
            # 已经合规的图片段不出条目。确有行距差异的才叫「解除固定行距」；
            # 只动了段前/段后距的用「图片段套版」——动作名必须与实际改动相符
            # （FIGACT：命名不实非虚报），note 按 fmt_delta 列出实际差异
            if before_fmt != after_fmt:
                line_diff = (before_fmt["line"], before_fmt["rule"]) \
                    != (after_fmt["line"], after_fmt["rule"])
                why = "固定行距会裁切内嵌图片，" if line_diff else ""
                audit.add(blk["idx"], t,
                          "图片段解除固定行距" if line_diff else "图片段套版",
                          note=why + fmt_delta(before_fmt, after_fmt, "单倍行距并居中"),
                          rule="figures")
            continue

        if t == "references_heading":
            # 文献表标题的写法逐刊不同（「参考文献」/「参 考 文 献」/「﹝参考文献﹞」），
            # 必须按规则集改写；版式默认同一级标题，可用 typography.references_heading 覆盖
            want = next((b.get("label") for b in rules["blocks"]
                         if b["type"] == "references" and b.get("label")), None)
            old = block_text(p)
            if want and old.strip() != want:
                sig = reject_sig(old)
                if audit.blocked(blk["idx"], "文献表标题归一", sig):
                    audit.add(blk["idx"], t, "文献表标题归一", old, old,
                              "已驳回：保留原写法", rejected=True,
                              rule="blocks:references/label", sig=sig)
                elif text_locked(blk):
                    locked_note(audit, blk, "文献表标题归一", "blocks:references/label")
                else:
                    replace_text(p, want)
                    audit.add(blk["idx"], t, "文献表标题归一", old, want, "仅改标题写法，条目未动",
                              rule="blocks:references/label", sig=sig)
            rh = rules["typography"].get("references_heading") or {}
            rh_size = rh.get("size_pt", lv1.get("size_pt") or size)
            rh_cn = rh.get("font_cn", lv1.get("font_cn", cn))
            restyle_audited(p, audit, blk["idx"], t, "文献表标题套版",
                            f"{rh_cn} {rh_size:g}pt {ALIGN_DESC[rh.get('align', lv1.get('align', 'center'))]}",
                            # 该刊没单列这条样式时随一级标题，依据就是一级标题
                            "typography.references_heading" if rh else "headings.level1",
                            tp=rh,
                            cn=rh_cn, latin=latin, line_multiple=line_mult,
                            size=rh_size, align=rh.get("align", lv1.get("align", "center")),
                            line_pt=head_line_pt(lv1, rh_size, line_pt), bold=False)
            continue
        # 作者信息与缺失要件是用户补录后注入的块（apply_details 产出）。
        # 这里只把"语义块"映射到规则集里的样式键，字号字体一律查表，
        # 因此新增期刊仍然不需要改这段代码。
        if t in SEMANTIC:
            key, align, bold = SEMANTIC[t]
            tp = rules["typography"].get(key) or rules["typography"]["body"]
            sz = tp.get("size_pt", size)
            use_latin = latin
            if t == "english_title":
                ten = rules["typography"].get("title_en")
                if ten:            # 期刊若明确规定英文标题字号（如三号 Times），以它为准
                    tp = ten
                    sz = ten.get("size_pt", sz)
                    align = ten.get("align", align)
                    bold = ten.get("bold", bold)
                    use_latin = ten.get("font_latin", latin)
                else:              # 未声明时取标题与副标题中较小的字号，避免英文标题压过中文标题
                    sub = rules["typography"].get("subtitle") or {}
                    sz = min(sz, sub.get("size_pt", sz))
            # 英文标题与英文作者单独给段前距，否则会与中文关键词贴在一起
            before = after = None
            if t == "english_title":
                before, after = round(sz * 0.9), round(sz * 0.15)
            elif t == "english_authors":
                after = round(sz * 0.6)
            sem_rule = ("typography.title_en" if (t == "english_title" and ten)
                        else f"typography.{key}" if rules["typography"].get(key)
                        else "typography.body")
            restyle_audited(p, audit, blk["idx"], t, "补录块套版",
                            f"{tp.get('font_cn', cn)} {sz:g}pt {ALIGN_DESC[align]}",
                            sem_rule, tp=tp,
                            cn=tp.get("font_cn", cn), latin=use_latin, size=sz, align=align,
                            line_multiple=line_mult,
                            line_pt=line_pt, bold=bold, before_pt=before, after_pt=after)
            continue

        if t in ("heading1", "heading2", "heading3", "heading4"):
            level = blk.get("level") or int(t[-1])
            levels = hd.get("levels") or []
            max_level = hd.get("max_level") or len(levels)
            if level > max_level or len(levels) < level:
                # 该刊只排到 X 级：四级标题不硬套成标题，按正文套版并留说明。
                # 不能静默（用户会以为工具漏了），也不能按 title 模板硬编号——
                # 没有目标模板，编号只能原样保留。落到末尾的正文套版。
                audit.add(blk["idx"], t, "标题层级超出本刊",
                          note=f"该刊只排到 {max_level} 级；本段按正文套版，编号不重写",
                          changed=False,
                          rule="headings.max_level" if "max_level" in hd else "headings")
            else:
                old = block_text(p)
                numeric = str(hd.get("numbering", "")).startswith("numeric")
                # 剥前缀要认"稿子原来是什么编号"，不能按目标刊的体系去剥：
                # 中文稿投英文刊时，源码是「一、」，目标是「1.」，按目标剥会剥不掉。
                # 但目标刊自己的写法也必须认（strip_number_prefix 第三条候选）：
                # 本工具上一轮输出的就是那个形式，不认就叠成「1.1.标题」。
                core, _stripped = strip_number_prefix(old.strip(), level, hd)
                core_norm = re.sub(r"\s", "", core)

                if numeric:
                    # 编号模板取自 levels：写「1」得「1 引言」（中国管理科学），
                    # 写「1.」得「1. Introduction」（英文刊）。数字按层级依次替换。
                    if level == 1:
                        sec_no[1] += 1
                        sec_no[2] = sec_no[3] = sec_no[4] = 0
                        nums = [sec_no[1]]
                    elif level == 2:
                        sec_no[2] += 1
                        sec_no[3] = sec_no[4] = 0
                        nums = [sec_no[1], sec_no[2]]
                    elif level == 3:
                        sec_no[3] += 1
                        sec_no[4] = 0
                        nums = [sec_no[1], sec_no[2], sec_no[3]]
                    else:
                        sec_no[4] += 1
                        nums = [sec_no[1], sec_no[2], sec_no[3], sec_no[4]]
                    it = iter(nums)
                    head = re.sub(r"\d+", lambda m: str(next(it)), levels[level - 1], count=len(nums))
                    new = f"{head} {core}"
                elif level == 1:
                    if core_norm in [re.sub(r"\s", "", s) for s in hd.get("unnumbered_sections", [])]:
                        new = hd.get("intro_title_text", core)
                    else:
                        sec_no[1] += 1
                        sec_no[2] = sec_no[3] = sec_no[4] = 0   # 下级编号按一级标题重新起算
                        new = fill_level_template(levels[0], sec_no[1]) + core
                elif level == 2:
                    sec_no[2] += 1
                    sec_no[3] = sec_no[4] = 0
                    new = fill_level_template(levels[1], sec_no[2]) + core
                elif level == 3:
                    sec_no[3] += 1
                    sec_no[4] = 0
                    new = fill_level_template(levels[2], sec_no[3]) + core
                else:
                    sec_no[4] += 1
                    new = fill_level_template(levels[3], sec_no[4]) + core

                if new != old:
                    sig = reject_sig(old)
                    if audit.blocked(blk["idx"], "标题重新编号", sig):
                        audit.add(blk["idx"], t, "标题重新编号", old, new,
                                  "已驳回：保留原编号", rejected=True,
                                  rule="headings.numbering", sig=sig)
                    elif text_locked(blk):
                        # 标题里挂 SEQ 域（Word 自动章节号）时编号归域管，
                        # 改字面文字会在 F9 更新时被域盖回去
                        locked_note(audit, blk, "标题重新编号", "headings.numbering")
                    else:
                        replace_text(p, new)
                        audit.add(blk["idx"], t, "标题重新编号", old, new,
                                  "仅改编号前缀，标题文字未动",
                                  rule="headings.numbering", sig=sig)
                if level == 1:
                    audit.sec = core  # 后续改动都归到这个一级标题下，审计里才能定位
                # 二三级标题多数刊要求黑体，此前一律继承正文字体；字体与字号
                # 由 rules.headings.levelN 声明，未声明才随正文
                lv = lv1 if level == 1 else (hd.get(f"level{level}") or {})
                inline = level >= 4 and bool(lv.get("inline_with_body"))
                lv_cn = lv.get("font_cn") or cn
                lv_size = lv.get("size_pt") or size
                lv_al = lv.get("align") or ("center" if level == 1 else "left")
                if inline:
                    # 行内四级标题：按正文套版，只重写编号前缀。行距、缩进、
                    # 加粗都跟正文走，不能把整段正文改成标题的黑体/居中。
                    restyle_audited(p, audit, blk["idx"], t, "标题套版",
                                    f"{cn} {size:g}pt，{ALIGN_DESC[lv_al]}"
                                    "（行内四级，按正文套版）",
                                    f"headings.level{level}~",
                                    tp=lv,
                                    cn=cn,
                                    latin=latin, size=size,
                                    align=lv_al,
                                    line_pt=line_pt,
                                    line_multiple=line_mult,
                                    indent=indent,
                                    bold=False,
                                    drop_numbering=True)
                else:
                    restyle_audited(p, audit, blk["idx"], t, "标题套版",
                                    f"{lv_cn} {lv_size:g}pt，{ALIGN_DESC[lv_al]}",
                                    # level1 是必填键；二至四级未声明的刊随正文约定，标可缺省
                                    "headings.level1" if level == 1 else f"headings.level{level}~",
                                    tp=lv,
                                    cn=lv_cn,
                                    latin=latin, size=lv_size,
                                    align=lv_al,
                                    line_pt=head_line_pt(lv, lv_size, line_pt),
                                    line_multiple=line_mult,
                                    indent=lv.get("indent_chars") if level > 1 else None,
                                    bold=lv.get("bold", False),
                                    drop_numbering=True)   # 编号已重写为字面文字，自动编号必须摘掉
                continue

        if t == "title":
            # parse_docx 不会产出 title 块（标题一律走 front_matter 分支），
            # 此分支是给外部直接构造 IR 时的防御，当前主流程不可达。
            tp_t = rules["typography"]["title"]
            t_def = [k for k, f in (("title_size_pt", "size_pt"),
                                    ("title_font_cn", "font_cn")) if f not in tp_t]
            t_def.append("title_line_mult")      # 标题行距折算恒为代码约定
            sz = tp_t.get("size_pt", DEFAULTS["title_size_pt"][0])
            restyle_audited(p, audit, blk["idx"], t, "文章标题套版",
                            f"{sz:g}pt，行距 {round(sz * DEFAULTS['title_line_mult'][0])}pt"
                            + presumed(t_def),
                            "typography.title" + ("~" if t_def else ""), tp=tp_t,
                            cn=tp_t.get("font_cn", DEFAULTS["title_font_cn"][0]), latin=latin,
                            size=sz, align="center",
                            line_pt=round(sz * DEFAULTS["title_line_mult"][0]), bold=False,
                            before_pt=round(sz * 0.2), after_pt=round(sz * 0.25),
                            letter_spacing=tp_t.get("letter_spacing_pt"))
            continue

        if t in ("abstract", "keywords"):
            # 段首手工缩进空白先删（N-01 段首＋N-08）：不删的话，段首空格会和
            # 下面的首行缩进设置叠成双重缩进，标签归一的正则也认不出段首的标签
            strip_leading_indent_audited(p, audit, blk["idx"], t, indent)
            want = next((b.get("label") for b in rules["blocks"]
                         if b["type"] == t and b.get("lang") == "zh" and b.get("label")), None)
            old = block_text(p)
            if want:
                m = re.match(r"^(内容摘要|内容提要|摘\s*要|提\s*要|关\s*键\s*词)\s*[:：]", old)
                if m and not old.startswith(want):
                    new = want + old[m.end():]
                    sig = reject_sig(old)
                    if audit.blocked(blk["idx"], "标签归一", sig):
                        audit.add(blk["idx"], t, "标签归一", old, new,
                                  "已驳回：保留原标签", rejected=True,
                                  rule=f"blocks:{t}/zh/label", sig=sig)
                    else:
                        # 只换前缀，不整段重写：标签的加粗不该传染给整段摘要
                        replace_prefix(p, m.end(), want)
                        audit.add(blk["idx"], t, "标签归一", old, new,
                                  "仅改标签前缀，正文文字未动",
                                  rule=f"blocks:{t}/zh/label", sig=sig)
            # 关键词分隔符归一：各刊要求「；」「，」或空格，稿件里常是混的
            sep = next((b.get("separator") for b in rules["blocks"]
                        if b["type"] == "keywords" and b.get("lang") == "zh"), None)
            if t == "keywords" and sep and want and block_text(p).startswith(want):
                before_kw = block_text(p)
                items = split_keywords(before_kw[len(want):], sep)
                new_kw = want + sep.join(items)
                if new_kw != before_kw and len(items) >= 2:
                    sig = reject_sig(before_kw)
                    if audit.blocked(blk["idx"], "关键词分隔符归一", sig):
                        audit.add(blk["idx"], t, "关键词分隔符归一", before_kw, new_kw,
                                  "已驳回：保留原分隔符", rejected=True,
                                  rule="blocks:keywords/zh/separator", sig=sig)
                    else:
                        # 同理：只重写标签之后的正文，标签本身的加粗不能扩散
                        replace_after(p, len(want), sep.join(items))
                        # after 重读段落实际文字（与文献编号路径同法）：
                        # 审计记的必须是落盘结果，不是替换前的期望值
                        audit.add(blk["idx"], t, "关键词分隔符归一", before_kw, block_text(p),
                                  f"改为「{sep}」，共 {len(items)} 个关键词，词条文字未动",
                                  rule="blocks:keywords/zh/separator", sig=sig)
            # 摘要与关键词常各自定字体（统计研究：提要宋体、关键词黑体），必须分别取
            key = ("abstract_en" if t == "abstract" else "keywords_en") if blk.get("lang") == "en" \
                else ("abstract" if t == "abstract" else "keywords")
            absp = rules["typography"].get(key) or rules["typography"].get("abstract") or {}
            abs_own_line = (absp.get("line_spacing") or {}).get("value_pt") \
                or (absp.get("line_spacing") or {}).get("equivalent_exact_pt")
            abs_ls = abs_own_line or line_pt
            abs_rule = (f"typography.{key}" if rules["typography"].get(key)
                        else "typography.abstract" if rules["typography"].get("abstract")
                        else "typography.body")
            if not abs_own_line and line_def:
                abs_rule += "~"                  # 行距落到了正文缺省值上（F30）
            abs_line = (f"行距 {abs_ls:g}pt" + (presumed(line_def) if not abs_own_line else "")
                        if abs_ls else line_desc)
            abs_target = (f"{absp.get('font_cn', cn)} {absp.get('size_pt', size):g}pt，"
                          + abs_line)
            restyle_audited(p, audit, blk["idx"], t,
                            "摘要套版" if t == "abstract" else "关键词套版",
                            abs_target, abs_rule, tp=absp,
                            cn=absp.get("font_cn", cn), latin=latin, line_multiple=line_mult,
                            size=absp.get("size_pt", size), align="both", line_pt=abs_ls,
                            indent=indent)
            continue

        if t == "author_note":
            # 文末作者简介：标签一段、每人一段，字体字号单独声明
            an = rules["typography"].get("author_note") or {}
            an_own_line = (an.get("line_spacing") or {}).get("value_pt")
            an_def = [] if an_own_line else ["author_note_line_mult"]
            sz = an.get("size_pt", size)
            an_ls = an_own_line or round(sz * DEFAULTS["author_note_line_mult"][0], 1)
            lab = (next((b.get("label") for b in rules["blocks"]
                         if b["type"] == "author_note"), "") or "作者简介").strip()
            an_rule = ("typography.author_note" if an else "typography.body") \
                + ("~" if an_def else "")
            if p.text.strip() == lab:
                restyle_audited(p, audit, blk["idx"], t, "作者简介套版",
                                f"{an.get('label_font_cn', DEFAULTS['author_note_label_font_cn'][0])}"
                                f" {sz:g}pt" + presumed(an_def),
                                an_rule, tp=an,
                                cn=an.get("label_font_cn",
                                          DEFAULTS["author_note_label_font_cn"][0]),
                                latin=latin, size=sz,
                                align=an.get("heading_align", "left"), line_pt=an_ls,
                                indent=an.get("heading_indent_chars") or indent)
            else:
                restyle_audited(p, audit, blk["idx"], t, "作者简介套版",
                                f"{an.get('font_cn', cn)} {sz:g}pt" + presumed(an_def),
                                an_rule, tp=an,
                                cn=an.get("font_cn", cn), latin=latin, size=sz,
                                align="both", line_pt=an_ls, indent=indent)
            continue

        if t in ("reference_item",):
            rl = rules["citation"].get("reference_list", {})
            hang = rl.get("hanging_indent_chars")

            # 编号前缀按目标刊改写法（[1] / 〔1〕 / 1．）。只动前缀——条目正文里的
            # 作者、年份、标点一律不动，那属于改写内容，不归这个产品管。
            old_prefix, num = ref_number_prefix(block_text(p))
            if old_prefix is not None:
                new_prefix = format_ref_number(rl.get("number_format"), num,
                                               rl.get("space_after_number"))
                # 比较时忽略尾随空白：前缀区含空白是为了「整段替换」不叠空格，
                # 但该刊声明 space_after=None 时新前缀本就不带空白——`[1] ` 与
                # `[1]` 是同一个编号，不能因为差一个空格就反复重写、每跑一次
                # 吃掉一个空格
                if new_prefix and new_prefix.rstrip() != old_prefix.rstrip():
                    before_txt = block_text(p)
                    sig = reject_sig(before_txt)
                    if audit.blocked(blk["idx"], "文献条目编号归一", sig):
                        audit.add(blk["idx"], t, "文献条目编号归一", before_txt, before_txt,
                                  "已驳回：保留原编号写法", rejected=True,
                                  rule="citation.reference_list.number_format", sig=sig)
                    elif text_locked(blk):
                        # EndNote 文献表是一整个 EN.REFLIST 域：条目编号是该域的
                        # 显示结果，改了前缀，用户 F9 更新就整表回滚
                        locked_note(audit, blk, "文献条目编号归一",
                                    "citation.reference_list.number_format")
                    else:
                        replace_prefix(p, len(old_prefix), new_prefix)
                        audit.add(blk["idx"], t, "文献条目编号归一", before_txt, block_text(p),
                                  f"编号写法改为 {new_prefix.strip()}，条目内容未动",
                                  rule="citation.reference_list.number_format", sig=sig)

            ref_tp = rules["typography"].get("references") or {}
            r_size = ref_tp.get("size_pt", size)
            # 参考文献行距逐刊不同（经济研究 14.2、统计研究 12.3），不能沿用正文行距
            ref_ls = ref_tp.get("line_spacing") or {}
            ref_own_line = ref_ls.get("value_pt") or ref_ls.get("equivalent_exact_pt")
            ref_line = ref_own_line or line_pt
            before_fmt = fmt_snapshot(p)
            restyle(p, cn=ref_tp.get("font_cn", cn), latin=latin, size=r_size, line_multiple=line_mult,
                    align="both", line_pt=ref_line, indent=None, bold=False)
            if ref_tp.get("line_single"):
                # 规范要求「单倍行距」时不能沿用正文的固定行距，否则六号字会被撑开
                p.paragraph_format.line_spacing = 1.0
            if hang:
                set_hanging(p, hang)
            elif rl.get("indent_chars"):
                # 部分期刊要求首行缩进而非悬挂缩进
                set_chars_indent(p, rl["indent_chars"])
            after_fmt = fmt_snapshot(p)
            # 缩进有自己的规则集字段，单独一条；ind 已经是目标值时不再报
            # （重跑合规稿不再重出这批条目）
            ind_dedicated = False
            if before_fmt["ind"] != after_fmt["ind"]:
                if hang:
                    audit.add(blk["idx"], t, "参考文献悬挂缩进", after=f"{hang:g} 字符",
                              rule="citation.reference_list.hanging_indent_chars")
                    ind_dedicated = True
                elif rl.get("indent_chars"):
                    audit.add(blk["idx"], t, "参考文献缩进",
                              after=f"{rl['indent_chars']:g} 字符",
                              rule="citation.reference_list.indent_chars")
                    ind_dedicated = True
            if ind_dedicated:
                before_fmt["ind"] = after_fmt["ind"]   # 缩进已单列，套版对比不再计
            if before_fmt != after_fmt:
                if ref_tp.get("line_single"):
                    ref_desc = "单倍行距"
                elif ref_line:
                    ref_desc = (f"行距 {ref_line:g}pt"
                                + (presumed(line_def) if not ref_own_line else ""))
                else:
                    ref_desc = line_desc
                ref_rule = ("typography.references" if ref_tp else "typography.body") \
                    + ("~" if (not ref_own_line and line_def) else "")
                audit.add(blk["idx"], t, "参考文献条目套版",
                          note=fmt_delta(before_fmt, after_fmt,
                                         f"{ref_tp.get('font_cn', cn)} {r_size:g}pt，{ref_desc}"),
                          rule=ref_rule)
            continue

        if t in ("table_caption", "figure_caption"):
            figure_no += t == "figure_caption"
            size_cap = rules["tables"].get("caption_size_pt", size)
            cap_font_def = "caption_font_cn" not in rules["tables"]
            font_cap = rules["tables"].get("caption_font_cn", DEFAULTS["caption_font_cn"][0])
            # 「表1」与标题之间的间隔由规则集给出（一字空/半字空…），不再写死两空格
            gap_kind = (rules["tables"].get("caption_gap")
                        or (rules.get("figures") or {}).get("caption_gap")
                        or DEFAULTS["caption_gap"][0])
            pad = CAPTION_GAP.get(gap_kind, "  ")
            gap_rule = ("tables.caption_gap" if "caption_gap" in rules["tables"]
                        else "figures.caption_gap" if "caption_gap" in (rules.get("figures") or {})
                        else "tables.caption_gap~")
            # 稿件里的表题常用制表符把「表1」和标题撑开，需按目标刊体例归一
            m = re.match(r"^(表|图)\s*(\d+)\s*(.*)$", block_text(p).strip())
            if m:
                old_cap = block_text(p)
                new_cap = f"{m.group(1)}{m.group(2)}{pad}{m.group(3)}".rstrip()
                if new_cap != old_cap:
                    sig = reject_sig(old_cap)
                    if audit.blocked(blk["idx"], "题注归一", sig):
                        audit.add(blk["idx"], t, "题注归一", old_cap, new_cap,
                                  "已驳回：保留原题注", rejected=True, rule=gap_rule, sig=sig)
                    elif text_locked(blk):
                        # 「表1」的 1 常是 SEQ 域：改字面间隔，F9 更新后域只重出
                        # 数字，间隔照样丢，还可能把域结果变成死文本
                        locked_note(audit, blk, "题注归一", gap_rule)
                    else:
                        replace_text(p, new_cap)
                        audit.add(blk["idx"], t, "题注归一", old_cap, new_cap,
                                  f"编号与标题间隔改为「{gap_kind}」，题注文字未动",
                                  rule=gap_rule, sig=sig)
            cap_def = ((["caption_font_cn"] if cap_font_def else [])
                       + (["body_line_pt"] if line_def else []))
            cap_rule = ("tables.caption_size_pt" if "caption_size_pt" in rules["tables"]
                        else "tables.caption_size_pt~")
            if cap_def:
                cap_rule = "tables.caption_size_pt~"   # 用到缺省值的条目标可缺省（F30）
            restyle_audited(p, audit, blk["idx"], t, "题注套版",
                            f"{font_cap} {size_cap:g}pt 居中" + presumed(cap_def),
                            cap_rule,
                            cn=font_cap, latin=latin, size=size_cap, align="center",
                            line_multiple=line_mult, line_pt=line_pt, bold=False)
            continue

        if t == "table_note":
            _nl = rules["tables"].get("note_layers") or [{}]
            note_def = ["note_size_pt"] if (not _nl[0] or "size_pt" not in _nl[0]) else []
            note_sz = _nl[0].get("size_pt", DEFAULTS["note_size_pt"][0])
            restyle_audited(p, audit, blk["idx"], t, "表注套版",
                            f"{cn} {note_sz:g}pt 左对齐" + presumed(note_def),
                            "tables.note_layers"
                            + ("~" if (note_def or not rules["tables"].get("note_layers"))
                               else ""),
                            tp=_nl[0],
                            cn=cn, latin=latin, size=note_sz, align="left", line_pt=line_pt,
                            line_multiple=line_mult)
            continue

        # 独立公式段落：行距用「最小值」，不能用「固定值」。固定值会把分式、
        # 积分、上下标这类向上向下伸展的公式裁掉——和图片被裁是同一类事故。
        # 对齐方式保持原样，不强制居中：短段落里含行内公式时，解析器也会判为
        # equation，强行居中会把普通正文段落挪位置。
        if t == "equation":
            before_fmt = fmt_snapshot(p)
            restyle(p, cn=cn, latin=latin, size=size, align=None, line_pt=line_pt, line_multiple=line_mult,
                    line_rule="at_least")
            after_fmt = fmt_snapshot(p)
            # 已经是「最小值」行距的公式段不再报
            if before_fmt != after_fmt:
                tgt = ("行距随正文倍数设置" if line_mult
                       else f"行距最小值 {line_pt:g}pt" + presumed(line_def))
                audit.add(blk["idx"], t, "公式段行距",
                          note=fmt_delta(before_fmt, after_fmt, tgt)
                          + "：固定值会裁掉分式与上下标",
                          rule="typography.body" + ("~" if line_def else ""))
            continue

        # 正文、公式、伪表格、英文区等一律按正文字体处理。
        # 套版本身属 A 档（不动文字），但仍是「改动」，必须进报告——
        # 否则报告只记标签、编号这些小事，最显眼的正文重排反倒没有记录。
        if t == "body":
            # 段首手工缩进空白删除（N-01 段首＋N-08）：不删就会与下面的首行
            # 缩进设置叠成双重缩进
            strip_leading_indent_audited(p, audit, blk["idx"], t, indent)
        before_fmt = fmt_snapshot(p)
        restyle(p, cn=cn, latin=latin, size=size, align="both", line_pt=line_pt, indent=indent,
                line_multiple=line_mult)
        after_fmt = fmt_snapshot(p)
        if before_fmt != after_fmt:
            audit.add(blk["idx"], t, "正文套版",
                      note=fmt_delta(before_fmt, after_fmt,
                                     f"{cn} {size:g}pt，{line_desc}"),
                      rule="typography.body" + ("~" if line_def else ""))

    reorder_notes(elems, ir, rules, audit)

    # 引用体系转换放在最后：前面已经把编号归一成 [1] 形式，这里才好定位
    if a.citation:
        apply_citation(doc, elems, ir, rules, audit)

    out = Path(a.out)

    fn_count, fn_desc, fn_external = style_footnotes(doc, rules)
    if fn_count:
        audit.add(-2, "footnote", "脚注正文套版",
                  note=f"{fn_count} 条脚注改为 {fn_desc}",
                  rule="typography.footnote")
    if fn_external:
        # 外部脚注关系没有可解析部件，套版与体检都够不着（EXTFOOT）：
        # 跳过了多少条必须进审计文件——下游只读 *.audit.json，看不到 stdout
        # 汇总（EXTFOOTAUD），不能让「这里没处理」只留在屏幕上
        audit.add(-2, "footnote", "外部脚注关系未处理",
                  note=f"{fn_external} 条外部脚注关系未处理（TargetMode=External"
                       f" 的外部超链接式脚注没有可解析部件，套版与体检都够不着，"
                       f"已原样保留）",
                  changed=False)

    doc.save(str(out))
    # 正文里的作者信息归 anonymize.py 管；这里清掉 Word 看不见的文档属性：
    # 作者、最后修改者、公司，以及 WPS 写进 custom.xml 的账号标识。
    sanitize_docx(out)

    audit_path = out.with_suffix(".audit.json")
    # 推定项的区块归属（F16NOISE）：assumptions 保持原数组（逐条完整落盘，
    # 展示分层不许删披露），另存一份「条目 → 区块」映射，审计页据此分组折叠。
    # 键是 assumption 全文；规则集没带 extraction 留痕时映射全落「其他推定」。
    audit_path.write_text(json.dumps({"source": ir["source"], "target": rules["journal"]["name_zh"],
                                      "items": audit.items,
                                      "assumptions": rules.get("assumptions") or [],
                                      "assumption_sections": assume_groups.section_map(
                                          rules.get("assumptions"),
                                          (rules.get("verification") or {}).get("extraction")),
                                      "rejected": sorted(rejects),
                                      # 覆盖率：核对过多少块、改动落在多少块上。
                                      # 两个数一起给，「没改的块是看过还是漏了」
                                      # 才有答案（见 render_audit 的统计条）。
                                      # 负数 idx 是页面/脚注/引用这类非块条目，
                                      # 数进来会让改动块数超过核对块数。
                                      "checked_blocks": audit.blocks_seen,
                                      "changed_blocks": len({it["idx"] for it in audit.items
                                                             if it.get("changed", True)
                                                             and not it.get("rejected")
                                                             and not str(it["idx"]).startswith("-")}),
                                      # 域统计：报告要能说清「这段文字由谁掌管」
                                      "fields": ir.get("fields") or {}},
                                     ensure_ascii=False, indent=2),
                          encoding="utf-8")

    print(f"已生成: {out}")
    print(f"审计条目: {len(audit.items)} 条 → {audit_path.name}")
    if rejects:
        print(f"已驳回: {len(rejects)} 条改动（保留原文）")
    kinds = {}
    for it in audit.items:
        kinds[it["action"]] = kinds.get(it["action"], 0) + 1
    for k, v in sorted(kinds.items(), key=lambda x: -x[1]):
        print(f"   {v:4d}  {k}")
    if ir.get("unsupported_content_blocks"):
        # 解析时记数进 IR 的仍未处理的内容控件结构（F24/SDTNEST）：渲染汇总
        # 必须报出来，不能让「这里没处理」变成静默的
        print(f"未处理内容块: {ir['unsupported_content_blocks']} 个"
              f"（内容控件里未识别的块级节点、取空或嵌套超过 "
              f"{SDT_MAX_DEPTH} 层的内容控件，已原样保留）")
    if fn_external:
        # 外部脚注关系没有可解析部件，套版与体检都够不着（EXTFOOT）：
        # 跳过了多少条必须说明，不能静默
        print(f"{fn_external} 条外部脚注关系未处理"
              f"（外部超链接式脚注没有可解析部件，已原样保留）")


if __name__ == "__main__":
    main()
