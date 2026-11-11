"""把真实稿件解析成论文中间表示（Paper IR）。

设计依据来自两份真实稿件的结构普查：
1. 96%~100% 的段落没有使用 Word 样式，全靠直接格式 —— 因此完全不能依赖 pStyle
2. 自动编号 numPr 为 0 —— 标题编号全是手工敲的文本，只能靠文本模式 + 上下文识别
3. 表格边框可能定义在 tblBorders，也可能只定义在 tcBorders —— 两处都要查
4. 存在"伪表格"：用 1 行 2 列的表格排公式与编号
5. 公式可能是 OMML 对象，也可能是一行纯文本

规则识别之后还有一个可选的 AI 补漏环节（DESIGN §14）：AI 只能给块改标签，
不能改文字，且只对规则自己不确定的块生效。指纹守卫不通过则整批作废。

用法：python tools/parse_docx.py <file.docx> [-o out.json] [--ai | --no-ai]
"""

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ai_assist  # noqa: E402  同目录模块，AI 不可用时内部自行降级
import xmlsafe  # noqa: E402  同目录模块

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

MAX_HEADING_LEN = 40

# AI 可以"补漏"的起点：规则判不出来时才落到 body，front_matter 更是只按位置猜的
AI_UPGRADE_FROM = {"body", "front_matter"}

# 标签括号体例（全库实有：【摘要】〔…〕﹝参考文献﹞）：闭合括号本身就是
# 标签与正文的分隔，冒号/空白不再是必要条件（IDMPOT：解析必须能读回渲染器
# 按规则集写出的每一种标签形态，否则二次解析整段退化成正文、重跑即漂移）
_LABEL_BRA = "[【〔﹝\\[（(]"
_LABEL_KET = "[】〕﹞\\]）)]"

PAT = {
    # 内容提要／提要 是「内容摘要／摘要」的常见变体（《经济研究》即用「内容提要」）。
    # 这里必须与 render_from_ir 的标签归一正则保持一致，否则解析认不出、
    # 渲染就进不了摘要分支：字号不套、字数不查、标签也不归一。
    # 标签后允许「空白」代替冒号（IDMPOT）：有的刊体例标签后不加冒号、直接接
    # 全角空格（如《世界经济》式「内容提要　」），渲染器按规则集写出的就是
    # 这个形态——二次解析认不出它，摘要/关键词整段退化成正文，重跑即漂移。
    # 括号体例（【摘要】）由闭合括号充当分隔。无分隔的「内容提要本文」仍不当摘要。
    "abstract": re.compile(
        r"^(?:[" + _LABEL_BRA[1:-1] + r"]\s*(?:内容摘要|内容提要|摘\s*要|提\s*要)\s*"
        r"[" + _LABEL_KET[1:-1] + r"]"
        r"|(?:内容摘要|内容提要|摘\s*要|提\s*要)(?:\s*[:：]|\s))"),
    "keywords": re.compile(
        r"^(?:[" + _LABEL_BRA[1:-1] + r"]\s*关\s*键\s*词\s*[" + _LABEL_KET[1:-1]
        + r"]|关\s*键\s*词(?:\s*[:：]|\s))"),
    "classification": re.compile(r"^(中图分类号|JEL分类号|JEL|文献标识码)"),
    # 文献表标题必须能读回渲染器写出的全部体例：「参考文献：」「参 考 文 献」
    # 「﹝参考文献﹞」（ciejournal）与英文刊的「References」（cwe 等，re.I
    # 覆盖 REFERENCES），否则二次解析时文献区整段失效、条目全部退化成正文
    # （IDMPOT）。
    "ref_head": re.compile(
        r"^(?:[" + _LABEL_BRA[1:-1] + r"]\s*参\s*考\s*文\s*献\s*[" + _LABEL_KET[1:-1]
        + r"]|参\s*考\s*文\s*献\s*[:：]?|references)\s*[:：]?$", re.I),
    "table_caption": re.compile(r"^表\s*\d+"),
    "figure_caption": re.compile(r"^图\s*\d+"),
    "note": re.compile(r"^(注|资料来源|数据来源)\s*[:：]"),
    "h1": re.compile(r"^[一二三四五六七八九十]{1,3}、"),
    # 「一　引言」：中文数字＋空白分隔的一级体例（《世界经济》等，levels[0]=「一　」）。
    # 渲染器按规则集重写出的就是这个形态，二次解析认不出它标题就退化成正文
    # （IDMPOT）。分隔空白必须存在（顿号体例由 h1 管）。护栏与 h1_num 同一口径
    # （短、无句读、分类词负向断言，见下）：只有「短＋无句读」挡不住「一 些
    # 研究表明」「一 种新的方法」这类把数量词拆开写的正文短句——它们没有
    # 句读、也够短，而判成一级是 high 置信，AI 不得纠正（DESIGN §14.2），
    # 必须在正则处排除。负向断言与 h1_num 的分类词清单一致，另加「些」——
    # 「一些」被空格拆开是中文正文的高频形态，节名不会以「些」起头。
    "h1_ws": re.compile(
        r"^[一二三四五六七八九十]{1,3}\s+(?![个项种类条位名年月日次篇本些])\S"),
    "h2": re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]"),
    "h3": re.compile(r"^\d{1,2}\s*[.．](?!\d)\s*\S"),
    "h4": re.compile(r"^[（(]\d{1,2}[）)]"),
    # 阿拉伯数字分级、编号后不加点号：「1 引言」「2.1 研究方法」。
    # 这是 Word 默认多级列表的另一种写法，与上面的「1. 引言」并存，
    # 中文期刊稿件里非常常见，漏判会让整篇标题退化成正文。
    # 否定量词的断言是防「3 个案例」这种正文短句冒充一级标题。
    "h1_num": re.compile(r"^\d{1,2}\s+(?![\d.．个项种类条位名年月日次篇本])"),
    "h2_num": re.compile(r"^\d{1,2}[.．]\d{1,2}\s+\S"),
    "h3_num": re.compile(r"^\d{1,2}[.．]\d{1,2}[.．]\d{1,2}\s+\S"),
    "ref_item_num": re.compile(r"^\[\d+\]"),
    "appendix": re.compile(r"^(在线)?附\s*录"),
    "en_abstract": re.compile(r"^Abstract\s*[:：]", re.I),
    "en_keywords": re.compile(r"^Key\s*words?\s*[:：]", re.I),
    "en_section": re.compile(r"^(Abstract|Key\s*words|参考文献|References)", re.I),
}

YEAR = re.compile(r"(19|20)\d{2}")

# 「1.」「2.」这类单级阿拉伯编号在中文期刊里是**两级歧义**的：
#   《经济管理》《南开管理评论》体例里它是二级（一级「一、」）
#   「一、／（一）／1.」体系里它是三级
# 仅凭文本无法定级。硬编码成三级，会让前一类刊的合规投稿稿整篇二级标题
# 被降成正文（`level==3 and 2 not in seen_level` 的降级保护正好触发），
# 首轮就把标题格式毁掉，且重跑不幂等——同一份成品转两次，第二次报一批
# 「改对齐」，标题被套成正文样式。因此这里只返回候选标记，由 parse()
# 按「本文档已出现过哪几级」定级。
AMBIGUOUS_H3 = "2or3"

# 标题不用句读标点。数字编号标题全靠这一条与正文短句区分（如「3 个案例」）。
# 冒号不算在内——「3 分析框架：科研自主性分级」是正常标题写法。
HEADING_PUNCT = re.compile(r"[。；，！？]")

# 不带编号的一级标题。不少刊的一级标题不编号（规则集 unnumbered_sections），
# 渲染器会把「一、引言」改写成「引  言」——这个形式必须能被读回来，否则重跑时
# 它认不出是标题，落进「前三段即前置区」的按位置推断里，被当成作者署名套版。
# 同一份稿子越跑越错，就是这么来的。
# 收紧到「整段只有这几个词、字间可有空白」：正文里不会单独成段，误判风险低。
UNNUMBERED_H1 = re.compile(r"^(引\s*言|绪\s*论|结\s*论|结\s*语|余\s*论|致\s*谢|后\s*记)$")


def looks_like_ref(text):
    """参考文献条目几乎必然含年份；用它作为退出参考文献区的软判据。

    真实反例：参考文献之后紧跟「在线附录」并继续出现表 8–12，
    若状态不复位，附录内容会被整段吞进参考文献。
    """
    if PAT["ref_item_num"].match(text):
        return True
    return len(text) > 15 and bool(YEAR.search(text))


def para_text(p):
    """段落聚合文本；w:txbxContent（文本框）子树不计。

    文本框文字不属于段落：混进来会污染标题/摘要/表格题注这些分类判据
    （F24），而文本框对象本身原样保留、不参与分类与改写。mc:AlternateContent
    的 Choice/Fallback 双份同文都在各自的 txbxContent 里，一并排除。
    """
    parts = []

    def walk(node, depth=1):
        if depth > XML_MAX_DEPTH:
            # DEPTHCAP：parse() 入口已先整树把关，正常到不了这里；
            # 兜住绕过入口直接拿深树调本函数的路径（F2-2）
            raise ValueError(f"段落 XML 嵌套超过 {XML_MAX_DEPTH} 层，疑似异常文档，已拒收")
        for child in node:
            if child.tag == W + "txbxContent":
                continue
            if child.tag == W + "t":
                parts.append(child.text or "")
            walk(child, depth + 1)

    walk(p)
    return "".join(parts).strip()


def has_omml(p):
    return next(p.iter(M + "oMath"), None) is not None


# ---------------------------------------------------------------- Word 域

# 域类型识别。顺序即优先级：ADDIN 必须先于通用前缀判断，
# 因为文献管理域的指令开头是「ADDIN EN.CITE ...」，不是任何关键字。
FIELD_PATTERNS = (
    # 文献管理软件插入的引文／文献表：这是唯一一类「文字不归作者也不归我们」的内容
    ("citation", re.compile(r"^\s*ADDIN\s+(?:EN\.|ZOTERO|MENDELEY|CNKI|NOTE|REFMAN|CITATION)", re.I)),
    # 图表自动编号，改动它会与题注文字脱钩，但重建时按目标刊重编，属正常业务
    ("seq", re.compile(r"^\s*SEQ\b", re.I)),
    # 交叉引用与目录：指向别处，拆平后编号变死文本
    ("ref", re.compile(r"^\s*(?:REF|NOTEREF|PAGEREF)\b", re.I)),
    ("toc", re.compile(r"^\s*TOC\b", re.I)),
    # 页眉页脚类，重排不涉及
    ("page", re.compile(r"^\s*(?:PAGE|NUMPAGES|SECTIONPAGES|STYLEREF)\b", re.I)),
    ("eq", re.compile(r"^\s*EQ\b", re.I)),
    ("date", re.compile(r"^\s*(?:DATE|TIME|SAVEDATE|PRINTDATE)\b", re.I)),
    ("other", re.compile(r".*")),
)

FIELD_INSTR_MAX = 120
FIELD_RESULT_MAX = 200


def field_kind(instr):
    for kind, rx in FIELD_PATTERNS:
        if rx.match(instr or ""):
            return kind
    return "other"


def scan_fields(p):
    """扫描段落里的 Word 域，返回 [{kind, instr, result}]。

    复杂域在 XML 里是三段结构：
        fldChar(begin) → instrText(指令) → fldChar(separate) → 域结果 run → fldChar(end)

    这条管道此前对域完全隐形：`para_text` 只收 `w:t`，`instrText` 天然读不到，
    于是解析层看到的是「域结果」（长得和普通文字一样），IR 里只有纯文本，
    重建时域结构凭空蒸发——没有任何一处记录这件事。后果是 EndNote 插入的
    文献表被拆成死文本：作者返修想补一篇文献时，根本更新不了。

    简化域 w:fldSimple 没有 fldChar 三段，指令在属性上，直接取。

    这里只做识别与计数，不改动任何字节。文本框子树与 para_text 同口径跳过。
    """
    out = []
    stack = []  # 每层：{"instr": [...], "result": [...], "phase": "instr"|"result"}

    def walk(node, depth=1):
        if depth > XML_MAX_DEPTH:
            # DEPTHCAP：与 para_text 同口径（F2-2），入口闸之外的兜底
            raise ValueError(f"段落 XML 嵌套超过 {XML_MAX_DEPTH} 层，疑似异常文档，已拒收")
        for child in node:
            tag = child.tag
            if tag == W + "txbxContent":
                continue
            if tag == W + "fldSimple":
                instr = child.get(W + "instr") or ""
                result = "".join(t.text or "" for t in child.iter(W + "t"))
                out.append(_field_rec(instr, result))
            elif tag == W + "fldChar":
                kind = child.get(W + "fldCharType")
                if kind == "begin":
                    stack.append({"instr": [], "result": [], "phase": "instr"})
                elif kind == "separate" and stack:
                    stack[-1]["phase"] = "result"
                elif kind == "end" and stack:
                    f = stack.pop()
                    out.append(_field_rec("".join(f["instr"]), "".join(f["result"])))
            elif tag == W + "instrText" and stack:
                stack[-1]["instr"].append(child.text or "")
            elif tag == W + "t" and stack and stack[-1]["phase"] == "result":
                stack[-1]["result"].append(child.text or "")
            walk(child, depth + 1)

    walk(p)
    return out


def _field_rec(instr, result):
    instr = (instr or "").strip()
    return {"kind": field_kind(instr),
            "instr": instr[:FIELD_INSTR_MAX],
            "result": result[:FIELD_RESULT_MAX]}


def field_stats(blocks):
    """全文域的汇总：总数、按类型分布、含域的块数。

    报告要用它说一句「本文档含 N 个域」——尤其是 citation 类，
    那是「文字不归本工具改」的区域，必须让用户看见边界在哪。
    """
    by_kind = Counter()
    blocks_with = 0
    for b in blocks:
        fs = b.get("fields") or []
        if not fs:
            continue
        blocks_with += 1
        for f in fs:
            by_kind[f["kind"]] += 1
    return {"total": sum(by_kind.values()), "by_kind": dict(by_kind),
            "blocks_with_fields": blocks_with}


def digit_ratio(s):
    if not s:
        return 0
    return sum(c.isdigit() for c in s) / len(s)


def ascii_ratio(s):
    if not s:
        return 0
    return sum(ord(c) < 128 for c in s) / len(s)


STYLE_LEVEL = {
    "title": 1, "标题": 1,
    "heading1": 1, "heading2": 2, "heading3": 3, "heading4": 4,
    "标题1": 1, "标题2": 2, "标题3": 3, "标题4": 4,
}


def style_level(p):
    """读段落的 Word 样式级别（Heading 1 / 标题 2…），没套样式返回 None。

    真实稿件大多靠直接格式，但**用了**样式时这是最可靠的信号，比任何文本
    正则都强。此前 HEADING_STYLES 定义了却没人调用，等于把这条线索丢掉了。
    """
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    st = ppr.find(W + "pStyle")
    if st is None:
        return None
    val = (st.get(W + "val") or "").strip().lower().replace(" ", "")
    return STYLE_LEVEL.get(val)


def style_offset(body):
    """标题样式的级别偏移。

    不少作者把「文档标题」直接套成 Heading 1，正文一级标题顺延成 Heading 2。
    照朴素映射会整体高一级：文档标题被当成「一级标题」，真正的节标题被当成
    二级。判据是——全文第一个非空段落用了 Heading 1，且之后再没有 Heading 1，
    说明这个 1 被文档标题占用了。
    """
    seq = []
    for child in body:
        if child.tag != W + "p":
            continue
        txt = para_text(child)
        if not txt.strip():
            continue
        seq.append(style_level(child))
    if not seq or seq[0] != 1:
        return 0
    return 1 if 1 not in seq[1:] else 0


def classify_heading(text):
    """返回 (level, 置信度)。长度与数字占比是关键的排除条件。"""
    if not text or len(text) > MAX_HEADING_LEN:
        return None
    if text.endswith(("。", "；", "，", "：")) and not PAT["table_caption"].match(text):
        return None

    # 数字编号标题要排在数字占比检查之前：短标题里数字的占比天然偏高
    # （「1 引言」1/4，「2 节名」更高），用正文的 0.3 阈值判会把它们误杀。
    # 它们靠「短 + 不带句读标点」区分于正文句子，不靠数字占比。
    if PAT["h3_num"].match(text):
        return (3, "high")
    if PAT["h2_num"].match(text):
        return (2, "high")
    if PAT["h1_num"].match(text) and len(text) <= 30 and not HEADING_PUNCT.search(text):
        return (1, "high")

    # 中文序号前缀与阿拉伯序号同属强信号，同样要排在数字占比之前。
    # 「十、第10章」的占比是 2/6=0.33，按正文的 0.3 阈值会被判成正文——
    # 这一条掉进正文，后面所有章节号就整体错位一格。
    if PAT["h1"].match(text):
        return (1, "high")
    if PAT["h1_ws"].match(text) and len(text) <= 30 and not HEADING_PUNCT.search(text):
        return (1, "high")
    if PAT["h2"].match(text):
        return (2, "high")

    # 无编号的一级标题（「引  言」「结  论」）：本工具自己上一轮就是这么写的，
    # 认不出来就意味着重跑不幂等。置信度给 medium——它没有编号这个强信号，
    # 但有「整段只有节名」这个结构信号，且不会与正文句子混淆。
    if UNNUMBERED_H1.match(text):
        return (1, "medium")

    if digit_ratio(text) > 0.3:
        return None
    if PAT["h4"].match(text):
        return (4, "high")
    if PAT["h3"].match(text):
        # 「1.」是二级还是三级取决于本文档的编号体系，见 AMBIGUOUS_H3
        return (AMBIGUOUS_H3, "medium")
    return None


def detect_arabic_level(children):
    """判定「1.」「（1）」这类歧义编号的层级。

    返回 (arabic_level, arabic_sure, paren_level)：
      arabic_level —— 「1.」的级别；arabic_sure —— 是否有明确依据；
      paren_level  —— 「（1）」的级别（缺省 4）。

    单看「1.标题」本身定不了级：`（一）` 体系里它是三级，《经济管理》体例
    （一、／1.／（1））里它是二级。两种都在中文经管期刊里常见，猜错不是
    「排得不好看」，是整篇该级标题被套成另一级的样式（字号、对齐、缩进全错）。

    判据只有一条，但必须**全文预扫**而不能边扫边判：全文有没有中文括号二级。
    有「（一）」说明二级另有写法，「1.」只能是三级；没有，它就是二级。
    边扫边判会在「（一）」出现在第一个「1.」之后时判反——引言章直接进正文
    小节、二级标题晚于三级出现，在真实稿件里并不罕见。

    「经济管理」式推断（「1.」是二级）的前提是全文存在中文一级标题「一、」：
    那个体例是「一、／1.／（1）」。全文**一个中文编号都没有**的稿子（阿拉伯
    层级体例 1／1.1，英文刊常见，也是本工具给 numeric 刊重编号后的自身产物，
    IDMPOT），「1.」只能是顶级编号，判二级反而把整篇一级标题降级重排。

    「（1）」同理（IDMPOT）：它挂在哪个层级下面由「1.」的层级决定——
    一、／1.／（1）体系（有「一、」无「（一）」）里它是三级；有「（一）」的
    （一）／1.／（1.1）体系里它才是四级。渲染器给前一类刊写出的三级标题就是
    「（1）标题」，二次解析若仍判四级，渲染端没有四级分支、整段退回正文套版。

    第二个返回值是「有没有依据」。有「一、」但既没有「（一）」也没有「（1）」
    时，判二级是**推断**而非依据；调用方据此把置信度标成 ambiguous，别让
    推断混进「期刊要求」里。
    """
    has_cn1 = has_cn2 = has_arabic3 = has_h2num = False
    for _slot, child, _pos in children:
        if child.tag != W + "p":
            continue
        t = para_text(child)
        if not t:
            continue
        if PAT["h2"].match(t):
            has_cn2 = True
            break                 # 有中文括号二级就够了，无须扫完
        if PAT["h1"].match(t):
            has_cn1 = True
        if PAT["h4"].match(t):
            has_arabic3 = True
        if PAT["h2_num"].match(t):
            has_h2num = True
    if has_cn2:
        return 3, True, 4
    if not has_cn1:
        # 全文没有「一、」：「1.」是一级（阿拉伯层级体例的还原）。有「1.1」式
        # 二级（has_h2num）是直接依据；连它也没有（只有「N.」标题）时，一级
        # 仍是唯一自洽的层级——「一、／1.」体系必须有「一、」才成立。
        # 「（1）」不能推翻这一判级：上面的预扫用 PAT["h4"] 直接匹配原文，
        # 不带标题护栏，正文枚举句「（1）样本选择说明。」同样命中，分不出
        # 标题与正文句——只有「一、」在场（一、／1.／（1）体系铁证）时
        # 「（1）」才是层级证据。这里「（1）」保持缺省四级，不随「1.」下探
        return 1, True, 4
    # 有「一、」无「（一）」：一、／1.／（1）体系。「1.」是二级——有「（1）」
    # 在场是铁证（sure），否则是推断；「（1）」顺延为三级。
    return 2, has_arabic3, 3


def table_borders(tbl):
    """表格边框可能只在 tcBorders 里，必须两处都查。"""
    found = {}
    tp = tbl.find(W + "tblPr")
    if tp is not None:
        tb = tp.find(W + "tblBorders")
        if tb is not None:
            for e in tb:
                v = e.get(W + "val")
                if v and v != "none":
                    found[e.tag.replace(W, "")] = (v, e.get(W + "sz"))
    for tc in tbl.iter(W + "tc"):
        cb = tc.find(W + "tcPr/" + W + "tcBorders")
        if cb is not None:
            for e in cb:
                v = e.get(W + "val")
                if v and v != "nil" and v != "none":
                    key = "cell_" + e.tag.replace(W, "")
                    found.setdefault(key, (v, e.get(W + "sz")))
    return found


def has_drawing(p):
    for tag in ("drawing", "pict", "object"):
        if next(p.iter(W + tag), None) is not None:
            return True
    return False


# sdt 嵌套下钻的深度上限（SDTCAP）：真实文档的内容控件嵌到不了这个层数；
# 恶意/异常 docx 的超长嵌套链此前会让递归下钻打穿（RecursionError，批六
# F24 验证 depth=60 链可复现——解析器整篇崩掉）。到顶后不再下钻，整个深层
# 子树按「一个未识别结构」计数进 unsupported_content_blocks，对象原样保留，
# 解析汇总必须能看见这笔账。
SDT_MAX_DEPTH = 8


# XML 嵌套深度上限（DEPTHCAP）：para_text / scan_fields 的子树递归没有
# SDTCAP 那样的下钻限制（F2-2），普通段落里 2000 层嵌套的 w:r 链会把
# Python 递归打穿 RecursionError——解析器整篇崩掉。真实文档的段落/单元格
# 子树（嵌套表格、文本框、多层 OMML 公式）远到不了这个数：两个 walk 各自
# 带同口径深度计数，超限拒收。整树级不限——sdt 链的深度由 SDTCAP 单独
# 管（超深内容控件计数进 unsupported，对象原样保留），文档其余部分照常解析。
XML_MAX_DEPTH = 100


def _sdt_blocks(content, depth=1):
    """sdtContent 递归下钻：返回（可处理块元素列表, 未识别叶节点数）。

    depth 是当前内容控件的嵌套层数（顶层 sdt 为 1，SDTCAP）：超过
    SDT_MAX_DEPTH 的更深 w:sdt 不再下钻，整体计 1 个未识别结构。

    w:p / w:tbl 直接收进块列表；内嵌 w:sdt 继续下钻，其块按文档序并入同一
    列表（SDTNEST：批五只下钻一层，外层有直接段落时内嵌 sdt 的段落既不进
    IR 也不计数，深层内容控件里的正文被无声丢掉）；其余子节点是不认识的
    叶节点，逐个计数。一个块都没取到、也没有未识别节点（空内容控件）时
    计 1——整块没有可处理段落/表格的对象要报出来（F24 口径，原样保留）。

    渲染端 render_from_ir._sdt_blocks 必须与这里保持同一展开顺序与同一深度
    上限：IR 的 sdt_pos 是槽位内序，两边顺序不一致渲染就套错段。
    """
    blocks, unknown = [], 0
    for e in (list(content) if content is not None else []):
        if e.tag in (W + "p", W + "tbl"):
            blocks.append(e)
        elif e.tag == W + "sdt":
            if depth >= SDT_MAX_DEPTH:
                unknown += 1
                continue
            sub_blocks, sub_unknown = _sdt_blocks(e.find(W + "sdtContent"),
                                                  depth + 1)
            blocks.extend(sub_blocks)
            unknown += sub_unknown
        else:
            unknown += 1
    if not blocks and not unknown:
        unknown = 1
    return blocks, unknown


def bind_captions(blocks):
    """题注必须靠邻接关系判定，不能只看文本开头。

    反例（真实存在）：「表1报告了主要变量的描述性统计结果。……」是正文引用句，
    以「表1」开头却位于正文中间，不是表题。
    """
    is_cap_table = re.compile(r"^表\s*\d+")
    is_cap_figure = re.compile(r"^图\s*\d+")

    def short_cap(block, pat):
        return (block["type"] == "body" and block["text"]
                and pat.match(block["text"]) and len(block["text"]) <= MAX_HEADING_LEN
                and not block["text"].endswith(("。", "；")))

    for i, b in enumerate(blocks):
        if b["type"] == "table":
            for j in range(i - 1, max(i - 4, -1), -1):
                if short_cap(blocks[j], is_cap_table):
                    blocks[j]["type"] = "table_caption"
                    b["caption"] = blocks[j]["text"]
                    break
                if blocks[j]["type"] in ("table", "body") and blocks[j]["text"] and \
                        len(blocks[j]["text"]) > MAX_HEADING_LEN:
                    break
        elif b["type"] == "figure":
            for j in range(i + 1, min(i + 3, len(blocks))):
                if short_cap(blocks[j], is_cap_figure):
                    blocks[j]["type"] = "figure_caption"
                    b["caption"] = blocks[j]["text"]
                    break


def _normalize_ai_headings(blocks):
    """AI 标出的层级不能跳级：没有二级就不该有三级。

    只处理 AI 改过的块，规则自己判出来的标题原样保留——否则等于
    顺手改了既有行为，样本回归会立刻炸。
    """
    seen = set()
    for b in blocks:
        t = b["type"]
        if not (t.startswith("heading") and t[-1].isdigit()):
            continue
        lv = int(t[-1])
        if lv == 1 or lv - 1 in seen:
            seen.add(lv)
            continue
        if b.get("ai_from"):
            b["type"] = "body"
            b["ai_demoted"] = True
            b.pop("level", None)


def apply_ai(blocks, cfg, log=None):
    """让 AI 补结构标签，返回 trace。

    三条不对称规则，是这套接入的安全边界：
      1. 规则标 high confidence 的块，AI 只许在它落到 body/front_matter 时才动
      2. AI 只能改 type / level，输入输出都不含文本字段
      3. 动手前后做指纹比对，不一致则恢复现场并判为失败
    """
    log = log or (lambda *_: None)
    before = ai_assist.fingerprint(blocks)
    # 快照必须含 text：回滚只还原标签、不还原文本，就等于没守住这条线。
    # 表格块本来没有 text 键，用 MISSING 区分"原本没有"与"原本是空串"。
    MISSING = object()
    snapshot = [(b["idx"], b["type"], b.get("level"), b.get("confidence"),
                 b["text"] if "text" in b else MISSING) for b in blocks]

    cands = [b for b in blocks
             if b["type"] in AI_UPGRADE_FROM
             or b.get("confidence") in ("low", "medium", "ambiguous")]
    if not cands:
        return {"used": False, "reason": "规则已给出确定判断，无需补漏"}

    # trace（F50）：截断情况随返回值落进 ir["ai"]，审计 JSON 有留痕，
    # 不能只留在 stdout 日志行里
    ai_trace = {}
    prop = ai_assist.propose_types(blocks, cands, cfg, log=log, trace=ai_trace)
    if not prop:
        if ai_trace:
            return {"used": False, "reason": "无有效提案（未配置、调用失败或返回不合规）",
                    "truncation": ai_trace}
        return {"used": False, "reason": "无有效提案（未配置、调用失败或返回不合规）"}

    by_idx = {b["idx"]: b for b in blocks}
    items = []
    for idx, p in sorted(prop.items()):
        b = by_idx.get(idx)
        if b is None or p["type"] == b["type"]:
            continue
        if b["type"] not in AI_UPGRADE_FROM and b.get("confidence") == "high":
            items.append({"idx": idx, "from": b["type"], "to": p["type"],
                          "status": "rejected", "why": "规则高置信，AI 不得推翻"})
            continue
        old = b["type"]
        b["type"] = p["type"]
        b["ai_from"] = old
        b["ai_why"] = p["why"]
        if p["type"].startswith("heading"):
            b["level"] = int(p["type"][-1])
            b["confidence"] = "ai-" + p["conf"]
        items.append({"idx": idx, "from": old, "to": p["type"],
                      "status": "applied", "conf": p["conf"], "why": p["why"]})

    _normalize_ai_headings(blocks)

    if not ai_assist.text_unchanged(before, ai_assist.fingerprint(blocks)):
        for (idx, t, lv, cf, tx) in snapshot:
            b = by_idx.get(idx)
            if b is None:
                continue
            if tx is MISSING:
                b.pop("text", None)
            else:
                b["text"] = tx
            b["type"] = t
            b.pop("ai_from", None)
            b.pop("ai_why", None)
            b.pop("ai_demoted", None)
            if lv is None:
                b.pop("level", None)
            else:
                b["level"] = lv
            if cf is None:
                b.pop("confidence", None)
            else:
                b["confidence"] = cf
        if ai_trace:
            return {"used": False, "reason": "指纹校验失败：AI 路径动到了文本，整批已回滚",
                    "truncation": ai_trace}
        return {"used": False, "reason": "指纹校验失败：AI 路径动到了文本，整批已回滚"}

    applied = [i for i in items if i["status"] == "applied"]
    log(f"AI 补漏：候选 {len(cands)} 块，提案 {len(prop)} 条，采纳 {len(applied)} 条")
    if ai_trace:
        return {"used": True, "model": cfg.get("model"),
                "candidates": len(cands), "proposed": len(prop),
                "applied": len(applied), "items": items,
                "truncation": ai_trace}
    return {"used": True, "model": cfg.get("model"),
            "candidates": len(cands), "proposed": len(prop),
            "applied": len(applied), "items": items}


def parse(path, ai_cfg=None):
    z = zipfile.ZipFile(path)
    doc = xmlsafe.fromstring(xmlsafe.read_part(z, "word/document.xml"))
    body = doc.find(W + "body")

    blocks = []
    seen_level = set()
    in_references = False
    unsupported_content = 0

    def emit(block, el=None):
        if sdt_pos is not None:
            block["sdt_pos"] = sdt_pos    # sdt 内容块：同槽内的序号（F24）
        # 域随块一起进 IR：重建端要靠它判断「这块的文字归 EndNote 所有，
        # 本工具不动」，报告也要靠它把「本文档含 N 个域」说清楚
        if el is not None:
            fs = scan_fields(el)
            if fs:
                block["fields"] = fs
        blocks.append(block)
        return block

    # idx 对 body 全部子元素统一编号（dense）：bookmarkStart、sdt 等异类
    # 子元素也占一个号。消费端（render_from_ir / anonymize / apply_details）
    # 一律按 body 子元素位置定位，生产端只数 w:p/w:tbl 的话，body 里一旦
    # 出现位于块元素之前的书签等元素，双方就整体错位——套版错段、改错正文、
    # 匿名路径漏删（F08）。sdt 内容控件也占自己的槽号；其内的 w:p/w:tbl
    # 递归下钻（内嵌 sdt 继续展开直到取尽段落/表格，SDTNEST）为块、共用该
    # 槽号并带 sdt_pos 序号（F24）——不重新编号，按 body 子元素位置定位的
    # 契约不破。取尽后仍不认识的叶节点与取空的内容控件记数进 IR，不让
    # 「这里没处理」变成静默。段落/表格内部的行内 sdt 文字已随段落文本
    # 处理，不算未处理块。
    # 标题样式级别整体偏移（文档标题占用了 Heading 1 时）；规则识别的兜底仍走文本
    offset = style_offset(body)
    children = []
    for slot, child in enumerate(body):
        if child.tag == W + "sdt":
            # 变量名避开 parse() 本地的 blocks（IR 块列表）——下钻产物是
            # 原始 XML 元素，混进来就是类型灾难
            inner_els, unknown = _sdt_blocks(child.find(W + "sdtContent"))
            for pos, e in enumerate(inner_els):
                children.append((slot, e, pos))
            unsupported_content += unknown
            continue
        if child.tag in (W + "tbl", W + "p"):
            children.append((slot, child, None))

    # 「1.」「（1）」的级别全文预扫一次（见 detect_arabic_level）：必须在进
    # 主循环之前拿到结论，否则「（一）」晚于第一个「1.」出现时会判反
    arabic_level, arabic_sure, paren_level = detect_arabic_level(children)

    for idx, child, sdt_pos in children:
        if child.tag == W + "tbl":
            text = "".join(t.text or "" for t in child.iter(W + "t")).strip()
            rows = child.findall(W + "tr")
            cols = len(rows[0].findall(W + "tc")) if rows else 0
            # 伪表格：1 行且列数很少、文字极短，实际是公式加编号
            if len(rows) == 1 and cols <= 3 and len(text) <= 60:
                emit({"type": "equation_pseudo_table", "text": text, "idx": idx,
                      "n_rows": len(rows), "n_cols": cols}, child)
                continue
            emit({
                "type": "table", "idx": idx, "n_rows": len(rows), "n_cols": cols,
                "borders": table_borders(child),
                "header": [para_text(tc) for tc in rows[0].findall(W + "tc")] if rows else [],
                # 表内文字按行平铺进 cells（表格块本身没有 text 键）：声明
                # 「全文（含图表）」口径的刊统计时要数表格内文字，缺了它
                # 标注与计算就自相矛盾。图内文字在图片对象里，文本层没有，天然不计
                "cells": [para_text(tc) for tr in rows for tc in tr.findall(W + "tc")],
            }, child)
            continue

        text = para_text(child)
        has_img = has_drawing(child)
        if not text and has_img:
            emit({"type": "figure", "text": "", "idx": idx}, child)
            continue
        if not text:
            # 纯公式段落没有文字，但它是内容，不能像空行一样丢掉——
            # 丢了渲染器就不知道这里有个公式，也就不会给它上行距保护。
            if has_omml(child):
                emit({"type": "equation", "text": "", "idx": idx, "no_text": True}, child)
            continue

        level = conf = None
        if PAT["abstract"].match(text):
            btype = "abstract"
        elif PAT["keywords"].match(text):
            btype = "keywords"
        elif PAT["classification"].match(text):
            btype = "classification"
        elif PAT["ref_head"].match(text):
            btype = "references_heading"
            in_references = True
        elif PAT["note"].match(text):
            btype = "table_note"
        elif PAT["appendix"].match(text):
            btype = "appendix_heading"
            in_references = False
        elif PAT["en_abstract"].match(text):
            btype = "english_abstract"
        elif PAT["en_keywords"].match(text):
            btype = "english_keywords"
        elif in_references and looks_like_ref(text):
            btype = "reference_item"
        else:
            if in_references:
                in_references = False
            # 作者套了标题样式就以样式为准（含偏移校正），这是作者自己声明的
            # 层级；没套样式才退回文本模式识别
            h = None
            slv = style_level(child)
            if slv is not None and slv - offset >= 1:
                h = (slv - offset, "style")
            if h is None:
                h = classify_heading(text)
            if h:
                cand, conf = h
                if cand == AMBIGUOUS_H3:
                    # 级别由全文预扫给出（detect_arabic_level）：同一份稿子里
                    # 「1.」不会一会儿二级一会儿三级。全文没有任何二级编号
                    # 痕迹时，这个级别是推出来的，置信度标 ambiguous——
                    # 报告里要能看出「这一级是猜的」，不能混进「期刊要求」。
                    level = arabic_level
                    if not arabic_sure:
                        conf = "ambiguous"
                else:
                    level = cand
                    if level == 4 and paren_level != 4 and h[1] != "style":
                        # 「（1）」的层级随编号体系顺延（detect_arabic_level）：
                        # 一、／1.／（1）体系里它是三级，判成四级会让这一级标题
                        # 按 level4 重排；全文没有 （一） 的依据，重跑即漂移（IDMPOT）。
                        # 但作者显式套了 Heading 4 样式时以样式为准——那是作者
                        # 自己声明的层级，不能按文本编号再降一级。
                        level = paren_level
                if level is None:
                    btype, conf = "body", "low"
                elif level == 3 and 2 not in seen_level:
                    btype, conf = "body", "low"
                else:
                    btype = f"heading{level}"
                    seen_level.add(level)
            elif has_omml(child) and len(text) < 120:
                # 公式检测必须排在这里：标题、摘要这些结构一旦被认成 equation，
                # 后面整篇的标题编号与正文套版都会失准——带公式的标题很常见
                # （「二、模型：y=α+βx」），它不是公式段落。
                btype = "equation"
            elif not seen_level and len(text) < 250 and ascii_ratio(text) > 0.7:
                btype = "english_title"
            elif len(blocks) <= 3 and len(text) < 80 and sdt_pos is None:
                # 前置区是按「前三段」的位置猜的：藏在内容控件里的段落不参与
                # 这个推断——补录占位符路径会把非段落元素当段落包一层，直接崩
                btype = "front_matter"
            else:
                btype = "body"

        block = {"type": btype, "text": text, "idx": idx}
        if has_img:
            block["has_image"] = True
        if level:
            block["level"] = level
            block["confidence"] = conf
        emit(block, child)

    # 顺序不可换：规则先给基线 → AI 补漏（只改标签）→ 题注绑定。
    # 题注绑定放最后，因为 AI 可能刚把某个 body 认出是表题。
    trace = {"used": False, "reason": "未启用"}
    if ai_cfg is None:
        ai_cfg = ai_assist.load_config()
    if ai_assist.available(ai_cfg):
        # log 进解析 stdout（WEBLOG2）：截断、调用失败、采纳数不再静默；
        # 截断数值另由 apply_ai 内部 trace 带回，落进 ir["ai"]（F50）。
        trace = apply_ai(blocks, ai_cfg, log=print)
    bind_captions(blocks)

    stats = Counter(b["type"] for b in blocks)
    ir = {
        "source": str(Path(path).name),
        "stats": dict(stats),
        "block_total": len(blocks),
        "blocks": blocks,
        "ai": trace,
    }
    if unsupported_content:
        # 取尽段落/表格后仍不认识的内容控件结构（未识别的块级节点、取空的
        # 内容控件）记数进 IR（F24/SDTNEST），渲染汇总据此提示，
        # 不让「这里没处理」静默
        ir["unsupported_content_blocks"] = unsupported_content
    fstats = field_stats(blocks)
    if fstats["total"]:
        # 域统计：重建端与报告都要用。不记录的话，「文献表被拆成死文本」
        # 这件事在全流程里没有一处留下痕迹
        ir["fields"] = fstats
    return ir


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--ai", dest="ai", action="store_true", default=None,
                    help="强制启用 AI 补漏（需 config/ai.json 已配好）")
    ap.add_argument("--no-ai", dest="ai", action="store_false",
                    help="强制关闭，纯规则路径")
    args = ap.parse_args()

    cfg = ai_assist.load_config()
    if args.ai is not None:
        cfg["enabled"] = args.ai
    ir = parse(args.docx, ai_cfg=cfg)
    out = Path(args.out) if args.out else Path(args.docx).with_suffix(".ir.json")
    out.write_text(json.dumps(ir, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"=== {ir['source']} ===")
    print(f"识别 {ir['block_total']} 个块\n")
    for k, v in sorted(ir["stats"].items(), key=lambda x: -x[1]):
        print(f"  {v:5d}  {k}")

    if ir.get("unsupported_content_blocks"):
        print(f"\n  未处理内容块: {ir['unsupported_content_blocks']} 个"
              f"（内容控件里未识别的块级节点、取空或嵌套超过 "
              f"{SDT_MAX_DEPTH} 层的内容控件，对象原样保留）")

    fs = ir.get("fields") or {}
    if fs.get("total"):
        kinds = "、".join(f"{k} {v}" for k, v in sorted(fs["by_kind"].items(), key=lambda x: -x[1]))
        print(f"\n  检测到 {fs['total']} 个 Word 域（{kinds}）"
              f"，分布在 {fs['blocks_with_fields']} 个块")
        if fs["by_kind"].get("citation"):
            print("  ⚠ 含文献管理域：这些位置的文字由 EndNote/Zotero 掌管，"
                  "本工具只改版式、不改文字")

    heads = [b for b in ir["blocks"] if b["type"].startswith("heading")]
    print(f"\n-- 识别出的标题（{len(heads)} 个） --")
    for b in heads:
        flag = "" if b.get("confidence") == "high" else f"  ← {b.get('confidence')}"
        print(f"  H{b['level']}  {b['text'][:40]}{flag}")

    low = [b for b in ir["blocks"] if b.get("confidence") == "low"]
    if low:
        print(f"\n-- 存疑（可能漏判的标题） --")
        for b in low[:10]:
            print(f"  {b['text'][:50]}")

    ai = ir.get("ai") or {}
    if ai.get("used"):
        print(f"\n-- AI 补漏（{ai.get('model')}） --")
        for it in ai.get("items", []):
            if it["status"] == "applied":
                print(f"  #{it['idx']}  {it['from']} → {it['to']}  ({it['why']})")
        print(f"  采纳 {ai.get('applied', 0)} 条")
    else:
        print(f"\n-- AI 补漏：未参与（{ai.get('reason', '未启用')}） --")

    print(f"\n已写入 {out}")


if __name__ == "__main__":
    main()
