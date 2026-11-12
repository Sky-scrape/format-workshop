"""文字白名单守卫：把「正文文字零改动」落成运行期可执行的判定。

这份实现是从 tests/conftest.py 搬来的（F31 的「同一不变式只有一份实现」）。
搬出来不是为了整理代码：判定要能在 pytest 之外复用（产物导出前自检、桌面端
复核），留在夹具里就只有测试能调用，运行期反而没有守卫。

口径（DESIGN §4.1，2026-11-13）：渲染器真正落地的字符级归一为段首缩进
删除（N-01 段首＋N-08，仅 body/abstract/keywords）、段尾空白删除（N-01
段尾）、全角字母数字转半角（N-05，仅 body/abstract/keywords）、标题编号
前缀重写、题注编号间隔归一、文献条目编号前缀（N-07）。N-02～N-04、N-06
（正文内部空格、全半角标点、数字与中文间空格）未实现，守卫逐字比对——
「无条件删全部空白」的宽松口径不许回来。
"""

import re
from collections import Counter

from docx import Document
from docx.oxml.ns import qn

# 段首缩进只对这三类块允许删（N-01 段首＋N-08 的唯一落地处）。
INDENT_LEADING_TYPES = {"body", "abstract", "keywords"}
# N-05：全角字母数字转半角。只认这三段连续区，与渲染器 FULLWIDTH_ALNUM 同口径。
FULLWIDTH_ALNUM = re.compile(r"[Ａ-Ｚａ-ｚ０-９]")
HEADING_TYPES = {"heading1", "heading2", "heading3", "heading4"}
CAPTION_TYPES = {"table_caption", "figure_caption"}

# 编号前缀的合法形态（23 本规则库实有的体例，F43 全库矩阵补齐）：中文数字＋
# 顿号（一、）、中文数字加括号（（一））、阿拉伯数字加括号（（1）——ciejournal/
# jjgl/nkglpl 的三级体例）、中文数字＋空白分隔（「一　」——sjjj 的一级体例）、
# 阿拉伯数字加点（1.／1.2.）。空白分隔形态要求分隔符处真有空白，否则「一带一路
# 研究」这类以数字开头的标题会被剥掉第一个字；括号与点号两侧的空白可容忍
# （剥前缀在塌缩留白之前进行，见 normalize）。
HEADING_NUMBER_PREFIXES = (re.compile(r"^[一二三四五六七八九十]{1,3}\s*[、]"),
                           re.compile(r"^[（(]\s*[一二三四五六七八九十]{1,3}\s*[）)]"),
                           re.compile(r"^[（(]\s*\d{1,2}\s*[）)]"),
                           re.compile(r"^[一二三四五六七八九十]{1,3}\s+"),
                           # 阿拉伯数字前缀必须带分隔（点号或空白）才算编号：
                           # 旧写法允许「1」后面直接跟字母，把「5G技术研究」
                           # 剥成「G技术研究」。多级形式的最终分隔符后若还是
                           # 数字（1.2.3），说明是「1.2.」被截了，不能算匹配。
                           re.compile(r"^\d{1,2}(?:\s*[.．]\s*\d{1,2})+[.．](?=[^\dA-Za-z]|$)"),
                           re.compile(r"^\d{1,2}[.．](?=[^\dA-Za-z]|$)"),
                           re.compile(r"^\d{1,2}(?:\s*[.．]\s*\d{1,2})+(?=[^\dA-Za-z.．]|$)"),
                           re.compile(r"^\d{1,2}(?=[^\dA-Za-z.．]|$)"))
CAPTION_PREFIX = re.compile(r"^(表|图)\s*\d+")
REF_NUMBER_PREFIX = re.compile(
    r"^\s*[\[［〔【(（]?\s*\d{1,3}\s*[\]］〕】)）]?\s*[.．、]?\s*")


def normalize(text, btype=""):
    """把该块类型下白名单允许的字符级归一抹平，剩下的差异才算真改动。

    口径按块类型收紧（F31）：段首缩进空白只允许 body/abstract/keywords 删；
    标题允许抹编号前缀与留白（未编号节的标题文字整段来自规则集
    intro_title_text）；题注允许抹「表1/图2」前缀与两端空白（题注归一只改
    间隔）；文献条目允许抹编号前缀（N-07）。其余——包括全角↔半角、中文与
    拉丁之间的空格——一律不在白名单：normalize('数字金融 A 股') 不再等于
    normalize('数字金融A股')。

    2026-11-13 起 body/abstract/keywords 另吃两项白名单（N-01 段尾、N-05）：
    段尾空白删除与全角字母数字（Ａ１→A1）转半角。仍只删空白/换编码，不动
    任何文字、数字或标点；其余块类型不适用。
    """
    s = text
    if btype in INDENT_LEADING_TYPES:
        s = s.lstrip("\t \u3000").rstrip("\t \u3000")
        if FULLWIDTH_ALNUM.search(s):
            s = "".join(chr(ord(c) - 0xFEE0) if FULLWIDTH_ALNUM.match(c) else c
                        for c in s)
    if btype in HEADING_TYPES:
        # 先剥编号前缀、再塌留白（F43）：编号与标题之间的空白是分隔符
        # （「一　文献综述」），先塌缩会把编号粘进标题，前缀就再也剥不下来；
        # 剥完剩下的空白才是版式，按原口径塌平。
        s = s.strip()
        for pat in HEADING_NUMBER_PREFIXES:
            s = pat.sub("", s).strip()
        s = re.sub(r"\s+", "", s)
    elif btype in CAPTION_TYPES:
        s = CAPTION_PREFIX.sub("", s.strip()).strip()
    elif btype == "reference_item":
        s = REF_NUMBER_PREFIX.sub("", s, count=1)
    return s


def content_of(text, btype, label=None):
    """标签块只比对冒号之后的正文。

    「内容摘要：」→「摘  要：」这类属于标签归一（B 档，规则集明确定义），改的
    是标签本身，不是内容。比对时把标签剥掉，剩下的差异才算真改动。

    冒号要同时认半角——真实稿件里「内容摘要:」很常见，标签归一顺带把它改成
    全角，这是允许的，不该被判成改字。

    label 是目标刊规则集 blocks[type].zh.label 的值（F43：与渲染器读同一
    字段，守卫不硬编码标签样式——承诺④）。它整段替换源稿标签且可以不带冒号
    （「【摘要】」「内容提要　」「关  键  词　」），冒号启发式够不着，所以
    前缀对得上时直接剥；对不上（源稿是投稿人自己的标签）时退回冒号口径。

    文献表标题是纯标签（「参考文献」本身没有正文），比对指纹一律记为空——
    否则指纹差集会把「参考文献 → 参考文献：」误报成丢字/生字。但改字不能
    无条件放行：whitelist_violations 对这类块单独要求同 idx 的
    `文献表标题归一` 审计条目（见 references_heading_blocks）。
    """
    if btype not in ("abstract", "keywords", "references_heading"):
        return text
    if btype == "references_heading":
        return ""
    if label and text.startswith(label):
        return text[len(label):]
    m = re.search(r"[:：]", text)
    # 标签很短（"内容摘要:" 6 个字），指望号出现在前 12 个字内才算标签分隔
    if m and m.start() <= 12:
        return text[m.end():]
    return text.rstrip("：:")


def keyword_terms(body, sep=None):
    """关键词串切词条。口径与渲染器 split_keywords 同一（tools/render_from_ir.py）。

    逗号、顿号、分号是明确的分隔符，总是切；空白分两等——目标分隔符本身是
    空白时全部当分隔符，否则只有「连续两个以上的空白」或全角空格才算
    （英文词组「digital finance」里的单词空格是词间空格）。sep 缺省按非空白
    分支（保守侧）：守卫欠切只是把词条连成大块来比，渲染器错把一个关键词
    拆成两个会被抓到；过切则会把这类错放行。
    """
    pattern = (r"[\s,，;；、]+" if sep is not None and not sep.strip()
               else r"(?:\s{2,}|\u3000|[,，;；、])+")
    return [t for t in re.split(pattern, body or "") if t.strip()]


def signature(text, btype, kw_sep=None, label=None):
    """块的可比对指纹。

    关键词单独处理：分隔符从全角空格换成「；」属于规则集定义的归一
    （`关键词分隔符归一`），所以按目标刊的分隔符口径拆成词条再比，词条
    本身逐字比（kw_sep 是目标刊 blocks:keywords/zh/separator 的值）。

    标签块按目标刊规则集的 label 剥标签（F43，见 content_of）。

    参考文献条目剥掉编号前缀（见 normalize）：`[1]`→`〔1〕` 这类改写法属
    N-07 白名单，但条目正文必须逐字不变，所以只比前缀之后的部分。
    """
    body = content_of(text, btype, label)
    if btype == "keywords":
        return tuple(normalize(t) for t in keyword_terms(body, kw_sep))
    return (normalize(body, btype),)


# 会真正改写文字的动作。只有这些条目的 before/after 是文字快照，
# 其余（页面设置、套版、悬挂缩进）存的是格式值。
# 会真正改写文字、且审计 before/after 存的是**整段快照**的动作。
TEXT_ACTIONS = {"标签归一", "关键词分隔符归一", "题注归一", "文献表标题归一",
                "标题重新编号", "文献条目编号归一"}

# 同样改文字、但审计只存改动附近片段的动作（N-01 段首/段尾、N-05）。
# 它们不进 TEXT_ACTIONS——那一组的口径是「before/after 与源稿成品逐字相
# 等」，片段快照不满足；槽位审计仍按「这里动过字」处理。
FRAGMENT_TEXT_ACTIONS = {"段首缩进空格删除", "段尾空白删除", "全角字母数字归一"}

# --citation 的豁免凭证（F39）：这两个动作会改写正文/条目文字，属产品唯一
# 越线开关。豁免以审计条目本身为准——条目存在且 rule 指对才豁免，改了字
# 没记录、或 rule 指错位置的，一律仍按白名单外改动处理。
CITATION_REWRITE_RULES = {
    "正文引用改作者-年份": "citation.in_text",
    "条目按目标刊体例重排": "citation.reference_list.templates",
}

# 槽位覆盖只问「这条审计有没有动文字」，口径比 TEXT_ACTIONS 宽一格：段首缩进
# 删除（N-01＋N-08）同样改字，只是记的是片段快照、没进全文快照登记；槽位审计
# 认不出它就会把白名单内的缩进归一误判成槽内改字。引用改写不在这里——它是
# 产品唯一越线开关，动作与 rule 对得上才认（见 _is_text_action）。
SLOT_TEXT_ACTIONS = TEXT_ACTIONS | FRAGMENT_TEXT_ACTIONS


class DocumentTexts(dict):
    """document_texts 的返回表，额外记下哪些下标是 sdt 槽位。

    键仍是 body 子元素下标（int）：下游按 int 下标取普通块文字（审计 before/
    after 对账、匿名路径）的老口径不能变。槽位身份单独放在 sdt_slots 属性里，
    而不是把键改成 "slot.5"——同一份文字表混两套键，任何 `texts[idx]` 的既有
    调用都要分情况，迟早漏掉一路。
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.sdt_slots = set()


def _wtext(el):
    """元素子树里的全部 w:t（不含 w:delText 等非正文文字节点）。"""
    return "".join(n.text or "" for n in el.iter(qn("w:t")))


def document_texts(path):
    """按 body 子元素下标取文字；下标与 IR 的 idx 同源。

    - w:p：取后代全部 w:t（段内的行内 sdt、超链接文字都随段落走）；
    - w:tbl：逐单元格取 w:t，单元格之间用 "\\x00" 分隔——直接拼接会让相邻
      单元格的文字粘成一段，表格里换一个字就可能落在边界上看不出来；
    - w:sdt：整槽取全部后代 w:t。槽内的段落/表格在 IR 里共用槽号（另带
      sdt_pos），守卫按 body 子元素对账，槽内改字只能整槽看；槽号登记进
      sdt_slots，whitelist_violations 靠它区分「槽位变化」与普通段落变化。

    和 tests/conftest.py 的 body_texts 相比，只有 sdt 这一支是新增的：
    旧实现整槽跳过，内容控件里的正文等于没人守。
    """
    doc = Document(str(path))
    out = DocumentTexts()
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            out[i] = _wtext(child)
        elif child.tag == qn("w:tbl"):
            cells = [_wtext(tc) for tc in child.iter(qn("w:tc"))]
            out[i] = "\x00".join(cells)
        elif child.tag == qn("w:sdt"):
            out[i] = _wtext(child)
            out.sdt_slots.add(i)
    return out


def _changed_indices(source_texts, output_texts):
    """成品里与源稿不同的块；源稿有、成品没有的块由 whitelist_violations 另报。"""
    return {i for i in output_texts if output_texts[i] != source_texts.get(i)}


def _block_type(idx, block_types):
    """块类型；sdt 槽位也可能被调用方用 "slot.5" 这种字符串下标传进来。"""
    if idx in block_types:
        return block_types.get(idx) or ""
    if isinstance(idx, str) and idx.startswith("slot."):
        num = idx[len("slot."):].split(".", 1)[0]
        if num.isdigit():
            return block_types.get(int(num)) or ""
    return ""


def _order_key(idx):
    """int 下标按数值排，字符串槽位键排在后面——混排时不能靠 str 比较。"""
    if isinstance(idx, int):
        return (0, idx, "")
    return (1, 0, str(idx))


def _sdt_slots(texts):
    return set(getattr(texts, "sdt_slots", None) or ())


def _is_text_action(item):
    action = (item.get("action") or "").strip()
    if action in SLOT_TEXT_ACTIONS:
        return True
    # 引用改写：动作与 rule 都对得上才算「这条审计动了文字」
    return CITATION_REWRITE_RULES.get(action) == item.get("rule")


def _slot_no(idx):
    """槽位下标 → 槽号；兼容 int、"slot.5" 与渲染器的 "5.0"（槽.序）三种写法。"""
    if isinstance(idx, int):
        return idx
    s = str(idx)
    if s.startswith("slot."):
        s = s[len("slot."):]
    head = s.split(".", 1)[0]
    return int(head) if head.isdigit() else None


def _slot_audited(items, slot):
    """该槽位有没有文字动作审计条目。

    槽位在审计里有两种 idx 写法：调用方/测试用 "slot.<槽号>"，渲染器实际写
    "<槽号>.<槽内序号>"（render_from_ir 的 sdt 审计键）。只认前者会把槽内的
    合法段首缩进删除判成白名单外改字。契约是整槽一条：该槽有任意一条文字动作
    条目即视为已进审计，但必须按槽判断——一个槽的条目不能给别的槽背书。
    """
    for it in items:
        if not _is_text_action(it):
            continue
        s = str(it.get("idx", ""))
        if slot is not None and (s == f"slot.{slot}" or s.startswith(f"slot.{slot}.")
                                 or s.startswith(f"{slot}.")):
            return True
    return False


def _declared_slots(items):
    """从文字动作审计条目反推槽号，供调用方传普通 dict 时识别槽位。"""
    out = set()
    for it in items:
        if not _is_text_action(it):
            continue
        s = str(it.get("idx", ""))
        if "." not in s:
            continue
        num = _slot_no(s)
        if num is not None:
            out.add(num)
    return out


def note_reorder_blocks(source_texts, output_texts, block_types):
    """只换了位置、没换文字的表格注。

    目标刊要求「注」与「资料来源」的先后不同时，渲染器会交换这两段的内容
    （`表注／图注层序调整`）。逐块比对必然误判成「改了字」，所以先把这类纯
    换位识别出来：这些块各自的文字都变了，但整组文字守恒。

    只认 table_note——题注（表1 xxx）不参与层序调整，它们的空格归一是另一回
    事，混进来会让整组文字不守恒，反而识别不出真正的换位。
    """
    changed = _changed_indices(source_texts, output_texts)
    notes = {i for i in changed if _block_type(i, block_types) == "table_note"}
    if len(notes) < 2:
        return set()

    def sig(idx, texts):
        return signature(texts.get(idx, ""), "table_note")

    changed_notes = {i for i in notes if sig(i, source_texts) != sig(i, output_texts)}
    if not changed_notes:
        return set()

    before = Counter(sig(i, source_texts) for i in notes)
    after = Counter(sig(i, output_texts) for i in notes)
    return changed_notes if before == after else set()


def citation_rewrite_blocks(items):
    """凭引用改写获得豁免的块：动作与 rule 都对得上的审计条目所在块。"""
    return {it["idx"] for it in items
            if CITATION_REWRITE_RULES.get(it.get("action")) == it.get("rule")}


def references_heading_blocks(items):
    """文献表标题归一获得豁免的块：必须有同 idx 的「文献表标题归一」条目。

    标题是纯标签、没有可逐字比的正文，但也不能无条件放行——否则「参考文献」
    被改成任意文字都算合规。渲染器每次改标题都会写这条审计，要求条目在场即可。
    """
    return {it["idx"] for it in items
            if it.get("action") == "文献表标题归一" and _is_text_action(it)}


def whitelist_violations(source_texts, output_texts, block_types, items,
                         labels=None, kw_sep=None):
    """超出白名单的文字改动 [(idx, 块类型, before, after)]。

    引用改写块豁免逐块比对，但豁免只在条目（action+rule）对得上时成立：
    改了字而没有引用改写条目的，照旧算违规——豁免不豁免审计。

    sdt 槽位是另一条口径：槽内多段在 IR 里共用槽号，逐块文字对不上号，只能
    按整槽看。槽位有变化而该槽没有文字动作审计条目时，整槽算违规；有则视为
    已覆盖（条目里的 before/after 快照供报告核对）。

    少块同样是违规：运行期闸只拿输出下标遍历，成品末尾段/表被删就会完全
    看不见。这里把源稿有、成品没有的下标一并报出来，宁可失败也不交付丢字稿。
    """
    labels = labels or {}
    exempt = citation_rewrite_blocks(items) | references_heading_blocks(items)
    reordered = note_reorder_blocks(source_texts, output_texts, block_types)
    slots = (_sdt_slots(source_texts) | _sdt_slots(output_texts)
             | _declared_slots(items))

    offenders = []
    for idx in sorted(set(source_texts) - set(output_texts), key=_order_key):
        offenders.append((idx, _block_type(idx, block_types),
                          source_texts.get(idx, "")[:40], ""))

    for idx in sorted(_changed_indices(source_texts, output_texts), key=_order_key):
        before = source_texts.get(idx, "")
        after = output_texts.get(idx, "")
        btype = _block_type(idx, block_types)
        is_slot = idx in slots or (isinstance(idx, str) and idx.startswith("slot."))
        if is_slot:
            # 槽位不套 normalize：槽里可能是正文、题注、表格混排，逐块白名单
            # 无从判定；这正需要一条整槽审计来交代。
            if not _slot_audited(items, _slot_no(idx)):
                offenders.append((idx, btype, before[:40], after[:40]))
            continue
        if idx in reordered:
            continue
        if btype == "references_heading":
            # 标题是纯标签，指纹恒为空；改字必须由同 idx 的审计条目背书。
            if before != after and idx not in exempt:
                offenders.append((idx, btype, before[:40], after[:40]))
            continue
        label = labels.get(btype)
        if signature(before, btype, kw_sep=kw_sep, label=label) == \
                signature(after, btype, kw_sep=kw_sep, label=label):
            continue
        if idx in exempt:
            continue
        offenders.append((idx, btype, before[:40], after[:40]))
    return offenders
