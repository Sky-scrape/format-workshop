"""字符级归一的纯文本实现（DESIGN §4.1 N-02/N-03/N-04/N-06）。

渲染器与交付前的文字闸（tools/text_guard.py）必须消费同一套规则，否则
「渲染器改了、守卫不认」会阻断交付，「守卫放了、渲染器没做」会让报告与
成品漂移。这里只做纯文本变换，不碰 docx 结构；渲染器负责把变换结果落回
物理 run，守卫负责把同一变换抹平后再逐字比对。

四条规则的口径（2026-11-15 落地）：

- N-02 正文内连续空格压缩：两个及以上半角空格/制表符压成一个半角空格。
  全角空格 U+3000 不动——它在中文稿里常被当作有意加宽（「摘　要」）。
- N-06 数字与中文之间的多余空格删除：空白两侧都是中文、或一侧中文一侧
  数字时删除。中文与拉丁字母之间的空格仍然保留（「数字金融 A 股」不在
  白名单里），这是 F31 钉死的边界，不许放宽。
- N-03 中文语境半角标点转全角：逗号、分号、冒号、问号、叹号、圆括号，
  以及句末句点。只在中文语境判定成立时转换——数字之间（1,000 / 12:30）
  与西文单词之间（Smith, 2020）一律不动。
- N-04 英文语境全角标点转半角：西文相邻的全角逗号/分号/冒号/句点/问号/
  叹号/圆括号转半角，句读符号后补一个半角空格（`(Smith，2020)` →
  `(Smith, 2020)`）；全角小数点两侧都是数字时只转字符、不补空格。
  只要有一侧紧邻中文就不转，中文语境归 N-03 管。

所有函数都是幂等的：同一段文字跑第二遍不再产生变化，渲染重跑与守卫
二次抹平都不会漂移。
"""

import difflib
import re

# 汉字与中文标点所在的常用区段。只认这几段，避免把日文假名、韩文
# 也当前文语境——经管稿件里出现假名多半是引文，不该被改标点。
CJK_RANGES = (
    (0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xF900, 0xFAFF),
    (0x3000, 0x303F),  # 中文标点（、。《》「」等）
    (0xFF00, 0xFFEF),  # 全角形式（含全角字母数字，但由 N-05 另管）
)


def is_cjk(ch):
    """该字符是否属于中文语境。空串、None 一律 False。"""
    if not ch:
        return False
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in CJK_RANGES)


def is_ascii_alnum(ch):
    return bool(ch) and ch.isascii() and ch.isalnum()


def _neighbor(text, i, step):
    """跳过空白后的相邻字符；越界返回空串。"""
    j = i + step
    while 0 <= j < len(text) and text[j] in " \t\u3000":
        j += step
    return text[j] if 0 <= j < len(text) else ""


# ---------------------------------------------------------------- N-02

_MULTI_SPACE = re.compile(r"[ \t]{2,}")


def collapse_spaces(text):
    """N-02：正文内连续空格（含制表符）压缩为一个半角空格。

    段首/段尾的空白不归本条管（N-01 段首/段尾删除单独成条、可单独驳回）。
    若这里也压，用户驳回「段尾空白删除」后段尾仍被压成一个空格，驳回就
    只生效了一半。
    """
    s = text or ""
    if not s.strip(" \t"):
        return s
    lead = re.match(r"[ \t]*", s).group(0)
    trail = re.search(r"[ \t]*$", s).group(0)
    mid = s[len(lead):len(s) - len(trail)]
    return lead + _MULTI_SPACE.sub(" ", mid) + trail


# ---------------------------------------------------------------- N-06

# 两侧都是中文/中文标点，或一侧中文一侧数字；全角空格不参与（见模块说明）。
# 中拉之间的空格必须保留：F31 钉死的边界，N-05 转换后的全角字母数字
# （现在是半角拉丁）也不在左/右侧字符集里。
_CN_CHAR = (r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
            r"\u3001-\u303f"          # 中文标点（、。《》「」等；不含 U+3000）
            r"\uff0c\uff1b\uff1a\uff1f\uff01\uff08\uff09\uff0e\uff1f")
_CJK_SPACE = re.compile(
    rf"(?<=[{_CN_CHAR}])[ \t]+(?=[{_CN_CHAR}0-9])"
    rf"|(?<=[0-9])[ \t]+(?=[{_CN_CHAR}])"
)


def drop_cjk_spaces(text):
    """N-06：删除中文与中文、中文与数字之间的多余空格。"""
    return _CJK_SPACE.sub("", text or "")


# ---------------------------------------------------------------- N-03

_CN_HALF = {",": "，", ";": "；", ":": "：", "?": "？", "!": "！"}


def cn_punct(text):
    """N-03：中文语境下的半角标点转全角。"""
    s = text or ""
    out = list(s)
    for i, ch in enumerate(s):
        if ch in _CN_HALF:
            prev, nxt = _neighbor(s, i, -1), _neighbor(s, i, 1)
            # 数字之间是千分位/时刻，西文之间是英文标点：都不转
            if is_ascii_alnum(prev) and is_ascii_alnum(nxt):
                continue
            if prev.isdigit() and nxt.isdigit():
                continue
            # 西文单词后紧跟中文：`Note: 研究` 的冒号属英文语境。数字后
            # 紧跟中文不在此列——「２０２０, 年」转半角后逗号仍是中文标点。
            if prev.isascii() and prev.isalpha() and not is_ascii_alnum(nxt):
                continue
            if is_cjk(prev) or is_cjk(nxt):
                out[i] = _CN_HALF[ch]
        elif ch == ".":
            prev, nxt = _neighbor(s, i, -1), _neighbor(s, i, 1)
            if is_cjk(prev) and (not nxt or is_cjk(nxt)
                                 or nxt in "”’\"'）】」』"):
                out[i] = "。"
        elif ch == "(":
            if is_cjk(_neighbor(s, i, 1)):
                out[i] = "（"
        elif ch == ")":
            if is_cjk(_neighbor(s, i, -1)):
                out[i] = "）"
    return "".join(out)


# ---------------------------------------------------------------- N-04

_EN_FULL = {"，": ",", "；": ";", "：": ":", "．": ".",
            "？": "?", "！": "!", "（": "(", "）": ")"}


def _ascii_context(prev, nxt):
    """西文语境：两侧都不紧邻中文，且至少一侧是 ASCII 字母/数字。"""
    if is_cjk(prev) or is_cjk(nxt):
        return False
    return is_ascii_alnum(prev) or is_ascii_alnum(nxt)


def en_punct(text):
    """N-04：英文语境下的全角标点转半角，句读后补一个空格。"""
    s = text or ""
    out = []
    pending_space = False
    for i, ch in enumerate(s):
        if ch in _EN_FULL:
            prev, nxt = _neighbor(s, i, -1), _neighbor(s, i, 1)
            if not _ascii_context(prev, nxt):
                out.append(ch)
                continue
            if ch == "（":
                out.append("(")
                continue
            if ch == "）":
                out.append(")")
                continue
            # 全角小数点两侧都是数字：只转字符，不补空格
            if ch == "．" and prev.isdigit() and nxt.isdigit():
                out.append(".")
                continue
            out.append(_EN_FULL[ch])
            if ch in ("，", "；", "：", "？", "！") or ch == "．":
                pending_space = True
            continue
        if pending_space:
            pending_space = False
            if ch not in " \t" and ch not in "),.;:?!":
                out.append(" ")
        out.append(ch)
    return "".join(out)


# ---------------------------------------------------------------- 组合

ACTIONS = (
    ("中文标点半角转全角", cn_punct),
    ("英文标点全角转半角", en_punct),
    ("正文连续空格压缩", collapse_spaces),
    ("数字与中文间空格删除", drop_cjk_spaces),
)
ACTION_NAMES = {name for name, _fn in ACTIONS}


def apply_action(text, action):
    """按动作名套一次归一；未知动作返回原文（渲染器不静默改字）。"""
    for name, fn in ACTIONS:
        if name == action:
            return fn(text or "")
    return text or ""


def normalize(text, include_spaces=True):
    """按渲染器的固定顺序套归一。

    include_spaces=False 用于关键词块：目标刊的分隔符本身可能就是两个空格
    （如 gjjryj 的 `"  "`），N-02 把它压成一个后再按两空格切词就切不开，
    分隔符会被整段吞掉。关键词只做标点与全角字母数字归一。
    """
    s = text or ""
    for name, fn in ACTIONS:
        if not include_spaces and name in ("正文连续空格压缩", "数字与中文间空格删除"):
            continue
        s = fn(s)
    return s


def edits_for(before, after):
    """把 before→after 的差异拆成 [(start, end, 新文字)]，供渲染器落回 run。"""
    if before == after:
        return []
    sm = difflib.SequenceMatcher(None, before, after, autojunk=False)
    return [(i1, i2, after[j1:j2])
            for tag, i1, i2, j1, j2 in sm.get_opcodes() if tag != "equal"]
