"""从稿约（PDF/Word/TXT）抽期刊规则集草稿。

分工（对应 DESIGN §14.7 与 §15）
--------------------------------------------------------------------
**结构由代码保证，AI 只负责填值；AI 关着时也可以纯手工填。**

规则集有十几个区块、上百个字段，让模型直接吐完整 JSON，它必然漏字段、
改结构、编造键名，产物连渲染器都跑不起来。所以这里反过来做：

1. 先复制一份已知良好的基线规则集，结构天然完整
2. AI 只回答一张扁平的事实清单（约 45 个键，值 + 一句原文引文）
3. 引文必须在原文里逐字搜得到，搜不到整条丢弃——这是防编造的主闸
4. 没抽到的键保留基线值，但逐条记进 `assumptions` 并在复核页标红
5. 没有 AI 时生成 `mode: manual` 的空草稿，用户在复核页逐项填写；
   手工修正记进 `extraction.manual`，不伪造原文引文

产物永远能渲染，但**永远标 `unverified`**，必须人工核对后才允许发布。

用法
----
    python tools/extract_journal.py 稿约.pdf --id jryj --name 金融研究 -o /tmp/draft.json
    python tools/extract_journal.py 投稿须知.docx --id jryj --manual -o /tmp/draft.json
"""

import argparse
import json
import math
import re
import sys
from copy import deepcopy
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ai_assist  # noqa: E402
import ocr        # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
RULES_DIR = ROOT / "rules" / "journal"

MAX_PDF_CHARS = 24000

# 键 → (写入路径, 类型, 中文名)。类型只用于校验与格式化，不做业务判断。
FACTS = {
    # 刊物标识
    "journal_name_zh":     (("journal", "name_zh"), "str", "中文刊名"),
    "journal_name_en":     (("journal", "name_en"), "str", "英文刊名"),
    "journal_language":    (("journal", "language"), "enum:zh|en", "刊物语言"),
    "journal_issn":        (("journal", "issn"), "str", "ISSN"),
    "journal_cn_number":   (("journal", "cn_number"), "str", "CN 号"),
    "journal_sponsor":     (("journal", "sponsor"), "str", "主办单位"),
    # 版心
    "page_size":           (("page", "size"), "enum:A4|B5|16开", "纸张"),
    "margin_top_mm":       (("page", "margin_mm", "top"), "float", "上边距(mm)"),
    "margin_bottom_mm":    (("page", "margin_mm", "bottom"), "float", "下边距(mm)"),
    "margin_left_mm":      (("page", "margin_mm", "left"), "float", "左边距(mm)"),
    "margin_right_mm":     (("page", "margin_mm", "right"), "float", "右边距(mm)"),
    "columns":             (("page", "columns"), "int", "分栏数"),
    # 正文
    "body_font_cn":        (("typography", "body", "font_cn"), "str", "正文字体"),
    "body_font_latin":     (("typography", "body", "font_latin"), "str", "正文西文字体"),
    "body_size_pt":        (("typography", "body", "size_pt"), "float", "正文字号(pt)"),
    "body_indent_chars":   (("typography", "body", "first_line_indent_chars"), "float", "首行缩进(字)"),
    "body_align":          (("typography", "body", "align"), "enum:both|left|justify", "正文对齐"),
    # 题名与摘要
    "title_font_cn":       (("typography", "title", "font_cn"), "str", "标题字体"),
    "title_size_pt":       (("typography", "title", "size_pt"), "float", "标题字号(pt)"),
    "title_bold":          (("typography", "title", "bold"), "bool", "标题加粗"),
    "subtitle_size_pt":    (("typography", "subtitle", "size_pt"), "float", "副标题字号(pt)"),
    "abstract_font_cn":    (("typography", "abstract", "font_cn"), "str", "摘要字体"),
    "abstract_size_pt":    (("typography", "abstract", "size_pt"), "float", "摘要字号(pt)"),
    "title_en_size_pt":    (("typography", "title_en", "size_pt"), "float", "英文标题字号(pt)"),
    "title_en_bold":       (("typography", "title_en", "bold"), "bool", "英文标题加粗"),
    "footnote_size_pt":    (("typography", "footnote", "size_pt"), "float", "脚注字号(pt)"),
    # 标题体系
    "heading_numbering":   (("headings", "numbering"), "enum:cn-hierarchical|arabic|none", "标题编号体系"),
    "heading_max_level":   (("headings", "max_level"), "int", "标题最大层级"),
    "heading_level1_font": (("headings", "level1", "font_cn"), "str", "一级标题字体"),
    "heading_level1_size": (("headings", "level1", "size_pt"), "float", "一级标题字号(pt)"),
    "heading_level1_bold": (("headings", "level1", "bold"), "bool", "一级标题加粗"),
    "heading_level1_align": (("headings", "level1", "align"), "enum:center|left", "一级标题对齐"),
    # 编号体例
    "numbering_tables":    (("numbering", "tables"), "str", "表编号写法"),
    "numbering_figures":   (("numbering", "figures"), "str", "图编号写法"),
    "numbering_equations": (("numbering", "equations"), "str", "公式编号写法"),
    # 引用
    "citation_system":     (("citation", "system"), "enum:numbered|author-year", "引用制"),
    "citation_style_id":   (("citation", "style_id"), "str", "文献著录标准"),
    # 限制
    "limit_total":         (("limits", "total", "value"), "int", "总字数上限"),
    "limit_total_metric":  (("limits", "total", "metric"), "enum:A|B|C", "字数口径"),
    "limit_abstract":      (("limits", "abstract", "value"), "int", "摘要字数上限"),
    "limit_title_chars":   (("limits", "title_chars_max"), "int", "标题字数上限"),
    "limit_keywords_min":  (("limits", "keywords", "min"), "int", "关键词最少"),
    "limit_keywords_max":  (("limits", "keywords", "max"), "int", "关键词最多"),
    # 投稿形态
    "submission_anonymous": (("submission", "anonymous"), "bool", "是否匿名评审"),
    "author_page_required": (("submission", "author_page", "separate_page"), "bool", "是否需作者信息页"),
    # 表图
    "table_caption_font":  (("tables", "caption_font_cn"), "str", "表题字体"),
    "table_caption_size":  (("tables", "caption_size_pt"), "float", "表题字号(pt)"),
    "figure_caption_font": (("figures", "caption_font_cn"), "str", "图题字体"),
    "figure_caption_size": (("figures", "caption_size_pt"), "float", "图题字号(pt)"),
}

# 需要成组处理、不能直接写单一路径的键
GROUPED = {"body_line_mode", "body_line_value", "heading_levels"}

ENUM_HINT = """
- journal_language: zh 或 en
- page_size: A4 / B5 / 16开
- body_align: both（两端对齐）/ left / justify
- heading_level1_align: center（居中）/ left
- heading_numbering: cn-hierarchical（一、（一）1.）/ arabic（1 / 1.1 / 1.1.1）/ none
- citation_system: numbered（顺序编码制，正文标 [1]）/ author-year（作者年制，正文标（张三，2020））
- limit_total_metric: A（字符数不计空格）/ B（中文字符数+英文单词数，对应 Word 中文版『字数』）/ C（字符数计空格）
"""

EXTRACT_PROMPT = """你是期刊投稿规范的抽取器。输入是一篇期刊稿约（投稿须知）的正文，输出是一张事实清单。

规则：
1. 只报原文明确写了的事实。没写、写得含糊、需要推测的，一律返回 null，不要填「大概」「通常」。
2. 每个有值的事实必须附一句 evidence，从原文中**逐字复制**（不改标点、不补字）。程序会拿它回原文里检索，搜不到的条目会被整条丢弃。
3. 数值只写数字，不带单位（"五号"换算成 10.5，"小四"换成 12，"1.5 倍行距"的 value 写 1.5）。
4. 中文号数换算：初号42 / 小初36 / 一号26 / 小一24 / 二号22 / 小二18 / 三号16 / 小三15 / 四号14 / 小四12 / 五号10.5 / 小五9 / 六号7.5。
5. 页边距若原文给的是厘米，乘以 10 换成毫米。

可抽取的键（其余键一律不要输出）：
{keys}

枚举取值要求：
{enums}

只输出 JSON，不要 markdown 代码块，形如：
{{"facts":{{"body_size_pt":10.5}},"evidence":{{"body_size_pt":"正文用小五号宋体"}},"notes":"稿约未提及页边距"}}"""

# 截断后的降级重试提示（F52）：只要 facts、不要 evidence。没有引文的值走
# F15 的无引文通道进 unverified、不进规则集——宁缺勿编，也不整链报废。
DEGRADE_HINT = """

注意：上一次输出被截断，没有得到完整 JSON。这次改用最简格式，只输出
{{"facts":{{...}}}} 一个对象：每个键一个取值，不要 evidence、不要 notes、
不要任何解释文字。没有把握的键直接不要输出。"""


# ---------------------------------------------------------------- 取文

# 稿约可以是 PDF、Word（.docx）或纯文本。旧版只收 PDF，用户手里最常见的
# 其实是「投稿须知.docx」；.doc 不直接读——老二进制格式没有可靠的纯 Python
# 解析器，页面会明确提示先另存为 .docx。
SOURCE_SUFFIXES = {".pdf", ".docx", ".txt", ".md"}


def pdf_text(path, max_chars=MAX_PDF_CHARS, log=None):
    """取 PDF 文字：有文字层用文字层，没有才 OCR（见 tools/ocr.py）。

    扫描件不再拒收——这是稿约场景的常态，纸质须知扫描成 PDF 很常见。
    """
    return ocr.extract_pdf_text(path, max_chars=max_chars, log=log)


def _truncate_text(text, max_chars):
    """按字符数截断并给出 truncated 标记，口径与 PDF 取文一致。"""
    text = text or ""
    return text[:max_chars], len(text) > max_chars


def docx_text(path, max_chars=MAX_PDF_CHARS):
    """取 .docx 正文与表格文字。

    按文档 body 的子元素顺序读，段落和表格的相对位置不会丢——稿约常把
    「字体字号」放进表格，只读 paragraphs 会漏掉整张表。页眉页脚不计入：
    那里通常是刊名/页码，不是格式要求。
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))
    lines = []

    def table_text(table):
        rows = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
                for nested in cell.tables:
                    nested_text = table_text(nested)
                    if nested_text:
                        parts.append(nested_text)
                cells.append(" ".join(parts))
            if any(cells):
                rows.append("\t".join(cells))
        return "\n".join(rows)

    def walk(parent):
        for child in parent.element.body.iterchildren():
            if child.tag == qn("w:p"):
                text = Paragraph(child, parent).text.strip()
                if text:
                    lines.append(text)
            elif child.tag == qn("w:tbl"):
                text = table_text(Table(child, parent))
                if text:
                    lines.append(text)

    walk(doc)
    text, truncated = _truncate_text("\n".join(lines), max_chars)
    return text, {"method": "docx", "pages": None, "chars": len(text),
                  "truncated": truncated}


def txt_text(path, max_chars=MAX_PDF_CHARS):
    """取纯文本稿约。按 UTF-8（含 BOM）→ GB18030 → UTF-16 依次尝试。

    国内稿约 TXT 常见 GBK 编码；直接 utf-8 读会半路抛异常，不能让用户
    自己猜编码。解码成功但正文为空时由调用方拦下。
    """
    raw = Path(path).read_bytes()
    text = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("TXT 编码无法识别，请另存为 UTF-8 或 Word（.docx）")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text, truncated = _truncate_text(text, max_chars)
    return text, {"method": "txt", "pages": None, "chars": len(text),
                  "truncated": truncated}


def read_source(path, max_chars=MAX_PDF_CHARS, log=None):
    """按扩展名取稿约文字，返回 (text, meta)。"""
    suffix = Path(path).suffix.lower()
    if suffix == ".pdf":
        return pdf_text(path, max_chars=max_chars, log=log)
    if suffix == ".docx":
        return docx_text(path, max_chars=max_chars)
    if suffix in (".txt", ".md"):
        return txt_text(path, max_chars=max_chars)
    raise ValueError("只支持 PDF、Word（.docx）或 TXT 稿约")


def norm(s):
    """比对引文用：去空白与全半角差异，避免只因排版空格就判失败。"""
    s = str(s or "")
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    return re.sub(r"\s+", "", s)


# ---------------------------------------------------------------- 抽取

def _cast(value, kind):
    """按声明类型取值，不合法返回 (None, 原因)。"""
    if value is None:
        return None, "空值"
    try:
        if kind == "str":
            v = str(value).strip()
            return (v, "") if v else (None, "空串")
        if kind == "float":
            v = float(str(value).replace("pt", "").strip())
            # Infinity/1e400 能被 float 收下，写进 JSON 会变非严格 JSON
            if not math.isfinite(v):
                return None, "数值超出范围"
            return v, ""
        if kind == "int":
            v = float(str(value).strip())
            if not math.isfinite(v):
                return None, "数值超出范围"
            return int(v), ""
        if kind == "bool":
            if isinstance(value, bool):
                return value, ""
            s = str(value).strip().lower()
            if s in ("true", "是", "需要", "yes", "1"):
                return True, ""
            if s in ("false", "否", "不需要", "no", "0"):
                return False, ""
            return None, "不是布尔值"
        if kind.startswith("enum:"):
            opts = kind[5:].split("|")
            s = str(value).strip()
            return (s, "") if s in opts else (None, f"不在 {opts} 内")
    except (TypeError, ValueError, OverflowError):
        return None, "类型不符"
    return None, "未知类型"


def _displayable(v):
    """dropped/unverified 的 value 只作复核展示：非有限浮点转字符串，
    否则裸 Infinity 会写进草稿 JSON，产出非严格 JSON。"""
    if isinstance(v, float) and not math.isfinite(v):
        return str(v)
    return v


def _set_path(root, path, value):
    cur = root
    for k in path[:-1]:
        cur = cur.setdefault(k, {})
        if not isinstance(cur, dict):
            return False
    cur[path[-1]] = value
    return True


def _get_path(root, path):
    cur = root
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return None
        cur = cur[k]
    return cur


def ai_facts(text, cfg, log=None):
    """让模型抽事实。返回 (facts, evidence, notes, raw)。

    事实抽取要求 49+3 个键逐个输出值和逐字引文，输出体量大，默认的
    max_tokens 会把 JSON 截成半截——而这条链路原本没有重试，一次即失败
    （F52）。截断/非法 JSON 时降级重试一次：只要取值、不要引文，没有引文
    的值经 build() 进 unverified，不会混进规则集。
    """
    keys = "\n".join(f"- {k}: {meta[2]}" for k, meta in FACTS.items())
    keys += "\n- body_line_mode: exact（固定行距，给磅值）或 multiple（倍数行距）\n"
    keys += "- body_line_value: 行距数值（固定值给磅，倍数给倍数）\n"
    keys += "- heading_levels: 各级标题的编号写法数组，如 [\"一、\",\"（一）\",\"1.\"]"
    prompt = EXTRACT_PROMPT.format(keys=keys, enums=ENUM_HINT)

    messages = [
        {"role": "system", "content": prompt},
        {"role": "user", "content": "以下是稿约正文：\n\n" + text},
    ]
    try:
        content = ai_assist._chat(cfg, messages, max_tokens=8192)
        obj = json.loads(_json_blob(content))
    except ValueError:
        # JSONDecodeError 是 ValueError 子类；_json_blob 找不到花括号也抛 ValueError
        if log:
            log("事实抽取的输出被截断或不是合法 JSON，降级重试一次：只要取值，不要引文")
        degrade = [dict(messages[0], content=prompt + DEGRADE_HINT), messages[1]]
        content = ai_assist._chat(cfg, degrade, max_tokens=8192)
        obj = json.loads(_json_blob(content))
    facts = obj.get("facts") or {}
    ev = obj.get("evidence") or {}
    if not isinstance(facts, dict):
        facts = {}
    if not isinstance(ev, dict):
        ev = {}
    return facts, ev, str(obj.get("notes") or ""), content


def _json_blob(s):
    s = (s or "").strip()
    if s.startswith("```"):
        s = s.split("\n", 1)[1] if "\n" in s else s
        s = s.rstrip("`").rsplit("```", 1)[0]
    lo, hi = s.find("{"), s.rfind("}")
    if lo < 0 or hi <= lo:
        raise ValueError("模型没有返回 JSON 对象")
    return s[lo:hi + 1]


def baseline_rules(prefer=None):
    """结构基线：默认取一本已核验的中文规则集，任何字段缺失都由它兜住。

    prefer 可指定仓库里任意规则集 id（如英文刊），用于目标刊结构与默认
    基线差异较大时少改一些字段。名字先过 JID 格式，避免被当成路径。
    """
    if prefer and re.match(r"^[a-z0-9_]{2,24}$", str(prefer)):
        p = RULES_DIR / f"{prefer}.json"
        if p.exists():
            return deepcopy(json.loads(p.read_text(encoding="utf-8"))), prefer
    for name in ("gjjryj", "jryj", "cjyj"):
        p = RULES_DIR / f"{name}.json"
        if p.exists():
            return deepcopy(json.loads(p.read_text(encoding="utf-8"))), name
    cands = sorted(RULES_DIR.glob("*.json"))
    if not cands:
        raise RuntimeError("rules/journal 下没有任何规则集可供作为结构基线")
    return deepcopy(json.loads(cands[0].read_text(encoding="utf-8"))), cands[0].stem


def iter_leaves(node, path=()):
    """产出 (路径元组, 叶子值)：dict 逐层下钻，其余（标量、列表）整体算一片叶子。

    列表不当 dict 拆——blocks 是一整份区块清单、headings.levels 是一串编号，
    拆开既没法回填也没法人工核对，整片披露才有意义。
    """
    if isinstance(node, dict):
        for k, v in node.items():
            yield from iter_leaves(v, path + (k,))
    else:
        yield path, node


# 抽取会覆写的非 FACTS 路径（成组项）。继承披露的分界线靠它算：
# 路径在这里面 → 值来自本次抽取；不在 → 值与基线刊相同，必须逐条披露。
# 必须登记 iter_leaves 产出的叶子粒度：父路径永远匹配不上叶子，已抽取的
# 字段会被再报成「沿用基线」（出处失实、计数虚增）。
_GROUPED_PATHS = {
    "body_line": (
        ("typography", "body", "line_spacing", "mode"),
        ("typography", "body", "line_spacing", "value_pt"),
        ("typography", "body", "line_spacing", "value"),
        ("typography", "body", "line_spacing", "equivalent_exact_pt"),
    ),
    "heading_levels": (("headings", "levels"), ("headings", "max_level")),
}


def written_paths(used, name_zh=""):
    """本次抽取（含刊物标识覆写）实际落盘的路径集合。"""
    ps = {tuple(FACTS[k][0]) for k in used if k in FACTS}
    for gk, gpaths in _GROUPED_PATHS.items():
        if gk in used:
            ps.update(gpaths)
    ps.add(("journal", "id"))            # 刊物代号是用户指定的，永远不算继承
    if name_zh:
        ps.add(("journal", "name_zh"))   # 用户亲手填的刊名同上，否则复核页出假风险
    return ps


def manual_paths(manual):
    """手工修正实际落盘的路径集合（与 written_paths 同口径）。"""
    ps = set()
    for key, entry in (manual or {}).items():
        if key in FACTS:
            ps.add(tuple(FACTS[key][0]))
        elif key in _GROUPED_PATHS:
            ps.update(_GROUPED_PATHS[key])
        elif isinstance(entry, dict) and isinstance(entry.get("path"), (list, tuple)):
            # 结构开关（匿名/作者页/blocks[].required/statements.*.required）
            # 不是 FACTS 键，落盘路径由 entry.path 自带
            ps.add(tuple(entry["path"]))
    return ps


def collect_inherited(base, rules, written):
    """收集「未被本次覆写、且值仍与基线相同」的全部叶子字段。

    披露范围是全部叶子，不只 FACTS 的 49 个事实键——blocks、statements、
    personal_fields、表格边框这些结构字段也必须逐条可见，否则「沿用 N 项」
    会变成完整假象（F16）。
    """
    by_path = {m[0]: (k, m[2]) for k, m in FACTS.items()}
    inherited = []
    for path, val in iter_leaves(base):
        if not path or path[0] in ("verification", "assumptions"):
            continue                  # 草稿的核验区由本次重建，不算继承
        if path in written:
            continue
        if _get_path(rules, path) != val:
            continue                  # 已被本次抽取/手工修正覆写成别的值
        key, label = by_path.get(path, (None, None))
        inherited.append({
            "key": ".".join(str(p) for p in path),
            "label": label or ".".join(str(p) for p in path),
            "value": val,
            "section": path[0],
            "fact": key is not None,
        })
    return inherited


def derive_assumptions(inherited, base_id):
    """由继承披露生成 assumptions：事实键逐条，结构字段收一条汇总。

    渲染器把 assumptions 逐条当字符串转义，这里只能是字符串数组。
    """
    fact_inherited = [it for it in inherited if it["fact"]]
    n_struct = len(inherited) - len(fact_inherited)
    assumptions = [
        f"{it['label']} = {it['value']}（未从稿约抽到，沿用基线值，须人工核对）"
        for it in fact_inherited if it["value"] not in (None, "", [], {})
    ]
    if n_struct:
        assumptions.append(
            f"另有 {n_struct} 个结构字段（blocks 区块清单、statements、personal_fields、"
            f"tables 边框等）未经稿约逐项抽取，逐字段沿用基线《{base_id}》原值；"
            f"逐字段清单见 verification.extraction.inherited，发布前须人工核对")
    return assumptions


# 复核页手工修正表单的分组。分组只影响展示顺序，字段本身仍以 FACTS 为准；
# 这里多出的 body_line / heading_levels 是成组项，由 apply_manual 特殊处理。
FIELD_GROUPS = (
    ("刊物标识", ("journal_name_zh", "journal_name_en", "journal_language",
                  "journal_issn", "journal_cn_number", "journal_sponsor")),
    ("页面", ("page_size", "margin_top_mm", "margin_bottom_mm",
              "margin_left_mm", "margin_right_mm", "columns")),
    ("正文", ("body_font_cn", "body_font_latin", "body_size_pt",
              "body_indent_chars", "body_align", "body_line")),
    ("题名与摘要", ("title_font_cn", "title_size_pt", "title_bold",
                    "subtitle_size_pt", "abstract_font_cn", "abstract_size_pt",
                    "title_en_size_pt", "title_en_bold", "footnote_size_pt")),
    ("标题体系", ("heading_numbering", "heading_max_level", "heading_level1_font",
                  "heading_level1_size", "heading_level1_bold",
                  "heading_level1_align", "heading_levels")),
    ("编号体例", ("numbering_tables", "numbering_figures", "numbering_equations")),
    ("引用", ("citation_system", "citation_style_id")),
    ("字数限制", ("limit_total", "limit_total_metric", "limit_abstract",
                  "limit_title_chars", "limit_keywords_min", "limit_keywords_max")),
    ("投稿形态", ("submission_anonymous", "author_page_required")),
    ("表图", ("table_caption_font", "table_caption_size",
              "figure_caption_font", "figure_caption_size")),
)

# 枚举值的中文展示。没有登记的枚举直接显示原值，不猜含义。
ENUM_LABELS = {
    "journal_language": {"zh": "中文", "en": "英文"},
    "page_size": {"A4": "A4", "B5": "B5", "16开": "16 开"},
    "body_align": {"both": "两端对齐", "left": "左对齐", "justify": "两端对齐（justify）"},
    "heading_level1_align": {"center": "居中", "left": "左对齐"},
    "heading_numbering": {"cn-hierarchical": "一、（一）1.",
                          "arabic": "1 / 1.1 / 1.1.1", "none": "不编号"},
    "citation_system": {"numbered": "顺序编码制（[1]）",
                        "author-year": "作者-年份制（（张三，2020））"},
    "limit_total_metric": {"A": "A：字符数（不计空格）",
                           "B": "B：中文字符数 + 英文单词数",
                           "C": "C：字符数（计空格）"},
}


def _field_options(key, kind):
    """枚举字段的可选项；布尔、文本、数值由模板按 kind 单独渲染。"""
    if kind.startswith("enum:"):
        labels = ENUM_LABELS.get(key) or {}
        return [{"value": v, "label": labels.get(v, v)}
                for v in kind[5:].split("|")]
    return []


def editable_fields(rules):
    """把当前规则集拆成复核页可编辑的分组字段。

    只暴露稿约最常规定、且渲染器真正消费的字段；blocks/statements 等
    结构字段仍按基线继承并在「沿用基线」清单里逐条披露，不在本表单里
    伪装成可安全手改。
    """
    groups = []
    for title, keys in FIELD_GROUPS:
        fields = []
        for key in keys:
            if key == "body_line":
                ls = _get_path(rules, ("typography", "body", "line_spacing")) or {}
                mode = ls.get("mode") or "multiple"
                value = ls.get("value_pt") if mode == "exact" else ls.get("value")
                fields.append({"key": "body_line", "label": "行距",
                               "kind": "line", "mode": mode, "value": value})
                continue
            if key == "heading_levels":
                levels = _get_path(rules, ("headings", "levels")) or []
                fields.append({"key": "heading_levels", "label": "标题编号（每行一级）",
                               "kind": "levels",
                               "value": "\n".join(str(x) for x in levels)})
                continue
            path, kind, label = FACTS[key]
            field = {"key": key, "label": label, "kind": kind,
                     "value": _get_path(rules, path)}
            options = _field_options(key, kind)
            if options:
                field["options"] = options
            fields.append(field)
        groups.append({"title": title, "fields": fields})
    return groups


def build(extracted, evidence, text, *, jid, name_zh="", source="", meta=None,
          manual=False, baseline_id=""):
    """把事实清单映射成完整规则集。返回 (rules, report)。

    manual=True 时 facts/evidence 通常为空：不调用 AI，只从结构基线出发，
    所有字段先按「沿用基线」如实披露，用户在复核页逐项修正。
    baseline_id 可指定仓库里另一本规则集作结构基线。
    """
    base, base_id = baseline_rules(baseline_id or None)
    rules = deepcopy(base)
    used = {}
    dropped = []
    unverified = []

    for key, raw in extracted.items():
        if key in GROUPED:
            continue
        if key not in FACTS:
            continue
        path, kind, label = FACTS[key]
        val, why = _cast(raw, kind)
        if val is None:
            dropped.append({"key": key, "label": label,
                            "value": _displayable(raw), "why": why})
            continue
        quote = str(evidence.get(key) or "").strip()
        if not quote:
            # 无引文的事实不能进规则集——防编造主闸不允许「没引文照写入」的旁路
            unverified.append({"key": key, "label": label, "value": val,
                               "why": "模型未给出原文引文，无法核验"})
            continue
        if norm(quote) not in norm(text):
            # 引文在原文里搜不到 —— 判为编造，整条丢弃
            dropped.append({"key": key, "label": label, "value": raw,
                            "why": f"引文在原文中找不到：{quote[:40]}"})
            continue
        if _set_path(rules, path, val):
            used[key] = {"label": label, "value": val, "evidence": quote}

    # 成组项：与单值字段过同一道闸——没引文不落盘，引文对不上判编造丢弃
    mode = str(extracted.get("body_line_mode") or "").strip()
    value = extracted.get("body_line_value")
    if mode in ("exact", "multiple") and value not in (None, ""):
        quote = str(evidence.get("body_line_value")
                    or evidence.get("body_line_mode") or "").strip()
        if not quote:
            unverified.append({"key": "body_line", "label": "行距",
                               "value": f"{mode} {value}",
                               "why": "模型未给出原文引文，无法核验"})
        elif norm(quote) not in norm(text):
            dropped.append({"key": "body_line_value", "label": "行距", "value": value,
                            "why": f"引文在原文中找不到：{quote[:40]}"})
        else:
            try:
                v = float(value)
            except (TypeError, ValueError):
                v = None
            if v is None or not math.isfinite(v):
                # 成组项绕过 _cast 直接 float()，inf/1e400 在这里同样拦下，
                # 不能让非有限值经引文闸写进规则集
                dropped.append({"key": "body_line_value", "label": "行距",
                                "value": value,
                                "why": "不是数字" if v is None else "数值超出范围"})
            else:
                size = rules["typography"]["body"].get("size_pt") or 10.5
                ls = {"mode": mode}
                if mode == "exact":
                    ls["value_pt"] = v
                else:
                    ls["value"] = v
                    ls["equivalent_exact_pt"] = round(size * v, 1)
                rules["typography"]["body"]["line_spacing"] = ls
                used["body_line"] = {"label": "行距", "value": f"{mode} {v}",
                                     "evidence": quote}

    levels = extracted.get("heading_levels")
    if isinstance(levels, list) and levels and all(isinstance(x, str) and x.strip() for x in levels):
        quote = str(evidence.get("heading_levels") or "").strip()
        if not quote:
            unverified.append({"key": "heading_levels", "label": "标题编号",
                               "value": levels,
                               "why": "模型未给出原文引文，无法核验"})
        elif norm(quote) not in norm(text):
            dropped.append({"key": "heading_levels", "label": "标题编号",
                            "value": levels,
                            "why": f"引文在原文中找不到：{quote[:40]}"})
        else:
            rules["headings"]["levels"] = [x.strip() for x in levels]
            rules["headings"]["max_level"] = min(len(levels), 4)
            used["heading_levels"] = {"label": "标题编号", "value": rules["headings"]["levels"],
                                      "evidence": quote}
            if "heading_max_level" in used:
                # 编号写法派生的层级数覆写了单值事实的落盘值：used 里留着
                # 模型读到的旧值，复核页证据表会与草稿实际值矛盾——改写成
                # 生效值并注明原因，证据引文本身不动
                used["heading_max_level"]["value"] = rules["headings"]["max_level"]
                used["heading_max_level"]["why"] = (
                    "与标题编号同时抽出，最大层级按编号写法的条数取 "
                    f"{rules['headings']['max_level']}")

    # 刊物标识：以用户在页面上填的为准，AI 抽到的只作补充
    rules["journal"]["id"] = jid
    if name_zh:
        rules["journal"]["name_zh"] = name_zh
    rules.setdefault("journal", {})["name_zh"] = rules["journal"].get("name_zh") or name_zh or jid

    # 继承披露：基线里每个未被本次覆写、且值原样保留的叶子，逐条列出。
    # 只披露 FACTS 键会漏掉 blocks 标签、statements、personal_fields、tables
    # 边框等几百个静默继承的字段，「沿用 N 项」就成了完整假象。
    written = written_paths(used, name_zh)
    inherited = collect_inherited(base, rules, written)

    rules["verification"] = {
        "status": "unverified",
        # 渲染器直接下标取 source_url，缺键会抛 KeyError；草稿阶段无官方链接，留空串
        "source_url": "",
        "template_url": "",
        "template_date": "",
        "verified_at": "",
        "verified_by": ("手工建库，未经稿约抽取" if manual
                        else "AI 从稿约抽取，未经人工核对"),
        "notes": ((f"手工建库（来源 {source}）；{len(inherited)} 项沿用基线值，"
                   f"{len(dropped)} 项被丢弃，{len(unverified)} 项未证实未采用")
                  if manual else
                  (f"抽取自 {source}；{len(used)} 项有原文依据，"
                   f"{len(inherited)} 项沿用基线值，{len(dropped)} 项被丢弃，"
                   f"{len(unverified)} 项未证实未采用")),
        "typography_authority": "guide",
        "extraction": {
            "source_file": source,
            "extracted_at": date.today().isoformat(),
            "baseline": base_id,
            "mode": "manual" if manual else "ai",
            "pages": (meta or {}).get("pages"),
            "truncated": (meta or {}).get("truncated", False),
            # 文字从哪来的必须留痕：OCR 一定有错字，复核时要区别对待
            "text_method": (meta or {}).get("method", "text_layer"),
            "ocr_pages": (meta or {}).get("ocr_pages") or [],
            # 哪些页识别失败必须留痕——整页内容静默丢失比读错更严重
            "ocr_failed_pages": (meta or {}).get("ocr_failed_pages") or [],
            # 文字层异常少的页（F51）：20 字阈值的旁路可能让扫描页静默跳过
            # OCR，复核时要知道哪些页可疑
            "ocr_low_text_pages": (meta or {}).get("ocr_low_text_pages") or [],
            "ocr_engine": (meta or {}).get("engine", ""),
            "ai_model": None,
            "used": {k: v for k, v in used.items()},
            "manual": {},
            "inherited": inherited,
            "dropped": dropped,
            "unverified": unverified,
        },
    }
    rules["assumptions"] = derive_assumptions(inherited, base_id)

    report = validate(rules, used=used, inherited=inherited, dropped=dropped)
    report["stats"] = {"extracted": len(used), "inherited": len(inherited),
                       "manual": 0, "dropped": len(dropped),
                       "unverified": len(unverified)}
    return rules, report


def _line_value(spacing, mode):
    """从 line_spacing 取当前值：固定值取 value_pt，倍数取 value。"""
    if mode == "exact":
        return spacing.get("value_pt")
    if mode == "multiple":
        return spacing.get("value")
    return None


# 结构要件（blocks[].required / statements / 匿名）的中文名。与 compliance_report
# 的展示口径保持一致；只用于复核页，不参与渲染。
STRUCT_BLOCK_LABELS = {
    "title": "标题", "subtitle": "副标题", "authors": "作者", "affiliations": "单位",
    "corresponding_author": "通讯作者", "abstract": "摘要", "keywords": "关键词",
    "classification_code": "分类号", "document_code": "文献标识码",
    "article_number": "文章编号", "body": "正文", "references": "参考文献",
    "appendix": "附录", "acknowledgement": "致谢", "highlights": "Highlights",
    "received_date": "收稿日期", "author_note": "作者简介", "endnotes": "文末尾注",
    "conflict_of_interest": "利益冲突声明", "data_availability": "数据可得性",
    "english_title": "英文标题", "english_abstract": "英文摘要",
    "english_keywords": "英文关键词", "english_authors": "英文作者",
    "english_affiliations": "英文单位",
}
STRUCT_STATEMENT_LABELS = {
    "copyright_transfer": "版权转让协议",
    "conflict_of_interest": "利益冲突声明",
    "data_availability": "数据可得性声明",
    "ai_declaration": "AI 使用声明",
    "originality": "原创性声明",
    "confidentiality": "保密审查",
}


def structural_fields(rules):
    """复核页的结构开关：匿名、作者页、区块必备/可选、声明必备/可选。

    这些字段原先只按基线继承、在「沿用基线」清单里逐条披露，用户想改只能
    手改 draft.json。目标刊结构与基线不同很常见（要不要英文摘要、要不要
    作者信息页、声明是不是必备），这里做成可视开关；保存仍走同一条
    apply_manual 事务，改动计入「手工修正」并从继承披露里移出。
    """
    sub = rules.get("submission") or {}
    blocks = []
    for i, b in enumerate(rules.get("blocks") or []):
        if not isinstance(b, dict) or not b.get("type"):
            continue
        lang = b.get("lang")
        label = STRUCT_BLOCK_LABELS.get(b["type"], b.get("type"))
        blocks.append({
            "index": i, "type": b["type"], "lang": lang or "",
            "label": label + (f"（{lang}）" if lang else ""),
            "required": bool(b.get("required")),
        })
    statements = []
    for key, conf in (rules.get("statements") or {}).items():
        conf = conf if isinstance(conf, dict) else {}
        statements.append({
            "key": key, "label": STRUCT_STATEMENT_LABELS.get(key, key),
            "required": bool(conf.get("required")),
        })
    ap = (sub.get("author_page") or {}).get("separate_page")
    return {
        "anonymous": bool(sub.get("anonymous")),
        "author_page": bool(ap),
        "blocks": blocks,
        "statements": statements,
    }


def _apply_structural(rules, structural, manual):
    """把结构开关并入草稿；与 apply_manual 同一事务，出错前不落盘。

    structural 是「开关名 → 期望值」的完整映射（未提交的复选框由调用方
    填 False）。只处理与当前值不同的项，避免一次保存把所有结构字段都记成
    「手工修正」、清零继承披露。
    """
    def record(key, label, value, previous, path):
        manual[key] = {"label": label, "value": value,
                       "previous": _displayable(previous), "path": list(path)}

    sub = rules.setdefault("submission", {})
    if "anonymous" in structural:
        old = bool(sub.get("anonymous"))
        new = bool(structural["anonymous"])
        if old != new:
            sub["anonymous"] = new
            record("structural:anonymous", "正文匿名（双盲）", new, old,
                   ("submission", "anonymous"))
    if "author_page" in structural:
        ap = sub.setdefault("author_page", {})
        old = bool(ap.get("separate_page"))
        new = bool(structural["author_page"])
        if old != new:
            ap["separate_page"] = new
            record("structural:author_page", "作者信息单独成页", new, old,
                   ("submission", "author_page", "separate_page"))
    for i, want in (structural.get("blocks") or {}).items():
        blocks = rules.get("blocks") or []
        if not (0 <= int(i) < len(blocks)) or not isinstance(blocks[int(i)], dict):
            continue
        b = blocks[int(i)]
        old = bool(b.get("required"))
        new = bool(want)
        if old == new:
            continue
        b["required"] = new
        label = STRUCT_BLOCK_LABELS.get(b.get("type"), b.get("type")) + "·必备"
        record(f"structural:blocks:{i}", label, new, old, ("blocks", int(i), "required"))
    for key, want in (structural.get("statements") or {}).items():
        st = rules.setdefault("statements", {})
        conf = st.get(key)
        if not isinstance(conf, dict):
            continue
        old = bool(conf.get("required"))
        new = bool(want)
        if old == new:
            continue
        conf["required"] = new
        label = STRUCT_STATEMENT_LABELS.get(key, key) + "·必备"
        record(f"structural:statements:{key}", label, new, old,
               ("statements", key, "required"))


def apply_manual(rules, values, structural=None):
    """把复核页提交的字段修正并入草稿。返回 (rules, report, errors)。

    表单每次提交全部字段，所以这里只处理**与当前值不同**的字段：若把所有
    字段都记成「手工修正」，「沿用基线」披露会瞬间清零，也就失去了复核页
    存在的意义。

    有错误时返回原 rules 与错误清单，调用方不得落盘半份修改。
    """
    out = deepcopy(rules)
    verification = out.setdefault("verification", {})
    ext = verification.setdefault("extraction", {})
    # 手工修正后的继承披露必须对回这份草稿自己的基线；固定用默认基线会
    # 让选过英文/其他基线的草稿把「与基线相同」判错。
    base, base_id = baseline_rules(ext.get("baseline") or None)
    raw_used = ext.get("used")
    used = dict(raw_used) if isinstance(raw_used, dict) else {}
    raw_manual = ext.get("manual")
    manual = dict(raw_manual) if isinstance(raw_manual, dict) else {}
    unverified = list(ext.get("unverified") or [])
    errors = []
    body_size_changed = False

    def take_evidence(key):
        item = used.pop(key, None)
        return item.get("evidence", "") if isinstance(item, dict) else ""

    def drop_unverified(key):
        return [u for u in unverified
                if not (isinstance(u, dict) and u.get("key") == key)]

    for key, (path, kind, label) in FACTS.items():
        raw = values.get(key)
        if raw is None:
            continue
        if isinstance(raw, str) and not raw.strip():
            continue                  # 留空 = 不改，不拿空串清掉当前值
        val, why = _cast(raw, kind)
        if val is None:
            errors.append(f"{label}：{why}")
            continue
        old = _get_path(out, path)
        if old == val:
            continue                  # 未改动，不记「手工修正」
        if not _set_path(out, path, val):
            errors.append(f"{label}：字段路径不可写")
            continue
        evidence = take_evidence(key)
        unverified = drop_unverified(key)
        entry = {"label": label, "value": val, "previous": _displayable(old)}
        if evidence:
            entry["evidence"] = evidence
        manual[key] = entry
        if key == "body_size_pt":
            body_size_changed = True

    if "heading_levels" in values:
        raw = values.get("heading_levels")
        levels = []
        if isinstance(raw, str):
            for line in raw.replace("|", "\n").splitlines():
                part = line.strip()
                if part:
                    levels.append(part)
        if not levels:
            errors.append("标题编号：至少保留一级（每行一级，如 一、/（一）/1.）")
        elif len(levels) > 8:
            errors.append("标题编号：最多 8 级")
        elif any(len(x) > 20 for x in levels):
            errors.append("标题编号：单级编号不超过 20 个字符")
        else:
            current = _get_path(out, ("headings", "levels")) or []
            if levels != current:
                evidence = take_evidence("heading_levels")
                unverified = drop_unverified("heading_levels")
                out.setdefault("headings", {})["levels"] = levels
                out["headings"]["max_level"] = min(len(levels), 4)
                entry = {"label": "标题编号", "value": levels,
                         "previous": _displayable(current)}
                if evidence:
                    entry["evidence"] = evidence
                manual["heading_levels"] = entry

    if "body_line_mode" in values or "body_line_value" in values:
        mode = str(values.get("body_line_mode") or "").strip()
        if mode:
            if mode not in ("exact", "multiple"):
                errors.append("行距：模式只能是固定值或倍数")
            else:
                spacing = _get_path(out, ("typography", "body", "line_spacing")) or {}
                raw_value = str(values.get("body_line_value") or "").strip()
                if not raw_value:
                    current_value = _line_value(spacing, mode)
                    raw_value = "" if current_value is None else str(current_value)
                try:
                    value = float(raw_value)
                except (TypeError, ValueError):
                    value = None
                if value is None or not math.isfinite(value):
                    errors.append("行距：数值无效")
                elif value <= 0:
                    errors.append("行距：必须大于 0")
                else:
                    old_mode = spacing.get("mode")
                    old_value = _line_value(spacing, old_mode) if old_mode else None
                    if old_mode != mode or old_value != value:
                        evidence = take_evidence("body_line")
                        unverified = drop_unverified("body_line")
                        replacement = {"mode": mode}
                        if mode == "exact":
                            replacement["value_pt"] = value
                        else:
                            size = ((out.get("typography") or {}).get("body") or {}) \
                                .get("size_pt") or 10.5
                            replacement["value"] = value
                            replacement["equivalent_exact_pt"] = round(size * value, 1)
                        out.setdefault("typography", {}).setdefault("body", {})[
                            "line_spacing"] = replacement
                        entry = {"label": "行距", "value": f"{mode} {value:g}",
                                 "previous": _displayable(spacing)}
                        if evidence:
                            entry["evidence"] = evidence
                        manual["body_line"] = entry

    # 字号改了、行距是倍数制：折算磅值要跟着变，否则规格页与审计的依据
    # 还是旧字号算出来的数。
    if body_size_changed:
        body = (out.get("typography") or {}).get("body") or {}
        spacing = body.get("line_spacing") or {}
        if spacing.get("mode") == "multiple" and spacing.get("value"):
            spacing["equivalent_exact_pt"] = round(
                float(body.get("size_pt") or 10.5) * float(spacing["value"]), 1)

    # 结构开关（匿名/作者页/区块必备/声明必备）与字段修正同一事务：
    # 有错时整单不落盘，草稿不会出现「字段改了、结构没改」的半份状态。
    if structural is not None:
        _apply_structural(out, structural, manual)

    if errors:
        return rules, None, errors

    written = written_paths(used,
                            name_zh=(out.get("journal") or {}).get("name_zh") or "")
    written |= manual_paths(manual)
    inherited = collect_inherited(base, out, written)

    ext["used"] = used
    ext["manual"] = manual
    ext["unverified"] = unverified
    ext["inherited"] = inherited
    ext["edited_at"] = date.today().isoformat()
    ext.setdefault("mode", "manual" if not used else "ai")

    if manual and used:
        verified_by = "AI 抽取 + 人工修正，未经人工核对"
    elif manual:
        verified_by = "手工建库，未经稿约抽取"
    elif used:
        verified_by = "AI 从稿约抽取，未经人工核对"
    else:
        verified_by = "基线草稿，未经稿约核对"
    verification["verified_by"] = verified_by

    source = ext.get("source_file") or "未知来源"
    dropped = ext.get("dropped") or []
    verification["notes"] = (
        f"来源 {source}；{len(used)} 项有原文依据，{len(manual)} 项手工修正，"
        f"{len(inherited)} 项沿用基线值，{len(dropped)} 项被丢弃，"
        f"{len(unverified)} 项未证实未采用")
    out["assumptions"] = derive_assumptions(inherited, base_id)

    report = validate(out, used=used, inherited=inherited, dropped=dropped)
    report["stats"] = {"extracted": len(used), "inherited": len(inherited),
                       "manual": len(manual), "dropped": len(dropped),
                       "unverified": len(unverified)}
    return out, report, []


def validate(rules, *, used=None, inherited=None, dropped=None):
    """结构性校验。errors 不阻断保存，但阻断发布。"""
    errs, warns = [], []

    # blocks 是列表，其余是字典——混在一个循环里判会误报
    for k in ("journal", "page", "typography", "headings", "citation",
              "submission", "tables", "figures"):
        if not isinstance(rules.get(k), dict):
            errs.append({"field": k, "msg": f"缺少 {k} 区块"})
    if not isinstance(rules.get("blocks"), list):
        errs.append({"field": "blocks", "msg": "缺少 blocks 区块"})

    if not (rules.get("journal") or {}).get("name_zh"):
        errs.append({"field": "journal.name_zh", "msg": "缺中文刊名"})
    if not re.match(r"^[a-z0-9_]{2,24}$", str((rules.get("journal") or {}).get("id") or "")):
        errs.append({"field": "journal.id", "msg": "期刊代号只能用小写字母、数字、下划线"})
    if not (rules.get("blocks") or []):
        errs.append({"field": "blocks", "msg": "必备区块清单为空"})

    body = (rules.get("typography") or {}).get("body") or {}
    size = body.get("size_pt")
    if not isinstance(size, (int, float)) or not (5 <= size <= 72):
        errs.append({"field": "typography.body.size_pt", "msg": f"正文字号异常：{size}"})
    ls = body.get("line_spacing") or {}
    if ls.get("mode") == "exact" and not ls.get("value_pt"):
        errs.append({"field": "typography.body.line_spacing", "msg": "固定行距缺磅值"})
    if ls.get("mode") == "multiple" and not ls.get("value"):
        errs.append({"field": "typography.body.line_spacing", "msg": "倍数行距缺倍数"})

    mm = (rules.get("page") or {}).get("margin_mm") or {}
    for side in ("top", "bottom", "left", "right"):
        v = mm.get(side)
        if not isinstance(v, (int, float)) or not (5 <= v <= 80):
            warns.append({"field": f"page.margin_mm.{side}", "msg": f"{side} 边距异常：{v}"})

    lv = (rules.get("headings") or {}).get("levels") or []
    if not lv:
        errs.append({"field": "headings.levels", "msg": "标题编号体系为空"})

    kw = (rules.get("limits") or {}).get("keywords") or {}
    if kw.get("min") and kw.get("max") and kw["min"] > kw["max"]:
        warns.append({"field": "limits.keywords", "msg": "关键词下限大于上限"})

    for d in (dropped or []):
        warns.append({"field": d["label"], "msg": f"已丢弃：{d['why']}"})

    return {"errors": errs, "warnings": warns,
            "ok_to_publish": not errs,
            "needs_review": len(inherited or [])}


def extract(source, *, jid, name_zh="", cfg=None, log=None, manual=False,
            baseline_id=""):
    """完整流程：读稿约（PDF/Word/TXT）→ 抽事实（或手工模式）→ 建规则集。

    manual=True 时不调用 AI：读取原文后直接以结构基线起一份空草稿，所有
    字段先进「沿用基线」清单，由用户在复核页逐项修正。AI 未配置时也能
    走通建库，只是抽取项为 0。baseline_id 选择结构基线刊。
    """
    cfg = cfg or ai_assist.load_config()
    use_ai = bool(not manual and ai_assist.available(cfg))
    text, meta = read_source(source, log=log)
    if not (text or "").strip():
        raise RuntimeError("稿约里没有读到任何文字：请确认文件不是空文档，"
                           "扫描件请装 OCR 或改用文字版 PDF/Word/TXT")
    if use_ai:
        facts, ev, notes, _raw = ai_facts(text, cfg, log=log)
        rules, report = build(facts, ev, text, jid=jid, name_zh=name_zh,
                              source=Path(source).name, meta=meta,
                              baseline_id=baseline_id)
        rules["verification"]["extraction"]["ai_model"] = cfg.get("model")
        if notes:
            rules["verification"]["extraction"]["ai_notes"] = notes[:400]
    else:
        if not manual and log:
            log("AI 未配置或未启用，自动改走手工建库：先生成基线草稿，"
                "再在复核页逐项填写")
        rules, report = build({}, {}, text, jid=jid, name_zh=name_zh,
                              source=Path(source).name, meta=meta, manual=True,
                              baseline_id=baseline_id)
    return rules, report, text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", help="稿约文件：PDF / Word(.docx) / TXT")
    ap.add_argument("--id", required=True, help="期刊代号，小写字母，如 jryj")
    ap.add_argument("--name", default="", help="中文刊名")
    ap.add_argument("--manual", action="store_true",
                    help="不调用 AI，只建基线草稿（可在复核页手工填写）")
    ap.add_argument("--baseline", default="",
                    help="结构基线刊 id（如 tjyj）；缺省用已核验中文刊")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    try:
        rules, report, _ = extract(a.source, jid=a.id, name_zh=a.name,
                                   log=print, manual=a.manual,
                                   baseline_id=a.baseline)
    except Exception as e:  # noqa: BLE001
        print(f"失败：{e}")
        return 1

    Path(a.out).write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    ext_meta = rules["verification"]["extraction"]
    # OCR 丢页与未证实项都必须在命令行里喊出来，不能只在 JSON 里躺着
    for m in ocr.failure_warnings(ext_meta):
        print(m, file=sys.stderr)
    for m in ocr.low_text_warnings(ext_meta):
        print(m, file=sys.stderr)
    s = report["stats"]
    if s.get("unverified"):
        print(f"警告：{s['unverified']} 项模型给了值但没有原文引文（未证实），"
              f"未写入规则集，需人工核对（见草稿 verification.extraction.unverified）",
              file=sys.stderr)
    print(f"抽取 {s['extracted']} 项 / 手工修正 {s.get('manual', 0)} 项 / "
          f"沿用基线 {s['inherited']} 项 / 丢弃 {s['dropped']} 项 / "
          f"未证实 {s['unverified']} 项")
    for e in report["errors"]:
        print("  错误:", e["msg"])
    for w in report["warnings"][:10]:
        print("  提示:", w["msg"])
    print("可发布" if report["ok_to_publish"] else "不可发布：先修掉错误项")
    return 0 if report["ok_to_publish"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
