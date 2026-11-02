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

PAT = {
    # 内容提要／提要 是「内容摘要／摘要」的常见变体（《经济研究》即用「内容提要」）。
    # 这里必须与 render_from_ir 的标签归一正则保持一致，否则解析认不出、
    # 渲染就进不了摘要分支：字号不套、字数不查、标签也不归一。
    "abstract": re.compile(r"^(内容摘要|内容提要|摘\s*要|提\s*要)\s*[:：]"),
    "keywords": re.compile(r"^关\s*键\s*词\s*[:：]"),
    "classification": re.compile(r"^(中图分类号|JEL分类号|JEL|文献标识码)"),
    "ref_head": re.compile(r"^参\s*考\s*文\s*献\s*[:：]?$"),
    "table_caption": re.compile(r"^表\s*\d+"),
    "figure_caption": re.compile(r"^图\s*\d+"),
    "note": re.compile(r"^(注|资料来源|数据来源)\s*[:：]"),
    "h1": re.compile(r"^[一二三四五六七八九十]{1,3}、"),
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

# 标题不用句读标点。数字编号标题全靠这一条与正文短句区分（如「3 个案例」）。
# 冒号不算在内——「3 分析框架：科研自主性分级」是正常标题写法。
HEADING_PUNCT = re.compile(r"[。；，！？]")


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

    def walk(node):
        for child in node:
            if child.tag == W + "txbxContent":
                continue
            if child.tag == W + "t":
                parts.append(child.text or "")
            walk(child)

    walk(p)
    return "".join(parts).strip()


def has_omml(p):
    return next(p.iter(M + "oMath"), None) is not None


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
    if PAT["h2"].match(text):
        return (2, "high")

    if digit_ratio(text) > 0.3:
        return None
    if PAT["h4"].match(text):
        return (4, "high")
    if PAT["h3"].match(text):
        # 三级标题只在已出现二级标题之后可信，否则可能是数据行
        return (3, "medium")
    return None


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
             if b["type"] in AI_UPGRADE_FROM or b.get("confidence") in ("low", "medium")]
    if not cands:
        return {"used": False, "reason": "规则已给出确定判断，无需补漏"}

    prop = ai_assist.propose_types(blocks, cands, cfg, log=log)
    if not prop:
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
        return {"used": False, "reason": "指纹校验失败：AI 路径动到了文本，整批已回滚"}

    applied = [i for i in items if i["status"] == "applied"]
    log(f"AI 补漏：候选 {len(cands)} 块，提案 {len(prop)} 条，采纳 {len(applied)} 条")
    return {"used": True, "model": cfg.get("model"),
            "candidates": len(cands), "proposed": len(prop),
            "applied": len(applied), "items": items}


def parse(path, ai_cfg=None):
    z = zipfile.ZipFile(path)
    doc = xmlsafe.fromstring(z.read("word/document.xml"))
    body = doc.find(W + "body")

    blocks = []
    seen_level = set()
    in_references = False
    unsupported_content = 0

    def emit(block):
        if sdt_pos is not None:
            block["sdt_pos"] = sdt_pos    # sdt 内容块：同槽内的序号（F24）
        blocks.append(block)
        return block

    # idx 对 body 全部子元素统一编号（dense）：bookmarkStart、sdt 等异类
    # 子元素也占一个号。消费端（render_from_ir / anonymize / apply_details）
    # 一律按 body 子元素位置定位，生产端只数 w:p/w:tbl 的话，body 里一旦
    # 出现位于块元素之前的书签等元素，双方就整体错位——套版错段、改错正文、
    # 匿名路径漏删（F08）。sdt 内容控件也占自己的槽号；其内的 w:p/w:tbl
    # 下钻为块、共用该槽号并带 sdt_pos 序号（F24）——不重新编号，按 body
    # 子元素位置定位的契约不破。下不出来的（里面没有段落/表格）记数进 IR；
    # 外层下钻了、内嵌套的内容控件（sdt 套 sdt）没有——其内段落不进 IR，
    # 也逐个记数进 IR，不让「这里没处理」变成静默（F24）。只数 sdtContent
    # 的直接子级：段落/表格内部的行内 sdt 文字已随段落文本处理，不算未处理块。
    # 标题样式级别整体偏移（文档标题占用了 Heading 1 时）；规则识别的兜底仍走文本
    offset = style_offset(body)
    children = []
    for slot, child in enumerate(body):
        if child.tag == W + "sdt":
            content = child.find(W + "sdtContent")
            inner = [e for e in (content if content is not None else [])
                     if e.tag in (W + "p", W + "tbl")]
            for pos, e in enumerate(inner):
                children.append((slot, e, pos))
            if not inner:
                unsupported_content += 1
            else:
                unsupported_content += sum(
                    1 for e in content if e.tag == W + "sdt")
            continue
        if child.tag in (W + "tbl", W + "p"):
            children.append((slot, child, None))

    for idx, child, sdt_pos in children:
        if child.tag == W + "tbl":
            text = "".join(t.text or "" for t in child.iter(W + "t")).strip()
            rows = child.findall(W + "tr")
            cols = len(rows[0].findall(W + "tc")) if rows else 0
            # 伪表格：1 行且列数很少、文字极短，实际是公式加编号
            if len(rows) == 1 and cols <= 3 and len(text) <= 60:
                emit({"type": "equation_pseudo_table", "text": text, "idx": idx,
                      "n_rows": len(rows), "n_cols": cols})
                continue
            emit({
                "type": "table", "idx": idx, "n_rows": len(rows), "n_cols": cols,
                "borders": table_borders(child),
                "header": [para_text(tc) for tc in rows[0].findall(W + "tc")] if rows else [],
                # 表内文字按行平铺进 cells（表格块本身没有 text 键）：声明
                # 「全文（含图表）」口径的刊统计时要数表格内文字，缺了它
                # 标注与计算就自相矛盾。图内文字在图片对象里，文本层没有，天然不计
                "cells": [para_text(tc) for tr in rows for tc in tr.findall(W + "tc")],
            })
            continue

        text = para_text(child)
        has_img = has_drawing(child)
        if not text and has_img:
            emit({"type": "figure", "text": "", "idx": idx})
            continue
        if not text:
            # 纯公式段落没有文字，但它是内容，不能像空行一样丢掉——
            # 丢了渲染器就不知道这里有个公式，也就不会给它上行距保护。
            if has_omml(child):
                emit({"type": "equation", "text": "", "idx": idx, "no_text": True})
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
                level, conf = h
                if level == 3 and 2 not in seen_level:
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
        emit(block)

    # 顺序不可换：规则先给基线 → AI 补漏（只改标签）→ 题注绑定。
    # 题注绑定放最后，因为 AI 可能刚把某个 body 认出是表题。
    trace = {"used": False, "reason": "未启用"}
    if ai_cfg is None:
        ai_cfg = ai_assist.load_config()
    if ai_assist.available(ai_cfg):
        trace = apply_ai(blocks, ai_cfg)
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
        # 仍不支持的 sdt 结构（整块没有可处理的段落/表格，或内嵌套了下钻
        # 不到的内容控件）记数进 IR（F24），渲染汇总据此提示，
        # 不让「这里没处理」静默
        ir["unsupported_content_blocks"] = unsupported_content
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
              f"（整块没有可处理的段落/表格，或内嵌了未下钻的内容控件，"
              f"对象原样保留）")

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
