"""把稿件 IR 与目标期刊规则集对撞，生成投稿合规体检报告（HTML）。

这是产品的入口功能：先告诉作者"这篇稿子投这本刊缺什么"，
再谈自动重排。所有结论只出报告，不修改稿件。

用法：python tools/compliance_report.py <ir.json> <ruleset.json> -o report.html
"""

import argparse
import html
import json
import re
import sys
from pathlib import Path

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from report_theme import BASE

# 期刊 block 类型 → IR block 类型的映射
NEED_MAP = {
    "title": {"title", "front_matter", "english_title"},
    "authors": {"authors"},
    "affiliations": {"affiliations"},
    "abstract": {"abstract", "english_abstract"},
    "keywords": {"keywords", "english_keywords"},
    "classification_code": {"classification"},
    "document_code": {"classification"},
    "body": {"body"},
    "references": {"references_heading", "reference_item"},
    "appendix": {"appendix_heading"},
    "acknowledgement": {"table_note"},
}

LABEL = {
    "title": "标题", "subtitle": "副标题", "authors": "作者", "affiliations": "单位",
    "abstract": "摘要", "keywords": "关键词", "classification_code": "分类号",
    "document_code": "文献标识码", "article_number": "文章编号", "body": "正文",
    "references": "参考文献", "appendix": "附录", "acknowledgement": "致谢",
    "data_availability": "数据可得性", "conflict_of_interest": "利益冲突声明",
    "highlights": "Highlights",
}

# 规则集 limits.total.scope 是各刊投稿须知的原话（「正文」「全文」「全文（含图表）」）。
# 已认识的写法映射到统计过滤：声明「正文」就只数正文块，全文类数全部块；
# 没认出的写法一律保守按全文计，由 check() 在报告里注明回退——
# 绝不把全文数字偷偷对到正文上限上（那会出假超限判决）。
SCOPE_FILTER = {"正文": "body", "全文": "all", "全文（含图表）": "all"}

# 这些不是作者能补的（由编辑部编号或按刊物体例给定），不当成待办派给作者
NON_ACTIONABLE = {"article_number", "document_code"}

# ---- 必备声明（规则集 statements）----
# 23 本刊都声明了这一项，但报告里从来没有出现过。作者不会知道
# 「这本刊要 AI 使用声明」「那本刊投稿即视为签版权转让」。
STATEMENT_LABELS = {
    "copyright_transfer": "版权转让协议",
    "conflict_of_interest": "利益冲突声明",
    "data_availability": "数据可得性声明",
    "ai_declaration": "AI 使用声明",
    "originality": "原创性声明",
    "confidentiality": "保密审查",
}

# 只用来判断"稿件里提到过没有"。命中不代表格式合规，未命中也不代表一定缺失——
# 有些声明是投稿系统里单独勾选的，不在正文里。
STATEMENT_PATTERNS = {
    "copyright_transfer": r"版权|著作权|转让协议",
    "conflict_of_interest": r"利益冲突|作者声明|无利益",
    "data_availability": r"数据可得|数据可用|数据获取|数据和代码|数据与代码",
    "ai_declaration": r"AI\s*使用|人工智能使用|生成式|ChatGPT|大模型|AI\s*辅助",
    "originality": r"原创|未曾发表|未公开发表|一稿多投",
    "confidentiality": r"保密审查|涉密",
}

# 前置区在 IR 里统一是 front_matter，没有细分角色；
# 这里用机构名特征把"作者行"和"单位行"区分开，避免把署名判成缺失。
ORG = re.compile(r"(大学|学院|研究院|研究所|学校|中心|实验室|集团|公司|银行|医院|系|研究所)")

# 英文要件单独一套，不能与中文混判：只补了英文摘要不等于中文摘要存在
NEED_MAP_EN = {
    "title": {"english_title"},
    "authors": {"english_authors"},
    "affiliations": {"english_affiliations"},
    "abstract": {"english_abstract"},
    "keywords": {"english_keywords"},
    "body": {"body"},
    "references": {"references_heading", "reference_item"},
}


def count_words(blocks, scope="all"):
    """按 scope 统计三种口径。'body' 只数正文块，其余数全部有文字的块，
    表格块本身没有 text 键，表内文字在 parse_docx 存的 cells 里——声明
    「全文（含图表）」的刊必须把表格内文字算进去，否则标注与计算自相矛盾；
    取值来自规则集 limits.total.scope 经 SCOPE_FILTER 归一后的结果。"""
    if scope == "body":
        texts = [b.get("text", "") for b in blocks if b.get("type", "") == "body"]
    else:
        texts = []
        for b in blocks:
            if b.get("text"):
                texts.append(b["text"])
            elif b.get("cells"):
                texts.append("".join(b["cells"]))
    joined = "".join(texts)
    no_space = re.sub(r"\s", "", joined)
    cn_chars = len(re.findall(r"[\u4e00-\u9fff]", joined))
    en_words = len(re.findall(r"[A-Za-z][A-Za-z'-]*", joined))
    return {
        "A": len(no_space),
        "B": cn_chars + en_words,
        "C": len(joined),
        "cn": cn_chars,
        "en": en_words,
    }


def measure(text, metric):
    """按期刊声明的口径量一段文字。metric 取值见规则集 limits.*.metric。"""
    text = text or ""
    if metric == "chars_cn":
        return len(re.findall(r"[\u4e00-\u9fff]", text))
    if metric == "words_en":
        return len(re.findall(r"[A-Za-z][A-Za-z'-]*", text))
    if metric == "A":
        return len(re.sub(r"\s", "", text))
    if metric == "C":
        return len(text)
    return len(re.findall(r"[\u4e00-\u9fff]", text)) + len(re.findall(r"[A-Za-z][A-Za-z'-]*", text))


def detect_citation_style(ir):
    # 块可能整条不是对象（COMPNOND）：跳过，判定基于其余块照常给出
    refs = [b.get("text", "") for b in ir["blocks"]
            if isinstance(b, dict) and b.get("type", "") == "reference_item"]
    if not refs:
        return None, 0
    numbered = sum(1 for r in refs if re.match(r"^\[\d+\]", r))
    return ("numbered" if numbered > len(refs) / 2 else "author-year"), len(refs)


def _need(container, key, what):
    """取必需键；缺失时给出人话说明，而不是让 KeyError 冒成一屏 traceback。"""
    if not isinstance(container, dict) or container.get(key) is None:
        raise ValueError(f"{what}缺少必需键「{key}」——不是有效的数据结构，无法继续体检")
    return container[key]


def check(ir, rules):
    rows = []
    rows_raw = []
    ir_blocks = _need(ir, "blocks", "稿件 IR")
    rules_blocks = _need(rules, "blocks", "规则集")
    # 缺 type 键的块不再冒 KeyError（F26 坏文件隔离没盖到的同类残留）：
    # 按未知类型处理，并把缺在哪份数据的第几块记进 warnings 点名。
    warnings = []

    def _typed(blocks, what):
        out = []
        for i, b in enumerate(blocks):
            if not isinstance(b, dict):
                # 非 dict 元素（坏文件里出过字符串）同样不能崩：取值处的
                # b.get/b["type"] 会冒 AttributeError/TypeError——跳过并
                # 点名，判定基于打了折扣的输入照常给出（COMPNOND）。
                warnings.append(f"{what}第 {i + 1} 块不是对象"
                                f"（{type(b).__name__}），已跳过")
                continue
            if not b.get("type"):
                warnings.append(f"{what}第 {i + 1} 块缺 type 键，按未知类型处理")
                b = {**b, "type": ""}
            out.append(b)
        return out

    ir_blocks = _typed(ir_blocks, "稿件 IR ")
    rules_blocks = _typed(rules_blocks, "规则集 ")
    cite_cfg = _need(rules, "citation", "规则集")
    want = _need(cite_cfg, "system", "规则集 citation")
    present = {b.get("type", "") for b in ir_blocks}
    present |= {"front_matter"} if "front_matter" in present else set()

    fm = [b.get("text", "") for b in ir_blocks if b["type"] == "front_matter"]
    has_org = any(ORG.search(t) for t in fm)
    has_author = any(not ORG.search(t) and not t.strip().startswith(("——", "—", "-"))
                     for t in fm[1:])
    # 分类号常有中英两套，必须按语言分别判断
    cls_zh = any(b["type"] == "classification" and (b.get("lang") or "zh") != "en"
                 for b in ir_blocks)
    cls_en = any(b["type"] == "classification" and b.get("lang") == "en"
                 for b in ir_blocks)

    for blk in rules_blocks:
        t = blk.get("type", "")
        lang = blk.get("lang")
        if t == "classification_code":
            ok = cls_en if lang == "en" else cls_zh
        elif lang == "en":
            ok = bool(present & NEED_MAP_EN.get(t, {t}))
        elif t == "title":
            ok = "front_matter" in present
        elif t == "abstract":
            ok = "abstract" in present
        elif t == "keywords":
            ok = "keywords" in present
        elif t == "authors":
            ok = has_author
        elif t == "affiliations":
            ok = has_org
        else:
            ok = bool(present & NEED_MAP.get(t, {t}))
        name = LABEL.get(t, t) + (f"（{lang}）" if lang else "")
        rows.append((name, "有" if ok else "缺", blk.get("required", False)))
        # 原始行供 Web 端生成补录表单：网页需要知道类型与语言，而不只是中文标签
        rows_raw.append((t, lang, name, ok, blk.get("required", False)))

    limits = rules.get("limits", {})
    total = limits.get("total", {})
    metric = total.get("metric")
    # 总字数按规则集声明的 scope 过滤后再对上限：声明「正文」只数正文块，
    # 声明「全文」数全部块；写法未识别或未声明时保守按全文计，并在报告注明。
    # 标注的统计范围必须与实际参与统计的块一致，否则计数与口径自相矛盾。
    raw_scope = (total.get("scope") or "").strip()
    eff_scope = SCOPE_FILTER.get(raw_scope, "all")
    scope_note = ""
    if raw_scope not in SCOPE_FILTER:
        scope_note = (f"规则集声明为「{raw_scope}」，写法未识别，按全文口径统计"
                      if raw_scope else "规则集未声明统计范围，按全文口径统计")
    words = count_words(ir_blocks, scope=eff_scope)
    w_res = []
    for key, label in (("A", "字符数（不计空格）"), ("B", "中文字符 + 英文单词"), ("C", "字符数（计空格）")):
        val = words[key]
        mark = ""
        if metric == key and total.get("value"):
            over = val - total["value"]
            mark = f"｜上限 {total['value']:,}，{'超出 ' + format(over, ',') if over > 0 else '余量 ' + format(-over, ',')}"
        w_res.append((label, f"{val:,}{mark}"))

    total_info = None
    if metric and total.get("value"):
        actual = words.get(metric, 0)
        total_info = {"metric": metric, "value": total["value"], "actual": actual,
                      "over": actual - total["value"], "scope": raw_scope,
                      "scope_label": "正文" if eff_scope == "body" else "全文",
                      "scope_note": scope_note,
                      "note": total.get("note", ""), "strict": total.get("strict", False)}

    lim_abs = limits.get("abstract") or {}
    abstract_info = None
    abs_lo = lim_abs.get("min")
    abs_hi = lim_abs.get("max") or lim_abs.get("value")
    if abs_lo or abs_hi:
        abs_text = "".join(b.get("text", "") for b in ir_blocks if b["type"] == "abstract")
        actual = measure(abs_text, lim_abs.get("metric", "chars_cn"))
        abstract_info = {"lo": abs_lo, "hi": abs_hi, "value": abs_hi, "actual": actual,
                         "over": max(0, actual - abs_hi) if abs_hi else 0,
                         "under": max(0, abs_lo - actual) if abs_lo else 0,
                         "note": lim_abs.get("note", ""),
                         "strict": lim_abs.get("strict", False), "found": bool(abs_text)}

    # 关键词条数：须知写的是「3–5 个」，多半是硬性要求，但增减关键词属于改内容，
    # 只能报给作者，工具不替他删词也不替他补词
    lim_kw = limits.get("keywords") or {}
    keywords_info = None
    if lim_kw.get("min") or lim_kw.get("max"):
        kw_text = "".join(b.get("text", "") for b in ir_blocks if b["type"] == "keywords")
        body = kw_text.split("：", 1)[-1] if "：" in kw_text else kw_text
        count = len([x for x in re.split(r"[；;、,，\s\u3000]+", body) if x.strip()])
        keywords_info = {"min": lim_kw.get("min"), "max": lim_kw.get("max"),
                         "count": count, "found": bool(kw_text.strip())}

    # 标题字数：少数刊有明文限制（如国际金融研究 20 字）
    lim_title = limits.get("title_chars_max")
    title_info = None
    if lim_title:
        title_text = next((b.get("text", "") for b in ir_blocks
                           if b.get("type") == "front_matter"), "")
        n_title = measure(title_text, "chars_cn")
        title_info = {"max": lim_title, "actual": n_title,
                      "over": max(0, n_title - lim_title), "found": bool(title_text)}

    style, nrefs = detect_citation_style(ir)
    cite_ok = style == want if style else None

    # 匿名要求：多数中文经管期刊明确要求正文匿名（会计研究更是「否则直接退稿」）。
    # 这是信息层面的判断，不改稿件，只在报告里点名，措辞与「格式不合规」区分。
    sub = rules.get("submission") or {}
    anon_req = bool(sub.get("anonymous"))
    identity_lines = [t.strip() for t in fm[1:]
                      if t.strip() and not t.strip().startswith(("——", "—", "-"))]
    anon = {"required": anon_req, "lines": identity_lines[:3], "count": len(identity_lines)}

    # 必备声明：规则集 statements 逐条列出。只判断"稿件里提到过没有"，
    # 不判断是否合规——版权转让这类常在投稿系统里单独勾选，正文本来就不会有。
    st_conf = rules.get("statements") or {}
    all_text = "\n".join(b.get("text") or "" for b in ir_blocks)
    statements = []
    for key, conf in st_conf.items():
        conf = conf if isinstance(conf, dict) else {}
        pat = STATEMENT_PATTERNS.get(key)
        statements.append({
            "key": key,
            "label": STATEMENT_LABELS.get(key, key),
            "required": bool(conf.get("required")),
            "note": conf.get("note", ""),
            "found": bool(pat and re.search(pat, all_text)),
        })
    statements.sort(key=lambda s: (not s["required"], s["label"]))

    # 分节字数：帮作者定位"哪一节最占篇幅"，属定位信息，不是改写建议
    sections = []
    cur = None
    for b in ir_blocks:
        if b["type"] == "heading1":
            title = re.sub(r"^[一二三四五六七八九十]{1,3}\s*[、.．]\s*", "", b.get("text", "")).strip()
            if re.search(r"[\u4e00-\u9fff]", title):
                title = re.sub(r"\s+", "", title)
            cur = {"title": title or b.get("text", ""), "words": 0, "tables": 0, "figures": 0}
            sections.append(cur)
        elif cur is not None:
            if b["type"] == "body":
                s = b.get("text", "")
                cur["words"] += (len(re.findall(r"[\u4e00-\u9fff]", s))
                                 + len(re.findall(r"[A-Za-z][A-Za-z'-]*", s)))
            elif b["type"] == "table":
                cur["tables"] += 1
            elif b["type"] == "figure":
                cur["figures"] += 1

    return {
        "rows": rows,
        "rows_raw": rows_raw,
        # 数据本身可疑之处的点名（块缺 type 键这类，COMPSUB）——工具不崩，
        # 但读报告的人要知道判定是在打了折扣的输入上做出的
        "warnings": warnings,
        "words": w_res,
        # 本次统计实际用的口径与规则集原话：「统计范围」一格必须照实展示，
        # 没声明总字数上限时数字照常统计，标注也不能与计算相反
        "words_scope": eff_scope,
        "words_scope_declared": raw_scope,
        "words_scope_note": scope_note,
        "sections": sections,
        "total_limit": total_info,
        "abstract_limit": abstract_info,
        "keywords_limit": keywords_info,
        "title_limit": title_info,
        "citation": {"actual": style, "want": want, "ok": cite_ok, "n": nrefs},
        "anonymous": anon,
        "statements": statements,
        "tables": sum(1 for b in ir_blocks if b["type"] == "table"),
        "figures": sum(1 for b in ir_blocks if b["type"] == "figure"),
        "equations_text": sum(1 for b in ir_blocks
                              if b["type"] in ("equation_pseudo_table",) or
                              (b["type"] == "body" and re.match(r"^[a-zA-Z_]+\s*=\s*\S+（\d+）$", b.get("text", "")))),
        "equations_omml": sum(1 for b in ir_blocks if b["type"] == "equation"),
    }


def render(ir, rules, res):
    j = rules.get("journal") or {}
    ver = rules.get("verification") or {}
    src = ir.get("source") or "未命名稿件"
    rows_html = "".join(
        f'<tr><td>{html.escape(n)}</td>'
        f'<td class="{"ok" if v == "有" else "miss"}">{v}</td>'
        f'<td>{"必备" if req else "可选"}</td></tr>'
        for n, v, req in res["rows"]
    )
    words_html = "".join(f'<tr><td>{html.escape(k)}</td><td class="num">{html.escape(v)}</td></tr>'
                         for k, v in res["words"])

    c = res["citation"]
    cite_text = {
        (True, None): "稿件未检出参考文献条目",
    }.get((c["ok"], c["actual"]), "")
    if not cite_text:
        cite_text = ("一致，无需转换" if c["ok"] else
                     f"稿件为{'顺序编码制' if c['actual'] == 'numbered' else '作者年制'}，"
                     f"目标刊要求{'顺序编码制' if c['want'] == 'numbered' else '作者年制'}——需要重新格式化")
    # 文章编号、文献标识码由编辑部给定，作者补不了，不计入"还差几项"也不进待办
    missing = [n for t, lang, n, ok, req in res.get("rows_raw", [])
               if not ok and req and t not in NON_ACTIONABLE]
    publisher_missing = [n for t, lang, n, ok, req in res.get("rows_raw", [])
                         if not ok and req and t in NON_ACTIONABLE]
    tl, al = res.get("total_limit"), res.get("abstract_limit")

    def limit_cell(info, unit=""):
        """把「实测值 / 允许范围」渲染成一格，并区分硬性超限与建议值超出。

        规则集里字数上限有两种写法：单向的 `value`，和 `min`/`max` 区间
        （如统计研究「300–500 字为宜」）。两种都要能显示，否则写区间的刊
        在报告里等于没有这条限制。
        """
        if not info:
            return '<span class="muted">该刊未声明</span>'
        lo, hi = info.get("lo"), info.get("hi") or info.get("value")
        actual = info["actual"]
        target = f"{lo:,}–{hi:,}" if (lo and hi) else (f"{hi:,}" if hi else f"≥{lo:,}")
        if hi and actual > hi:
            word = "超出" if info.get("strict") else "高于建议值"
            cls = "bad" if info.get("strict") else "warn"
            tail = f'<span class="pill {cls}">{word} {actual - hi:,}{unit}</span>'
        elif lo and actual < lo:
            tail = f'<span class="pill warn">低于下限 {lo - actual:,}{unit}</span>'
        elif hi:
            tail = f'<span class="pill good">余量 {hi - actual:,}{unit}</span>'
        else:
            tail = '<span class="pill good">在范围内</span>'
        return f'{actual:,} / {target}{unit} {tail}'

    metric_name = {"A": "字符数（不计空格）", "B": "中文字符 + 英文单词",
                   "C": "字符数（计空格）"}.get((tl or {}).get("metric"), "待确认")

    # 「统计范围」一格必须如实写明本次统计用的范围：规则集声明了且认识，
    # 就原样照抄；没声明或写法未识别而回退全文的，必须注明回退——
    # 不能标签写「正文」、数字却是全文口径（F18 的自相矛盾就在这）。
    # 没声明总字数上限（tl 为空）时 A/B/C 数字照常按实际口径统计，标注
    # 同样跟随计算，不再写与计算相反的默认文案。
    scope_cell = '<span class="muted">规则集未声明统计范围，按全文口径统计</span>'
    if tl:
        if tl.get("scope"):
            scope_cell = html.escape(tl["scope"])
            if tl.get("scope_note"):
                scope_cell += f'<span class="muted">（{html.escape(tl["scope_note"])}）</span>'
        else:
            scope_cell = '<span class="muted">规则集未声明统计范围，按全文口径统计</span>'
    else:
        declared = res.get("words_scope_declared", "")
        if declared in SCOPE_FILTER:
            scope_cell = html.escape(declared)
        elif declared:
            scope_cell = (html.escape(declared)
                          + '<span class="muted">（写法未识别，按全文口径统计）</span>')

    kw = res.get("keywords_limit") or {}
    if not kw:
        kw_cell = '<span class="muted">该刊未声明关键词数量</span>'
    elif not kw.get("found"):
        kw_cell = '<span class="muted">稿件未检出关键词</span>'
    else:
        lo_kw, hi_kw = kw.get("min"), kw.get("max")
        rng = (f"{lo_kw}–{hi_kw} 个" if lo_kw and hi_kw
               else f"{hi_kw} 个以内" if hi_kw else f"不少于 {lo_kw} 个")
        bad = (hi_kw and kw["count"] > hi_kw) or (lo_kw and kw["count"] < lo_kw)
        kw_cell = (f'{kw["count"]} 个 / {rng} ' +
                   ('<span class="pill warn">不在范围内</span>' if bad
                    else '<span class="pill good">在范围内</span>'))

    ti = res.get("title_limit") or {}
    if not ti:
        title_cell = '<span class="muted">该刊未声明标题字数</span>'
    elif not ti.get("found"):
        title_cell = '<span class="muted">未检出标题</span>'
    else:
        title_cell = (f'{ti["actual"]} / {ti["max"]} 字 ' +
                      ('<span class="pill warn">超出</span>' if ti["over"]
                       else '<span class="pill good">在范围内</span>'))

    manual = []
    anon = res.get("anonymous") or {}
    if anon.get("required") and anon.get("count"):
        manual.append(
            f"该刊要求正文匿名：稿件前置区仍保留 {anon['count']} 处作者信息"
            f"（{'、'.join(anon['lines'])}），投稿前须删除或改用匿名稿——"
            f"这属于身份信息，不是格式错误，本工具不会代删")
    if res["equations_text"]:
        manual.append(f"检出 {res['equations_text']} 处公式以纯文本或伪表格形式存在，"
                      f"目标刊要求使用 Word 公式编辑器或 MathType，需作者自行重建")
    if missing:
        manual.append("缺失必备项：" + "、".join(missing) + "，需作者补充，工具不会代写")
    if publisher_missing:
        manual.append("、".join(publisher_missing) + " 由编辑部编号或按刊物体例给定，无须你处理")
    if tl and tl["over"] > 0:
        manual.append(f"总字数比该刊{'上限' if tl.get('strict') else '建议值'}多 {tl['over']:,}，"
                      f"如何缩减由作者决定——本工具不会为迁就字数删改任何内容")
    if al and al.get("found") and al["over"] > 0:
        manual.append(f"摘要 {al['actual']} 字，{'上限' if al.get('strict') else '建议值'} {al['value']} 字，"
                      f"超出 {al['over']} 字，需作者自行精简")
    if al and al.get("found") and al.get("under"):
        manual.append(f"摘要 {al['actual']} 字，低于该刊下限 {al['lo']} 字，需作者补充内容")
    kw = res.get("keywords_limit") or {}
    if kw.get("found"):
        if kw.get("max") and kw["count"] > kw["max"]:
            manual.append(f"关键词 {kw['count']} 个，超过该刊上限 {kw['max']} 个，删哪个由作者决定")
        elif kw.get("min") and kw["count"] < kw["min"]:
            manual.append(f"关键词 {kw['count']} 个，少于该刊下限 {kw['min']} 个，补什么由作者决定")
    ti = res.get("title_limit") or {}
    if ti.get("over"):
        manual.append(f"标题 {ti['actual']} 字，超过该刊 {ti['max']} 字的上限，如何精简由作者决定")

    st_rows = res.get("statements") or []
    for s in st_rows:
        if s["required"] and not s["found"]:
            manual.append(f"该刊要求{s['label']}，稿件中未检出相关文字——"
                          f"若该声明由投稿系统单独收集，可忽略此条")
    data_warns = res.get("warnings") or []
    warn_html = ""
    if data_warns:
        warn_html = ('<section class="warn"><h2>数据提示（判定基于有缺口的输入）</h2><ul>'
                     + "".join(f"<li>{html.escape(w)}</li>" for w in data_warns)
                     + "</ul></section>")
    st_html = ""
    if st_rows:
        st_html = (
            '<section><h2>该刊要求的声明</h2><table>'
            '<tr><td class="k">声明</td><td class="k">要求</td>'
            '<td class="k">稿件中</td><td class="k">说明</td></tr>'
            + "".join(
                f'<tr><td class="k">{html.escape(s["label"])}</td>'
                f'<td>{"必填" if s["required"] else "可选"}</td>'
                f'<td>{"已提到" if s["found"] else "<span class=muted>未检出</span>"}</td>'
                f'<td>{html.escape(s["note"] or "—")}</td></tr>'
                for s in st_rows)
            + '</table><p class="meta">「未检出」只表示正文里没有相关文字，'
              '不代表一定缺失——版权转让、伦理审查这类声明常由投稿系统单独收集，'
              '正文里本来就不会出现。</p></section>'
        )

    # ---- 判决句：作者只想知道"能不能投" ----
    n_missing = len(missing)
    over_total = tl["over"] if tl else 0
    over_abs = al["over"] if (al and al.get("found")) else 0
    cite_bad = c["ok"] is False

    bits = [f"还差 {n_missing} 项必备要件" if n_missing else "必备要件齐全"]
    if tl:
        scope_label = tl.get("scope_label", "正文")
        if over_total > 0:
            bits.append(f"{scope_label}超{'上限' if tl.get('strict') else '建议值'} {over_total:,} 字")
        else:
            bits.append(f"{scope_label}字数在范围内")
    if al and al.get("found"):
        bits.append(f"摘要超 {over_abs} 字" if over_abs > 0 else "摘要字数合规")
    if cite_bad:
        bits.append("引用体系需重排")
    elif c["ok"]:
        bits.append("引用体系一致")
    if anon.get("required") and anon.get("count"):
        bits.append("正文未匿名")
    verdict = "；".join(bits) + "。"

    if n_missing:
        nxt = f"先补齐这 {n_missing} 项必备要件，其余版式现在就可以排。"
    elif over_total > 0 or over_abs > 0:
        nxt = "格式上已可排版；字数怎么缩减由你决定，工具不会代删。"
    elif anon.get("required") and anon.get("count"):
        nxt = "版式已可排版；投稿前记得按该刊要求做匿名处理（本工具不改动身份信息）。"
    else:
        nxt = "格式与字数都不挡路，可以直接进入排版。"

    todos = [f"补齐{m}" for m in missing]
    if anon.get("required") and anon.get("count"):
        todos.append(f"删除前置区 {anon['count']} 处作者信息（该刊要求匿名）")
    if over_total > 0:
        todos.append(f"{(tl or {}).get('scope_label', '正文')}缩减 {over_total:,} 字（口径：{metric_name}）")
    if over_abs > 0:
        todos.append(f"摘要压缩 {over_abs} 字")
    if res["equations_text"]:
        todos.append(f"重建 {res['equations_text']} 处文本公式")
    if cite_bad:
        todos.append("按目标刊重排参考文献")
    todos_html = "".join(
        f'<label class="todo"><input type="checkbox" data-i="{i}"><span>{html.escape(t)}</span></label>'
        for i, t in enumerate(todos))

    sec_total = sum(s["words"] for s in res["sections"]) or 1
    sec_rows = "".join(
        f'<tr><td>{html.escape(s["title"][:22])}</td><td class="num">{s["words"]:,}</td>'
        f'<td class="num">{s["words"] / sec_total:.0%}</td>'
        f'<td class="num">{s["tables"] or "—"}</td></tr>'
        for s in res["sections"])

    kpi = f"""
<div class="kpi">
  <div class="{'hl-bad' if missing else 'hl-ok'}"><b>{len(missing)}</b><span>缺失必备项</span></div>
  <div><b>{limit_cell(tl) if tl else '—'}</b><span>{(tl or {}).get('scope_label', '正文')}口径字数 {('· ' + metric_name) if tl else ''}</span></div>
  <div><b>{limit_cell(al, ' 字') if al else '—'}</b><span>中文摘要字数</span></div>
  <div class="{'hl-bad' if c['ok'] is False else 'hl-ok'}"><b>{'需重排' if c['ok'] is False else '一致'}</b><span>引用体系</span></div>
</div>"""

    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>投稿体检 · {html.escape(src.replace('.docx', ''))} → {html.escape(j.get('name_zh') or j.get('name_en') or '未知期刊')}</title>
<style>{BASE}
td.ok{{color:var(--accent)}} td.miss{{color:var(--danger);font-weight:500}}
td.num{{font-variant-numeric:tabular-nums;white-space:nowrap}}
td.k{{color:var(--ink3);white-space:nowrap;width:150px}}
.kpi{{display:flex;margin:0 0 18px;background:var(--surface);border:1px solid var(--line);
border-radius:14px;overflow:hidden}}
.kpi>div{{flex:1;padding:14px 18px;border-left:1px solid var(--line2);min-width:0}}
.kpi>div:first-child{{border-left:none}}
.kpi b{{display:block;font-size:14px;font-weight:600;color:var(--ink);line-height:1.6;
font-variant-numeric:tabular-nums}}
.kpi span{{display:block;font-size:11.5px;color:var(--ink3);margin-top:2px}}
.kpi .hl-bad b{{color:var(--danger)}} .kpi .hl-ok b{{color:var(--accent)}}
.pill{{display:inline-block;font-size:11px;padding:1px 8px;border-radius:999px;
margin-left:5px;vertical-align:1px}}
.pill.bad{{background:var(--danger-bg);color:var(--danger)}}
.pill.warn{{background:var(--gold-bg);color:var(--gold)}}
.pill.good{{background:var(--accent-soft);color:var(--accent)}}
.muted{{color:var(--ink3)}}
.verdict{{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--accent);
border-radius:12px;padding:16px 20px;margin:0 0 14px}}
.verdict .v{{font-size:15px;font-weight:600;margin:0 0 6px;line-height:1.7}}
.verdict .vn{{font-size:12.5px;color:var(--ink2);margin:0;line-height:1.8}}
.todos-wrap{{background:var(--surface);border:1px solid var(--line);border-radius:12px;
padding:14px 20px;margin:0 0 18px}}
.todos-wrap h2{{margin:0 0 10px;border:none;padding:0;font-size:12px;color:var(--ink3);
font-weight:500;letter-spacing:.4px}}
.todos-wrap .empty{{font-size:12.5px;color:var(--accent);margin:0}}
.todo{{display:flex;align-items:flex-start;gap:9px;font-size:13px;padding:5px 0;cursor:pointer;
line-height:1.6}}
.todo input{{margin:3px 0 0;width:14px;height:14px;accent-color:var(--accent);flex:none}}
.todo input:checked+span{{text-decoration:line-through;color:var(--ink3)}}
ul{{margin:0;padding-left:19px;font-size:13.5px;color:var(--ink);line-height:1.85}}
.chips{{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 14px}}
.chip{{font-size:12px;padding:3px 10px;border-radius:999px;background:var(--soft);
border:1px solid var(--line);color:var(--ink2)}}
.chip.miss{{background:var(--danger-bg);border-color:#efcfc7;color:var(--danger)}}
@media(max-width:680px){{
  .kpi{{flex-wrap:wrap}}.kpi>div{{flex:1 1 48%;border-left:none;border-top:1px solid var(--line2)}}
  .kpi>div:nth-child(-n+2){{border-top:none}}
  td.k{{width:auto}}
}}
</style></head><body><div class="wrap">
<header>
  <h1>{html.escape(src.replace('.docx', ''))}</h1>
  <p class="sub">目标期刊：<strong>{html.escape(j.get('name_zh') or j.get('name_en') or '未知期刊')}</strong>
  · 依据 {html.escape(ver.get('source_title') or ver.get('source_url') or '（未记录出处）')}</p>
</header>
<div class="verdict">
  <p class="v">{verdict}</p>
  <p class="vn">{nxt}</p>
</div>
<div class="todos-wrap no-print">
  <h2>投稿前待办（可勾选，只存在你自己的浏览器里）</h2>
  {f'<div class="todos">{todos_html}</div>' if todos else '<p class="empty">没有需要作者处理的事项。</p>'}
</div>
{kpi}

<section><h2>结构对照</h2>
<div class="chips">{''.join(f'<span class="chip miss">{html.escape(n)} · 缺</span>' for n, v, req in res["rows"] if v == "缺") or '<span class="chip">结构齐全，该刊要求的组成部分均已具备</span>'}</div>
<table>
<tr><th>期刊要求的组成部分</th><th>稿件现状</th><th>要求级别</th></tr>
{rows_html}
</table></section>

<section><h2>字数（三种口径同时给出，仅统计不修改）</h2><table>
{words_html}
<tr><td class="k">本刊采用口径</td><td>{metric_name}{'（见规则集 limits.total.metric）' if tl else '（规则集未声明，待确认）'}</td></tr>
<tr><td class="k">统计范围</td><td>{scope_cell}</td></tr>
<tr><td class="k">摘要</td><td>{limit_cell(al, ' 字') if al else '<span class="muted">该刊未声明摘要字数</span>'}</td></tr>
<tr><td class="k">关键词</td><td>{kw_cell}</td></tr>
<tr><td class="k">标题字数</td><td>{title_cell}</td></tr>
</table>
{f'''<h2 style="margin-top:22px">分节字数（定位用，不是修改建议）</h2><table>
<thead><tr><th>章节</th><th>字数</th><th>占比</th><th>表</th></tr></thead>
{sec_rows}</table>''' if res["sections"] else ''}
</section>

<section><h2>引用体系</h2><table>
<tr><td class="k">稿件现状</td><td>{html.escape(str(c['n']))} 条参考文献，
{'顺序编码制' if c['actual'] == 'numbered' else '作者年制' if c['actual'] else '未检出'}</td></tr>
<tr><td class="k">目标期刊要求</td><td>{'顺序编码制' if c['want'] == 'numbered' else '作者年制'}</td></tr>
<tr><td class="k">结论</td><td>{html.escape(cite_text)}
<span class="tag {'good' if c['ok'] else 'bad'}">{'无需处理' if c['ok'] else '需重排'}</span></td></tr>
</table></section>

{st_html}

{warn_html}

<section><h2>图表</h2><table>
<tr><td class="k">表格</td><td>{res['tables']} 个（含伪表格形式的公式 {res['equations_text']} 处已剔除）</td></tr>
<tr><td class="k">插图</td><td>{res['figures']} 个</td></tr>
<tr><td class="k">公式对象</td><td>{res['equations_omml']} 个（OMML 原生公式）</td></tr>
</table></section>

<section class="warn"><h2>需要作者处理，工具不代做</h2>
<ul>{''.join(f'<li>{html.escape(m)}</li>' for m in manual) or '<li>无</li>'}</ul>
</section>

<footer>
  本报告只读取结构，未修改稿件任何字符<br>
  字数与结构结论仅作提示，是否调整由作者决定
</footer>
<script>
/* 待办勾选只写 localStorage，报告本身是静态文件，不会回传任何数据 */
(function(){{
  const key='triage-todo:'+document.title;
  const boxes=[...document.querySelectorAll('.todo input')];
  let saved={{}};
  try{{saved=JSON.parse(localStorage.getItem(key)||'{{}}')}}catch(e){{}}
  boxes.forEach(b=>{{
    b.checked=!!saved[b.dataset.i];
    b.addEventListener('change',()=>{{
      saved[b.dataset.i]=b.checked;
      try{{localStorage.setItem(key,JSON.stringify(saved))}}catch(e){{}}
    }});
  }});
}})();
</script>
</div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ir")
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    parsed = {}
    for label, path in (("稿件 IR", a.ir), ("规则集", a.ruleset)):
        try:
            parsed[label] = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:  # noqa: BLE001 —— 坏文件点名报错，不甩原始 traceback
            print(f"失败：文件 {path}（{label}）：{type(e).__name__}: {e}", file=sys.stderr)
            return 1
    ir, rules = parsed["稿件 IR"], parsed["规则集"]
    try:
        res = check(ir, rules)
        out_html = render(ir, rules, res)
    except ValueError as e:  # 缺键等数据问题：说明缺什么，而不是冒 traceback
        print(f"失败：{e}", file=sys.stderr)
        return 1
    Path(a.out).write_text(out_html, encoding="utf-8")
    print("已生成:", a.out)
    for w in res.get("warnings") or []:
        print("  提示:", w)
    print("  结构缺失:", [n for n, v, r in res["rows"] if v == "缺" and r] or "无")
    print("  引用体系:", res["citation"]["actual"], "→", res["citation"]["want"],
          "一致" if res["citation"]["ok"] else "需重排")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
