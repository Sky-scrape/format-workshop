"""从稿约 PDF 抽期刊规则集草稿。

分工（对应 DESIGN §14.7 与 §15）
--------------------------------------------------------------------
**结构由代码保证，AI 只负责填值。**

规则集有十几个区块、上百个字段，让模型直接吐完整 JSON，它必然漏字段、
改结构、编造键名，产物连渲染器都跑不起来。所以这里反过来做：

1. 先复制一份已知良好的基线规则集，结构天然完整
2. AI 只回答一张扁平的事实清单（约 45 个键，值 + 一句原文引文）
3. 引文必须在原文里逐字搜得到，搜不到整条丢弃——这是防编造的主闸
4. 没抽到的键保留基线值，但逐条记进 `assumptions` 并在复核页标红

产物永远能渲染，但**永远标 `unverified`**，必须人工核对后才允许发布。

用法
----
    python tools/extract_journal.py 稿约.pdf --id jryj --name 金融研究 -o /tmp/draft.json
"""

import argparse
import json
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
- body_align / heading_level1_align: both（两端对齐）/ left / center
- heading_numbering: cn-hierarchical（一、（一）1.）/ arabic（1 / 1.1 / 1.1.1）/ none
- citation_system: numbered（顺序编码制，正文标 [1]）/ author-year（作者年制，正文标（张三，2020））
- limit_total_metric: A（字符数不计空格）/ B（中文汉字数）/ C（字符数计空格）
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


# ---------------------------------------------------------------- PDF

def pdf_text(path, max_chars=MAX_PDF_CHARS, log=None):
    """取 PDF 文字：有文字层用文字层，没有才 OCR（见 tools/ocr.py）。

    扫描件不再拒收——这是稿约场景的常态，纸质须知扫描成 PDF 很常见。
    """
    return ocr.extract_pdf_text(path, max_chars=max_chars, log=log)


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
            return float(str(value).replace("pt", "").strip()), ""
        if kind == "int":
            return int(float(str(value).strip())), ""
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
    except (TypeError, ValueError):
        return None, "类型不符"
    return None, "未知类型"


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
    """让模型抽事实。返回 (facts, evidence, notes, raw)。"""
    keys = "\n".join(f"- {k}: {meta[2]}" for k, meta in FACTS.items())
    keys += "\n- body_line_mode: exact（固定行距，给磅值）或 multiple（倍数行距）\n"
    keys += "- body_line_value: 行距数值（固定值给磅，倍数给倍数）\n"
    keys += "- heading_levels: 各级标题的编号写法数组，如 [\"一、\",\"（一）\",\"1.\"]"
    prompt = EXTRACT_PROMPT.format(keys=keys, enums=ENUM_HINT)

    messages = [
        {"role": "system", "content": prompt},
        {"role": "user", "content": "以下是稿约正文：\n\n" + text},
    ]
    content = ai_assist._chat(cfg, messages)
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


def baseline_rules():
    """结构基线：取一本已核验的规则集，任何字段缺失都由它兜住。"""
    for name in ("gjjryj", "jryj", "cjyj"):
        p = RULES_DIR / f"{name}.json"
        if p.exists():
            return deepcopy(json.loads(p.read_text(encoding="utf-8"))), name
    cands = sorted(RULES_DIR.glob("*.json"))
    if not cands:
        raise RuntimeError("rules/journal 下没有任何规则集可供作为结构基线")
    return deepcopy(json.loads(cands[0].read_text(encoding="utf-8"))), cands[0].stem


def build(extracted, evidence, text, *, jid, name_zh="", source="", meta=None):
    """把事实清单映射成完整规则集。返回 (rules, report)。"""
    base, base_id = baseline_rules()
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
            dropped.append({"key": key, "label": label, "value": raw, "why": why})
            continue
        quote = str(evidence.get(key) or "").strip()
        if not quote:
            unverified.append({"key": key, "label": label, "value": val,
                               "why": "模型未给出原文引文，无法核验"})
        elif norm(quote) not in norm(text):
            # 引文在原文里搜不到 —— 判为编造，整条丢弃
            dropped.append({"key": key, "label": label, "value": raw,
                            "why": f"引文在原文中找不到：{quote[:40]}"})
            continue
        if _set_path(rules, path, val):
            used[key] = {"label": label, "value": val, "evidence": quote}

    # 成组项
    mode = str(extracted.get("body_line_mode") or "").strip()
    value = extracted.get("body_line_value")
    if mode in ("exact", "multiple") and value not in (None, ""):
        try:
            v = float(value)
            size = rules["typography"]["body"].get("size_pt") or 10.5
            ls = {"mode": mode}
            if mode == "exact":
                ls["value_pt"] = v
            else:
                ls["value"] = v
                ls["equivalent_exact_pt"] = round(size * v, 1)
            rules["typography"]["body"]["line_spacing"] = ls
            used["body_line"] = {"label": "行距", "value": f"{mode} {v}",
                                 "evidence": str(evidence.get("body_line_value") or "")}
        except (TypeError, ValueError):
            dropped.append({"key": "body_line_value", "label": "行距",
                            "value": value, "why": "不是数字"})

    levels = extracted.get("heading_levels")
    if isinstance(levels, list) and levels and all(isinstance(x, str) and x.strip() for x in levels):
        rules["headings"]["levels"] = [x.strip() for x in levels]
        rules["headings"]["max_level"] = min(len(levels), 4)
        used["heading_levels"] = {"label": "标题编号", "value": rules["headings"]["levels"],
                                  "evidence": str(evidence.get("heading_levels") or "")}

    # 刊物标识：以用户在页面上填的为准，AI 抽到的只作补充
    rules["journal"]["id"] = jid
    if name_zh:
        rules["journal"]["name_zh"] = name_zh
    rules.setdefault("journal", {})["name_zh"] = rules["journal"].get("name_zh") or name_zh or jid

    # 用户亲手填的字段不算"沿用基线"，否则复核页会提示一条假风险
    handwritten = {"journal_name_zh"} if name_zh else set()
    inherited = [{"key": k, "label": meta[2],
                  "value": _get_path(rules, meta[0])}
                 for k, meta in FACTS.items()
                 if k not in used and k not in handwritten]

    rules["verification"] = {
        "status": "unverified",
        # 渲染器直接下标取 source_url，缺键会抛 KeyError；草稿阶段无官方链接，留空串
        "source_url": "",
        "template_url": "",
        "template_date": "",
        "verified_at": "",
        "verified_by": "AI 从稿约 PDF 抽取，未经人工核对",
        "notes": (f"抽取自 {source}；{len(used)} 项有原文依据，"
                  f"{len(inherited)} 项沿用基线值，{len(dropped)} 项被丢弃"),
        "typography_authority": "guide",
        "extraction": {
            "source_file": source,
            "extracted_at": date.today().isoformat(),
            "baseline": base_id,
            "pages": (meta or {}).get("pages"),
            "truncated": (meta or {}).get("truncated", False),
            # 文字从哪来的必须留痕：OCR 一定有错字，复核时要区别对待
            "text_method": (meta or {}).get("method", "text_layer"),
            "ocr_pages": (meta or {}).get("ocr_pages") or [],
            "ocr_engine": (meta or {}).get("engine", ""),
            "ai_model": None,
            "used": {k: v for k, v in used.items()},
            "inherited": inherited,
            "dropped": dropped,
            "unverified": unverified,
        },
    }
    # 渲染器把 assumptions 逐条当字符串转义，这里只能是字符串数组
    rules["assumptions"] = [
        f"{it['label']} = {it['value']}（未从稿约抽到，沿用基线值，须人工核对）"
        for it in inherited if it["value"] not in (None, "", [], {})
    ]

    report = validate(rules, used=used, inherited=inherited, dropped=dropped)
    report["stats"] = {"extracted": len(used), "inherited": len(inherited),
                       "dropped": len(dropped), "unverified": len(unverified)}
    return rules, report


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


def extract(pdf, *, jid, name_zh="", cfg=None, log=None):
    """完整流程：读 PDF → 抽事实 → 建规则集 → 校验。"""
    cfg = cfg or ai_assist.load_config()
    if not ai_assist.available(cfg):
        raise RuntimeError("未配置可用的 AI 服务，无法抽取规则集")
    text, meta = pdf_text(pdf, log=log)
    facts, ev, notes, _raw = ai_facts(text, cfg, log=log)
    rules, report = build(facts, ev, text, jid=jid, name_zh=name_zh,
                          source=Path(pdf).name, meta=meta)
    rules["verification"]["extraction"]["ai_model"] = cfg.get("model")
    if notes:
        rules["verification"]["extraction"]["ai_notes"] = notes[:400]
    return rules, report, text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--id", required=True, help="期刊代号，小写字母，如 jryj")
    ap.add_argument("--name", default="", help="中文刊名")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    try:
        rules, report, _ = extract(a.pdf, jid=a.id, name_zh=a.name, log=print)
    except Exception as e:  # noqa: BLE001
        print(f"失败：{e}")
        return 1

    Path(a.out).write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    s = report["stats"]
    print(f"抽取 {s['extracted']} 项 / 沿用基线 {s['inherited']} 项 / 丢弃 {s['dropped']} 项")
    for e in report["errors"]:
        print("  错误:", e["msg"])
    for w in report["warnings"][:10]:
        print("  提示:", w["msg"])
    print("可发布" if report["ok_to_publish"] else "不可发布：先修掉错误项")
    return 0 if report["ok_to_publish"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
