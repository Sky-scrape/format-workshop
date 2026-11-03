"""从期刊规则集生成人类可读的格式规格页（HTML）。

与 render_sample.py 同源：两者的输入是同一份 JSON，因此不会出现
"说明书写的" 与 "样板长的" 不一致。

用法：python tools/render_spec.py rules/journal/gjjryj.json -o rules/journal/gjjryj.spec.html
"""

import argparse
import base64
import html
import json
import sys
from pathlib import Path

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path，与
# render_audit.py 同一姿势显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

import assume_groups

FR_NAMES = {22: "二号", 16: "三号", 15: "小三", 14: "四号", 12: "小四", 10.5: "五号", 9: "小五"}


def pt_name(v):
    if v is None:
        return ""
    return f"{v} pt" + (f"（{FR_NAMES[v]}）" if v in FR_NAMES else "")


def row(k, v):
    return f'<tr><th>{html.escape(str(k))}</th><td>{html.escape(str(v))}</td></tr>'


def table(rows, title):
    body = "".join(rows)
    return f'<section><h2>{html.escape(title)}</h2><table>{body}</table></section>'


def embed_preview(pdf_path, dpi=100, quality=84):
    """把样板 PDF 的每一页转成内嵌图片。

    必须从样板文件转，不能用 HTML 另绘一份——那会立刻产生"页面看到的"与
    "下载到的 docx" 不一致的问题，比没有预览更糟。
    """
    try:
        import fitz
    except ImportError:
        return ""
    if not pdf_path or not Path(pdf_path).exists():
        return ""
    doc = fitz.open(str(pdf_path))
    figs = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=dpi)
        b64 = base64.b64encode(pix.tobytes("jpeg", jpg_quality=quality)).decode()
        figs.append(
            f'<figure><img src="data:image/jpeg;base64,{b64}" alt="样板第 {i} 页">'
            f'<figcaption>第 {i} 页</figcaption></figure>'
        )
    return (
        '<section><h2>成品样式预览</h2>'
        '<p class="hint">以下页面由本规则集自动渲染生成，与同目录的 '
        f'<code>{html.escape(Path(pdf_path).stem)}</code> 为同一份文件，非另行绘制。</p>'
        f'<div class="preview">{"".join(figs)}</div></section>'
    )


SORT_LABEL = {
    "cn-first-by-pinyin-then-en-by-alphabet": "中文在前按拼音，英文在后按字母",
    "cn-first-by-alphabet-then-en-by-alphabet": "中文在前按字母，英文在后按字母",
    "by_appearance_order": "按正文引用出现的先后顺序",
    "by-appearance": "按正文引用出现的先后顺序",
    "alphabetical": "按字母顺序",
}


def build(rules, sample_pdf=None):
    j = rules["journal"]
    v = rules["verification"]
    tp = rules["typography"]
    hd = rules["headings"]
    ci = rules["citation"]
    tb = rules["tables"]
    fig = rules["figures"]
    lim = rules["limits"]
    body = tp["body"]

    parts = []

    status = {"verified": ("已核验", "#176b57", "#e6f2ed"),
              "partial": ("部分核验", "#8a6626", "#FAEEDA"),
              "unverified": ("未核验", "#a44b3f", "#fbede9")}
    label, fg, bg = status.get(v["status"], ("未知", "#52655f", "#eef2ee"))
    auth = v.get("typography_authority", "published")
    auth_text = {"template": "官方投稿模板", "guide": "官方明文规定",
                 "published": "已发表论文实测"}.get(auth, "未声明")
    auth_color = {"template": "#176b57", "guide": "#0e4f41",
                  "published": "#8a6626"}.get(auth, "#52655f")
    src_url = str(v.get("source_url", "") or "")
    # html.escape 只转义 <>&'"，拦不住 javascript: 这类协议——规格页由应用
    # 同源提供，恶意 href 点击即在应用源内执行（F45）。仅 http(s) 渲染成
    # 链接，其余值降级为纯文本，出处本身仍如实展示。scheme 按不区分大小写
    # 判（浏览器就是这么解析的）："HTTPS://example.com" 是链接不是纯文本。
    src_html = (f'<a href="{html.escape(src_url)}">{html.escape(src_url)}</a>'
                if src_url.lower().startswith(("http://", "https://"))
                else html.escape(src_url))

    parts.append(f"""
<header>
  <div class="badge" style="color:{fg};background:{bg}">{label}</div>
  <h1>{html.escape(j['name_zh'])}</h1>
  <p class="sub">{html.escape(j.get('name_en',''))} · {html.escape(j.get('sponsor',''))}</p>
  <p class="meta">核验于 {html.escape(str(v.get('verified_at','')))} · 依据 {html.escape(v.get('verified_by',''))}<br>
  版式依据：<strong style="color:{auth_color}">{auth_text}</strong><br>
  来源：{src_html}</p>
</header>""")

    # 交叉核对结论
    cc = v.get("cross_check")
    if cc:
        li = lambda xs: "".join(f"<li>{html.escape(x)}</li>" for x in xs)
        blocks = []
        if cc.get("confirmed"):
            blocks.append(f'<div class="cc good"><h3>与实物吻合</h3><ul>{li(cc["confirmed"])}</ul></div>')
        if cc.get("discrepancies") or cc.get("corrected"):
            items = (cc.get("discrepancies") or []) + (cc.get("corrected") or [])
            # 差异之后取了哪一边，由规则集的 typography_authority 决定，不在代码里认刊名
            kept_official = (cc.get("resolution") == "official"
                             or (cc.get("resolution") is None
                                 and v.get("typography_authority") in ("guide", "template")))
            tail = "保留官方值" if kept_official else "已按实测修正"
            blocks.append(f'<div class="cc warn2"><h3>指南与实物不符（{tail}）</h3><ul>{li(items)}</ul></div>')
        if cc.get("observed_difference"):
            blocks.append(f'<div class="cc info"><h3>与实物有差异，但保留官方值</h3><ul>{li(cc["observed_difference"])}</ul></div>')
        if cc.get("conclusion"):
            blocks.append(f'<p class="hint" style="margin:12px 0 0">{html.escape(cc["conclusion"])}</p>')
        parts.append('<section><h2>交叉核对（{0} 篇样本）</h2>{1}</section>'.format(
            cc.get("samples", "—"), "".join(blocks)))

    conflict = v.get("conflict")
    if conflict:
        parts.append(
            '<section class="warn"><h2>官方与实物冲突</h2><table>'
            + row("官方要求", conflict.get("official"))
            + row("实物实测", conflict.get("published"))
            + row("本稿取", conflict.get("resolution"))
            + row("风险", conflict.get("risk"))
            + '</table></section>')

    preview = embed_preview(sample_pdf)
    if preview:
        parts.append(preview)

    # 结构顺序
    order = []
    for b in rules["blocks"]:
        t = b["type"]
        lang = b.get("lang", "")
        name = {
            "title": "标题", "subtitle": "副标题", "abstract": "摘要", "keywords": "关键词",
            "classification_code": "分类号", "document_code": "文献标识码", "body": "正文",
            "references": "参考文献", "appendix": "附录", "acknowledgement": "致谢",
        }.get(t, t)
        order.append(f"{name}({lang})" if lang else name)
    parts.append(table(
        [row("稿件顺序", " → ".join(order)),
         row("作者信息页", "独立成页置于正文前" if rules["submission"]["author_page"]["separate_page"] else "无"),
         row("正文匿名", "是" if rules["submission"]["anonymous"] else "否")],
        "整体结构"))

    # 页面
    pm = rules["page"]["margin_mm"]
    parts.append(table([
        row("纸张", rules["page"]["size"]),
        row("页边距", f"上 {pm['top']} / 下 {pm['bottom']} / 左 {pm['left']} / 右 {pm['right']} mm"),
        row("分栏", f"{rules['page']['columns']} 栏"),
    ], "页面设置"))

    # 字体字号
    ls = body["line_spacing"]
    ls_text = f"{ls.get('value_pt')} pt 固定值" if ls["mode"] == "exact" else f"{ls.get('value')} 倍"
    title_row = row("文章标题", f"{tp['title'].get('font_cn', '')} {pt_name(tp['title'].get('size_pt'))}，"
                              + ("居中" if tp["title"].get("align") == "center" else "左对齐")
                              + ("（推定）" if tp["title"].get("assumed") else ""))
    subtitle_row = ""
    if tp.get("subtitle"):
        subtitle_row = row("副标题", f"{tp['subtitle'].get('font_cn', '')} "
                                   f"{pt_name(tp['subtitle'].get('size_pt'))}，居中")
    lv1 = hd["level1"]
    parts.append(table([
        row("正文", f"{body['font_cn']} / {body.get('font_latin', '')} {pt_name(body['size_pt'])}"),
        row("行距", ls_text),
        row("首行缩进", f"{body['first_line_indent_chars']} 字符"),
        row("对齐", "两端对齐" if body.get("align") == "both" else body.get("align", "")),
        title_row,
        subtitle_row,
        row("一级标题", f"{lv1.get('font_cn', '')} {pt_name(lv1.get('size_pt'))}，"
                       f"{'加粗，' if lv1.get('bold') else '不加粗，'}"
                       f"{'居中' if lv1.get('align') == 'center' else '左对齐'}"),
        row("二 / 三级标题", "字体字号随正文，左对齐，不加粗"),
        row("标题编号", " / ".join(hd["levels"]) +
            ("（{0}不编号）".format("、".join(hd.get("unnumbered_sections", [])))
             if hd.get("unnumbered_sections") else "")),
        row("最多层级", f"{hd['max_level']} 级"),
        row("图表标题", f"{tb['caption_font_cn']} {pt_name(tb['caption_size_pt'])}"
                       + ("（推定）" if tb.get("caption_assumed") else "")),
        row("脚注", f"{tp['footnote']['font_cn']} {pt_name(tp['footnote']['size_pt'])}"),
    ], "字体与行距"))

    # 表格
    tl = "、".join(f"{x['kind']}={x.get('prefix','') or x.get('marker','')}（{'加标点' if x['trailing_punct'] else '不加标点'}）"
                  for x in tb["note_layers"])
    bp = tb.get("border_pt") or {}
    parts.append(table([
        row("样式", "三线表" + ("，两侧不加墙线" if not tb.get("side_walls") else "")),
        row("线宽", f"顶线 {bp.get('top','—')} pt / 栏目线 {bp.get('header_bottom','—')} pt / 底线 {bp.get('bottom','—')} pt"),
        row("表题位置", "表格上方，居中"),
        row("表号与表题", tb.get("caption_gap", "")),
        row("表注层次", tl or "无"),
        row("表注字号", f"{tb['note_layers'][0].get('size_pt','')} pt"
            if tb.get("note_layers") else "该刊表下不加注释"),
        row("跨页表头", "重复" if tb.get("repeat_header_on_break") else "不重复"),
    ], "表格规范"))

    # 图
    parts.append(table([
        row("边框", "无边框" if not fig["border"] else "有边框"),
        row("图题位置", "插图下方，居中"),
        row("图注层次", "、".join(x["kind"] for x in fig["note_layers"]) or "无"),
        row("图注字号", f"{fig['note_layers'][0].get('font_cn','')} {fig['note_layers'][0].get('size_pt','')} pt"
            if fig.get("note_layers") else "该刊图下不加注释"),
        row("图形格式", "优先矢量图" if fig.get("prefer_vector") else ""),
    ], "插图规范"))

    # 引用
    rl = ci["reference_list"]
    it = ci["in_text"]
    if ci["system"] == "numbered":
        # 编号制的正文引用是上标序号，不是作者—年份；规则集给了说明就以它为准
        in_text = (it.get("note") or it.get("format")
                   or "正文用上标方括号序号 [n]，与文末条目一一对应（实引制）")
        sort_text = SORT_LABEL.get(rl["sort"], rl["sort"])
        num_text = rl.get("number_format", "[序号]") + (
            f"，序号后 {rl['space_after_number']} 个空格" if "space_after_number" in rl else "")
    else:
        ex = (it.get("examples") or ["（Lee et al.，1993）"])[0]
        in_text = it.get("note") or (
            f"作者年制，括注或叙述式，如 {ex}；两位作者中文用「{it.get('two_authors_conjunction_cn', '和')}」、"
            f"英文用「{it.get('two_authors_conjunction_en', 'and')}」；"
            f"三位及以上：首位作者 + 「{it.get('three_plus_ellipsis_cn', '等')}」/「"
            f"{it.get('three_plus_ellipsis_en', 'et al.')}」")
        sort_text = SORT_LABEL.get(rl["sort"], rl["sort"])
        hin = rl.get("hanging_indent_note")
        num_text = "不编号" + (f"；{hin}" if hin else "")

    # 正文引用体系与文献表是否编号是两件事：统计研究正文用作者—年份，
    # 文献表却是 [1] 编号列表（编号只是序号，不对应正文）
    if rl.get("numbered") and ci["system"] != "numbered":
        num_text = (rl.get("number_format", "[序号]") +
                    ("（仅作列表序号，正文不标注）" if rl.get("number_role") == "list_index_only" else ""))
    parts.append(table([
        row("体系", "顺序编码制" if ci["system"] == "numbered" else "作者年制"),
        row("依据标准", ci["style_id"] + ("（期刊自定体例）" if not ci["style_id"].startswith(("GB", "APA", "MLA", "Chicago")) else "")),
        row("正文引用", in_text),
        row("文献表排序", sort_text),
        row("版式", num_text),
        row("标点规则", rl.get("punctuation_rule")
            or (f"中文用{rl['cn_punctuation']}，英文用{rl['en_punctuation']}"
                if rl.get("cn_punctuation") and rl.get("en_punctuation")
                else "规则集未单列，按上方文献示例执行")),
        row("英文姓名", rl.get("english_author_name_rule", "姓在前名在后，姓不缩写，名缩写")),
    ], "参考文献"))

    # 限制
    total = lim.get("total") or {}
    if total.get("value"):
        lim_rows = [row("正文", f"{total['value']} 字（{total['metric']} 口径，" +
                              ("硬性" if total.get("strict") else "建议值") + "）")]
    else:
        # 少数期刊只规定必备构件，不设总字数上限
        lim_rows = [row("正文", "该刊未规定总字数上限")]
    for key, label in (("abstract", "中文摘要"), ("abstract_en", "英文摘要")):
        v = lim.get(key)
        if not v:
            continue
        unit = "词" if v.get("metric") == "words_en" else ("字符" if v.get("metric") == "A" else "字")
        # 有的期刊只给区间（如英文长摘要 3600～4100 字，占一个页码），没有单一上限
        if v.get("value"):
            txt = f"{v['value']} {unit}（" + ("硬性" if v.get("strict") else "建议值") + "）"
        elif v.get("min") and v.get("max"):
            txt = f"{v['min']}～{v['max']} {unit}"
        elif v.get("max"):
            txt = f"不超过 {v['max']} {unit}"
        elif v.get("min"):
            txt = f"不少于 {v['min']} {unit}"
        else:
            txt = "官方未给字数上下限"
        lim_rows.append(row(label, txt))
    if lim.get("keywords"):
        kw = lim["keywords"]
        span = f"{kw['min']}–{kw['max']}" if kw.get("min") else f"不超过 {kw['max']}"
        lim_rows.append(row("关键词", f"{span} 个"))
    if lim.get("classification_code_count"):
        lim_rows.append(row("分类号", f"{lim['classification_code_count']} 个（{rules['blocks'][2].get('scheme','') if len(rules['blocks'])>2 else ''}）"))
    if lim.get("title_chars_max"):
        lim_rows.append(row("标题长度", f"不超过 {lim['title_chars_max']} 字"))
    lim_rows.append(row("处理方式", "仅提醒，不修改稿件内容"))
    parts.append(table(lim_rows, "字数与限制"))

    # 个人信息
    pf_rows = []
    for f in rules["personal_fields"]:
        flag = "必填" if f.get("required") else "选填"
        sens = "，敏感" if f.get("sensitive") else ""
        pf_rows.append(row(f["label_zh"], f"{flag}{sens} · 出现在{f['placement']}"))
    parts.append(table(pf_rows, f"需收集的个人信息（共 {len(rules['personal_fields'])} 项）"))

    if rules.get("assumptions"):
        groups = assume_groups.group_assumptions(
            rules["assumptions"], rules.get("verification", {}).get("extraction"))
        n = sum(len(rows) for _, rows in groups)
        if assume_groups.should_group(groups):
            # 大批量（AI 建库草稿发布后）：按区块分组折叠——披露一条不删，
            # 但不能把复核重点淹没。逐条仍可达，来源区块随组名亮出。
            blocks = []
            for sec, rows in groups:
                lis = "".join(f"<li>{html.escape(a)}</li>" for a in rows)
                blocks.append(
                    f'<details><summary>{html.escape(sec)} '
                    f'<span class="abadge">{len(rows)} 条</span>'
                    f'（官方未规定，沿用基线/推定）</summary>'
                    f'<ul style="margin:4px 0 8px;padding-left:20px">{lis}</ul></details>')
            parts.append('<section class="warn"><h2>推定项（官方未规定，非期刊要求）</h2>'
                         f'<p class="hint">共 {n} 条，按区块折叠，逐条可达；'
                         '其中多数沿用自基线规则集，发布前须人工核对。</p>'
                         + "".join(blocks) + '</section>')
        else:
            items = "".join(f"<li>{html.escape(a)}</li>" for a in rules["assumptions"])
            parts.append('<section class="warn"><h2>推定项（官方未规定，非期刊要求）</h2>'
                         f'<ul style="margin:0;padding-left:20px;font-size:13px;color:#7d3a30">{items}</ul></section>')

    if tb.get("open_issue"):
        parts.append(f'<section class="warn"><h2>待确认</h2><p>{html.escape(tb["open_issue"])}</p></section>')

    body_html = "\n".join(parts)
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(j['name_zh'])} · 格式规格</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;padding:40px 24px 64px;background:#f4f6f3;color:#18332e;
font-family:Inter,"PingFang SC","Microsoft YaHei",system-ui,-apple-system,sans-serif;line-height:1.65}}
.wrap{{max-width:820px;margin:0 auto}}
header{{margin-bottom:32px;padding-bottom:24px;border-bottom:1px solid #dfe6e1}}
.badge{{display:inline-block;font-size:12px;padding:3px 10px;border-radius:20px;margin-bottom:12px}}
h1{{font-family:"Songti SC","Noto Serif SC",SimSun,serif;font-size:26px;font-weight:600;margin:0 0 6px;letter-spacing:.8px}}
.sub{{margin:0 0 10px;color:#52655f;font-size:14px}}
.meta{{margin:0;color:#64756e;font-size:12px;word-break:break-all}}
.meta a{{color:#0e4f41;text-decoration:none}}
section{{margin:0 0 26px;background:#fff;border:1px solid #dfe6e1;border-radius:12px;padding:18px 22px}}
section.warn{{background:#fbede9;border-color:#efcfc7}}
section h2{{font-size:14px;font-weight:600;margin:0 0 12px;color:#18332e;
padding-left:10px;border-left:3px solid #176b57}}
section.warn h2{{border-left-color:#a44b3f}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{text-align:left;font-weight:400;color:#52655f;width:118px;padding:7px 10px 7px 0;
vertical-align:top;white-space:nowrap}}
td{{padding:7px 0;color:#18332e;word-break:break-word}}
tr+tr th,tr+tr td{{border-top:1px solid #eef2ee}}
footer{{margin-top:36px;color:#64756e;font-size:12px;text-align:center}}
.hint{{margin:0 0 14px;font-size:12px;color:#64756e}}
.hint code{{font-family:ui-monospace,Consolas,monospace;font-size:11px;
background:#eef2ee;padding:1px 5px;border-radius:3px}}
.preview{{background:#eef2ee;border-radius:10px;padding:16px}}
.preview figure{{margin:0 0 14px;text-align:center}}
.preview figure:last-child{{margin-bottom:0}}
.preview img{{width:100%;max-width:700px;background:#fff;border:1px solid #dfe6e1;
border-radius:2px;display:block;margin:0 auto}}
.preview figcaption{{font-size:12px;color:#64756e;margin-top:6px}}
.cc{{border-left:3px solid #B4B2A9;padding:2px 0 2px 12px;margin:0 0 12px}}
.cc:last-child{{margin-bottom:0}}
.cc h3{{font-size:12px;font-weight:600;margin:0 0 6px;color:#52655f}}
.cc ul{{margin:0;padding-left:18px;font-size:13px;line-height:1.7}}
.cc.good{{border-left-color:#176b57}} .cc.good h3{{color:#176b57}}
.cc.warn2{{border-left-color:#a44b3f}} .cc.warn2 h3{{color:#a44b3f}}
.cc.info{{border-left-color:#176b57}} .cc.info h3{{color:#0e4f41}}
details{{margin:2px 0}}
details summary{{cursor:pointer;font-size:13px;color:#7d3a30}}
.abadge{{display:inline-block;min-width:14px;text-align:center;font-size:11px;
font-weight:600;color:#a44b3f;background:#fff;border:1px solid #efcfc7;
border-radius:999px;padding:0 7px;margin:0 2px}}
</style></head>
<body><div class="wrap">
{body_html}
<footer>本页由规则集 {html.escape(j['id'])}.json 自动生成 · 与版式样板同源，二者不会不一致</footer>
</div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--sample-pdf", default=None, help="样板 PDF，用于在规格页内嵌成品预览")
    args = ap.parse_args()
    with open(args.ruleset, encoding="utf-8") as f:
        rules = json.load(f)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(build(rules, sample_pdf=args.sample_pdf))
    print("已生成:", args.out)


if __name__ == "__main__":
    main()
