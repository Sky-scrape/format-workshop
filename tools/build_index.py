"""把 23 本期刊的规则集汇成一页总览，每本带一张首页缩略图和规格页入口。

为什么要有这一页：规则集是 JSON、规格页是本刊专用的一整页，两者都不方便横向比较。
加总览页之后，"这 23 本到底长什么样、依据是什么、缺什么"一屏能看完。

缩略图从 {id}.sample.pdf 的第一页转，不用 HTML 另画——另画会与下载到的 docx 不一致。

用法：python tools/build_index.py [-o rules/index.html]
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JDIR = ROOT / "rules" / "journal"

AUTH_LABEL = {
    "template": ("官方模板", "#0e4f41", "#dceee6"),
    "guide": ("官方明文", "#0e4f41", "#e4f0ea"),
    "published": ("实物实测", "#8a5a14", "#f6ecd8"),
    None: ("未标注", "#7a4a3f", "#f7e7e2"),
}


def thumb(pdf: Path, dpi=72, quality=72):
    if not pdf.exists():
        return None
    try:
        import fitz
    except ImportError:
        return None
    d = fitz.open(pdf)
    if not d.page_count:
        return None
    pix = d[0].get_pixmap(dpi=dpi)
    return base64.b64encode(pix.tobytes("jpeg", jpg_quality=quality)).decode()


def line_spacing_text(ls):
    if not isinstance(ls, dict):
        return "—"
    if ls.get("mode") == "exact":
        return f"固定 {ls.get('value_pt')} 磅"
    return f"{ls.get('value')} 倍"


def summarize(r):
    j, t, p = r["journal"], r.get("typography", {}), r.get("page", {})
    body = t.get("body", {})
    lim = r.get("limits", {})
    total = lim.get("total") or {}
    cite = r.get("citation", {})
    auth = (r.get("verification") or {}).get("typography_authority")
    label, fg, bg = AUTH_LABEL.get(auth, AUTH_LABEL[None])
    return {
        "id": j["id"],
        "name": j.get("name_zh") or j.get("name_en"),
        "name_en": j.get("name_en"),
        "lang": j.get("language"),
        "status": (r.get("verification") or {}).get("status"),
        "auth": auth,
        "auth_label": label, "auth_fg": fg, "auth_bg": bg,
        "page": f"{p.get('size', '—')}　{'单栏' if p.get('columns') == 1 else str(p.get('columns')) + ' 栏'}",
        "body": f"{body.get('font_cn') or body.get('font_latin') or '—'} {body.get('size_pt', '—')}pt",
        "spacing": line_spacing_text(body.get("line_spacing")),
        "headings": (r.get("headings", {}).get("levels") or [])[:3],
        "citation": cite.get("style_id") or cite.get("system") or "—",
        "total": f"{total.get('value')} {total.get('metric') or ''}" if total.get("value") else "未规定",
        "spec": f"{j['id']}.spec.html",
        "docx": f"{j['id']}.sample.docx",
        "official": (ROOT / "rules" / "sources" / j["id"]).is_dir(),
        "conflict": bool((r.get("verification") or {}).get("conflict")),
    }


CARD = """
<article class="card">
  <div class="thumb">{thumb}</div>
  <div class="body">
    <div class="row1"><h3>{name}</h3><span class="chip" style="color:{auth_fg};background:{auth_bg}">{auth_label}</span></div>
    <p class="en">{name_en}</p>
    <dl>
      <dt>页面</dt><dd>{page}</dd>
      <dt>正文</dt><dd>{body}　{spacing}</dd>
      <dt>标题</dt><dd>{headings}</dd>
      <dt>引用</dt><dd>{citation}</dd>
      <dt>字数</dt><dd>{total}</dd>
      <dt>依据</dt><dd>{basis}</dd>
    </dl>
    <p class="links"><a href="{spec}">版式规格页</a><a href="{docx}">样板 docx</a>{conflict}</p>
  </div>
</article>"""

STYLE = """
*{box-sizing:border-box}
body{margin:0;padding:44px 24px 72px;background:#f4f6f3;color:#18332e;
font-family:Inter,"PingFang SC","Microsoft YaHei",system-ui,-apple-system,sans-serif;line-height:1.65}
.wrap{max-width:1180px;margin:0 auto}
header{margin-bottom:28px;padding-bottom:22px;border-bottom:1px solid #dfe6e1}
h1{font-family:"Songti SC","Noto Serif SC",SimSun,serif;font-size:28px;font-weight:600;margin:0 0 8px;letter-spacing:1px}
.sub{margin:0 0 14px;color:#52655f;font-size:14px}
.stats{display:flex;flex-wrap:wrap;gap:10px;margin:0}
.stat{background:#fff;border:1px solid #dfe6e1;border-radius:10px;padding:10px 16px;font-size:13px;color:#52655f}
.stat b{display:block;font-size:20px;color:#18332e;font-weight:600;line-height:1.3}
h2.group{font-size:15px;font-weight:600;margin:34px 0 14px;color:#18332e;padding-left:10px;border-left:3px solid #176b57}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:16px}
.card{display:flex;gap:14px;background:#fff;border:1px solid #dfe6e1;border-radius:12px;padding:14px}
.thumb{flex:0 0 108px}
.thumb img{width:108px;border:1px solid #e4eae6;border-radius:3px;display:block;background:#fff}
.body{flex:1;min-width:0}
.row1{display:flex;align-items:baseline;gap:8px;justify-content:space-between}
h3{font-size:15px;margin:0;font-weight:600}
.chip{font-size:11px;padding:2px 8px;border-radius:20px;white-space:nowrap}
.en{margin:2px 0 8px;font-size:11px;color:#7c8a84;word-break:break-word}
dl{margin:0;display:grid;grid-template-columns:34px 1fr;gap:2px 8px;font-size:12px}
dt{color:#7c8a84}
dd{margin:0;color:#2b4640}
.links{margin:9px 0 0;font-size:12px;display:flex;gap:12px;flex-wrap:wrap}
.links a{color:#0e4f41;text-decoration:none;border-bottom:1px solid #bcd6cb}
.links .warn{color:#a44b3f;border-bottom:1px dotted #e3c2b9;cursor:help}
footer{margin-top:40px;padding-top:18px;border-top:1px solid #dfe6e1;color:#64756e;font-size:12px}
.note{background:#fff;border:1px solid #dfe6e1;border-left:3px solid #176b57;border-radius:10px;
padding:14px 18px;margin:0 0 8px;font-size:13px;color:#3d5751}
.note b{color:#18332e}
"""


def build(jdir, out):
    """生成总览页。坏文件逐个点名进返回值，不拖垮其余文件。

    返回失败清单 [(文件名, 原因)]——空列表表示全部成功。
    """
    jdir, out = Path(jdir), Path(out)
    rows, failures = [], []
    for p in sorted(jdir.glob("*.json")):
        try:
            rows.append(summarize(json.loads(p.read_text(encoding="utf-8"))))
        except Exception as e:  # noqa: BLE001 —— 一份手写坏 JSON 不能让总览整体崩溃
            failures.append((p.name, f"{type(e).__name__}: {e}"))
    zh = [r for r in rows if r["lang"] == "zh"]
    en = [r for r in rows if r["lang"] != "zh"]

    cards = []
    for group, items in (("中文刊", zh), ("英文刊", en)):
        cards.append(f'<h2 class="group">{group}　{len(items)} 本</h2><div class="grid">')
        for r in items:
            basis = r["auth_label"] + ("（有官方原件）" if r["official"] else "")
            t = thumb(jdir / f'{r["id"]}.sample.pdf')
            fields = dict(
                name=html.escape(r["name"]), name_en=html.escape(r["name_en"] or ""),
                page=html.escape(r["page"]), body=html.escape(r["body"]),
                spacing=html.escape(r["spacing"]),
                headings=html.escape("　".join(str(h) for h in r["headings"]) or "—"),
                citation=html.escape(r["citation"]), total=html.escape(r["total"]),
                basis=html.escape(basis),
                auth_label=html.escape(r["auth_label"]),
                auth_fg=r["auth_fg"], auth_bg=r["auth_bg"],
                thumb=(f'<img alt="{html.escape(r["name"])} 首页" '
                       f'src="data:image/jpeg;base64,{t}">' if t else ""),
                spec=r["spec"], docx=r["docx"],
                conflict=('<span class="warn" title="官方指南与现行排版的字号不一致，'
                      '两组值都记在规则集 verification.conflict 里">官方与实物有冲突</span>'
                      if r["conflict"] else ""),
            )
            cards.append(CARD.format(**fields))
        cards.append("</div>")

    n_official = sum(1 for r in rows if r["official"])
    n_template = sum(1 for r in rows if r["auth"] == "template")
    n_pub = sum(1 for r in rows if r["auth"] == "published")
    n_conflict = sum(1 for r in rows if r["conflict"])

    head = f"""
<header>
  <h1>期刊版式规则总览</h1>
  <p class="sub">{len(rows)} 本经管期刊的投稿版式，逐本对应一份样板 docx 与一页规格说明。依据顺序：官方模板 ＞ 官方明文 ＞ 已发表实物；实物只在官方未规定时补位。</p>
  <div class="stats">
    <div class="stat"><b>{len(rows)}</b>期刊总数</div>
    <div class="stat"><b>{n_official}</b>有官方原件存档</div>
    <div class="stat"><b>{n_template}</b>依据官方模板</div>
    <div class="stat"><b>{n_pub}</b>版式取自实物</div>
    <div class="stat"><b>{n_conflict}</b>官方与实物冲突</div>
  </div>
</header>
<div class="note">
  <b>怎么读：</b>「官方模板」＝有该刊下发的 Word 模板，版式照模板解析；「官方明文」＝官网体例/指南写明了字号行距，按文字落值；
  「实物实测」＝官方只给结构不给版式，字号行距取自该刊已发表论文，属出版版式，已按 0.98 系数反算缩印影响并落到半磅刻度。
  每本卡片右下角的「版式规格页」列出全部推定项与交叉核对记录，交付前应逐条过一遍。
</div>"""

    out_text = f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>期刊版式规则总览 · 格式工坊</title>
<style>{STYLE}</style></head>
<body><div class="wrap">{head}{''.join(cards)}
<footer>由 rules/journal/*.json 自动生成 · {date.today():%Y-%m-%d} · 与各刊样板、规格页同源</footer>
</div></body></html>"""

    Path(out).write_text(out_text, encoding="utf-8")
    print(f"总览页 → {out}（{len(rows)} 本，{len(out_text) // 1024} KB）")
    if failures:
        print(f"{len(failures)} 份规则集解析失败，已跳过：", file=sys.stderr)
        for name, why in failures:
            print(f"  失败：文件 {name}：{why}", file=sys.stderr)
    return failures


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default=str(JDIR / "index.html"))
    args = ap.parse_args()
    return 1 if build(JDIR, args.out) else 0


if __name__ == "__main__":
    raise SystemExit(main())
