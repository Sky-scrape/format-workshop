"""把审计 JSON 渲染成人类可读的 HTML。

审计报告的读者是作者本人，他需要能逐条核对"到底改了什么"。
因此这里的原则是：
  1. 原值/新值必须完整（超长才截断，且带省略号与全文悬停），绝不静默切字；
  2. 说明永远显示，不能因为已有 diff 就吞掉"为什么改"；
  3. 同类版式操作折叠计数，避免 24 条"参考文献悬挂缩进"淹没真正的文本改动。

用法：python tools/render_audit.py <audit.json> -o out.html
"""

import argparse
import html
import json
import re
from pathlib import Path

from report_theme import BASE

KIND = {
    "页面设置": "A",
    "文章标题套版": "A",
    "副标题套版": "A",
    "作者署名套版": "A",
    "标题重新编号": "B",
    "题注归一": "B",
    "标签归一": "B",
    "参考文献悬挂缩进": "A",
    "表格重设为三线表": "A",
    "表注／图注层序调整": "B",
}

GRADE_DESC = {
    "A": "纯版式，不触及任何字符",
    "B": "触及字符，但不动字词",
}


def clip(text, limit=48):
    """压缩空白并按需截断，超长时补省略号并保留全文供悬停查看。"""
    raw = re.sub(r"\s+", " ", text or "").strip()
    if len(raw) <= limit:
        return html.escape(raw)
    return (f'<span title="{html.escape(raw)}">{html.escape(raw[:limit])}…</span>')


def render(d, job=None):
    items = d["items"]
    for it in items:
        it["grade"] = KIND.get(it["action"], "A")

    text_items = [it for it in items
                  if it.get("before") and it.get("after") and it["before"] != it["after"]]
    layout_items = [it for it in items if it not in text_items]

    n_rejected = sum(1 for it in items if it.get("rejected"))
    # 可驳回的只有 B 档（触及字符）与表注层序：A 档是纯版式，驳回后成品就不合规了
    REJECTABLE = {"标题重新编号", "题注归一", "标签归一", "文献表标题归一", "表注／图注层序调整"}
    col = (5 if job else 4)

    def reject_cell(it):
        if not job:
            return ""
        if it["action"] not in REJECTABLE:
            return '<td class="rj"></td>'
        ck = " checked" if it.get("rejected") else ""
        return (f'<td class="rj"><label class="rjbox"><input type="checkbox" name="rid"'
                f' value="{html.escape(it.get("id", ""))}"{ck}>'
                f'<span>保留原文</span></label></td>')

    # 版式操作按动作分组，保留首次出现顺序
    groups = {}
    for it in layout_items:
        groups.setdefault(it["action"], []).append(it)

    text_rows = "".join(
        f'<tr data-k="{html.escape(it["action"] + " " + it.get("note", ""))}"'
        f'{" class=rejected" if it.get("rejected") else ""}>'
        f'<td class="pos">{it["idx"] if it["idx"] >= 0 else "页面"}'
        f'<span class="sec">{html.escape(it.get("sec") or "")}</span></td>'
        f'<td class="act">{html.escape(it["action"])}'
        f'{"<span class=tag>已驳回</span>" if it.get("rejected") else ""}</td>'
        f'<td class="diff"><span class="old">{clip(it["before"])}</span>'
        f'<span class="arrow">→</span><span class="new">{clip(it["after"])}</span></td>'
        f'<td class="note">{html.escape(it.get("note", ""))}'
        f'{"<code class=ref>" + html.escape(it["rule"]) + "</code>" if it.get("rule") else ""}</td>'
        f'{reject_cell(it)}</tr>'
        for it in text_items
    ) or f'<tr><td colspan="{col}" class="empty">本次没有触及任何字符的改动</td></tr>'

    group_html = []
    for action, rows in groups.items():
        grade = rows[0]["grade"]
        inner = "".join(
            f'<div class="grow"><span class="gpos">{it["idx"] if it["idx"] >= 0 else "页面"}</span>'
            f'<span class="gnote">{html.escape(it.get("note") or it.get("after") or "—")}</span>'
            f'{"<code class=gref>" + html.escape(it["rule"]) + "</code>" if it.get("rule") else ""}</div>'
            for it in rows)
        group_html.append(
            f'<details class="group"{" open" if len(rows) <= 6 else ""}>'
            f'<summary><span class="badge k{grade}">{grade}</span>'
            f'<span class="gname">{html.escape(action)}</span>'
            f'<span class="gcount">{len(rows)} 处</span></summary>'
            f'<div class="glist">{inner}</div></details>'
        )
    layout_html = "".join(group_html) or '<p class="empty">本次没有纯版式操作</p>'

    summary = " · ".join(f"{k} {len(v)}" for k, v in sorted(groups.items(),
                                                           key=lambda x: -len(x[1])))

    anon = d.get("anonymize") or {}
    anon_html = ""
    if anon.get("removed"):
        rows = "".join(f"<li>{html.escape(x)}</li>" for x in anon["removed"])
        anon_html = (f'<section id="anonSection"><h2>匿名稿处理（另存一份，不影响署名稿）</h2>'
                     f'<p class="meta">该刊要求正文匿名，工具另出一份匿名稿，删除下列 '
                     f'{len(anon["removed"])} 处身份信息；署名稿保持完整，供留档与校样使用。</p>'
                     f'<ul class="anon">{rows}</ul></section>')

    out = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>审计报告 · {html.escape(d['target'])}</title>
<style>{BASE}
.kpi{{display:flex;gap:0;margin:0 0 20px;background:var(--surface);
border:1px solid var(--line);border-radius:14px;overflow:hidden}}
.kpi div{{flex:1;padding:15px 18px;border-left:1px solid var(--line2)}}
.kpi div:first-child{{border-left:none}}
.kpi b{{display:block;font-size:21px;font-weight:600;color:var(--dark);line-height:1.3;
font-variant-numeric:tabular-nums;letter-spacing:.5px}}
.kpi span{{display:block;font-size:12px;color:var(--ink3);margin-top:3px}}
.kpi .hl b{{color:var(--accent)}}
.pos{{color:var(--ink3);font-variant-numeric:tabular-nums;white-space:nowrap;width:96px;
font-size:12.5px;line-height:1.5}}
.pos .sec{{display:block;font-size:10.5px;color:#8b9a94;max-width:88px;overflow:hidden;
text-overflow:ellipsis;white-space:nowrap;margin-top:1px}}
.act{{white-space:nowrap;width:132px}}
.diff{{line-height:1.9}}
.old{{color:var(--danger);text-decoration:line-through;text-decoration-color:#d9aca2}}
.old,.new{{background:var(--soft);border-radius:3px;padding:1px 4px;word-break:break-all}}
.new{{color:var(--accent)}}
.arrow{{color:var(--ink3);margin:0 7px}}
.note{{color:var(--ink2);font-size:12.5px;line-height:1.75}}
.ref{{display:block;margin-top:4px;font-size:11px;color:var(--ink3);letter-spacing:.2px}}
.gref{{margin-left:9px;font-size:10.5px;color:var(--ink3)}}
.empty{{color:var(--ink3);font-size:13px;padding:6px 0}}
.badge{{display:inline-block;font-size:11px;font-weight:500;padding:2px 8px;
border-radius:6px;margin-right:9px}}
.kA{{background:var(--accent-soft);color:var(--accent)}}
.kB{{background:var(--gold-bg);color:var(--gold)}}
details.group{{border-bottom:1px solid var(--line2)}}
details.group:last-child{{border-bottom:none}}
details.group>summary{{display:flex;align-items:center;gap:0;padding:11px 0;cursor:pointer;
font-size:13.5px;list-style:none}}
details.group>summary::-webkit-details-marker{{display:none}}
details.group>summary:before{{content:"▸";color:var(--ink3);margin-right:9px;
font-size:11px;transition:transform .15s}}
details.group[open]>summary:before{{transform:rotate(90deg)}}
.gname{{font-weight:500}}
.gcount{{margin-left:auto;color:var(--ink3);font-size:12px;font-variant-numeric:tabular-nums}}
table.inner{{margin:0 0 8px 20px;width:calc(100% - 20px)}}
#textTable{{table-layout:fixed}}
.glist{{margin:2px 0 14px 50px}}
.grow{{display:flex;gap:14px;padding:7px 0;border-top:1px dashed var(--line2);
font-size:12.5px;line-height:1.7}}
.grow:first-child{{border-top:none;padding-top:2px}}
.gpos{{width:34px;flex:none;color:var(--ink3);font-variant-numeric:tabular-nums}}
.gnote{{color:var(--ink2);min-width:0;word-break:break-word}}
table.inner td{{padding:6px 10px 6px 0;border-top:1px dashed var(--line2);font-size:12.5px}}
.toolbar{{display:flex;gap:8px;align-items:center;margin:0 0 14px;flex-wrap:wrap}}
.toolbar button{{border:1px solid var(--line);background:var(--surface);color:var(--ink2);
font-size:12.5px;padding:6px 13px;border-radius:999px;cursor:pointer}}
.toolbar button.on{{background:var(--accent-soft);border-color:#cfe3d9;color:var(--accent)}}
.toolbar input{{flex:1;min-width:150px;border:1px solid var(--line);background:var(--surface);
border-radius:999px;padding:7px 14px;font-size:12.5px;color:var(--ink);outline:none}}
.toolbar input:focus{{border-color:#a9c8ba}}
.legend{{color:var(--ink3);font-size:12px;margin:14px 0 0;line-height:1.9}}
.rejectbar{{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:0 0 12px;
padding:10px 14px;border:1px solid #e6dcc6;background:var(--gold-bg);border-radius:10px;
font-size:12.5px;color:var(--gold)}}
.rejectbar span{{flex:1;min-width:200px;line-height:1.7}}
.rejectbar button{{border:1px solid #ddc9a5;background:var(--surface);color:var(--gold);
font-size:12.5px;padding:7px 16px;border-radius:999px;cursor:pointer;white-space:nowrap}}
.rejectbar button:hover{{background:#fffdf7}}
.rj{{width:104px}}
.rjbox{{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--ink3);cursor:pointer}}
.rjbox input{{width:14px;height:14px;cursor:pointer;accent-color:var(--gold)}}
.tag{{display:inline-block;margin-left:7px;font-size:10.5px;padding:1px 6px;border-radius:5px;
background:#f2ece0;color:var(--gold)}}
tr.rejected .old{{text-decoration:none;color:var(--ink3)}}
tr.rejected .new{{color:var(--ink3);text-decoration:line-through}}
.anon li{{margin:4px 0;font-size:12.5px;color:var(--ink2)}}
@media(max-width:640px){{
  .kpi{{flex-wrap:wrap}}.kpi div{{flex:1 1 45%;border-left:none;border-top:1px solid var(--line2)}}
  .kpi div:nth-child(-n+2){{border-top:none}}
  table{{font-size:12.5px}}.act{{width:auto;white-space:normal}}
}}
</style></head><body><div class="wrap">
<header>
  <h1>审计报告</h1>
  <p class="sub">{html.escape(d['source'].replace('.docx', ''))} → <strong>{html.escape(d['target'])}</strong></p>
  <p class="meta">共 {len(items)} 处调整 · {html.escape(summary) if summary else '无'}{f' · 已驳回 {n_rejected} 条' if n_rejected else ''}</p>
</header>

<div class="kpi">
  <div class="hl"><b>{len(text_items)}</b><span>字符级改动（B 档）</span></div>
  <div><b>{len(layout_items)}</b><span>版式操作（A 档）</span></div>
  <div><b>{n_rejected}</b><span>已驳回的改动</span></div>
  <div><b>3</b><span>字数口径（A / B / C）</span></div>
</div>

<section id="textSection">
  <h2>触及字符的改动</h2>
  <div class="toolbar no-print">
    <button class="on" data-f="all" type="button">全部</button>
    <button data-f="text" type="button">仅字符改动</button>
    <button data-f="layout" type="button">仅版式</button>
    <input id="q" type="search" placeholder="搜索改动内容或说明">
  </div>
  {('<form method="post" action="/job/' + job + '/reject" target="_top" id="rejectForm">'
    '<div class="rejectbar no-print"><span>勾选「保留原文」即驳回该条改动，重渲染时按你的选择执行，'
    '其余改动保持不变。</span>'
    '<button type="submit">按选择重渲染</button></div>') if job else ''}
  <table id="textTable">
    <colgroup><col style="width:96px"><col style="width:96px"><col><col style="width:152px">
    <col style="width:104px"></colgroup>
    <thead><tr><th>段落</th><th>动作</th><th>原值 → 新值</th><th>说明</th>
    <th>驳回</th></tr></thead>
    <tbody>{text_rows}</tbody>
  </table>
  {('</form>') if job else ''}
  <p class="legend">原值与新值为完整文本，过长时以「…」省略，悬停可看全文。
  段落编号指稿件解析后的段落顺序（从 0 开始），「页面」表示整篇页面设置。</p>
</section>

<section id="layoutSection">
  <h2>版式操作</h2>
  {layout_html}
</section>

{anon_html}

<p class="meta" style="margin-top:20px">A 档＝纯版式，不触及文字；B 档＝触及字符但不动字词，全部可回溯。
正文与表格内文字零改动，本报告可直接用于投稿说明或导师核对。</p>
</div>
<script>
const q=document.getElementById('q'),textSection=document.getElementById('textSection'),
      layoutSection=document.getElementById('layoutSection');
document.querySelectorAll('.toolbar button').forEach(b=>b.addEventListener('click',()=>{{
  document.querySelectorAll('.toolbar button').forEach(x=>x.classList.toggle('on',x===b));
  const f=b.dataset.f;
  textSection.style.display=(f==='layout')?'none':'';
  layoutSection.style.display=(f==='text')?'none':'';
}}));
q.addEventListener('input',()=>{{
  const v=q.value.trim().toLowerCase();
  document.querySelectorAll('#textTable tbody tr').forEach(tr=>{{
    tr.hidden=!!v&&!tr.textContent.toLowerCase().includes(v);
  }});
  document.querySelectorAll('details.group').forEach(g=>{{
    g.hidden=!!v&&!g.textContent.toLowerCase().includes(v);
    if(v)g.open=true;
  }});
}});
</script>
</body></html>"""
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("audit")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--job", help="任务号；给了才会渲染「驳回」勾选与重渲染表单")
    a = ap.parse_args()
    d = json.loads(Path(a.audit).read_text(encoding="utf-8"))
    Path(a.out).write_text(render(d, job=a.job), encoding="utf-8")
    n = sum(1 for it in d["items"] if it.get("rejected"))
    print("已生成:", a.out, f"({len(d['items'])} 条" + (f"，已驳回 {n} 条)" if n else ")"))


if __name__ == "__main__":
    main()
