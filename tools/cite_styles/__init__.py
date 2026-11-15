"""引用格式化器插件目录（DESIGN §9.2）。

`citation.style_id` 指向这里的实现。规则集里写了 `reference_list.templates`
时模板优先（数据驱动）；没有模板、但 style_id 能解析到插件时，由插件按
结构化字段（作者/年/题/刊/卷期页/DOI）渲染条目。

支持的样式族：
- APA-7 / APA
- HARVARD-ELSEVIER、HARVARD-WILEY（其余 HARVARD* 走 Elsevier 版式）
- CHICAGO / CHICAGO-AUTHOR-DATE

插件只改文献条目的写法，不增删文献、不改作者名与题名——与 §1.4 的
「唯一越线开关」口径一致。没有对应插件时返回 None，调用方保持原样。
"""

from . import apa, chicago, harvard


def resolve(style_id):
    """style_id → 格式化模块；不认识的样式返回 None（不猜）。"""
    s = (style_id or "").strip().upper()
    if not s:
        return None
    if s.startswith("APA"):
        return apa
    if s.startswith("HARVARD"):
        return harvard
    if s.startswith("CHICAGO"):
        return chicago
    return None


def format_entry(entry, rules):
    """按规则集声明的 style_id 渲染一条文献，返回 (text, italic_spans)。

    没有插件或该样式不处理这类条目时返回 (None, [])。
    """
    style = ((rules or {}).get("citation") or {}).get("style_id")
    mod = resolve(style)
    if mod is None:
        return None, []
    return mod.format_entry(entry, style or "")
