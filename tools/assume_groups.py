"""推定项展示分组的共享逻辑（F16NOISE，DESIGN §6 / journal-schema §6.1）。

背景：批五 F16 把继承披露扩到全量后，AI 建库草稿的顶层 `assumptions` 从几条
涨到 48+ 条。披露不许删（落盘条数只增不减），但规格页与审计页照旧平铺会把
复核重点淹没——修法是**展示分层**：≥12 条时按区块分组折叠（默认收起、计数
徽标、逐条可达），少量推定项（手工维护规则集的常态）保持平铺。

分组依据是规则集自己的留痕（`verification.extraction.inherited`，见
journal-schema §6.1）：extract_journal 生成的派生假设是
「`{label} = {value}（未从稿约抽到，沿用基线值，须人工核对）`」格式的字符串，
以继承尾注认领派生条目、再以「label = 」前缀对回 inherited 的 label → section。
对不上号的条目（手写假设、结构字段汇总）归入「其他推定」组，照常展示，
不因分组丢失任何一条。

规则集是唯一真源：本模块只读它给的结构化留痕，不内置任何刊名或字段表。
"""

from __future__ import annotations

import re

# extract_journal 派生事实键假设的固定尾注（tools/extract_journal.py 的
# rules["assumptions"] 生成处）。以此认领「哪条是继承派生的」，手写假设
# 不带这个尾注，天然分流到「其他推定」。
INHERITED_SUFFIX = "（未从稿约抽到，沿用基线值，须人工核对）"

# 「其他推定」组的展示名：手写假设与结构字段汇总条的对号失败兜底。
OTHER = "其他推定"

# 少于阈值保持平铺——手工维护的规则集通常三五条，平铺读起来更快。
GROUP_THRESHOLD = 12

# 「label = 值…」前缀。label 是 extract_journal FACTS 的中文名（或退化成
# 点路径），不含「 = 」；值里可能含「 = 」（如映射的 tojson），所以从左侧
# 只切第一个「 = 」。
_LABEL_RE = re.compile(r"^([^=]{1,60}?)\s*=\s*(.*)$", re.S)


def group_assumptions(assumptions, extraction=None):
    """把 assumptions 字符串数组整理成 [(section_name, [str, …])]。

    extraction 传 `verification.extraction`（有就按它的 inherited 对号）。
    返回列表保持出现顺序（Python dict 保序）；对不上号的一律进「其他推定」。
    纯函数：同输入同输出，不修改入参。
    """
    items = list(assumptions or [])
    if not items:
        return []

    by_label = {}
    for it in ((extraction or {}).get("inherited") or []):
        label = (it or {}).get("label")
        section = (it or {}).get("section")
        if label and section and label not in by_label:
            by_label[str(label)] = str(section)

    groups: dict[str, list[str]] = {}
    for a in items:
        section = OTHER
        text = str(a)
        if text.endswith(INHERITED_SUFFIX):
            m = _LABEL_RE.match(text)
            if m and m.group(1) in by_label:
                section = by_label[m.group(1)]
        groups.setdefault(section, []).append(text)
    return list(groups.items())


def should_group(groups):
    """是否走分组折叠展示：总条数到阈值才折叠，其余平铺。"""
    return sum(len(rows) for _, rows in groups) >= GROUP_THRESHOLD


def section_map(assumptions, extraction=None):
    """assumption → section 的一一映射（audit JSON 的 assumption_sections 用）。

    与 group_assumptions 同一对号口径；渲染器侧据此自组分组，避免审计 JSON
    里重复存整个数组。
    """
    m = {}
    for section, rows in group_assumptions(assumptions, extraction):
        for r in rows:
            m[r] = section
    return m
