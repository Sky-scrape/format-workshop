"""规则集键名归一：入库样本只许使用规范写法（2026-11-13 收口）。

历史上有两种总字数下限写法（limits.total.min / limits.total_min）与两种
标题上限写法（title_chars_max / title_max_chars）。读取侧为兼容外部旧
规则集保留双读，但仓库自带的样本必须只有规范键——否则新增规则会照着
旧写法抄，永远收不拢。迁移工具：tools/normalize_keys.py。
"""

import json
import sys
from pathlib import Path

from conftest import RULES, SHIPPED_JOURNALS, TOOLS

sys.path.insert(0, str(TOOLS))

from normalize_keys import migrate  # noqa: E402


def test_shipped_rulesets_use_canonical_limit_keys():
    offenders = []
    for jid in SHIPPED_JOURNALS:
        rules = json.loads((RULES / f"{jid}.json").read_text(encoding="utf-8"))
        limits = rules.get("limits") or {}
        if "total_min" in limits:
            offenders.append(f"{jid}: limits.total_min")
        if "title_max_chars" in limits:
            offenders.append(f"{jid}: limits.title_max_chars")
        total = limits.get("total") or {}
        if "total_min" in total:
            offenders.append(f"{jid}: limits.total.total_min")
    assert not offenders, (
        "入库规则集仍在用旧键名，先跑 py -3 tools/normalize_keys.py：\n"
        + "\n".join(offenders))


def test_migrate_moves_legacy_keys_and_is_idempotent():
    rules = {"limits": {"total": {"value": 20000}, "total_min": 15000,
                        "title_max_chars": 22}}
    out, changes = migrate(rules)
    assert out["limits"]["total"]["min"] == 15000
    assert out["limits"]["title_chars_max"] == 22
    assert "total_min" not in out["limits"]
    assert "title_max_chars" not in out["limits"]
    assert {c[0] for c in changes} == {"limits.total_min", "limits.title_max_chars"}

    again, changes2 = migrate(out)
    assert again == out, "第二次迁移又改了内容，不是幂等操作"
    assert changes2 == []


def test_migrate_prefers_existing_canonical_value():
    rules = {"limits": {"total": {"min": 12000}, "total_min": 15000,
                        "title_chars_max": 20, "title_max_chars": 22}}
    out, changes = migrate(rules)
    assert out["limits"]["total"]["min"] == 12000, "迁移覆盖了已核对过的规范值"
    assert out["limits"]["title_chars_max"] == 20
    assert len(changes) == 2
