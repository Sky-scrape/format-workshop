"""把规则集里两种历史键名收敛到规范写法（数据侧迁移，可重复执行）。

    limits.total_min       → limits.total.min
    limits.title_max_chars → limits.title_chars_max

读取侧（compliance_report / render_spec / app.py）仍保留旧键兼容：外部手工
维护的规则集、旧版本导出的 JSON 不该因为升级工具就整份失效。但**仓库自带
的规则集**必须只有规范键——两种写法并存会让新增规则时「照着隔壁抄」抄出
第三种。tests/test_fix_ruleset_keys.py 对入库样本做静态断言。

用法：
    python tools/normalize_keys.py            # 就地迁移 rules/journal/*.json
    python tools/normalize_keys.py --check    # 只检查，不写；有旧键时退出码 1
"""

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JDIR = ROOT / "rules" / "journal"


def migrate(rules):
    """返回 (rules, changes)。changes 是 [(旧键, 新键, 值)]，便于--check 汇报。"""
    changes = []
    limits = rules.get("limits")
    if not isinstance(limits, dict):
        return rules, changes

    if "total_min" in limits:
        value = limits.pop("total_min")
        total = limits.get("total")
        if not isinstance(total, dict):
            total = {}
            limits["total"] = total
        # 已有规范键时以规范键为准：迁移不覆盖人工核对过的值
        if total.get("min") is None:
            total["min"] = value
            changes.append(("limits.total_min", "limits.total.min", value))
        else:
            changes.append(("limits.total_min", "（已存在 limits.total.min，丢弃重复值）", value))

    if "title_max_chars" in limits:
        value = limits.pop("title_max_chars")
        if limits.get("title_chars_max") is None:
            limits["title_chars_max"] = value
            changes.append(("limits.title_max_chars", "limits.title_chars_max", value))
        else:
            changes.append(("limits.title_max_chars", "（已存在 limits.title_chars_max，丢弃重复值）", value))
    return rules, changes


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只检查，不写文件")
    ap.add_argument("--dir", default=str(JDIR), help="规则集目录")
    a = ap.parse_args(argv)

    jdir = Path(a.dir)
    files = sorted(jdir.glob("*.json"))
    if not files:
        print(f"没有找到规则集：{jdir}", file=sys.stderr)
        return 2

    stale = 0
    for path in files:
        try:
            raw = path.read_text(encoding="utf-8")
            rules = json.loads(raw)
        except Exception as e:  # noqa: BLE001
            print(f"失败：文件 {path.name}：{type(e).__name__}: {e}", file=sys.stderr)
            return 2
        rules, changes = migrate(rules)
        if not changes:
            continue
        stale += 1
        for old, new, value in changes:
            print(f"{path.name}: {old} = {value!r} → {new}")
        if not a.check:
            path.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n",
                            encoding="utf-8")
    if a.check:
        if stale:
            print(f"\n{stale} 份规则集仍在使用旧键名；跑一遍 tools/normalize_keys.py 迁移。",
                  file=sys.stderr)
            return 1
        print(f"检查通过：{len(files)} 份规则集都使用规范键名。")
    else:
        print(f"已处理 {stale} 份规则集（其余无需改动）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
