"""审计里的 `rule` 字段指向规则集的哪一处。

DESIGN §6 要求每条改动都能回溯到规则集条目，§8.5 要求这个回溯**可验证**。
只写一个好看的名字不算回溯——指过去必须真能落到一个值上。

两种写法：

    "typography.body.size_pt"           纯字典路径
    "blocks:keywords/zh/separator"      在 rules["blocks"] 里按 type + lang 找到条目，再取键

第二种是为 blocks 准备的——它是个数组，光靠字典路径指不进去。
lang 可省略（`blocks:references/label`），表示该类型只有一条。

解析不到一律抛 KeyError。审计里出现一个指不到实处的引用，比没有引用更糟：
它看起来像证据。
"""


def resolve(rules, ref):
    """把引用解析成规则集里的实际值。解析不到抛 KeyError。

    引用以 `~` 结尾表示"该键可缺省"（部分期刊没写这一项，此时渲染器走代码内
    约定）。这类引用解析不到时返回 None，而不是报错——但它的父路径仍须存在，
    否则一样抛 KeyError：不能拿"可缺省"当挡箭牌去指向不存在的地方。
    """
    if not ref:
        raise KeyError("空的规则引用")

    optional = ref.endswith("~")
    if optional:
        ref = ref[:-1]

    if ref.startswith("blocks:"):
        parts = ref[len("blocks:"):].split("/")
        btype, rest = parts[0], parts[1:]
        lang = None
        if rest and rest[0] in ("zh", "en"):
            lang = rest.pop(0)

        node = next((b for b in rules.get("blocks", [])
                     if b.get("type") == btype
                     and (lang is None or b.get("lang") == lang)), None)
        if node is None:
            if optional:
                return None
            raise KeyError(f"规则集里没有 blocks type={btype} lang={lang}")
    else:
        node, rest = rules, ref.split(".")

    for i, key in enumerate(rest):
        if not isinstance(node, dict) or key not in node:
            if optional and i == len(rest) - 1:
                return None
            raise KeyError(f"引用 {ref} 在 {key} 处断掉")
        node = node[key]
    return node


def resolvable(rules, ref):
    """便于在测试里做批量断言，不抛异常。"""
    try:
        resolve(rules, ref)
        return True
    except KeyError:
        return False
