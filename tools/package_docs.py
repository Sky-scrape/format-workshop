"""分发包根目录允许携带的文档白名单（D-08）。

AGPL 包必须带 LICENSE 与项目文档，但根目录散落的内部材料不随包外流——
实测报告引用私有稿件（训练论文/…）与本机完整规则库规模（23 本），两者
都是仓库刻意不公开的素材（.gitignore 的排除理由）。build_portable 复制时
按此白名单过滤根目录 *.md，build_installer 对旧构建残留的 dist 再清一遍。
"""

from pathlib import Path

ROOT_DOC_WHITELIST = {"README.md", "DESIGN.md", "CONTRIBUTING.md"}


def is_private_doc(rel_parts, name):
    """根目录一层（且仅根目录一层）的 *.md 不在白名单即私有。

    子目录的 .md（如 rules/ 下的说明）不受影响；LICENSE 无 .md 后缀、
    AGPL 要求随包携带，不在本列。
    """
    return len(rel_parts) == 1 and name.endswith(".md") and name not in ROOT_DOC_WHITELIST


def strip_private_docs(portable):
    """清掉已构建包根目录里不该分发的散落文档（旧构建残留）。"""
    for p in Path(portable).glob("*.md"):
        if p.name not in ROOT_DOC_WHITELIST:
            p.unlink(missing_ok=True)
