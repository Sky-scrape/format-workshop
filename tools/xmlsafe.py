"""实体安全的 XML 解析入口。

合法的 docx 部件（document.xml、styles.xml、footnotes.xml……）不会带 DTD 或
实体声明；带了的只可能是实体扩展炸弹（billion laughs）一类的东西，一律拒收。
各工具不要直接碰 ET / lxml 的裸解析入口，从这里走。
"""


class UnsafeXml(ValueError):
    """XML 里带 DOCTYPE 或实体声明。"""


def reject_dtd(data):
    """bytes 里出现 DOCTYPE 或 ENTITY 声明就抛 UnsafeXml。

    XML 的声明区分大小写且必须出现在根元素之前，按字面量检查即完备；
    数字字符引用（&#x...;）不在此列，它们不是实体扩展。
    """
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise UnsafeXml("XML 带 DOCTYPE/实体声明，已拒收")


def fromstring(data):
    """reject_dtd 之后的 ElementTree.fromstring。只收 bytes。"""
    import xml.etree.ElementTree as ET

    reject_dtd(data)
    return ET.fromstring(data)
