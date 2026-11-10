"""实体安全的 XML 解析入口。

合法的 docx 部件（document.xml、styles.xml、footnotes.xml……）不会带 DTD 或
实体声明；带了的只可能是实体扩展炸弹（billion laughs）一类的东西，一律拒收。
各工具不要直接碰 ET / lxml 的裸解析入口，从这里走。zip 部件的读取也在这里
把关：解压后大小超限的部件同样拒收（read_part）。
"""


class UnsafeXml(ValueError):
    """XML 里带 DOCTYPE 或实体声明，或编码可疑。"""


class UnsafeZip(ValueError):
    """zip 部件解压后大小超限。"""


# 单部件解压上限：恶意 docx 用 ~1000:1 的压缩比可把 80MB 的上传闸门放大成
# 数十 GB 内存（F2-3）。真实稿件的 XML 部件（正文+样式+脚注全在内）远到不了
# 这个数；图片等大块媒体部件本工具从不读，不设总闸免得误伤图多的真实稿件。
PART_MAX_BYTES = 50 * 1024 * 1024


def _wide_encoding(data):
    """按 BOM / 首 2-4 字节识别 UTF-16/UTF-32 宽编码，返回解码名或 None。

    XML 规范附录 F 规定：无 BOM 的 UTF-16/32 文档必须以「<?」的宽编码形式
    （3C 00 / 00 3C / 3C 00 00 00 / 00 00 00 3C）开头才可能被解析器认出，
    所以按首字节判别与解析器的自检测口径一致。
    """
    # UTF-32 的 BOM 与 UTF-16LE 的 BOM 有公共前缀（FF FE），必须先判宽的
    if data.startswith(b"\x00\x00\xfe\xff"):
        return "utf-32-be"
    if data.startswith(b"\xff\xfe\x00\x00"):
        return "utf-32-le"
    if data.startswith(b"\xff\xfe"):
        return "utf-16-le"
    if data.startswith(b"\xfe\xff"):
        return "utf-16-be"
    if data.startswith(b"\x3c\x00\x00\x00"):
        return "utf-32-le"
    if data.startswith(b"\x00\x00\x00\x3c"):
        return "utf-32-be"
    if data.startswith(b"\x3c\x00"):
        return "utf-16-le"
    if data.startswith(b"\x00\x3c"):
        return "utf-16-be"
    return None


def reject_dtd(data):
    """bytes 里出现 DOCTYPE 或 ENTITY 声明就抛 UnsafeXml。

    XML 的声明区分大小写且必须出现在根元素之前，编码展开后按字面量检查即
    完备；数字字符引用（&#x...;）不在此列，它们不是实体扩展。

    UTF-16/32 编码的部件里 ASCII 字面量被 NUL 隔开、直接匹配不到（F2-1），
    守卫先按 BOM/首字节识别宽编码展开成 UTF-8 再查——否则
    encoding="utf-16" 的实体炸弹能绕过这里直达解析器。展开不了的带 NUL
    数据不是合法 XML（U+0000 非法字符，pyexpat 也解不动），一并拒收。
    """
    enc = _wide_encoding(data)
    if enc is not None:
        try:
            data = data.decode(enc).encode("utf-8")
        except UnicodeDecodeError:
            raise UnsafeXml(f"XML 声称是 {enc} 但解不开，已拒收")
    elif b"\x00" in data:
        # 不是宽编码却含 NUL：任何合法 XML 都不会有这个字节，别让
        # 未知的编码戏法把 DOCTYPE 藏过去（fail-closed）
        raise UnsafeXml("XML 含 NUL 字节且不是 UTF-16/32，已拒收")
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise UnsafeXml("XML 带 DOCTYPE/实体声明，已拒收")


def read_part(zf, name):
    """读 zip 部件，声明解压大小超限即抛 UnsafeZip（F2-3）。

    zipfile 的 read 受条目声明的 file_size 约束，读前查它就能挡住压缩比
    炸弹：不查的话 ~200KB 的压缩包可以在内存里当场解出 200MB。
    """
    info = zf.getinfo(name)
    if info.file_size > PART_MAX_BYTES:
        raise UnsafeZip(
            f"zip 部件 {name} 解压后 {info.file_size} 字节，"
            f"超过 {PART_MAX_BYTES} 上限，已拒收")
    return zf.read(name)


def fromstring(data):
    """reject_dtd 之后的 ElementTree.fromstring。只收 bytes。"""
    import xml.etree.ElementTree as ET

    reject_dtd(data)
    return ET.fromstring(data)
