""".doc 接入：认出它到底是什么，再决定怎么转。

`.doc` 这个后缀在真实投稿里至少代表四种东西：Word 97-2003 二进制、改了后缀的
.docx、RTF、以及根本不是 Word 文档的垃圾。之前一律交给 LibreOffice，于是：

- 改了后缀的 .docx 白白绕一趟 Office（5 秒起），还可能被版本转换改掉版式
- 垃圾文件被 Word「尽力打开」成一份 1 个块的稿子，全程不报错，用户拿到空稿
  还以为工具坏了

这里守的就是「先认身份，再动转换」。涉及真实 Office 套件的用例在没有相应
环境的机器上跳过——CI 上既没有 Word 也不保证有 LibreOffice。
"""

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app  # noqa: E402

SAMPLE_DOCX = ROOT / "rules" / "journal" / "tjyj.sample.docx"
# 样板是 build_journal.py 的渲染产物（开源仓库只放行三本示例的 json+spec），
# 公开检出里没有——用到它的用例跳过，而不是报「文件不存在」的假回归
needs_sample = pytest.mark.skipif(
    not SAMPLE_DOCX.is_file(), reason="tjyj.sample.docx 由 build_journal 生成，不进开源仓库")

OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _has_converter():
    """这台机器上有没有能真正转 .doc 的软件。"""
    if shutil.which("soffice"):
        return True
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if Path(p).is_file():
            return True
    return False


needs_converter = pytest.mark.skipif(
    not _has_converter(), reason="本机没有 LibreOffice，真实 .doc 转换无法验证")


# ---------------------------------------------------------------- 身份识别

def test_sniff_recognizes_real_doc(tmp_path):
    """OLE2 复合文档头 = Word 97-2003 的真 .doc。"""
    p = tmp_path / "a.doc"
    p.write_bytes(OLE2 + b"\x00" * 64)
    assert app._sniff_doc(p) == "doc"


@needs_sample
def test_sniff_recognizes_renamed_docx(tmp_path):
    """.docx 本质是 zip，改了后缀也认得出来。"""
    p = tmp_path / "b.doc"
    shutil.copy2(SAMPLE_DOCX, p)
    assert app._sniff_doc(p) == "docx"


def test_sniff_recognizes_rtf(tmp_path):
    p = tmp_path / "c.doc"
    p.write_bytes(b"{\\rtf1\\ansi hello}")
    assert app._sniff_doc(p) == "rtf"


def test_sniff_rejects_plain_bytes(tmp_path):
    """纯文本、乱码、损坏文件一律判 unknown，不许交给 Word 去「尽力打开」。"""
    p = tmp_path / "d.doc"
    p.write_text("这只是一段纯文本。", encoding="utf-8")
    assert app._sniff_doc(p) == "unknown"


def test_sniff_missing_file_is_unknown(tmp_path):
    """读不到就按 unknown 处理，不能因为一次 OSError 把请求打成 500。"""
    assert app._sniff_doc(tmp_path / "nope.doc") == "unknown"


# ---------------------------------------------------------------- 分派逻辑

@needs_sample
def test_renamed_docx_is_used_directly(tmp_path):
    """改了后缀的 docx 直接复制，不绕 Office。

    这条同时守一个性能事实：绕一圈要 5 秒以上，而用户根本不需要转。
    """
    src = tmp_path / "renamed.doc"
    shutil.copy2(SAMPLE_DOCX, src)
    out = app.doc_to_docx(src, tmp_path)
    assert out.exists() and out.suffix == ".docx"
    assert out.stat().st_size == src.stat().st_size


def test_garbage_file_is_rejected_with_reason(tmp_path):
    """垃圾文件必须在转换前拦下，且说清是「不是 Word 文档」而不是「转不了」。"""
    src = tmp_path / "garbage.doc"
    src.write_bytes(b"NOT-A-WORD-FILE" * 40)
    with pytest.raises(app.ConvertError) as ei:
        app.doc_to_docx(src, tmp_path)
    assert "不是 Word 文档" in ei.value.title
    # 不许留下半成品：下游拿到一个残缺 docx 会一路错下去
    assert not (tmp_path / "garbage.docx").exists()


@needs_converter
@needs_sample
def test_real_doc_converts_and_keeps_structure(tmp_path):
    """真 .doc 能转出来，且转换不吞内容。

    用 LibreOffice 现造一份 .doc（它自带的 Word 97 导出），再走本模块的
    转换链，最后回读段落数——只断言「能转」是不够的，转出一份空文档
    同样会让下游一路空跑。
    """
    doc = tmp_path / "src.docx"
    shutil.copy2(SAMPLE_DOCX, doc)
    exe = shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"
    subprocess.run([exe, "--headless",
                    f"-env:UserInstallation={(tmp_path / 'lo').as_uri()}",
                    "--convert-to", "doc", "--outdir", str(tmp_path), str(doc)],
                   capture_output=True, timeout=180)
    legacy = tmp_path / "src.doc"
    if not legacy.exists():
        pytest.skip("LibreOffice 没能造出 .doc 样本")
    assert app._sniff_doc(legacy) == "doc"

    before = _paragraphs(doc)
    out = app.doc_to_docx(legacy, tmp_path)
    after = _paragraphs(out)
    assert len(after) == len(before), f"转换前后段落数不一致：{len(before)} → {len(after)}"
    assert "".join(after).strip(), "转出来的文档没有文字"


def _paragraphs(docx_path):
    """读 docx 的段落文本，用于比较转换前后的内容量。"""
    from docx import Document

    return [p.text for p in Document(str(docx_path)).paragraphs if p.text.strip()]
