"""F19 回归钉子：OCR 失败不能静默——丢页必须进 meta、CLI 必须报警、
纯扫描件 + 引擎缺失时的报错要指向缺引擎，而不是误导成「图像质量」问题。

复现（问题单）：ocr.py 引擎缺失与单页异常都 return "",0 无日志；仅成功页进
ocr_pages；只要一页出过字就不触发 RuntimeError；CLI 不查 ocr_ok/ocr_engine。
"""

import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import ocr  # noqa: E402

fitz = pytest.importorskip("fitz")

# 文字层用 ASCII，避免 fitz 内置字体写不了中文
PAGE1_TEXT = ("Submission guide: body text must be Songti 10.5pt, "
              "margins 25mm on all sides, double columns. ")


@pytest.fixture
def mixed_pdf(tmp_path):
    """第 1 页有文字层，第 2 页空白（有效字符 < 20 → 判为需 OCR）。"""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), PAGE1_TEXT)
    doc.new_page()
    f = tmp_path / "mixed.pdf"
    doc.save(str(f))
    doc.close()
    return f


@pytest.fixture
def pure_scan_pdf(tmp_path):
    """整本没有文字层（纯扫描件的样子）。"""
    doc = fitz.open()
    doc.new_page()
    f = tmp_path / "scan.pdf"
    doc.save(str(f))
    doc.close()
    return f


@pytest.fixture
def no_engine(monkeypatch):
    monkeypatch.setattr(ocr, "engine", lambda: None)


@pytest.fixture
def boom_engine(monkeypatch):
    class Boom:
        def __call__(self, img):
            raise RuntimeError("识别崩了")

    monkeypatch.setattr(ocr, "engine", lambda: Boom())


# ---------------------------------------------------------------- meta 留痕

def test_engine_missing_page_recorded(mixed_pdf, no_engine):
    """引擎缺失时该页必须进 meta.ocr_failed_pages，并带原因 engine_missing。"""
    text, meta = ocr.extract_pdf_text(str(mixed_pdf))
    assert meta["ocr_failed_pages"] == [{"page": 2, "why": "engine_missing"}]
    assert meta["ocr_pages"] == [], "失败页不得混进成功页清单"
    assert "Submission guide" in text  # 第 1 页文字层不受影响


def test_engine_error_page_recorded(mixed_pdf, boom_engine):
    """单页识别抛错必须带原因 page_error 进 meta，且不拖垮其余页。"""
    text, meta = ocr.extract_pdf_text(str(mixed_pdf))
    assert meta["ocr_failed_pages"] == [{"page": 2, "why": "page_error"}]
    assert "Submission guide" in text


def test_failure_warnings_name_pages_and_package(mixed_pdf, no_engine):
    _, meta = ocr.extract_pdf_text(str(mixed_pdf))
    msgs = ocr.failure_warnings(meta)
    assert any("RapidOCR" in m and "rapidocr-onnxruntime" in m for m in msgs)
    assert any("2" in m for m in msgs), "警告必须点名失败页号"


def test_no_failure_no_warning(mixed_pdf):
    """正常路径（空白页只是没内容、不是失败）不得产生假警告。"""
    text, meta = ocr.extract_pdf_text(str(mixed_pdf))
    assert meta["ocr_failed_pages"] == []
    assert ocr.failure_warnings(meta) == []


# ---------------------------------------------------------------- 报错改口

def test_pure_scan_without_engine_message(pure_scan_pdf, no_engine):
    """纯扫描件 + 引擎缺失：报错必须指向未安装 RapidOCR，而不是图像质量。"""
    with pytest.raises(RuntimeError) as ei:
        ocr.extract_pdf_text(str(pure_scan_pdf))
    msg = str(ei.value)
    assert "RapidOCR" in msg and "rapidocr-onnxruntime" in msg
    assert "图像" not in msg, "引擎缺失不得再被误导为「图像质量」问题"


def test_pure_scan_with_engine_error_names_pages(pure_scan_pdf, boom_engine):
    """引擎在但页页失败：报错要带上失败页号，让人知道丢的是哪些页。"""
    with pytest.raises(RuntimeError) as ei:
        ocr.extract_pdf_text(str(pure_scan_pdf))
    assert "识别失败页" in str(ei.value) and "[1]" in str(ei.value)


# ---------------------------------------------------------------- CLI 警告

def test_cli_warns_on_stderr(mixed_pdf, no_engine, monkeypatch, capsys, tmp_path):
    """ocr.py CLI：有失败页时必须在 stderr 打显式警告，且正文输出不被污染。"""
    out = tmp_path / "out.txt"
    monkeypatch.setattr(sys, "argv", ["ocr.py", str(mixed_pdf), "-o", str(out)])
    assert ocr.main() == 0
    captured = capsys.readouterr()
    assert "RapidOCR" in captured.err, "失败页 >0 时 CLI 没有 stderr 警告"
    assert "RapidOCR" not in captured.out


def test_cli_pure_scan_no_engine_exits_nonzero(pure_scan_pdf, no_engine,
                                               monkeypatch, capsys):
    """纯扫描件 + 无引擎：退出码非零，报错点名 RapidOCR。"""
    monkeypatch.setattr(sys, "argv", ["ocr.py", str(pure_scan_pdf)])
    assert ocr.main() == 1
    captured = capsys.readouterr()
    assert "RapidOCR" in captured.out + captured.err


# ---------------------------------------------------------------- 上游贯通

def test_draft_carries_failed_pages():
    """extract_journal.build 必须把失败页写进草稿的 extraction 留痕。"""
    import extract_journal as ej
    meta = {"ocr_failed_pages": [{"page": 3, "why": "page_error"}]}
    rules, _ = ej.build({}, {}, "无关正文", jid="demof", name_zh="演示期刊丁", meta=meta)
    assert rules["verification"]["extraction"]["ocr_failed_pages"] == \
        meta["ocr_failed_pages"]
    assert ocr.failure_warnings(rules["verification"]["extraction"])


def test_extract_journal_cli_warns(monkeypatch, tmp_path, capsys):
    """extract_journal.py CLI：OCR 失败页与未证实项都要在 stderr 喊出来。"""
    import extract_journal as ej
    rules = {"verification": {"extraction": {
        "ocr_failed_pages": [{"page": 1, "why": "engine_missing"}]}}}
    report = {"stats": {"extracted": 0, "inherited": 0, "dropped": 0, "unverified": 2},
              "errors": [], "warnings": [], "ok_to_publish": True}
    monkeypatch.setattr(ej, "extract", lambda *a, **k: (rules, report, "文本"))
    monkeypatch.setattr(sys, "argv",
                        ["extract_journal.py", "x.pdf", "--id", "democli",
                         "-o", str(tmp_path / "draft.json")])
    assert ej.main() == 0
    err = capsys.readouterr().err
    assert "RapidOCR" in err, "OCR 失败页没有 CLI 警告"
    assert "未证实" in err and "2" in err, "未证实项没有 CLI 警告"


def test_subprocess_cli_pure_scan_message(tmp_path):
    """真子进程跑 tools/ocr.py（本机若装了引擎则跳过——引擎缺失是前提）。"""
    try:
        import rapidocr_onnxruntime  # noqa: F401
        pytest.skip("本机装了 RapidOCR，引擎缺失场景无法在子进程复现")
    except ImportError:
        pass
    doc = fitz.open()
    doc.new_page()
    f = tmp_path / "scan.pdf"
    doc.save(str(f))
    doc.close()
    r = subprocess.run([sys.executable, str(ROOT / "tools" / "ocr.py"), str(f)],
                       capture_output=True, text=True)
    assert r.returncode != 0, "纯扫描件取不到文必须以非零退出"
    assert "RapidOCR" in r.stdout + r.stderr
