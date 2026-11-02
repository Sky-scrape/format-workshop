"""F26 / F27 回归钉子：离线工具链坏输入隔离 + build_journal 不静默吞 soffice 失败。

- F26  一份手写坏 JSON 让体检/总览整体崩溃且不指明坏文件；字符串型字号被静默
  跳过；compliance_report 缺键即原始 traceback。
  → 四个工具逐文件 try/except（点名「文件 X：原因」、末尾汇总失败清单、退出码
    非零），字符串字号输出 warning，缺键给明确说明。
- F27  soffice 调用无 timeout；退出码 0 但 PDF 未落盘时静默降级仍报「完成」；
  旧 PDF 残留时静默内嵌过期 PDF 进规格页。
  → 带超时 + 转换前删旧 PDF + 转换后 PDF 不存在则显式失败（非零退出）。
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))


def good_ruleset(jid="goodone", name="演示刊"):
    """能通过各工具的最小规则集，不含任何真实期刊数据。"""
    return {
        "journal": {"id": jid, "name_zh": name, "language": "zh"},
        "typography": {"body": {"font_cn": "宋体", "size_pt": 10.5}},
        "verification": {"status": "verified", "source_url": "https://example.com"},
        "blocks": [{"type": "title", "required": True}],
    }


def good_ir():
    return {"source": "a.docx", "blocks": [{"type": "body", "text": "正文内容。"}]}


# ---------------------------------------------------------------- F26 audit_rulesets

def test_audit_isolates_bad_file(tmp_path):
    """非 JSON 文件点名报错，好文件照常体检——不再整体崩溃。"""
    import audit_rulesets as ar
    (tmp_path / "good.json").write_text(
        json.dumps(good_ruleset(), ensure_ascii=False), encoding="utf-8")
    (tmp_path / "bad.json").write_text("{这不是JSON", encoding="utf-8")
    (tmp_path / "broken.json").write_text(json.dumps({"foo": 1}), encoding="utf-8")

    rows, failures = ar.audit_all(tmp_path)
    assert [r["id"] for r in rows] == ["goodone"], "坏文件拖垮了好文件的体检"
    names = {f[0] for f in failures}
    assert {"bad.json", "broken.json"} <= names, f"失败清单没点名坏文件：{failures}"
    assert all(reason for _, reason in failures), "失败清单必须带原因"


def test_audit_cli_exit_code_and_report(tmp_path, monkeypatch, capsys):
    """CLI：末尾汇总失败清单，退出码非零，好文件行照常打印。"""
    import audit_rulesets as ar
    (tmp_path / "good.json").write_text(
        json.dumps(good_ruleset(), ensure_ascii=False), encoding="utf-8")
    (tmp_path / "bad.json").write_text("{oops", encoding="utf-8")
    monkeypatch.setattr(ar, "JOURNAL_DIR", tmp_path)
    monkeypatch.setattr(sys, "argv", ["audit_rulesets.py"])

    assert ar.main() == 1
    out = capsys.readouterr().out
    assert "goodone" in out
    assert "bad.json" in out


def test_audit_warns_on_string_size(tmp_path):
    """命中 SIZE_KEYS 但值是字符串：不参与刻度检查，但必须点名报警。"""
    import audit_rulesets as ar
    rs = good_ruleset()
    rs["typography"]["body"]["size_pt"] = "五号"
    (tmp_path / "s.json").write_text(json.dumps(rs, ensure_ascii=False), encoding="utf-8")

    rows, failures = ar.audit_all(tmp_path)
    assert not failures
    assert rows[0]["string_sizes"] == [("typography.body.size_pt", "五号")]


# ---------------------------------------------------------------- F26 normalize_rulesets

def test_normalize_isolates_bad_file(tmp_path, capsys):
    """坏文件点名报错、退出码非零；好文件照常归整落盘。"""
    import normalize_rulesets as nr
    good = good_ruleset(jid="goodtwo")
    good["typography"]["body"]["size_pt"] = 10.37  # 不在半磅刻度，应被归整
    (tmp_path / "good.json").write_text(json.dumps(good, ensure_ascii=False), encoding="utf-8")
    (tmp_path / "bad.json").write_text("[[[", encoding="utf-8")

    assert nr.main(tmp_path) == 1
    assert "bad.json" in capsys.readouterr().out

    fixed = json.loads((tmp_path / "good.json").read_text(encoding="utf-8"))
    assert fixed["typography"]["body"]["size_pt"] == nr.snap(10.37), \
        "好文件没有被继续归整"


def test_normalize_warns_on_string_size(tmp_path):
    import normalize_rulesets as nr
    rs = good_ruleset(jid="goodthree")
    rs["typography"]["body"]["size_pt"] = "五号"
    (tmp_path / "s.json").write_text(json.dumps(rs, ensure_ascii=False), encoding="utf-8")

    nr.TYPE_WARNINGS.clear()
    try:
        assert nr.main(tmp_path) == 0  # 字符串字号是警告，不是失败
        warns = list(nr.TYPE_WARNINGS)
    finally:
        nr.TYPE_WARNINGS.clear()
    assert any(p == "typography.body.size_pt" and v == "五号" for p, v in warns), \
        "字符串型字号被静默跳过了"


# ---------------------------------------------------------------- F26 build_index

def test_build_index_isolates_bad_file(tmp_path, capsys):
    """坏 JSON / 缺 journal 区块都点名跳过，好文件的卡片照常进总览页。"""
    import build_index as bi
    (tmp_path / "good.json").write_text(
        json.dumps(good_ruleset(), ensure_ascii=False), encoding="utf-8")
    (tmp_path / "bad.json").write_text("{}", encoding="utf-8")  # 缺 journal → KeyError
    (tmp_path / "oops.json").write_text("not json", encoding="utf-8")

    failures = bi.build(tmp_path, tmp_path / "index.html")
    assert {"bad.json", "oops.json"} <= {f[0] for f in failures}

    page = (tmp_path / "index.html").read_text(encoding="utf-8")
    assert "演示刊" in page, "好文件没有照常出结果"
    assert "bad.json" in capsys.readouterr().err


def test_build_index_exit_code(tmp_path, monkeypatch):
    import build_index as bi
    (tmp_path / "bad.json").write_text("{}", encoding="utf-8")
    monkeypatch.setattr(bi, "JDIR", tmp_path)
    monkeypatch.setattr(sys, "argv", ["build_index.py", "-o", str(tmp_path / "i.html")])
    assert bi.main() == 1


# ---------------------------------------------------------------- F26 compliance_report

def _run_compliance(ir_path, rules_path, out_path):
    return subprocess.run(
        [sys.executable, str(TOOLS / "compliance_report.py"),
         str(ir_path), str(rules_path), "-o", str(out_path)],
        capture_output=True, text=True)


def test_compliance_bad_ir_named_not_traceback(tmp_path):
    """坏 IR：退出码非零、点名文件与原因，不冒原始 traceback。"""
    ir = tmp_path / "ir.json"
    ir.write_text("{bad json", encoding="utf-8")
    rules = tmp_path / "rules.json"
    rules.write_text(json.dumps(good_ruleset(), ensure_ascii=False), encoding="utf-8")
    r = _run_compliance(ir, rules, tmp_path / "r.html")
    assert r.returncode == 1
    blob = r.stdout + r.stderr
    assert "ir.json" in blob and "失败" in blob
    assert "Traceback" not in blob


def test_compliance_missing_key_explained(tmp_path):
    """规则集缺 citation.system：给「缺哪个键」的人话说明，不甩 KeyError。"""
    rules = tmp_path / "rules.json"
    rules.write_text(json.dumps(good_ruleset(), ensure_ascii=False), encoding="utf-8")
    ir = tmp_path / "ir.json"
    ir.write_text(json.dumps(good_ir(), ensure_ascii=False), encoding="utf-8")
    r = _run_compliance(ir, rules, tmp_path / "r.html")
    assert r.returncode == 1
    blob = r.stdout + r.stderr
    assert "citation" in blob, f"缺键说明没有点名 citation：{blob[-300:]}"
    assert "Traceback" not in blob
    assert not (tmp_path / "r.html").exists(), "失败时不应留下半成品报告"


def test_compliance_still_works_on_good_input(tmp_path):
    """隔离加固不得碰坏正常路径：好输入照常出报告。"""
    rules = good_ruleset()
    rules["citation"] = {"system": "numbered"}
    rp = tmp_path / "rules.json"
    rp.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    ip = tmp_path / "ir.json"
    ip.write_text(json.dumps(good_ir(), ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "r.html"
    r = _run_compliance(ip, rp, out)
    assert r.returncode == 0, r.stderr[-500:]
    assert out.exists() and "投稿体检" in out.read_text(encoding="utf-8")


# ---------------------------------------------------------------- F27 build_journal

@pytest.fixture
def bj():
    import build_journal as bjmod
    return bjmod


def test_run_timeout_kills_and_exits(bj):
    """soffice 挂死必须在超时后被终止并让 build 以非零退出，而不是无限阻塞。"""
    with pytest.raises(SystemExit) as ei:
        bj.run([sys.executable, "-c", "import time; time.sleep(30)"], timeout=2)
    assert ei.value.code == 1


def test_convert_fails_loudly_when_pdf_missing(tmp_path, bj, monkeypatch):
    """soffice 退出码 0 但 PDF 没落盘 → 显式失败；旧 PDF 必须先被删掉。"""
    sample = tmp_path / "x.sample.docx"
    sample.write_bytes(b"docx")
    pdf = tmp_path / "x.sample.pdf"
    pdf.write_bytes(b"OLD PDF")  # 上一轮残留的旧 PDF

    calls = []

    def fake_run(cmd, timeout=None, **kw):
        calls.append((list(cmd), timeout))  # 桩：退出码 0，但不产出 PDF

    monkeypatch.setattr(bj, "run", fake_run)
    with pytest.raises(SystemExit) as ei:
        bj.convert_pdf("soffice-fake", sample, pdf)
    assert ei.value.code not in (None, 0), "PDF 缺失必须让 build 非零退出"
    assert not pdf.exists(), "旧 PDF 必须在转换前删掉，不能被静默内嵌进规格页"
    cmd, timeout = calls[0]
    assert timeout == bj.SOFFICE_TIMEOUT and timeout >= 1, "soffice 调用必须带 timeout"


def test_convert_keeps_fresh_pdf(tmp_path, bj, monkeypatch):
    """正常桩（真的产出 PDF）不误伤：转换成功后规格页照常拿得到预览。"""
    sample = tmp_path / "y.sample.docx"
    sample.write_bytes(b"docx")
    pdf = tmp_path / "y.sample.pdf"

    def fake_run(cmd, timeout=None, **kw):
        pdf.write_bytes(b"NEW PDF")

    monkeypatch.setattr(bj, "run", fake_run)
    bj.convert_pdf("soffice-fake", sample, pdf)  # 不应抛
    assert pdf.read_bytes() == b"NEW PDF"


def test_convert_timeout_wiring(tmp_path, bj, monkeypatch):
    """桩超时（TimeoutExpired）必须被 run 转成非零 SystemExit，而不是冒泡。"""
    sample = tmp_path / "z.sample.docx"
    sample.write_bytes(b"docx")
    pdf = tmp_path / "z.sample.pdf"

    def fake_subprocess_run(cmd, timeout=None, **kw):
        raise subprocess.TimeoutExpired(cmd, timeout)

    # 只替换 subprocess 层——真正的 run 包装器（超时→SystemExit）要真跑
    monkeypatch.setattr(bj.subprocess, "run", fake_subprocess_run)
    with pytest.raises(SystemExit) as ei:
        bj.convert_pdf("soffice-fake", sample, pdf)
    assert ei.value.code == 1
