"""一条命令重建某本期刊的全部产物。

    python tools/build_journal.py rules/journal/jryj.json

依次完成：
1. 由规则集渲染版式样板 docx
2. 转 PDF（供预览与打印核对）
3. 渲染规格页 HTML，并把样板 PDF 逐页内嵌进去

三步共用同一份 JSON，因此样板、PDF、规格页不可能互相不一致。
"""

import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = sys.executable

SOFFICE_CANDIDATES = [
    shutil.which("soffice"),
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        if c and Path(c).exists():
            return c
    return None


def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        print("失败:", " ".join(map(str, cmd)))
        print(r.stdout[-800:])
        print(r.stderr[-800:])
        raise SystemExit(1)
    return r


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    ruleset = Path(sys.argv[1]).resolve()
    rules = json.loads(ruleset.read_text(encoding="utf-8"))
    jid = rules["journal"]["id"]
    outdir = ruleset.parent

    sample = outdir / f"{jid}.sample.docx"
    pdf = outdir / f"{jid}.sample.pdf"
    spec = outdir / f"{jid}.spec.html"

    run([PY, str(ROOT / "tools/render_sample.py"), str(ruleset), "-o", str(sample)])
    print("  样板 →", sample.name)

    soffice = find_soffice()
    if soffice:
        run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(sample)])
        print("  PDF  →", pdf.name)
    else:
        print("  跳过 PDF：未找到 LibreOffice，规格页将不含成品预览")

    run([PY, str(ROOT / "tools/render_spec.py"), str(ruleset), "-o", str(spec),
         "--sample-pdf", str(pdf) if pdf.exists() else ""])
    print("  规格 →", spec.name)

    size = spec.stat().st_size / 1024
    print(f"完成：{jid}（规格页 {size:.0f} KB，含 {len(rules.get('assumptions', []))} 条推定项）")


if __name__ == "__main__":
    main()
