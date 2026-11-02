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

# soffice 偶发挂死会无限期阻塞整个构建；给个上限，超时杀进程报错
SOFFICE_TIMEOUT = 180

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


def run(cmd, timeout=None, **kw):
    try:
        # cmd 一律是参数列表（绝不拼 shell 字符串），shell=False 显式声明
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                           shell=False, **kw)
    except subprocess.TimeoutExpired:
        print("失败（超时）:", " ".join(map(str, cmd)),
              f"— 超过 {timeout} 秒未结束，已终止")
        raise SystemExit(1)
    if r.returncode != 0:
        print("失败:", " ".join(map(str, cmd)))
        print(r.stdout[-800:])
        print(r.stderr[-800:])
        raise SystemExit(1)
    return r


def remove_old_pdf(pdf):
    """清理旧 PDF。Windows 上旧件常被阅读器占用着删不掉，
    裸 PermissionError 的 traceback 不是给人看的——点名文件、给出出路、非零退出。"""
    try:
        pdf.unlink()
    except PermissionError:
        print(f"失败：旧 PDF {pdf} 被占用（通常是阅读器开着它），"
              "请关闭阅读器后重试。")
        raise SystemExit(1)


def convert_pdf(soffice, sample, pdf, timeout=SOFFICE_TIMEOUT):
    """soffice 转 PDF。三条防线：

    1. 转换前删掉旧 PDF——残留的旧件是上一轮样板的影，会被规格页静默内嵌
    2. 带超时调用——soffice 挂死不能无限期阻塞
    3. 退出码 0 但 PDF 没落盘 → 显式失败。漂移的样板比没有样板更糟，
       宁可让 build 以非零退出，也不带着过期预览假装「完成」。
    """
    if pdf.exists():
        remove_old_pdf(pdf)
    run([soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(pdf.parent), str(sample)], timeout=timeout)
    if not pdf.exists():
        raise SystemExit(
            f"失败：soffice 退出码为 0，但 {pdf.name} 没有生成——转换实际没有成功。"
            "规格页不会带着旧 PDF 假装完成，请检查 LibreOffice 是否正常。")


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
        convert_pdf(soffice, sample, pdf)
        print("  PDF  →", pdf.name)
    else:
        if pdf.exists():
            # 没有转换器也不能拿上一轮的 PDF 冒充本轮样板的预览
            remove_old_pdf(pdf)
            print("  已删除旧 PDF：未找到 LibreOffice，无法重新生成，规格页将不含成品预览")
        else:
            print("  跳过 PDF：未找到 LibreOffice，规格页将不含成品预览")

    run([PY, str(ROOT / "tools/render_spec.py"), str(ruleset), "-o", str(spec),
         "--sample-pdf", str(pdf) if pdf.exists() else ""])
    print("  规格 →", spec.name)

    size = spec.stat().st_size / 1024
    print(f"完成：{jid}（规格页 {size:.0f} KB，含 {len(rules.get('assumptions', []))} 条推定项）")


if __name__ == "__main__":
    main()
