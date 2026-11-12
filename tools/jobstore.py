"""任务目录的读写原语：与 Flask 解耦，供 app.py 与测试共用。

这里只放「怎么读写」——目录布局、原子写、坏 meta 推断、Windows 短暂占用
的退避。任务锁、TTL 清扫与 HTTP 404 语义留在 app.py：它们要么依赖进程内
锁表，要么直接决定页面响应，搬进工具层反而制造隐式全局。

app.py 仍以同名符号转出这些函数（read_meta / write_json_atomic / ...），
既有调用与测试不需要改 import。
"""

import json
import os
import time
import uuid
from pathlib import Path


def read_json_safe(path):
    """读 JSON 文件，坏文件（截断 / 编码错 / 读失败）返回 None，绝不抛异常。

    faulthandler 都装上了，说明「进程被硬杀」是预期场景，半截 JSON 会留下。
    """
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def read_json(path, default=None):
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default


def replace_with_retry(src, dst):
    """原子替换；Windows 上目标被读者短暂打开时做有限退避，不无限等待。

    4 次退避共约 0.24s，最后一次直接抛出（FX1 实测 0.247s）。
    """
    delays = (0.02, 0.04, 0.08, 0.1, 0.2)
    for attempt in range(5):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(delays[attempt])


def write_json_atomic(path, data, replace=None):
    """把 JSON 完整写入同目录唯一临时名，再原子替换目标。

    临时名必须带随机后缀：固定的 meta.json.tmp / details.json.tmp 在两个
    写者并发时会互相覆盖，甚至一个 finally 把另一个尚未落位的临时文件删掉。
    同目录替换保证不跨卷；读方只会看到完整旧值或完整新值。

    replace 允许调用方注入自己的替换实现（app.py 的测试会 monkeypatch
    模块级 replace_with_retry 观察临时名，默认绑定在这里就看不见了）。
    """
    replace = replace or replace_with_retry
    p = Path(path)
    tmp = p.with_name(f".{p.name}.{uuid.uuid4().hex}.tmp")
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        replace(tmp, p)
    finally:
        tmp.unlink(missing_ok=True)


def write_meta(d, meta):
    """meta.json 原子写（F34），并支持同任务的并发写者。"""
    write_json_atomic(d / "meta.json", meta)


def read_meta(d):
    """读任务 meta；坏文件按「从现场文件推断最低限度状态」兜底。

    进程硬杀会留下截断的 JSON，一条坏 meta 不许把 history 和任务页整页
    拖成 500（F34）。推断只回答「这个任务走到了哪一步、还剩什么可用」，
    救不回的字段留空，由各页面按缺失兜底（fail_page 会把推断结果写回，
    相当于就地自愈，不再等 24 小时 TTL 整目录删除）。
    """
    d = Path(d)
    p = d / "meta.json"
    if p.exists():
        m = read_json_safe(p)
        if isinstance(m, dict):
            return m
    m = {}
    if (d / "error.json").is_file():
        m["stage"] = "failed"          # 最后一次落笔是失败页 → 停在失败态
    ir = read_json_safe(d / "manuscript.ir.json")
    if isinstance(ir, dict) and ir.get("source"):
        # 解析成功时 IR 里记录了用户上传的原始文件名，可以由此救回
        m["original"] = ir["source"]
        m["stem"] = Path(ir["source"]).stem or "稿件"
    delivered = [q for q in sorted(d.glob("manuscript.*.docx"))
                 if not q.name.endswith(".anon.docx")]
    if delivered:                      # 成品文件名 manuscript.<刊id>.docx 反推目标刊
        m["docx"] = delivered[0].name
        m["journal"] = delivered[0].name[len("manuscript."):-len(".docx")]
    return m
