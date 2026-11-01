"""AI 辅助层的自检：不需要真实模型，用打桩验证三条安全边界。

跑法：python tools/selftest_ai.py [样本.docx]

验证项
    1 采纳路径   合法提案被写进 IR，且文字一字未动
    2 白名单     AI 提出 table / 未知类型一律丢弃
    3 锁定      规则 high confidence 的块，AI 不得推翻
    4 降级       模型调用抛错时退回纯规则路径，IR 与关闭 AI 时逐块相同
    5 守卫       文本被篡改时整批回滚并报失败
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ai_assist  # noqa: E402
import parse_docx  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "训练论文" / "目标落空之后.docx"

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(f"  {'PASS' if cond else 'FAIL'}  {name}{'  ' + detail if detail and not cond else ''}")


def cfg():
    return {**ai_assist.DEFAULTS, "enabled": True, "base_url": "http://x/v1",
            "api_key": "sk-test", "model": "stub"}


def stub(proposals=None, raises=None):
    def _chat(_cfg, messages):
        if raises:
            raise raises
        return json.dumps({"labels": proposals}, ensure_ascii=False)
    return _chat


def baseline(docx):
    """纯规则路径的基线。"""
    old_avail = ai_assist.available
    ai_assist.available = lambda *_a, **_k: False
    try:
        return parse_docx.parse(docx, ai_cfg={"enabled": False})["blocks"]
    finally:
        ai_assist.available = old_avail


def main():
    docx = Path(sys.argv[1]) if len(sys.argv) > 1 else SAMPLE
    if not docx.exists():
        print(f"找不到样本：{docx}")
        return 2

    base = baseline(docx)
    bodies = [b for b in base if b["type"] == "body"]
    heads = [b for b in base if b["type"].startswith("heading")]
    print(f"样本 {docx.name}：{len(base)} 块，body {len(bodies)}，标题 {len(heads)}\n")

    # ---- 1/2/3 采纳、白名单、锁定
    print("[1-3] 采纳路径 / 白名单 / 锁定")
    target_body = bodies[0]["idx"]
    target_head = heads[0]["idx"]
    ai_assist.available = lambda *_a, **_k: True
    ai_assist._chat = stub([
        {"idx": target_body, "type": "heading1", "conf": "high", "why": "打桩：补一个漏判标题"},
        {"idx": bodies[1]["idx"], "type": "table", "conf": "high", "why": "打桩：非法类型"},
        {"idx": target_head, "type": "heading2", "conf": "high", "why": "打桩：推翻规则"},
        {"idx": 999999, "type": "body", "conf": "high", "why": "打桩：不存在的块"},
    ])
    ir = parse_docx.parse(docx, ai_cfg=cfg())
    after = {b["idx"]: b for b in ir["blocks"]}
    check("合法提案已采纳", after[target_body]["type"] == "heading1"
          and after[target_body].get("ai_from") == "body")
    check("非法类型被丢弃", after[bodies[1]["idx"]]["type"] == bodies[1]["type"])
    check("不存在的块被忽略", 999999 not in after)
    rejected = [i for i in ir["ai"]["items"] if i["status"] == "rejected"]
    check("规则高置信块不可推翻", len(rejected) == 1
          and after[target_head]["type"] == heads[0]["type"])
    check("文字一字未动", ai_assist.fingerprint(ir["blocks"]) == ai_assist.fingerprint(base))

    # ---- 4 降级
    print("\n[4] 模型调用失败 → 静默降级")
    ai_assist._chat = stub(raises=TimeoutError("模拟超时"))
    ir2 = parse_docx.parse(docx, ai_cfg=cfg())
    check("AI 失败不阻断解析", ir2["ai"]["used"] is False)
    check("降级结果与纯规则逐块相同", ir2["blocks"] == base)

    # ---- 5 守卫
    print("\n[5] 文本指纹守卫")
    ai_assist._chat = stub([
        {"idx": target_body, "type": "heading1", "conf": "high", "why": "打桩"},
    ])
    real_norm = parse_docx._normalize_ai_headings

    def rogue(blocks):
        real_norm(blocks)
        blocks[0]["text"] = blocks[0]["text"] + "（AI 偷偷加了字）"

    parse_docx._normalize_ai_headings = rogue
    try:
        ir3 = parse_docx.parse(docx, ai_cfg=cfg())
    finally:
        parse_docx._normalize_ai_headings = real_norm
    check("篡改文本触发回滚", ir3["ai"]["used"] is False and "指纹" in ir3["ai"]["reason"])
    check("回滚后与纯规则逐块相同", ir3["blocks"] == base)

    print(f"\n{len(PASS)} 通过，{len(FAIL)} 失败")
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
