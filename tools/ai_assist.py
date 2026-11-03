"""AI 辅助层：只给块打类型标签，不碰任何一个字符。

契约（DESIGN §14，改动前先读那一节）
--------------------------------------------------------------------
1. 入参是块列表，出参是 `{块序号: {type, conf, why}}`。**永不返回文本。**
   调用方拿到提案后自己决定要不要采纳，本模块不做任何落笔动作。
2. 任何失败（无配置、超时、鉴权错、返回不是合法 JSON、类型不在白名单）
   → 返回空字典，流水线静默退回纯规则路径。AI 不可用不得导致转换失败。
3. 提案必须经调用方的文本指纹守卫校验（`fingerprint()` 前后一致），
   不一致则整批提案作废——这是"AI 不可能改坏文字"的物理保证。

为什么是"标注"而不是"改写"
--------------------------------------------------------------------
§1.1 把"不改变任何一个字词"定为产品第一性原则。块类型（这段是标题还是
正文）属于**结构判断**，判断错了还有审计记录可回溯；文字属于**内容**，
改一个就再也说不清是作者写的还是机器写的。所以 AI 的收益（鲁棒性）
全部放在结构层，风险（不可信）一个都不进内容层。

配置
--------------------------------------------------------------------
读 `config/ai.json`，环境变量可覆盖单值：

    PAPER_AI_ENABLED=1
    PAPER_AI_BASE_URL=https://api.deepseek.com/v1
    PAPER_AI_API_KEY=sk-xxx
    PAPER_AI_MODEL=deepseek-chat

任何 OpenAI 兼容的 `/chat/completions` 端点都行（DeepSeek、通义、Kimi、
本地 Ollama 的 /v1）。默认 `enabled: false`——不配就不联网。
"""

import ipaddress
import json
import os
import time
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config" / "ai.json"
# 环境变量统一前缀：PAPER_AI_<SUFFIX>
_ENV_PREFIX = "PAPER_AI_"

DEFAULTS = {
    "enabled": False,
    "base_url": "",
    "api_key": "",
    "model": "",
    "timeout": 60,
    "batch_size": 60,
    "max_blocks": 300,
    # snippet: 只发首尾片段；full: 发全文（隐私风险，见 §14.3）
    "send_mode": "snippet",
    "snippet_chars": 60,
    "retries": 1,
}

# 允许 AI 判定的类型。table / figure / equation* 是结构性块，由 XML 直判，
# AI 不得染指——它看不到单元格，猜"这不是表格"只会把表格改坏。
ALLOWED_TYPES = {
    "body", "front_matter",
    "heading1", "heading2", "heading3", "heading4",
    "abstract", "keywords", "classification",
    "english_title", "english_abstract", "english_keywords",
    "references_heading", "reference_item", "appendix_heading",
    "table_caption", "figure_caption", "table_note",
}

SYSTEM_PROMPT = """你是中文学术论文的结构标注器。你的唯一任务是判断「这一段属于哪一类」，绝对不改写、补写、翻译、删减任何文字。

可选类型（只能从中选一个）：
body 正文段落
heading1 / heading2 / heading3 / heading4 一级到四级标题
abstract 中文摘要 / keywords 中文关键词 / classification 分类号或文献标识码
english_title 英文标题 / english_abstract 英文摘要 / english_keywords 英文关键词
references_heading 参考文献表的标题行 / reference_item 参考文献条目
appendix_heading 附录标题
table_caption 表题 / figure_caption 图题 / table_note 表注（以「注：」「资料来源」「数据来源」开头）
front_matter 位于摘要之前的题录行：题名、副标题、作者、单位、通讯方式

判断要点：
- 编号体系要认全：「一、」「（一）」「1.」「（1）」是标题层级，「第一章」「1 引言」「1.1」「2.3.1」同样是标题
- 标题通常较短、无句末标点（。；，）、数字与英文占比低；正文段落通常较长且有句号
- 以「表1」开头但含句号、后接完整论证句的，是正文里的引用句，不是表题
- 参考文献区从 references_heading 之后开始，条目几乎都含年份；遇到新的一级标题或附录即结束
- 摘要之前的作者、单位行归 front_matter；摘要之后一律不是 front_matter
- 拿不准时给 conf="low"，不要硬猜

只输出 JSON，形如：
{"labels":[{"idx":12,"type":"heading1","conf":"high","why":"以「一、」开头的短句"}]}
不要输出解释、不要用 markdown 代码块。"""


# ---------------------------------------------------------------- 配置

def load_config():
    cfg = dict(DEFAULTS)
    if CONFIG_PATH.exists():
        try:
            raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                cfg.update({k: v for k, v in raw.items() if k in DEFAULTS})
        except Exception:
            # 配置坏了就当没配，绝不让格式转换因为读不到配置而失败
            pass
    env_map = {
        "enabled": "ENABLED",
        "base_url": "BASE_URL",
        "api_key": "API_KEY",
        "model": "MODEL",
    }
    for k, suffix in env_map.items():
        v = os.environ.get(_ENV_PREFIX + suffix)
        if v is None or v == "":
            continue
        cfg[k] = (v.strip().lower() not in ("0", "false", "no", "off")
                  if k == "enabled" else v.strip())
    try:
        cfg["timeout"] = max(5, int(cfg.get("timeout") or 60))
        cfg["batch_size"] = max(10, int(cfg.get("batch_size") or 60))
        cfg["max_blocks"] = max(10, int(cfg.get("max_blocks") or 300))
        cfg["snippet_chars"] = max(20, int(cfg.get("snippet_chars") or 60))
    except (TypeError, ValueError):
        cfg.update({k: DEFAULTS[k] for k in ("timeout", "batch_size", "max_blocks", "snippet_chars")})
    return cfg


def available(cfg=None):
    """能不能用。缺 base_url / api_key / model 任一项即视为不可用。"""
    cfg = cfg or load_config()
    return bool(cfg.get("enabled")
                and str(cfg.get("base_url") or "").strip()
                and str(cfg.get("api_key") or "").strip()
                and str(cfg.get("model") or "").strip())


def validate_config(cfg):
    """返回错误列表，空列表即配置可用。给网页表单用。"""
    errs = []
    if not str(cfg.get("base_url") or "").strip():
        errs.append("接口地址不能为空，例如 https://api.deepseek.com/v1")
    elif not str(cfg.get("base_url")).startswith(("http://", "https://")):
        errs.append("接口地址必须以 http:// 或 https:// 开头")
    if not str(cfg.get("model") or "").strip():
        errs.append("模型名不能为空")
    if cfg.get("enabled") and not str(cfg.get("api_key") or "").strip():
        errs.append("启用 AI 时必须填写 API Key")
    if str(cfg.get("send_mode") or "") not in ("snippet", "full"):
        errs.append("发送范围只能是 snippet 或 full")
    return errs


def save_config(values):
    """写回 config/ai.json。

    只接受 DEFAULTS 里的键——表单被伪造也写不进别的字段。
    直接改文件而不是改 load_config() 的结果：后者会混入环境变量，
    把环境变量的值反写进文件，等于悄悄把配置固化下来。
    """
    raw = {}
    if CONFIG_PATH.exists():
        try:
            cur = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(cur, dict):
                raw = {k: v for k, v in cur.items() if k in DEFAULTS}
        except Exception:
            raw = {}
    for k, v in (values or {}).items():
        if k in DEFAULTS and v is not None:
            raw[k] = v
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n",
                           encoding="utf-8")
    return raw


def test_connection(cfg=None, timeout=None):
    """真发一次最小请求。返回 (ok, 人话)。

    不抛异常：网页上点「测试」不该看到 traceback。
    """
    # 测的是「这套配置能不能用」，与当前开没开无关——
    # 沿用表单里 enabled=False 会让测试永远失败，用户会以为密钥填错了。
    cfg = dict(cfg or load_config())
    cfg["enabled"] = True
    errs = validate_config(cfg)
    if errs:
        return False, "；".join(errs)
    if not str(cfg.get("api_key") or "").strip():
        return False, "未填写 API Key"
    if timeout:
        cfg["timeout"] = timeout

    fake = [{"type": "body", "text": "一、引言", "idx": 0, "confidence": None}]
    try:
        prop = propose_types(fake, fake, cfg)
    except Exception as e:  # propose_types 内部已兜住，这里是最后一道
        return False, f"{type(e).__name__}: {e}"
    if not prop:
        return False, "连通但没拿到合法标签，检查模型是否支持 JSON 输出"
    # 不认 idx：有的服务会改写请求里的序号，能回出类型就说明链路是通的
    got = next(iter(prop.values()), {}).get("type")
    if got != "heading1":
        return True, f"连通正常，但样例「一、引言」被判为 {got}，标注可能不准"
    return True, "连通正常，样例判定正确"


# ---------------------------------------------------------------- 守卫

def fingerprint(blocks):
    """块序号与文本的指纹。AI 前后必须完全一致，否则提案整批作废。"""
    return [(b.get("idx"), b.get("text") or "") for b in blocks]


def text_unchanged(before, after):
    return before == after


# ---------------------------------------------------------------- 请求

def _validate_endpoint(cfg, url):
    """端点只允许 http(s)；环回地址放行，其余私网 IP 需显式 opt-in。

    127.0.0.1/localhost 是本机自身，指向它只可能打到本机服务——这是设计文档
    明确支持的零外发选项（DESIGN §14.3 的本地 Ollama）。其余私网段可能是
    内网第三方服务，默认不发；确有内网模型服务时在 config/ai.json 设
    "allow_lan_endpoint": true。
    """
    if not url.lower().startswith(("http://", "https://")):
        # urllib 对 file:// 等协议照发不误；端点只该是 http(s)。
        raise ValueError(f"AI 端点只允许 http/https：{url}")
    host = (urlparse(url).hostname or "").strip().lower()
    if not host:
        raise ValueError(f"AI 端点缺主机名：{url}")
    if host == "localhost" or host.endswith(".localhost"):
        return
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        return  # 常规域名，边界交给 http(s) scheme
    if ip.is_loopback:
        return
    if not ip.is_global and cfg.get("allow_lan_endpoint") is not True:
        raise ValueError(
            f"AI 端点指向私网地址 {host}。若模型服务确在内网，"
            '请在 config/ai.json 设置 "allow_lan_endpoint": true'
        )


def _post(cfg, payload, timeout):
    url = str(cfg["base_url"]).rstrip("/") + "/chat/completions"
    _validate_endpoint(cfg, url)
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST", headers={
        "Content-Type": "application/json",
        "Authorization": "Bearer " + str(cfg["api_key"]),
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _chat(cfg, messages, max_tokens=2048):
    """调一次模型。返回文本内容；失败抛异常由上层统一兜。

    max_tokens 默认 2048（标签批的输出很小，维持原值不变）；事实抽取这类
    要求逐键输出值+逐字引文的长应答由调用方传更大的值，否则密稿在 2048
    处被截成半截 JSON（F52）。
    """
    payload = {
        "model": cfg["model"],
        "messages": messages,
        "temperature": 0,
        "stream": False,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }
    try:
        return _post(cfg, payload, cfg["timeout"])["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        # 部分端点不认 response_format，去掉再试一次
        if e.code in (400, 404, 422):
            payload.pop("response_format", None)
            return _post(cfg, payload, cfg["timeout"])["choices"][0]["message"]["content"]
        raise


def _parse_labels(text):
    """从模型输出里抠出 JSON。容忍代码块围栏与前后多余文字。"""
    if not text:
        return {}
    s = text.strip()
    if s.startswith("```"):
        s = s.split("```")[1] if "```" in s[3:] else s.lstrip("`")
        s = s.split("\n", 1)[1] if "\n" in s else s
    lo, hi = s.find("{"), s.rfind("}")
    if lo < 0 or hi <= lo:
        return {}
    try:
        obj = json.loads(s[lo:hi + 1])
    except json.JSONDecodeError:
        return {}
    labels = obj.get("labels") if isinstance(obj, dict) else None
    if not isinstance(labels, list):
        return {}
    out = {}
    for it in labels:
        if not isinstance(it, dict):
            continue
        try:
            idx = int(it.get("idx"))
        except (TypeError, ValueError):
            continue
        t = str(it.get("type") or "").strip()
        if t not in ALLOWED_TYPES:
            continue
        out[idx] = {"type": t,
                    "conf": str(it.get("conf") or "low").strip().lower(),
                    "why": str(it.get("why") or "")[:60]}
    return out


def _snippet(text, n):
    if len(text) <= n * 1.6:
        return text
    return text[:n] + "…" + text[-max(12, n // 3):]


def _candidate_payload(blocks, cands, cfg):
    """只发候选块 + 极短的邻接上下文。标题判定靠前后文，正文不需要。"""
    by_idx = {b.get("idx"): i for i, b in enumerate(blocks)}
    items = []
    for b in cands:
        i = by_idx.get(b.get("idx"))
        prev_t = blocks[i - 1].get("text", "") if i and i > 0 else ""
        next_t = blocks[i + 1].get("text", "") if i is not None and i + 1 < len(blocks) else ""
        item = {
            "idx": b.get("idx"),
            "cur": b.get("type"),
            "len": len(b.get("text") or ""),
            "text": (_snippet(b.get("text") or "", cfg["snippet_chars"])
                     if cfg["send_mode"] != "full" else (b.get("text") or "")),
        }
        if b.get("has_image"):
            item["img"] = True
        if prev_t:
            item["prev"] = _snippet(prev_t, 30)
        if next_t:
            item["next"] = _snippet(next_t, 30)
        items.append(item)
    return items


# ---------------------------------------------------------------- 对外

def propose_types(blocks, cands, cfg=None, log=None, trace=None):
    """让模型给候选块提类型。返回 {idx: {type, conf, why}}。

    任何异常都收敛成空字典——AI 挂掉不能让格式转换挂掉。

    trace（可选）：传入一个 dict，函数往里写本次截断情况——
    {"truncated": bool, "dropped": int, "kept": int}。候选超过 max_blocks
    时并不是「跳过 AI」，而是只标前 max_blocks 块、其余留在规则路径
    （F50：旧日志说「AI 跳过」，与实际相反，且被截的块无处留痕）。
    """
    # 调用方可能只传部分键（网页表单就没带 batch_size），这里统一补齐，
    # 否则 cfg["max_blocks"] 会 KeyError——库函数不该假设调用方给全了配置
    cfg = {**DEFAULTS, **(cfg or load_config())}
    if not cands or not available(cfg):
        return {}
    if len(cands) > cfg["max_blocks"]:
        # 候选太多说明解析本身有问题，此时批量标注的成本与收益不成比例。
        # 如实说：是截断续跑，不是跳过；多出来的块没有提案，靠规则判定。
        if log:
            log(f"AI：候选块 {len(cands)} 个超过上限 {cfg['max_blocks']}，"
                f"仅标注前 {cfg['max_blocks']} 块，其余按规则判定")
        if isinstance(trace, dict):
            trace["truncated"] = True
            trace["dropped"] = len(cands) - cfg["max_blocks"]
            trace["kept"] = cfg["max_blocks"]
        cands = cands[:cfg["max_blocks"]]
    elif isinstance(trace, dict):
        trace["truncated"] = False
        trace["dropped"] = 0
        trace["kept"] = len(cands)

    out = {}
    size = cfg["batch_size"]
    for start in range(0, len(cands), size):
        chunk = cands[start:start + size]
        items = _candidate_payload(blocks, chunk, cfg)
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(
                {"blocks": items}, ensure_ascii=False, separators=(",", ":"))},
        ]
        for attempt in range(int(cfg.get("retries", 1)) + 1):
            try:
                out.update(_parse_labels(_chat(cfg, messages)))
                break
            except Exception as e:
                if attempt >= int(cfg.get("retries", 1)):
                    if log:
                        log(f"AI 调用失败，退回纯规则路径：{type(e).__name__}: {e}")
                    return out
                time.sleep(1.5)
    return out


# ---------------------------------------------------------------- 自检入口

def _mask(s):
    s = str(s or "")
    return (s[:4] + "…" + s[-4:]) if len(s) > 10 else ("已设置" if s else "未设置")


def main():
    """配好之后先跑这个，别等转换跑到一半才发现鉴权错。

        python tools/ai_assist.py --test
        python tools/ai_assist.py --show-config
    """
    import argparse
    ap = argparse.ArgumentParser(description="AI 辅助层连通性与配置检查")
    ap.add_argument("--test", action="store_true", help="发一次最小请求验证连通")
    ap.add_argument("--show-config", action="store_true", help="显示生效配置（key 打码）")
    args = ap.parse_args()

    cfg = load_config()
    if args.show_config or not args.test:
        print(f"配置文件    {CONFIG_PATH}（{'存在' if CONFIG_PATH.exists() else '不存在'}）")
        print(f"enabled     {cfg['enabled']}")
        print(f"base_url    {cfg['base_url'] or '未设置'}")
        print(f"model       {cfg['model'] or '未设置'}")
        print(f"api_key     {_mask(cfg['api_key'])}")
        print(f"send_mode   {cfg['send_mode']}（snippet 只发片段，full 发全文）")
        print(f"可用        {'是' if available(cfg) else '否——缺 enabled/base_url/model/api_key 之一'}")

    if not args.test:
        return 0
    print("\n发起测试请求…")
    ok, msg = test_connection(cfg)
    print(msg)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
