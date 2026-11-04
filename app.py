"""论文格式转换 · 本地 Web 应用

两段式流水线（对应 DESIGN §10.2 的执行顺序）：

    POST /convert            上传 → 解析 → 投稿体检            → 302 /job/<id>/details
    GET  /job/<id>/details   个人信息与缺失要件补录（预填已有值）
    POST /job/<id>/details   注入补录内容 → 渲染 → 审计          → 302 /result/<id>
    GET  /result/<id>        成品、报告、换刊重排、差异表入口

其余路由：
    GET  /journals                      期刊规则库
    GET  /query                         论文格式查询（单刊深查 / 多刊对比）
    GET  /history                       本机最近的转换记录
    GET  /job/<id>/diff?to=<jid>        两本刊的规则差异
    POST /job/<id>/retry                同一稿件换刊重排
    GET  /job/<id>/download/<kind>      带可读文件名的下载
    GET  /journal/<jid>/spec            期刊格式规格页
    GET  /journal/<jid>/sample(.pdf)    期刊版式样板（docx 下载 / PDF 内联预览）
"""

import argparse
import importlib.util
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid
import webbrowser
from functools import wraps
from logging.handlers import RotatingFileHandler
from pathlib import Path
from urllib.parse import quote

from flask import (Flask, abort, jsonify, redirect, render_template, request,
                   send_file, send_from_directory, url_for)

# pythonw.exe 下没有控制台，sys.stdout/stderr 是 None，print() 会直接抛错；
# 后台启动正是走 pythonw，所以进来先兜住这两个流：换成一个不打开任何文件
# 的空设备写入器（等价于写到 NUL，但没有文件句柄）。
class _DevNull:
    def write(self, s):
        return len(s)

    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = _DevNull()
if sys.stderr is None:
    sys.stderr = _DevNull()

ROOT = Path(__file__).resolve().parent
PY = sys.executable
RULES_DIR = ROOT / "rules" / "journal"
DRAFT_DIR = RULES_DIR / "_drafts"
WORK = ROOT / "outputs" / "web"
JOB_RE = re.compile(r"^[0-9a-f]{12}$")
JID_RE = re.compile(r"^[a-z0-9_]{2,24}$")
JOB_TTL = 24 * 3600
STEP_TIMEOUT = 300

sys.path.insert(0, str(ROOT / "tools"))
import ai_assist        # noqa: E402  AI 配置与调用
import docx_meta        # noqa: E402  DOCX 元数据隐私闸门
import extract_journal  # noqa: E402  稿约 PDF → 规则集草稿
import ocr              # noqa: E402  PDF 取文（文字层优先，缺则 OCR）
import rules_ref        # noqa: E402  审计引用 → 规则集字段解析

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
app.config["TEMPLATES_AUTO_RELOAD"] = True

# 服务只绑 127.0.0.1，但绑定挡不住两件事：用户浏览器里第三方页面对本机的跨站
# 表单 POST（简单请求无预检），以及 DNS rebinding 把攻击者域名解析到回环地址。
# 这两条都在进入任何路由之前拦下（F07）。
LOCAL_HOSTS = {"127.0.0.1", "localhost", "::1"}
MUTATING_METHODS = {"POST", "PUT", "PATCH", "DELETE"}


def _is_local_host(host_header):
    """取 Host 头里的主机名（去端口、去 IPv6 方括号）判断是否本机。"""
    h = (host_header or "").strip().lower()
    if h.startswith("["):
        h = h[1:].split("]", 1)[0]
    elif ":" in h:
        h = h.rsplit(":", 1)[0]
    return h in LOCAL_HOSTS


@app.after_request
def _add_security_headers(resp):
    """本地服务也补基本浏览器防线；不限制内联脚本，报告预览 iframe 仍可用。"""
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    resp.headers.setdefault("Content-Security-Policy",
                            "frame-ancestors 'self'; base-uri 'self'; object-src 'none'")
    # 稿件、审计与设置都含本机私密内容；静态资源不属于敏感面，仍允许缓存。
    if not request.path.startswith("/static/"):
        resp.headers.setdefault("Cache-Control", "no-store")
    return resp


@app.before_request
def _reject_remote_and_cross_site():
    """只接受本机自发、本站发起的请求。

    ① Host 必须是回环地址：Host 是别的域名只可能是 rebinding，一律 403。
    ② 写操作（POST 等）带了 Sec-Fetch-Site 且不是 same-origin 就 403——
       现代浏览器跨站表单必带 cross-site；curl / 本机脚本不带这个头，不受影响。
       GET 不拦：地址栏、书签打开页面时该头是 none，拦了就没法正常进入。
    """
    if not _is_local_host(request.host):
        app.logger.warning("拒绝非本机 Host 的请求：%s %s", request.method, request.path)
        return "拒绝访问：本服务只在本机（127.0.0.1）开放。", 403
    if request.method in MUTATING_METHODS:
        site = (request.headers.get("Sec-Fetch-Site") or "").strip().lower()
        if site and site != "same-origin":
            app.logger.warning("拒绝跨站写请求：%s %s（Sec-Fetch-Site: %s）",
                               request.method, request.path, site)
            return "拒绝访问：不接受来自其他网站的提交。", 403
    return None

AUTH_LABEL = {"template": "官方模板", "guide": "官方规定", "published": "实物实测"}
AUTH_CLASS = {"template": "a-tpl", "guide": "a-guide", "published": "a-pub"}
STATUS_LABEL = {"verified": "已核验", "partial": "部分核验", "unverified": "未核验"}
REPORT_FILES = {"triage": "triage.html", "audit": "audit.html"}

# 可由用户补录、并真的写进正文的要件：字段键 → (块类型, 语言, 标签, 输入提示)
DETAIL_FIELDS = [
    ("affiliation_zh", "affiliations", "zh", "作者单位",
     "多位作者用分号分隔，按署名顺序"),
    ("classification", "classification", "zh", "分类号",
     "按该刊要求填写，例如 JEL: E62; O17; D22"),
    ("english_title", "english_title", "en", "英文标题", ""),
    ("english_authors", "english_authors", "en", "英文作者",
     "按署名顺序，通常用拼音姓大写或全拼"),
    ("english_affiliation", "english_affiliations", "en", "英文单位", ""),
    ("english_abstract", "english_abstract", "en", "英文摘要",
     "请粘贴你已确定的译文；工具不会替你翻译或改写"),
    ("english_keywords", "english_keywords", "en", "英文关键词",
     "分号分隔，与该刊英文摘要配套"),
    ("english_classification", "classification", "en", "英文分类号",
     "个别刊要求与中文分类号并列的英文分类号，例如 JEL: E62; O17; D22"),
]

FOOTER_FIELDS = [
    ("funding", "基金项目", "写在首页脚注，例：国家自然科学基金青年项目（72003012）"),
    ("author_bio", "作者简介", "写在首页脚注，例：作者姓名，男，经济学博士，研究方向……"),
    ("corresponding", "通讯作者", "如需在首页注明，例：作者姓名，邮箱 xxx@example.com"),
]


# ---------------------------------------------------------------- 数据加载

def rules_of(jid):
    if not JID_RE.match(jid or ""):      # 不合法的 id 不参与路径拼接（F47，同 draft_dir）
        return None
    p = RULES_DIR / f"{jid}.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


try:                                  # 有 pypinyin 就按真实拼音排（多音字也对，如「会计」= kuai）
    from pypinyin import lazy_pinyin as _lazy_pinyin
except Exception:                     # 没有就退回 GB2312 近似序，不让功能因此挂掉
    _lazy_pinyin = None


def _pinyin_key(text):
    if _lazy_pinyin is not None:
        # 用空格连接各个音节：这样「jin rong」排在「jing ji」前面，
        # 直接拼字符串会把 jing 的 g 拿去和 r 比较，把「经济」排到「金融」前面
        return " ".join(_lazy_pinyin(text)).lower()
    return _gb_sort_key(text)


def _gb_sort_key(text):
    """中文按拼音排序：GB2312 一级汉字本身就是按拼音排的，用它当排序键。

    不引入第三方拼音库；二级汉字（生僻字）在 GB2312 里按部首排，
    对刊名这种常用字场景够用，取不到的字统一排到后面。
    """
    buf = bytearray()
    for ch in text:
        try:
            buf += ch.encode("gb2312")
        except UnicodeEncodeError:
            buf += b"\xff\xff"
    return bytes(buf)


def _load_journals_uncached():
    """读取全部期刊规则集。单本字段缺失只让这本刊消失，不让整站 500。"""
    items = []
    for f in sorted(RULES_DIR.glob("*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            body = d["typography"]["body"]
            ls = body.get("line_spacing", {}) or {}
            line_pt = ls.get("value_pt") or ls.get("equivalent_exact_pt")
            # 倍数行距（如「1.2 倍」）如实展示倍数，固定值展示磅值
            line_label = (f'{ls["value"]:g} 倍行距（≈{ls.get("equivalent_exact_pt", 0):g}pt）'
                          if ls.get("mode") == "multiple" and ls.get("value")
                          else (f"{line_pt:g}pt 行距" if line_pt else ""))
            v = d.get("verification", {})
            cc = v.get("cross_check", {})
            auth = v.get("typography_authority") or "published"
            lang = d["journal"].get("language")
            if lang not in ("zh", "en"):
                lang = "zh" if re.search(r"[\u4e00-\u9fff]", d["journal"]["name_zh"]) else "en"
            en_name = d["journal"].get("name_en", "")
            # 英文刊的 name_zh 往往就是英文名的截断版，标题一律用完整英文名
            display = en_name or d["journal"]["name_zh"] if lang == "en" else d["journal"]["name_zh"]
            items.append({
                "id": d["journal"]["id"],
                "lang": lang,
                "name": display,
                "en": en_name,
                "body": f'{body["font_cn"]} {body["size_pt"]:g}pt',
                "line": body.get("line_spacing", {}).get("value_pt"),
                "line_label": line_label,
                "levels": len(d["headings"]["levels"]),
                "cite": "编号制" if d["citation"]["system"] == "numbered" else "作者年制",
                "note_order": "→".join(
                    {"general": "注", "source": "来源", "data_source": "数据",
                     "cell": "①②"}.get(x["kind"], x["kind"])
                    for x in d["tables"]["note_layers"]),
                "words": (d.get("limits", {}).get("total") or {}).get("value"),
                "abstract": (d.get("limits", {}).get("abstract") or {}).get("value"),
                "status": v.get("status", "unverified"),
                "status_label": STATUS_LABEL.get(v.get("status"), "未知"),
                "authority": auth,
                "auth_label": AUTH_LABEL.get(auth, "未声明"),
                "auth_class": AUTH_CLASS.get(auth, "a-pub"),
                "samples": cc.get("samples"),
                "spec": (RULES_DIR / f'{d["journal"]["id"]}.spec.html').exists(),
            })
        except Exception as e:  # noqa: BLE001
            print(f"[warn] 跳过规则集 {f.name}：{e}", file=sys.stderr)
            continue

    # 中文按拼音、英文按字母，各自升序；两组分开排，避免混在一起
    items.sort(key=lambda j: (0, _pinyin_key(j["name"])) if j["lang"] == "zh"
               else (1, j["name"].lower()))
    return items


_JOURNALS_CACHE = {"key": None, "items": []}
_JOURNALS_CACHE_LOCK = threading.Lock()


def _journals_signature():
    """规则 JSON、规格页的存在性与 mtime/size：任一变化即让列表缓存失效。"""
    parts = []
    for f in sorted(RULES_DIR.glob("*.json")):
        st = f.stat()
        spec = f.with_suffix(".spec.html")
        spec_st = spec.stat() if spec.exists() else None
        parts.append((f.name, st.st_mtime_ns, st.st_size,
                      bool(spec_st), spec_st.st_mtime_ns if spec_st else 0,
                      spec_st.st_size if spec_st else 0))
    return (str(RULES_DIR), tuple(parts))


def load_journals():
    """带签名校验的期刊列表缓存。

    页面一次渲染会多次要列表（工作台、导航、换刊下拉）。每次重解析二十多份
    规则 JSON 没有意义；规则发布或规格页重建会改变签名，下一次请求自动刷新。
    返回浅拷贝，调用方 append/改列表不会污染缓存。
    """
    try:
        key = _journals_signature()
    except OSError:
        return _load_journals_uncached()   # 文件正被替换时退回无缓存读取
    with _JOURNALS_CACHE_LOCK:
        if _JOURNALS_CACHE["key"] != key:
            _JOURNALS_CACHE["items"] = _load_journals_uncached()
            _JOURNALS_CACHE["key"] = key
        return [dict(j) for j in _JOURNALS_CACHE["items"]]


def journal_groups(items=None):
    items = items if items is not None else load_journals()
    zh = [j for j in items if j["lang"] == "zh"]
    en = [j for j in items if j["lang"] != "zh"]
    groups = []
    if zh:
        groups.append({"key": "zh", "title": "中文期刊", "list": zh})
    if en:
        groups.append({"key": "en", "title": "英文期刊", "list": en})
    return groups


def journal_by_id(jid):
    for j in load_journals():
        if j["id"] == jid:
            return j
    return None


_cr = None


def compliance():
    """按需加载体检模块，复用它的判定逻辑，避免 Web 层再写一遍。"""
    global _cr
    if _cr is None:
        spec = importlib.util.spec_from_file_location(
            "compliance_report", ROOT / "tools" / "compliance_report.py")
        m = importlib.util.module_from_spec(spec)
        sys.modules["compliance_report"] = m
        sys.path.insert(0, str(ROOT / "tools"))
        spec.loader.exec_module(m)
        _cr = m
    return _cr


# ---------------------------------------------------------------- 任务目录

def remove_job_dir(d):
    """在持有任务锁时删除目录；渲染中返回 False，删除失败如实抛出。

    绝不能先 pop 锁再 rmtree：旧锁仍被渲染线程持有时，新请求会拿到第二把锁。
    """
    lock = job_lock(d.name)
    if not lock.acquire(blocking=False):
        return False
    try:
        if d.exists():
            shutil.rmtree(d)
        with _JOB_LOCKS_GUARD:
            if _JOB_LOCKS.get(d.name) is lock:
                _JOB_LOCKS.pop(d.name, None)
        return True
    finally:
        lock.release()


def sweep(now=None):
    now = time.time() if now is None else now
    if not WORK.exists():
        return
    for d in WORK.iterdir():
        try:
            # 只清任务目录，outputs/web 下若有其他人工目录不能按年龄误删。
            if d.is_dir() and JOB_RE.match(d.name) and now - d.stat().st_mtime > JOB_TTL:
                remove_job_dir(d)       # 忙碌任务保留，本轮完成后再按新 mtime 判断
        except OSError:
            continue


def job_dir(job):
    if not JOB_RE.match(job or ""):
        abort(404)
    d = WORK / job
    try:
        expired = d.is_dir() and time.time() - d.stat().st_mtime > JOB_TTL
    except OSError:
        expired = False
    if expired:
        sweep()
        abort(404)                      # 删除被占用也不能让过期内容继续可下载
    if not d.is_dir():
        abort(404)
    return d


def read_json_safe(path):
    """读 JSON 文件，坏文件（截断 / 编码错 / 读失败）返回 None，绝不抛异常。

    faulthandler 都装上了，说明「进程被硬杀」是预期场景，半截 JSON 会留下。
    """
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def read_meta(d):
    """读任务 meta。进程硬杀会留下截断的 JSON，一条坏 meta 不许把 history
    和任务页整页拖成 500（F34）：把它当作缺失，从现场文件推断最低限度的状态。

    推断只回答「这个任务走到了哪一步、还剩什么可用」，能救回多少救多少，
    救不回的字段留空，由各页面按缺失兜底（fail_page 会把推断结果写回，
    相当于就地自愈，不再等 24 小时 TTL 整目录删除）。
    """
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


def write_json_atomic(path, data):
    """把 JSON 完整写入同目录唯一临时名，再原子替换目标。

    临时名必须带随机后缀：固定的 meta.json.tmp / details.json.tmp 在两个
    写者并发时会互相覆盖，甚至一个 finally 把另一个尚未落位的临时文件删掉。
    同目录替换保证不跨卷；读方只会看到完整旧值或完整新值。
    """
    p = Path(path)
    tmp = p.with_name(f".{p.name}.{uuid.uuid4().hex}.tmp")
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        replace_with_retry(tmp, p)
    finally:
        tmp.unlink(missing_ok=True)


def write_meta(d, meta):
    """meta.json 原子写（F34），并支持同任务的并发写者。"""
    write_json_atomic(d / "meta.json", meta)


def read_json(path, default=None):
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default


# ---------------------------------------------------------------- 错误翻译

class ConvertError(Exception):
    def __init__(self, title, hint, detail=""):
        super().__init__(title)
        self.title, self.hint, self.detail = title, hint, detail


class RenderBusy(ConvertError):
    """同任务的上一次渲染还没结束：这个请求被挡在门外（F11）。

    做成 ConvertError 的子类，路由的 except 链照常能接住；但必须排在
    ConvertError 之前处理——忙碌态的处置是「直接响应、绝不落笔」，因为
    此刻另一轮渲染正在写现场文件（见 busy_page）。
    """

    def __init__(self, job):
        super().__init__("正在处理中",
                         "这个任务上一次排版还没有结束，请稍候片刻后刷新页面，"
                         "不要重复提交。")


_JOB_LOCKS = {}
_JOB_LOCKS_GUARD = threading.Lock()


def job_lock(job):
    """每个任务一把渲染锁。服务 threaded=True，双击提交会让两个请求同时
    进入 stage_render 互删互写同一批成品，产出与审计可能来自不同轮（F11）。"""
    with _JOB_LOCKS_GUARD:
        return _JOB_LOCKS.setdefault(job, threading.Lock())


def busy_page(err):
    """「正在处理」页。与 fail_page 的关键差别：不写 error.json / meta——
    此刻另一轮渲染正在写这些文件，任何落笔都会把两轮的现场搅在一起（F11）。"""
    return render_template("error.html", code="…", title=err.title, hint=err.hint), 200


def job_read_locked(view):
    """下载/报告预览与任务写锁互斥，避免读到跨轮提交窗口。"""
    @wraps(view)
    def wrapped(job, *args, **kwargs):
        if not JOB_RE.match(job or ""):
            return view(job, *args, **kwargs)
        job_dir(job)  # 先做过期清理；若进入视图后才做，会与这里的读锁自相排斥
        lock = job_lock(job)
        if not lock.acquire(blocking=False):
            return render_template(
                "error.html", code="…", title="任务正在更新",
                hint="排版或重解析尚未结束，请稍候再下载或预览，以免拿到不同轮次的文件。"), 409
        try:
            return view(job, *args, **kwargs)
        finally:
            lock.release()
    return wrapped


_TIMEOUT_MSG = ("转换超时",
                "这份稿件可能体积过大或图片特别多。请确认文件小于 80 MB，"
                "或删掉冗余的嵌入图片后重试。")


def explain(raw):
    # 超时必须先于正则：run() 直接把 TimeoutExpired 实例传进来，re.search
    # 对非字符串必抛 TypeError，超时友好页就永远出不来（F33）
    if isinstance(raw, subprocess.TimeoutExpired):
        return _TIMEOUT_MSG
    t = raw or ""
    if re.search(r"BadZipFile|not a zip|PackageNotFound|zipfile", t):
        return ("这份文件打不开",
                "它可能不是真正的 DOCX（例如把 .doc 后缀直接改成了 .docx），或者文件已损坏。"
                "用 Word 打开后「另存为」一份新的 .docx，通常就能解决。")
    if "TimeoutExpired" in t or "timed out" in t:
        return _TIMEOUT_MSG
    if re.search(r"KeyError|IndexError|TypeError|AttributeError", t):
        return ("稿件结构超出当前识别范围",
                "可能是封面、文本框或多级列表的结构比较特殊。可以先把可疑的封面页简化后重试，"
                "或把下方技术细节发给开发者排查。")
    return ("转换没有完成",
            "处理途中出现异常，你的原稿没有被修改。可以重试一次，或换一本期刊再试。")


def run(args):
    try:
        r = subprocess.run([PY] + [str(a) for a in args], capture_output=True,
                           text=True, cwd=ROOT, timeout=STEP_TIMEOUT,
                           encoding="utf-8", errors="replace")
    except subprocess.TimeoutExpired as e:
        raise ConvertError(*explain(e)) from e
    if r.returncode != 0:
        raw = (r.stderr or r.stdout or "")
        raise ConvertError(*explain(raw), detail=raw.strip()[-1500:])
    return r.stdout


SOFFICE = [shutil.which("soffice"),
           r"C:\Program Files\LibreOffice\program\soffice.exe",
           r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
           "/usr/bin/soffice"]

DOC2DOCX_PS1 = ROOT / "tools" / "doc2docx.ps1"

# .doc 转换单独设超时：起一个 COM 实例、打开、另存，正常 3–10 秒，120 秒是
# 留给几十页的大稿。沿用流水线的 STEP_TIMEOUT(300) 的话，Word 卡住时用户
# 要对着转圈的页面等五分钟，还得不到任何可行动的提示。
DOC_CONVERT_TIMEOUT = 120


def _com_convert(src, out, prog_id):
    """用本机 Office 套件把 .doc 另存成 .docx，返回 ps1 的退出码。

    0 = 成功；2 = 套件在但这份文件转不出来；3 = 套件没装。

    返回值分三档而不是布尔，是因为「没装 Word」和「Word 打不开这份稿」
    对用户是两句完全不同的话：前者让他换台机器，后者让他检查文件。

    走 PowerShell 而不是 pywin32：绿色包不想为一个「只在 Windows 有、只在
    老 .doc 上用一次」的能力背一份 dll 和 COM 的引用计数；进程隔离还顺带
    解决一个真实风险——Word 崩了不会把 Flask 一起带走。
    """
    if not DOC2DOCX_PS1.is_file():
        return 3
    # 绝对路径是必须的：Word 的 SaveAs 按**它自己进程**的当前目录解析相对
    # 路径（往往是「我的文档」），不是按 Python 的 cwd。传相对路径时 ps1 会
    # 报退出码 0 却把文件存到别处，或者干脆存不出来。
    cmd = ["powershell", "-NoProfile", "-NonInteractive",
           "-ExecutionPolicy", "Bypass", "-File", str(DOC2DOCX_PS1),
           "-In", str(Path(src).resolve()), "-Out", str(Path(out).resolve()),
           "-ProgId", prog_id]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace",
                           timeout=DOC_CONVERT_TIMEOUT)
    except subprocess.TimeoutExpired:
        app.logger.warning("%s 转换 .doc 超时：%s", prog_id, Path(src).name)
        return 2
    except OSError:
        return 3                      # 这台机器没有 PowerShell（非 Windows）
    if r.returncode == 0 and Path(out).exists():
        return 0
    if r.returncode == 3:
        return 3
    return 2


def _sniff_doc(src):
    """看文件头判断上传的「.doc」到底是什么，返回 doc / docx / rtf / unknown。

    为什么必须嗅探：Word 的「打开并修复」会把任何字节都当文本尽力打开，
    一个随机字节的文件能被转成一份 1 个块的垃圾 docx，全程不报错——用户
    拿到一份空稿子还以为是工具坏了。而「把 .docx 改名成 .doc」是极常见的
    操作失误，认出来的话可以直接用，省掉 5 秒 COM 往返。

    这三类文件头都是公开格式规范：
      OLE2 复合文档（D0 CF 11 E0…）—— Word 97-2003 及更早，真正的 .doc
      ZIP（PK\\x03\\x04）—— .docx，本质是个 zip 包
      RTF（{\\rtf）—— 写字板也能存的富文本
    """
    try:
        with open(src, "rb") as f:
            head = f.read(8)
    except OSError:
        return "unknown"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "doc"
    if head.startswith(b"PK\x03\x04"):
        return "docx"
    if head.startswith(b"{\\rtf"):
        return "rtf"
    return "unknown"


def doc_to_docx(src, d):
    """老 .doc 转成 docx，按「保真度 + 可得性」依次试本机已有的工具。

    顺序是有理由的：
      1. Word / WPS 的 COM —— 与作者手工「另存为 .docx」完全等价，这是保真度
         的上限。国内作者装 WPS 的比装 Word 的多，两个都要试
      2. LibreOffice —— 免费后备，但样式映射偶有偏差（分节、域、脚注编号）
    三条都没有时只说事实和出路，**绝不建议在线转换**：让人把未发表稿传上云
    转个格式，比转不了严重得多，「稿子不出本机」是这个产品的前提。
    """
    out = d / (src.stem + ".docx")
    out.unlink(missing_ok=True)

    kind = _sniff_doc(src)
    if kind == "docx":
        # 后缀写错了而已，内容本来就是 docx：直接拿来用。绕一趟 Office 不但
        # 白等 5 秒，还可能被 Word 的版本转换改掉版式。
        shutil.copy2(src, out)
        app.logger.info("上传文件实为 docx（后缀写成 .doc），已直接使用：%s", src.name)
        return out
    if kind == "unknown":
        raise ConvertError(
            "这个文件不是 Word 文档",
            "它的文件头既不像 .doc（Word 97-2003），也不像 .docx 或 RTF，"
            "多半是别的东西改了后缀，或者上传时文件损坏了。"
            "请回原文件确认格式后再上传。")

    # 至少有一个工具装了，就说明失败是文件的问题而不是环境的问题。
    # 这两句话对用户必须分开：「去装个 Word」和「这份稿子坏了」完全不同。
    any_installed = False

    for prog_id in ("Word.Application", "KWPS.Application"):
        code = _com_convert(src, out, prog_id)
        if code == 0:
            # 留下「用了哪个工具」的痕迹：转换保真度直接取决于它，用户报
            # 「.doc 转出来公式丢了」时，第一件事就是查这条日志
            app.logger.info(".doc 经 %s 转换为 docx：%s", prog_id, src.name)
            return out
        if code == 2:
            any_installed = True

    exe = next((c for c in SOFFICE if c and Path(c).exists()), None)
    if exe:
        any_installed = True
        try:
            subprocess.run([exe, "--headless", "--convert-to", "docx",
                            "--outdir", str(d), str(src)],
                           capture_output=True, timeout=DOC_CONVERT_TIMEOUT)
        except (OSError, subprocess.TimeoutExpired):
            pass
        if out.exists():
            app.logger.info(".doc 经 LibreOffice 转换为 docx：%s", src.name)
            return out
        # LibreOffice 偶有「文件名被改写」的情况（带空格或点号的 stem）
        cand = [p for p in d.glob("*.docx") if p.stem != "manuscript"]
        if cand:
            app.logger.info(".doc 经 LibreOffice 转换为 docx（文件名被改写）：%s", cand[0].name)
            return cand[0]

    if any_installed:
        raise ConvertError(
            "这份 .doc 转不出 DOCX",
            "本机的 Word / WPS / LibreOffice 都打不开它，最可能是文件损坏、"
            "设了打开密码，或者保存时用了很老的 Word 6.0/95 格式。"
            "请在原软件里打开确认一下，能打开就「另存为 .docx」再上传。")
    raise ConvertError(
        "这台机器上没有能读 .doc 的软件",
        "已试过 Microsoft Word、WPS 文字和 LibreOffice，都没找到。"
        ".doc 是 Word 97-2003 的老格式，必须由其中一款软件另存为 .docx 才能读。"
        "可以在另一台装了 Word 或 WPS 的电脑上另存后拷回来再上传。"
        "请不要使用在线转换网站——那等于把未发表的稿子传到别人的服务器上。")


# ---------------------------------------------------------------- 流水线

def stage_parse(d, jid, original):
    """解析与重解析共用任务锁，防止进行中目录被删除或重复写入。"""
    lock = job_lock(d.name)
    if not lock.acquire(blocking=False):
        raise RenderBusy(d.name)
    try:
        return _stage_parse(d, jid, original)
    finally:
        lock.release()


def _stage_parse(d, jid, original):
    """解析 + 体检。产出 manuscript.ir.json 与 triage.html。"""
    src, ir = d / "manuscript.docx", d / "manuscript.ir.json"
    # WEBLOG2 收尾：AI 补漏事件（调用失败/采纳数）已由 parse() log 进解析
    # stdout，但网页流水线以子进程跑解析器，stdout 本来无人接收——逐行转进
    # app.logger（server.log），识别汇总与存疑标题随之可在后台关窗后回溯。
    out = run(["tools/parse_docx.py", src, "-o", ir])
    for ln in (out or "").splitlines():
        ln = ln.strip()
        if ln:
            app.logger.info("[parse] %s", ln)
    ir_data = read_json(ir)
    # 空稿在这里就拦下：0 个内容块的文件能一路渲染出「0 块成品」，
    # 用户拿到一份什么都没排的稿子，比直接报错更糟。必须在生成任何
    # 交付物之前失败，避免任务目录里留下可下载的空成品。
    if not isinstance(ir_data, dict) or not (ir_data.get("blocks") or []):
        raise ConvertError(
            "这份稿件没有可排版的内容",
            "解析出 0 个内容块。请确认上传的不是空文档；如果原稿是扫描件或 PDF "
            "改的后缀，请提供可编辑的 .docx 源文件。")
    # 解析器看到的是任务目录里的 manuscript.docx，把来源名改回用户上传的文件名
    if original and ir_data.get("source") != original:
        ir_data["source"] = original
        # IR 会被 retry/details 当作恢复依据；半截 JSON 不能被误判成有效现场。
        write_json_atomic(ir, ir_data)
    run(["tools/compliance_report.py", ir, RULES_DIR / f"{jid}.json", "-o", d / "triage.html"])
    return ir_data


def _check_meta(path, strict_people=False):
    """交付前的硬闸门：元数据没清干净就不让文件出门。"""
    try:
        return docx_meta.assert_clean(path, strict_people=strict_people)
    except docx_meta.MetaLeakError as e:
        raise ConvertError(
            "成品未通过隐私检查",
            "文档属性里仍残留作者、公司或 WPS 账号信息。已停止交付，请重试；"
            "若反复出现，把下方技术细节发给开发者。",
            str(e)) from e


def stage_render(d, jid, original, prepare=None):
    """补录内容注入 → 排版 → 审计的入口：同任务一次只放一轮进来（F11）。

    prepare 用于提交本轮输入（details/reject/citation），也必须在持锁后执行。
    否则被 RenderBusy 挡回的请求虽已显示「没有提交」，却已改写下一轮输入。
    所有产物先写 .staging 临时名，全部成功后才进入提交阶段（F36）。
    """
    lock = job_lock(d.name)
    if not lock.acquire(blocking=False):
        raise RenderBusy(d.name)
    try:
        if prepare is not None:
            prepare()
        return _stage_render(d, jid, original)
    finally:
        lock.release()


def _stage_render(d, jid, original):
    """排版与审计的主体。调用方必须已持有本任务的渲染锁。"""
    src, ir = d / "manuscript.docx", d / "manuscript.ir.json"
    rules = RULES_DIR / f"{jid}.json"
    outdocx = d / f"manuscript.{jid}.docx"
    tmp_docx = d / f".staging.{jid}.docx"
    tmp_docx.unlink(missing_ok=True)      # 上次中断可能留下的半成品

    t0 = time.time()

    # details.json 由路由在抢锁前原子写（tmp + os.replace）：半截 JSON 已不可能，
    # 但 Windows 上读方仍可能在替换瞬间撞上 sharing violation（毫秒不到的窗口）。
    # 这里重试而不是静默当空表——静默会把用户刚提交的补录内容整轮丢掉。
    details = {}
    for attempt in range(3):
        try:
            details = read_json(d / "details.json", {}) or {}
            break
        except OSError:
            if attempt == 2:
                raise
            time.sleep(0.02)
    filled_docx, filled_ir = d / "filled.docx", d / "filled.ir.json"
    injected = []
    if any((details.get(k) or "").strip() for k in details):
        run(["tools/apply_details.py", src, ir, d / "details.json", rules,
             "--docx-out", filled_docx, "--ir-out", filled_ir])
        for row in read_json(filled_ir, {}).get("blocks", []):
            if row.get("injected"):
                injected.append(row["type"])
        render_ir = filled_ir
    else:
        render_ir = ir
        for p in (filled_docx, filled_ir):
            p.unlink(missing_ok=True)

    rules_data = read_json(rules, {}) or {}
    reject_file = d / "reject.json"
    meta = read_meta(d)
    cmd = ["tools/render_from_ir.py", render_ir, rules, "-o", tmp_docx]
    if reject_file.is_file():
        cmd += ["--reject", reject_file]
    if meta.get("citation"):
        cmd.append("--citation")
    run(cmd)
    _check_meta(tmp_docx)
    # 体检报告按"注入之后"的稿子重跑，否则填完英文摘要仍会显示"缺 6 项"
    tmp_triage = d / ".staging.triage.html"
    run(["tools/compliance_report.py", render_ir, rules, "-o", tmp_triage])

    audit_json = tmp_docx.with_suffix(".audit.json")
    # 匿名稿：该刊要求正文匿名时另出一份，只删身份信息（§11.4 双版本交付）
    anon_docx = None
    anon_tmp = None
    if (rules_data.get("submission") or {}).get("anonymous"):
        anon_docx = d / f"manuscript.{jid}.anon.docx"
        anon_tmp = d / f".staging.{jid}.anon.docx"
        anon_tmp.unlink(missing_ok=True)
        try:
            out = run(["tools/anonymize.py", tmp_docx, render_ir, rules, "-o", anon_tmp])
            info = json.loads(out.strip().splitlines()[-1]) if out.strip() else {}
            _check_meta(anon_tmp, strict_people=True)
        except (ConvertError, ValueError) as e:
            # 该刊要求匿名时，匿名稿是交付要件；生成失败或隐私检查不过，
            # 不能再把署名稿当成品交付，否则用户会以为已经匿名。
            raise ConvertError(
                "匿名稿未通过隐私检查",
                "这份稿子要求匿名，但匿名稿没有生成成功，或元数据里仍有作者信息。"
                "已停止交付，请重试；若反复出现，把下方技术细节发给开发者。",
                str(e)) from e
        if info.get("removed"):
            data = read_json(audit_json, {"items": []}) or {"items": []}
            data["anonymize"] = {"removed": info["removed"],
                                 "file": anon_docx.name if anon_docx else ""}
            audit_json.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                  encoding="utf-8")

    tmp_audit_html = d / ".staging.audit.html"
    run(["tools/render_audit.py", audit_json, "-o", tmp_audit_html, "--job", d.name])

    audit = read_json(audit_json, {"items": []})
    audit_items = audit.get("items", [])
    # 首页数字只能数真改动：渲染器把「看过但没动」的说明项（含域锁定、
    # 跳过的引用转换、伪表格公式）也记进同一个列表，直接取 len 会把说明
    # 当改动报给用户——报告自己虚报，整套审计就没有说服力了。
    n_changed = sum(1 for it in audit_items if it.get("changed", True))
    n_notes = len(audit_items) - n_changed
    ir_data = read_json(render_ir) or {}
    j = journal_by_id(jid) or {"id": jid, "name": jid}

    meta = read_meta(d)
    meta.update({
        "job": d.name,
        "journal": jid,
        "journal_name": j["name"],
        "original": original,
        "stem": Path(original).stem or "稿件",
        "blocks": ir_data.get("block_total"),
        "audit_total": n_changed,
        "audit_notes": n_notes,
        "audit_checked": audit.get("checked_blocks"),
        "elapsed": round(time.time() - t0, 1),
        "created": time.strftime("%Y-%m-%d %H:%M"),
        "docx": outdocx.name,
        "docx_anon": anon_docx.name if anon_docx else "",
        "anon_removed": len((read_json(audit_json, {}) or {}).get("anonymize", {}).get("removed", [])),
        "rejected": len(audit.get("rejected", [])),
        "injected": injected,
        "stage": "rendered",
    })
    # ---- 提交阶段：先备份本轮会覆盖的旧产物，再逐项原子落位，最后发布 meta。
    # meta 若先写，后续 replace 失败会让结果页指向不存在的文件；同一刊重渲染
    # 还会把新成品与旧审计混在一轮。这里对异常路径回滚已落位项（F36）。
    final_audit = outdocx.with_suffix(".audit.json")
    tmp_audit = tmp_docx.with_suffix(".audit.json")
    keep = {outdocx.name, final_audit.name}
    replacements = []
    if anon_docx:
        keep.add(anon_docx.name)
        replacements.append((anon_tmp, anon_docx))
    replacements.extend([
        (tmp_docx, outdocx),
        (tmp_audit, final_audit),
        (tmp_triage, d / "triage.html"),
        (tmp_audit_html, d / "audit.html"),
    ])

    token = uuid.uuid4().hex
    backups, applied = {}, []
    preserve_backups = set()
    try:
        for _src, dst in replacements:
            bak = None
            if dst.exists():
                bak = d / f".rollback.{token}.{len(backups)}"
                shutil.copy2(dst, bak)
            backups[dst] = bak
        for src, dst in replacements:
            replace_with_retry(src, dst)
            applied.append(dst)
        write_meta(d, meta)             # 元数据最后发布，始终指向完整的新一轮
    except Exception as commit_error:
        rollback_errors = []
        for dst in reversed(applied):
            bak = backups[dst]
            try:
                if bak is None:
                    dst.unlink(missing_ok=True)
                else:
                    replace_with_retry(bak, dst)
            except OSError as e:
                if bak is not None:
                    preserve_backups.add(bak)  # 恢复失败时，备份就是唯一旧版，绝不能清理
                rollback_errors.append(f"{dst.name}: {e}")
        if rollback_errors:
            raise ConvertError(
                "排版结果未能完整落位",
                "文件可能被 Word 或预览程序占用。请关闭相关窗口后重试；"
                "旧产物已尽力恢复；未恢复项的备份保留在任务目录（.rollback.*），请先核对再下载。",
                "; ".join(rollback_errors)) from commit_error
        raise
    finally:
        for bak in backups.values():
            if bak is not None and bak not in preserve_backups:
                try:
                    bak.unlink(missing_ok=True)
                except OSError:
                    app.logger.warning("回滚备份暂未删除：%s", bak)

    # 新 meta 已发布后再清旧刊遗留；清理失败不应把已成功的一轮误报为失败。
    for old in (list(d.glob("manuscript.*.docx")) + list(d.glob("manuscript.*.audit.json"))
                + list(d.glob(".staging.*"))):
        if old.name == "manuscript.docx" or old.name in keep:
            continue
        try:
            old.unlink(missing_ok=True)
        except OSError:
            app.logger.warning("旧产物暂未删除：%s", old)
    try:
        (d / "error.json").unlink(missing_ok=True)
    except OSError:
        app.logger.warning("旧错误标记暂未删除：%s", d / "error.json")
    return meta


def _scrub_local_paths(text):
    """技术细节脱敏（F44）：traceback / format_exc / 子进程 stderr 里的本机
    绝对路径会原样进 error.json 并渲染到失败页，页面还引导转发给开发者——
    贴公开 issue 即泄露目录结构（解释器目录还带 C:\\Users\\<用户名>）。
    落盘与渲染前统一把仓库根、用户主目录与解释器安装目录替换成「.」；
    正斜杠写法（as_posix / 手拼路径）一并替换。异常信息本身保留，只抹路径。

    解释器目录（SCRUBEXE）：标准库帧的路径形如 <prefix>\\Lib\\traceback.py，
    只抹仓库根与主目录时仍会残留安装盘的目录结构。
    """
    if not text:
        return text
    exe_dir = Path(sys.executable).parent
    seen = set()
    for p in (ROOT, Path.home(), Path(sys.prefix), exe_dir):
        for s in (str(p), p.as_posix()):
            if s and s != "." and s not in seen:
                seen.add(s)
                text = text.replace(s, ".")
    return text


def fail_page(d, err, journals=None):
    meta = read_meta(d)
    err.detail = _scrub_local_paths(err.detail)   # error.json 与页面同一份 detail（F44）
    write_json_atomic(d / "error.json",
                      {"title": err.title, "hint": err.hint, "detail": err.detail})
    write_meta(d, meta)
    return render_template("failed.html", job=d.name, err=err, meta=meta,
                           journals=journals if journals is not None else load_journals(),
                           selected=meta.get("journal"),
                           # IR 不存在 = 解析就没成功过：换刊重试不可能有用，
                           # 页面要给「重新上传」的出路而不是诱导再试（F12）
                           can_retry=isinstance(read_json_safe(d / "manuscript.ir.json"), dict)), 200


# ---------------------------------------------------------------- 补录页数据

MISSING_TO_FIELD = {
    ("affiliations", "zh"): "affiliation_zh",
    ("classification_code", None): "classification",
    ("classification_code", "en"): "english_classification",
    ("title", "en"): "english_title",
    ("authors", "en"): "english_authors",
    ("affiliations", "en"): "english_affiliation",
    ("abstract", "en"): "english_abstract",
    ("keywords", "en"): "english_keywords",
}


def author_page_fields(rules, saved=None):
    """该刊要求作者信息单独成页时，列出这一页还需要作者填的项。

    字段清单来自规则集 personal_fields（§10.2：要问什么由目标期刊决定），
    已经出现在正文补录或首页脚注里的项不重复问。
    """
    ap = (rules.get("submission") or {}).get("author_page") or {}
    if not ap.get("separate_page"):
        return []
    known = {k for k, *_ in DETAIL_FIELDS} | {k for k, *_ in FOOTER_FIELDS}
    # 这几项可由稿件或首页脚注自动带出，不再让用户重输（§10.3）
    known |= {"paper_title", "author_name", "authors", "abstract", "keywords",
              "paper_title_en", "funding_info", "funding"}
    wanted = set(ap.get("fields") or [])
    saved = saved or {}
    out = []
    for f in rules.get("personal_fields") or []:
        fid = f.get("id")
        if fid in known or (wanted and fid not in wanted):
            continue
        out.append({"key": fid, "label": f.get("label_zh") or fid,
                    "hint": f.get("hint") or "", "value": saved.get(fid, "")})
    return out


def details_context(d, meta):
    jid = meta.get("journal", "")
    rules = rules_of(jid) or {}
    ir = read_json(d / "manuscript.ir.json", {}) or {}
    try:
        res = compliance().check(ir, rules)
        rows_raw = res.get("rows_raw") or []
    except Exception:  # noqa: BLE001 —— 体检失败不该挡住补录页
        rows_raw = []

    missing_keys = []
    unknown_missing = []
    for t, lang, name, ok, req in rows_raw:
        if ok or not req:
            continue
        key = MISSING_TO_FIELD.get((t, lang))
        if key:
            missing_keys.append(key)
        elif t in compliance().NON_ACTIONABLE:
            continue
        else:
            unknown_missing.append(name)

    # 同一类型可能有中英两套（分类号），判断"已有"必须带上语言
    present = {(b["type"], "en" if b["type"].startswith("english_") or b.get("lang") == "en" else "zh")
               for b in ir.get("blocks", [])}
    field_rows = []
    for key, btype, lang, label, hint in DETAIL_FIELDS:
        need = key in missing_keys
        has = (btype, lang) in present
        if need or has or key == "affiliation_zh":
            field_rows.append({"key": key, "label": label, "hint": hint,
                               "need": need, "value": ""})

    fm = [b.get("text", "") for b in ir.get("blocks", []) if b["type"] == "front_matter"]
    authors = fm[-1] if fm else ""
    stats = ir.get("stats") or {}
    keys = [re.sub(r"^关键词\s*[:：]\s*", "", b.get("text", ""))
            for b in ir.get("blocks", []) if b["type"] == "keywords"]
    abstract = [b.get("text", "") for b in ir.get("blocks", []) if b["type"] == "abstract"]

    saved = read_json(d / "details.json", {}) or {}
    for f in field_rows:
        f["value"] = saved.get(f["key"], "")

    personal = [p for p in (rules.get("personal_fields") or [])
                if p.get("placement") != "author_page"]

    # 引用体系重排只在「规则集给了条目模板 + 稿子确实用编号引用」时才提供
    cite_cfg = rules.get("citation") or {}
    cite_rl = cite_cfg.get("reference_list") or {}
    has_cite = any("[" in (b.get("text") or "") and "]" in (b.get("text") or "")
                   for b in ir.get("blocks", []) if b["type"] == "body")
    cite_label = {"author-year": "作者-年份制（张三，2020）",
                  "numbered": "顺序编码制（[1]）"}.get(cite_cfg.get("system"), "")
    # 勾选闸与渲染层 apply_citation 的硬闸同一判据（P4-1 阶段 D）：只对声明
    # 作者-年份体的刊开。编号实引刊勾了也会被硬闸整批拒绝，UI 不该开出
    # 永远落空的勾选框。
    citation_offered = (cite_cfg.get("system") == "author-year"
                        and bool(cite_rl.get("templates")) and has_cite)
    return {
        "rules": rules,
        "citation_offered": citation_offered,
        "citation_label": cite_label,
        "citation_on": bool(meta.get("citation")),
        "journal": meta.get("journal_name", jid),
        "title": fm[0] if fm else "",
        "authors": authors,
        "abstract_len": len(re.sub(r"\s", "", abstract[0])) if abstract else 0,
        "keywords": keys[0] if keys else "",
        "fields": field_rows,
        "n_need": sum(1 for f in field_rows if f["need"]),
        "unknown_missing": unknown_missing,
        "footer_fields": [(k, label, hint, saved.get(k, "")) for k, label, hint in FOOTER_FIELDS],
        "author_fields": author_page_fields(rules, saved),
        "personal": personal,
        "stats": stats,
        "saved": saved,
    }


# ---------------------------------------------------------------- 页面

@app.route("/")
def index():
    return render_template("index.html", journal_count=len(load_journals()))


@app.route("/favicon.ico")
def favicon():
    """浏览器不带 <link> 也会来要这个路径；--app 窗口的任务栏图标同样取自它。

    真文件在 static/favicon.ico（tools/make_icon.py 生成）；旧的内联 SVG
    data-URI 在独立窗口壳里支持不稳，ICO 是全场景兜住的格式。
    """
    return send_from_directory(ROOT / "static", "favicon.ico")


def render_workspace(error=None, preselect=""):
    """工作台只有这一个入口：报错页也必须带全期刊列表。

    曾经 @app.route("/workspace") 与 /convert 的报错分支各写一遍 render_template，
    报错分支漏了 groups 和 default_id —— 页面能出来，但「选择期刊」里一本都没有，
    一个单选也勾不上，用户就卡在死页上再也提交不了。
    """
    items = load_journals()
    known = {j["id"] for j in items}
    default_id = preselect if preselect in known else (items[0]["id"] if items else "")
    return render_template("workspace.html", journals=items, groups=journal_groups(items),
                           error=error, default_id=default_id, preselect=preselect)


@app.route("/workspace")
def workspace():
    jid = request.args.get("journal") or ""
    if jid and not JID_RE.match(jid):    # 预选参数也要过体例：不合法 → 400（F47）
        abort(400)
    return render_workspace(preselect=jid if (RULES_DIR / f"{jid}.json").exists() else "")


@app.route("/journals")
def journals():
    return render_template("journals.html", journals=load_journals(),
                           drafts=draft_list())


@app.context_processor
def inject_ai_state():
    """每个页面都要知道 AI 开没开：导航要显示状态，「本地安全」的措辞也要跟着变。"""
    try:
        cfg = ai_assist.load_config()
        return {"ai_on": ai_assist.available(cfg),
                "ai_enabled": bool(cfg.get("enabled")),
                "ai_send_mode": cfg.get("send_mode") or "snippet"}
    except Exception:
        return {"ai_on": False, "ai_enabled": False, "ai_send_mode": "snippet"}


# ---------------------------------------------------------------- AI 设置

def _ai_form():
    """从表单取值。api_key 留空表示沿用已保存的，避免把密钥回显到页面上。"""
    cur = ai_assist.load_config()
    key = (request.form.get("api_key") or "").strip() or str(cur.get("api_key") or "")

    def num(name, default):
        try:
            return max(1, int(request.form.get(name) or default))
        except (TypeError, ValueError):
            return default

    return {
        "enabled": request.form.get("enabled") == "on",
        "base_url": (request.form.get("base_url") or "").strip(),
        "model": (request.form.get("model") or "").strip(),
        "api_key": key,
        "send_mode": request.form.get("send_mode") or "snippet",
        "timeout": num("timeout", 60),
        "snippet_chars": num("snippet_chars", 60),
    }


@app.route("/settings", methods=["GET", "POST"])
def settings():
    error = success = None
    if request.method == "POST":
        cfg = _ai_form()
        errs = ai_assist.validate_config(cfg)
        if errs:
            error = "；".join(errs)
        else:
            ai_assist.save_config(cfg)
            success = "已保存。改动下一次转换即可生效。"
    cfg = ai_assist.load_config()
    return render_template("settings.html", cfg=cfg, error=error, success=success,
                           key_set=bool(str(cfg.get("api_key") or "").strip()),
                           config_path=str(ai_assist.CONFIG_PATH))


@app.route("/settings/test", methods=["POST"])
def settings_test():
    """测当前表单里的配置，不必先保存——填完就能验，省一轮来回。"""
    cfg = _ai_form() if (request.form.get("base_url") or "").strip() else ai_assist.load_config()
    ok, msg = ai_assist.test_connection(cfg, timeout=25)
    return jsonify(ok=ok, msg=msg)


# ---------------------------------------------------------------- 新建期刊

def draft_dir(jid):
    if not JID_RE.match(jid or ""):
        abort(404)
    d = DRAFT_DIR / jid
    if not d.is_dir():
        abort(404)
    return d


def draft_list():
    rows = []
    if DRAFT_DIR.exists():
        for d in sorted(DRAFT_DIR.iterdir()):
            f = d / "draft.json"
            if not (d.is_dir() and f.exists()):
                continue
            try:
                r = json.loads(f.read_text(encoding="utf-8"))
                rep = read_json(d / "report.json", {}) or {}
            except Exception:
                continue
            rows.append({
                "id": d.name,
                "name": (r.get("journal") or {}).get("name_zh") or d.name,
                "publishable": bool(rep.get("ok_to_publish")),
                "extracted": (rep.get("stats") or {}).get("extracted", 0),
                "inherited": (rep.get("stats") or {}).get("inherited", 0),
            })
    return rows


@app.route("/journals/new")
def journal_new():
    return render_template("journal_new.html", ai_ok=ai_assist.available(),
                           cfg=ai_assist.load_config(), error=None,
                           ocr_ok=ocr.available(), ocr_engine=ocr.engine_name(),
                           drafts=draft_list(),
                           existing={j["id"] for j in load_journals()})


@app.route("/journals/extract", methods=["POST"])
def journal_extract():
    def back(msg, status=400):
        return render_template("journal_new.html", ai_ok=ai_assist.available(),
                               cfg=ai_assist.load_config(), error=msg,
                               ocr_ok=ocr.available(), ocr_engine=ocr.engine_name(),
                               drafts=draft_list(),
                               existing={j["id"] for j in load_journals()}), status

    if not ai_assist.available():
        return back("AI 未配置或未启用。请先到「设置」填写接口与密钥。")
    f = request.files.get("file")
    jid = (request.form.get("jid") or "").strip().lower()
    name = (request.form.get("name") or "").strip()
    if not f or not f.filename:
        return back("请选择一份稿约 PDF")
    if not f.filename.lower().endswith(".pdf"):
        return back("请上传 PDF：稿约通常是 PDF，Word 稿约请先另存为 PDF")
    if not JID_RE.match(jid):
        return back("期刊代号只能用小写字母、数字、下划线，2–24 位，例如 jryj")
    lock = job_lock(f"journal:{jid}")
    if not lock.acquire(blocking=False):
        return back("这本期刊正在抽取或发布，请稍候再试。", 409)
    try:
        # 持锁后再查一次，避免长时间抽取期间另一请求已完成同名发布。
        if (RULES_DIR / f"{jid}.json").exists():
            return back(f"代号 {jid} 已被占用，换一个")

        d = DRAFT_DIR / jid
        d.mkdir(parents=True, exist_ok=True)
        src = d / "source.pdf"
        f.save(src)

        try:
            # 抽取过程的事件（F52 截断降级重试、OCR 丢页/低文字层）必须记进
            # 服务日志——CLI 链路走 stderr，web 链路不传 log 就全部静默，
            # draft.json 里只有结果，查不到「发生过」。
            rules, report, text = extract_journal.extract(
                src, jid=jid, name_zh=name, log=app.logger.warning)
        except Exception as e:  # noqa: BLE001
            app.logger.exception("稿约抽取失败")
            return back(f"抽取失败：{e}")

        write_json_atomic(d / "draft.json", rules)
        write_json_atomic(d / "report.json", report)
        (d / "source.txt").write_text(text, encoding="utf-8")
        return redirect(url_for("journal_draft", jid=jid))
    finally:
        lock.release()


@app.route("/journals/draft/<jid>")
def journal_draft(jid):
    d = draft_dir(jid)
    rules = read_json(d / "draft.json", {}) or {}
    report = read_json(d / "report.json", {}) or {}
    ext = (rules.get("verification") or {}).get("extraction") or {}
    return render_template("journal_draft.html", jid=jid, rules=rules, report=report,
                           ext=ext, journal=rules.get("journal") or {},
                           evidence=ext.get("used") or {},
                           inherited=ext.get("inherited") or [],
                           dropped=ext.get("dropped") or [],
                           ai_on=ai_assist.available())


@app.route("/journals/draft/<jid>/publish", methods=["POST"])
def journal_publish(jid):
    if not JID_RE.match(jid or ""):
        abort(404)
    lock = job_lock(f"journal:{jid}")
    if not lock.acquire(blocking=False):
        return render_template(
            "error.html", code="…", title="这本期刊正在发布",
            hint="另一轮构建或清理尚未结束，本次没有重复发布。请稍候再试。"), 409
    try:
        return _journal_publish(jid)
    finally:
        lock.release()


def _journal_publish(jid):
    # 持锁后重新读取草稿与目标状态：覆盖确认不能停留在构建开始前的旧现场。
    d = draft_dir(jid)
    rules = read_json(d / "draft.json", {}) or {}
    report = read_json(d / "report.json", {}) or {}

    def back(msg, status=400):
        ext = (rules.get("verification") or {}).get("extraction") or {}
        return render_template("journal_draft.html", jid=jid, rules=rules, report=report,
                               ext=ext, journal=rules.get("journal") or {},
                               evidence=ext.get("used") or {},
                               inherited=ext.get("inherited") or [],
                               dropped=ext.get("dropped") or [],
                               ai_on=ai_assist.available(), error=msg), status

    if not report.get("ok_to_publish"):
        return back("还有阻断性错误，先把错误项修掉")
    if request.form.get("confirm") != "on":
        return back("请先勾选「我已逐项核对」——AI 抽出来的东西不能未经核对就进规则库")
    target = RULES_DIR / f"{jid}.json"
    if target.exists() and request.form.get("overwrite") != "on":
        return back(f"{jid}.json 已存在。确实要覆盖请勾选覆盖选项")

    # 先在草稿目录完成全部构建。以前先覆盖正式 JSON，再吞掉构建异常并删除
    # 草稿/备份，会留下「新规则 + 旧样板」且无法恢复的现场。
    try:
        run(["tools/build_journal.py", str(d / "draft.json")])
        for suffix in (".sample.docx", ".spec.html"):
            if not (d / f"{jid}{suffix}").is_file():
                raise ConvertError("规则集产物不完整", f"缺少 {jid}{suffix}")
    except Exception as e:  # noqa: BLE001
        app.logger.exception("规则集构建失败，正式库未修改：%s", jid)
        return back(f"发布未完成：{e}。正式规则库未修改，草稿已保留，请修正后重试。", 500)

    destinations = [target] + [RULES_DIR / f"{jid}{s}"
                               for s in (".sample.docx", ".sample.pdf", ".spec.html")]
    token = uuid.uuid4().hex
    backups, applied = {}, []
    try:
        for dst in destinations:
            bak = None
            if dst.exists():
                bak = d / f".publish-backup.{token}.{len(backups)}"
                shutil.copy2(dst, bak)
            backups[dst] = bak

        write_json_atomic(target, rules)
        applied.append(target)
        for suffix in (".sample.docx", ".sample.pdf", ".spec.html"):
            src = d / f"{jid}{suffix}"
            if src.is_file():               # PDF 可因未安装 LibreOffice 而缺省
                dst = RULES_DIR / src.name
                replace_with_retry(src, dst)
                applied.append(dst)
    except Exception:  # noqa: BLE001
        app.logger.exception("规则集落位失败：%s", jid)
        rollback_errors = []
        for dst in reversed(applied):
            bak = backups[dst]
            try:
                if bak is None:
                    dst.unlink(missing_ok=True)
                else:
                    replace_with_retry(bak, dst)
            except OSError as rollback_error:
                rollback_errors.append(f"{dst.name}: {rollback_error}")
        msg = "发布落位失败，已恢复旧版；草稿与备份均已保留，请关闭占用文件后重试。"
        if rollback_errors:
            msg = "发布落位失败，且部分旧版未能自动恢复；备份保留在草稿目录，请先人工核对。"
        return back(msg, 500)

    # 本轮没有 PDF 时，旧 PDF 不能留下冒充新样板预览；清理失败只记日志，
    # 规格页本身已不含旧预览，不影响已经成功的发布。
    if not (d / f"{jid}.sample.pdf").exists():
        try:
            (RULES_DIR / f"{jid}.sample.pdf").unlink(missing_ok=True)
        except OSError:
            app.logger.warning("旧样板 PDF 暂未删除：%s", jid)
    try:
        shutil.rmtree(d)                  # 只有全部成功才允许清草稿与备份
    except OSError:
        app.logger.warning("发布成功，但草稿目录暂未清理：%s", jid)
    return redirect(url_for("journal_spec", jid=jid))


@app.route("/journals/draft/<jid>/discard", methods=["POST"])
def journal_discard(jid):
    if not JID_RE.match(jid or ""):
        abort(404)
    lock = job_lock(f"journal:{jid}")
    if not lock.acquire(blocking=False):
        return render_template("error.html", code="…", title="期刊草稿正在处理中",
                               hint="构建或发布尚未结束，本次没有丢弃草稿。"), 409
    try:
        try:
            shutil.rmtree(draft_dir(jid))
        except OSError:
            return render_template("error.html", code="…", title="草稿暂未丢弃",
                                   hint="有文件仍被占用，请关闭预览窗口后重试。"), 409
        return redirect(url_for("journals"))
    finally:
        lock.release()


@app.route("/history")
def history():
    sweep()
    rows = []
    for d in WORK.iterdir() if WORK.exists() else []:
        if not d.is_dir() or not JOB_RE.match(d.name):
            continue
        # 逐条兜底：目录在遍历途中被清理、meta 读出意外结构，都只跳过这一行，
        # 不让一条坏数据把整个 history 页拖成 500（F34）
        try:
            m = read_meta(d)
            if not m.get("original"):
                continue
            rows.append({
                "job": d.name,
                "original": m.get("original"),
                "journal": m.get("journal_name") or m.get("journal", ""),
                "created": m.get("created", ""),
                "stage": m.get("stage") or ("rendered" if m.get("docx") else "details"),
                "audit": m.get("audit_total"),
                "mtime": d.stat().st_mtime,
            })
        except OSError:
            continue
    rows.sort(key=lambda r: r["mtime"], reverse=True)
    return render_template("history.html", rows=rows[:30])


@app.route("/job/<job>/delete", methods=["POST"])
def job_delete(job):
    """单任务删除（F57）：任务目录里有论文标题与成品，24 小时 TTL 之内
    同机他人仍能看到——历史页要给「现在就删」的入口，隐私承诺才算落地。

    两层校验分开：id 格式不对是请求本身的问题（400），格式对但目录已不在
    （TTL 清掉了 / 已删过）按 404 兜底，都过才 rmtree。
    """
    if not JOB_RE.match(job or ""):
        abort(400)
    d = WORK / job
    if not d.is_dir():
        abort(404)
    try:
        removed = remove_job_dir(d)
    except OSError:
        app.logger.exception("任务删除失败：%s", job)
        return render_template(
            "error.html", code="…", title="任务暂未删除",
            hint="有文件仍被 Word、预览程序或其他进程占用。请关闭相关窗口后重试；"
                 "本次没有显示为删除成功。"), 409
    if not removed:
        return render_template(
            "error.html", code="…", title="任务正在处理中",
            hint="排版尚未结束，为避免破坏成品，本次没有删除。请稍候再试。"), 409
    return redirect(url_for("history"))


@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    jid = request.form.get("journal") or ""
    # keep 判断先过 JID_RE：..\ 反斜杠 id 不能参与 RULES_DIR 拼接（F47），
    # 不合法的落到 keep="" 走「请选择目标期刊」的既有 400 分支
    keep = jid if (JID_RE.match(jid) and (RULES_DIR / f"{jid}.json").exists()) else ""
    if not f or not f.filename:
        return render_workspace("请选择一份 DOCX 稿件", keep), 400
    name = Path(f.filename).name
    if not name.lower().endswith((".docx", ".doc")):
        return render_workspace(
            "成品只能从 .docx / .doc 生成：排版是在你的原文件上改的，"
            "图片、公式、脚注都靠原对象原样保留。PDF 里这些已是画面，"
            "取不回可编辑对象，硬转出来的稿子交给编辑部反而更危险。"
            "手头只有 PDF 的话，请向作者要源文件。", keep), 400
    if not keep:
        return render_workspace("请选择一本目标期刊后再开始转换", ""), 400

    sweep()
    job = uuid.uuid4().hex[:12]
    d = WORK / job
    d.mkdir(parents=True, exist_ok=True)
    write_meta(d, {"job": job, "journal": jid, "original": name,
                   "stem": Path(name).stem or "稿件", "stage": "uploaded"})

    try:
        if name.lower().endswith(".doc"):
            raw = d / "upload.doc"
            f.save(raw)
            doc = doc_to_docx(raw, d)
            shutil.move(str(doc), str(d / "manuscript.docx"))
            raw.unlink(missing_ok=True)
        else:
            f.save(d / "manuscript.docx")
        stage_parse(d, jid, name)
    except ConvertError as e:
        return fail_page(d, e)
    except Exception:  # noqa: BLE001
        return fail_page(d, ConvertError(*explain(""), detail=traceback.format_exc()[-1500:]))

    meta = read_meta(d)
    meta["stage"] = "details"
    write_meta(d, meta)
    return redirect(url_for("details", job=job))


@app.route("/job/<job>/details", methods=["GET", "POST"])
def details(job):
    d = job_dir(job)
    meta = read_meta(d)
    if not meta.get("journal"):
        return redirect(url_for("workspace"))

    # 文件存在但 JSON 已截断时不能被「存在即有效」骗过；原稿仍在就直接重解析。
    ir_path = d / "manuscript.ir.json"
    if ir_path.exists() and not isinstance(read_json_safe(ir_path), dict):
        try:
            stage_parse(d, meta["journal"], meta.get("original", "稿件.docx"))
        except RenderBusy as e:
            return busy_page(e)
        except ConvertError as e:
            return fail_page(d, e)
        except Exception:  # noqa: BLE001
            return fail_page(d, ConvertError(*explain(""), detail=traceback.format_exc()[-1500:]))
        meta = read_meta(d)

    if request.method == "POST":
        data = {}
        for key, *_ in DETAIL_FIELDS:
            data[key] = (request.form.get(key) or "").strip()
        for key, *_ in FOOTER_FIELDS:
            data[key] = (request.form.get(key) or "").strip()
        for f in author_page_fields(rules_of(meta["journal"]) or {}):
            data[f["key"]] = (request.form.get(f["key"]) or "").strip()
        citation = request.form.get("citation") == "on"

        def prepare():
            # 输入与开关都在持锁后落笔：忙碌请求必须真正零副作用（F11）。
            write_json_atomic(d / "details.json", data)
            current = read_meta(d)
            # 引用体系重排是唯一会改写正文/条目文字的开关，默认关，需显式勾选。
            current["citation"] = citation
            write_meta(d, current)

        try:
            stage_render(d, meta["journal"], meta.get("original", "稿件.docx"),
                         prepare=prepare)
        except RenderBusy as e:      # 忙碌态先于 ConvertError：不能落笔写现场（F11）
            return busy_page(e)
        except ConvertError as e:
            return fail_page(d, e)
        except Exception:  # noqa: BLE001
            return fail_page(d, ConvertError(*explain(""),
                                             detail=traceback.format_exc()[-1500:]))
        return redirect(url_for("result", job=job))

    return render_template("details.html", job=job, meta=meta, **details_context(d, meta))


@app.route("/result/<job>")
def result(job):
    d = job_dir(job)
    meta = read_meta(d)
    if not meta.get("docx"):
        return redirect(url_for("details", job=job))
    j = journal_by_id(meta["journal"]) or {
        "id": meta["journal"], "name": meta.get("journal_name", meta["journal"]),
        "auth_class": "a-pub", "auth_label": "未声明",
        "status_label": "未知", "status": "unverified",
    }
    return render_template("result.html", job=job, journal=j, meta=meta,
                           journals=load_journals())


@app.route("/job/<job>/reject", methods=["POST"])
def reject_edits(job):
    """按用户勾选的「保留原文」重渲染：其余改动保持不变，可反复提交。"""
    d = job_dir(job)
    meta = read_meta(d)
    if not meta.get("docx"):
        abort(404)
    ids = [x for x in request.form.getlist("rid") if x and len(x) < 200]

    def prepare():
        write_json_atomic(d / "reject.json", {"ids": ids})

    try:
        stage_render(d, meta["journal"], meta.get("original", "稿件.docx"),
                     prepare=prepare)
    except RenderBusy as e:      # 忙碌态先于 ConvertError：不能落笔写现场（F11）
        return busy_page(e)
    except ConvertError as e:
        return fail_page(d, e)
    except Exception:  # noqa: BLE001
        return fail_page(d, ConvertError(*explain(""), detail=traceback.format_exc()[-1500:]))
    return redirect(url_for("result", job=job))


@app.route("/job/<job>/retry", methods=["POST"])
def retry(job):
    d = job_dir(job)
    meta = read_meta(d)
    jid = request.form.get("journal") or meta.get("journal")
    # 换刊目标先过 JID_RE：不合法的 id 会被拼进渲染路径、原样写进 meta（F47）
    if not (jid and JID_RE.match(jid) and (RULES_DIR / f"{jid}.json").exists()):
        abort(400)
    if not (d / "manuscript.docx").exists():
        abort(404)
    if not meta.get("docx"):
        # 还没有成品：换刊只是改后续渲染的目标，改完回补录页
        meta["journal"] = jid
        meta["journal_name"] = (journal_by_id(jid) or {}).get("name", jid)
        if not isinstance(read_json_safe(d / "manuscript.ir.json"), dict):
            # IR 不存在或已截断：去补录页再渲染只会把死循环走下去（F12）。
            # 先按新刊重跑解析；再失败会落到 failed 页的「重新上传」出路。
            try:
                stage_parse(d, jid, meta.get("original", "稿件.docx"))
            except RenderBusy as e:
                return busy_page(e)
            except ConvertError as e:
                return fail_page(d, e)
            except Exception:  # noqa: BLE001
                return fail_page(d, ConvertError(*explain(""),
                                                 detail=traceback.format_exc()[-1500:]))
            meta["stage"] = "details"
        write_meta(d, meta)
        return redirect(url_for("details", job=job))
    try:
        # 换刊字段由 stage_render 成功后随 meta 一起改写；提前写入的话，
        # 一次失败的换刊会让 meta 指向新刊、成品还是旧刊（F36）
        stage_render(d, jid, meta.get("original", "稿件.docx"))
    except RenderBusy as e:      # 忙碌态先于 ConvertError：不能落笔写现场（F11）
        return busy_page(e)
    except ConvertError as e:
        return fail_page(d, e)
    except Exception:  # noqa: BLE001  与 details / reject 同款兜底：非 ConvertError
        return fail_page(d, ConvertError(*explain(""),   # 也不许裸 500、不写现场（F35）
                                         detail=traceback.format_exc()[-1500:]))
    return redirect(url_for("result", job=job))


@app.route("/job/<job>/diff")
def diff(job):
    d = job_dir(job)
    meta = read_meta(d)
    a_id = meta.get("journal")
    b_id = request.args.get("to") or ""
    ra, rb = rules_of(a_id), rules_of(b_id)
    if not ra or not rb:
        abort(404)

    def g(r, *path, fmt=None):
        cur = r
        for p in path:
            if not isinstance(cur, dict) or p not in cur:
                return "—"
            cur = cur[p]
        try:
            return fmt(cur) if fmt else str(cur)
        except Exception:  # noqa: BLE001
            return str(cur)

    def limit_txt(r, key):
        """该刊没声明上限时（如数量经济技术经济研究不设总字数）如实写"未声明"。"""
        lim = (r.get("limits") or {}).get(key) or {}
        if lim.get("value") is None:
            return "未声明"
        return (f'{lim["value"]:,}（{lim.get("metric", "?")} 口径）' if key == "total"
                else f'{lim["value"]} 字')

    pct = lambda v: f"{v:g}pt"  # noqa: E731
    NOTE_KIND = {"general": "注", "source": "来源", "data_source": "数据", "cell": "①②"}

    def border_txt(r):
        b = r.get("tables", {}).get("border_pt") or {}
        if not isinstance(b, dict):
            return str(b)
        lab = {"top": "顶", "header_bottom": "表头下", "bottom": "底"}
        return " / ".join(f"{lab.get(k, k)} {v:g}" for k, v in b.items()) + " pt"

    def note_txt(r):
        return "→".join(NOTE_KIND.get(x["kind"], x["kind"]) for x in r["tables"]["note_layers"])

    def line_txt(r):
        """行距：倍数行距如实展示倍数，同时给出折算磅值，两刊才可比。"""
        ls = ((r.get("typography") or {}).get("body") or {}).get("line_spacing") or {}
        if ls.get("mode") == "multiple" and ls.get("value"):
            eq = ls.get("equivalent_exact_pt")
            return f'{ls["value"]:g} 倍' + (f'（≈{eq:g}pt）' if eq else "")
        v = ls.get("value_pt")
        return f"{v:g}pt" if isinstance(v, (int, float)) else "—"

    rows = [
        ("纸张", g(ra, "page", "size"), g(rb, "page", "size")),
        ("页边距（上/下/左/右 mm）",
         g(ra, "page", "margin_mm", fmt=lambda m: "/".join(f"{m[k]:g}" for k in ("top", "bottom", "left", "right"))),
         g(rb, "page", "margin_mm", fmt=lambda m: "/".join(f"{m[k]:g}" for k in ("top", "bottom", "left", "right")))),
        ("正文字体", g(ra, "typography", "body", "font_cn"), g(rb, "typography", "body", "font_cn")),
        ("正文字号", g(ra, "typography", "body", "size_pt", fmt=pct),
         g(rb, "typography", "body", "size_pt", fmt=pct)),
        ("行距", line_txt(ra), line_txt(rb)),
        ("标题字号", g(ra, "typography", "title", "size_pt", fmt=pct),
         g(rb, "typography", "title", "size_pt", fmt=pct)),
        ("摘要字号", g(ra, "typography", "abstract", "size_pt", fmt=pct),
         g(rb, "typography", "abstract", "size_pt", fmt=pct)),
        ("标题层级", " / ".join(ra["headings"]["levels"]), " / ".join(rb["headings"]["levels"])),
        ("引用体系",
         "编号制" if ra["citation"]["system"] == "numbered" else "作者年制",
         "编号制" if rb["citation"]["system"] == "numbered" else "作者年制"),
        ("总字数上限", limit_txt(ra, "total"), limit_txt(rb, "total")),
        ("摘要上限", limit_txt(ra, "abstract"), limit_txt(rb, "abstract")),
        ("表格线宽", border_txt(ra), border_txt(rb)),
        ("表注层次", note_txt(ra), note_txt(rb)),
    ]
    changes = [{"name": n, "a": a, "b": b, "same": a == b} for n, a, b in rows]
    ja = journal_by_id(a_id) or {"name": a_id}
    jb = journal_by_id(b_id) or {"name": b_id}
    return render_template("diff.html", job=job, changes=changes,
                           a=ja, b=jb, n_diff=sum(1 for c in changes if not c["same"]))


@app.route("/job/<job>/failed")
def failed(job):
    d = job_dir(job)
    p = d / "error.json"
    if not p.exists():
        return redirect(url_for("workspace"))
    e = read_json_safe(p) or {}
    return fail_page(d, ConvertError(e.get("title", "转换没有完成"),
                                     e.get("hint", ""), e.get("detail", "")))


def attachment(path, display_name, fallback, mimetype):
    """附件下载：中文名走 RFC 5987，同时给一个可读的 ASCII 兜底名。"""
    resp = send_file(path, as_attachment=True, download_name=fallback, mimetype=mimetype)
    resp.headers["Content-Disposition"] = (
        f'attachment; filename="{fallback}"; '
        f"filename*=UTF-8''{quote(display_name)}")
    return resp


@app.route("/job/<job>/download/<kind>")
@job_read_locked
def download(job, kind):
    d = job_dir(job)
    meta = read_meta(d)
    stem = meta.get("stem") or "稿件"
    jn = meta.get("journal_name") or meta.get("journal", "")
    if kind in ("docx", "docx-anon"):
        jid = meta.get("journal") or ""
        if not JID_RE.match(jid):
            abort(404)
        anonymous = kind == "docx-anon"
        field = "docx_anon" if anonymous else "docx"
        name = meta.get(field) or ""
        # meta 也可能在硬杀后由人工修复；下载只允许当前刊对应的单一文件名，
        # 不接受绝对路径、../、反斜杠或把署名稿冒充匿名稿的字段值。
        expected = f"manuscript.{jid}{'.anon' if anonymous else ''}.docx"
        if name != expected:
            abort(404)
        p = d / name
        label = "匿名稿" if anonymous else "署名稿"
        if not p.is_file():
            abort(404)
        return attachment(p, f"{stem}-{jn}-{label}.docx",
                          f"manuscript-{meta.get('journal', 'paper')}-{kind}.docx",
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if kind in REPORT_FILES:
        name = REPORT_FILES[kind]
        p = d / name
        if not p.is_file():
            abort(404)
        label = "投稿体检" if kind == "triage" else "审计报告"
        return attachment(p, f"{stem}-{label}-{jn}.html", f"paper-{kind}.html", "text/html")
    abort(404)


@app.route("/job/<job>/<name>")
@job_read_locked
def job_file(job, name):
    """仅用于 iframe 预览两份报告。"""
    if name not in REPORT_FILES.values():
        abort(404)
    d = job_dir(job)
    if not (d / name).is_file():
        abort(404)
    return send_from_directory(d, name)


@app.route("/journal/<jid>/spec")
def journal_spec(jid):
    if not JID_RE.match(jid or ""):      # 同 draft_dir：..\ 穿透先拦下（F47）
        abort(404)
    p = RULES_DIR / f"{jid}.spec.html"
    if not p.exists():
        abort(404)
    return send_from_directory(RULES_DIR, p.name)


@app.route("/journal/<jid>/sample.pdf")
def journal_sample_pdf(jid):
    """样张 PDF 内联预览：查询页 iframe 直出，浏览器原生缩放翻页。

    只服务 build_journal.py 渲染出的既有产物，不做任何转换——预览即交付物。
    """
    if not JID_RE.match(jid or ""):      # 同 spec：..\ 穿透先拦下（F47）
        abort(404)
    p = RULES_DIR / f"{jid}.sample.pdf"
    if not p.exists():
        abort(404)
    return send_from_directory(RULES_DIR, p.name)


@app.route("/journal/<jid>/sample")
def journal_sample_docx(jid):
    """样板 docx 下载：在 Word/WPS 里亲手核对版式（DESIGN §11.2 用途一）。"""
    if not JID_RE.match(jid or ""):
        abort(404)
    p = RULES_DIR / f"{jid}.sample.docx"
    if not p.exists():
        abort(404)
    name = ((rules_of(jid) or {}).get("journal") or {}).get("name_zh") or jid
    return attachment(p, f"{name}-版式样板.docx", f"{jid}-sample.docx",
                      "application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document")


# ---------------------------------------------------------------- 论文格式查询

# 一次最多并排对比的期刊数。三本并排每列还有 ~200px，再宽就要横向滚了。
QUERY_MAX = 3

_ALIGN_LABEL = {"both": "两端对齐", "center": "居中", "left": "左对齐", "right": "右对齐"}
_METRIC_UNIT = {"chars_cn": "字", "words_en": "词", "A": "字符", "B": "字", "C": "字符"}
_POS_LABEL = {"before_body": "正文之前", "after_body": "正文之后"}
_ANON_LABEL = {"authors": "作者署名", "affiliations": "作者单位", "acknowledgement": "致谢",
               "funding_statement": "基金声明", "corresponding_author": "通讯作者"}
_STATEMENT_LABEL = {"conflict_of_interest": "利益冲突声明", "data_availability": "数据可用性声明",
                    "ai_declaration": "AI 使用声明", "copyright_transfer": "版权转让协议"}


def _style_txt(d):
    """把 typography 子字典转成「宋体 10pt · 居中」式的短文案；没有可展示项返回 None。"""
    if not isinstance(d, dict):
        return None
    cn, latin = d.get("font_cn"), d.get("font_latin")
    fonts = [f for f in (cn, latin if latin and latin != cn else None) if f]
    parts = ([" / ".join(fonts)] if fonts else []) \
        + ([f'{d["size_pt"]:g}pt'] if d.get("size_pt") else []) \
        + (["加粗"] if d.get("bold") else []) \
        + ([_ALIGN_LABEL[d["align"]]] if d.get("align") in _ALIGN_LABEL else
           ([str(d["align"])] if d.get("align") else [])) \
        + ([f'标签「{d["label"]}」'] if d.get("label") else []) \
        + ([f'悬挂缩进 {d["hanging_indent_chars"]:g} 字符'] if d.get("hanging_indent_chars") else []) \
        + ([f'首行缩进 {d["first_line_indent_chars"]:g} 字符'] if d.get("first_line_indent_chars") else [])
    return " · ".join(parts) or None


def _fmt_line(ls):
    """行距：与换刊差异页同一口径——倍数行距如实展示倍数并折算磅值。"""
    if not isinstance(ls, dict):
        return None
    if ls.get("mode") == "multiple" and ls.get("value"):
        eq = ls.get("equivalent_exact_pt")
        return f'{ls["value"]:g} 倍行距' + (f'（≈{eq:g}pt）' if eq else "")
    v = ls.get("value_pt")
    return f"{v:g}pt 行距" if isinstance(v, (int, float)) else None


def _fmt_heading_level(d):
    if not isinstance(d, dict) or d.get("inherit"):
        base = "同正文" if isinstance(d, dict) and d.get("inherit") else None
    else:
        base = _style_txt(d)
    if not base:
        return None
    return base + (" · 随正文接排" if d.get("inline_with_body") else "")


def _fmt_limit(lim):
    """单值上限（摘要等）：单位跟口径走，口径认不出就按「字」，不编。"""
    v = lim.get("value") if isinstance(lim, dict) else None
    if v is None:
        return None
    unit = _METRIC_UNIT.get(lim.get("metric"), "字")
    return f"≤ {v:,} {unit}"


def _fmt_total(lim):
    if not isinstance(lim, dict) or lim.get("value") is None:
        return None
    v = lim["value"]
    unit = _METRIC_UNIT.get(lim.get("metric"), "字")
    lo = lim.get("min")
    bits = [f"{lo:,}–{v:,} {unit}" if lo else f"≤ {v:,} {unit}"]
    if lim.get("metric"):
        bits.append(f"口径 {lim['metric']}")
    if lim.get("scope"):
        bits.append(str(lim["scope"]))
    return " · ".join(bits)


def _fmt_count(lim):
    lo, hi = lim.get("min"), lim.get("max")
    if lo and hi:
        return f"{lo}–{hi} 个"
    if hi:
        return f"≤ {hi} 个"
    if lo:
        return f"≥ {lo} 个"
    return None


def _fmt_highlights(lim):
    if not isinstance(lim, dict):
        return None
    lo, hi = lim.get("min"), lim.get("max")
    n = f"{lo}–{hi}" if lo and hi else (f"≤ {hi}" if hi else (f"≥ {lo}" if lo else None))
    parts = [f"{n} 条"] if n else []
    if lim.get("chars_max"):
        parts.append(f"每条 ≤ {lim['chars_max']:,} 字符")
    return " · ".join(parts) or None


def _fmt_paper(p):
    size = p.get("size")
    mm = p.get("size_mm") or {}
    if isinstance(mm, dict) and mm.get("width") and mm.get("height"):
        return f"{size}（{mm['width']:g}×{mm['height']:g} mm）"
    return size


def _fmt_margin(m):
    return "/".join(f"{m[k]:g}" for k in ("top", "bottom", "left", "right")) + " mm"


def _fmt_columns(v):
    return "单栏" if v == 1 else f"{v:g} 栏"


def _fmt_levels(levels):
    return " / ".join(levels) + f"（{len(levels)} 级）" if levels else None


def _fmt_note_layers(layers):
    kind = {"general": "注", "source": "来源", "data_source": "数据", "cell": "①②"}
    return "→".join(kind.get(x.get("kind"), x.get("kind")) for x in layers) or None


def _fmt_border(b):
    """ref 直接指向 tables.border_pt 字典本身。"""
    if not isinstance(b, dict) or not b:
        return None
    lab = {"top": "顶", "header_bottom": "表头下", "bottom": "底"}
    return " / ".join(f"{lab.get(k, k)} {v:g}" for k, v in b.items()) + "pt"


def _fmt_table_style(t):
    if not isinstance(t, dict):
        return None
    style = {"three-line": "三线表"}.get(t.get("style"), t.get("style"))
    if not style:
        return None
    if t.get("side_walls"):
        style += "（带墙线）"
    return style


def _fmt_caption(t):
    """表/图题：位置 · 对齐 · 字体字号 · 编号样式，有一说一。"""
    if not isinstance(t, dict):
        return None
    pos = {"above": "上方", "below": "下方"}.get(t.get("caption_position"))
    parts = ([f"题在{pos}"] if pos else [])
    style = _style_txt({"font_cn": t.get("caption_font_cn") or t.get("font_cn"),
                        "size_pt": t.get("caption_size_pt"),
                        "align": t.get("caption_align")})
    if style:
        parts.append(style)
    if t.get("caption_numbering"):
        parts.append(f'编号「{t["caption_numbering"]}」')
    return " · ".join(parts) or None


def _fmt_equations(eq):
    if not isinstance(eq, dict):
        return None
    parts = (["居中" if eq.get("align") == "center" else _ALIGN_LABEL.get(eq.get("align"))] if eq.get("align") else []) \
        + ([str(eq["numbering"])] if eq.get("numbering") else []) \
        + (["字母斜体"] if eq.get("letter_italic") else []) \
        + (["简单式可行内"] if eq.get("inline_simple") else [])
    return " · ".join(p for p in parts if p) or None


def _fmt_footnote_numbering(n):
    if not isinstance(n, dict) or not (n.get("footnotes") or n.get("footnotes_scope")):
        return None
    scope = {"per_page": "每页重新编号", "continuous": "全文连续编号"}.get(n.get("footnotes_scope"))
    parts = ([f"脚注标「{n['footnotes']}」"] if n.get("footnotes") else []) \
        + ([scope] if scope else [])
    return " · ".join(p for p in parts if p) or None


def _fmt_in_text(it):
    if not isinstance(it, dict):
        return None
    pos = {"inline": "行内", "superscript": "上标"}.get(it.get("position"), it.get("position"))
    parts = ([f"标注位置：{pos}"] if pos else []) \
        + ([f'括号「{it["bracket"]}」'] if it.get("bracket") else [])
    return " · ".join(parts) or None


def _fmt_cite_system(v):
    return "编号制" if v == "numbered" else ("作者年制" if v == "author-year" else v)


def _fmt_sort(v):
    return {"by_appearance_order": "按正文引用顺序", "alphabetical": "按作者字母序",
            "cn-first-by-pinyin-then-en-by-alphabet": "中文按拼音、英文按字母序"}.get(v, v)


def _fmt_lang(v):
    return {"zh": "中文", "en": "英文"}.get(v, v)


def _fmt_examples(v):
    return v[0] if isinstance(v, list) and v and isinstance(v[0], str) else None


def _fmt_author_page(ap):
    if not isinstance(ap, dict):
        return None
    parts = (["独立成页"] if ap.get("separate_page") else []) \
        + (["置于" + _POS_LABEL.get(ap["position"], ap["position"])] if ap.get("position") else []) \
        + ([f'{len(ap["fields"])} 个字段'] if ap.get("fields") else [])
    return " · ".join(parts) or None


def _fmt_anon_blocks(items):
    labs = [_ANON_LABEL.get(x, x) for x in items]
    return f"{len(labs)} 类（{'、'.join(labs)}）" if labs else None


def _fmt_statements(sts):
    if not isinstance(sts, dict) or not sts:
        return None
    bits = []
    for k, v in sts.items():
        req = v.get("required") if isinstance(v, dict) else None
        bits.append(f"{_STATEMENT_LABEL.get(k, k)} · " + ("必备" if req else "未要求"))
    return "；".join(bits)


def _fmt_subtitle(blocks):
    """副标题：独立副标题块优先，其次题名块的 allow_subtitle；两者都没有才算未规定。"""
    hits = [b for b in blocks if isinstance(b, dict) and b.get("type") == "subtitle"]
    if hits:
        req = next((b.get("required") for b in hits if b.get("required") is not None), None)
        return {True: "必备", False: "可选"}.get(req, "未声明")
    if any(isinstance(b, dict) and b.get("type") == "title" and b.get("allow_subtitle")
           for b in blocks):
        return "允许（随题名）"
    return None


def _block_req(btype):
    def fmt(blocks):
        hits = [b for b in blocks if isinstance(b, dict) and b.get("type") == btype]
        if not hits:
            return None
        req = next((b.get("required") for b in hits if b.get("required") is not None), None)
        base = {True: "必备", False: "可选"}.get(req, "未声明")
        langs = sorted({b.get("lang") for b in hits if b.get("lang")})
        return base + ("（中英）" if len(langs) > 1 else "")
    return fmt


def _block_rows(types):
    return [{"key": t, "label": lab, "ref": "blocks", "fmt": _block_req(t),
             "missing": "未设"} for t, lab in types]


# 每一行 = 一个格式项。ref 用 rules_ref 的引用语法（含 `~` 可缺省后缀），
# 页面上每个非空取值都必须能解析回规则集——查询页不许出现指不到出处的「规定」。
QUERY_SECTIONS = [
    {"key": "identity", "title": "刊物信息", "rows": [
        {"key": "language", "label": "刊物语言", "ref": "journal.language", "fmt": _fmt_lang},
        {"key": "sponsor", "label": "主管 · 主办", "ref": "journal.sponsor~"},
        {"key": "cn", "label": "刊号", "ref": "journal.cn_number~"},
        {"key": "issn", "label": "ISSN", "ref": "journal.issn~"},
    ]},
    {"key": "page", "title": "页面设置", "rows": [
        {"key": "paper", "label": "纸张", "ref": "page", "fmt": _fmt_paper},
        {"key": "margin", "label": "页边距（上/下/左/右）", "ref": "page.margin_mm", "fmt": _fmt_margin},
        {"key": "columns", "label": "分栏", "ref": "page.columns", "fmt": _fmt_columns},
        {"key": "binding", "label": "订边", "ref": "page.binding_side~",
         "fmt": lambda v: {"left": "左侧", "right": "右侧"}.get(v, v)},
    ]},
    {"key": "body", "title": "正文排版", "rows": [
        {"key": "font", "label": "字体（中文 / 西文）", "ref": "typography.body", "fmt": _style_txt},
        {"key": "size", "label": "字号", "ref": "typography.body.size_pt", "fmt": lambda v: f"{v:g}pt"},
        {"key": "line", "label": "行距", "ref": "typography.body.line_spacing", "fmt": _fmt_line,
         "note_refs": ["typography.body.measured.note~"]},
        {"key": "indent", "label": "首行缩进", "ref": "typography.body.first_line_indent_chars~",
         "fmt": lambda v: f"{v:g} 字符"},
        {"key": "align", "label": "对齐", "ref": "typography.body.align~",
         "fmt": lambda v: _ALIGN_LABEL.get(v, v)},
    ]},
    {"key": "headings", "title": "标题体系", "rows": [
        {"key": "numbering", "label": "编号样式", "ref": "headings.levels", "fmt": _fmt_levels,
         "note_refs": ["headings.numbering_source~"]},
        {"key": "level1", "label": "一级标题", "ref": "headings.level1~", "fmt": _fmt_heading_level},
        {"key": "level2", "label": "二级标题", "ref": "headings.level2~", "fmt": _fmt_heading_level},
        {"key": "level3", "label": "三级标题", "ref": "headings.level3~", "fmt": _fmt_heading_level},
        {"key": "level4", "label": "四级标题", "ref": "headings.level4~", "fmt": _fmt_heading_level},
    ]},
    {"key": "titlepage", "title": "题名与作者区", "rows": [
        {"key": "doc_title", "label": "文题", "ref": "typography.title~", "fmt": _style_txt},
        {"key": "subtitle", "label": "副标题", "ref": "blocks", "fmt": _fmt_subtitle, "missing": "未设"},
        {"key": "authors", "label": "作者署名", "ref": "typography.authors~", "fmt": _style_txt},
        {"key": "affil", "label": "作者单位", "ref": "typography.affiliations~", "fmt": _style_txt},
        {"key": "title_en", "label": "英文题名", "ref": "typography.title_en~", "fmt": _style_txt},
    ]},
    {"key": "abstract", "title": "摘要与关键词", "rows": [
        {"key": "abs_style", "label": "摘要版式", "ref": "typography.abstract~", "fmt": _style_txt},
        {"key": "kw_style", "label": "关键词版式", "ref": "typography.keywords~", "fmt": _style_txt},
        {"key": "kw_sep", "label": "关键词分隔符", "ref": "blocks:keywords/separator~",
         "fmt": lambda v: f"「{v}」"},
        {"key": "en_abs", "label": "英文摘要标签", "ref": "blocks:abstract/en/label~",
         "fmt": lambda v: f"「{v}」"},
    ]},
    {"key": "citation", "title": "引用与参考文献", "rows": [
        {"key": "system", "label": "引用体系", "ref": "citation.system", "fmt": _fmt_cite_system},
        {"key": "in_text", "label": "正文标注", "ref": "citation.in_text~", "fmt": _fmt_in_text,
         "note_refs": ["citation.in_text.note~"]},
        {"key": "sort", "label": "文献排序", "ref": "citation.reference_list.sort~", "fmt": _fmt_sort},
        {"key": "style_id", "label": "体例代码", "ref": "citation.style_id~"},
        {"key": "ref_style", "label": "文献表版式", "ref": "typography.references~", "fmt": _style_txt},
        {"key": "example", "label": "条目体例示例", "ref": "citation.reference_list.examples~",
         "fmt": _fmt_examples},
    ]},
    {"key": "figure", "title": "图", "rows": [
        {"key": "caption", "label": "图题", "ref": "figures", "fmt": _fmt_caption},
        {"key": "notes", "label": "图注层次", "ref": "figures.note_layers", "fmt": _fmt_note_layers},
        {"key": "border", "label": "图加边框", "ref": "figures.border~",
         "fmt": lambda v: "不加边框" if v is False else ("加边框" if v else None)},
        {"key": "resolution", "label": "分辨率下限", "ref": "figures.resolution_dpi_min~",
         "fmt": lambda v: f"≥ {v:g} dpi"},
    ]},
    {"key": "table", "title": "表", "rows": [
        {"key": "style", "label": "表线", "ref": "tables", "fmt": _fmt_table_style,
         "note_refs": ["tables.border_pt_source~"]},
        {"key": "border", "label": "线宽（顶/表头下/底）", "ref": "tables.border_pt", "fmt": _fmt_border},
        {"key": "cell_size", "label": "表内字号", "ref": "tables.font_size_pt",
         "fmt": lambda v: f"{v:g}pt"},
        {"key": "caption", "label": "表题", "ref": "tables", "fmt": _fmt_caption},
        {"key": "notes", "label": "表注层次", "ref": "tables.note_layers", "fmt": _fmt_note_layers,
         "note_refs": ["tables.note_order_note~"]},
        {"key": "repeat", "label": "跨页续表头", "ref": "tables.repeat_header_on_break~",
         "fmt": lambda v: "续页重复表头" if v else "续页不重复表头"},
    ]},
    {"key": "equation", "title": "公式", "rows": [
        {"key": "eq", "label": "公式体例", "ref": "equations~", "fmt": _fmt_equations,
         "note_refs": ["equations.note~"]},
    ]},
    {"key": "footnote", "title": "脚注", "rows": [
        {"key": "style", "label": "脚注版式", "ref": "typography.footnote~", "fmt": _style_txt},
        {"key": "numbering", "label": "脚注编号", "ref": "numbering", "fmt": _fmt_footnote_numbering,
         "note_refs": ["numbering.rule~"]},
    ]},
    {"key": "limits", "title": "字数与数量", "rows": [
        {"key": "total", "label": "全文篇幅", "ref": "limits.total~", "fmt": _fmt_total,
         "missing": "不设上限", "note_refs": ["limits.total.note~"]},
        {"key": "abstract", "label": "摘要上限", "ref": "limits.abstract~", "fmt": _fmt_limit},
        {"key": "abstract_en", "label": "英文摘要上限", "ref": "limits.abstract_en~", "fmt": _fmt_limit},
        {"key": "keywords", "label": "关键词个数", "ref": "limits.keywords~", "fmt": _fmt_count},
        {"key": "title", "label": "题名字数", "ref": "limits.title_chars_max~",
         "fmt": lambda v: f"≤ {v} 字"},
        {"key": "highlights", "label": "亮点条目", "ref": "limits.highlights~", "fmt": _fmt_highlights},
        {"key": "jel", "label": "JEL 分类号", "ref": "limits.jel~", "fmt": _fmt_count},
        {"key": "vitae", "label": "作者简介", "ref": "limits.author_vitae_words_max~",
         "fmt": lambda v: f"≤ {v:,} 词"},
    ]},
    {"key": "submission", "title": "投稿要求", "rows": [
        {"key": "anon", "label": "匿名要求", "ref": "submission.anonymous~",
         "fmt": lambda v: "需要匿名稿" if v else "不要求匿名"},
        {"key": "anon_scope", "label": "匿名删除范围", "ref": "submission.anonymize_blocks~",
         "fmt": _fmt_anon_blocks},
        {"key": "author_page", "label": "作者信息页", "ref": "submission.author_page~",
         "fmt": _fmt_author_page, "note_refs": ["submission.author_page.note~"]},
        {"key": "file_format", "label": "提交文件格式", "ref": "submission.file_format~"},
        {"key": "statements", "label": "声明要件", "ref": "statements", "fmt": _fmt_statements},
    ]},
    {"key": "blocks", "title": "篇章结构要件", "rows": _block_rows([
        ("title", "题名"), ("subtitle", "副标题"), ("highlights", "亮点条目"),
        ("authors", "作者署名"), ("affiliations", "作者单位"),
        ("corresponding_author", "通讯作者"), ("abstract", "摘要"), ("keywords", "关键词"),
        ("classification_code", "分类号"), ("document_code", "文献标识码"),
        ("article_number", "文章编号"), ("doi", "DOI"),
        ("body", "正文"), ("references", "参考文献"),
        ("appendix", "附录"), ("acknowledgement", "致谢"), ("funding_statement", "基金声明"),
    ])},
    {"key": "source", "title": "规则来源与核验", "rows": [
        {"key": "status", "label": "核验状态", "ref": "verification.status",
         "fmt": lambda v: STATUS_LABEL.get(v, v)},
        {"key": "authority", "label": "版式依据", "ref": "verification.typography_authority~",
         "fmt": lambda v: AUTH_LABEL.get(v, v)},
        {"key": "source", "label": "规则依据", "ref": "verification.source_title~"},
        {"key": "source_url", "label": "官方来源", "ref": "verification.source_url~", "url": True},
        {"key": "verified_at", "label": "核验时间", "ref": "verification.verified_at~"},
        {"key": "verified_by", "label": "核验方式", "ref": "verification.verified_by~"},
        {"key": "samples", "label": "样张交叉核对", "ref": "verification.cross_check.samples~",
         "fmt": lambda v: f"{v} 篇已发表论文"},
    ]},
]


def _qcell(rules, ref, fmt=None):
    """按 rules_ref 引用取一个展示值；查不到返回 None，页面如实写「未规定」。

    格式化函数抛错同样按查不到处理：查询是只读页面，任何一个值的展示问题
    都不该把整页拖成 500——但引用本身仍然指回规则集，错了在测试里现形。
    """
    try:
        v = rules_ref.resolve(rules, ref)
    except KeyError:
        return None
    if v is None:
        return None
    if fmt is None:
        return v if isinstance(v, str) else None
    try:
        return fmt(v)
    except Exception:  # noqa: BLE001
        return None


def _qnotes(rules, refs):
    """行注：把 *_source / 推定说明 这类伴生键拼成小字，取不到就留空。"""
    if not refs:
        return ""
    bits = []
    for ref in refs:
        try:
            v = rules_ref.resolve(rules, ref)
        except KeyError:
            continue
        if isinstance(v, str) and v.strip():
            bits.append(v.strip())
    return "；".join(bits)


def journal_query_view(jid):
    """把一本刊的规则集摊平成查询页的分区视图。

    行的集合是全刊统一的（QUERY_SECTIONS），单本刊缺哪个字段哪格就空——
    缺席如实写「未规定」，不从同类刊猜、也不替期刊补。
    """
    rules = rules_of(jid) or {}
    j = journal_by_id(jid) or {}
    sections = []
    for sec in QUERY_SECTIONS:
        rows = []
        for row in sec["rows"]:
            fmt = row.get("fmt")
            rows.append({
                "key": row["key"],
                "label": row["label"],
                "cells": [_qcell(rules, row["ref"], fmt)],
                "notes": _qnotes(rules, row.get("note_refs")),
                "url": row.get("url", False),
                "missing": row.get("missing", "未规定"),
            })
        sections.append({"key": sec["key"], "title": sec["title"], "rows": rows})
    ver = rules.get("verification") or {}
    cc = ver.get("cross_check") or {}
    return {
        "id": jid,
        "name": j.get("name", jid),
        "en": j.get("en", ""),
        "status": ver.get("status", "unverified"),
        "status_label": STATUS_LABEL.get(ver.get("status"), "未知"),
        "auth_label": j.get("auth_label", "未声明"),
        "auth_class": j.get("auth_class", "a-pub"),
        "spec": (RULES_DIR / f"{jid}.spec.html").exists(),
        "sample": (RULES_DIR / f"{jid}.sample.pdf").exists(),
        "sections": sections,
        "assumptions": [a for a in (rules.get("assumptions") or []) if isinstance(a, str)],
    }


def _mark_diffs(views):
    """对比模式给每行打 same 标记，返回「规定不同」的行数。

    各刊行集合同构（同一份 QUERY_SECTIONS），按位 zip 即可。全空行
    （各刊都未规定）不算「不同」——没规定的东西不构成差异。
    """
    n_diff = 0
    if len(views) > 1:
        for cols in zip(*[v["sections"] for v in views]):
            for rows in zip(*[c["rows"] for c in cols]):
                cells = [tuple(r["cells"]) for r in rows]
                same = len(set(cells)) == 1
                for r in rows:
                    r["same"] = same
                if not same and any(any(c is not None for c in t) for t in cells):
                    n_diff += 1
    return n_diff


@app.route("/query")
def query():
    """论文格式查询：选一本深查，勾到两三本自动进入逐项对比。

    选择状态放在 URL（?journals=a,b,c）：查到的结果可以收藏、发给自己，
    后退键符合直觉。id 先过 JID_RE 再过 rules_of，不合法或不存在的直接忽略。
    """
    sel = []
    for tok in (request.args.get("journals") or "").split(","):
        if len(sel) >= QUERY_MAX:
            break
        tok = tok.strip()
        if tok and tok not in sel and JID_RE.match(tok) and rules_of(tok):
            sel.append(tok)
    views = [journal_query_view(jid) for jid in sel]
    n_diff = _mark_diffs(views)
    for v in views:
        others = [x for x in sel if x != v["id"]]
        v["remove_url"] = "/query?journals=" + ",".join(others) if others else "/query"
    return render_template("query.html", views=views, ids=sel,
                           groups=journal_groups(), journals=load_journals(),
                           n_diff=n_diff, query_max=QUERY_MAX)


# ---------------------------------------------------------------- 错误页

@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code="404", title="页面不存在",
                           hint="链接可能已经过期。转换结果默认只在本机保留 24 小时。"), 404


@app.errorhandler(413)
def too_large(e):
    return render_template("error.html", code="413", title="文件超过 80 MB",
                           hint="请压缩稿件中的图片，或删掉附录里的大图后重新上传。"), 413


@app.errorhandler(500)
def server_error(e):
    return render_template("error.html", code="500", title="服务内部错误",
                           hint="转换过程出错，稿件没有被修改。请回到工作台重试。"), 500


def setup_logging():
    """请求与异常同时写一份到 outputs/server.log：后台窗口被关掉也能回溯。"""
    path = ROOT / "outputs" / "server.log"
    path.parent.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    handler.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)
    # werkzeug 打请求行、未捕获异常也从这里过一遍，进程被强杀后仍能回溯最后做了什么
    wz = logging.getLogger("werkzeug")
    wz.addHandler(handler)
    wz.setLevel(logging.INFO)
    # 返回 logger 本身：调用方拿它去记心跳与退出事件。返回路径会让
    # 「log.info(...)」报 AttributeError，心跳线程直接死掉且不显眼。
    return app.logger


def launch_detached(port, no_browser):
    """把服务交给一个完全不依附控制台的进程，然后本进程立刻返回。

    关键在 DETACHED_PROCESS：只换成 pythonw 还不够——从 cmd 里 start 出来的进程
    仍然继承了那个控制台，窗口一关就会收到 CTRL_CLOSE_EVENT 被一起结束。
    这也是“点完脚本、页面却拒绝连接”的真正原因。
    """
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    target = str(pyw if pyw.exists() else exe)
    cmd = [target, str(Path(__file__).resolve())]
    if port != 5000:
        cmd += ["--port", str(port)]
    if no_browser:
        cmd += ["--no-browser"]

    kwargs = {}
    if os.name == "nt":
        DETACHED_PROCESS = 0x00000008
        CREATE_NEW_PROCESS_GROUP = 0x00000200
        kwargs["creationflags"] = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen(cmd, cwd=str(ROOT), close_fds=True,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL, **kwargs)


def port_in_use(port=5000):
    """Windows 的 SO_REUSEADDR 允许两个进程同时绑同一端口，请求会被随机分到新旧代码，
    表现成"改了不生效"。启动前先探一下，宁愿拒绝启动也不要跑出两个实例。"""
    import socket
    with socket.socket() as s:
        s.settimeout(0.4)
        return s.connect_ex(("127.0.0.1", port)) == 0


_KEEP = []          # 持有文件句柄，防止 faulthandler 写入目标被回收


def install_lifecycle_hooks(log):
    """让"进程为什么消失"留下证据。

    服务被外部 taskkill 或被作业对象回收时不会抛异常，日志只会停在一半，事后无法
    区分"自己退出""收到信号""被强杀"。这里装三样东西：
      1. faulthandler —— 原生崩溃（栈溢出、段错误）也能写出 C 层堆栈
      2. atexit + 信号处理 —— 正常退出与收到终止信号各留一行
      3. 心跳线程 —— 每 30 秒记一行，事后能确定它最后活到哪一秒
    """
    import atexit
    import faulthandler
    import signal

    crash = ROOT / "outputs" / "crash.log"
    crash.parent.mkdir(parents=True, exist_ok=True)
    # 必须持有文件引用，否则被 GC 后 faulthandler 写不进去
    fh = crash.open("a", encoding="utf-8", buffering=1)
    _KEEP.append(fh)
    faulthandler.enable(file=fh, all_threads=True)

    atexit.register(lambda: log.info("进程正常退出 pid=%s", os.getpid()))

    def on_signal(signum, _frame):
        log.info("收到终止信号 %s，进程即将退出 pid=%s", signum, os.getpid())
        raise SystemExit(0)

    for name in ("SIGTERM", "SIGINT", "SIGBREAK", "SIGHUP"):
        sig = getattr(signal, name, None)
        if sig is not None:
            try:
                signal.signal(sig, on_signal)
            except (ValueError, OSError):
                pass

    def beat():
        while True:
            time.sleep(30)
            log.info("心跳 pid=%s", os.getpid())

    threading.Thread(target=beat, daemon=True).start()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="格式工坊 · 本地服务")
    ap.add_argument("--port", type=int, default=int(os.environ.get("PAPER_PORT", "5000")),
                    help="监听端口，默认 5000，也可用环境变量 PAPER_PORT 指定")
    ap.add_argument("--no-browser", action="store_true", help="启动后不自动打开浏览器")
    ap.add_argument("--detach", action="store_true",
                    help="以脱离控制台的方式在后台启动（关掉窗口也不会停）")
    args = ap.parse_args()

    url = f"http://127.0.0.1:{args.port}"

    if args.detach:
        if port_in_use(args.port):
            print(f"服务已经在运行：{url}")
            raise SystemExit(0)
        launch_detached(args.port, args.no_browser)
        for _ in range(20):                     # 最多等 10 秒确认端口起来
            time.sleep(0.5)
            if port_in_use(args.port):
                print(f"已在后台启动，窗口可以关闭：{url}")
                raise SystemExit(0)
        print("后台启动失败：端口没有起来，请查看 outputs/server.log")
        raise SystemExit(1)

    if port_in_use(args.port):
        print(f"端口 {args.port} 已被占用：已经有一个实例在跑。")
        print(f"直接打开 {url} 即可；要重启就先运行 停止.bat，再运行 启动.bat。")
        raise SystemExit(1)

    WORK.mkdir(parents=True, exist_ok=True)
    sweep()
    log = setup_logging()
    install_lifecycle_hooks(log)
    log_path = ROOT / "outputs" / "server.log"
    app.logger.info(f"服务启动 {url} pid={os.getpid()} 日志={log_path.name}")
    print("格式工坊 · 本地服务已启动 → " + url)
    print(f"日志文件：{log_path.relative_to(ROOT)}（超过 1 MB 自动轮转）")
    print("关闭这个窗口即可停止服务；或在浏览器里按 Ctrl+C。")
    if not args.no_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True,
            use_reloader=False)
