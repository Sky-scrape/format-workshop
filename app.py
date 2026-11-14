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
import importlib
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

sys.path.insert(0, str(ROOT / "tools"))
import app_paths        # noqa: E402  可写数据目录（只读安装时回退到用户目录）

# 程序目录可写（绿色包/开发版）时数据仍留在程序目录；装到 Program Files
# 这类只读位置时回退到 %LOCALAPPDATA%\格式工坊，普通用户也能跑。
def data_root():
    """跟随 ROOT 动态解析：测试会 monkeypatch ROOT，不能钉死在 import 时。"""
    return app_paths.data_root(ROOT)


DATA_ROOT = data_root()

RULES_DIR = ROOT / "rules" / "journal"
DRAFT_DIR = RULES_DIR / "_drafts"
WORK = DATA_ROOT / "outputs" / "web"
JOB_RE = re.compile(r"^[0-9a-f]{12}$")
JID_RE = re.compile(r"^[a-z0-9_]{2,24}$")
JOB_TTL = 24 * 3600
STEP_TIMEOUT = 300

import ai_assist        # noqa: E402  AI 配置与调用
import atomic_io        # noqa: E402  原子写（小配置文件）
import author_profile   # noqa: E402  本机作者档案（默认关闭，可一键清除）
import docx_meta        # noqa: E402  DOCX 元数据隐私闸门
import extract_journal  # noqa: E402  稿约 PDF → 规则集草稿
import jobstore         # noqa: E402  任务目录读写原语（2026-11-13 拆分）
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
REPORT_FILES = {"triage": "triage.html", "audit": "audit.html",
                "preflight": "preflight.html"}

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

# 补录字段服务端长度上限：表单输入只在本机，但 80MB 的 POST 体不该原样
# 落盘——超长的「英文摘要」注入进成品只会以不可解释的方式在渲染层失败。
# 一篇论文的英文摘要远用不到 20k 字符，正常输入永不触界。
DETAIL_FIELD_MAX = 20000
# 驳回列表同理：审计条目到不了这个量级，条数与单条长度都设闸（F1-5）。
REJECT_IDS_MAX = 2000


# ---------------------------------------------------------------- 数据加载

def rules_of(jid):
    if not JID_RE.match(jid or ""):      # 不合法的 id 不参与路径拼接（F47，同 draft_dir）
        return None
    p = RULES_DIR / f"{jid}.json"
    if not p.exists():
        return None
    # 坏 JSON（硬杀留下的截断 / 手改坏）只让这本刊不可用，与列表侧
    # _load_journals_uncached「坏文件只跳过」同一口径，别把整页拖成 500。
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):        # 与 read_json_safe 同款异常集
        return None


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
                    x.get("marker") or {"general": "注", "source": "来源",
                                        "data_source": "数据"}.get(x["kind"], x["kind"])
                    for x in d["tables"]["note_layers"]),
                "words": (d.get("limits", {}).get("total") or {}).get("value"),
                "abstract": _abstract_card(d.get("limits", {}).get("abstract")),
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


_LAZY_IMPORT_LOCK = threading.Lock()


def _lazy_module(name):
    """tools/ 下的重模块按需加载一次。

    sys.path 已在模块顶部挂好 tools/，这里用标准 import 而不是
    spec_from_file_location 再插一次路径：后者并发首调时每个线程各自 exec
    一份模块对象，虽然后面会覆盖 sys.modules 同名项，但先拿到的调用方仍
    攥着旧实例——运行期闸与测试夹具本该共用同一份判定（见 text_guard 的
    模块头），多实例就是两份实现。双检锁只让一个线程真正执行 import。
    """
    mod = sys.modules.get(name)
    if mod is not None:
        return mod
    with _LAZY_IMPORT_LOCK:
        return importlib.import_module(name)


_cr = None


def compliance():
    """按需加载体检模块，复用它的判定逻辑，避免 Web 层再写一遍。"""
    global _cr
    if _cr is None:
        _cr = _lazy_module("compliance_report")
    return _cr


_tg = None


def text_guard():
    """按需加载文字白名单守卫：运行期与测试共用同一份判定（tests/conftest.py
    转出的也是这个模块），避免「测试守得住、上线没人守」的两份实现。"""
    global _tg
    if _tg is None:
        _tg = _lazy_module("text_guard")
    return _tg


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


# 任务目录的读写原语在 tools/jobstore.py（2026-11-13 拆分）：与 Flask 解耦，
# 坏文件兜底与原子写的口径只有一份。写入口在 app 层薄包一层，把
# replace_with_retry 按调用时的模块属性传下去——测试 monkeypatch 它来观察
# 临时名，默认绑定在 jobstore 里就看不见了。
from jobstore import (  # noqa: E402,F401  （公开名，供路由与测试 import）
    read_json, read_json_safe, read_meta, replace_with_retry,
)


def write_json_atomic(path, data):
    """JSON 原子写；替换实现取模块级 replace_with_retry（可被测试替换）。"""
    return jobstore.write_json_atomic(path, data, replace=replace_with_retry)


def write_meta(d, meta):
    """meta.json 原子写（F34），并支持同任务的并发写者。"""
    write_json_atomic(d / "meta.json", meta)


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


def conversion_running():
    """此刻是否有任务真的在处理：解析/排版线程持着任务锁（D-05）。

    stage_parse / stage_render 全程持锁（含编排 Office 的子进程），锁被
    持有 = 转换正在跑；锁还在 _JOB_LOCKS 里但没人持（上一轮结束后的残留）
    不算忙。桌面壳的关窗确认用：转换线程会被 os._exit 连同成果一起硬杀，
    关窗前必须先问引擎忙不忙。journal:{jid} 的锁也一并计入——新建刊的
    AI 抽取同样是几分钟量级的在途工作。
    """
    with _JOB_LOCKS_GUARD:
        locks = list(_JOB_LOCKS.values())
    for lock in locks:
        if not lock.acquire(blocking=False):
            return True
        lock.release()
    return False


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
    _run_preflight(d, jid)
    return ir_data


def _run_preflight(d, jid):
    """输入稿版式体检（只读）。产物 preflight.json + preflight.html。

    体检失败不挡排版：它回答的是「原稿哪里不合本刊要求」，缺了这份附加
    信息的代价远小于整次上传失败。但必须留痕——把失败写进 preflight.json
    的 error，详情页据此显示「未完成」，后台日志里也能回溯原因。
    """
    out_json, out_html = d / "preflight.json", d / "preflight.html"
    try:
        run(["tools/preflight.py", d / "manuscript.docx",
             RULES_DIR / f"{jid}.json", "-o", out_json, "--html", out_html])
    except Exception as e:  # noqa: BLE001 —— 体检是附加信息，不能拖垮上传
        app.logger.warning("版式体检未完成（任务 %s）：%s", d.name, e)
        detail = getattr(e, "detail", "") or str(e) or type(e).__name__
        try:
            write_json_atomic(out_json, {"error": _scrub_local_paths(detail)[-1500:]})
            # 上一轮留下的旧报告必须清掉，否则页面会把旧结论当成这次的结果。
            out_html.unlink(missing_ok=True)
        except OSError:
            app.logger.warning("版式体检失败现场暂未写入：%s", out_json)


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


def _cleanup_staging(d):
    """删掉本轮未提交的 .staging 临时产物。调用方必须持任务锁。"""
    for p in d.glob(".staging.*"):
        try:
            p.unlink(missing_ok=True)
        except OSError as e:
            app.logger.warning("临时产物暂未删除：%s（%s）", p, e)


def _guard_product_text(d, src, filled_docx, out_docx, render_ir, audit_json, rules_data):
    """运行时文字白名单闸：成品相对「补录后的稿子」只允许已登记的字符归一。

    基线必须是 filled.docx（DESIGN §10.5）：补录注入本来就插入文字，拿原稿比
    会把注入段落整批判成改字。匿名稿不参与——匿名化是故意删身份文字，不在
    正文零改动的白名单口径里。
    """
    tg = text_guard()
    ir_data = read_json(render_ir, {}) or {}
    block_types = {b.get("idx"): b.get("type", "")
                   for b in ir_data.get("blocks") or [] if isinstance(b, dict)}
    items = (read_json(audit_json, {}) or {}).get("items") or []
    blocks = rules_data.get("blocks") or []
    labels = {b.get("type"): b.get("label") for b in blocks
              if b.get("lang") == "zh"
              and b.get("type") in ("abstract", "keywords") and b.get("label")}
    # 关键词切词条口径与渲染器 split_keywords 同一：分隔符归一是规则集定义的
    # 白名单动作，不按目标分隔符切就会把它误判成改字。
    kw_sep = next((b.get("separator") for b in blocks
                   if b.get("type") == "keywords" and b.get("lang") == "zh"), None)
    baseline = filled_docx if filled_docx.is_file() else src
    violations = tg.whitelist_violations(
        tg.document_texts(baseline), tg.document_texts(out_docx),
        block_types, items, labels=labels, kw_sep=kw_sep)
    if not violations:
        return
    # 不进入提交阶段：临时产物先清掉，旧一轮成品原样可下载（F36）。
    _cleanup_staging(d)
    shown = "；".join(
        f"块 {idx}（{btype or '未知类型'}）：{before} → {after}"
        for idx, btype, before, after in violations[:5])
    more = f"（共 {len(violations)} 处，只列前 5 处）" if len(violations) > 5 else ""
    raise ConvertError(
        "成品未通过文字白名单校验",
        "排版只允许已登记的字符归一（段首/段尾空白、全角字母数字、"
        "题注/标签/编号前缀、关键词分隔符）。"
        "检测到白名单之外的文字改动，已停止交付；请把下方技术细节发给开发者。",
        f"白名单外改动 {len(violations)} 处：{shown}{more}")


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
    audit_json = tmp_docx.with_suffix(".audit.json")
    # 运行时文字白名单闸（README 的「非白名单差异不交付」）：必须在匿名稿
    # 生成之前——匿名稿是故意删文字，不适用正文比对口径。
    _guard_product_text(d, src, filled_docx, tmp_docx, render_ir, audit_json, rules_data)
    # 体检报告按"注入之后"的稿子重跑，否则填完英文摘要仍会显示"缺 6 项"
    tmp_triage = d / ".staging.triage.html"
    run(["tools/compliance_report.py", render_ir, rules, "-o", tmp_triage])

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
    # 驳回有没有真的落上（S4）：reject.json 里的 id 若与审计条目对不上
    # （编码错位、稿子重传后指纹变了、手工拼的 id），渲染器会静默跳过——
    # 用户以为驳回成功实则原改动照做。这里数出未命中数，日志留痕 + meta
    # 带到结果页提示重驳回。
    requested_ids = []
    if reject_file.is_file():
        raw_reject = read_json_safe(reject_file)
        if isinstance(raw_reject, dict):
            requested_ids = [x for x in raw_reject.get("ids", []) if x]
    if requested_ids:
        known = {it.get("id") for it in audit_items if it.get("id")}
        unmatched = [x for x in requested_ids if x not in known]
        if unmatched:
            app.logger.warning(
                "驳回 id 未命中任何审计条目（任务 %s，%d/%d 条）：%s",
                d.name, len(unmatched), len(requested_ids), unmatched[:10])
    else:
        unmatched = []
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
        "rejected_unmatched": len(unmatched),
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
    # stage 不在这里改成 failed：重渲染失败的任务旧一轮已完好落位（F36 回滚
    # 契约要求回滚后 meta 逐字节不变，stage 仍是 rendered、成品可下载），
    # 一律标「失败」会把可用成品藏起来。失败态由 history 按
    # 「error.json 在场且无可用成品」推导（见 history 路由）。
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
    # 通讯作者不是独立输入框：补录页底部的「通讯作者」脚注字段会经
    # apply_details 落成 corresponding_author 块，登记它才算补齐。
    ("corresponding_author", None): "corresponding",
    # 作者简介在补录页是首页脚注字段 author_bio，apply_details 会按规则集
    # 落成文末的 author_note 块（见 tools/apply_details.py）。不登记这条，
    # 统计研究这类要求作者简介的刊会永远停在「缺 author_note」。
    ("author_note", None): "author_bio",
}

# rows_raw 的标签取自 compliance_report.LABEL；作者简介还没中文名，补录页
# 按刊面用词展示，不把内部类型名甩给用户。
BLOCK_LABEL = {"author_note": "作者简介"}


def _rows_raw(d, rules, strict=False):
    """跑一遍体检取 rows_raw。坏 IR / 坏规则只让这一页没有缺件提示，
    不能把补录页整页拖成 500——与 fail_page 的降级口径一致。

    strict=True 专给补录页提交时的阻断判断：这时「查不到」不能当成「没有
    缺件」放行渲染（失败开放），异常交给调用方按失败页处理。展示型调用
    （details_context）仍用宽松口径，页面少一块缺件提示好过整页 500。
    """
    ir = read_json(d / "manuscript.ir.json", {}) or {}
    try:
        return compliance().check(ir, rules).get("rows_raw") or []
    except Exception:  # noqa: BLE001 —— 体检失败不该挡住补录页
        if strict:
            raise
        return []


def _anonymized_types(rules):
    """该刊要求匿名时必须从正文删掉的块类型。

    这些块「缺」是投稿要求的常态，不是作者漏填：统计研究须知明确要求署名、
    作者简介只在投稿系统填写，英文作者同属署名。拿它们阻断排版，会让所有
    匿名刊稿件永远卡在补录页，与规则集自己的 anonymize_blocks 冲突。
    """
    submission = (rules or {}).get("submission") or {}
    if not submission.get("anonymous"):
        return set()
    return {t for t in (submission.get("anonymize_blocks") or [])
            if isinstance(t, str)}


def blocking_details(rows_raw, data, rules=None):
    """会阻断排版的必备要件，带补录页锚点。

    与 blocking_missing 同一判据；返回值多带 anchor（MISSING_TO_FIELD 的
    字段键），补录页据此把缺件横幅变成可点的定位链接。补不了项的 anchor
    为空串，页面按纯文本展示——不能给一个点了没反应的链接。
    """
    cr = compliance()
    anonymized = _anonymized_types(rules)
    out, seen = [], set()
    for t, lang, name, ok, req in rows_raw or []:
        if ok or not req:
            continue
        if t in cr.NON_ACTIONABLE or t in cr.EXTERNAL_BLOCKS or t in anonymized:
            continue
        key = MISSING_TO_FIELD.get((t, lang))
        if key and (data.get(key) or "").strip():
            continue
        label = BLOCK_LABEL.get(t, name)
        if label in seen:
            continue
        seen.add(label)
        out.append({"label": label, "anchor": key or "", "block": t})
    return out


def blocking_missing(rows_raw, data, rules=None):
    """列出会阻断开始排版的「稿件本体必备要件」。

    只看 required 且判缺的块：article_number / document_code 这类由编辑部
    给定、highlights 这类另附文件不派给作者；匿名刊要求删掉的署名类块也不
    算缺件（见 _anonymized_types）；能由补录字段补齐的（英文摘要、英文标题、
    作者简介…），表单里有非空值就算补齐——否则用户填完还得回 Word 改，
    补录页就白做了。
    """
    return [x["label"] for x in blocking_details(rows_raw, data, rules)]


def optional_missing(rows_raw):
    """该刊要求但非必填、稿子里也没有的块：补录页留空即可跳过。"""
    cr = compliance()
    return [name for t, lang, name, ok, req in rows_raw or []
            if not ok and not req
            and t not in cr.NON_ACTIONABLE and t not in cr.EXTERNAL_BLOCKS]


def _author_page_enabled(rules):
    ap = ((rules or {}).get("submission") or {}).get("author_page") or {}
    return bool(ap.get("separate_page"))


def _footnote_fields(rules):
    out = {}
    for f in (rules or {}).get("personal_fields") or []:
        if isinstance(f, dict) and f.get("id") and f.get("placement") == "footnote":
            out[f["id"]] = f.get("label_zh") or f["id"]
    return out


def author_info_handoff(old, new, saved=None):
    """换刊时人员信息的承接提醒（DESIGN §13.3 的已知缺口）。

    作者页上的字段不会重复写进首页脚注；从「要求作者页」的刊换到「不要求」
    的刊时，已填的作者页内容不会自动跟着进成品，只能靠首页脚注/署名行。
    这里只提示，不替用户搬内容、不生成占位符。
    """
    if not old or not new:
        return []
    old_name = (old.get("journal") or {}).get("name_zh") or "原刊"
    new_name = (new.get("journal") or {}).get("name_zh") or "目标刊"
    notes = []
    old_ap, new_ap = _author_page_enabled(old), _author_page_enabled(new)
    if old_ap and not new_ap:
        labels = [f.get("label_zh") or f.get("id")
                  for f in old.get("personal_fields") or []
                  if isinstance(f, dict) and f.get("placement") == "author_page"]
        filled = [f["id"] for f in old.get("personal_fields") or []
                  if isinstance(f, dict) and f.get("placement") == "author_page"
                  and (saved or {}).get(f.get("id") or "")]
        detail = ("已填的 " + "、".join(
            f.get("label_zh") or f["id"] for f in old.get("personal_fields") or []
            if isinstance(f, dict) and f["id"] in filled)) if filled else \
                 ("、".join(labels[:4]) + ("等" if len(labels) > 4 else "") if labels else "作者信息页")
        notes.append(
            f"《{old_name}》要求作者信息单独成页；《{new_name}》不设作者页。"
            f"原作者页上的 {detail} 不会自动写进新成品——"
            f"这些信息只会出现在首页脚注与署名行，请在下方补录页确认通讯作者与基金信息。")
    elif not old_ap and new_ap:
        notes.append(
            f"《{new_name}》要求作者信息单独成页，下方补录页已列出需要填写的字段；"
            f"匿名稿会整页删除，署名稿保留。")
    old_fn, new_fn = _footnote_fields(old), _footnote_fields(new)
    lost = [label for fid, label in old_fn.items() if fid not in new_fn]
    if lost:
        notes.append(
            f"原刊写在首页脚注的 {'、'.join(lost[:4])} 在《{new_name}》规则集里没有对应字段，"
            f"不会自动带过去；如需保留，请在补录页的「首页脚注」里手工填写。")
    return notes


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


# personal_fields[].placement → 补录页分组（标题，一句人话说明）。
# 分组顺序按「会不会进正文」由近到远排：先看会落进稿子的，再看不进稿子的。
PLACEMENT_GROUPS = (
    ("author_page", "作者信息页", "单独成页放在正文之前，匿名稿整页删除"),
    ("byline", "署名行", "写在标题下方的作者与单位行"),
    ("footnote", "首页脚注", "写在首页底部，不占正文篇幅"),
    ("before_references", "参考文献之前", "放在正文末尾、参考文献之前"),
    ("back_matter", "文末", "排在参考文献之后"),
    ("separate_file", "另附文件", "单独成文件随稿提交，不写进正文"),
    ("submission_system", "投稿系统", "在投稿系统表单里逐项填写，不写进稿件"),
    ("title_page", "独立标题页", "只出现在独立标题页，正文稿不含"),
)
# 库内 system 与 submission_system 两种拼写并存（rules/journal-schema.md
# §10.2），语义相同，合并成一组展示。
PLACEMENT_ALIASES = {"system": "submission_system"}
_PLACEMENT_KEYS = {key for key, *_ in PLACEMENT_GROUPS}


def placement_groups(rules):
    """把 personal_fields 按 placement 分组交给模板。

    旧页面把 author_page 之外的一律归到「投稿系统里填」：byline（署名行）、
    footnote（首页脚注）、separate_file（Highlights 这类另附文件）全被说成
    系统字段，用户看不到稿子里真正会多出什么。分组顺序按「会不会进正文」
    由近到远排，先看会落进稿子的，再看不进稿子的。

    没写 placement 或写了未登记的值时不猜「投稿系统」——Highlights 正是这样
    被旧实现说成系统字段的。归入「其他位置」并如实提示按编辑部要求处理，
    既不丢字段，也不替它编位置。
    """
    buckets = {}
    for f in rules.get("personal_fields") or []:
        if not isinstance(f, dict) or not f.get("id"):
            continue
        key = (f.get("placement") or "").strip()
        key = PLACEMENT_ALIASES.get(key, key)
        if key not in _PLACEMENT_KEYS:
            key = "other"
        buckets.setdefault(key, []).append({
            "id": f["id"],
            "label": f.get("label_zh") or f["id"],
            "required": bool(f.get("required")),
            "note": f.get("note") or "",
            "prefill": bool(f.get("prefill_from_paper")),
        })
    groups = [{"key": key, "title": title, "desc": desc,
               "items": buckets[key]}
              for key, title, desc in PLACEMENT_GROUPS if buckets.get(key)]
    if buckets.get("other"):
        groups.append({"key": "other", "title": "其他位置",
                       "desc": "规则集未登记放置位置，请按编辑部要求处理",
                       "items": buckets["other"]})
    return groups


def _preflight_summary(raw):
    """详情页用的体检摘要。坏文件/未生成返回 None——旧任务没有这份报告时
    页面应安静地少一块，而不是整页 500。"""
    if not isinstance(raw, dict):
        return None
    if raw.get("error"):
        return {"error": str(raw["error"])[:400]}
    st = raw.get("stats") or {}
    cov = raw.get("coverage") or {}
    return {
        "auto": st.get("auto", 0), "manual": st.get("manual", 0),
        "ok": st.get("ok", 0), "inherited": st.get("inherited", 0),
        "unchecked": st.get("unchecked", 0),
        "coverage": cov.get("text", ""),
    }


def details_context(d, meta):
    jid = meta.get("journal", "")
    rules = rules_of(jid) or {}
    rows_raw = _rows_raw(d, rules)
    saved = read_json(d / "details.json", {}) or {}
    ir = read_json(d / "manuscript.ir.json", {}) or {}

    missing_keys = []
    for t, lang, name, ok, req in rows_raw:
        if ok or not req:
            continue
        key = MISSING_TO_FIELD.get((t, lang))
        if key:
            missing_keys.append(key)

    # IR 被外部改坏成「合法 JSON 但结构不对」时（blocks 元素非 dict / 缺 type /
    # blocks 是 null），逐项取 b["type"] 会把补录页整页拖成 500。这里只取结构
    # 完好的块，坏块当作不存在——与 _fmt_subtitle/_block_req 的 isinstance 同款。
    blocks = [b for b in (ir.get("blocks") or []) if isinstance(b, dict) and b.get("type")]
    # 同一类型可能有中英两套（分类号），判断"已有"必须带上语言
    present = {(b["type"], "en" if b["type"].startswith("english_") or b.get("lang") == "en" else "zh")
               for b in blocks}
    field_rows = []
    for key, btype, lang, label, hint in DETAIL_FIELDS:
        need = key in missing_keys
        has = (btype, lang) in present
        if need or has or key == "affiliation_zh":
            field_rows.append({"key": key, "label": label, "hint": hint,
                               "need": need, "value": ""})

    fm = [b.get("text", "") for b in blocks if b["type"] == "front_matter"]
    authors = fm[-1] if fm else ""
    stats = ir.get("stats") or {}
    keys = [re.sub(r"^关键词\s*[:：]\s*", "", b.get("text", ""))
            for b in blocks if b["type"] == "keywords"]
    abstract = [b.get("text", "") for b in blocks if b["type"] == "abstract"]

    for f in field_rows:
        f["value"] = saved.get(f["key"], "")
    # 本机作者档案（§10.7）：默认关闭；启用后只补空位，不覆盖本次已填值。
    # 档案里没有论文专属字段（摘要/标题/基金），不会把上一篇的内容带进来。
    profile_vals, _profile_saved_n = author_profile.prefill(saved)
    profile_prefilled = 0
    for f in field_rows:
        if not f["value"] and profile_vals.get(f["key"]):
            f["value"] = profile_vals[f["key"]]
            profile_prefilled += 1
    footer_values = []
    for k, label, hint, value in [(k, label, hint, saved.get(k, ""))
                                  for k, label, hint in FOOTER_FIELDS]:
        if not value and profile_vals.get(k):
            value = profile_vals[k]
            profile_prefilled += 1
        footer_values.append((k, label, hint, value))
    author_values = author_page_fields(rules, saved)
    for f in author_values:
        if not f["value"] and profile_vals.get(f["key"]):
            f["value"] = profile_vals[f["key"]]
            profile_prefilled += 1

    # 引用体系重排只在「规则集给了条目模板 + 稿子确实用编号引用」时才提供
    cite_cfg = rules.get("citation") or {}
    cite_rl = cite_cfg.get("reference_list") or {}
    has_cite = any("[" in (b.get("text") or "") and "]" in (b.get("text") or "")
                   for b in blocks if b["type"] == "body")
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
        "footer_fields": footer_values,
        "author_fields": author_values,
        "profile_enabled": author_profile.load()["enabled"],
        "profile_prefilled": profile_prefilled,
        "placement_groups": placement_groups(rules),
        # 阻断项按已保存的补录值算：填过的英文摘要不该在刷新后又变成「缺件」。
        "blocking": blocking_missing(rows_raw, saved, rules),
        "blocking_items": blocking_details(rows_raw, saved, rules),
        "optional_missing": optional_missing(rows_raw),
        # 版式体检报告由 _stage_parse 在解析后生成；旧任务没有这份文件时
        # 页面不给死链接，也不显示空摘要。体检脚本直接写文件（非原子），
        # 硬杀可能留下半截 JSON——坏文件按缺失处理，别拖垮补录页。
        "preflight": _preflight_summary(read_json_safe(d / "preflight.json")),
        "has_preflight": (d / "preflight.html").is_file(),
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


APP_VERSION = "1.0.2"       # 版本真源：构建脚本经 tools/app_version.py 解析这里（D-10）


@app.route("/healthz")
def healthz():
    """服务身份标记：app_ctl start 靠它认出端口上的是不是自家服务（D-06）。

    端口活≠自家服务：映像名核对认不出「5000 上的另一个 python 应用」，
    /healthz 的应答带身份标记（tools/app_ctl.py 的 HEALTHZ_MARK），对得上
    才算自家的。只回常量——不碰盘、不碰引擎，也不暴露路径之外的信息。
    """
    return {"app": "geshi-gongfang", "version": APP_VERSION}


@app.context_processor
def inject_version():
    """状态栏（桌面壳）与帮助弹窗都要展示版本号。"""
    return {"app_version": APP_VERSION}


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


# ---------------------------------------------------------------- 设置页

def _ai_form():
    """从表单取值。api_key 留空表示沿用已保存的，避免把密钥回显到页面上。"""
    cur = ai_assist.load_config()
    key = (request.form.get("api_key") or "").strip() or str(cur.get("api_key") or "")

    def num(name, default, lo=1, hi=None):
        try:
            v = int(request.form.get(name) or default)
        except (TypeError, ValueError):
            return default
        # 按表单声明的区间钳制（settings.html 的 min/max 只约束浏览器控件）：
        # 超大的 timeout 会让慢端点拖满解析子进程的 STEP_TIMEOUT，整轮转换
        # 失败还误报「文件体积过大」，违背「超时即放弃 AI」的页面承诺（F1-6）。
        v = max(lo, v)
        return min(hi, v) if hi else v

    return {
        "enabled": request.form.get("enabled") == "on",
        "base_url": (request.form.get("base_url") or "").strip(),
        "model": (request.form.get("model") or "").strip(),
        "api_key": key,
        "send_mode": request.form.get("send_mode") or "snippet",
        "timeout": num("timeout", 60, lo=5, hi=300),
        "snippet_chars": num("snippet_chars", 60, lo=20, hi=400),
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
    try:
        shell_cfg = json.loads((data_root() / "config" / "close.json").read_text(encoding="utf-8"))
    except Exception:
        shell_cfg = {}
    prof = author_profile.load()
    return render_template("settings.html", cfg=cfg, error=error, success=success,
                           key_set=bool(str(cfg.get("api_key") or "").strip()),
                           config_path=str(ai_assist.CONFIG_PATH),
                           close_action=shell_cfg.get("action") or "exit",
                           profile_enabled=prof["enabled"],
                           profile_count=len(prof["fields"]),
                           profile_path=str(author_profile.PROFILE_PATH),
                           profile_saved=bool(request.args.get("profile_saved")))


@app.route("/settings/profile", methods=["POST"])
def settings_profile():
    """本机作者档案开关与一键清除（DESIGN §10.7）。

    默认关闭：补录页不勾「记住这些信息」就不会落盘。开启后新任务自动
    预填空位；关闭只停止预填，不清已存字段；清除才删文件。
    """
    action = request.form.get("action")
    if action == "clear":
        author_profile.clear()
    elif action == "enable":
        author_profile.set_enabled(True)
    elif action == "disable":
        author_profile.set_enabled(False)
    else:
        abort(400)
    return redirect(url_for("settings", profile_saved=1))


@app.route("/settings/shell", methods=["POST"])
def settings_shell():
    """桌面壳偏好（关窗语义）。只在意桌面壳里可见，浏览器里提交也无害。"""
    action = request.form.get("close_action", "exit")
    # 原子写：进程被硬杀不留半截 JSON（读侧 read_json_safe 本就按半截兜底）
    atomic_io.write_text(
        data_root() / "config" / "close.json",
        json.dumps({"action": "tray" if action == "tray" else "exit"}))
    return redirect(url_for("settings", shell_saved=1))


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


def _journal_new_ctx(error=None):
    journals = load_journals()
    return dict(ai_ok=ai_assist.available(), cfg=ai_assist.load_config(),
                error=error, ocr_ok=ocr.available(), ocr_engine=ocr.engine_name(),
                drafts=draft_list(), journals=journals,
                existing={j["id"] for j in journals})


@app.route("/journals/new")
def journal_new():
    return render_template("journal_new.html", **_journal_new_ctx())


@app.route("/journals/extract", methods=["POST"])
def journal_extract():
    def back(msg, status=400):
        return render_template("journal_new.html", **_journal_new_ctx(msg)), status

    f = request.files.get("file")
    jid = (request.form.get("jid") or "").strip().lower()
    name = (request.form.get("name") or "").strip()
    baseline = (request.form.get("baseline") or "").strip().lower()
    if baseline and (not JID_RE.match(baseline)
                     or not (RULES_DIR / f"{baseline}.json").exists()):
        baseline = ""                # 选择项被篡改/规则集已删：退回自动基线
    mode = (request.form.get("mode") or "auto").strip().lower()
    if mode not in ("auto", "manual"):
        mode = "auto"
    if not f or not f.filename:
        return back("请选择一份稿约：PDF、Word（.docx）或 TXT 都可以")
    suffix = Path(f.filename).suffix.lower()
    if suffix not in extract_journal.SOURCE_SUFFIXES:
        return back("只收 PDF、Word（.docx）或 TXT。.doc 老格式请先另存为 .docx；"
                    "扫描件请用 PDF/Word，并确认本机 OCR 可用。")
    if not JID_RE.match(jid):
        return back("期刊代号只能用小写字母、数字、下划线，2–24 位，例如 jryj")
    # AI 不可用而用户仍点了「抽取」时，不把整条路堵死：改走手工建库，
    # 读文后生成基线草稿，字段在复核页填。用户明确选 manual 时同理。
    manual = mode == "manual" or not ai_assist.available()
    lock = job_lock(f"journal:{jid}")
    if not lock.acquire(blocking=False):
        return back("这本期刊正在抽取或发布，请稍候再试。", 409)
    try:
        # 持锁后再查一次，避免长时间抽取期间另一请求已完成同名发布。
        if (RULES_DIR / f"{jid}.json").exists():
            return back(f"代号 {jid} 已被占用，换一个")

        d = DRAFT_DIR / jid
        d.mkdir(parents=True, exist_ok=True)
        src = d / f"source{suffix}"
        f.save(src)

        try:
            # 抽取过程的事件（F52 截断降级重试、OCR 丢页/低文字层）必须记进
            # 服务日志——CLI 链路走 stderr，web 链路不传 log 就全部静默，
            # draft.json 里只有结果，查不到「发生过」。
            rules, report, text = extract_journal.extract(
                src, jid=jid, name_zh=name, log=app.logger.warning,
                manual=manual, baseline_id=baseline)
        except Exception as e:  # noqa: BLE001
            app.logger.exception("稿约抽取失败")
            return back(f"读取稿约失败：{e}")

        write_json_atomic(d / "draft.json", rules)
        write_json_atomic(d / "report.json", report)
        # 取文结果单独落一份，供复核时检索；不能用 source.txt 命名——上传的
        # 稿约本身可能就叫 source.txt，会把原件覆盖掉。
        (d / "source_text.txt").write_text(text, encoding="utf-8")
        return redirect(url_for("journal_draft", jid=jid))
    finally:
        lock.release()


def _draft_report(d):
    """读草稿的 report.json 并归一成模板可安全下钻的形状。

    report.json 缺失 / 截断 / 手改坏时（比如只留下 draft.json 的现场），
    journal_draft.html 第 59 行的 report.stats.extracted 会把页面拖成
    jinja2 UndefinedError 500——发布闸前的报错回显（_journal_publish 的
    back()）复用同一模板，同样中招。这里补齐模板用到的每个键，坏报告
    按空报告渲染：错误清单一空，ok_to_publish=False 天然挡住发布。
    """
    report = read_json_safe(d / "report.json")
    if not isinstance(report, dict):
        report = {}
    if not isinstance(report.get("stats"), dict):
        report["stats"] = {}
    if not isinstance(report.get("errors"), list):
        report["errors"] = []
    report.setdefault("ok_to_publish", False)
    return report


def _draft_ctx(jid, rules, report, error=None):
    """复核页的统一上下文。GET、保存失败回显与发布失败回显共用一份，
    避免三处各拼一遍、漏掉新字段（比如手工修正表单）。"""
    ext = (rules.get("verification") or {}).get("extraction") or {}
    if not isinstance(ext, dict):
        ext = {}
    if not isinstance(ext.get("manual"), dict):
        ext["manual"] = {}
    return dict(jid=jid, rules=rules, report=report, ext=ext,
                journal=rules.get("journal") or {},
                evidence=ext.get("used") or {},
                inherited=ext.get("inherited") or [],
                dropped=ext.get("dropped") or [],
                field_groups=extract_journal.editable_fields(rules),
                structural=extract_journal.structural_fields(rules),
                ai_on=ai_assist.available(), error=error)


def _render_draft(jid, rules, report, error=None, status=200):
    return render_template("journal_draft.html",
                           **_draft_ctx(jid, rules, report, error)), status


@app.route("/journals/draft/<jid>")
def journal_draft(jid):
    d = draft_dir(jid)
    rules = read_json(d / "draft.json", {}) or {}
    return _render_draft(jid, rules, _draft_report(d))


@app.route("/journals/draft/<jid>/edit", methods=["POST"])
def journal_draft_edit(jid):
    """把复核页手工修正的字段写回草稿，并重新计算继承披露与校验结果。"""
    if not JID_RE.match(jid or ""):
        abort(404)
    lock = job_lock(f"journal:{jid}")
    if not lock.acquire(blocking=False):
        return render_template(
            "error.html", code="…", title="这本期刊正在处理",
            hint="抽取或发布尚未结束，本次修改没有保存。请稍候再试。"), 409
    try:
        d = draft_dir(jid)
        rules = read_json(d / "draft.json", {}) or {}
        if not rules:
            return render_template(
                "error.html", code="404", title="草稿不存在",
                hint="这份草稿可能已被删除；请返回规则库重新建刊。"), 404
        values = {k[2:]: v for k, v in request.form.items() if k.startswith("f_")}
        structural = None
        if request.form.get("s_structural"):
            # 复选框未勾选时浏览器不提交该键：按 context 里的完整清单补齐
            # False，否则「取消必备」在服务端读不到、等于没保存。
            ctx = extract_journal.structural_fields(rules)
            structural = {
                "anonymous": "s_anon" in request.form,
                "author_page": "s_author_page" in request.form,
                "blocks": {b["index"]: f"s_blk_{b['index']}" in request.form
                           for b in ctx["blocks"]},
                "statements": {s["key"]: f"s_st_{s['key']}" in request.form
                               for s in ctx["statements"]},
            }
        new_rules, report, errors = extract_journal.apply_manual(
            rules, values, structural=structural)
        if errors:
            # 有错就整单不落盘：不能把一半修正写进草稿，复核页显示的
            # 值与磁盘上的值必须是同一份。
            return _render_draft(jid, rules, _draft_report(d),
                                 error="；".join(errors), status=400)
        write_json_atomic(d / "draft.json", new_rules)
        write_json_atomic(d / "report.json", report)
        return redirect(url_for("journal_draft", jid=jid))
    finally:
        lock.release()


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
    report = _draft_report(d)

    def back(msg, status=400):
        return _render_draft(jid, rules, report, error=msg, status=status)

    if not report.get("ok_to_publish"):
        return back("还有阻断性错误，先把错误项修掉")
    if request.form.get("confirm") != "on":
        return back("请先勾选「我已逐项核对」——草稿不能未经核对就进规则库")
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
    # 复制用 os.replace 是「移动」：复制后草稿里的 PDF 会消失。必须在移动
    # 之前记下本轮是否真的生成了 PDF，否则后面的「无 PDF 清理旧预览」会把
    # 刚发布的 PDF 当成旧件删掉。
    built_pdf = (d / f"{jid}.sample.pdf").is_file()
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
    if not built_pdf:
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
            # 解析失败（stage=uploaded 且无 IR）的任务不能给「继续补录」：
            # 补录页渲染得出表单、提交后渲染必败，落到 F12 的「重新上传」
            # 死路。IR 存在但损坏时不误伤——details 路由会重解析自救。
            ir_ok = isinstance(read_json_safe(d / "manuscript.ir.json"), dict)
            stage = m.get("stage") or ("rendered" if m.get("docx") else "details")
            # 失败态现场推导（S3）：error.json 在场且没有可用成品——解析失败
            # 的任务 fail_page 不改 meta.stage（F36 回滚契约不许碰），历史页
            # 若仍按 stage 渲染就会把失败任务照「待补录」摆出来、诱导用户去
            # 把同一错误再踩一遍。已交付任务的重渲染失败不算失败：成品还在、
            # 可下载，标「失败」反而把可用产物藏起来。
            docx_usable = bool(m.get("docx")) and (d / m["docx"]).is_file()
            rows.append({
                "job": d.name,
                "original": m.get("original"),
                "journal": m.get("journal_name") or m.get("journal", ""),
                "created": m.get("created", ""),
                "stage": stage,
                "failed": (d / "error.json").is_file() and not docx_usable,
                "parse_failed": not ir_ok,
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


def remember_details(d, data):
    """保存补录表单状态；与渲染轮次共用任务锁，抢不到返回 False。

    阻断（缺件）与正常提交都要先把用户填的值落盘：缺件被退回时页面要带着
    已填内容重开，不能让人重输一遍。抢锁失败时不落笔——此刻另一轮渲染正在
    写同一份 details.json（F11）。
    """
    lock = job_lock(d.name)
    if not lock.acquire(blocking=False):
        return False
    try:
        write_json_atomic(d / "details.json", data)
        return True
    finally:
        lock.release()


@app.route("/job/<job>/details", methods=["GET", "POST"])
def details(job):
    d = job_dir(job)
    meta = read_meta(d)
    if not meta.get("journal"):
        return redirect(url_for("workspace"))

    # 换刊补录（结果页换到要求更多构件的刊）只在本页显示/本轮渲染用目标刊：
    # meta 仍由 stage_render 成功后改写，失败时旧成品与旧刊信息原样保留。
    requested_jid = request.values.get("journal")
    target_jid = requested_jid or meta["journal"]
    rules = rules_of(target_jid)
    if requested_jid and not rules:
        abort(404)          # 显式换刊目标必须真实存在；旧任务规则被删仍按空规则降级
    rules = rules or {}

    # 换刊补录页不能继续展示旧目标刊的体检结论：预览的是「原稿 vs 目标刊」，
    # 目标一变就必须重跑。抢不到任务锁说明另一轮正在写现场，宁可先不展示，
    # 也不给旧刊口径的结论。
    switched = target_jid != meta.get("journal")
    preflight_current = not switched
    if switched and (d / "manuscript.docx").exists():
        lock = job_lock(d.name)
        if lock.acquire(blocking=False):
            try:
                _run_preflight(d, target_jid)
                preflight_current = True
            finally:
                lock.release()

    def page_context(base_meta):
        ctx = details_context(d, base_meta)
        # 换刊时提示人员信息承接（DESIGN §13.3）：作者页/首页脚注里的内容
        # 不会自动跟着换刊产物走，提前说清楚，别等出稿后才发现少了。
        ctx["switch_notes"] = (author_info_handoff(
            rules_of(meta.get("journal") or "") or {}, rules,
            ctx.get("saved") or {}) if switched else [])
        if not preflight_current:
            ctx["preflight"] = None
            ctx["has_preflight"] = False
        return ctx

    def view_meta(base):
        if target_jid == base.get("journal"):
            return base
        return {**base, "journal": target_jid,
                "journal_name": (journal_by_id(target_jid) or {}).get("name", target_jid)}

    # 文件存在但 JSON 已截断时不能被「存在即有效」骗过；原稿仍在就直接重解析。
    ir_path = d / "manuscript.ir.json"
    if ir_path.exists() and not isinstance(read_json_safe(ir_path), dict):
        try:
            stage_parse(d, target_jid, meta.get("original", "稿件.docx"))
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
        for f in author_page_fields(rules):
            data[f["key"]] = (request.form.get(f["key"]) or "").strip()
        # 服务端长度闸（F1-5）：前端没有 maxlength，80MB 的 POST 体不该
        # 原样落盘再注入成品。超限是请求本身的问题，按 400 拒掉。
        oversized = [k for k, v in data.items() if len(v) > DETAIL_FIELD_MAX]
        if oversized:
            return render_template(
                "error.html", code="400", title="填写内容过长",
                hint=f"{'、'.join(oversized)} 超过 {DETAIL_FIELD_MAX} 字上限，"
                     "请删减后再提交（一篇论文用不到这么长）。"), 400
        citation = request.form.get("citation") == "on"
        # 本机作者档案：用户显式勾选「记住这些信息」才落盘（§10.7 默认
        # 会话级）。只收与论文无关的身份字段，论文专属字段由模块白名单挡掉。
        if request.form.get("remember_profile") == "on":
            author_profile.remember(data)

        # 结构性要件阻断（DESIGN §10.4）：稿件本体缺件时不允许开始排版。
        # 放在长度闸之后——超长输入先按请求错误拒掉，不必跑体检。
        try:
            rows_raw = _rows_raw(d, rules, strict=True)
        except Exception:  # noqa: BLE001 —— 查不到缺件不等于没有缺件，不能失败开放
            return fail_page(d, ConvertError(
                "无法核对稿件必备要件",
                "体检过程出错，未开始排版。请重试；若反复出现，把下方技术细节发给开发者。",
                traceback.format_exc()[-1500:]))
        blocking = blocking_missing(rows_raw, data, rules)
        if blocking:
            # 表单状态必须留下：用户填的英文摘要不能因为正文还缺件就被丢掉。
            if not remember_details(d, data):
                return busy_page(RenderBusy(job))
            page_meta = view_meta(read_meta(d))
            return render_template("details.html", job=job, meta=page_meta,
                                   **page_context(page_meta)), 400

        def prepare():
            # 输入与开关都在持锁后落笔：忙碌请求必须真正零副作用（F11）。
            write_json_atomic(d / "details.json", data)
            current = read_meta(d)
            # 引用体系重排是唯一会改写正文/条目文字的开关，默认关，需显式勾选。
            current["citation"] = citation
            write_meta(d, current)

        try:
            stage_render(d, target_jid, meta.get("original", "稿件.docx"),
                         prepare=prepare)
        except RenderBusy as e:      # 忙碌态先于 ConvertError：不能落笔写现场（F11）
            return busy_page(e)
        except ConvertError as e:
            return fail_page(d, e)
        except Exception:  # noqa: BLE001
            return fail_page(d, ConvertError(*explain(""),
                                             detail=traceback.format_exc()[-1500:]))
        # 换刊后输入稿体检的目标刊变了，旧报告不能再挂着；同刊重排沿用原报告。
        if target_jid != meta.get("journal"):
            _run_preflight(d, target_jid)
        return redirect(url_for("result", job=job))

    page_meta = view_meta(meta)
    return render_template("details.html", job=job, meta=page_meta,
                           **page_context(page_meta))


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
                           journals=load_journals(),
                           # 旧任务没有版式体检：第三个 tab 不渲染，也不报错
                           has_preflight=(d / "preflight.html").is_file())


@app.route("/job/<job>/reject", methods=["POST"])
def reject_edits(job):
    """按用户勾选的「保留原文」重渲染：其余改动保持不变，可反复提交。"""
    d = job_dir(job)
    meta = read_meta(d)
    if not meta.get("docx"):
        abort(404)
    ids = [x for x in request.form.getlist("rid") if x and len(x) < 200]
    if len(ids) > REJECT_IDS_MAX:    # 条数闸（F1-5）：审计条目到不了这个量级
        abort(400)

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
        else:
            # IR 还能用：目标刊已变，输入稿体检必须按新刊重跑，否则结果页会
            # 挂着旧刊口径的结论（报告标题写的是新目标刊，内容却是旧刊判的）。
            _run_preflight(d, jid)
        write_meta(d, meta)
        return redirect(url_for("details", job=job))

    # 换刊重排同样受结构性要件阻断：目标刊要的构件稿件里没有、补录也没填，
    # 直接渲染就会交付缺件稿。先落到目标刊的补录页，让用户在表单里补齐。
    rules = rules_of(jid) or {}
    try:
        rows_raw = _rows_raw(d, rules, strict=True)
    except Exception:  # noqa: BLE001 —— 查不到缺件不等于没有缺件，不能失败开放
        return fail_page(d, ConvertError(
            "无法核对稿件必备要件",
            "换刊前的体检过程出错，未开始排版。请重试；若反复出现，把下方技术细节发给开发者。",
            traceback.format_exc()[-1500:]))
    saved = read_json(d / "details.json", {}) or {}
    if blocking_missing(rows_raw, saved, rules):
        return redirect(url_for("details", job=job, journal=jid))
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
    if jid != meta.get("journal"):
        _run_preflight(d, jid)
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
        """该刊没声明上限时（如数量经济技术经济研究不设总字数）如实写"未声明"。

        区间/仅下限型（如 tjyj 摘要 300–500、zgrz 摘要 ≥600）也照规则集给，
        不因缺单一 value 就当「未声明」。
        """
        lim = (r.get("limits") or {}).get(key) or {}
        v, lo, hi = lim.get("value"), lim.get("min"), lim.get("max")
        if v is not None:
            return (f'{v:,}（{lim.get("metric", "?")} 口径）' if key == "total"
                    else f"{v} 字")
        if lo and hi:
            return f"{lo}–{hi} 字"
        if hi:
            return f"≤ {hi} 字"
        if lo:
            return f"≥ {lo} 字"
        return "未声明"

    pct = lambda v: f"{v:g}pt"  # noqa: E731
    NOTE_KIND = {"general": "注", "source": "来源", "data_source": "数据"}

    def border_txt(r):
        b = (r.get("tables") or {}).get("border_pt") or {}
        if not isinstance(b, dict):
            return str(b)
        lab = {"top": "顶", "header_bottom": "表头下", "bottom": "底"}
        return " / ".join(f"{lab.get(k, k)} {v:g}" for k, v in b.items()) + " pt"

    def note_txt(r):
        # 规则集缺 tables/note_layers（或元素不是 dict）时显「—」，不 500：
        # 与本页其余行走同一 g()/.get 口径（查询页「任何值的展示问题都不该
        # 把整页拖成 500」，F1-3）
        layers = (r.get("tables") or {}).get("note_layers")
        if not isinstance(layers, list):
            return "—"
        parts = []
        for x in layers:
            if not isinstance(x, dict):
                continue
            parts.append(x.get("marker") or NOTE_KIND.get(x.get("kind"), x.get("kind") or ""))
        return "→".join(p for p in parts if p)

    def line_txt(r):
        """行距：倍数行距如实展示倍数，同时给出折算磅值，两刊才可比。"""
        body = r.get("typography")
        body = body.get("body") if isinstance(body, dict) else None
        ls = (body or {}).get("line_spacing") or {}
        if not isinstance(ls, dict):
            return "—"
        if ls.get("mode") == "multiple" and ls.get("value"):
            eq = ls.get("equivalent_exact_pt")
            return f'{ls["value"]:g} 倍' + (f'（≈{eq:g}pt）' if eq else "")
        v = ls.get("value_pt")
        return f"{v:g}pt" if isinstance(v, (int, float)) else "—"

    def cite_txt(r):
        # 缺 citation.system 显「—」；有值才判编号/作者年（F1-3）
        s = g(r, "citation", "system")
        return "—" if s == "—" else ("编号制" if s == "numbered" else "作者年制")

    def levels_txt(r):
        # 缺 headings.levels 显「—」；非列表不给 join 拆字符（F1-3）
        lv = g(r, "headings", "levels")
        if lv == "—":
            return "—"
        return " / ".join(lv) if isinstance(lv, list) else str(lv)

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
        ("标题层级", levels_txt(ra), levels_txt(rb)),
        ("引用体系", cite_txt(ra), cite_txt(rb)),
        ("总字数上限", limit_txt(ra, "total"), limit_txt(rb, "total")),
        ("摘要上限", limit_txt(ra, "abstract"), limit_txt(rb, "abstract")),
        ("表格线宽", border_txt(ra), border_txt(rb)),
        ("表注层次", note_txt(ra), note_txt(rb)),
    ]
    changes = [{"name": n, "a": a, "b": b, "same": a == b} for n, a, b in rows]
    ja = journal_by_id(a_id) or {"name": a_id}
    jb = journal_by_id(b_id) or {"name": b_id}
    return render_template("diff.html", job=job, changes=changes,
                           a=ja, b=jb, n_diff=sum(1 for c in changes if not c["same"]),
                           handoff=author_info_handoff(ra, rb))


@app.route("/job/<job>/failed")
@job_read_locked
def failed(job):
    """失败页回访。fail_page 会重写 error.json 与 meta（带写副作用的 GET），
    必须与渲染提交互斥：渲染进行中访问会把旧失败写回、盖掉已提交的新 meta
    （error.json 复活 / meta 回退），与 download/job_file 同款闸（F1-2）。"""
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
        label = {"triage": "投稿体检", "audit": "审计报告",
                 "preflight": "版式体检"}[kind]
        return attachment(p, f"{stem}-{label}-{jn}.html", f"paper-{kind}.html", "text/html")
    abort(404)


@app.route("/job/<job>/<name>")
@job_read_locked
def job_file(job, name):
    """仅用于 iframe 预览报告（投稿体检 / 审计 / 输入稿体检）。"""
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
    """字数上限（摘要等）：单一上限 / 区间 / 仅下限都照规则集给（与 render_spec
    的三分支同口径，§17.1），单位跟口径走，口径认不出就按「字」，不编。"""
    if not isinstance(lim, dict):
        return None
    unit = _METRIC_UNIT.get(lim.get("metric"), "字")
    v = lim.get("value")
    if v is not None:
        return f"≤ {v:,} {unit}"
    lo, hi = lim.get("min"), lim.get("max")
    if lo and hi:
        return f"{lo:,}–{hi:,} {unit}"
    if hi:
        return f"≤ {hi:,} {unit}"
    if lo:
        return f"≥ {lo:,} {unit}"
    return None


def _abstract_card(lim):
    """列表卡片的「摘要」栏：模板会在取值后面接「字」，所以这里给不带单位的
    数/区间；区间与仅下限型（如 tjyj 300–500、zgrz ≥600）不再掉成「未规定」。"""
    if not isinstance(lim, dict):
        return None
    v = lim.get("value")
    if v is not None:
        return v
    lo, hi = lim.get("min"), lim.get("max")
    if lo and hi:
        return f"{lo}–{hi}"
    if hi:
        return f"≤ {hi}"
    if lo:
        return f"≥ {lo}"
    return None


def _fmt_total(limits):
    """总字数：上限在 limits.total，下限还有顶层 limits.total_min 的兼容写法。

    只读 total.min 会让 cmjj 这类「不少于 1.8 万字」在下限登记到顶层时整行
    消失；只读 total.value 同样会把只写下限的刊误报成「未规定」。
    """
    if not isinstance(limits, dict):
        return None
    lim = limits.get("total") if isinstance(limits.get("total"), dict) else {}
    v = lim.get("value")
    unit = _METRIC_UNIT.get(lim.get("metric"), "字")
    lo = lim.get("min") or lim.get("total_min") or limits.get("total_min")
    if v is None and lo is None:
        return None
    bits = [f"{lo:,}–{v:,} {unit}" if (lo and v is not None)
            else (f"≤ {v:,} {unit}" if v is not None else f"≥ {lo:,} {unit}")]
    if lim.get("metric"):
        bits.append(f"口径 {lim['metric']}")
    if lim.get("scope"):
        bits.append(str(lim["scope"]))
    return " · ".join(bits)


def _fmt_title_max(limits):
    """题名字数上限：多数刊写 title_chars_max，cjwt 写 title_max_chars。"""
    if not isinstance(limits, dict):
        return None
    v = limits.get("title_chars_max") or limits.get("title_max_chars")
    return f"≤ {v} 字" if v else None


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
    kind = {"general": "注", "source": "来源", "data_source": "数据"}
    return "→".join(x.get("marker") or kind.get(x.get("kind"), x.get("kind"))
                    for x in layers) or None


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
        {"key": "total", "label": "全文篇幅", "ref": "limits", "fmt": _fmt_total,
         "missing": "不设上限", "note_refs": ["limits.total.note~"]},
        {"key": "abstract", "label": "摘要上限", "ref": "limits.abstract~", "fmt": _fmt_limit},
        {"key": "abstract_en", "label": "英文摘要上限", "ref": "limits.abstract_en~", "fmt": _fmt_limit},
        {"key": "keywords", "label": "关键词个数", "ref": "limits.keywords~", "fmt": _fmt_count},
        {"key": "title", "label": "题名字数", "ref": "limits", "fmt": _fmt_title_max},
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
        {"key": "file_format", "label": "提交文件格式", "ref": "submission.file_format~",
         "fmt": lambda v: " / ".join(v) if isinstance(v, list) else str(v),
         "note_refs": ["submission.file_format_note~"]},
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
    return render_template("error.html", code="500", title="软件内部错误",
                           hint="转换过程出错，稿件没有被修改。请回到工作台重试。"), 500


def setup_logging(name="server.log"):
    """请求与异常同时写一份到 outputs/<name>：后台窗口被关掉也能回溯。

    name 分流（D-07）：server.log 归监听端口的浏览器服务模式，桌面壳写
    desktop.log——两个进程各持一个 RotatingFileHandler 写同一文件时，
    轮转 rename 在 Windows 上会互相踩（WinError 32）。
    """
    path = data_root() / "outputs" / name
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

    crash = data_root() / "outputs" / "crash.log"
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


def init_app(install_hooks=True, log_name="server.log"):
    """引擎初始化的唯一入口：建目录、清过期任务、装日志、装生命周期钩子。

    服务模式 main() 与桌面壳 desktop.engine() 共用这一份——壳曾把前三步
    抄过去、漏掉 install_lifecycle_hooks：faulthandler/crash.log/atexit/
    心跳全没装，壳进程（pythonnet/WebView2 COM、WSGI、编排 Office 的线程
    池同住一屋）原生崩了只剩 pythonw 静默消失，无从取证（D-07）。

    注意：信号处理只有主线程装得上（其他线程 ValueError，被钩子自己吞掉）；
    壳的引擎引导跑在后台线程里，装的是 faulthandler/atexit/心跳三样，这
    三样已覆盖壳形态最想要的崩溃取证。
    """
    WORK.mkdir(parents=True, exist_ok=True)
    sweep()
    log = setup_logging(log_name)
    if install_hooks:
        install_lifecycle_hooks(log)
    return log


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
        print(f"后台启动失败：端口没有起来，请查看 {data_root() / 'outputs' / 'server.log'}")
        raise SystemExit(1)

    if port_in_use(args.port):
        print(f"端口 {args.port} 已被占用：已经有一个实例在跑。")
        print(f"直接打开 {url} 即可；要重启就先运行 停止.bat，再运行 启动.bat。")
        raise SystemExit(1)

    log = init_app()
    log_path = data_root() / "outputs" / "server.log"
    app.logger.info(f"服务启动 {url} pid={os.getpid()} 日志={log_path.name}")
    print("格式工坊 · 本地服务已启动 → " + url)
    print(f"日志文件：{log_path}（超过 1 MB 自动轮转）")
    print("关闭这个窗口即可停止服务；或在浏览器里按 Ctrl+C。")
    if not args.no_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True,
            use_reloader=False)
