"""论文格式转换 · 本地 Web 应用

两段式流水线（对应 DESIGN §10.2 的执行顺序）：

    POST /convert            上传 → 解析 → 投稿体检            → 302 /job/<id>/details
    GET  /job/<id>/details   个人信息与缺失要件补录（预填已有值）
    POST /job/<id>/details   注入补录内容 → 渲染 → 审计          → 302 /result/<id>
    GET  /result/<id>        成品、报告、换刊重排、差异表入口

其余路由：
    GET  /journals                      期刊规则库
    GET  /history                       本机最近的转换记录
    GET  /job/<id>/diff?to=<jid>        两本刊的规则差异
    POST /job/<id>/retry                同一稿件换刊重排
    GET  /job/<id>/download/<kind>      带可读文件名的下载
    GET  /journal/<jid>/spec            期刊格式规格页
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


def load_journals():
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

def sweep(now=None):
    now = now or time.time()
    if not WORK.exists():
        return
    for d in WORK.iterdir():
        try:
            if d.is_dir() and now - d.stat().st_mtime > JOB_TTL:
                shutil.rmtree(d, ignore_errors=True)
                with _JOB_LOCKS_GUARD:      # 任务没了，锁也一并回收（F11）
                    _JOB_LOCKS.pop(d.name, None)
        except OSError:
            continue


def job_dir(job):
    if not JOB_RE.match(job or ""):
        abort(404)
    d = WORK / job
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


def write_meta(d, meta):
    """meta.json 原子写：先落临时名再 os.replace，硬杀也只会留下完整的旧值或
    完整的新值，不会留下半个 JSON（F34；同款姿势见 docx_meta.save）。"""
    p = d / "meta.json"
    tmp = d / "meta.json.tmp"
    try:
        tmp.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, p)
    finally:
        tmp.unlink(missing_ok=True)


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


def doc_to_docx(src, d):
    """老 .doc 先用 LibreOffice 转成 docx，再进流水线。"""
    exe = next((c for c in SOFFICE if c and Path(c).exists()), None)
    if not exe:
        raise ConvertError("这台机器上没有找到 LibreOffice",
                          "转 .doc 需要它。请先用 Word 把稿件另存为 .docx 再上传。")
    subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", str(d), str(src)],
                   capture_output=True, timeout=STEP_TIMEOUT)
    out = d / (src.stem + ".docx")
    if not out.exists():
        cand = [p for p in d.glob("*.docx") if p.stem != "manuscript"]
        if not cand:
            raise ConvertError("这份 .doc 转不出 DOCX",
                              "文件可能损坏或受保护。请在 Word 里另存为 .docx 后重新上传。")
        out = cand[0]
    return out


# ---------------------------------------------------------------- 流水线

def stage_parse(d, jid, original):
    """解析 + 体检。产出 manuscript.ir.json 与 triage.html。"""
    src, ir = d / "manuscript.docx", d / "manuscript.ir.json"
    run(["tools/parse_docx.py", src, "-o", ir])
    ir_data = read_json(ir)
    # 解析器看到的是任务目录里的 manuscript.docx，把来源名改回用户上传的文件名
    if original and ir_data.get("source") != original:
        ir_data["source"] = original
        ir.write_text(json.dumps(ir_data, ensure_ascii=False, indent=2), encoding="utf-8")
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


def stage_render(d, jid, original):
    """补录内容注入 → 排版 → 审计的入口：同任务一次只放一轮进来（F11）。

    先抢本任务的渲染锁，抢不到说明上一轮还没结束，立刻以 RenderBusy
    把第二个并发请求挡回去，绝不与新一轮互删互写。抢到后的流程见
    _stage_render：所有产物先写 .staging 临时名，全部成功后才原子落位
    （os.replace），中途任何一步失败，上一轮已交付的成品、审计与 meta
    都原封不动——否则一次失败的换刊/驳回就会毁掉已经交付的结果页（F36）。
    """
    lock = job_lock(d.name)
    if not lock.acquire(blocking=False):
        raise RenderBusy(d.name)
    try:
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
        "audit_total": len(audit.get("items", [])),
        "elapsed": round(time.time() - t0, 1),
        "created": time.strftime("%Y-%m-%d %H:%M"),
        "docx": outdocx.name,
        "docx_anon": anon_docx.name if anon_docx else "",
        "anon_removed": len((read_json(audit_json, {}) or {}).get("anonymize", {}).get("removed", [])),
        "rejected": len(audit.get("rejected", [])),
        "injected": injected,
        "stage": "rendered",
    })
    write_meta(d, meta)

    # ---- 全部成功，此刻才落位：成品/匿名稿/审计原子替换旧文件，体检页与
    # 审计页随后换新，再清掉旧刊遗留的成品与审计。顺序保证 meta 指向的
    # 文件在每个时点都是完整的；失败路径根本走不到这里（F36）。
    final_audit = outdocx.with_suffix(".audit.json")
    tmp_audit = tmp_docx.with_suffix(".audit.json")
    keep = {outdocx.name, final_audit.name}
    if anon_docx:
        keep.add(anon_docx.name)
        os.replace(anon_tmp, anon_docx)
    os.replace(tmp_docx, outdocx)
    os.replace(tmp_audit, final_audit)
    os.replace(tmp_triage, d / "triage.html")
    os.replace(tmp_audit_html, d / "audit.html")
    for old in (list(d.glob("manuscript.*.docx")) + list(d.glob("manuscript.*.audit.json"))
                + list(d.glob(".staging.*"))):
        if old.name == "manuscript.docx" or old.name in keep:
            continue
        old.unlink(missing_ok=True)
    (d / "error.json").unlink(missing_ok=True)
    return meta


def fail_page(d, err, journals=None):
    meta = read_meta(d)
    (d / "error.json").write_text(json.dumps(
        {"title": err.title, "hint": err.hint, "detail": err.detail}, ensure_ascii=False),
        encoding="utf-8")
    write_meta(d, meta)
    return render_template("failed.html", job=d.name, err=err, meta=meta,
                           journals=journals if journals is not None else load_journals(),
                           selected=meta.get("journal"),
                           # IR 不存在 = 解析就没成功过：换刊重试不可能有用，
                           # 页面要给「重新上传」的出路而不是诱导再试（F12）
                           can_retry=(d / "manuscript.ir.json").is_file()), 200


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
    citation_offered = bool(cite_rl.get("templates")) and has_cite
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
    def back(msg):
        return render_template("journal_new.html", ai_ok=ai_assist.available(),
                               cfg=ai_assist.load_config(), error=msg,
                               ocr_ok=ocr.available(), ocr_engine=ocr.engine_name(),
                               drafts=draft_list(),
                               existing={j["id"] for j in load_journals()}), 400

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
    if (RULES_DIR / f"{jid}.json").exists():
        return back(f"代号 {jid} 已被占用，换一个")

    d = DRAFT_DIR / jid
    d.mkdir(parents=True, exist_ok=True)
    src = d / "source.pdf"
    f.save(src)

    try:
        rules, report, text = extract_journal.extract(src, jid=jid, name_zh=name)
    except Exception as e:  # noqa: BLE001
        app.logger.exception("稿约抽取失败")
        return back(f"抽取失败：{e}")

    (d / "draft.json").write_text(json.dumps(rules, ensure_ascii=False, indent=2),
                                  encoding="utf-8")
    (d / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
    (d / "source.txt").write_text(text, encoding="utf-8")
    return redirect(url_for("journal_draft", jid=jid))


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
    d = draft_dir(jid)
    rules = read_json(d / "draft.json", {}) or {}
    report = read_json(d / "report.json", {}) or {}

    def back(msg):
        ext = (rules.get("verification") or {}).get("extraction") or {}
        return render_template("journal_draft.html", jid=jid, rules=rules, report=report,
                               ext=ext, journal=rules.get("journal") or {},
                               evidence=ext.get("used") or {},
                               inherited=ext.get("inherited") or [],
                               dropped=ext.get("dropped") or [],
                               ai_on=ai_assist.available(), error=msg), 400

    if not report.get("ok_to_publish"):
        return back("还有阻断性错误，先把错误项修掉")
    if request.form.get("confirm") != "on":
        return back("请先勾选「我已逐项核对」——AI 抽出来的东西不能未经核对就进规则库")
    target = RULES_DIR / f"{jid}.json"
    if target.exists() and request.form.get("overwrite") != "on":
        return back(f"{jid}.json 已存在。确实要覆盖请勾选覆盖选项")

    if target.exists():
        bak = d / f"{jid}.json.bak"
        shutil.copy2(target, bak)
    target.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")

    # 三件产物必须同源：json 落盘后立刻重建样板与规格页
    try:
        run(["tools/build_journal.py", str(target)])
    except Exception:  # noqa: BLE001
        app.logger.exception("规则集已保存，但样板/规格页重建失败")

    shutil.rmtree(d, ignore_errors=True)
    return redirect(url_for("journal_spec", jid=jid))


@app.route("/journals/draft/<jid>/discard", methods=["POST"])
def journal_discard(jid):
    shutil.rmtree(draft_dir(jid), ignore_errors=True)
    return redirect(url_for("journals"))


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


@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    jid = request.form.get("journal") or ""
    keep = jid if (RULES_DIR / f"{jid}.json").exists() else ""
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

    if request.method == "POST":
        data = {}
        for key, *_ in DETAIL_FIELDS:
            data[key] = (request.form.get(key) or "").strip()
        for key, *_ in FOOTER_FIELDS:
            data[key] = (request.form.get(key) or "").strip()
        for f in author_page_fields(rules_of(meta["journal"]) or {}):
            data[f["key"]] = (request.form.get(f["key"]) or "").strip()
        # 原子写（F11）：这个文件在抢锁前落笔，而持锁的上一轮渲染开头就要读它
        # （_stage_render 首步 read_json 对半截 JSON 抛异常）。同值幂等只保证值
        # 相同，不保证写入动作与持锁者读不交错——非原子 write_text 的毫秒级窗口
        # 命中时，第一轮会以 JSONDecodeError 误报失败页。先落临时名再 os.replace，
        # 读方要么见完整旧值、要么见完整新值，永远见不到半个 JSON（同款姿势见
        # write_meta / F34）。Windows 上对方进程恰好开着这个文件时 os.replace
        # 会抛 PermissionError（毫秒不到的窗口），稍候重试即可，不必失败整轮。
        p = d / "details.json"
        tmp = d / "details.json.tmp"
        try:
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                           encoding="utf-8")
            for attempt in range(5):      # 对方进程读完即放手；5 次共约 0.44s
                try:
                    os.replace(tmp, p)
                    break
                except PermissionError:
                    if attempt == 4:
                        raise
                    time.sleep((0.02, 0.04, 0.08, 0.1, 0.2)[attempt])
        finally:
            tmp.unlink(missing_ok=True)
        # 引用体系重排是唯一会改写正文/条目文字的开关，默认关，由用户显式勾选
        meta["citation"] = request.form.get("citation") == "on"
        write_meta(d, meta)
        try:
            stage_render(d, meta["journal"], meta.get("original", "稿件.docx"))
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
    (d / "reject.json").write_text(json.dumps({"ids": ids}, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
    try:
        stage_render(d, meta["journal"], meta.get("original", "稿件.docx"))
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
    if not (RULES_DIR / f"{jid}.json").exists():
        abort(400)
    if not (d / "manuscript.docx").exists():
        abort(404)
    if not meta.get("docx"):
        # 还没有成品：换刊只是改后续渲染的目标，改完回补录页
        meta["journal"] = jid
        meta["journal_name"] = (journal_by_id(jid) or {}).get("name", jid)
        if not (d / "manuscript.ir.json").is_file():
            # 解析阶段就没成功（IR 不存在）：去补录页再渲染只会把死循环走下去，
            # 坏文件换任何刊都不可能成功（F12）。先按新刊重跑解析；再失败会落到
            # failed 页的「重新上传」出路（failed.html 按 can_retry 区分文案）。
            try:
                stage_parse(d, jid, meta.get("original", "稿件.docx"))
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
def download(job, kind):
    d = job_dir(job)
    meta = read_meta(d)
    stem = meta.get("stem") or "稿件"
    jn = meta.get("journal_name") or meta.get("journal", "")
    if kind in ("docx", "docx-anon"):
        if kind == "docx-anon":
            if not meta.get("docx_anon"):
                abort(404)
            p = d / meta["docx_anon"]
            label = "匿名稿"
        else:
            p = d / (meta.get("docx") or "")
            label = "署名稿"
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
    p = RULES_DIR / f"{jid}.spec.html"
    if not p.exists():
        abort(404)
    return send_from_directory(RULES_DIR, p.name)


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
