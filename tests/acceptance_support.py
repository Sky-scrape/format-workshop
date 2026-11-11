"""验收层的合成数据与操作助手。

与 `tests/conftest.py` 的分工：那边是底层契约夹具（渲染流水线、指纹判定），
这边是「用户会怎么操作」的小工具——上传、交付、造空稿/大稿/多标题稿。

所有数据都是现场生成的合成件，不读取仓库里的真实稿件目录，也不写 outputs/。
"""

import base64
import io
import json
from pathlib import Path

from docx import Document
from docx.shared import Pt

from conftest import RULES, SHIPPED_JOURNALS, body_texts

ROOT = Path(__file__).resolve().parent.parent

# 1x1 PNG：只用于图片段落行距测试，不含真实图像内容。
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)

# 统计研究的必备要件里，网页补录能补齐的部分。结构性要件阻断上线后，
# 空表单会被 400 拦下；合成稿只放中文骨架的用例统一带这几个值。
TJYJ_FILL = {
    "classification": "JEL: E62; O17",
    "english_title": "Digital Finance and Financing Constraints",
    "english_authors": "Zhang San, Li Si",
    "english_abstract": "This paper studies digital finance and financing constraints.",
    "english_keywords": "digital finance; financing constraints",
    "author_bio": "李明，男，经济学博士，研究方向：数字金融。",
}


# ---------------------------------------------------------------- 上传与交付

def upload(client, payload, name="source.docx", journal="tjyj"):
    """POST /convert。payload 可以是 Path，也可以是 bytes。"""
    data = Path(payload).read_bytes() if isinstance(payload, (str, Path)) else payload
    return client.post(
        "/convert",
        data={"journal": journal, "file": (io.BytesIO(data), name)},
        content_type="multipart/form-data",
    )


def job_id_from(response):
    location = response.headers.get("Location", "")
    assert "/job/" in location and location.endswith("/details"), location
    return location.split("/job/")[1].split("/")[0]


def upload_job(client, payload, name="source.docx", journal="tjyj"):
    response = upload(client, payload, name, journal)
    assert response.status_code == 302, response.get_data(as_text=True)[:500]
    return job_id_from(response)


def deliver(client, payload, name="source.docx", journal="tjyj", details=None):
    """上传 → 补录（默认全空）→ 出稿，返回 job id。真实跑完整流水线。"""
    job = upload_job(client, payload, name, journal)
    response = client.post(f"/job/{job}/details", data=details or {})
    assert response.status_code == 302, response.get_data(as_text=True)[:800]
    assert response.headers["Location"].startswith(f"/result/{job}")
    return job


# ---------------------------------------------------------------- 读现场

def meta_of(work, job):
    return json.loads((work / job / "meta.json").read_text(encoding="utf-8"))


def audit_of(work, job, journal="tjyj"):
    return json.loads(
        (work / job / f"manuscript.{journal}.audit.json").read_text(encoding="utf-8"))


def details_of(work, job):
    p = work / job / "details.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def para_texts(docx_path):
    doc = Document(str(docx_path))
    return [p.text for p in doc.paragraphs]


def docx_contains(docx_path, needle):
    return any(needle in t for t in para_texts(docx_path))


# ---------------------------------------------------------------- 合成数据

def build_empty_docx(path):
    """真正的空 docx：0 段。"""
    Document().save(str(path))
    return path


def build_title_only_docx(path, title="数字金融与企业融资约束"):
    doc = Document()
    doc.add_paragraph(title)
    doc.save(str(path))
    return path


def build_clean_docx(path, title="数字金融与企业融资约束",
                     abstract="内容提要：本文考察数字金融对企业融资约束的影响。",
                     keywords="关键词：数字金融；融资约束；企业创新"):
    """干净骨架：标题→摘要→关键词→正文→参考文献，没有作者页/英文区。

    补录页与注入用例用这份，避免 render_sample 的「作者信息页占位行」
    把 front_matter 的第一段抢走（那会让详情页显示「文章题目：」而不是真标题）。
    """
    doc = Document()
    doc.add_paragraph(title)
    doc.add_paragraph(abstract)
    doc.add_paragraph(keywords)
    doc.add_paragraph("一、引言")
    doc.add_paragraph("本文使用上市公司数据，考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.save(str(path))
    return path


def build_spacing_docx(path):
    """正文内部有连续空格、半角标点、全角数字——按现行白名单，这些一律不许改。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("研究发现,  数字金融 显著缓解了融资约束 。样本 共 １２００ 个。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.save(str(path))
    return path


def _cn_num(i):
    cn = "一二三四五六七八九十"
    return cn[i] if i < 10 else f"十{cn[i - 10]}"


def build_many_headings_docx(path, sections=12):
    """12 个一级标题：钉住中文序号「十一、十二、」与不崩。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    for i in range(sections):
        doc.add_paragraph(f"{_cn_num(i)}、第{i + 1}节")
        doc.add_paragraph(f"这是第{i + 1}节的正文，用于核验标题编号与正文版式。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.save(str(path))
    return path


def build_large_docx(path, sections=12, body_per_section=20, tables=12, refs=60):
    """大量数据：300+ 段、12 张表、12 个一级标题、60 条文献、3 张图。

    不测真实 80MB 文件（那属于手动性能验收），只测「块数量多时流水线不崩、
    对象不丢、文字守恒」。
    """
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")

    for i in range(sections):
        doc.add_paragraph(f"{_cn_num(i)}、第{i + 1}节")
        for j in range(body_per_section):
            doc.add_paragraph(
                f"第{i + 1}节第{j + 1}段。本文使用上市公司数据，考察数字金融对"
                f"企业融资约束的影响，并控制企业规模、资产负债率与盈利能力。")

    for t in range(tables):
        doc.add_paragraph(f"表{t + 1}  主要变量描述性统计")
        table = doc.add_table(rows=3, cols=3)
        for r, row in enumerate((("变量", "均值", "标准差"),
                                 ("Size", "8.42", "0.31"),
                                 ("Lev", "0.43", "0.12"))):
            for c, value in enumerate(row):
                table.cell(r, c).text = value
        doc.add_paragraph("注：括号内为标准差；*** 表示 1% 水平显著。")

    for f in range(3):
        fig = doc.add_paragraph()
        fig.add_run().add_picture(io.BytesIO(PNG_1PX))
        fig.paragraph_format.line_spacing = Pt(20)
        doc.add_paragraph(f"图{f + 1}  趋势图")
        doc.add_paragraph("资料来源：作者绘制。")

    doc.add_paragraph("参考文献")
    for n in range(1, refs + 1):
        doc.add_paragraph(
            f"[{n}] 作者{n}，{2000 + n % 20}：《数字金融研究》，"
            f"《经济研究》第{n % 12 + 1}期。")

    doc.save(str(path))
    return path


# ---------------------------------------------------------------- 草稿现场

def make_draft(tmp_rules, jid="testjournal", ok=True):
    """在临时规则库里造一份可发布的草稿，返回 (jid, draft_dir)。"""
    d = tmp_rules / "_drafts" / jid
    d.mkdir(parents=True, exist_ok=True)
    rules = json.loads((tmp_rules / "tjyj.json").read_text(encoding="utf-8"))
    rules["journal"]["id"] = jid
    rules["journal"]["name_zh"] = "测试期刊"
    report = {"ok_to_publish": ok, "errors": [] if ok else ["缺字段"],
              "warnings": [], "stats": {"extracted": 0, "inherited": 0}}
    (d / "draft.json").write_text(
        json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    (d / "report.json").write_text(
        json.dumps(report, ensure_ascii=False), encoding="utf-8")
    return jid, d
