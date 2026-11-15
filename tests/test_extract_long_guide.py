"""长稿约分段抽取与结构基线推荐（2026-11-15）。

旧版把稿约截断到 24000 字，后半本「图表规范」整章静默丢失。现在按
12000 字分段送模型、合并事实清单，引文核验仍对完整原文做；没有显式
选择基线时，按语言、引用制、印刷分栏三项推荐一本结构最接近的已入库刊。
"""

import json
import sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parent.parent / "tools"
sys.path.insert(0, str(TOOLS))

import extract_journal as ej  # noqa: E402


def test_chunk_text_splits_at_paragraphs_and_reports_truncation():
    text = "\n".join(f"第{i}段 " + "字" * 80 for i in range(200))
    chunks, truncated = ej._chunk_text(text, size=1000, limit=10)
    assert len(chunks) == 10 and truncated is True
    assert sum(len(c) for c in chunks) < len(text)
    assert all(len(c) <= 1000 for c in chunks)

    short, trunc2 = ej._chunk_text("短稿约", size=1000, limit=10)
    assert short == ["短稿约"] and trunc2 is False


def test_ai_facts_merges_chunks(monkeypatch):
    calls = []

    def fake_chat(cfg, messages, max_tokens=2048):
        user = messages[-1]["content"]
        calls.append(user)
        if "第 1/2 段" in user:
            return json.dumps({"facts": {"body_size_pt": 10.5},
                               "evidence": {"body_size_pt": "正文 10.5 磅"}},
                              ensure_ascii=False)
        return json.dumps({"facts": {"body_size_pt": 12, "limit_total": 18000},
                           "evidence": {"body_size_pt": "正文 12 磅",
                                        "limit_total": "不超过 1.8 万字"}},
                          ensure_ascii=False)

    monkeypatch.setattr(ej.ai_assist, "_chat", fake_chat)
    # 超过一段的长度，强迫分两段（按 12000 字）
    text = "甲" * 13000
    facts, ev, _notes, _raw = ej.ai_facts(text, {}, log=lambda _m: None)
    assert len(calls) == 2
    assert facts["body_size_pt"] == 10.5, "首段已有的值不应被后段覆盖"
    assert facts["limit_total"] == 18000
    assert ev["body_size_pt"] == "正文 10.5 磅"


def _rules(jid, lang, system, printed=1):
    return {"journal": {"id": jid, "language": lang},
            "citation": {"system": system},
            "page": {"printed_columns": printed}}


def test_recommend_baseline_matches_language_citation_and_columns(tmp_path, monkeypatch):
    (tmp_path / "zh_num.json").write_text(json.dumps(
        _rules("zh_num", "zh", "numbered")), encoding="utf-8")
    (tmp_path / "zh_ay.json").write_text(json.dumps(
        _rules("zh_ay", "zh", "author-year"), ensure_ascii=False), encoding="utf-8")
    (tmp_path / "en_ay.json").write_text(json.dumps(
        _rules("en_ay", "en", "author-year", printed=2)), encoding="utf-8")
    monkeypatch.setattr(ej, "RULES_DIR", tmp_path)

    zh_author_year = "本刊采用作者—年份制，正文引注写作（张三，2020）。"
    assert ej.recommend_baseline(zh_author_year) == "zh_ay"

    zh_numbered = "正文引用采用顺序编码制，标注为[1]。"
    assert ej.recommend_baseline(zh_numbered) == "zh_num"

    english = ("References should follow APA 7th edition. The journal uses a "
               "two-column layout and author-year citations.")
    assert ej.recommend_baseline(english) == "en_ay"

    assert ej.recommend_baseline("") == ""
