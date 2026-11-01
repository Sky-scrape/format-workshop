# 期刊规则集

每本期刊一套三件产物，全部由 `{id}.json` 生成，**禁止手写**。

| 文件 | 是否入库 | 说明 |
|---|---|---|
| `{id}.json` | 是 | 唯一真源。规则集本体 |
| `{id}.spec.html` | 是 | 规格页。逐条列出要求 + 内嵌成品预览 |
| `{id}.sample.docx` | 否 | 版式样板。由 JSON 渲染，属构建产物 |
| `{id}.sample.pdf` | 否 | 样板的 PDF 版，供预览与打印核对 |

样板和 PDF 未入库，是因为它们完全可由 JSON 重建，入库只会让每次规则修订都产生二进制 diff。需要时本地生成：

```bash
python tools/build_journal.py rules/journal/{id}.json
```

## 本仓库收录范围

开源仓库只带 **3 本示例**：

- `cjyj` 财经研究
- `tjyj` 统计研究
- `gjjryj` 国际金融研究

其余规则集通过 `.gitignore` 排除，留在本地。**不要**把本机全部规则集提交上来——完整规则库不随仓库分发，这是有意的。

如果你要往仓库里加自己的期刊，在 `.gitignore` 里加两行例外：

```gitignore
!rules/journal/{id}.json
!rules/journal/{id}.spec.html
```

## 新增一本期刊

1. 收集材料：官方投稿须知（PDF/网页）+ 2–3 篇该刊近期已发表论文
2. 生成草稿：

   ```bash
   python tools/extract_journal.py 投稿须知.pdf -o rules/journal/{id}.json
   ```

3. **人工逐条核对草稿**。自动提取只能省掉录入，不能省掉判断——字号行距这类版式要求，投稿须知里往往根本不写，得从已发表论文反推
4. 填 `verification` 段：来源 URL、核验日期、核验方式、置信度。反推得来的必须标 `typography_source: published_layout`
5. 生成样板与规格页并人工比对：

   ```bash
   python tools/build_journal.py rules/journal/{id}.json
   ```

6. 打开 `{id}.sample.docx`，逐项与该刊真实论文比对字号、行距、版心、图表注层次
7. 三件产物一次提交

字段定义见 [`../journal-schema.md`](../journal-schema.md)。

## 一条硬规矩

规则集里出现的每一个格式值，都必须能在样板里找到对应，且能标出来源。找不到出处的值不要写——那说明渲染器里有硬编码，会让样板与规则集悄悄漂移。

查不到就写 `null`，并在报告里标"待确认"。宁可标不确定，不要编一个看起来合理的数字。
