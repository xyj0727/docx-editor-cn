# docx-editor-cn（《岩土工程学报》格式版）

用于 AI Agent（Claude Code 等）创建、编辑和排版 Word 文档（.docx）的 Skill。除 MathType 公式外，所有格式都按
**《岩土工程学报》（CJGE）** 的刊出版式设置。

## 安装

按 [skills](https://github.com/vercel-labs/skills) 的方式安装到 Claude Code：

```bash
npx --yes skills add https://github.com/xyj0727/docx-editor-cn -g -a claude-code -y --copy
```

- 用这种方式安装不会自动更新，更新命令：`npx skills update -g`。
- 也可以直接克隆到 `~/.claude/skills/docx-editor-cn`，更新用 `git pull`。
- 安装后需重启 Claude Code 才会加载。

## 能做什么

| 任务 | 命令 |
|---|---|
| 把已有论文整体排成 CJGE 格式 | `python scripts/apply_format.py apply 输入.docx 输出.docx` |
| 检查格式是否符合 CJGE | `python scripts/apply_format.py audit 输出.docx` |
| 新建 CJGE 论文 / Markdown 转 Word | `node scripts/new_doc.js [论文.md 输出.docx]` |
| 在解包的文档里插入三线表 | `python scripts/table.py unpacked/ "1" "题名" --caption-en "Table 1 Title" --headers … --rows …` |

## 主要格式

| 元素 | 格式 |
|---|---|
| 页面 | A4；左 1.65 cm、右 1.53 cm、上下 2.15 cm；题名至 Key words 通栏，正文双栏（栏宽 8.53 cm，栏距 0.74 cm） |
| 正文 | 宋体 + Times New Roman，五号，两端对齐，首行缩进 2 字符，行距固定值 15.6 磅 |
| 中文题名 / 作者 / 单位 | 黑体二号 / 仿宋五号 / 宋体六号，居中 |
| 摘要、关键词 | 小五，左右各缩进 0.75 cm，行距固定值 14 磅，标签黑体 |
| 英文题名 / 作者 / 单位 | TNR 加粗四号 / TNR 五号 / TNR 六号，居中；Abstract、Key words 标签加粗 |
| 标题 | 一级“0  引    言”黑体四号；二级“1.1  标题”黑体五号；三级“（1）标题”宋体五号缩进 2 字符；均不加粗 |
| 图表题 | “图1 题名”“表1 题名”黑体小五，英文题注 TNR 小五；图题在图下，表题在表上 |
| 表格 | 三线表：顶线、底线 0.75 磅，栏目线 0.5 磅，无竖线；表内小五 |
| 参考文献 | “参考文献：”黑体五号；条目小五，行距固定值 16 磅，悬挂缩进 2 字符 |
| 文末声明 | 利益冲突声明、作者贡献，小五，标签加粗 |
| 页眉页脚 | 页码在页眉：奇数页“第N期 / 第一作者，等. 题名 / 页码”，偶数页“页码 / 岩 土 工 程 学 报 / 年份”；无页脚 |

格式数值集中在 `scripts/format_profiles/cjge.json`，`new_doc.js` 使用同一套数值。

## MathType 公式

本 Skill 不处理 MathType 公式（排版时保持原样）。公式请配合 MathType 相关工具另行处理。

## 目录

```
docx-editor-cn/
├── SKILL.md                     ← Agent 读取的说明
├── scripts/
│   ├── apply_format.py          ← 已有文档排版 / 格式检查
│   ├── format_profiles/cjge.json← CJGE 格式数值
│   ├── new_doc.js               ← 新建 CJGE 论文、Markdown 转 Word
│   ├── mathml-to-docx.js        ← MathML → Word 公式（仅限不含 MathType 的文档）
│   ├── table.py                 ← 三线表 XML 插入
│   ├── formula.py               ← OMML 公式 XML 插入（仅限不含 MathType 的文档）
│   ├── comment.py / accept_changes.py
│   └── office/                  ← 解包、打包、校验、LibreOffice 调用
└── LICENSE.txt
```

## 依赖

- Python 3.9+，`python-docx`（`office/validate.py` 另需 `defusedxml`）
- Node.js 与 `npm install docx temml fast-xml-parser`（新建文档时需要）
- 可选：LibreOffice（`.doc` 转换、接受修订）、Pandoc（内容提取）

## 致谢与许可

基于 Gostyan 的原 docx-editor Skill 改造，按《岩土工程学报》版式重写了格式部分。
MIT 许可，见 [LICENSE.txt](LICENSE.txt)。
