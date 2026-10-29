---
name: docx-editor-cn
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). All formatting follows the 《岩土工程学报》 (Chinese Journal of Geotechnical Engineering, CJGE) layout, except MathType equations. Triggers include: any mention of \"Word doc\", \"word document\", \".docx\", or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a \"report\", \"memo\", \"letter\", \"template\", or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
license: MIT. LICENSE.txt has complete terms
---

# DOCX creation, editing, and analysis — 《岩土工程学报》(CJGE) format

## Overview

A .docx file is a ZIP archive containing XML files.

Every format this skill produces or applies — page, front matter, headings, body, captions, tables,
references, statements, running headers — follows the layout of 《岩土工程学报》 (CJGE), measured from the
published PDF and the journal's author guidelines. The single source of truth for existing documents is
`scripts/format_profiles/cjge.json`; `scripts/new_doc.js` uses the same values for new documents.
**MathType equations are out of scope** (they follow the CJGE MathType spec through the
`mathtype-for-word` MCP).

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` or unpack for raw XML |
| Create new CJGE paper | `node scripts/new_doc.js` (edit FRONT / CONTENT first) or `node scripts/new_doc.js paper.md out.docx` |
| Edit existing document | Unpack → edit XML → repack - see Editing Existing Documents below |
| **Format an existing paper to CJGE** | `python scripts/apply_format.py apply IN.docx OUT.docx [--header "第一作者，等. 题名"] [--year 2026 --volume 48 --issue 9 --month 9]`, then `python scripts/apply_format.py audit OUT.docx` — see "Formatting Existing Documents" |
| Superscript in-text citations of an existing .docx | `python scripts/superscript_citations.py IN.docx OUT.docx` ([n], [n,m], [n-m]; reference-list entries untouched) |
| Insert 三线表 (XML editing) | `python scripts/table.py unpacked/ "1" "标题" --caption-en "Table 1 Title" --headers … --rows …` |
| Insert block formula (OMML, only for documents without MathType) | `python scripts/formula.py unpacked/ "LaTeX" 1 --anchor "锚文本"` |

## CJGE Format at a Glance

Word sizes: 二号 22 pt · 四号 14 pt · 小四 12 pt · 五号 10.5 pt · 小五 9 pt · 六号 7.5 pt.
Chinese text uses 宋体 / 黑体 / 仿宋 as listed; **all Latin letters and digits use Times New Roman**.

| Element | Font | Size | Paragraph |
|---|---|---|---|
| Page | — | — | A4; margins left 1.65 cm, right 1.53 cm, top/bottom 2.15 cm (版心 17.8 × 25.4 cm); header 1.5 cm |
| Columns | — | — | 题名 … Key words: one column; body: two columns 8.53 cm, gap 0.74 cm (continuous section break) |
| 中文题名 | 黑体 | 二号 | centred, fixed 30 pt, 8 pt after |
| 作者 | 仿宋 (序号上标) | 五号 | centred, fixed 15.6 pt; “简富献¹˒²，蔡正银*¹” |
| 单位 | 宋体 | 六号 | centred, fixed 12 pt; “（1. 单位，江苏 南京 210024；2. …）” |
| 摘  要：/ 关键词： | label 黑体, text 宋体 | 小五 | indent 0.75 cm left and right, justified, fixed 14 pt |
| 中图分类号 / 文献标识码 / 文章编号, 作者简介 | label 黑体 | 小五 | same indents, fixed 15.6 pt |
| English title | TNR bold | 四号 | centred, fixed 20 pt, 13.5 pt before, 11.8 pt after |
| English authors / affiliation | TNR | 五号 / 六号 | centred, fixed 15.6 pt |
| Abstract: / Key words: | label TNR bold (colon not bold) | 小五 | indent 0.75 cm, fixed 15.6 pt |
| First-page notes (基金项目, 通信作者, 收稿日期) | label 黑体 / English label bold | 六号 | fixed 12 pt |
| Body | 宋体 + TNR | 五号 | justified, first line 2 characters, **fixed 15.6 pt**, no space before/after |
| 一级标题 “0  引    言”, “1  试验设计” | 黑体, not bold | 四号 | flush left, 15 pt before, fixed 18 pt |
| 二级标题 “1.1  试验装置” | 黑体, not bold | 五号 | flush left, fixed 15.6 pt |
| 三级标题 “（1）级配影响” | 宋体, not bold | 五号 | first line 2 characters, own line, no end punctuation |
| 四级标题 “a）相似性。” | 宋体 + TNR | 五号 | runs in with the body text |
| 图题 “图1 题名” + “Fig. 1 Title” | 黑体 / TNR, not bold | 小五 | centred **below** the figure, fixed 15.6 pt |
| 表题 “表1 题名” + “Table 1 Title” | 黑体 / TNR, not bold | 小五 | centred **above** the table, fixed 14 pt |
| 分图 “(a) …” | 宋体 + TNR | 小五 | centred below the sub-figure |
| Table | 宋体 + TNR | 小五 | 三线表: top/bottom 0.75 pt, header rule 0.5 pt, no vertical rules, centred |
| 利益冲突声明/Conflict of Interests：, 作者贡献/Authors' Contributions： | label bold (中文黑体) | 小五 | flush left, fixed 15.6 pt, after the conclusions |
| 参考文献： | 黑体 | 五号 | flush left |
| Reference entries | 宋体 + TNR | 小五 | justified, fixed 16 pt, hanging indent 2 characters (18 pt), “[1] ” |
| Running header | 宋体 + TNR | 小五 | odd: 第N期 / 第一作者，等. 题名 / page; even: page / 岩 土 工 程 学 报 / YYYY年; 0.5 pt rule; **no footer** |

Text rules (content, not styling): number and unit separated by a space (“124 m”), “%”/“°” attached,
ranges with “～” (“0.5～1.0 mm”), “2.0×10<sup>7</sup> m<sup>3</sup>”, citations as superscript “[n]”,
references to figures/tables/equations as “图 1”“表 1”“式（5）”, full-width punctuation in Chinese
text and half-width in English and references, variables in TNR italic.

**In-text citations (all documents, including fund proposals):** every citation mark in body text is a
superscript square bracket placed directly after the cited author or statement, with no space and no box
or frame: “Xu 等<sup>[13]</sup>报道……”“……高于打入桩<sup>[27]</sup>，”. Grouped and ranged citations
are one superscript mark: <sup>[1,2]</sup>, <sup>[10,13,31,46]</sup>, <sup>[40-42]</sup>. The brackets
and digits are ordinary characters in a run with `w:vertAlign w:val="superscript"` (not a text box,
field or character border). Reference-list entries (“[1] 作者. 题名…”) stay on the baseline. For an
existing document run `python scripts/superscript_citations.py IN.docx OUT.docx` (skips reference-list
paragraphs that start with “[n] ”), then check the result with OfficeCLI.

### Converting .doc to .docx

Legacy `.doc` files must be converted before editing:

```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content

```bash
# Text extraction with tracked changes
pandoc --track-changes=all document.docx -o output.md

# Raw XML access
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes

To produce a clean document with all tracked changes accepted (requires LibreOffice):

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## Creating New Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Setup
```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation
After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.
```bash
python scripts/office/validate.py doc.docx
```

### Page Size

```javascript
// 《岩土工程学报》: A4, margins L 1.65 / R 1.53 / T,B 2.15 cm, header 1.5 cm
sections: [
  { // front matter (题名 … Key words): one column
    properties: { page: PAGE, titlePage: true, column: { count: 1 } },
    headers: buildHeaders(),
    children: FRONT,
  },
  { // body: two columns 8.53 cm, gap 0.74 cm, continuous break
    properties: { type: SectionType.CONTINUOUS, page: PAGE, column: { count: 2, space: 420 } },
    children: CONTENT,
  },
]
// PAGE = { size: { width: 11906, height: 16838 },
//          margin: { top: 1219, bottom: 1219, left: 935, right: 867, header: 850, footer: 567 } }
```

Text width 10104 DXA (17.82 cm); column width 4842 DXA (8.54 cm). Single-column figures 8.0 cm,
full-width figures 17.0 cm; single-column tables ≤ 8.5 cm, full-width tables ≤ 17.0 cm.

### Styles (CJGE)

**CRITICAL**: follow the table in "CJGE Format at a Glance". `scripts/new_doc.js` defines them as
paragraph styles (STYLES): Normal, Title, AuthorsCn, AffiliationCn, AbstractCn, FrontNote, TitleEn,
AuthorsEn, AffiliationEn, AbstractEn, PageNote, Heading1–3, RefHeading, FigureCaption(En),
TableCaption(En), SubCaption, TableText, Statement, Reference, HeaderText.

- Always give fonts in object form `font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'SimSun' }`
  (a string only sets the ASCII slot).
- Body and most elements use **exact** line spacing (`lineRule: EXACT`). A paragraph that holds an
  inline equation or picture must use **at least** instead (`lineRule: AT_LEAST`), otherwise the object is clipped;
  picture paragraphs use single spacing.
- Headings are **not bold** (黑体 carries the weight). Heading numbers are part of the text, separated from the
  title by two spaces; the introduction is “0  引    言” and two-character titles get four spaces (“结    论”).

**Creating a new document — use the template script:**
```bash
node scripts/new_doc.js                    # sample CJGE paper → output.docx
node scripts/new_doc.js paper.md out.docx  # Markdown → CJGE .docx
```

`scripts/new_doc.js` helpers — front matter: `titleCn`, `authorsCn(name, ['1','2'], …)`, `affiliationCn`,
`abstractCn`, `keywordsCn`, `clcLine`, `bioCn`, `titleEn`, `authorsEn`, `affiliationEn`, `abstractEn`,
`keywordsEn`; body: `h1/h2/h3(text)` (auto numbers “0”, “1.1”, “（1）”) or `h1Manual/h2Manual/h3Manual`,
`body(text)`, `figCaption('图1 题名', 'Fig. 1 Title')`, `tableCaption('表1 题名', 'Table 1 Title')`
(both return arrays — spread them), `subCaption('(a) …')`, `threeLineTable(headers, rows, colWidths)`,
`statement(label, text)`, `refHeading()`, `ref(text)`. Set `JOURNAL` (year, volume, issue, month,
runningHeader) for the running headers.

### Lists (NEVER use unicode bullets)

```javascript
// ❌ WRONG - never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })  // BAD
new Paragraph({ children: [new TextRun("\u2022 Item")] })  // BAD

// ✅ CORRECT - use numbering config with LevelFormat.BULLET
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Numbered item")] }),
    ]
  }]
});

// ⚠️ Each reference creates INDEPENDENT numbering
// Same reference = continues (1,2,3 then 4,5,6)
// Different reference = restarts (1,2,3 then 1,2,3)
```

### Tables (三线表 / Three-Line Table)

**CJGE tables are 三线表**: top and bottom rules 0.75 pt, one 0.5 pt rule under the header row, no vertical
rules, no fill, header not bold, 小五 text centred at fixed 14 pt. Caption **above** the table:
“表1 题名” (黑体小五) then “Table 1 Title” (TNR 小五). Header quantities are written “量/单位”, with
“/(°)” for degrees.

**For new documents — use `threeLineTable()` in `scripts/new_doc.js`:**
```javascript
...tableCaption('表1 试样参数', 'Table 1 Parameters of samples'),
threeLineTable(
  ['参数', '数值'],
  [['跨度/m', '1.2'], ['θ/(°)', '36']],
  [2000, 2842]            // optional DXA widths; default = one column (4842), full width = 10104
),
```

**For editing existing documents (XML) — use `scripts/table.py`:**
```bash
python scripts/table.py unpacked/ "1" "符号说明" --caption-en "Table 1 Symbols" \
    --headers "符号,说明" \
    --rows '[["S","状态空间"],["A","动作空间"]]' \
    --anchor "以下是符号说明"

# Print XML snippet only (no file modification):
python scripts/table.py --caption "表1 示例" --headers "列1,列2" --rows '[["a","b"]]'
```

**Key rules:**
- Width: one column 4842 DXA (8.5 cm), full width 10104 DXA (17.0 cm)
- Always use `WidthType.DXA` — never `WidthType.PERCENTAGE`
- Table width = sum of `columnWidths`; set matching `width` on each cell too

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  children: [new ImageRun({
    type: "png", // Required: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" } // All three required
  })]
})
```

### Page Breaks

```javascript
// CRITICAL: PageBreak must be inside a Paragraph
new Paragraph({ children: [new PageBreak()] })

// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Table of Contents

```javascript
// The overridden Heading1/2/3 styles (defined in paragraphStyles above) include outlineLevel,
// which TableOfContents uses to build the TOC. Using heading: HeadingLevel.HEADING_X on
// each paragraph applies that overridden style automatically — giving both TOC support AND
// the custom SimHei/size formatting. Do NOT skip paragraphStyles and rely on built-in styles.
new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-3" })

// Apply to heading paragraphs — heading property applies the overridden style:
new Paragraph({
  heading: HeadingLevel.HEADING_1, // → w:pStyle "Heading1" (SimHei 16pt, outlineLevel:0)
  children: [new TextRun("一、引言")]
})
new Paragraph({
  heading: HeadingLevel.HEADING_2, // → w:pStyle "Heading2" (SimHei 14pt, outlineLevel:1)
  children: [new TextRun("1.1 研究背景")]
})
```

### Headers/Footers (CJGE running headers)

No footer: the page number sits in the header. Odd and even pages differ, the first page has the
journal masthead (printed by the editorial office; `new_doc.js` prints it only when `JOURNAL.volume`
and `JOURNAL.issue` are set).

| Page | Left | Centre | Right | Rule |
|---|---|---|---|---|
| first | 第48卷 第9期 / 2026年 9月 | 岩 土 工 程 学 报 (小四) / Chinese Journal of Geotechnical Engineering | Vol.48 No.9 / Sep. 2026 | thick-thin double |
| odd | 第9期 | 第一作者，等. 中文题名 | page | 0.5 pt |
| even | page | 岩 土 工 程 学 报 | 2026年 | 0.5 pt |

All header text 宋体 + TNR 小五. Build with centre/right tab stops (`tabStops` at 5052 and 10104 DXA),
`evenAndOddHeaderAndFooters: true` on the Document and `titlePage: true` on the first section.

### Heading Numbering (CJGE)

Levels: `1 2 3` → `1.1 1.2` → `（1）（2）` → `a）b）`. Numbers are literal text:

```javascript
h1('引    言')      // → "0  引    言"   (first h1 is 0)
h1('试验设计')      // → "1  试验设计"
h2('试验装置')      // → "1.1  试验装置"
h3('级配影响')      // → "（1）级配影响"  (宋体, indented 2 characters)
body('a）相似性。…') // 四级标题 runs in with the text
```

No table of contents and no page breaks in a CJGE paper. Conclusions start with “为……，主要得到以下 N 点结论。”
followed by one paragraph per item “（1）…”.

### Captions

Captions are literal text, numbered per paper (not per chapter): “图1 题名” / “Fig. 1 Title” below the
figure, “表1 题名” / “Table 1 Title” above the table; no end punctuation. Sub-figures “(a)”, “(b)” go
below each sub-figure. `stripCaptionNumber()` is not needed — keep the number in the text.

### Block Formula Layout (OMML — only when the paper has no MathType)

CJGE equations are MathType objects formatted by the `mathtype-for-word` MCP; never mix them with the
OMML helpers here. For documents without MathType, `scripts/formula.py` / `formula()` in `new_doc.js`
build a 3-column borderless table sized to one column: [454 DXA | formula centred | 454 DXA number],
single line spacing (never exact, or the equation is clipped).

### References (CJGE, GB/T 7714)

Heading “参考文献：” (黑体五号, flush left). Entries 小五, fixed 16 pt, hanging indent 2 characters, numbered
“[1] ” in order of first citation; all punctuation half-width.

- Authors: more than three → first three + “, 等.” / “, et al.”; foreign names “DONG J Y” (surname in capitals,
  initials without dots); pinyin in translations “SHAN Yibo”.
- Volume in bold: “**53**(6)”; combined issues “(1/2)”; article numbers “**265**: 105428.”; page ranges “1847-1859”.
- Journal names in full. Non-English entries are followed by the full English entry in parentheses ending
  with “(in Chinese)”.

| Type | Tag | Format |
|------|-----|--------|
| 期刊 | [J] | 作者. 题名[J]. 刊名, 年, **卷**(期): 起止页码. |
| 专著 | [M] | 作者. 书名[M]. 版本项. 出版地: 出版者, 出版年: 引文页码. |
| 专著析出 | [M] | 析出作者. 析出题名[M]//专著作者. 专著题名. 出版地: 出版者, 年: 起止页码. |
| 学位论文 | [D] | 作者. 题名[D]. 保存地: 保存单位, 年份. |
| 会议 | [C] | 作者. 题名[C]//会议名. 出版地: 出版者, 年: 起止页码. |
| 标准 | [S] | 起草单位. 标准名称: 标准号[S]. 出版地: 出版者, 年. |
| 专利 | [P] | 申请者. 专利题名: 专利号[P]. 公告日期. |
| 电子资源 | [EB/OL] | 作者. 题名[EB/OL]. (更新日期)[引用日期]. 访问路径. |

Example: `[3] 单熠博, 陈生水, 钟启明. 堰塞体稳定性快速评价方法研究[J]. 岩石力学与工程学报, 2020, 39(9): 1847-1859. (SHAN Yibo, CHEN Shengshui, ZHONG Qiming. A rapid evaluation method of landslide dam stability[J]. Chinese Journal of Rock Mechanics and Engineering, 2020, 39(9): 1847-1859. (in Chinese))`

### Statements (文末声明)

After the conclusions and before the references, 小五, flush left, fixed 15.6 pt:
“**利益冲突声明/Conflict of Interests：**所有作者声明不存在利益冲突。All authors disclose no relevant conflict of
interest.” and “**作者贡献/Authors' Contributions：**…All the authors have read the final version of manuscript and
consented for submission.” Labels bold, Chinese part 黑体.

---

### Critical Rules for docx-js

- **Page**: A4 (11906 × 16838 DXA), margins L 935 / R 867 / T,B 1219 DXA, header 850; front matter one column, body two columns (space 420) after a continuous section break
- **Fonts**: object form with `ascii`/`hAnsi` = Times New Roman and `eastAsia` = SimSun / SimHei / FangSong
- **Line spacing**: exact values (15.6 pt body); at-least for paragraphs holding equations or pictures
- **Never use `\n`** - use separate Paragraph elements
- **Never use unicode bullets** - use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** - standalone creates invalid XML (CJGE papers normally have none)
- **ImageRun requires `type`** - always specify png/jpg/etc; figures 8.0 cm (single column) or 17.0 cm (full width)
- **Headings**: number in the text, two spaces before the title, not bold
- **Captions**: Chinese + English lines, 小五, not bold; figure caption below, table caption above
- **Tables MUST be 三线表** - top/bottom `size: 6` (0.75 pt), header rule `size: 4` (0.5 pt), all other borders `BorderStyle.NONE`, no fill
- **Always set table `width` with DXA** - never `WidthType.PERCENTAGE`; `columnWidths` and cell `width` must match
- **Use `ShadingType.CLEAR`** - never SOLID for table shading
- **Override built-in styles** - use exact IDs: "Heading1", "Heading2", etc., and include `outlineLevel`

---

## Formatting Existing Documents

`scripts/apply_format.py` applies the CJGE layout (profile `scripts/format_profiles/cjge.json`) to an
existing .docx **in place**: paragraph, run, table and section properties only, never rebuilding the
document, so fields, bookmarks and inline objects keep their positions. Output always goes to a new file.

```bash
python scripts/apply_format.py apply IN.docx OUT.docx [--header "第一作者，等. 题名"] \
       [--year 2026] [--volume 48 --issue 9 --month 9] [--no-journal-header] [--no-columns]
python scripts/apply_format.py audit OUT.docx
```

What it does:
- **Recognises every element** by style and text: 题名 (first line before “摘  要：”), 作者, 单位, 摘要, 关键词,
  中图分类号, 作者简介, the English block, first-page notes (基金项目 / 通信作者 / 收稿日期), headings (Heading 1–3
  styles or “0  引言”, “1.1  …”, “（1）…” text), captions (图/Fig./表/Table), sub-captions “(a)”, tables,
  statements, “参考文献：” and its entries; labels such as “摘  要：” get their own fonts.
- **Page and columns**: A4 + CJGE margins; a continuous section break after “Key words” puts the front matter
  in one column and the body in two (`--no-columns` keeps one column).
- **Running headers**: odd/even/first-page headers with the page number, no footer. The odd header defaults
  to “第一作者，等. 题名” from the front matter; the first-page masthead needs `--volume` and `--issue`.
  `--no-journal-header` keeps the existing header/footer text and only sets its fonts.
- **三线表** for every non-MathType table (existing table styles are removed).
- Word footnotes (first-page notes) get the 六号 note format.
- `audit` resolves styles like Word does (direct formatting → styles → defaults) and reports fonts, sizes,
  bold, line spacing, indents, alignment, table rules, margins and columns that differ from the profile.

Division of work:
- **MathType equations are out of scope.** Runs that hold objects, MathType display paragraphs and equation
  tables (tables holding `Equation.DSMT` objects) are skipped; paragraphs with inline objects get “at least”
  instead of “exact” spacing so nothing is clipped. MathType format follows the CJGE MathType spec and is
  applied by the `mathtype-for-word` MCP. Do **not** use `scripts/formula.py` (OMML) in documents that use MathType.
- Non-MathType formatting is checked with OfficeCLI (`officecli` skill), using its default workflow.
  MathType equations are checked twice: first with the `mathtype-for-word` MCP (`validate_mathtype_word_document`),
  then visually (render and inspect each equation). OfficeCLI is not used for MathType equations.

Pipeline for a paper with MathType: `apply_format.py apply` → `render_mathtype_word_document` (CJGE) →
`apply_format.py apply` again → `validate_mathtype_word_document` + visual check of the equations + OfficeCLI check.

## Editing Existing Documents

**Follow all 3 steps in order.**

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```
Extracts XML, pretty-prints, merges adjacent runs, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing. Use `--merge-runs false` to skip run merging.

### Step 2: Edit XML

Edit files in `unpacked/word/`. See XML Reference below for patterns.

**Use "Claude" as the author** for tracked changes and comments, unless the user explicitly requests use of a different name.

**Use the Edit tool directly for string replacement. Do not write Python scripts.** Scripts introduce unnecessary complexity. The Edit tool shows exactly what is being replaced.

**CRITICAL: Use smart quotes for new content.** When adding text with apostrophes or quotes, use XML entities to produce smart quotes:
```xml
<!-- Use these entities for professional typography -->
<w:t>Here&#x2019;s a quote: &#x201C;Hello&#x201D;</w:t>
```
| Entity | Character |
|--------|-----------|
| `&#x2018;` | ‘ (left single) |
| `&#x2019;` | ’ (right single / apostrophe) |
| `&#x201C;` | “ (left double) |
| `&#x201D;` | ” (right double) |

**Adding comments:** Use `comment.py` to handle boilerplate across multiple XML files (text must be pre-escaped XML):
```bash
python scripts/comment.py unpacked/ 0 "Comment text with &amp; and &#x2019;"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0  # reply to comment 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"  # custom author name
```
Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```
Validates with auto-repair, condenses XML, and creates DOCX. Use `--validate false` to skip.

**Auto-repair will fix:**
- `durableId` >= 0x7FFFFFFF (regenerates valid ID)
- Missing `xml:space="preserve"` on `<w:t>` with whitespace

**Auto-repair won't fix:**
- Malformed XML, invalid element nesting, missing relationships, schema violations

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings. Don't inject tracked change tags inside a run.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.
- **Never regenerate `styles.xml`**: When editing an existing document, all template styles live in `word/styles.xml`. Edit only `document.xml` content; do not overwrite or recreate `styles.xml` unless explicitly asked — doing so erases the user's template.
- **Preserve `<w:pStyle>` references**: When inserting new paragraphs, copy the `<w:pStyle w:val="..."/>` from an adjacent paragraph of the same type. Omitting `<w:pStyle>` silently falls back to the document default style, losing all heading/body formatting from the template.
- **Smart quotes are XML-encoded after unpack**: The unpack step converts `"` / `"` to `&#x201C;` / `&#x201D;` entities. When searching `document.xml` for heading text that contains Chinese quotation marks, search for the entity form (`&#x201C;`), not the raw Unicode character.

---

## XML Reference

### Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: Add `xml:space="preserve"` to `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)

### Tracked Changes

**Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Inside `<w:del>`**: Use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

**Minimal edits** - only mark what changes:
```xml
<!-- Change "30 days" to "60 days" -->
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting entire paragraphs/list items** - when removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Add `<w:del/>` inside `<w:pPr><w:rPr>`:
```xml
<w:p>
  <w:pPr>
    <w:numPr>...</w:numPr>  <!-- list numbering if present -->
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```
Without the `<w:del/>` in `<w:pPr><w:rPr>`, accepting changes leaves an empty paragraph/list item.

**Rejecting another author's insertion** - nest deletion inside their insertion:
```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion** - add insertion after (don't modify their deletion):
```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py` (see Step 2), add markers to document.xml. For replies, use `--parent` flag and nest markers inside the parent's.

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment markers are direct children of w:p, never inside w:r -->
<w:commentRangeStart w:id="0"/>
<w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted</w:delText></w:r>
</w:del>
<w:r><w:t> more text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment 0 with reply 1 nested inside -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

### Images

1. Add image file to `word/media/`
2. Add relationship to `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId5" Type=".../image" Target="media/image1.png"/>
```
3. Add content type to `[Content_Types].xml`:
```xml
<Default Extension="png" ContentType="image/png"/>
```
4. Reference in document.xml:
```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="914400" cy="914400"/>  <!-- EMUs: 914400 = 1 inch -->
    <a:graphic>
      <a:graphicData uri=".../picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

---

## Dependencies

- **pandoc**: Text extraction AND LaTeX→OMML formula conversion (`scripts/formula.py` requires pandoc ≥ 2.0)
- **docx**: `npm install docx` (new documents)
- **temml**: `npm install temml` (LaTeX → MathML conversion for Word native math)
- **fast-xml-parser**: `npm install fast-xml-parser` (MathML parsing for docx conversion)
- **LibreOffice**: PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images

---

## Markdown to Word Conversion (Chinese Academic Papers)

This section documents comprehensive solutions for converting Markdown papers with LaTeX formulas, tables, and citations to properly formatted Word documents following Chinese academic standards.

### Quick Start

```bash
# Install dependencies
npm install docx temml fast-xml-parser

# Run conversion
node scripts/new_doc.js
```

### Critical Issues & Solutions

The following 8 issues were identified and solved during production usage. All solutions are implemented in `scripts/new_doc.js` and `scripts/mathml-to-docx.js`.

#### Issue 1: Three-Line Table Middle Borders Visible

**Problem**: Body row borders in 三线表 appeared visible instead of invisible.

**Solution**: Set ALL body row borders to `NONE`, only keep (CJGE widths):
- Header top: `THICK` (0.75pt)
- Header bottom: `THIN` (0.5pt)
- Last row bottom: `THICK` (0.75pt)

```javascript
const THICK = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const THIN  = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// Header row
cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)

// Body rows - ALL borders NONE except last row bottom
cellOf(cell, colWidths[i], {
  top: NONE,
  bottom: isLastRow ? THICK : NONE,
  left: NONE,
  right: NONE,
})
```

#### Issue 2: Formula Table Borders Visible

**Problem**: Block formula tables (3-column layout) showed visible borders.

**Solution**: Set ALL borders including `insideHorizontal` and `insideVertical` to `NONE`:

```javascript
return new Table({
  width: { size: COLUMN_W, type: WidthType.DXA },
  columnWidths: [454, COLUMN_W - 908, 454],
  borders: {
    top: NONE,
    bottom: NONE,
    left: NONE,
    right: NONE,
    insideHorizontal: NONE,  // CRITICAL: Must include these
    insideVertical: NONE,    // CRITICAL: Must include these
  },
  rows: [new TableRow({ children: [leftCell, formulaCell, numberCell] })],
});
```

#### Issue 3: Heading Numbering (superseded for CJGE — numbers are literal text, see "Heading Numbering (CJGE)")

**Problem**: Originally, heading numbering was baked into paragraph text via manual JS counters. When content was added/removed in Word, numbers did not update.

**Solution**: H2/H3 use Word's native auto-numbering via per-chapter numbering references. H1 keeps Chinese numerals in text (一、二、三). The `buildNumberingConfig(chapterCount)` function generates a `sections_c{N}` numbering reference for each chapter with 2 levels. At each chapter boundary, `_chapter` increments and H2/H3 switch to a fresh numbering reference, achieving automatic reset.

```javascript
let _chapter = 0;

function h1Chinese(text) {  // Chinese numeral in text
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h2(text) {  // Word auto-numbered (level 0 in per-chapter ref)
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h3(text) {  // Word auto-numbered (level 1 in per-chapter ref)
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}
```

#### Issue 4: Heading English/Numbers in Wrong Font (SimSun instead of Times New Roman)

**Problem**: English text and numbers in headings, figure captions, and table captions displayed in SimSun (宋体) instead of Times New Roman.

**Solution**: Use mixed font configuration with `ascii`, `eastAsia`, and `hAnsi` properties:

```javascript
// Style definition for headings
{
  id: 'Heading1', name: 'Heading 1', basedOn: 'Normal',
  run: {
    font: {
      ascii: 'Times New Roman',    // English letters, numbers (CJGE: TNR)
      eastAsia: 'SimHei',        // Chinese characters (黑体)
      hAnsi: 'Times New Roman',     // Western European characters
    },
    size: 28, bold: false,          // 黑体四号, not bold
  },
  // ... paragraph settings
}

// For captions (SimSun body font for Chinese)
font: { ascii: 'Times New Roman', eastAsia: 'SimSun', hAnsi: 'Times New Roman' }
```

#### Issue 5: Block Formulas Not Using Word Equation Editor

**Problem**: LaTeX formulas were rendered as plain text or images instead of native Word equations.

**Solution**: Use `temml` (LaTeX→MathML) + `mathml-to-docx.js` (MathML→OMML) pipeline:

```javascript
const temml = require('temml');
const { mathmlToDocxChildren } = require('./mathml-to-docx');

function latexToMath(latex) {
  const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
  const children = mathmlToDocxChildren(mathml);
  if (children && children.length) {
    return new Math({ children });
  }
  // Fallback
  return new Math({ children: [new MathRun(latex)] });
}

// Usage in formula table
const mathObj = latexToMath('Q_n(x, a) = r + \\gamma V_{n-1}(y)');
```

#### Issue 6: Main Title English in Wrong Font

**Problem**: Document title's English text showed in SimSun instead of Times New Roman.

**Solution**: Apply same mixed font to title paragraph:

```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: '论文标题 Paper Title',
    bold: false,
    size: 44,   // 黑体二号
    font: { ascii: 'Times New Roman', eastAsia: 'SimHei', hAnsi: 'Times New Roman' },
  })],
})
```

#### Issue 7: Inline Math Detection Too Aggressive

**Problem**: Inline math regex matched plain numbers (like "1992") and English words (like "Agent", "Watkins"), incorrectly converting them to formula objects.

**Solution**: Use strict regex that ONLY matches actual mathematical content:

```javascript
// Detection function - returns true only for real math content
function containsMath(text) {
  // Greek letters
  if (/[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text)) return true;
  // Unicode subscript/superscript characters
  if (/[₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]/.test(text)) return true;
  // Math operators and special symbols
  if (/[∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩]/.test(text)) return true;
  // Starred symbols like π*, Q*
  if (/[A-Z]\*/.test(text)) return true;
  // Explicit $...$ LaTeX
  if (/\$[^$]+\$/.test(text)) return true;
  return false;
}

// Parsing regex - strict matching for inline formulas
const mathPattern = /\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;

// What it matches:
// - $...$ explicit LaTeX
// - Qₙ(x,a) - function with subscripts
// - V(αₙ) - function with Greek params
// - αₙ - Greek with subscript
// - xₙ - variable with subscript
// - Q* - starred variable

// What it does NOT match:
// - Plain numbers: 1992, 500
// - English words: Agent, Watkins, Dayan
// - Plain parentheses: (1992), (optional)
```

#### Issue 8: Citations Not in Superscript

**Problem**: Reference citations like [1], [2] appeared as normal text instead of superscript; grouped
or ranged citations such as [1,2] or [40-42] were missed by the single-number pattern.

**Solution**: Detect citation pattern (single, grouped and ranged) and apply `superScript: true`. For
existing .docx files use `scripts/superscript_citations.py`:

```javascript
function containsCitation(text) {
  return /\[\d+(?:\s*[-–,，]\s*\d+)*\]/.test(text);
}

// In parseInlineContentWithCitations():
const combinedPattern = /(\[\d+(?:\s*[-–,，]\s*\d+)*\])|...; // Citation first in alternation

if (match[1]) {
  // Citation [n] - convert to superscript
  children.push(new TextRun({
    text: match[1],
    superScript: true,
  }));
}
```

#### Issue 10: Page Breaks (not used in CJGE papers)

**Note**: CJGE papers run continuously — no page break after the abstract or before the references; the
front matter ends with a continuous section break instead. The helper is kept for other documents.

**Problem (other documents)**: Abstract page should end after keywords, and references should start on a new page.

**Solution**: Add `pageBreak()` helper function:

```javascript
// ISSUE 10 FIX: Page break helper
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Usage in CONTENT:
// After keywords (end of abstract)
body('关键词：强化学习；Q-learning'),
pageBreak(),  // <- Start new page after abstract

// Before references
pageBreak(),  // <- Start new page for references
new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun('参考文献')],
}),
```

**Important**: `PageBreak` MUST be wrapped in a `Paragraph` - it cannot be used standalone.

#### Issue 11: Formula Table Cell Vertical Alignment

**Problem**: In the 3-column borderless formula table, the formula and equation number `(n)` are not vertically centered, causing misalignment.

**Solution**: Add `verticalAlign: VerticalAlign.CENTER` to all three TableCell definitions:

```javascript
// First, add VerticalAlign to imports:
const {
  // ... other imports
  VerticalAlign,  // <-- Add this
} = require('docx');

// Then apply to each cell in the formula() function:
const leftCell = new TableCell({
  width: { size: 454, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
});

const formulaCell = new TableCell({
  width: { size: COLUMN_W - 908, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [mathObj],
  })],
});

const numberCell = new TableCell({
  width: { size: 454, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    indent: { firstLine: 0 },
    children: [new TextRun(`(${number})`)],
  })],
});
```

**Result**: Formula and equation number now align horizontally on the same baseline.

### File Structure

```
scripts/
├── new_doc.js          # CJGE template (front matter, headings, captions, 三线表, headers) + Markdown converter
├── apply_format.py     # Apply / audit the CJGE layout on an existing .docx
├── format_profiles/cjge.json  # CJGE layout values used by apply_format.py
├── mathml-to-docx.js   # MathML→docx Math converter (~250 lines)
├── formula.py          # Block formula XML insertion (legacy)
├── table.py            # CJGE 三线表 XML insertion
└── office/
    ├── unpack.py       # DOCX→XML extraction
    ├── pack.py         # XML→DOCX assembly
    └── validate.py     # Document validation
```

### Unicode to LaTeX Mapping

The script includes comprehensive Unicode math symbol conversion:

```javascript
const UNICODE_TO_LATEX = {
  // Greek letters
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', // ... etc
  // Subscripts
  '₀': '_0', '₁': '_1', '₂': '_2', // ... etc
  // Superscripts
  '²': '^2', '³': '^3', // ... etc
  // Special symbols
  '∞': '\\infty', '∑': '\\sum', '∫': '\\int', // ... etc
};
```

### MathML to DOCX Conversion

`scripts/mathml-to-docx.js` converts MathML (from temml) to docx Math components:

- Fractions: `<mfrac>` → `MathFraction`
- Subscripts: `<msub>` → `MathSubScript`
- Superscripts: `<msup>` → `MathSuperScript`
- Combined: `<msubsup>` → `MathSubSuperScript`
- Radicals: `<msqrt>`, `<mroot>` → `MathRadical`
- Summation: `<munderover>` with ∑ → `MathSum`
- Integrals: `<munderover>` with ∫ → `MathIntegral`
- Matrices: `<mtable>` → `MathMatrix`

### Workflow for New Conversions

1. **Copy `new_doc.js` template** to your project
2. **Install dependencies**: `npm install docx temml fast-xml-parser`
3. **Edit CONTENT section** with your document structure
4. **Use helper functions**:
   - `h1/h2/h3(text)` - Headings with auto-numbering
   - `body(text)` - Body paragraph (auto-detects math/citations)
   - `formula(latex, number)` - Block formula
   - `threeLineTable(headers, rows, colWidths)` - Three-line table
   - `tableCaption/figCaption(label)` - Captions
   - `pageBreak()` - Page break (Issue 10)
   - `ref(text)` - Reference entry
5. **Run**: `node new_doc.js`

---

## 实战坑与做法（2026-10-20，用 new_doc.js 把 Markdown 结果报告转为 Word）

- **依赖没装**：技能目录没有 node_modules，`node scripts/new_doc.js` 会报找不到 docx。不要往技能目录装；在工作目录里复制 `package.json`、`package-lock.json` 后 `npm ci`，再把 `new_doc.js`、`mathml-to-docx.js` 复制过去运行。
- **希腊字母、Unicode 下标会被转成 Word 公式（OMML）**：`containsMath()` 一见到 θ、η、₀ 之类就走公式路径，每个 θ 都变成一个公式对象（Cambria Math 字体，不是 TNR 斜体）；`M₀.₂₅` 还会被拆成公式 M₀ 加普通文本 “.₂₅”。变量多的报告应改为文本 run：在副本里把 `containsMath/parseInlineContent/parseInlineContentWithCitations/buildInlineMathRuns` 换成一个解析 `<i>…</i>`、`<sub>…</sub>`、`<sup>…</sup>` 的函数（TextRun 的 italics/subScript/superScript），并让 `parseHtmlTable` 保留这三种标签。
- **Markdown 支持范围**：表格只认一行内的 `<table><tr><th>…</th></tr><tr><td>…</td></tr></table>`，不认管道表；不解析 `**粗体**`、反引号；只认顶层 “- ” 列表；表题须在表格上方两行（“表1 …”“Table 1 …”），图题在图片行下方。
- **图表编号**：`normalizeCaptionText` 只认 “图1”“表1” 的全文连续编号，“图2-1” 会被改成 “图2 -1”。先用 `{{F:key}}`、`{{T:key}}` 占位，拼接全文后按题注出现顺序统一编号。
- **引文上标**：所有 `[n]` 都排成上标，包括 “文献[1]的表1”（作句子成分时应平排）和 “错误代码 [102]”。要按 “前面是‘文献’或数字超过参考文献条数则平排” 处理。
- **报告用单栏**：模板正文是双栏、图宽 8 cm、表格默认一栏宽。写详细报告时在副本中把正文节改为 `column: { count: 1 }`，图最大宽 643 px（17 cm），表格按内容估列宽（汉字约 180 DXA、西文约 100 DXA，总宽不超过 10104 DXA），页眉换成报告题名加页码（去掉“岩 土 工 程 学 报”）。
- **检查格式改用 OfficeCLI（2026-10-22，Word MCP 已删除）**：查上标时选择器写 `run[superscript]`，不是 `run[vertAlign=superscript]`（`vertAlign` 只用于 `set`/`add`，读出来的键名是 `superscript`）。
- 参考实现：`D:\Pile\bucket_results\report_build\make_report_doc.py`（由 new_doc.js 生成单栏报告版 report_doc.js）和 `assemble_report.py`（拼接、编号、记号转换）。
