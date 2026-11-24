---
name: journal-docx-formatting-audit
description: >-
  Audits and typesets Chinese journal Word manuscripts against a merged template
  profile (default Proceedings of the CSEE). When the user supplies a new .doc
  or .docx template, import-template updates only the fields that template
  states or shows; every field the new template does not define stays unchanged.
  Line spacing, figure interior font size, and table font size are hard gates.
  Layout shortfalls are fixed by rewriting prose, not by compressing line
  spacing, figures, or chart type. Stop only when audit JSON stop.allowed is true.
---

# Journal Word (.docx) Formatting Audit & Two-Column Typesetting

## 通过条件

审计报告里的 `stop.allowed` 为 true 之前，不得结束，也不得把稿件说成已经通过。`passed` 与 `stop.allowed` 只有在档案里的全部门闩都成立时才为 true。其中包含本次模板没有改写、因此继续有效的旧规定。

栏底不齐、出现留白或末页溢出时，改写表述（删套话、拆长句、去掉重复判断，或在同一小节内移动段落），让文字按现有行距重新回填。不得把行距降到档案规定值以下，不得缩小图，不得把图内或表内字体降到 `figure_interior` / `table_cell` 的字号以下。把行距改回档案规定值是在纠正偏差；为了省出一行而再压低，则不行。

图内文字和表内文字的字号是硬性条件。表内字号由 OOXML 直接判定。图内文字若在 DrawingML 里，字号不符即失败；若只在位图里，OOXML 读不到，必须逐图核对等于 `roles.figure_interior.size_pt` 之后，才可以在审计命令上加 `--figure-fonts-verified`。公式的 MathType 尺寸必须等于 `roles.equation` 的 `mt_*_pt`，核对之后才可以加 `--mathtype-sizes-verified`。这两个开关只证明已经核对过的那一项，不能跳过其他闸门，也不能用来接受更小的字号。

下列模板要求同样是闸门。作者简介照片、定义和定理、篇首的大小写与空格写法不在审查范围内。

- 段末不能只剩一个汉字；正文里参考文献序号、以及数字后面的单位，不能出现在行首。这两项来自 `measure_and_export.ps1` 写出的 `line_breaks`，所以最终审计必须带 `--columns`。
- 从小数点起每 3 位留 1/4 字空；数字和单位之间留 1/4 字空。列项用 `1）`、`2）`。
- 图为嵌入型并居中。每条中文图题、表题都要有对应的 `Fig.`、`Tab.`，附录编号写成 `图A1` / `Fig. A1`。
- 表头里的单位写成分子在上的分数，百分数写成 `φB/%`，列与列用制表位对齐，不用连续空格撑位置。
- 参考文献编号后空 2 小格；除括号外用全角标点；条目带 GB/T 7714 的类型标志，如 `[J]`、`[M]`、`[D]`、`[R]`、`[C]`、`[S]`、`[N]`、`[EB/OL]`；外文刊名写全称，不用 `Trans.`、`Proc.`、`J.` 这类缩写。参考文献条目允许段中分页、分栏，不得勾选“段中不分页”（`keepLines`）。
- 上下紧挨的展示公式按后文引用关系合成公式框，规则见下文「连续展示公式如何分框」。一个公式框是一个段落，编号只放在该段末尾。
- 缩写第一次出现写成“中文全称(英文全称，缩写)”。3 个及以上的大写缩写，若所在段没有把它放进括号，审计失败。
- 页眉必须含档案 `rules.header_must_contain` 里的刊名，字号和居中 8.5 cm、居右 16.8 cm 制表位也要在页眉里。每行字数按页宽核算，允许和 45 字、半栏 21 字差 1 个字。
- 责任编辑一行里，“编辑”用黑体，姓名用楷体。

变量是否符合 GB 3100～3102、同一符号的大小写和上下标是否全文一致、图中主线和副线是否分开、坐标标目是否与轴平行且标值不过密、录波图是否标出每格数值和单位，这些要从图和符号本身看。逐项核对之后才可以加 `--notation-verified`。这个开关不代替上面任何自动闸门。

## 新模板如何合并

用户给出新的 `.doc` 或 `.docx` 模板时，先运行 `import-template`，再按更新后的档案审计。

合并规则：

- 模板用句子写明的数值，或在对应样本段落上设置的字体、字号、行距、缩进、页面，写入档案并覆盖同名字段。
- 模板没有出现的角色、没有写明的字段，以及 `layout_policy`，保持导入前的值。不删除，也不改回 Word 默认页边距或默认行距。
- 合并报告里的 `updated` 是这次改动的档案字段，`kept` 是继续有效的字段。审计对两边一视同仁，缺一项 `stop.allowed` 就是 false。

## Overview

1. **Template profile**: `references/csee_profile.json` is the live standard. `import-template` refreshes it and regenerates `references/csee_template_spec.md`.
2. **Layout audit**: page, document grid, headers, role fonts, line spacing, caption and table fonts, figure-interior font, equation spacing, references, biography, and two-column metrics all have gates.
3. **MathType**: special characters and adjacent units become `Equation.DSMT4`. Equation paragraph spacing follows `roles.equation` (single spacing on the current CSEE profile). Do not impose a large exact line pitch, and do not squeeze spacing below the profile.
4. **Cross-references**: `REF` / `SEQ` with `\* CHARFORMAT`.
5. **Prose**: clichés in [academic_prose_checklist.md](./references/academic_prose_checklist.md). Wording changes are also how column bottoms are filled.
6. **Word COM**: [measure_and_export.ps1](./scripts/measure_and_export.ps1) writes `col_metrics.json`. Pass that file to `audit --columns`. Single-column pages are not failed for a missing right column. Thresholds come from `layout_policy.column_abs_d_pt` and `column_bottom_min_pt`.

## Dependencies

- **`uv`**: Python 3 standard library only inside the CLI.
- **Microsoft Word**: required for `.doc` import and for `measure_and_export.ps1`.

## Quick Start

Script path: `scripts/journal_docx_cli.py` beside this file. The commands below use the installed copy.

```powershell
$cli = "C:\Users\14434\AppData\Local\Cursor\AgentStores\cursor_agent_stores\u452247633\files\skills\journal-docx-formatting-audit\scripts\journal_docx_cli.py"
$measure = "C:\Users\14434\AppData\Local\Cursor\AgentStores\cursor_agent_stores\u452247633\files\skills\journal-docx-formatting-audit\scripts\measure_and_export.ps1"

# 0. Only when the user supplies a new template. Unspecified fields stay as they are.
uv run $cli import-template `
  --input "E:\path\to\new-template.doc" `
  --report "E:\path\to\merge_report.json"

# 1. Audit. stop.allowed is false until --columns is supplied and every gate passes.
uv run $cli audit `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\audit_report.json"

# 2. MathType size and equation spacing follow the profile.
uv run $cli build-mt `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\manuscript_mt.docx" `
  --report "E:\path\to\mt_report.json"

# 3. REF/SEQ with \* CHARFORMAT
uv run $cli fix-refs `
  --input "E:\path\to\manuscript_mt.docx" `
  --output "E:\path\to\manuscript_refs.docx" `
  --report "E:\path\to\ref_report.json"

# 4. Measure. Read MaxDeltaPt and MinBottomPt from csee_profile.json layout_policy.
powershell -NoProfile -ExecutionPolicy Bypass -File $measure `
  -DocxPath "E:\path\to\manuscript_refs.docx" `
  -PdfPath "E:\path\to\verification.pdf" `
  -JsonOutPath "E:\path\to\col_metrics.json" `
  -MaxDeltaPt 5.25 `
  -MinBottomPt 745

# 5. Re-audit. Add a verified switch only after the check named in 通过条件.
uv run $cli audit `
  --input "E:\path\to\manuscript_refs.docx" `
  --columns "E:\path\to\col_metrics.json" `
  --output "E:\path\to\audit_report.json"
```

## Utility Scripts

### [journal_docx_cli.py](./scripts/journal_docx_cli.py)

- **`import-template`**: `--input` template (`.doc` or `.docx`), `--profile`, `--spec`, `--report`. Deep-merges a partial patch. Missing template fields are left untouched.
- **`audit`**: `--input`, `--output`, optional `--profile`, `--columns`, `--figure-fonts-verified`, `--mathtype-sizes-verified`. Writes `gates`, `passed`, and `stop`. `stop.allowed` is the only completion signal.
- **`build-mt`**: normalizes OLE aspect ratio and baseline. Sets displayed-equation spacing from `roles.equation` instead of a fixed `atLeast` / 3 pt padding.
- **`fix-refs`**: appends `\* CHARFORMAT` to `REF` fields.

### [template_profile.py](./scripts/template_profile.py)

Extract, merge, render the spec, and evaluate gates. `merge_profile` never deletes a leaf the new patch omits.

### [measure_and_export.ps1](./scripts/measure_and_export.ps1)

Word COM, file lock at `$env:USERPROFILE\.gemini\tmp\word_com.lock`. `-MaxDeltaPt` and `-MinBottomPt` default to 5.25 and 745; pass the current profile values after a template import changes `layout_policy`. A page with only one column is not scored. If no two-column page exists, alignment fails.

## 连续展示公式如何分框

上下紧挨、各自单独成段的展示公式（段内是 MathType 对象和制表符，不是夹在句子中的行内式）要收成若干公式框。一个公式框是一个 MathType 对象，按 n×1 矩阵的样式排成一列：两式就是 2 行 1 列，三式就是 3 行 1 列。不要加矩阵括号，不要拆成上下两个对象，也不要用 Word 换行把它们分成两段。编号只在这个对象右侧出现一次：对象前一个居中制表符，编号前一个右对齐制表符。多出来的制表符会把公式推出栏外。

先查后文的 `REF`。书签若从未被引用，就不构成单独引用需求。

- 这一组里没有哪一式需要被单独指到，或者后文只把整组当作同一个式号：合成一个公式框，编号放在最后一行末尾。
- 后文要单独指到其中某一式，或单独指到连续的某几式：只在引用边界处拆开。同一次引用共同指向的连续几式放进同一个框，共用一个编号。没有被单独引用、并且彼此相连的式子，合成它们自己的框并编号，不要并进旁边那个必须被单独指到的框。
- 某一式和它相邻的式各自都有独立引用时，这两式必须分成两个框，各编一个号。
- 例：连续 5 式，后文要把正中间的两式（第 2 式和第 3 式）作为单独的引用目标，而且这两式本身不再彼此拆开。分组大小是 1、2、2：第 1 式一个框，第 2–3 式一个框，第 4–5 式一个框，共三个框、三个编号。若第 2 式和第 3 式各自还有独立引用，则中间这两式再拆成两个框。

变量用斜体，单位、数字和文字型下标用正体。单位和普通标量不加粗；加粗只用于矢量或矩阵。

## 格式修改顺序

改稿按下面的顺序做。先做完前面一档并再跑一次 `audit`，再进入下一档。不要为了过后面的栏位检查，回过头去改已经符合档案的字体或行距。

1. **字体和字号。** 只改 `eastAsia`、`ascii` 和 `sz`：正文、各级标题、摘要和关键词、图表题、表内、页眉、责任编辑。这一档不动段前段后，不改图片大小。
2. **行距和段间距。** 把该用单倍的段落改回单倍并与网格对齐，把摘要、图表题、表内、参考文献、附录、作者简介改成档案里的固定行距（当前是 14 pt）。段前、段后只改到档案规定的磅值，例如标题 6 pt、英文单位段后 12 pt。不要为了少一行把行距再压低。
3. **缩进和对齐。** 正文首行 0.74 cm，参考文献悬挂 0.63 cm，标题悬挂为 0，图表题居中，编号公式右对齐。
4. **页面。** 纸张、页边距、栏数、栏距、文档网格、页眉刊名和两个制表位。
5. **字符级排版。** 特殊字符（希腊字母、±、≤、≥ 等）和带单位的数字都改成 MathType，不能留在普通文字里。公式的正文字号必须和它所在段落的字号相同：正文 10.5 pt，摘要、题注、参考文献 9 pt，表内 7.5 pt。同时处理千分位、列项 `1）` 和参考文献标点。
6. **公式和交叉引用。** 先按「连续展示公式如何分框」合并或拆开展示公式，再运行 `build-mt`、`fix-refs`。`build-mt` 按所在位置的字号缩放公式，不再统一缩到 0.875。公式行距仍按第 2 档。公式对象必须保持 MathType 的 `Equation.DSMT4`，不要转成 Word 自带公式。
7. **图表。** 嵌入型、居中、中英文题成对、三线表、制表位对齐。三线表只给标题行加上下框线，只给末行加下框线，数据行之间不画横线，也不画竖线。图内和表内字号停在第 1 档的 7.5 pt，不把图缩小来换栏位。
8. **表述和双栏。** 最后才改句子。栏底不齐或段末单字时，删套话、拆长句或在同一小节内移动段落，然后运行 `measure_and_export.ps1`，再带 `--columns` 审计。目视项（位图字号、GB 3100、主副线、坐标标目、录波格值）放在这一档核对，核对完才加对应开关。

## Workflow

1. If the user supplied a new template, run `import-template` and read `updated` and `kept`. Do not loosen a kept field.
2. Read [csee_template_spec.md](./references/csee_template_spec.md) and [academic_prose_checklist.md](./references/academic_prose_checklist.md). Run `audit`.
3. Repair failed gates in the order in **格式修改顺序**, starting with font and size, then line spacing. Re-run `audit` after each stage before touching figures, prose, or column breaks.
4. Run `build-mt` and `fix-refs` only after font, line spacing, indents, and page setup match the profile.
5. Run `measure_and_export.ps1`, then `audit --columns`. If a column is short, rewrite or move prose in that column. Do not compress line spacing, figures, or chart fonts.
6. Inspect the PDF. Stop only when `stop.allowed` is true.

## Concurrency

Word COM calls share `$env:USERPROFILE\.gemini\tmp\word_com.lock` (120 s, poll 1 s). Do not run `import-template` on a `.doc` and `measure_and_export.ps1` at the same time.

## Common Mistakes

1. **Stopping on a partial audit.** A report with `stop.allowed: false` is not a pass, including when the only pending gate is `columns` because `--columns` was omitted.
2. **Treating an omitted template field as permission to drop it.** Import leaves that field at its previous value, and the audit still enforces it.
3. **Starting with column fitting.** Fonts and line spacing come first. Do not move paragraphs, shrink figures, or change spacing below the profile in order to clear a later gate. Figure interior type stays at `figure_interior.size_pt` (currently 7.5 pt). Table type stays at `table_cell.size_pt` (currently 7.5 pt) with its exact line spacing.
4. **Omitting `\* CHARFORMAT`.** Caption bold then leaks into body cross-references and changes line breaks.
5. **A trailing single-column section after 作者简介.** The last `sectPr` stays the two-column section.
6. **A large exact line pitch on displayed equations.** Follow `roles.equation`. The current profile uses single spacing, not `line="646" lineRule="exact"`, and not a tighter pitch than that profile.
7. **Numbering only the last line of a stacked display.** Consecutive display equations that are not cited separately belong in one paragraph, with one `SEQ` number at the end. Do not leave the upper equation in its own unnumbered paragraph.
