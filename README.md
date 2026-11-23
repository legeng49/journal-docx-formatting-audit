# journal-docx-formatting-audit

中文与 SCI 期刊 Word 稿的格式审查、MathType 公式、交叉引用、文风收紧和双栏对齐工具。默认对照《中国电机工程学报》模板。数值以 `references/csee_profile.json` 为准。

上传新的 `.doc` 或 `.docx` 模板时，只覆盖模板里写明或在样本上设置过的字段。模板没有的字段保持原值，不退回 Word 默认值，也不从档案中删除。

修改顺序是字体字号、行距和段间距、缩进对齐、页面、字符、公式与交叉引用、图表，最后才改表述并测双栏。栏底不齐时改写表述，不压低行距，不缩小图，不把图内或表内字体降到档案规定的字号以下。`audit` 报告里的 `stop.allowed` 为 true 之前，不算通过。

## 功能

1. **模板合并。** `import-template` 读取新模板，写出合并报告里的 `updated` 和 `kept`。同一结果只改 `latex/assets/template/csee-report.cls` 里对应的页边距、栏距、行距、公式尺寸和刊名；类文件其余部分保持不动。
2. **版式审查。** 页面、文档网格、页眉刊名与制表位、角色字体和行距、图表字号、嵌入型居中的图、中英文题成对、参考文献著录、段末单字和行首禁排。图内字号、MathType 尺寸，以及 GB 3100、主副线、坐标标目、录波格值，需要核对后再打开对应开关。
3. **MathType 公式。** 特殊字符和相邻单位进入 `Equation.DSMT4`。编号公式的行距按档案里的 `roles.equation`，当前模板是单倍行距。
4. **交叉引用。** `REF` 域带 `\* CHARFORMAT`。
5. **文风。** 按 `references/academic_prose_checklist.md` 收紧套话。版面不够时同样用改写解决。
6. **双栏对齐。** Word 测量左右栏顶底。允许误差和栏底下限在 `layout_policy` 里，默认 5.25 pt 和 745 pt。单栏页不按缺右栏判失败。
7. **LaTeX 课程报告。** `latex/` 用 XeLaTeX 写机报正文。`python latex/scripts/new_report.py <报告文件夹>` 新建，`python latex/scripts/build.py <报告文件夹>` 编译。要求交 Word 时用 `python latex/scripts/to_word.py <报告文件夹> --pdf`。写法见 `latex/references/format.md`。

规格说明由档案生成，见 `references/csee_template_spec.md`。

## 环境

- Windows，已安装桌面版 Word
- Python 3
- [uv](https://docs.astral.sh/uv/)
- 脚本只用 Python 标准库，不需要安装 MathType 图形界面

`.doc` 导入和双栏测量都会启动 Word。同一时间只运行一个 Word 脚本；脚本用 `%USERPROFILE%\.gemini\tmp\word_com.lock` 串行加锁。

## 安装

克隆仓库：

```powershell
git clone https://github.com/legeng49/journal-docx-formatting-audit.git
cd journal-docx-formatting-audit
```

### Cursor

个人技能放在用户 Agent Store 的 `skills/journal-docx-formatting-audit/`。当前这台机器的目录是：

```text
%LOCALAPPDATA%\Cursor\AgentStores\cursor_agent_stores\u452247633\files\skills\journal-docx-formatting-audit
```

若当前环境没有 Agent Store，则放到：

```text
%USERPROFILE%\.cursor\skills\journal-docx-formatting-audit\
```

不要放到 `%USERPROFILE%\.cursor\skills-cursor\`。

只对某一个项目生效时，复制到该项目的 `.cursor/skills/journal-docx-formatting-audit/`。目录里需要有 `SKILL.md`、`references/` 和 `scripts/`。

### Gemini

```powershell
$dest = Join-Path $env:USERPROFILE ".gemini\config\skills\journal-docx-formatting-audit"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\* $dest
```

安装后新开一次对话，让代理读取该 skill 的 `SKILL.md`。

## 使用

在 skill 根目录依次执行：

```powershell
uv run scripts/journal_docx_cli.py import-template `
  --input "E:\path\to\new-template.doc" `
  --report "E:\path\to\merge_report.json"

uv run scripts/journal_docx_cli.py audit `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\audit_report.json"

uv run scripts/journal_docx_cli.py build-mt `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\manuscript_mt.docx" `
  --report "E:\path\to\mt_report.json"

uv run scripts/journal_docx_cli.py fix-refs `
  --input "E:\path\to\manuscript_mt.docx" `
  --output "E:\path\to\manuscript_refs.docx" `
  --report "E:\path\to\ref_report.json"

powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/measure_and_export.ps1" `
  -DocxPath "E:\path\to\manuscript_refs.docx" `
  -PdfPath "E:\path\to\verification.pdf" `
  -JsonOutPath "E:\path\to\col_metrics.json"

uv run scripts/journal_docx_cli.py audit `
  --input "E:\path\to\manuscript_refs.docx" `
  --columns "E:\path\to\col_metrics.json" `
  --output "E:\path\to\audit_report.json"
```

`audit` 只读稿件并写出 JSON。`build-mt` 和 `fix-refs` 写出新的 docx，不覆盖 `--input`。`--figure-fonts-verified`、`--mathtype-sizes-verified` 和 `--notation-verified` 只在 `SKILL.md` 写明的核对完成之后使用，不能用来跳过其他闸门。

## 文件

| 路径 | 作用 |
| --- | --- |
| `SKILL.md` | 代理执行时的完整步骤和停止条件 |
| `references/csee_profile.json` | 当前模板档案；导入时按字段合并 |
| `references/csee_template_spec.md` | 由档案生成的规格说明 |
| `references/academic_prose_checklist.md` | 需要收紧的套话，以及必须保留的物理边界 |
| `scripts/journal_docx_cli.py` | 导入模板、审查、MathType、交叉引用 |
| `scripts/template_profile.py` | 模板字段提取、合并和闸门判定 |
| `scripts/mt_builder.py` | 不打开 MathType 界面，直接生成公式对象 |
| `scripts/measure_and_export.ps1` | 测量双栏、段末单字和行首禁排，并导出 PDF |
| `latex/` | XeLaTeX 课程报告：类文件、GB/T 7714 样式、新建、编译和转 Word |
