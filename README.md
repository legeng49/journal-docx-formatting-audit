# journal-docx-formatting-audit

中文与 SCI 期刊 Word 稿的格式审查、MathType 公式、交叉引用、文风收紧和双栏对齐工具。默认对照《中国电机工程学报》模板，也可以换用自定义模板。

在仓库根目录运行命令。示例里的稿件路径请换成自己的文件。

## 功能

1. **版式审查。** 检查 A4 页边距（上 3.2 cm、下 2.5 cm、左右 2.0 cm）、自英文摘要起的双栏（栏距 0.78 cm）、页眉、英文摘要断词、三线表，以及“图/表、题注、解释”的阅读顺序。文末不留空的单栏节。
2. **MathType 公式。** 把正文里的特殊字符、相邻物理单位和 `$...$` 公式写成 `Equation.DSMT4` 对象，不打开 MathType 窗口。行内公式按 1:1 对齐基线；单独成行的编号公式使用“至少”行距，段前段后 3 pt，公式居中、编号在行末。
3. **交叉引用。** 标题、图题、表题和公式编号使用 `SEQ` 与书签。正文里的“第1节”“图1”“表1”“式(1)”改成带 `\* CHARFORMAT` 的 `REF` 域，避免题注的粗体被带进正文。
4. **文风。** 按 `references/academic_prose_checklist.md` 标出拿腔用词和重复的否定句。气动近似、必要下界、开环前提和模型适用范围予以保留。
5. **双栏对齐。** 用 Word 测量每页左右栏的顶部和底部。允许误差不超过 5.25 pt（半个 10.5 pt 字符）。导出 PDF 后核对页数，不留空白页和栏底大块空白。

模板数值见 `references/csee_template_spec.md`。

## 环境

- Windows，已安装桌面版 Word
- Python 3
- [uv](https://docs.astral.sh/uv/)
- 脚本只用 Python 标准库，不需要安装 MathType 图形界面

双栏测量会启动 Word。同一时间只运行一个测量脚本；脚本用 `%USERPROFILE%\.gemini\tmp\word_com.lock` 串行加锁。

## 安装

克隆仓库：

```powershell
git clone https://github.com/legeng49/journal-docx-formatting-audit.git
cd journal-docx-formatting-audit
```

### Cursor

装到当前项目，只对该仓库生效：

```powershell
New-Item -ItemType Directory -Force -Path .cursor\skills | Out-Null
Copy-Item -Recurse -Force .\journal-docx-formatting-audit .cursor\skills\journal-docx-formatting-audit
```

如果仓库就是本 skill，则把它放到：

```text
.cursor/skills/journal-docx-formatting-audit/
```

目录里需要有 `SKILL.md`、`references/` 和 `scripts/`。

装到本机、所有项目可用：把整个目录复制到 Cursor 的个人 skills 目录，文件夹名保持 `journal-docx-formatting-audit`。Cursor 个人技能写在用户 Agent Store 的 `skills/` 下；若当前环境没有 Agent Store，则放到：

```text
%USERPROFILE%\.cursor\skills\journal-docx-formatting-audit\
```

不要放到 `%USERPROFILE%\.cursor\skills-cursor\`。那是 Cursor 自带技能的目录。

### Gemini

```powershell
$dest = Join-Path $env:USERPROFILE ".gemini\config\skills\journal-docx-formatting-audit"
New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
Copy-Item -Recurse -Force .\journal-docx-formatting-audit $dest
```

安装后新开一次对话，让代理读取该 skill 的 `SKILL.md`。

## 使用

在 skill 根目录依次执行：

```powershell
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
```

`audit` 只读稿件并写出 JSON。`build-mt` 和 `fix-refs` 写出新的 docx，不覆盖 `--input`。测量脚本会打开 Word、更新域并导出 PDF。左右栏底部差大于 5.25 pt，或栏底高于约 745 pt 的留白，按 `SKILL.md` 里的间距和段落顺序调整后再测一次。

## 文件

| 路径 | 作用 |
| --- | --- |
| `SKILL.md` | 代理执行时的完整步骤 |
| `references/csee_template_spec.md` | 学报页面、字体、公式和图表规格 |
| `references/academic_prose_checklist.md` | 需要收紧的套话，以及必须保留的物理边界 |
| `scripts/journal_docx_cli.py` | 审查、生成 MathType、写入交叉引用 |
| `scripts/mt_builder.py` | 不打开 MathType 界面，直接生成公式对象 |
| `scripts/measure_and_export.ps1` | 测量双栏坐标并导出核对用 PDF |
