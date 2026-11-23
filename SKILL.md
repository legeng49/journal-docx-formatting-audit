---
name: journal-docx-formatting-audit
description: >-
  Audits, repairs, and typesets Chinese and SCI academic journal Word (.docx)
  manuscripts (with built-in CSEE template standards and custom template
  support). Converts all special characters and adjacent physical units into
  native MathType (Equation.DSMT4) OLE objects without opening MathType UI,
  fixes inline formula baselines and displayed equation spacing, converts
  section/figure/table/equation cross-references to Word REF/SEQ fields with
  CHARFORMAT, tightens academic prose by removing clichés and redundant
  defensive disclaimers, and calibrates strict two-column top/bottom alignment
  (error <= 0.5 chars / 5.25 pt) with zero blank regions via Word COM
  coordinate measurement and PDF export verification.
---

# Journal Word (.docx) Formatting Audit & Two-Column Typesetting

## Overview

This skill provides a complete, reproducible workflow and CLI toolkit for auditing and typesetting academic journal Word (`.docx`) manuscripts (defaulting to *Proceedings of the CSEE* / 《中国电机工程学报》 specifications while supporting custom templates):

1. **Layout & Template Audit**: Verifies page margins (`3.2 cm` top, `2.5 cm` bottom, `2.0 cm` left/right), removes trailing empty single-column sections so the 2-column layout (`0.78 cm` spacing) extends cleanly to the end of the manuscript, checks odd/even running headers, enforces justified hyphenated English `ABSTRACT`, checks three-line tables (numbers and units unified in cells or headers), and verifies strict reading order (`Figure/Table -> Caption -> Interpretation`) without compressing figures or separating captions across columns.
2. **Pure-Python MathType (`Equation.DSMT4`) Engine**: Scans all plain-text runs (`<w:t>`) for leaked special characters (`±, −, ·, ≤, ≥, ∈, †, ‡`, Greek letters, etc.) and adjacent physical units, compiles `$...$` LaTeX tokens or special-character expressions directly into genuine `Equation.DSMT4` OLE containers + WMF vector previews (without launching MathType's GUI), uniformly scales existing OLE formulas (`1:1` aspect ratio) with exact WMF baseline `<w:position>` offsets, and sets displayed numbered equations to `lineRule="atLeast"` with `3 pt` (`60` twips) before/after spacing.
3. **Cross-Reference Field Automation (`REF` / `SEQ` + `\* CHARFORMAT`)**: Wraps headings (`1`, `2`, `2.2`, etc.), Chinese table/figure captions, and equation numbers in bookmarks and `SEQ` fields, and replaces hardcoded section/figure/table/equation references with `REF ... \h \* CHARFORMAT` fields so updating fields or exporting PDF never inherits bold caption fonts into body text.
4. **Academic Prose Tightening & Cliché Audit**: Detects pretentious clichés, redundant negative judgments, and over-defensive disclaimers listed in [academic_prose_checklist.md](./references/academic_prose_checklist.md) while preserving essential physical and model-scope boundaries.
5. **Word COM Two-Column Alignment & PDF Verification**: Uses [measure_and_export.ps1](./scripts/measure_and_export.ps1) to refresh fields, binary-search column-splitting paragraphs (`SPLIT_Y`), measure per-page Left/Right column top (`L_top, R_top, dT`) and bottom (`L_bot, R_bot, dB`) coordinates, export a verification PDF, and verify `|dT| <= 5.25 pt` and `|dB| <= 5.25 pt` (half a `10.5 pt` character) with zero blank regions and exact page-count parity between Word and PDF.

## Dependencies

- **`uv`**: Required to invoke the Python CLI script via `uv run` in a standard Python 3 environment. The scripts rely strictly on Python 3 standard library modules (`zipfile`, `xml.etree.ElementTree`, `struct`, `ctypes`, `re`, `json`, `zlib`, `base64`).

## Quick Start

```powershell
# 1. Run a full diagnostic audit on a manuscript and write JSON report to disk
uv run scripts/journal_docx_cli.py audit `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\audit_report.json"

# 2. Convert $...$ tokens / fix MathType OLE aspect ratios, baselines, and equation line spacing
uv run scripts/journal_docx_cli.py build-mt `
  --input "E:\path\to\manuscript.docx" `
  --output "E:\path\to\manuscript_mt.docx" `
  --report "E:\path\to\mt_report.json"

# 3. Convert section, figure, table, and equation references into Word REF/SEQ fields with \* CHARFORMAT
uv run scripts/journal_docx_cli.py fix-refs `
  --input "E:\path\to\manuscript_mt.docx" `
  --output "E:\path\to\manuscript_refs.docx" `
  --report "E:\path\to\ref_report.json"

# 4. Refresh fields, measure per-page two-column coordinates, and export verification PDF via Word COM
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/measure_and_export.ps1" `
  -DocxPath "E:\path\to\manuscript_refs.docx" `
  -PdfPath "E:\path\to\verification.pdf" `
  -JsonOutPath "E:\path\to\col_metrics.json"
```

## Utility Scripts

### 1. [journal_docx_cli.py](./scripts/journal_docx_cli.py)

Multi-command CLI tool for OOXML inspection, MathType OLE generation/repair, cross-reference injection, and prose auditing. All subcommands require `--output` (or `--report`) and write structured JSON output to files.

- **`audit`**:
  - Arguments: `--input <docx_path>` (required), `--output <json_path>` (required), `--template <csee>` (default: `csee`).
  - Checks: section margins (`pgMar`), column definitions (`cols`) and trailing empty sections, leaked special characters in `<w:t>`, post-OLE plain-text units, non-uniformly scaled or unpositioned MathType OLEs, displayed equation spacing (`atLeast` + `60` twips before/after), missing `\* CHARFORMAT` on `REF` fields, figure/table mention order vs. placement order, and forbidden phrases from [academic_prose_checklist.md](./references/academic_prose_checklist.md).
- **`build-mt`**:
  - Arguments: `--input <docx_path>` (required), `--output <docx_path>` (required), `--report <json_path>` (required), `--scale <float>` (default: `0.875`).
  - Compiles any `$...$` inline LaTeX markers in text runs into genuine `Equation.DSMT4` OLE objects using [mt_builder.py](./scripts/mt_builder.py), normalizes existing MathType OLE aspect ratios and `<w:position>` baseline offsets from WMF metrics, and sets displayed numbered equation paragraphs to `lineRule="atLeast"` with `before="60"` and `after="60"`.
- **`fix-refs`**:
  - Arguments: `--input <docx_path>` (required), `--output <docx_path>` (required), `--report <json_path>` (required).
  - Adds bookmarks and `SEQ` fields to numbered headings and Chinese captions (`表1`..`表N`, `图1`..`图N`), and converts plain-text references (`第1节`, `详见第2节`, `第2.2节`, `图1`, `表1`, `式(1)`) into Word `REF` fields with `\* CHARFORMAT`.

### 2. [mt_builder.py](./scripts/mt_builder.py)

Pure-Python MathType (`Equation.DSMT4`) generator and WMF inspector:
- `tex_to_ast(tex)`, `build_wmf(ast, pt, sub_pt, ssub_pt)`, `build_mtef5_bytes(tpl_eq_native, tex, ast)`, `replace_ole_eq_native(tpl_ole, eq_native)`, `inspect_wmf_metrics(wmf_bytes)`, and `get_default_ole_bytes()`.
- Uses Windows `gdi32` font metrics (`Times New Roman` and `Symbol`) to compute glyph bounding boxes and baseline offsets without opening MathType.

### 3. [measure_and_export.ps1](./scripts/measure_and_export.ps1)

PowerShell Word COM automation script (must be run with `BypassSandbox: true` when Word COM is blocked inside the sandbox):
- Acquires a cross-process file lock at `$env:USERPROFILE\.gemini\tmp\word_com.lock` before launching `Word.Application`.
- Calls `$doc.Fields.Update()`, `$doc.Save()`, and `$doc.Repaginate()`.
- Computes `SPLIT_Y` via binary search on any paragraph that splits across columns or pages, and writes per-page `L_top, R_top, dT, L_bot, R_bot, dB`, blank-column warnings, and `word_pages` vs. `pdf_pages` to `-JsonOutPath`.
- If `-TargetDocxPath` is specified and open in a desktop Word window, safely closes the open document handle via `GetActiveObject("Word.Application")` before copying the updated `.docx` into place.

## Workflow

1. **Read Specifications & Run Initial Audit**:
   - Consult [csee_template_spec.md](./references/csee_template_spec.md) and [academic_prose_checklist.md](./references/academic_prose_checklist.md).
   - Run `journal_docx_cli.py audit` on the user's `.docx` and inspect the JSON report.
2. **Apply Prose, MathType, and Layout Fixes**:
   - Keep essential physical boundaries (e.g., constant aerodynamic input + negligible DC-link change for $P_{\mathrm{gap}}\approx -\dot{E}_d$, open-loop assumptions before exponential recovery time estimate, model scope in a dedicated section) while removing repetitive negative judgments and pretentious clichés (`揭示了`, `直接印证`, `刚性前置约束`, `清晰解耦`, `精确刻画`, `显著...`, `更为关键的是`, `综合上述`, `可以看出/可以看到`, `源—网—储`).
   - Ensure all special characters and adjacent units are enclosed in MathType (`$...$` tokens compiled via `mt_builder.py` / `build-mt`), and put numbers + units together in Table 1 value cells (or headers) consistently.
   - Ensure strict reading order (`Figure/Table -> Caption -> Interpretation`) and keep captions attached to their figures/tables in the same column (`keepNext`).
   - Run `build-mt` and `fix-refs` to normalize all MathType OLEs, displayed equation spacing (`atLeast`, `3 pt` before/after), and `REF ... \* CHARFORMAT` cross-references.
3. **Measure Two-Column Alignment & Eliminate Blank Regions**:
   - Run `measure_and_export.ps1` (`BypassSandbox: true`) and read the resulting `col_metrics.json`.
   - If a tall figure jumps to the next column leaving a blank region (`L_bot` or `R_bot < 745 pt`), move the paragraph immediately following the figure before the figure (within the same subsection, preserving `Figure -> Caption -> Interpretation` order) so the text fills the column bottom.
   - If `|dB| > 5.25 pt` on any page (`41` vs `42` lines due to `atLeast` equation heights), add `40–220` twips (`2–11 pt`) of `before`/`after` spacing to a heading, displayed equation, or caption inside the shorter column so `L_bot` and `R_bot` align within `<= 5.25 pt` without triggering a line wrap.
   - Visually inspect the exported verification `.pdf` using `read_file` to confirm `word_pages == pdf_pages`, zero blank areas, and clean typography.

## Concurrency & Rate Limiting

- Local `Word.Application` COM calls are serialized via the lock file `$env:USERPROFILE\.gemini\tmp\word_com.lock` (timeout `120 s`, poll interval `1 s`). Never run multiple `measure_and_export.ps1` instances in parallel without the lock.

## Common Mistakes

1. **Omitting `\* CHARFORMAT` on `REF` fields**: Without `\* CHARFORMAT`, Word's `ExportAsFixedFormat` copies the `9 pt` bold `黑体` font from Chinese figure/table captions into body-text cross-references (`图4`, `表1`), altering line widths during PDF export and causing 1-line page overflows on the final page. Always include `\* CHARFORMAT` and run `$doc.Fields.Update()` before measuring coordinates.
2. **Leaving a trailing single-column Section 3**: Do not append an empty continuous single-column section break after `作者简介`. End `<w:body>` directly with the 2-column `<w:sectPr>` (`num="2"`, `space="442"`) and calibrate the final page's `参考文献` / `作者简介` spacing so both columns reach the bottom margin (`~751–754 pt`).
3. **Fixed `32.3 pt` line spacing on displayed equations**: Do not use `line="646" lineRule="exact"` on single-line displayed equations (`~17 pt` tall), which creates excessive vertical whitespace. Use `lineRule="atLeast"` with `before="60"` (`3 pt`) and `after="60"` (`3 pt`), and absorb sub-grid remainders via in-column `before`/`after` padding.
