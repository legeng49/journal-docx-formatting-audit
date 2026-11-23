"""编译报告并汇总问题，再把每页渲染成 PNG 供目检。

用法：python build.py <报告文件夹> [--tex main.tex] [--png-dir 目录] [--dpi 80] [--keep-aux]

- 用 latexmk -xelatex 编译；失败时打印 .log 里第一个错误前后的内容并以非零状态退出。
- 成功时列出 LaTeX/宏包警告、Overfull/Underfull、BibTeX 警告，打印页数和页码标签。
- PNG 默认写到系统临时目录下的 csee-preview/<报告文件夹名>/，不弄脏报告文件夹。
- 默认编译后用 latexmk -c 清掉辅助文件，只留 PDF；--keep-aux 保留。
"""
import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf


def run(cmd: list[str], cwd: Path) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def first_error(log: str) -> str:
    lines = log.splitlines()
    for i, line in enumerate(lines):
        if line.startswith("!"):
            return "\n".join(lines[max(0, i - 3): i + 12])
    return "\n".join(lines[-40:])


def warnings(log: str) -> list[str]:
    out = []
    lines = log.splitlines()
    for i, line in enumerate(lines):
        if re.match(r"(LaTeX|Package \S+) Warning", line):
            out.append(line.strip())
        elif line.startswith(("Overfull", "Underfull")):
            ctx = lines[i + 1].strip() if i + 1 < len(lines) else ""
            out.append(f"{line.strip()}  |  {ctx}")
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("report_dir")
    ap.add_argument("--tex", default="main.tex")
    ap.add_argument("--png-dir")
    ap.add_argument("--dpi", type=int, default=80)
    ap.add_argument("--keep-aux", action="store_true")
    a = ap.parse_args()

    d = Path(a.report_dir).resolve()
    stem = Path(a.tex).stem
    r = run(["latexmk", "-xelatex", "-interaction=nonstopmode", "-halt-on-error", a.tex], d)
    log_path = d / f"{stem}.log"
    log = log_path.read_text(encoding="utf-8", errors="replace") if log_path.exists() else ""
    if r.returncode != 0:
        print("编译失败。\n--- .log 中的错误 ---")
        print(first_error(log) if log else r.stdout[-3000:] + r.stderr[-3000:])
        sys.exit(1)

    ws = warnings(log)
    blg = d / f"{stem}.blg"
    if blg.exists():
        ws += [l.strip() for l in blg.read_text(encoding="utf-8", errors="replace").splitlines()
               if l.startswith("Warning--")]
    print(f"编译成功。警告 {len(ws)} 条")
    for w in ws:
        print("  ", w)

    pdf = d / f"{stem}.pdf"
    doc = pymupdf.open(pdf)
    png_dir = Path(a.png_dir) if a.png_dir else Path(tempfile.gettempdir()) / "csee-preview" / d.name
    png_dir.mkdir(parents=True, exist_ok=True)
    for old in png_dir.glob("page-*.png"):
        old.unlink()
    for i, page in enumerate(doc, 1):
        page.get_pixmap(dpi=a.dpi).save(png_dir / f"page-{i:02d}.png")
    print(f"PDF：{pdf}（{len(doc)} 页，页码标签 {[p.get_label() for p in doc]}）")
    print(f"PNG：{png_dir}")
    doc.close()

    if not a.keep_aux:
        c = run(["latexmk", "-c", a.tex], d)
        if c.returncode != 0:
            raise RuntimeError(f"latexmk -c 失败：{c.stderr}")
        for ext in (".bbl", ".xdv"):
            (d / f"{stem}{ext}").unlink(missing_ok=True)


if __name__ == "__main__":
    main()
