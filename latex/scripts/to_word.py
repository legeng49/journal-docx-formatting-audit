"""把 csee-report 的 LaTeX 报告转成 Word（.docx），版式按机报 Word 模板。

用法：python to_word.py <报告文件夹> [--tex main.tex] [-o 输出.docx] [--pdf]

流程：
1. latexmk 编译一遍，从 .aux 取文献编号和交叉引用编号，从 .bbl 取 gbt7714 排好的参考文献；
2. 把正文改写成 pandoc 认识的 LaTeX（标题编号、引用、公式/图表编号都换成实际数字，
   自定义命令换成 ⟦…⟧ 标记），用 pandoc 转成 docx；
3. 用 python-docx 按标记补版式：通栏题名 + 双栏正文分节、首页/奇偶页页眉、中英文图表题、
   公式右侧编号、三线表、参考文献悬挂缩进、基金项目脚注。
只出正文（题名页即第 1 页），不做封面、封底。
--pdf 用本机 Word 把结果另存为 PDF（同名 .word.pdf），用来对照检查。

支持的写法即模板 main.tex 里用到的：\\section 三级标题、\\appendix、equation、align、gather、
multline（含带星号的，编号规则同 amsmath）、figure/figure*、table/table*（tabular/fulltabular +
booktabs）、\\bicaption、\\cite/\\citenum、\\ref/\\eqref、enumerate/itemize、\\num/\\SI、acknowledgment。
"""
import argparse
import copy
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Emu
from lxml import etree

SKILL = Path(__file__).resolve().parent.parent


# 版心（twip）：A4 宽 11906，左右各 1134；双栏栏距 442（0.78 cm）
TEXT_W = 11906 - 2 * 1134
COL_W = (TEXT_W - 442) // 2
COL_CM = COL_W / 567
TEXT_CM = TEXT_W / 567
QSP = "\u2005"  # 1/4 em 空格：数字分组、数与单位之间
ENSP = chr(0x2002)
ZWSP = chr(0x200B)  # 零宽空格：只提供断行点

SECTION_LEVELS = ("section", "subsection", "subsubsection")


# ====================================================================== LaTeX 解析
def strip_comments(src: str) -> str:
    return "\n".join(re.sub(r"(?<!\\)%.*$", "", line) for line in src.splitlines())


def group_end(s: str, i: int, open_ch="{", close_ch="}") -> int:
    """s[i] 是开括号，返回匹配的闭括号下标。"""
    assert s[i] == open_ch, (s[i:i + 30], open_ch)
    depth = 0
    for j in range(i, len(s)):
        c = s[j]
        if c == "\\":
            continue
        if s[j - 1] == "\\" and j > 0:
            continue
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return j
    raise ValueError(f"括号不配对：{s[i:i + 60]}")


def macro_args(src: str, name: str) -> list[str]:
    """所有 \\name{...} 的参数（只取第一个必选参数）。"""
    out = []
    for m in re.finditer(r"\\" + name + r"\s*\{", src):
        a = m.end() - 1
        out.append(src[a + 1:group_end(src, a)])
    return out


def macro_arg(src: str, name: str, default: str = "") -> str:
    args = macro_args(src, name)
    return args[-1].strip() if args else default


def replace_macro(src: str, name: str, nargs: int, fn) -> str:
    """把 \\name{a}{b}... 替换成 fn(a, b, ...)。"""
    out, pos = [], 0
    pat = re.compile(r"\\" + name + r"(?![A-Za-z@])\s*")
    while True:
        m = pat.search(src, pos)
        if not m:
            break
        j = m.end()
        args = []
        for _ in range(nargs):
            while src[j] in " \n":
                j += 1
            if src[j] == "[":  # 可选参数，丢弃
                j = group_end(src, j, "[", "]") + 1
                while src[j] in " \n":
                    j += 1
            e = group_end(src, j)
            args.append(src[j + 1:e])
            j = e + 1
        out.append(src[pos:m.start()])
        out.append(fn(*args))
        pos = j
    out.append(src[pos:])
    return "".join(out)


def read_aux(aux: Path) -> tuple[dict, dict]:
    text = aux.read_text(encoding="utf-8", errors="replace")
    cites = {m.group(1): int(m.group(2)) for m in re.finditer(r"\\bibcite\{([^}]*)\}\{\{(\d+)\}", text)}
    labels = {}
    for m in re.finditer(r"\\newlabel\{([^}]*)\}\{\{", text):
        a = m.end() - 1
        labels[m.group(1)] = text[a + 1:group_end(text, a)]
    return cites, labels


def read_bbl(bbl: Path) -> list[tuple[str, str]]:
    """[(key, LaTeX 文本)]，按 .bbl 中的顺序（即引用编号顺序）。"""
    s = bbl.read_text(encoding="utf-8", errors="replace")
    end = s.index(r"\end{thebibliography}")
    items = []
    for m in re.finditer(r"\\bibitem", s):
        j = m.end()
        if s[j] == "[":
            j = group_end(s, j, "[", "]") + 1
        e = group_end(s, j)
        items.append((s[j + 1:e], e + 1))
    out = []
    for n, (key, start) in enumerate(items):
        stop = items[n + 1][1] if n + 1 < len(items) else end
        if n + 1 < len(items):
            stop = s.rindex(r"\bibitem", 0, stop)
        body = s[start:stop]
        body = re.sub(r"\\newblock\s*", " ", body)
        # \allowbreak 处换成零宽空格，Word 才能在“[EB/OL]．(日期)[日期]．网址”这类无空格的长串里断行
        body = re.sub(r"\\allowbreak\s*", ZWSP, body)
        body = re.sub(r"(\\url\{)", ZWSP + r"\1", body)
        body = " ".join(body.split())
        body = re.sub(r"([．，：；])\s+", r"\1", body)  # 全角标点后不留空格
        body = re.sub(r"([．，：；])(?!" + ZWSP + ")", r"\1" + ZWSP, body)  # 但可以在其后断行
        out.append((key, body))
    return out


def group_digits(x: str) -> str:
    """与 cls 的 siunitx 设置一致（机报模板）：整数、小数部分各自 4 位及以上时每 3 位加 1/4 em 空。"""
    x = x.strip()
    m = re.fullmatch(r"([+-]?)(\d*)(?:\.(\d+))?(.*)", x)
    if not m:
        return x
    sign, ip, fp, rest = m.groups()
    if len(ip) >= 4:
        ip = re.sub(r"(?<=\d)(?=(\d{3})+$)", QSP, ip)
    if fp and len(fp) >= 4:
        fp = QSP.join(fp[i:i + 3] for i in range(0, len(fp), 3))
    return sign + ip + ("." + fp if fp else "") + rest


def compress(nums: list[int]) -> str:
    nums = sorted(set(nums))
    parts, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts.append(str(nums[i]) if i == j else f"{nums[i]}-{nums[j]}")
        i = j + 1
    return ",".join(parts)


class Numberer:
    """按 csee-report.cls 的规则给标题、公式、图表编号，并记录 \\label。"""

    def __init__(self):
        self.sec = [0, 0, 0]
        self.eq = self.fig = self.tab = 0
        self.appendix = False
        self.current = ""
        self.labels = {}

    def thesection(self) -> str:
        return chr(ord("A") + self.sec[0] - 1) if self.appendix else str(self.sec[0])

    def prefix(self) -> str:
        return self.thesection() if self.appendix else ""

    def heading(self, level: int) -> str:
        self.sec[level] += 1
        for k in range(level + 1, 3):
            self.sec[k] = 0
        if level == 0 and self.appendix:
            self.eq = self.fig = self.tab = 0
        nums = [self.thesection()] + [str(n) for n in self.sec[1:level + 1]]
        self.current = ".".join(nums)
        return self.current

    def step(self, kind: str) -> str:
        setattr(self, kind, getattr(self, kind) + 1)
        self.current = self.prefix() + str(getattr(self, kind))
        return self.current


# ====================================================================== 行间公式
def scan_depths(s: str, sep: str | None = None):
    """在 s 的顶层（不在 {} / \\begin…\\end / \\left…\\right 之内）按 sep（"\\\\" 或 "&"）切分。
    返回 (各段, 结尾处的 (花括号, 环境, left-right) 深度)。"""
    parts, start, i = [], 0, 0
    brace = env = lr = 0
    while i < len(s):
        if s.startswith("\\begin{", i):
            env += 1
            i += 7
            continue
        if s.startswith("\\end{", i):
            env -= 1
            i += 5
            continue
        if s.startswith("\\left", i) and not s[i + 5:i + 6].isalpha():
            lr += 1
            i += 5
            continue
        if s.startswith("\\right", i) and not s[i + 6:i + 7].isalpha():
            lr -= 1
            i += 6
            continue
        top = brace == 0 and env == 0 and lr == 0
        c = s[i]
        if c == "\\":
            if s[i + 1:i + 2] == "\\" and sep == "\\\\" and top:
                parts.append(s[start:i])
                i += 2
                opt = re.match(r"\s*\[[^\]]*\]", s[i:])  # \\[2pt]
                if opt:
                    i += opt.end()
                start = i
                continue
            i += 2  # 转义字符（\& \{ 等）或命令开头
            continue
        if c == "{":
            brace += 1
        elif c == "}":
            brace -= 1
        elif c == "&" and sep == "&" and top:
            parts.append(s[start:i])
            start = i + 1
        i += 1
    parts.append(s[start:])
    return parts, (brace, env, lr)


def balanced(s: str) -> bool:
    return scan_depths(s)[1] == (0, 0, 0)


def math_block(env: str, starred: bool, inner: str, num: "Numberer") -> str:
    """equation / align / gather / multline / eqnarray（含带星号的）改写成逐行的公式标记。
    每行：⟦EQ n⟧ 前是该行公式；align 类在第一个 & 处拆成左右两半，中间插 ⟦EQSPLIT⟧，
    左半为空时用 ⟦EQBLANK⟧ 占位。n 为 - 表示不编号。编号规则同 amsmath：
    带星号不编号，align/gather/eqnarray 每行编号，\\nonumber/\\notag 的行不编号。"""
    def take_labels(s, n):
        for lab in re.findall(r"\\label\{([^}]*)\}", s):
            if n is None:
                raise ValueError(f"\\label{{{lab}}} 写在了不编号的公式里")
            num.labels[lab] = n
        return re.sub(r"\\label\{[^}]*\}", "", s)

    def math(s):
        return f"\n\n\\[{s.strip()}\\]\n\n"

    def row(content, n, left=None):
        mark = f"⟦EQ {n or '-'}⟧"
        if left is None:
            return math(content) + mark + "\n\n"
        # 右半以“=”开头时 Word 不在它左边留关系符间距，补一个 \; 到左半末尾
        left_part = math(left + r"\;") if left.strip() else "\n\n⟦EQBLANK⟧\n\n"
        return left_part + "⟦EQSPLIT⟧" + math(content) + mark + "\n\n"

    if env in ("equation", "multline"):
        n = None if starred else num.step("eq")
        content = take_labels(inner, n)
        if env == "multline":
            content = "\\begin{gathered}" + content + "\\end{gathered}"
        return row(content, n)

    out = []
    for line in scan_depths(inner, "\\\\")[0]:
        if not line.strip():
            continue
        nonum = re.search(r"\\(nonumber|notag)(?![A-Za-z])", line)
        line = re.sub(r"\\(nonumber|notag)(?![A-Za-z])", "", line)
        n = None if (starred or nonum) else num.step("eq")
        line = take_labels(line, n)
        if env == "gather":
            out.append(row(line, n))
            continue
        cells = scan_depths(line, "&")[0]
        left, right = cells[0], "".join(cells[1:])
        if len(cells) > 1 and balanced(left) and balanced(right):
            out.append(row(right, n, left=left))
        else:
            out.append(row("".join(cells), n))
    return "".join(out)


def mark_group_rules(tab: str) -> str:
    """表体中间的 \\midrule、\\cmidrule（分组线）pandoc 会丢掉：删去它们，在紧接的那一行行尾加 ⟦HR⟧，
    转成 Word 后在该行上方补 0.5 pt 横线（\\cmidrule 也画成整行）。表头下的第一条 \\midrule 是栏目线，保留。"""
    rules = list(re.finditer(r"\\midrule|\\cmidrule(?:\([^)]*\))?\{[^}]*\}", tab))
    first_mid = next((r for r in rules if r.group(0) == "\\midrule"), None)
    for r in reversed(rules):
        if r is first_mid:
            continue
        end = tab.find("\\\\", r.end())
        if end < 0:
            raise ValueError(f"表格中 {r.group(0)} 之后没有表格行")
        tab = tab[:r.start()] + tab[r.end():end] + "⟦HR⟧" + tab[end:]
    return tab


# ====================================================================== 正文改写
def convert_body(body: str, cites: dict, aux_labels: dict, img_dir: Path, report: Path,
                 refs: list[tuple[str, str]]) -> str:
    num = Numberer()

    # \verb 里的内容原样保留，不参与下面的命令替换
    verbs = []

    def keep_verb(m):
        verbs.append(m.group(0))
        return f"⟦V{len(verbs) - 1}⟧"
    body = re.sub(r"\\verb\*?(.)(.*?)\1", keep_verb, body)

    # 列表改成普通段落：“1）……”“• ……”，与机报列项格式一致
    def lists(m):
        env, content = m.group(1), m.group(2)
        items = [it.strip() for it in re.split(r"\\item\b", content)[1:]]
        lead = (lambda k: f"{k}）") if env == "enumerate" else (lambda k: "•" + "\u2002")
        return "\n\n" + "\n\n".join(lead(k) + it for k, it in enumerate(items, 1)) + "\n\n"
    body = re.sub(r"\\begin\{(enumerate|itemize)\}(?:\[[^\]]*\])?(.*?)\\end\{\1\}", lists, body, flags=re.S)

    body = re.sub(r"\\begin\{acknowledgment\}", "\n\n⟦ACK⟧\n\n", body)
    body = re.sub(r"\\end\{acknowledgment\}", "\n\n", body)

    # 依出现顺序编号；label 归属于它之前最近一次编号
    body = re.sub(r"\\(begin|end)\{fulltabular\}", r"\\\1{tabular}", body)
    body = re.sub(r"(\\begin\{tabular\}\{(?:[^{}]|\{[^{}]*\})*\})(.*?)(\\end\{tabular\})",
                  lambda m: m.group(1) + mark_group_rules(m.group(2)) + m.group(3), body, flags=re.S)

    token = re.compile(
        r"\\(?P<sec>section|subsection|subsubsection)(?P<star>\*?)\s*\{"
        r"|(?P<app>\\appendix\b)"
        r"|\\begin\{(?P<math>equation|align|gather|multline|eqnarray)(?P<mstar>\*?)\}"
        r"|\\begin\{(?P<fbeg>figure\*?|table\*?)\}"
        r"|\\end\{(?P<fend>figure\*?|table\*?)\}"
        r"|(?P<cap>\\bicaption)\s*\{"
        r"|\\label\{(?P<label>[^}]*)\}"
    )
    out, pos, float_stack = [], 0, []
    while True:
        m = token.search(body, pos)
        if not m:
            break
        out.append(body[pos:m.start()])
        if m["sec"]:  # 标题
            level = SECTION_LEVELS.index(m["sec"])
            a = m.end() - 1
            title = body[a + 1:group_end(body, a)].strip()
            pos = group_end(body, a) + 1
            cmd = "\\" + m["sec"] + "*"
            if m["star"]:
                out.append(f"{cmd}{{{title}}}")
            else:
                n = num.heading(level)
                if num.appendix:
                    text = (f"附录{ENSP}{n}" + (f"\u2003{title}" if title else "") if level == 0
                            else n + ("\u2003" if level != 1 else "\u2002") + title)
                else:  # 正文标题用 Word 自动编号（模板要求），编号格式见 add_heading_numbering
                    text = f"⟦HN {level}⟧{title}"
                out.append(f"{cmd}{{{text}}}")
            continue
        if m["app"]:
            num.appendix = True
            num.sec = [0, 0, 0]
        elif m["math"]:
            end_tag = f"\\end{{{m['math']}{m['mstar']}}}"
            e = body.index(end_tag, m.end())
            out.append(math_block(m["math"], bool(m["mstar"]), body[m.end():e], num))
            pos = e + len(end_tag)
            # 公式后不空行接着写（如“式中：”）时，LaTeX 不缩进，Word 里也顶格
            if not re.match(r"[ \t]*\n[ \t]*\n|\s*$", body[pos:]):
                out.append("⟦NI⟧")
            continue
        elif m["fbeg"]:
            env = m["fbeg"]
            float_stack.append(env)
            if env.endswith("*"):
                out.append("\n\n⟦WIDE⟧\n\n")
            out.append(f"\\begin{{{env.rstrip('*')}}}")
        elif m["fend"]:
            env = m["fend"]
            float_stack.pop()
            out.append(f"\\end{{{env.rstrip('*')}}}")
            if env.endswith("*"):
                out.append("\n\n⟦/WIDE⟧\n\n")
        elif m["cap"]:
            kind = float_stack[-1].rstrip("*")
            n = num.step("fig" if kind == "figure" else "tab")
            a = m.end() - 1
            cn_end = group_end(body, a)
            cn = body[a + 1:cn_end]
            b = cn_end + 1
            while body[b] in " \n":
                b += 1
            en_end = group_end(body, b)
            en = body[b + 1:en_end]
            out.append(f"\\caption{{⟦CAP {float_stack[-1]} {n}⟧{cn}⟦EN⟧{en}}}")
            pos = en_end + 1
            continue
        elif m["label"] is not None:
            num.labels[m["label"]] = num.current
        pos = m.end()
    out.append(body[pos:])
    body = "".join(out)

    for key, n in num.labels.items():
        if key in aux_labels:
            aux_n = re.sub(r"[{}]", "", aux_labels[key])
            if aux_n != n:
                raise ValueError(f"\\label{{{key}}} 编号不一致：LaTeX 为 {aux_n}，转换脚本算出 {n}")

    def ref(key):
        if key not in num.labels:
            raise KeyError(f"\\ref{{{key}}} 没有对应的 \\label")
        return num.labels[key]
    body = replace_macro(body, "eqref", 1, lambda k: f"({ref(k.strip())})")
    body = replace_macro(body, "ref", 1, lambda k: ref(k.strip()))

    def cite_nums(keys):
        ks = [k.strip() for k in keys.split(",") if k.strip()]
        missing = [k for k in ks if k not in cites]
        if missing:
            raise KeyError(f"未在 .aux 中找到引用：{missing}")
        return [cites[k] for k in ks]
    body = replace_macro(body, "citenum", 1, lambda k: compress(cite_nums(k)))
    for name in ("cite", "citep"):
        body = replace_macro(body, name, 1, lambda k: f"\\textsuperscript{{[{compress(cite_nums(k))}]}}")

    body = replace_macro(body, "num", 1, group_digits)
    for name in ("SI", "qty"):
        body = replace_macro(body, name, 2, lambda x, u: group_digits(x) + QSP + f"\\si{{{u}}}")

    # 插图：PDF 转 PNG；宽度按栏宽/版心宽换成 cm
    def graphic(m):
        opts, name = m.group(1) or "", m.group(2).strip()
        path = find_graphic(name, report)
        natural = None  # 没写 width 时按自然尺寸插入（与 LaTeX 一致），PDF 取页面宽度
        if path.suffix.lower() == ".pdf":
            page = pymupdf.open(path)[0]
            natural = page.rect.width / 72 * 2.54
            png = img_dir / (path.stem + ".png")
            page.get_pixmap(dpi=300).save(png)
            path = png
        else:
            shutil.copy2(path, img_dir / path.name)
            path = img_dir / path.name
        w = re.search(r"width\s*=\s*([\d.]*)\s*\\(columnwidth|linewidth|textwidth)", opts)
        if w:
            frac = float(w.group(1) or 1)
            width = frac * (TEXT_CM if w.group(2) == "textwidth" else COL_CM)
        else:
            wabs = re.search(r"width\s*=\s*([\d.]+)\s*(cm|mm|in|pt|bp)", opts)
            if wabs:
                val, unit = float(wabs.group(1)), wabs.group(2)
                width = val * {"cm": 1, "mm": 0.1, "in": 2.54, "pt": 2.54 / 72.27, "bp": 2.54 / 72}[unit]
            else:
                width = min(natural, TEXT_CM) if natural else COL_CM
        return f"\\includegraphics[width={width:.2f}cm]{{{path.name}}}"
    body = re.sub(r"\\includegraphics\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}", graphic, body)

    # 参考文献插在 \bibliography 处；网址按模板排成普通文字（不做成蓝色超链接）
    def plain_url(m):
        return re.sub(r"([_%#&])", r"\\\1", m.group(1))
    ref_tex = "⟦REFHEAD⟧参考文献\n\n" + "\n\n".join(
        f"⟦REF {n}⟧" + re.sub(r"\\url\{([^}]*)\}", plain_url, text)
        for n, (key, text) in enumerate(refs, 1)) + "\n\n"
    if not re.search(r"\\bibliography\{[^}]*\}", body):
        raise ValueError("正文中没有 \\bibliography{...}")
    body = re.sub(r"\\bibliography\{[^}]*\}", lambda _: ref_tex, body)

    for cmd in ("centering", "noindent", "clearpage", "newpage", "raggedright"):
        body = re.sub(r"\\" + cmd + r"(?![A-Za-z])", "", body)
    body = re.sub(r"\\vspace\*?\{[^}]*\}", "", body)
    return re.sub(r"⟦V(\d+)⟧", lambda m: verbs[int(m.group(1))], body)


def find_graphic(name: str, report: Path) -> Path:
    img_exts = (".pdf", ".png", ".jpg", ".jpeg")
    exts = [""] if Path(name).suffix.lower() in img_exts else list(img_exts)
    for base in (report / "figures", report):
        for ext in exts:
            p = base / (name + ext)
            if p.is_file():
                return p
    for ext in exts:
        r = subprocess.run(["kpsewhich", name + ext], capture_output=True, text=True)
        if r.returncode == 0 and r.stdout.strip():
            return Path(r.stdout.strip())
    raise FileNotFoundError(f"找不到插图：{name}")


# ====================================================================== pandoc 与参考样式
def pandoc_exe() -> str:
    exe = shutil.which("pandoc")
    if exe:
        return exe
    local = Path.home() / "AppData" / "Local" / "Pandoc" / "pandoc.exe"
    if local.is_file():
        return str(local)
    raise FileNotFoundError("找不到 pandoc；安装：winget install --id JohnMacFarlane.Pandoc -e")


RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
             "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
             "w:effect", "w:vertAlign", "w:lang"]


def sort_children(el, order):
    """按 OOXML 规定的顺序重排子元素（Word 对顺序敏感）。"""
    kids = sorted(list(el), key=lambda e: next((i for i, t in enumerate(order) if e.tag == qn(t)), len(order)))
    for k in kids:
        el.remove(k)
        el.append(k)


def set_fonts(rpr, latin=None, cjk=None):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = etree.SubElement(rpr, qn("w:rFonts"))
        rpr.insert(0, rf)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        rf.attrib.pop(qn(a), None)
    if latin:
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(a), latin)
    if cjk:
        rf.set(qn("w:eastAsia"), cjk)


def set_child(parent, tag, attrs=None, order_first=False):
    el = parent.find(qn(tag))
    if el is None:
        el = etree.SubElement(parent, qn(tag))
        if order_first:
            parent.insert(0, el)
    el.attrib.clear()
    for k, v in (attrs or {}).items():
        el.set(qn(k), str(v))
    return el


def style_rpr(style):
    return style.element.get_or_add_rPr()


def style_ppr(style):
    return style.element.get_or_add_pPr()


def make_reference(path: Path):
    data = subprocess.run([pandoc_exe(), "--print-default-data-file", "reference.docx"],
                          capture_output=True, check=True).stdout
    path.write_bytes(data)
    doc = Document(path)
    st = doc.styles

    # 文档默认：Times New Roman + 宋体，五号
    rpr_default = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    set_fonts(rpr_default, "Times New Roman", "宋体")
    set_child(rpr_default, "w:sz", {"w:val": 21})
    set_child(rpr_default, "w:szCs", {"w:val": 21})
    set_child(rpr_default, "w:lang", {"w:val": "en-US", "w:eastAsia": "zh-CN"})

    def para(name, *, latin="Times New Roman", cjk="宋体", size=10.5, bold=False, jc="both",
             first_chars=0, before=0, after=0, exact=None, keep_next=False):
        s = st[name]
        rpr = style_rpr(s)
        for tag in ("w:color", "w:b", "w:bCs", "w:i", "w:iCs"):
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        set_fonts(rpr, latin, cjk)
        set_child(rpr, "w:sz", {"w:val": int(size * 2)})
        set_child(rpr, "w:szCs", {"w:val": int(size * 2)})
        if bold:
            etree.SubElement(rpr, qn("w:b"))
        # 显式关掉斜体，免得从 pandoc 的 Caption 基样式继承
        etree.SubElement(rpr, qn("w:i")).set(qn("w:val"), "0")
        etree.SubElement(rpr, qn("w:iCs")).set(qn("w:val"), "0")
        sort_children(rpr, RPR_ORDER)
        ppr = style_ppr(s)
        for tag in ("w:spacing", "w:ind", "w:jc", "w:keepNext", "w:outlineLvl", "w:numPr"):
            el = ppr.find(qn(tag))
            if el is not None:
                ppr.remove(el)
        sp = {"w:before": int(before * 20), "w:after": int(after * 20)}
        if exact:
            sp.update({"w:line": int(exact * 20), "w:lineRule": "exact"})
        else:
            sp.update({"w:line": 240, "w:lineRule": "auto"})
        if keep_next:
            ppr.append(parse_xml(f'<w:keepNext {nsdecls("w")}/>'))
        ppr.append(parse_xml(f'<w:spacing {nsdecls("w")} ' + " ".join(f'{k}="{v}"' for k, v in sp.items()) + "/>"))
        if first_chars:
            ppr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLineChars="{first_chars}" '
                                 f'w:firstLine="{int(size * 20 * first_chars / 100)}"/>'))
        else:  # firstLineChars 会从基样式继承并压过 firstLine，要一起清零
            ppr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLineChars="0" w:firstLine="0"/>'))
        ppr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))
        # pPr 子元素顺序：keepNext … spacing ind … jc
        order = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:widowControl", "w:numPr",
                 "w:pBdr", "w:shd", "w:tabs", "w:snapToGrid", "w:spacing", "w:ind", "w:jc", "w:outlineLvl", "w:rPr"]
        kids = sorted(list(ppr), key=lambda e: next((i for i, t in enumerate(order) if e.tag == qn(t)), 99))
        for k in kids:
            ppr.remove(k)
            ppr.append(k)

    para("Normal")
    para("Body Text", first_chars=200)
    para("First Paragraph", first_chars=200)
    para("Compact", jc="center", size=7.5)
    para("Heading 1", latin="Arial", cjk="黑体", size=12, jc="left", before=6, after=6, keep_next=True)
    para("Heading 2", latin="Arial", cjk="黑体", size=10.5, jc="left", keep_next=True)
    para("Heading 3", latin="Times New Roman", cjk="宋体", size=10.5, jc="left", keep_next=True)
    para("Image Caption", cjk="黑体", size=9, bold=True, jc="center", exact=14)
    para("Table Caption", cjk="黑体", size=9, bold=True, jc="center", exact=14)
    para("Captioned Figure", jc="center", before=3, after=3)  # 模板：图形段前段后空 3 pt
    para("Figure", jc="center", before=3, after=3)
    para("Footnote Text", size=7.5, first_chars=200)
    doc.save(path)


# ====================================================================== docx 后处理工具
def p_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def remove_text(p, s: str):
    """从段落中删去字符串 s（可以跨 run）。"""
    ts = list(p.iter(qn("w:t")))
    full = "".join(t.text or "" for t in ts)
    i = full.index(s)
    j = i + len(s)
    pos = 0
    for t in ts:
        txt = t.text or ""
        a, b = pos, pos + len(txt)
        if b > i and a < j:
            t.text = txt[:max(0, i - a)] + txt[min(len(txt), j - a):]
        pos = b


def split_at(p, token: str):
    """把段落在 token 处拆成两段，返回新段（token 删去）。"""
    for r in list(p):
        if r.tag != qn("w:r"):
            continue
        t = r.find(qn("w:t"))
        if t is None or token not in (t.text or ""):
            continue
        before, after = t.text.split(token, 1)
        t.text = before
        r2 = copy.deepcopy(r)
        r2.find(qn("w:t")).text = after
        r.addnext(r2)
        newp = etree.Element(qn("w:p"))
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            newp.append(copy.deepcopy(ppr))
        moving = [r2] + list(r2.itersiblings())
        for el in moving:
            newp.append(el)
        p.addnext(newp)
        return newp
    raise ValueError(f"段落中找不到 {token}：{p_text(p)[:40]}")


def make_run(text, *, latin=None, cjk=None, size=None, bold=None, superscript=False):
    r = etree.Element(qn("w:r"))
    rpr = etree.SubElement(r, qn("w:rPr"))
    if latin or cjk:
        set_fonts(rpr, latin, cjk)
    if bold is not None:
        etree.SubElement(rpr, qn("w:b")).set(qn("w:val"), "1" if bold else "0")
    if superscript:
        etree.SubElement(rpr, qn("w:vertAlign")).set(qn("w:val"), "superscript")
    if size:
        etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), str(int(size * 2)))
        etree.SubElement(rpr, qn("w:szCs")).set(qn("w:val"), str(int(size * 2)))
    sort_children(rpr, RPR_ORDER)
    t = etree.SubElement(r, qn("w:t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def tab_run():
    r = etree.Element(qn("w:r"))
    etree.SubElement(r, qn("w:tab"))
    return r


def ppr_of(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    return ppr


PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:widowControl", "w:numPr",
             "w:pBdr", "w:shd", "w:tabs", "w:snapToGrid", "w:spacing", "w:ind", "w:jc", "w:outlineLvl",
             "w:rPr", "w:sectPr"]


def set_ppr(p, *, jc=None, before=None, after=None, exact=None, ind=None, tabs=None, style=None, keep_next=False):
    ppr = ppr_of(p)
    if style is not None:
        set_child(ppr, "w:pStyle", {"w:val": style})
    if keep_next:
        set_child(ppr, "w:keepNext")
    if tabs is not None:
        old = ppr.find(qn("w:tabs"))
        if old is not None:
            ppr.remove(old)
        ppr.append(parse_xml(f'<w:tabs {nsdecls("w")}>' + "".join(
            f'<w:tab w:val="{v}" w:pos="{pos}"/>' for v, pos in tabs) + "</w:tabs>"))
    if before is not None or after is not None or exact is not None:
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            sp = etree.SubElement(ppr, qn("w:spacing"))
        if before is not None:
            sp.set(qn("w:before"), str(int(before * 20)))
        if after is not None:
            sp.set(qn("w:after"), str(int(after * 20)))
        if exact is not None:
            sp.set(qn("w:line"), str(int(exact * 20)))
            sp.set(qn("w:lineRule"), "exact")
    if ind is not None:
        # 样式里的 firstLineChars 优先于 firstLine，不显式清零的话缩进还在
        set_child(ppr, "w:ind", {"w:firstLineChars": 0, **ind} if "w:firstLine" in ind else ind)
    if jc is not None:
        set_child(ppr, "w:jc", {"w:val": jc})
    kids = sorted(list(ppr), key=lambda e: next((i for i, t in enumerate(PPR_ORDER) if e.tag == qn(t)), 99))
    for k in kids:
        ppr.remove(k)
        ppr.append(k)


def format_runs(p, *, latin=None, cjk=None, size=None, bold=None):
    for r in p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("w:rPr"))
            r.insert(0, rpr)
        if latin or cjk:
            set_fonts(rpr, latin, cjk)
        if bold is not None:
            for tag in ("w:b", "w:bCs"):
                el = rpr.find(qn(tag))
                if el is not None:
                    rpr.remove(el)
            if bold:
                etree.SubElement(rpr, qn("w:b"))
        if size:
            for tag in ("w:sz", "w:szCs"):
                set_child(rpr, tag, {"w:val": int(size * 2)})
        sort_children(rpr, RPR_ORDER)
    # 公式里的字号
    if size:
        for mr in p.iter(qn("m:r")):
            wrpr = mr.find(qn("w:rPr"))
            if wrpr is None:
                wrpr = etree.SubElement(mr, qn("w:rPr"))
                t = mr.find(qn("m:t"))
                if t is not None:
                    mr.remove(t)
                    mr.append(t)
            for tag in ("w:sz", "w:szCs"):
                set_child(wrpr, tag, {"w:val": int(size * 2)})


def sectpr_xml(*, cols=1, start="nextPage", top=1814, bottom=1417, header=1134, footer=0,
               title_pg=False, pg_start=None, grid=323, grid_type="lines") -> str:
    return (f'<w:sectPr {nsdecls("w")}><w:type w:val="{start}"/>'
            f'<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
            f'<w:pgMar w:top="{top}" w:right="1134" w:bottom="{bottom}" w:left="1134" '
            f'w:header="{header}" w:footer="{footer}" w:gutter="0"/>'
            + (f'<w:pgNumType w:start="{pg_start}"/>' if pg_start is not None else "")
            + f'<w:cols w:num="{cols}" w:space="442"/>'
            + ("<w:titlePg/>" if title_pg else "")
            + f'<w:docGrid w:type="{grid_type}" w:linePitch="{grid}"/></w:sectPr>')


def attach_sectpr(p, xml: str):
    ppr = ppr_of(p)
    old = ppr.find(qn("w:sectPr"))
    if old is not None:
        ppr.remove(old)
    ppr.append(parse_xml(xml))


def empty_para():
    return etree.Element(qn("w:p"))


# ====================================================================== 文字宽度
def text_width(s: str, size: float) -> float:
    w = 0.0
    for c in s:
        if ord(c) > 0x2E80:
            w += size
        elif c == " ":
            w += size * 0.25
        elif c.isupper() or c.isdigit():
            w += size * (0.5 if c.isdigit() else 0.72)
        else:
            w += size * 0.5
    return w


# ====================================================================== 页眉
def page_field() -> list:
    """PAGE 域（小五 Times New Roman）。用完整域而不是 fldSimple：fldSimple 更新后会丢掉字号。"""
    rpr = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
           '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>')
    parts = ['<w:fldChar w:fldCharType="begin"/>', '<w:instrText xml:space="preserve"> PAGE </w:instrText>',
             '<w:fldChar w:fldCharType="separate"/>', '<w:t>1</w:t>', '<w:fldChar w:fldCharType="end"/>']
    return [parse_xml(f'<w:r {nsdecls("w")}>{rpr}{x}</w:r>') for x in parts]


def fill_header_runs(p, text_parts):
    for part in text_parts:
        if part == "\t":
            p.append(tab_run())
        elif part == "PAGE":
            for r in page_field():
                p.append(r)
        else:
            p.append(make_run(part, latin="Times New Roman", cjk="宋体", size=9))


def header_para(header, text_parts, *, tabs, border):
    p = header.paragraphs[0]._p
    for r in list(p):
        if r.tag != qn("w:pPr"):
            p.remove(r)
    fill_header_runs(p, text_parts)
    set_ppr(p, tabs=tabs, jc="left", ind={"w:firstLine": 0}, exact=None)
    ppr = ppr_of(p)
    for tag in ("w:pStyle",):
        el = ppr.find(qn(tag))
        if el is not None:
            ppr.remove(el)
    if border:
        val, sz = border
        bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="{val}" w:sz="{sz}" w:space="1" w:color="000000"/></w:pBdr>')
        ppr.insert(0, bdr)
        set_ppr(p)
    return p


def build_headers(section, info):
    cn = "\u3000".join(info["journalcn"])
    tabs = [("center", 4819), ("right", 9525)]
    section.different_first_page_header_footer = True
    # 首页：两行，下方双细线
    first = section.first_page_header
    first.is_linked_to_previous = False
    p1 = header_para(first, ["\t", cn], tabs=tabs, border=None)
    p2 = copy.deepcopy(p1)
    for r in list(p2):
        if r.tag != qn("w:pPr"):
            p2.remove(r)
    fill_header_runs(p2, ["\t", info["journalen"], "\t", "PAGE"])
    ppr_of(p2).insert(0, parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>'))
    set_ppr(p2)
    p1.addnext(p2)
    # 偶数页：左页码、中刊名；奇数页：中 oddheader、右页码；下方 0.75 pt 单线
    even = section.even_page_header
    even.is_linked_to_previous = False
    header_para(even, ["PAGE", "\t", cn], tabs=tabs, border=("single", 6))
    odd = section.header
    odd.is_linked_to_previous = False
    header_para(odd, ["\t", info["oddheader"], "\t", "PAGE"], tabs=tabs, border=("single", 6))


# ====================================================================== 标题编号、作者简介
def add_heading_numbering(doc) -> int:
    """正文标题的多级自动编号：1 / 1.1 / 1.1.1，编号后分别接全角空、半角空、全角空，
    与 LaTeX 版的间距一致。返回 numId。"""
    numbering = doc.part.numbering_part.element
    abs_ids = [int(a.get(qn("w:abstractNumId"))) for a in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    aid, nid = max(abs_ids, default=0) + 1, max(num_ids, default=0) + 1
    texts = ["%1 ", "%1.%2 ", "%1.%2.%3 "]
    lvls = "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        f'<w:suff w:val="nothing"/><w:lvlText w:val="{t}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>'
        for i, t in enumerate(texts))
    abstract = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}">'
                         f'<w:multiLevelType w:val="multilevel"/>{lvls}</w:abstractNum>')
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/></w:num>'))
    return nid


def bio_block(doc, marker, bio):
    """作者简介：marker 段改作 4 cm、0.75 pt 的短线，其后一张无框两列表：左列照片（2.2 cm × 2.8 cm）
    和姓名，右列“作者简介：”与简介正文。小五宋体 / Times New Roman，行距 14 pt，首行缩进 0.63 cm。"""
    photo, name, text = bio
    for r in marker.findall(qn("w:r")):
        marker.remove(r)
    ppr = ppr_of(marker)
    ppr.insert(0, parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="1" '
                            f'w:color="000000"/></w:pBdr>'))
    set_ppr(marker, style="Normal", before=6, exact=3, ind={"w:left": 0, "w:right": COL_W - 2268, "w:firstLine": 0})

    left_w = 1304  # 2.3 cm：照片 2.2 cm + 0.1 cm 间隔
    table = doc.add_table(rows=1, cols=2)
    tbl = table._tbl
    tblpr = tbl.find(qn("w:tblPr"))
    for tag in ("w:tblStyle", "w:tblW", "w:tblLook"):
        el = tblpr.find(qn(tag))
        if el is not None:
            tblpr.remove(el)
    no_border = "".join(f'<w:{s} w:val="nil"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    tblpr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{COL_W}" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblBorders {nsdecls("w")}>{no_border}</w:tblBorders>'))
    tblpr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblpr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/>'
                           f'<w:right w:w="0" w:type="dxa"/></w:tblCellMar>'))
    for g, w in zip(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")), (left_w, COL_W - left_w)):
        g.set(qn("w:w"), str(w))
    left, right = table.rows[0].cells
    left.width, right.width = Emu(left_w * 635), Emu((COL_W - left_w) * 635)
    # 单元格显式无边框、顶端对齐：否则继承参考文档默认表格样式的底线和垂直居中
    for tc in (left._tc, right._tc):
        tcpr = tc.get_or_add_tcPr()
        tcpr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}>' + "".join(
            f'<w:{s} w:val="nil"/>' for s in ("top", "left", "bottom", "right")) + '</w:tcBorders>'))
        tcpr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="top"/>'))
        sort_children(tcpr, ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
                             "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign"])

    pic_p = left.paragraphs[0]
    pic_p.add_run().add_picture(str(photo), width=Cm(2.2), height=Cm(2.8))
    # 照片和姓名在左列的前 2.2 cm 内居中（右缩进 0.1 cm 即与正文的间隔）
    set_ppr(pic_p._p, style="Normal", jc="center", ind={"w:left": 0, "w:right": left_w - 1247, "w:firstLine": 0})
    name_p = left.add_paragraph()._p
    name_p.append(make_run(name, latin="Times New Roman", cjk="宋体", size=9))
    set_ppr(name_p, style="Normal", jc="center", exact=14, ind={"w:left": 0, "w:right": left_w - 1247, "w:firstLine": 0})

    label_p = right.paragraphs[0]._p
    label_p.append(make_run("作者简介：", cjk="黑体", size=9))
    body_p = right.add_paragraph()._p
    body_p.append(make_run(text, latin="Times New Roman", cjk="宋体", size=9))
    for p in (label_p, body_p):  # 左对齐：简介末尾的邮箱很长，两端对齐会把前一行撑散
        set_ppr(p, style="Normal", jc="left", exact=14, ind={"w:firstLine": 357})
    marker.addnext(tbl)


# ====================================================================== 组装
def postprocess(body_docx: Path, out: Path, info: dict):
    doc = Document(body_docx)
    body = doc.element.body
    final_sect = body.find(qn("w:sectPr"))
    if final_sect is not None:
        body.remove(final_sect)

    paras = [e for e in body if e.tag == qn("w:p")]
    title_sect_para = None
    table_width = {}
    eq_markers, eq_tags = set(), {}
    split_rows = []  # (表格, 左半, 右半)：align 类公式行，最后按整块重新分配左右格宽
    heading_num = add_heading_numbering(doc)
    for p in paras:
        text = p_text(p)
        m = re.match(r"⟦([A-Z/]+)(?: ([^⟧]*))?⟧", text)
        if not m:
            continue
        tag, arg = m.group(1), m.group(2)
        remove_text(p, m.group(0))
        # 题名区用固定行距，间距与 LaTeX 版（按机报模板实测）一致；单位与英文题名之间留出“空一行”
        if tag == "META":
            # 首页编号行：小五，项目名黑体、内容 Times New Roman，各项之间空 4 格（2 个全角空格）
            for r in p.findall(qn("w:r")):
                p.remove(r)
            for k, (label, value) in enumerate(info["meta"]):
                if k:
                    p.append(make_run("　　", size=9))
                p.append(make_run(f"{label}：", cjk="黑体", size=9))
                p.append(make_run(value, latin="Times New Roman", cjk="宋体", size=9))
            set_ppr(p, style="Normal", jc="left", exact=16.2, ind={"w:firstLine": 0})
        elif tag == "HN":
            # 正文标题：Word 多级自动编号（1 / 1.1 / 1.1.1），编号后的空格由编号格式给出
            level = int(arg)
            ppr = ppr_of(p)
            ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{level}"/>'
                                 f'<w:numId w:val="{heading_num}"/></w:numPr>'))
            set_ppr(p)
            first_t = next(t for t in p.iter(qn("w:t")) if t.text)
            first_t.text = first_t.text.lstrip()
        elif tag == "REFHEAD":
            # “参考文献”：小四黑体，段前段后 4 pt，左对齐（按模板实测）
            set_ppr(p, style="Normal", jc="left", before=4, after=4, keep_next=True, ind={"w:firstLine": 0})
            format_runs(p, latin="Arial", cjk="黑体", size=12)
        elif tag == "BIO":
            bio_block(doc, p, info["bio"])
        elif tag == "TITLE":
            set_ppr(p, style="Normal", jc="center", before=12, exact=30)
            format_runs(p, latin="Arial", cjk="黑体", size=22)
        elif tag == "AUTHOR":
            set_ppr(p, style="Normal", jc="center", before=8, exact=22)
            format_runs(p, latin="Times New Roman", cjk="仿宋", size=14)
        elif tag == "AFFIL":
            set_ppr(p, style="Normal", jc="center", before=15, exact=16.2)
            format_runs(p, latin="Times New Roman", cjk="楷体", size=10.5)
        elif tag == "ENTITLE":
            set_ppr(p, style="Normal", jc="center", before=52, exact=16.2)
            format_runs(p, latin="Times New Roman", size=12, bold=True)
        elif tag == "ENAUTHOR":
            set_ppr(p, style="Normal", jc="center", exact=16.2)
            format_runs(p, latin="Times New Roman", size=10.5)
        elif tag == "ENAFFIL":
            set_ppr(p, style="Normal", jc="center", after=12, exact=14)
            format_runs(p, latin="Times New Roman", size=9)
            title_sect_para = p
        elif tag in ("ENABS", "ENKW", "CNABS", "CNKW"):
            # 模板：ABSTRACT 段前不空（英文单位段后已空 12 pt），其余三段段首空 6 pt
            set_ppr(p, style="Normal", jc="both", exact=14, before=0 if tag == "ENABS" else 6,
                    ind={"w:firstLine": 0})
            format_runs(p, size=9)
            if tag in ("CNABS", "CNKW"):
                first = next(r for r in p.iter(qn("w:r")) if p_text(r))
                rpr = first.find(qn("w:rPr"))
                set_fonts(rpr, None, "黑体")
                for b in rpr.findall(qn("w:b")) + rpr.findall(qn("w:bCs")):
                    rpr.remove(b)
        elif tag == "BODY":
            body.remove(p)
        elif tag in ("EQSPLIT", "EQBLANK"):
            eq_markers.add(p)  # 由后面的 ⟦EQ⟧ 一起处理
            eq_tags[p] = tag
        elif tag == "EQ":
            # 行间公式放进无框表格：公式独立成段，保持分式等的显示样式，编号放右格。
            # 和编号写在同一段里的公式会被 Word 当成行内公式缩小排。
            prev = p.getprevious()
            omp = prev.find(qn("m:oMathPara"))
            if omp is None:
                raise ValueError(f"公式 ({arg}) 前一段不是行间公式")
            used = [prev, p]
            sp = prev.getprevious()
            if sp is not None and sp in eq_markers and "EQSPLIT" in eq_tags.get(sp, ""):
                left = sp.getprevious()
                left_omp = None
                if not (left in eq_markers and eq_tags.get(left) == "EQBLANK"):
                    left_omp = left.find(qn("m:oMathPara"))
                    if left_omp is None:
                        raise ValueError(f"公式 ({arg}) 的左半不是行间公式")
                used += [sp, left]
                tbl = equation_table(omp, arg, left_omp=left_omp, split=True)
                split_rows.append((tbl, left_omp, omp))
            else:
                tbl = equation_table(omp, arg)
            used[-1].addprevious(tbl)
            for el in used:
                body.remove(el)
        elif tag == "CAP":
            kind, n = arg.split()
            label_cn, label_en = ("图", "Fig.") if kind.startswith("figure") else ("表", "Tab.")
            en = split_at(p, "⟦EN⟧")
            if kind.startswith("table"):
                tbl = en.getnext()
                if tbl is None or tbl.tag != qn("w:tbl"):
                    raise ValueError(f"表 {n} 的题名后面不是表格")
                table_width[tbl] = TEXT_W if kind.endswith("*") else COL_W
            p.insert(1, make_run(f"{label_cn}{n}\u2002\u2002"))
            en.insert(1, make_run(f"{label_en} {n}\u2002\u2002"))
            format_runs(p, size=9, bold=True)
            format_runs(en, latin="Times New Roman", size=9, bold=True)
            # 中英文题名合成一段、中间换行并设段中不分页：双栏在分节处平衡时 Word 不管
            # “与下段同页”，分成两段的话中文题名和英文题名会被拆到两栏
            p.append(parse_xml(f'<w:r {nsdecls("w")}><w:br/></w:r>'))
            for r in en.findall(qn("w:r")):
                p.append(r)
            en.getparent().remove(en)
            set_child(ppr_of(p), "w:keepLines")  # set_ppr 会按 PPR_ORDER 排好子元素顺序
            set_ppr(p, after=3, keep_next=kind.startswith("table"))
            # 图、表题与图、表不分页：表题段后接表格，图片段后接图题
            if not kind.startswith("table"):
                set_ppr(p.getprevious(), keep_next=True)
        elif tag == "REF":
            # 含网址的条目左对齐（模板的电子文献示例如此），否则两端对齐会把长网址前的一行撑散
            jc = "left" if "://" in p_text(p) else "both"
            p.insert(1, make_run(f"[{arg}]"))
            p.insert(2, tab_run())
            set_ppr(p, style="Normal", jc=jc, exact=14, tabs=[("left", 357)],
                    ind={"w:left": 357, "w:hanging": 357})
            format_runs(p, size=9)
        elif tag == "ACK":
            p.append(make_run("致\u3000谢"))
            set_ppr(p, style="Normal", jc="center", before=6, after=6)
            format_runs(p, latin="Times New Roman", cjk="黑体", size=10.5)
        elif tag == "NI":
            set_ppr(p, ind={"w:firstLine": 0})
            first_t = next(t for t in p.iter(qn("w:t")) if t.text)
            first_t.text = first_t.text.lstrip()
        elif tag == "WIDE":
            attach_sectpr(p, sectpr_xml(cols=2, start="continuous"))
        elif tag == "/WIDE":
            attach_sectpr(p, sectpr_xml(cols=1, start="continuous"))
        else:
            raise ValueError(f"未知标记 ⟦{tag}⟧")
    if title_sect_para is None:
        raise ValueError("没有找到题名页内容")
    fit_split_equations(split_rows)
    attach_sectpr(title_sect_para, sectpr_xml(cols=1, start="nextPage", title_pg=True, pg_start=1))

    # 三线表：表宽等于栏宽（跨栏表为版心宽），列宽按内容宽度分配（见 fit_widths），固定列宽
    for tbl in body.iter(qn("w:tbl")):
        if tbl.find(qn("w:tblPr")).find(qn("w:tblStyle")) is None:  # 公式编号用的无框表格
            continue
        format_table(tbl, table_width.get(tbl, COL_W))

    # 正文的双栏节即文档最后一节
    body.append(parse_xml(sectpr_xml(cols=2, start="continuous")))

    # 基金项目脚注：不显示编号
    fix_footnotes(doc)

    doc.settings.odd_and_even_pages_header_footer = True
    set_compat(doc)
    secs = doc.sections
    build_headers(secs[0], info)
    doc.save(out)


EQ_NUM_W = 680  # 公式编号格宽（twip），左右各留一格保证公式居中


def equation_table(right_omp, number: str, left_omp=None, split=False):
    """一行公式的无框表格，四列：留白 | 左半 | 右半 | 编号。
    不拆分时左右两格合并、公式居中；align 的行在 & 处拆开，左半右对齐、右半左对齐，
    相邻几行的表格网格相同，Word 会把它们接成一张表，对齐位置一致。"""
    half = (COL_W - 2 * EQ_NUM_W) // 2
    no_border = "".join(f'<w:{s} w:val="nil"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))

    def cell(width, jc, span=1):
        gs = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
        return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{gs}<w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="0" w:after="0"/>'
                f'<w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="{jc}"/></w:pPr></w:p></w:tc>')
    body_cells = (cell(half, "right") + cell(half, "left")) if split else cell(2 * half, "center", span=2)
    xml = (f'<w:tbl {nsdecls("w", "m")}><w:tblPr><w:tblW w:w="{COL_W}" w:type="dxa"/><w:jc w:val="center"/>'
           f'<w:tblBorders>{no_border}</w:tblBorders><w:tblLayout w:type="fixed"/>'
           f'<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
           f'<w:tblLook w:val="0000"/></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="{EQ_NUM_W}"/><w:gridCol w:w="{half}"/><w:gridCol w:w="{half}"/>'
           f'<w:gridCol w:w="{EQ_NUM_W}"/></w:tblGrid>'
           f'<w:tr>{cell(EQ_NUM_W, "left")}{body_cells}{cell(EQ_NUM_W, "right")}</w:tr></w:tbl>')
    tbl = parse_xml(xml)
    ps = tbl.findall(f".//{qn('w:p')}")

    def put(p, omp, jc):
        jc_el = omp.find(qn("m:oMathParaPr"))
        if jc_el is not None and jc_el.find(qn("m:jc")) is not None:
            jc_el.find(qn("m:jc")).set(qn("m:val"), jc)
        p.append(omp)
    if split:
        if left_omp is not None:
            put(ps[1], left_omp, "right")
        else:  # 转行后的续行：模板要求明显缩进，对应 LaTeX 的 &\quad（pandoc 会丢掉行首的 \quad）
            set_child(ps[2].find(qn("w:pPr")), "w:ind", {"w:left": 210, "w:firstLineChars": 0, "w:firstLine": 0})
        put(ps[2], right_omp, "left")
    else:
        put(ps[1], right_omp, "center")
    if number != "-":
        ps[-1].append(make_run(f"({number})", latin="Times New Roman"))
    return tbl


def omml_width(omp) -> int:
    """行间公式的大致宽度（twip）：五号字，上下标按 0.7 倍，关系符和二元运算符两侧各加 1/4 em。"""
    if omp is None:
        return 0
    small = {qn("m:sub"), qn("m:sup")}
    pt = 0.0
    for t in omp.iter(qn("m:t")):
        size = 10.5 * (0.7 if any(a.tag in small for a in t.iterancestors()) else 1)
        txt = t.text or ""
        pt += text_width(txt, size) + sum(size * 0.5 for c in txt if c in "=+−-<>≈≤≥×")
    return int(pt * 20)


def fit_split_equations(split_rows):
    """align 类公式：同一块（相邻的几行）用同一个拆分位置。左格宽 = 留白 + 最宽左半，
    右格宽 = 最宽右半 + 留白，整块居中；估计宽度占满 85% 以上时左格只留 3 pt，
    余量都给右格（估计有误差，宁可不居中也不能截断或折行）。等分时长的右半会被截掉。"""
    avail = COL_W - 2 * EQ_NUM_W
    blocks, cur = [], []
    for row in split_rows:
        if cur and cur[-1][0].getnext() is not row[0]:
            blocks.append(cur)
            cur = []
        cur.append(row)
    if cur:
        blocks.append(cur)
    for block in blocks:
        lw = max(omml_width(left) for _, left, _ in block)
        rw = max(omml_width(right) for _, _, right in block)
        if lw + rw <= 0.85 * avail:
            left_w = lw + (avail - lw - rw) // 2
        else:
            left_w = lw + 60
        right_w = avail - left_w
        for tbl, _, _ in block:
            cols = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
            cols[1].set(qn("w:w"), str(left_w))
            cols[2].set(qn("w:w"), str(right_w))
            tcs = tbl.find(qn("w:tr")).findall(qn("w:tc"))
            tcs[1].find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(left_w))
            tcs[2].find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(right_w))


TABLE_FONT = 7.5   # 六号
CELL_MAR = 57      # 单元格左右内边距 0.1 cm（twip）


def fit_widths(natural: list, avail: int) -> list:
    """把各列内容宽度分配到表宽 avail。放得下时按比例放大；放不下时，不超过平均份额的
    短列（如“作用时机”这类项目列）保持内容宽度不折行，只按比例压缩其余长列。"""
    if sum(natural) <= avail:
        return [int(w * avail / sum(natural)) for w in natural]
    keep = set()
    while True:
        rest = [i for i in range(len(natural)) if i not in keep]
        share = (avail - sum(natural[i] for i in keep)) / len(rest)
        short = [i for i in rest if natural[i] <= share]
        if not short:
            break
        keep.update(short)
    budget = avail - sum(natural[i] for i in keep)
    long_total = sum(natural[i] for i in rest)
    return [int(natural[i]) if i in keep else int(natural[i] * budget / long_total) for i in range(len(natural))]


def format_table(tbl, avail: int):
    tblpr = tbl.find(qn("w:tblPr"))
    for tag in ("w:tblBorders", "w:jc", "w:tblW", "w:tblLayout", "w:tblCellMar"):
        el = tblpr.find(qn(tag))
        if el is not None:
            tblpr.remove(el)
    rows = tbl.findall(qn("w:tr"))
    # 分组线：带 ⟦HR⟧ 标记的行，上方画 0.5 pt 线（先去掉标记，再量列宽）
    ruled_rows = set()
    for k, tr in enumerate(rows):
        marked = [p for p in tr.iter(qn("w:p")) if "⟦HR⟧" in p_text(p)]
        for p in marked:
            remove_text(p, "⟦HR⟧")
        if marked:
            ruled_rows.add(k)
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    natural = [0.0] * len(grid)
    cells = []  # (tc, 行号, 起始列, 跨列数)
    for k, tr in enumerate(rows):
        col = 0
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            span_el = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            cells.append((tc, k, col, span))
            if span == 1:
                for p in tc.findall(qn("w:p")):
                    txt = "".join(t.text or "" for t in p.iter(qn("w:t"), qn("m:t")))
                    natural[col] = max(natural[col], text_width(txt, TABLE_FONT) * 20 + 2 * CELL_MAR + 60)
            col += span
    natural = [max(w, 300) for w in natural]
    widths = fit_widths(natural, avail)
    for g, w in zip(grid, widths):
        g.set(qn("w:w"), str(w))

    tblpr.find(qn("w:tblStyle")).addnext(parse_xml(
        f'<w:tblW {nsdecls("w")} w:w="{sum(widths)}" w:type="dxa"/>'))
    tblpr.find(qn("w:tblW")).addnext(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    tblpr.find(qn("w:jc")).addnext(parse_xml(
        f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:left w:val="nil"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'))
    tblpr.find(qn("w:tblBorders")).addnext(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblpr.find(qn("w:tblLayout")).addnext(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}><w:left w:w="{CELL_MAR}" w:type="dxa"/>'
        f'<w:right w:w="{CELL_MAR}" w:type="dxa"/></w:tblCellMar>'))

    line ='w:val="single" w:sz="4" w:space="0" w:color="000000"'
    for tc, k, col, span in cells:
        tcpr = tc.find(qn("w:tcPr"))
        if tcpr is None:
            tcpr = etree.Element(qn("w:tcPr"))
            tc.insert(0, tcpr)
        for tag in ("w:tcW", "w:tcBorders", "w:vAlign"):
            el = tcpr.find(qn(tag))
            if el is not None:
                tcpr.remove(el)
        tcpr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="{sum(widths[col:col + span])}" w:type="dxa"/>'))
        borders = (f"<w:top {line}/>" if k in ruled_rows else "") + \
                  (f"<w:bottom {line}/>" if k == 0 else "")  # 表头下为栏目线 0.5 pt
        if borders:
            tcpr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}>{borders}</w:tcBorders>'))
        tcpr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
        sort_children(tcpr, ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
                             "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
                             "w:hideMark"])
    # 整张表不跨页：除最后一行外每行的段落都“与下段同页”，行内不断开
    for k, tr in enumerate(rows):
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = etree.Element(qn("w:trPr"))
            tr.insert(0, trpr)
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.insert(0, etree.Element(qn("w:cantSplit")))
        for p in tr.iter(qn("w:p")):
            set_ppr(p, exact=14, jc="center", ind={"w:firstLine": 0}, keep_next=k < len(rows) - 1)
            format_runs(p, size=TABLE_FONT)


def set_compat(doc):
    """沿用机报 Word 模板的兼容选项：表格内行高对齐网格、下划线包括行尾空格、中西文字宽平衡等。"""
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = etree.SubElement(settings, qn("w:compat"))
    flags = ["w:spaceForUL", "w:balanceSingleByteDoubleByteWidth", "w:doNotLeaveBackslashAlone",
             "w:ulTrailSpace", "w:doNotExpandShiftReturn", "w:adjustLineHeightInTable", "w:useFELayout"]
    for k, tag in enumerate(flags):
        if compat.find(qn(tag)) is None:
            compat.insert(k, etree.Element(qn(tag)))
    # 标点挤压（中文 Word 的默认值）
    csc = settings.find(qn("w:characterSpacingControl"))
    if csc is None:
        raise ValueError("settings.xml 缺少 characterSpacingControl")
    csc.set(qn("w:val"), "compressPunctuation")


def fix_footnotes(doc):
    refs = list(doc.element.body.iter(qn("w:footnoteReference")))
    if not refs:
        return
    for fr in refs:
        fr.set(qn("w:customMarkFollows"), "1")
    part = next(p for p in doc.part.package.iter_parts() if str(p.partname) == "/word/footnotes.xml")
    is_xml_part = hasattr(part, "_element")
    root = part._element if is_xml_part else etree.fromstring(part.blob)
    for fn in root.findall(qn("w:footnote")):
        if fn.get(qn("w:type")):
            continue
        for r in list(fn.iter(qn("w:r"))):
            if r.find(qn("w:footnoteRef")) is not None:
                r.getparent().remove(r)
        for p in fn.findall(qn("w:p")):
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = etree.SubElement(p, qn("w:pPr"))
                p.insert(0, ppr)
            set_child(ppr, "w:pStyle", {"w:val": "FootnoteText"})
            runs = [r for r in p.findall(qn("w:r")) if p_text(r)]
            if runs and "基金项目" in p_text(runs[0]):
                rpr = runs[0].find(qn("w:rPr"))
                if rpr is None:
                    rpr = etree.SubElement(runs[0], qn("w:rPr"))
                    runs[0].insert(0, rpr)
                set_fonts(rpr, None, "黑体")
                for b in rpr.findall(qn("w:b")) + rpr.findall(qn("w:bCs")):
                    rpr.remove(b)
    if not is_xml_part:
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ====================================================================== 入口
def latex_build(report: Path, tex: str):
    r = subprocess.run(["latexmk", "-xelatex", "-interaction=nonstopmode", "-halt-on-error", tex],
                       cwd=report, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise RuntimeError(f"LaTeX 编译失败，先用 build.py 查错：\n{r.stdout[-2000:]}")


def word_to_pdf(docx: Path, pdf: Path):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(str(docx), False, True)
        d.Fields.Update()
        d.SaveAs2(str(pdf), 17)
        d.Close(False)
    finally:
        word.Quit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report_dir")
    ap.add_argument("--tex", default="main.tex")
    ap.add_argument("-o", "--output")
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()

    report = Path(a.report_dir).resolve()
    stem = Path(a.tex).stem
    latex_build(report, a.tex)
    cites, aux_labels = read_aux(report / f"{stem}.aux")
    refs = read_bbl(report / f"{stem}.bbl")
    for n, (key, _) in enumerate(refs, 1):
        if cites.get(key) != n:
            raise ValueError(f"文献 {key} 在 .bbl 中排第 {n}，.aux 中编号为 {cites.get(key)}")

    src = strip_comments((report / a.tex).read_text(encoding="utf-8"))
    begin = src.index(r"\begin{document}")
    pre, doc_body = src[:begin], src[begin:]
    j = doc_body.index(r"\maketitle") + len(r"\maketitle")
    k = doc_body.index(r"\end{document}")
    main_body = doc_body[j:k]

    def field(name, default=""):
        return macro_arg(pre, name, default)

    jh = re.search(r"\\journalheader\s*\{([^}]*)\}\s*\{([^}]*)\}", pre)
    title = field("title")
    info = {
        "journalcn": re.sub(r"\\[A-Za-z]+\s*|\s", "", jh.group(1)) if jh else "中国电机工程学报",
        "journalen": jh.group(2) if jh else "Proceedings of the CSEE",
        "oddheader": re.sub(r"\\[A-Za-z]+\*?|[{}]", "", field("oddheader")),
        # 首页编号行：(项目名, 内容)，只列出写了的项
        "meta": [(label, field(name)) for name, label in (("articleno", "文章编号"), ("clc", "中图分类号"),
                                                           ("doccode", "文献标志码"), ("subjectcode", "学科分类号"))
                 if field(name)],
        "bio": None,
    }
    mb = re.search(r"\\biography\s*\{", pre)
    if mb:  # \biography{照片}{姓名}{简介}
        args, pos = [], mb.end() - 1
        for _ in range(3):
            e = group_end(pre, pos)
            args.append(pre[pos + 1:e].strip())
            pos = e + 1
            while pos < len(pre) and pre[pos] in " \n":
                pos += 1
        text = args[2].replace("\\_", "_").replace("\\%", "%").replace("\\&", "&").replace("~", " ")
        info["bio"] = (find_graphic(args[0], report), args[1], text)

    warnings: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        body_tex = convert_body(main_body, cites, aux_labels, tmp, report, refs)
        fund = field("fund")
        footnote = f"\\footnote{{\\textbf{{基金项目：}}{fund}}}" if fund else ""
        front = "\n\n".join((["⟦META⟧编号行"] if info["meta"] else []) + [
            f"⟦TITLE⟧{title}",
            f"⟦AUTHOR⟧{field('author')}",
            f"⟦AFFIL⟧（{field('affil')}）",
            f"⟦ENTITLE⟧{field('entitle')}",
            f"⟦ENAUTHOR⟧{field('enauthor')}",
            f"⟦ENAFFIL⟧({field('enaffil')})",
            f"⟦ENABS⟧\\textbf{{ABSTRACT:}} {field('enabstract')}{footnote}",
            f"⟦ENKW⟧\\textbf{{KEY WORDS:}} {field('enkeywords')}",
            f"⟦CNABS⟧\\textbf{{摘要：}}{field('abstract')}",
            f"⟦CNKW⟧\\textbf{{关键词：}}{field('keywords')}",
            "⟦BODY⟧",
        ])
        bio_tex = "\n\n⟦BIO⟧作者简介\n" if info["bio"] else ""
        (tmp / "in.tex").write_text(front + "\n\n" + body_tex + bio_tex, encoding="utf-8")
        ref_docx = tmp / "reference.docx"
        make_reference(ref_docx)
        r = subprocess.run([pandoc_exe(), "in.tex", "-f", "latex", "-t", "docx", "--reference-doc", str(ref_docx),
                            "-o", "body.docx"], cwd=tmp, capture_output=True, text=True, encoding="utf-8")
        if r.returncode != 0:
            raise RuntimeError(f"pandoc 失败：{r.stderr}")
        if r.stderr.strip():
            warnings.append("pandoc：" + r.stderr.strip())
        out = Path(a.output).resolve() if a.output else report / f"{stem}.docx"
        postprocess(tmp / "body.docx", out, info)

    subprocess.run(["latexmk", "-c", a.tex], cwd=report, capture_output=True)
    for ext in (".bbl", ".xdv"):
        (report / f"{stem}{ext}").unlink(missing_ok=True)

    print(f"Word：{out}")
    for w in warnings:
        print("  注意：", w)
    if a.pdf:
        pdf = out.with_suffix(".word.pdf")
        word_to_pdf(out, pdf)
        print(f"Word 另存的 PDF：{pdf}")


if __name__ == "__main__":
    main()
