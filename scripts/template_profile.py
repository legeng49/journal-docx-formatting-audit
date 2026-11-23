# -*- coding: utf-8 -*-
"""Merge a journal Word template into the active profile, then audit against it.

A newly uploaded template updates only the fields it actually states or shows
on a labeled sample. Fields the new template does not define stay as they are.
"""
import copy
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
W = f"{{{W_NS}}}"
A = f"{{{A_NS}}}"
V = f"{{{V_NS}}}"
O = f"{{{O_NS}}}"

SKILL_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DEFAULT_PROFILE = os.path.join(SKILL_ROOT, "references", "csee_profile.json")
DEFAULT_SPEC = os.path.join(SKILL_ROOT, "references", "csee_template_spec.md")

FACTORY_MARGIN = {
    "top": 1440, "bottom": 1440, "left": 1440, "right": 1440, "header": 720, "footer": 720,
}
SIZE_WORDS = [
    ("小五", 9), ("六号", 7.5), ("小四", 12), ("小二", 18), ("四号", 14),
    ("五号", 10.5), ("二号", 22), ("三号", 16),
]
UNIT_ONLY = re.compile(
    r"^(MW|kW|GW|MWh|kWh|Hz|kHz|mHz|kV|V|A|kA|s|ms|pu|rad|mm|cm|km|%)$"
)
GLUED_UNIT = re.compile(
    r"\d(?:MW|kW|GW|MWh|kWh|Hz|kHz|mHz|kV|mm|cm|km|pu)\b"
)


def load_profile(path=None):
    path = path or DEFAULT_PROFILE
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_profile(profile, path=None):
    path = path or DEFAULT_PROFILE
    with open(path, "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)
        f.write("\n")


def leaf_paths(obj, prefix=""):
    out = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{prefix}.{k}" if prefix else k
            if isinstance(v, dict):
                out.extend(leaf_paths(v, p))
            else:
                out.append(p)
    return out


def merge_profile(base, patch):
    """Copy base and replace only leaves present in patch. Never delete a leaf the patch omits."""
    out = copy.deepcopy(base)
    updated = []

    def rec(dst, src, path):
        for k, v in src.items():
            p = f"{path}.{k}" if path else k
            if isinstance(v, dict) and isinstance(dst.get(k), dict):
                rec(dst[k], v, p)
            else:
                dst[k] = copy.deepcopy(v)
                updated.append(p)

    rec(out, patch, "")
    kept = [p for p in leaf_paths(base) if p not in set(updated)]
    return out, updated, kept


def cm_to_twips(value):
    return int(round(float(value) / 2.54 * 1440))


def pt_to_twips(value):
    return int(round(float(value) * 20))


def _local(tag):
    return tag.split("}")[-1] if tag else ""


def _attr(el, name):
    if el is None:
        return None
    return el.get(f"{W}{name}")


def _text_of(p):
    parts = []
    for t in p.iter(f"{W}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts).replace("\u0001", "").strip()


def _style_map(root):
    styles = {}
    defaults = {"pPr": None, "rPr": None}
    if root is None:
        return styles, defaults
    dd = root.find(f"{W}docDefaults")
    if dd is not None:
        defaults["rPr"] = dd.find(f"{W}rPrDefault/{W}rPr")
        defaults["pPr"] = dd.find(f"{W}pPrDefault/{W}pPr")
    for s in root.findall(f"{W}style"):
        sid = s.get(f"{W}styleId")
        if sid:
            styles[sid] = s
    return styles, defaults


def _chain(styles, sid):
    seen = set()
    seq = []
    while sid and sid not in seen and sid in styles:
        seen.add(sid)
        seq.append(styles[sid])
        based = styles[sid].find(f"{W}basedOn")
        sid = _attr(based, "val") if based is not None else None
    seq.reverse()
    return seq


def _overlay_rpr(dst, rpr):
    if rpr is None:
        return
    sz = rpr.find(f"{W}sz")
    if sz is not None and _attr(sz, "val"):
        dst["size_pt"] = int(_attr(sz, "val")) / 2.0
    fonts = rpr.find(f"{W}rFonts")
    if fonts is not None:
        ea = fonts.get(f"{W}eastAsia")
        ascii_name = fonts.get(f"{W}ascii") or fonts.get(f"{W}hAnsi")
        if ea:
            dst["eastAsia"] = ea
        if ascii_name:
            dst["ascii"] = ascii_name
    b = rpr.find(f"{W}b")
    if b is not None:
        dst["bold"] = _attr(b, "val") not in ("0", "false", "off")


def _overlay_ppr(dst, ppr):
    if ppr is None:
        return
    sp = ppr.find(f"{W}spacing")
    if sp is not None:
        rule = _attr(sp, "lineRule") or "auto"
        line = _attr(sp, "line")
        if rule == "exact" and line:
            dst["line_exact_pt"] = int(line) / 20.0
            dst.pop("line", None)
        elif rule == "atLeast":
            dst["line"] = "atLeast"
            if line:
                dst["line_at_least_twips"] = int(line)
            dst.pop("line_exact_pt", None)
        elif rule == "auto":
            dst["line"] = "single" if line in (None, "240") else f"auto:{line}"
            dst.pop("line_exact_pt", None)
        if _attr(sp, "before") is not None:
            dst["before_pt"] = int(_attr(sp, "before")) / 20.0
        if _attr(sp, "after") is not None:
            dst["after_pt"] = int(_attr(sp, "after")) / 20.0
    ind = ppr.find(f"{W}ind")
    if ind is not None:
        if _attr(ind, "firstLine") is not None:
            dst["first_line_twips"] = int(_attr(ind, "firstLine"))
            dst["hanging_twips"] = 0
        if _attr(ind, "hanging") is not None:
            dst["hanging_twips"] = int(_attr(ind, "hanging"))
            dst["first_line_twips"] = 0
    jc = ppr.find(f"{W}jc")
    if jc is not None and _attr(jc, "val"):
        val = _attr(jc, "val")
        dst["align"] = {"both": "both", "center": "center", "right": "right", "left": "left", "start": "left"}.get(val, val)
    snap = ppr.find(f"{W}snapToGrid")
    if snap is not None:
        dst["snap"] = _attr(snap, "val") not in ("0", "false", "off")
    _overlay_rpr(dst, ppr.find(f"{W}rPr"))


def paragraph_fields(p, styles, defaults):
    dst = {}
    if defaults.get("rPr") is not None:
        _overlay_rpr(dst, defaults["rPr"])
    if defaults.get("pPr") is not None:
        _overlay_ppr(dst, defaults["pPr"])
    ppr = p.find(f"{W}pPr")
    sid = None
    if ppr is not None:
        ps = ppr.find(f"{W}pStyle")
        sid = _attr(ps, "val") if ps is not None else None
    for st in _chain(styles, sid):
        _overlay_ppr(dst, st.find(f"{W}pPr"))
        _overlay_rpr(dst, st.find(f"{W}rPr"))
    if ppr is not None:
        _overlay_ppr(dst, ppr)
    for r in p.findall(f"{W}r"):
        _overlay_rpr(dst, r.find(f"{W}rPr"))
    return dst


def _instruction_fields(text):
    """Numbers and fonts named in a template sentence. These override the sample's direct formatting."""
    fields = {}
    if "小五黑" in text:
        fields.update(size_pt=9, eastAsia="黑体", bold=True)
    elif "小四黑" in text:
        fields.update(size_pt=12, eastAsia="黑体", bold=True)
    elif "六宋" in text or "六号" in text:
        fields.update(size_pt=7.5, eastAsia="宋体")
    else:
        for word, pt in SIZE_WORDS:
            if word in text:
                fields["size_pt"] = pt
                break
    if "黑体" in text and "eastAsia" not in fields:
        fields["eastAsia"] = "黑体"
    if ("宋体" in text or "小五宋" in text or "六宋" in text) and "仿宋" not in text and "eastAsia" not in fields:
        fields["eastAsia"] = "宋体"
    if "仿宋" in text:
        fields["eastAsia"] = "仿宋_GB2312"
    if "楷体" in text and "姓名" not in text:
        fields["eastAsia"] = "楷体_GB2312"
    if "Times New Roman" in text:
        fields["ascii"] = "Times New Roman"
    if re.search(r"(?<![A-Za-z])Arial(?![A-Za-z])", text):
        fields["ascii"] = "Arial"
    if "加黑" in text or "加粗" in text:
        fields["bold"] = True
    m_line = re.search(r"行距(?:为)?(?:固定值)?\s*([0-9.]+)\s*(?:pt|磅)", text, re.I)
    if m_line:
        fields["line_exact_pt"] = float(m_line.group(1))
        fields.pop("line", None)
    elif "单倍行距" in text:
        fields["line"] = "single"
        fields["snap"] = True
        fields.pop("line_exact_pt", None)
    if "段前段后均空" in text:
        m = re.search(r"段前段后均空\s*([0-9.]+)\s*pt", text, re.I)
        pt = float(m.group(1)) if m else 6
        fields["before_pt"] = pt
        fields["after_pt"] = pt
    elif "段前段后不空" in text:
        fields["before_pt"] = 0
        fields["after_pt"] = 0
    if re.search(r"段前空\s*([0-9.]+)\s*pt", text, re.I):
        fields["before_pt"] = float(re.search(r"段前空\s*([0-9.]+)\s*pt", text, re.I).group(1))
    if re.search(r"段后空\s*([0-9.]+)\s*pt", text, re.I):
        fields["after_pt"] = float(re.search(r"段后空\s*([0-9.]+)\s*pt", text, re.I).group(1))
    if re.search(r"段首空\s*([0-9.]+)\s*pt", text, re.I):
        fields["before_pt"] = float(re.search(r"段首空\s*([0-9.]+)\s*pt", text, re.I).group(1))
    m3 = re.search(r"段前及英文表题后空\s*([0-9.]+)\s*pt", text, re.I)
    if m3:
        fields["before_pt"] = float(m3.group(1))
        fields["after_en_pt"] = float(m3.group(1))
    if re.search(r"英文图题段后空\s*([0-9.]+)\s*pt", text, re.I):
        fields["after_en_pt"] = float(re.search(r"英文图题段后空\s*([0-9.]+)\s*pt", text, re.I).group(1))
    m_ind = re.search(r"首行缩进\s*([0-9.]+)\s*cm", text)
    if m_ind:
        fields["first_line_twips"] = cm_to_twips(m_ind.group(1))
    m_hang = re.search(r"悬挂缩进\s*([0-9.]+)\s*cm", text)
    if m_hang:
        fields["hanging_twips"] = cm_to_twips(m_hang.group(1))
    if "悬挂缩进为0" in text or "悬挂缩进为 0" in text:
        fields["hanging_twips"] = 0
    if "居中" in text:
        fields["align"] = "center"
    if "居右" in text or "右对齐" in text:
        fields["align"] = "right"
    if "两端对齐" in text:
        fields["align"] = "both"
    if "顶线" in text and "0.75" in text:
        fields["border_top_bottom_eighths"] = 6
    if "栏目线" in text and "0.5" in text:
        fields["border_inside_eighths"] = 4
    photo = re.search(r"(\d+)\s*[×xX]\s*(\d+)\s*像素", text)
    if photo:
        fields["photo_min_px"] = [int(photo.group(1)), int(photo.group(2))]
    mt = re.search(
        r"完全\s*([0-9.]+)\s*pt、上标/下标\s*([0-9.]+)\s*pt、次上标/下标\s*([0-9.]+)\s*pt、符号\s*([0-9.]+)\s*pt、次符号\s*([0-9.]+)\s*pt",
        text,
    )
    if mt:
        fields["mt_full_pt"] = float(mt.group(1))
        fields["mt_sub_pt"] = float(mt.group(2))
        fields["mt_ssub_pt"] = float(mt.group(3))
        fields["mt_symbol_pt"] = float(mt.group(4))
        fields["mt_subsymbol_pt"] = float(mt.group(5))
    return fields


def _role_of_instruction(text):
    t = text.strip()
    rules = [
        ("page", ("页边距", "每页行数", "页面设置")),
        ("header", ("页眉小五",)),
        ("front_meta", ("文章编号", "中图分类号")),
        ("title_zh", ("标题二号",)),
        ("authors", ("作者姓名",)),
        ("affiliation_zh", ("作者单位",)),
        ("title_en", ("英文标题",)),
        ("authors_en", ("英文姓名",)),
        ("affiliation_en", ("英文单位",)),
        ("abstract_en", ("ABSTRACT",)),
        ("keywords_en", ("KEY WORDS", "英文关键词")),
        ("abstract_zh", ("摘要：", "摘要:")),
        ("keywords_zh", ("关键词：", "关键词:")),
        ("heading1", ("二级标题",)),
        ("heading2", ("三级标题",)),
        ("heading3", ("四级标题",)),
        ("body", ("正文字号",)),
        ("equation", ("MathType", "尺寸定义")),
        ("table_cell", ("表中文字", "表内栏目")),
        ("figure_interior", ("图中中文", "图中文字")),
        ("caption_zh", ("中文表题", "图题与表题")),
        ("references", ("参考文献为",)),
        ("appendix", ("附录格式",)),
        ("biography", ("作者简介",)),
        ("editor", ("编辑姓名", "责任编辑")),
        ("acknowledgment_heading", ()),
    ]
    if t.startswith("页眉") and "页边距" not in t:
        return "header"
    if re.match(r"^参考文献\s*$", t):
        return "references_heading"
    if re.match(r"^致\s*谢\s*$", t):
        return "acknowledgment_heading"
    if re.match(r"^附录\s*[A-ZＡ]?\s*$", t) or re.match(r"^附录\s+[A-ZＡ]", t):
        return "appendix_heading"
    if "整篇文档行距" in t or "整篇文档" in t and "五号宋体" in t:
        return "body"
    if "以上部分行距" in t:
        return "front_common"
    for role, keys in rules:
        if any(k in t for k in keys):
            return role
    return None


def _gap_fill(sample, inst, text):
    """Use a sample paragraph only for values the instruction sentence does not already state.

    A long instructional paragraph is not itself the specimen: its first-line indent
    and a trailing Symbol run must not redefine the role.
    """
    instructional = len(text) > 36 and any(w in text for w in ("为", "要求", "格式", "定义", "行距"))
    gap = {}
    safe = ("line", "line_exact_pt", "before_pt", "after_pt", "snap")
    if not instructional:
        safe = safe + ("size_pt", "eastAsia", "ascii", "bold", "align", "first_line_twips", "hanging_twips")
    for key in safe:
        if key not in inst and key in sample:
            gap[key] = sample[key]
    if gap.get("ascii") in ("Symbol", "宋体"):
        gap.pop("ascii", None)
    if "line_exact_pt" in inst or "line_exact_pt" in gap:
        gap.pop("line", None)
    if inst.get("line") == "single" or gap.get("line") == "single":
        gap.pop("line_exact_pt", None)
    return gap


def _apply_role(patch, role, fields):
    if not role or not fields:
        return
    if role == "page":
        patch.setdefault("page", {}).update(fields)
        return
    if role == "header":
        page = patch.setdefault("page", {})
        if "size_pt" in fields:
            page["header_size_pt"] = fields["size_pt"]
        if "eastAsia" in fields:
            page["header_eastAsia"] = fields["eastAsia"]
        if "ascii" in fields:
            page["header_ascii"] = fields["ascii"]
        return
    if role == "figure_interior":
        keep = {k: fields[k] for k in ("size_pt", "eastAsia", "ascii") if k in fields}
        patch.setdefault("roles", {}).setdefault(role, {}).update(keep)
        return
    if role == "equation":
        drop = {"eastAsia", "ascii", "bold", "first_line_twips", "hanging_twips", "size_pt", "before_pt", "after_pt"}
        keep = {k: v for k, v in fields.items() if k not in drop}
        patch.setdefault("roles", {}).setdefault(role, {}).update(keep)
        return
    if role == "caption_zh":
        zh = {k: v for k, v in fields.items() if k not in ("after_en_pt", "after_pt")}
        patch.setdefault("roles", {}).setdefault("caption_zh", {}).update(zh)
        en = {}
        if "line_exact_pt" in fields:
            en["line_exact_pt"] = fields["line_exact_pt"]
        if "ascii" in fields:
            en["ascii"] = fields["ascii"]
        if "bold" in fields:
            en["bold"] = fields["bold"]
        if "align" in fields:
            en["align"] = fields["align"]
        if "size_pt" in fields:
            en["size_pt"] = fields["size_pt"]
        if "after_en_pt" in fields:
            en["after_pt"] = fields["after_en_pt"]
        if en:
            patch.setdefault("roles", {}).setdefault("caption_en", {}).update(en)
        return
    if role == "editor":
        slot = patch.setdefault("roles", {}).setdefault("editor", {})
        if "五号" in json.dumps(fields, ensure_ascii=False) or fields.get("size_pt"):
            slot["size_pt"] = fields.get("size_pt", 10.5)
        slot["label_eastAsia"] = "黑体"
        slot["name_eastAsia"] = "楷体_GB2312"
        if "align" in fields:
            slot["align"] = fields["align"]
        if "before_pt" in fields:
            slot["before_pt"] = fields["before_pt"]
        if "line" in fields:
            slot["line"] = fields["line"]
        return
    # Specimen formatting that merely repeats a Normal default is still explicit
    # on a labeled sample, so it belongs in the patch.
    drop = {"line_at_least_twips"}
    clean = {k: v for k, v in fields.items() if k not in drop and k != "after_en_pt"}
    patch.setdefault("roles", {}).setdefault(role, {}).update(clean)


def _rules_from_text(text):
    """Only the sentences this template actually contains. Unmentioned rules are omitted."""
    rules = {}
    if "段尾不要单字" in text or "单字成行" in text:
        rules["body_no_orphan_line"] = True
    if "每隔3位" in text or "每隔 3 位" in text:
        rules["thousands_quarter_space"] = True
    if "数字和单位之间" in text:
        rules["unit_quarter_space"] = True
    if "不能置于行首" in text:
        rules["body_no_line_start_unit_or_ref"] = True
    if "列项说明" in text:
        rules["list_marker"] = "1）"
    if "Word图片" in text or "锁在文字中" in text:
        rules["figure_inline"] = True
        rules["figure_center"] = True
    if "主线" in text and "副线" in text:
        rules["figure_line_weight"] = True
    if "坐标轴平行" in text:
        rules["figure_axis_label"] = True
    if "录波" in text:
        rules["figure_oscillograph"] = True
    if "图题" in text and "表题" in text:
        rules["caption_bilingual"] = True
    if "分数形式" in text:
        rules["table_unit_fraction"] = True
    if "φB/%" in text or "φB/%" in text:
        rules["table_percent_phi"] = True
    if "制表位对齐" in text:
        rules["table_tab_align"] = True
    if "编号后空2" in text or "编号后空 2" in text:
        rules["references_spaces_after_number"] = 2
    if "括弧为半角" in text or "括号为半角" in text:
        rules["references_fullwidth_punct"] = True
    if "7714" in text:
        rules["references_gbt7714"] = True
    if "缩写的外文期刊" in text:
        rules["references_full_journal_name"] = True
    if "第一次出现" in text and ("缩略" in text or "符号" in text):
        rules["notation_abbreviation_on_first_use"] = True
    if "3100" in text and "3102" in text:
        rules["notation_gb3100"] = True
    return rules


def _header_titles(zin):
    found = []
    blob = []
    for name in zin.namelist():
        if name.startswith("word/header") and name.endswith(".xml"):
            root = ET.fromstring(zin.read(name))
            blob.append("".join(t.text or "" for t in root.iter(f"{W}t")))
    text = "\n".join(blob)
    compact = re.sub(r"\s+", "", text)
    if "中国电机工程学报" in compact and "中国电机工程学报" not in found:
        found.append("中国电机工程学报")
    if re.search(r"Proceedings of the CSEE", text, re.I):
        found.append("Proceedings of the CSEE")
    return found


def _page_from_text(text, page):
    m = re.search(
        r"上\s*([0-9.]+)\s*cm[，,]\s*下\s*([0-9.]+)\s*cm[，,]\s*左\s*([0-9.]+)\s*cm[，,]\s*右\s*([0-9.]+)\s*cm",
        text,
    )
    if m:
        page["top"] = cm_to_twips(m.group(1))
        page["bottom"] = cm_to_twips(m.group(2))
        page["left"] = cm_to_twips(m.group(3))
        page["right"] = cm_to_twips(m.group(4))
    m2 = re.search(r"栏间距\s*([0-9.]+)\s*cm", text)
    if m2:
        page["col_space"] = cm_to_twips(m2.group(1))
    m3 = re.search(r"页眉\s*([0-9.]+)\s*cm", text)
    if m3 and "页眉小五" not in text:
        page["header"] = cm_to_twips(m3.group(1))
    if re.search(r"页脚\s*0(?!\d)", text):
        page["footer"] = 0
    m4 = re.search(r"每页行数\s*(\d+)\s*行[，,]\s*每行字数\s*(\d+)\s*[（(]\s*半栏\s*(\d+)", text)
    if m4:
        page["lines_per_page"] = int(m4.group(1))
        page["chars_per_line"] = int(m4.group(2))
        page["chars_per_half_line"] = int(m4.group(3))
    if "与网格对齐" in text or "根据页面设置确定行高线" in text:
        page["snap_to_grid"] = True
    m5 = re.search(r"第一制表位居中[，,]?\s*([0-9.]+)\s*cm[，,]?\s*第二制表位居右[，,]?\s*([0-9.]+)\s*cm", text)
    if m5:
        page["header_tab_center_cm"] = float(m5.group(1))
        page["header_tab_right_cm"] = float(m5.group(2))


def _page_from_sect(sect, page, emit_margins):
    pg_sz = sect.find(f"{W}pgSz")
    if pg_sz is not None:
        w, h = _attr(pg_sz, "w"), _attr(pg_sz, "h")
        if w and int(w) != 12240:
            page["width_twips"] = int(w)
        if h and int(h) != 15840:
            page["height_twips"] = int(h)
    pg = sect.find(f"{W}pgMar")
    if pg is not None and emit_margins:
        for key in ("top", "bottom", "left", "right", "header", "footer"):
            raw = _attr(pg, key)
            if raw is None:
                continue
            val = int(raw)
            if val != FACTORY_MARGIN[key]:
                page[key] = val
    cols = sect.find(f"{W}cols")
    if cols is not None and _attr(cols, "space") and int(_attr(cols, "space")) not in (0, 720):
        page["col_space"] = int(_attr(cols, "space"))
    if sect.find(f"{W}evenAndOddHeaders") is not None:
        page["odd_even_header"] = True
    if sect.find(f"{W}titlePg") is not None:
        page["different_first_page"] = True
    grid = sect.find(f"{W}docGrid")
    if grid is not None and _attr(grid, "linePitch") and "line_pitch_twips" not in page:
        page["line_pitch_twips"] = int(_attr(grid, "linePitch"))


def extract_patch(docx_path):
    """Return a partial profile. Omitted leaves are not a request to clear them."""
    patch = {"page": {}, "roles": {}}
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc = ET.fromstring(zin.read("word/document.xml"))
        styles_root = ET.fromstring(zin.read("word/styles.xml")) if "word/styles.xml" in zin.namelist() else None
        styles, defaults = _style_map(styles_root)
        body = doc.find(f"{W}body")
        paragraphs = []

        def walk(el):
            for ch in list(el):
                name = _local(ch.tag)
                if name == "p":
                    paragraphs.append(ch)
                else:
                    walk(ch)

        walk(body)
        texts = [_text_of(p) for p in paragraphs]
        full = "\n".join(texts)
        _page_from_text(full, patch["page"])
        rule_patch = _rules_from_text(full)
        if rule_patch:
            patch["rules"] = rule_patch
        header_bits = _header_titles(zin)
        if header_bits:
            patch.setdefault("rules", {})["header_must_contain"] = header_bits
        photo = re.search(r"(\d+)\s*[×xX]\s*(\d+)\s*像素", full)
        if photo:
            patch.setdefault("roles", {}).setdefault("biography", {})["photo_min_px"] = [int(photo.group(1)), int(photo.group(2))]
        instruction_lines = "lines_per_page" in patch["page"]
        instruction_defined_page = any(k in patch["page"] for k in ("top", "bottom", "left", "right", "header", "footer"))
        sects = doc.findall(f".//{W}sectPr")
        for sect in sects:
            _page_from_sect(sect, patch["page"], emit_margins=not instruction_defined_page)
        if instruction_lines:
            usable = patch["page"].get("height_twips", 16838) - patch["page"].get("top", 1814) - patch["page"].get("bottom", 1417)
            lines = patch["page"]["lines_per_page"]
            if lines:
                patch["page"]["line_pitch_twips"] = int(round(usable / lines))
        elif "line_pitch_twips" in patch["page"] and "lines_per_page" not in patch["page"]:
            usable = patch["page"].get("height_twips", 16838) - patch["page"].get("top", 1814) - patch["page"].get("bottom", 1417)
            pitch = patch["page"]["line_pitch_twips"]
            if pitch:
                patch["page"]["lines_per_page"] = int(round(usable / pitch))
        cols_counts = []
        for sect in sects:
            cols = sect.find(f"{W}cols")
            num = _attr(cols, "num") if cols is not None else None
            cols_counts.append(int(num) if num else 1)
        if cols_counts:
            patch["page"]["section1_cols"] = cols_counts[0]
        if len(cols_counts) >= 2:
            patch["page"]["section2_cols"] = cols_counts[1]
        if "word/settings.xml" in zin.namelist():
            settings = ET.fromstring(zin.read("word/settings.xml"))
            hy = settings.find(f"{W}autoHyphenation")
            if hy is not None:
                patch["page"]["auto_hyphenation"] = _attr(hy, "val") not in ("0", "false", "off")

        pending_ack = False
        for p, text in zip(paragraphs, texts):
            if not text:
                continue
            role = _role_of_instruction(text)
            fields = {}
            if role:
                sample = paragraph_fields(p, styles, defaults)
                inst = _instruction_fields(text)
                fields.update(_gap_fill(sample, inst, text))
                fields.update(inst)
                if role in ("title_en", "authors_en", "affiliation_en", "abstract_en", "keywords_en", "caption_en"):
                    if "eastAsia" not in inst:
                        fields.pop("eastAsia", None)
                    if "bold" not in inst:
                        fields.pop("bold", None)
                if "hanging_twips" in fields:
                    fields.pop("first_line_twips", None)
                elif "first_line_twips" in inst:
                    fields.pop("hanging_twips", None)
                if role == "page":
                    continue
                if role == "front_common":
                    fields = {"line": "single", "snap": True}
                if role == "body" and "整篇文档" in text:
                    fields = {k: fields[k] for k in fields if k in ("line", "snap", "size_pt", "eastAsia", "ascii")}
                _apply_role(patch, role, fields)
            if role == "acknowledgment_heading":
                pending_ack = True
                continue
            if pending_ack:
                pending_ack = False
                sample = paragraph_fields(p, styles, defaults)
                sample.update(_instruction_fields(text))
                keep = {k: sample[k] for k in ("size_pt", "eastAsia", "ascii", "align", "line", "first_line_twips", "before_pt", "after_pt") if k in sample}
                _apply_role(patch, "acknowledgment", keep)
    if not patch["page"]:
        patch.pop("page")
    if not patch["roles"]:
        patch.pop("roles")
    return patch


def convert_doc_to_docx(path):
    out = os.path.join(tempfile.gettempdir(), "journal_template_import.docx")
    if os.path.exists(out):
        os.remove(out)
    ps = r'''
$lockDir = Join-Path $env:USERPROFILE ".gemini\tmp"
if (-not (Test-Path $lockDir)) { New-Item -ItemType Directory -Path $lockDir -Force | Out-Null }
$lockFile = Join-Path $lockDir "word_com.lock"
$deadline = (Get-Date).AddSeconds(90)
$lockStream = $null
while ((Get-Date) -lt $deadline) {
  try { $lockStream = [System.IO.File]::Open($lockFile, 'OpenOrCreate', 'ReadWrite', 'None'); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $lockStream) { Write-Error "lock timeout"; exit 1 }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($env:SRC_DOC, $false, $true)
  $fmt = 12
  $doc.SaveAs([ref]$env:DST_DOCX, [ref]$fmt)
  $doc.Close($false)
} finally {
  $word.Quit()
  $lockStream.Close()
  Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
}
'''
    env = os.environ.copy()
    env["SRC_DOC"] = os.path.abspath(path)
    env["DST_DOCX"] = out
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=True, env=env,
    )
    return out


def import_template(template_path, profile_path=None, spec_path=None):
    profile_path = profile_path or DEFAULT_PROFILE
    spec_path = spec_path or DEFAULT_SPEC
    base = load_profile(profile_path)
    temp = None
    src = template_path
    if template_path.lower().endswith(".doc"):
        temp = convert_doc_to_docx(template_path)
        src = temp
    try:
        patch = extract_patch(src)
    finally:
        if temp and os.path.exists(temp):
            os.remove(temp)
    merged, updated, kept = merge_profile(base, patch)
    merged.setdefault("meta", {})
    merged["meta"]["last_template"] = os.path.abspath(template_path)
    merged["meta"]["updated_keys"] = updated
    save_profile(merged, profile_path)
    spec = render_spec(merged)
    os.makedirs(os.path.dirname(os.path.abspath(spec_path)), exist_ok=True)
    with open(spec_path, "w", encoding="utf-8") as f:
        f.write(spec)
    report = {
        "template": os.path.abspath(template_path),
        "profile": os.path.abspath(profile_path),
        "spec": os.path.abspath(spec_path),
        "updated": updated,
        "kept_count": len(kept),
        "kept": kept,
        "rule": "模板里能确定的字段已覆盖；模板没有给出的字段保持导入前的值。",
    }
    return report


def render_spec(profile):
    page = profile.get("page", {})
    roles = profile.get("roles", {})
    policy = profile.get("layout_policy", {})
    lines = [
        "# 当前学报模板版式档案",
        "",
        "数值以 `references/csee_profile.json` 为准。导入新模板时，只覆盖该模板实际写明或在对应样本上设置的字段；模板没有的字段保持原值，不回退到 Word 默认值，也不从档案中删除。",
        "",
        "栏底不齐、留白或末页溢出时，用改写表述改变行数。不得把行距压到档案规定值以下，不得缩小图，不得把图内或表内字体降到档案中的图表字号以下。",
        "",
        "`audit` 的 `stop.allowed` 为 true 之前，任务不得结束。档案里的每一条闸门都要满足，包括这次模板没有改写、因此仍然有效的旧规定。",
        "",
        "## 页面",
        "",
    ]
    def tw(key, label):
        if key in page:
            cm = round(page[key] / 1440 * 2.54, 2)
            lines.append(f"- {label}：`{page[key]}` twips（{cm} cm）")
    tw("width_twips", "纸宽")
    tw("height_twips", "纸高")
    tw("top", "上边距")
    tw("bottom", "下边距")
    tw("left", "左边距")
    tw("right", "右边距")
    tw("header", "页眉距边界")
    tw("footer", "页脚距边界")
    tw("col_space", "栏距")
    if "lines_per_page" in page:
        lines.append(
            f"- 文档网格：每页 {page.get('lines_per_page')} 行，每行 {page.get('chars_per_line')} 字，半栏 {page.get('chars_per_half_line')} 字，行距基准 `{page.get('line_pitch_twips')}` twips，与网格对齐：{page.get('snap_to_grid')}"
        )
    lines.append(
        f"- 第 1 节 {page.get('section1_cols', 1)} 栏，第 2 节 {page.get('section2_cols', 2)} 栏。模板若没有第 2 节，不改已有的第 2 节栏数。"
    )
    lines.append(
        f"- 页眉：{page.get('header_size_pt')} pt {page.get('header_eastAsia')}+{page.get('header_ascii')}，居中制表位 {page.get('header_tab_center_cm')} cm，居右制表位 {page.get('header_tab_right_cm')} cm。奇偶页眉 {page.get('odd_even_header')}，首页不同 {page.get('different_first_page')}。"
    )
    lines += ["", "## 角色", ""]
    for name, spec in roles.items():
        bits = [f"{k}={v}" for k, v in spec.items()]
        lines.append(f"- `{name}`：" + "；".join(bits))
    rules = profile.get("rules", {})
    if rules:
        lines += ["", "## 正文、图表、参考文献与页眉", ""]
        lines.append("- 段末不单字成行：`%s`。参考文献序号和数字后的单位不放在行首：`%s`。" % (
            rules.get("body_no_orphan_line"), rules.get("body_no_line_start_unit_or_ref")))
        lines.append("- 数字自小数点起每 3 位留 1/4 字空：`%s`。数字与单位之间留 1/4 字空：`%s`。列项标记：`%s`。" % (
            rules.get("thousands_quarter_space"), rules.get("unit_quarter_space"), rules.get("list_marker")))
        lines.append("- 图用嵌入型并居中：`%s` / `%s`。中英文题成对：`%s`。主副线、坐标标目、录波格值须目视核对：`%s` / `%s` / `%s`。" % (
            rules.get("figure_inline"), rules.get("figure_center"), rules.get("caption_bilingual"),
            rules.get("figure_line_weight"), rules.get("figure_axis_label"), rules.get("figure_oscillograph")))
        lines.append("- 表头单位用分数，百分数写成 `φB/%%`，列用制表位而不是空格对齐：`%s` / `%s` / `%s`。" % (
            rules.get("table_unit_fraction"), rules.get("table_percent_phi"), rules.get("table_tab_align")))
        lines.append("- 参考文献编号后空 %s 小格；括号半角、其余全角：`%s`；GB/T 7714 类型标志：`%s`；外文刊名不缩写：`%s`。" % (
            rules.get("references_spaces_after_number"), rules.get("references_fullwidth_punct"),
            rules.get("references_gbt7714"), rules.get("references_full_journal_name")))
        lines.append("- 符号按 GB 3100～3102 且全文一致，缩写首次出现时给出中文全称(英文全称，缩写)：`%s` / `%s`。" % (
            rules.get("notation_gb3100"), rules.get("notation_abbreviation_on_first_use")))
        if rules.get("header_must_contain"):
            lines.append("- 页眉须包含：" + "、".join(rules["header_must_contain"]) + "。并核对页眉字号与居中、居右制表位。")
    lines += [
        "",
        "## 图表字号",
        "",
        f"- 图题、表题见 `caption_zh` / `caption_en`。图内文字见 `figure_interior`，当前为 {roles.get('figure_interior', {}).get('size_pt')} pt。表内文字见 `table_cell`，当前为 {roles.get('table_cell', {}).get('size_pt')} pt，行距 {roles.get('table_cell', {}).get('line_exact_pt')} pt。",
        "- 图内文字若只存在于位图中，OOXML 读不到字号。必须逐图核对为档案中的图内字号后，审计才可以带 `--figure-fonts-verified`。该开关不跳过任何其他闸门，也不能用来接受更小的字号。",
        "",
        "## 版面与停止",
        "",
        f"- 调整手段：`{policy.get('optimize_by')}`。禁止压行距：{policy.get('do_not_compress_line_spacing')}。禁止缩图：{policy.get('do_not_shrink_figures')}。禁止缩小图表字体：{policy.get('do_not_reduce_chart_font')}。",
        f"- 双栏：左右栏顶、底坐标差绝对值不超过 {policy.get('column_abs_d_pt')} pt；两栏都有文字的页，栏底不低于 {policy.get('column_bottom_min_pt')} pt。",
        "- 编号公式的行距、段前段后和 MathType 尺寸以 `roles.equation` 为准。禁止用更大的固定行距把公式行撑开，也禁止为了省行把行距降到该值以下。",
        "",
    ]
    return "\n".join(lines)


def _near(actual, expected, tol):
    if actual is None:
        return False
    return abs(float(actual) - float(expected)) <= tol


def _font_ok(actual, expected):
    if not expected:
        return True
    if not actual:
        return False
    if expected.startswith("楷体"):
        return actual.startswith("楷体")
    if expected.startswith("仿宋"):
        return actual.startswith("仿宋")
    return actual == expected


def _run_props_list(p, styles, defaults):
    base = {}
    if defaults.get("rPr") is not None:
        _overlay_rpr(base, defaults["rPr"])
    ppr = p.find(f"{W}pPr")
    sid = None
    if ppr is not None:
        ps = ppr.find(f"{W}pStyle")
        sid = _attr(ps, "val") if ps is not None else None
        _overlay_rpr(base, ppr.find(f"{W}rPr"))
    for st in _chain(styles, sid):
        _overlay_rpr(base, st.find(f"{W}rPr"))
    runs = []
    for r in p.findall(f"{W}r"):
        if r.find(f".//{O}OLEObject") is not None:
            continue
        text = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not text:
            continue
        cur = dict(base)
        _overlay_rpr(cur, r.find(f"{W}rPr"))
        runs.append((text, cur))
    return runs


def _check_role(role, spec, props, runs, where, violations, gate):
    if not spec:
        return
    if "size_pt" in spec and not _near(props.get("size_pt"), spec["size_pt"], 0.01):
        violations.append({"gate": gate, "where": where, "role": role, "item": "size_pt", "expected": spec["size_pt"], "actual": props.get("size_pt")})
    if "eastAsia" in spec and not _font_ok(props.get("eastAsia"), spec["eastAsia"]):
        violations.append({"gate": gate, "where": where, "role": role, "item": "eastAsia", "expected": spec["eastAsia"], "actual": props.get("eastAsia")})
    if "ascii" in spec and not _font_ok(props.get("ascii"), spec["ascii"]):
        violations.append({"gate": gate, "where": where, "role": role, "item": "ascii", "expected": spec["ascii"], "actual": props.get("ascii")})
    if spec.get("bold") is True and props.get("bold") is not True:
        violations.append({"gate": gate, "where": where, "role": role, "item": "bold", "expected": True, "actual": props.get("bold")})
    if "align" in spec and props.get("align") not in (None, spec["align"]):
        violations.append({"gate": gate, "where": where, "role": role, "item": "align", "expected": spec["align"], "actual": props.get("align")})
    if spec.get("line") == "single":
        if props.get("line") not in (None, "single") or "line_exact_pt" in props:
            violations.append({"gate": gate, "where": where, "role": role, "item": "line", "expected": "single", "actual": props.get("line") or props.get("line_exact_pt")})
    if "line_exact_pt" in spec and not _near(props.get("line_exact_pt"), spec["line_exact_pt"], 0.2):
        violations.append({"gate": gate, "where": where, "role": role, "item": "line_exact_pt", "expected": spec["line_exact_pt"], "actual": props.get("line_exact_pt")})
    for key, tol in (("before_pt", 0.6), ("after_pt", 0.6)):
        if key in spec and not _near(props.get(key, 0), spec[key], tol):
            violations.append({"gate": gate, "where": where, "role": role, "item": key, "expected": spec[key], "actual": props.get(key, 0)})
    for key, tol in (("before_twips", 20), ("after_twips", 20), ("first_line_twips", 30), ("hanging_twips", 30)):
        if key in spec:
            actual = props.get(key, 0)
            if key.startswith("before"):
                actual = props.get("before_pt", 0) * 20
            elif key.startswith("after"):
                actual = props.get("after_pt", 0) * 20
            if not _near(actual, spec[key], tol):
                violations.append({"gate": gate, "where": where, "role": role, "item": key, "expected": spec[key], "actual": actual})
    if spec.get("snap") is True and props.get("snap") is False:
        violations.append({"gate": gate, "where": where, "role": role, "item": "snap", "expected": True, "actual": False})
    if spec.get("keep_next") is True and not props.get("keep_next"):
        violations.append({"gate": gate, "where": where, "role": role, "item": "keep_next", "expected": True, "actual": False})
    if "label_eastAsia" in spec or spec.get("label_bold"):
        label = ""
        if role == "abstract_zh":
            label = "摘要"
        elif role == "keywords_zh":
            label = "关键词"
        hit = False
        for text, rp in runs:
            if text.startswith(label):
                hit = True
                if "label_eastAsia" in spec and not _font_ok(rp.get("eastAsia"), spec["label_eastAsia"]):
                    violations.append({"gate": gate, "where": where, "role": role, "item": "label_eastAsia", "expected": spec["label_eastAsia"], "actual": rp.get("eastAsia")})
                if spec.get("label_bold") and rp.get("bold") is not True:
                    violations.append({"gate": gate, "where": where, "role": role, "item": "label_bold", "expected": True, "actual": rp.get("bold")})
        if label and not hit:
            violations.append({"gate": gate, "where": where, "role": role, "item": "label", "expected": label, "actual": None})


def _props_from_fields(fields, p):
    props = dict(fields)
    ppr = p.find(f"{W}pPr")
    props["keep_next"] = ppr is not None and ppr.find(f"{W}keepNext") is not None
    if "before_pt" not in props:
        props["before_pt"] = 0
    if "after_pt" not in props:
        props["after_pt"] = 0
    if props.get("line_exact_pt"):
        props.pop("line", None)
    elif "line" not in props:
        props["line"] = "single"
    return props


def _classify(text, region, in_table, instr):
    t = text.strip()
    if in_table:
        return "table_cell"
    if "责任编辑" in t[:20]:
        return "editor"
    if t.startswith("作者简介"):
        return "biography"
    if re.match(r"^参考文献\s*$", t):
        return "references_heading"
    if re.match(r"^致\s*谢\s*$", t):
        return "acknowledgment_heading"
    if re.match(r"^附录", t) and len(t) <= 12:
        return "appendix_heading"
    if "SEQ" in instr and "公式" in instr:
        return "equation"
    if re.match(r"^\s*图\s*[A-ZＡ]?\d+", t) or re.match(r"^\s*表\s*\d+", t):
        return "caption_zh"
    if re.match(r"^\s*Fig\.", t, re.I) or re.match(r"^\s*Tab\.", t, re.I):
        return "caption_en"
    if re.match(r"^ABSTRACT\b", t, re.I):
        return "abstract_en"
    if re.match(r"^KEY\s*WORDS\b", t, re.I):
        return "keywords_en"
    if re.match(r"^摘要\s*[:：]", t):
        return "abstract_zh"
    if re.match(r"^关键词\s*[:：]", t):
        return "keywords_zh"
    if _LIST_OK.match(t):
        return "list_item"
    if re.match(r"^\d+\.\d+\.\d+\s+\S", t):
        return "heading3"
    if re.match(r"^\d+\.\d+\s+\S", t):
        return "heading2"
    if re.match(r"^\d+\s+\S", t) and not re.match(r"^\d+[）)]", t):
        return "heading1"
    if region == "ref":
        return "references"
    if region == "app":
        return "appendix"
    if region == "ack":
        return "acknowledgment"
    if region == "bio":
        return "biography"
    if region == "front":
        return "front"
    return "body"


def _drawing_font_issues(p, expected_pt):
    issues = []
    rasters = 0
    for blip in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        if p.find(f".//{O}OLEObject") is None:
            rasters += 1
    for imagedata in p.iter(f"{V}imagedata"):
        if p.find(f".//{O}OLEObject") is None:
            rasters += 1
    for rpr in p.iter(f"{A}rPr"):
        sz = rpr.get("sz")
        if sz and abs(int(sz) / 100.0 - expected_pt) > 0.05:
            issues.append({"item": "drawing_sz", "expected": expected_pt, "actual": int(sz) / 100.0})
    for rpr in p.iter(f"{W}rPr"):
        parent_text = "".join(t.text or "" for t in rpr.iter(f"{W}t"))
        if not parent_text:
            continue
    return issues, rasters


_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_THOUSANDS = re.compile(r"(?<!\d)\d{5,}(?:\.\d+)?(?!\d)")
_DECIMAL_RUN = re.compile(r"\.\d{4,}(?!\d)")
_LIST_OK = re.compile(r"^\d+）")
_LIST_BAD = re.compile(r"^(?:[（(]\d+[）)]|\d+[、.](?!\d)\s*)")
_REF_TYPE = re.compile(r"\[(?:J|M|D|R|C|S|N|EB/OL|EB|OL|P|G|A|Z)\]")
_JOURNAL_ABBR = re.compile(r"\b(?:Trans|Proc|J)\.")
_ABBR = re.compile(r"\b[A-Z]{3,8}\b")
_ABBR_ALLOW = {"IEEE", "IEC", "DOI", "PDF", "URL", "PID", "AGC"}
_UNIT_HEAD = re.compile(r"^(MW|kW|GW|Hz|kHz|mHz|kV|mm|cm|km|pu|%)$")


def _norm_fig(token):
    table = str.maketrans("０１２３４５６７８９ＡＢＣＤＥＦＧ", "0123456789ABCDEFG")
    return token.translate(table).upper()


def _check_manuscript_rules(paragraphs, rules, styles, defaults, violations, seen_abbr):
    zh_fig, en_fig, zh_tab, en_tab = set(), set(), set(), set()
    for index, (p, in_table) in enumerate(paragraphs):
        text = _text_of(p)
        if not text:
            continue
        stripped = text.strip()
        if rules.get("thousands_quarter_space") and (_THOUSANDS.search(text) or _DECIMAL_RUN.search(text)):
            violations.append({"gate": "body", "where": index, "item": "thousands", "expected": "自小数点起每 3 位留 1/4 字空", "actual": stripped[:40]})
        if rules.get("list_marker") and _LIST_BAD.match(stripped) and not _LIST_OK.match(stripped):
            violations.append({"gate": "body", "where": index, "item": "list_marker", "expected": rules["list_marker"], "actual": stripped[:20]})
        if rules.get("caption_bilingual"):
            if re.match(r"^图\s*[A-ZＡ]?\d+", stripped):
                m = re.match(r"^图\s*([A-ZＡ]?\d+)", stripped)
                zh_fig.add(_norm_fig(m.group(1)))
            if re.match(r"^表\s*\d+", stripped):
                zh_tab.add(re.match(r"^表\s*(\d+)", stripped).group(1))
            if re.match(r"^Fig\.", stripped, re.I):
                m = re.match(r"^Fig\.\s*([A-Z]?\d+)", stripped, re.I)
                if m:
                    en_fig.add(_norm_fig(m.group(1)))
            if re.match(r"^Tab\.", stripped, re.I):
                m = re.match(r"^Tab\.\s*(\d+)", stripped, re.I)
                if m:
                    en_tab.add(m.group(1))
        if rules.get("figure_inline") or rules.get("figure_center"):
            has_pic = p.find(f".//{_WP}inline") is not None or p.find(f".//{_WP}anchor") is not None or (
                p.find(f".//{V}imagedata") is not None and p.find(f".//{O}OLEObject") is None
            )
            if has_pic and p.find(f".//{O}OLEObject") is None:
                if rules.get("figure_inline") and p.find(f".//{_WP}anchor") is not None:
                    violations.append({"gate": "figures", "where": index, "item": "inline", "expected": "嵌入型", "actual": "浮动"})
                if rules.get("figure_center"):
                    props = _props_from_fields(paragraph_fields(p, styles, defaults), p)
                    if props.get("align") not in ("center",):
                        violations.append({"gate": "figures", "where": index, "item": "center", "expected": "center", "actual": props.get("align")})
        if rules.get("table_tab_align") and in_table and re.search(r" {2,}", text):
            violations.append({"gate": "table_font", "where": index, "item": "tabs", "expected": "制表位对齐", "actual": stripped[:30]})
        if in_table and rules.get("table_percent_phi") and re.fullmatch(r"[A-Za-zφΦ]{0,6}\s*[%％]", stripped.replace(" ", "")):
            violations.append({"gate": "table_font", "where": index, "item": "percent", "expected": "φB/%", "actual": stripped})
        if in_table and rules.get("table_unit_fraction") and _UNIT_HEAD.match(stripped):
            violations.append({"gate": "table_font", "where": index, "item": "unit_fraction", "expected": "单位写成分子在上的分数或与量同格", "actual": stripped})
    if rules.get("caption_bilingual"):
        for num in sorted(zh_fig - en_fig):
            violations.append({"gate": "captions", "item": "fig_en", "expected": f"Fig. {num}", "actual": None})
        for num in sorted(en_fig - zh_fig):
            violations.append({"gate": "captions", "item": "fig_zh", "expected": f"图{num}", "actual": None})
        for num in sorted(zh_tab - en_tab):
            violations.append({"gate": "captions", "item": "tab_en", "expected": f"Tab. {num}", "actual": None})
        for num in sorted(en_tab - zh_tab):
            violations.append({"gate": "captions", "item": "tab_zh", "expected": f"表{num}", "actual": None})
    if rules.get("notation_abbreviation_on_first_use"):
        for index, token in seen_abbr:
            violations.append({"gate": "notation", "where": index, "item": "abbreviation", "expected": "首次出现写成中文全称(英文全称，缩写)", "actual": token})


def _reference_entry_ok(text, rules, index, violations):
    if not rules:
        return
    if rules.get("references_gbt7714") and not _REF_TYPE.search(text):
        violations.append({"gate": "references", "where": index, "item": "gbt7714", "expected": "[J]/[M]/[D]/[R]/[C]/[S]/[N]/[EB/OL]", "actual": text[:40]})
    if rules.get("references_spaces_after_number"):
        if re.match(r"^(\[\d+\]|\d+)", text) and not re.match(r"^(\[\d+\]|\d+)( {2,}|\t)", text):
            violations.append({"gate": "references", "where": index, "item": "spaces_after_number", "expected": rules["references_spaces_after_number"], "actual": text[:20]})
    if rules.get("references_fullwidth_punct") and re.search(r"[\u4e00-\u9fff][,:;]", text):
        violations.append({"gate": "references", "where": index, "item": "fullwidth", "expected": "除括号外用全角标点", "actual": text[:40]})
    if rules.get("references_full_journal_name") and _JOURNAL_ABBR.search(text):
        violations.append({"gate": "references", "where": index, "item": "journal_name", "expected": "外文刊名用全称", "actual": text[:40]})


def _check_header_details(zin, page, rules, violations):
    texts = []
    saw_center = saw_right = False
    saw_font = False
    center = int(round(float(page.get("header_tab_center_cm", 0)) / 2.54 * 1440)) if page.get("header_tab_center_cm") else None
    right = int(round(float(page.get("header_tab_right_cm", 0)) / 2.54 * 1440)) if page.get("header_tab_right_cm") else None
    for name in zin.namelist():
        if not (name.startswith("word/header") and name.endswith(".xml")):
            continue
        root = ET.fromstring(zin.read(name))
        texts.append("".join(t.text or "" for t in root.iter(f"{W}t")))
        for tab in root.iter(f"{W}tab"):
            pos = _attr(tab, "pos")
            kind = _attr(tab, "val")
            if not pos:
                continue
            pos = int(pos)
            if center and kind == "center" and abs(pos - center) <= 40:
                saw_center = True
            if right and kind == "right" and abs(pos - right) <= 40:
                saw_right = True
        for r in root.iter(f"{W}r"):
            rpr = r.find(f"{W}rPr")
            if rpr is None:
                continue
            sz = rpr.find(f"{W}sz")
            fonts = rpr.find(f"{W}rFonts")
            if page.get("header_size_pt") and sz is not None and abs(int(_attr(sz, "val")) / 2 - float(page["header_size_pt"])) < 0.2:
                ea = fonts.get(f"{W}eastAsia") if fonts is not None else ""
                ascii_name = fonts.get(f"{W}ascii") if fonts is not None else ""
                if (not page.get("header_eastAsia") or _font_ok(ea, page.get("header_eastAsia"))) and (
                    not page.get("header_ascii") or ascii_name in (None, "", page.get("header_ascii"))
                ):
                    saw_font = True
    blob = re.sub(r"\s+", "", "\n".join(texts))
    for needle in rules.get("header_must_contain") or []:
        if re.sub(r"\s+", "", needle) not in blob and needle not in "\n".join(texts):
            violations.append({"gate": "headers", "item": "header_text", "expected": needle, "actual": blob[:40]})
    if center and not saw_center:
        violations.append({"gate": "headers", "item": "header_tab_center_cm", "expected": page.get("header_tab_center_cm"), "actual": None})
    if right and not saw_right:
        violations.append({"gate": "headers", "item": "header_tab_right_cm", "expected": page.get("header_tab_right_cm"), "actual": None})
    if page.get("header_size_pt") and texts and not saw_font:
        violations.append({"gate": "headers", "item": "header_font", "expected": page.get("header_size_pt"), "actual": None})


def _check_char_count(doc, page, violations):
    if not page.get("chars_per_line"):
        return
    em = 210
    for idx, sect in enumerate(doc.findall(f".//{W}sectPr")):
        pg = sect.find(f"{W}pgMar")
        psz = sect.find(f"{W}pgSz")
        if pg is None or psz is None:
            continue
        width = int(_attr(psz, "w") or 0) - int(_attr(pg, "left") or 0) - int(_attr(pg, "right") or 0)
        cols = sect.find(f"{W}cols")
        num = int(_attr(cols, "num")) if cols is not None and _attr(cols, "num") else 1
        if num >= 2:
            width = (width - int(_attr(cols, "space") or 0)) / num
            expect = page.get("chars_per_half_line")
        else:
            expect = page.get("chars_per_line")
        if not expect:
            continue
        got = int(round(width / em))
        if abs(got - int(expect)) > 1:
            violations.append({"gate": "doc_grid", "where": idx + 1, "item": "chars_per_line", "expected": expect, "actual": got})


def _check_line_breaks(columns_path, rules, violations):
    if not (rules.get("body_no_orphan_line") or rules.get("body_no_line_start_unit_or_ref")):
        return
    if not columns_path or not os.path.exists(columns_path):
        violations.append({"gate": "line_break", "item": "measurement", "expected": "col_metrics.json 中的 line_breaks", "actual": None})
        return
    with open(columns_path, encoding="utf-8-sig") as f:
        data = json.load(f)
    breaks = data.get("line_breaks")
    if not breaks:
        violations.append({"gate": "line_break", "item": "measurement", "expected": "重新测量以得到 line_breaks", "actual": None})
        return
    if rules.get("body_no_orphan_line"):
        for item in (breaks.get("orphans") or [])[:15]:
            violations.append({"gate": "line_break", "where": item.get("page"), "item": "orphan", "expected": "段末不单字成行", "actual": item.get("text")})
    if rules.get("body_no_line_start_unit_or_ref"):
        for item in (breaks.get("line_starts") or [])[:15]:
            violations.append({"gate": "line_break", "where": item.get("page"), "item": "line_start", "expected": "序号和单位不放在行首", "actual": item.get("text")})


def _notation_pending(rules):
    labels = []
    if rules.get("notation_gb3100"):
        labels.append("变量和单位符合 GB 3100～3102，且大小写、上下标全文一致")
    if rules.get("figure_line_weight"):
        labels.append("图中主线与副线粗细可以区分")
    if rules.get("figure_axis_label"):
        labels.append("坐标标目与坐标轴平行，标值不过密，单字母标目可放在轴端")
    if rules.get("figure_oscillograph"):
        labels.append("录波图标出每格数值和单位")
    return labels


def evaluate(zin, profile, columns_path=None, figure_fonts_verified=False, mathtype_sizes_verified=False, notation_verified=False):
    doc = ET.fromstring(zin.read("word/document.xml"))
    styles_root = ET.fromstring(zin.read("word/styles.xml")) if "word/styles.xml" in zin.namelist() else None
    styles, defaults = _style_map(styles_root)
    page = profile.get("page", {})
    roles = profile.get("roles", {})
    rules = profile.get("rules", {})
    policy = profile.get("layout_policy", {})
    violations = []
    seen = set()
    body = doc.find(f"{W}body")
    paragraphs = []

    def walk(el, in_table=False):
        for ch in list(el):
            name = _local(ch.tag)
            if name == "tbl":
                yield from walk(ch, True)
            elif name == "p":
                yield ch, in_table
            else:
                yield from walk(ch, in_table)

    paragraphs = list(walk(body))
    region = "front"
    front_slots = []
    raster_figures = 0
    equation_count = 0
    abbr_seen = set()
    abbr_bad = []
    for index, (p, in_table) in enumerate(paragraphs):
        text = _text_of(p)
        instr = " ".join(i.text or "" for i in p.findall(f".//{W}instrText"))
        if in_table and not text:
            role = None
        else:
            role = _classify(text, region, in_table, instr) if text or "SEQ" in instr else None
        if role == "abstract_en" or role == "abstract_zh":
            region = "body"
        elif role == "acknowledgment_heading":
            region = "ack"
        elif role == "references_heading":
            region = "ref"
        elif role == "appendix_heading":
            region = "app"
        elif role == "biography":
            region = "bio"
        elif role in ("heading1", "heading2", "heading3") and region in ("ack", "app"):
            region = "body"
        if role:
            seen.add(role)
        if role == "front" and text:
            front_slots.append((index, p, text))
        if role == "equation":
            equation_count += 1
        fig_issues, rasters = _drawing_font_issues(p, roles.get("figure_interior", {}).get("size_pt", 7.5))
        raster_figures += rasters
        for issue in fig_issues:
            issue.update(gate="figure_interior", where=index)
            violations.append(issue)
        if role is None:
            continue
        fields = paragraph_fields(p, styles, defaults)
        props = _props_from_fields(fields, p)
        runs = _run_props_list(p, styles, defaults)
        gate = {
            "table_cell": "table_font",
            "caption_zh": "captions",
            "caption_en": "captions",
            "equation": "equation",
            "references": "references",
            "references_heading": "references",
            "biography": "biography",
            "body": "body",
            "heading1": "headings",
            "heading2": "headings",
            "heading3": "headings",
        }.get(role, "roles")
        if role == "body" and rules.get("unit_quarter_space", True) and GLUED_UNIT.search(text):
            violations.append({"gate": "body", "where": index, "role": "body", "item": "quarter_space", "expected": "数字与单位之间留 1/4 字空", "actual": text[:40]})
        spec = roles.get("body") if role == "list_item" else roles.get(role)
        if role == "list_item":
            gate = "body"
        if spec and role not in ("front",):
            _check_role(role, spec, props, runs, index, violations, gate)
        if role == "references":
            _reference_entry_ok(text, rules, index, violations)
        if role == "editor":
            label_font = roles.get("editor", {}).get("label_eastAsia")
            name_font = roles.get("editor", {}).get("name_eastAsia")
            if label_font and not any("编辑" in t and _font_ok(rp.get("eastAsia"), label_font) for t, rp in runs):
                violations.append({"gate": "roles", "where": index, "role": "editor", "item": "label_eastAsia", "expected": label_font, "actual": None})
            if name_font and not any("编辑" not in t and t.strip("()（） ") and _font_ok(rp.get("eastAsia"), name_font) for t, rp in runs):
                violations.append({"gate": "roles", "where": index, "role": "editor", "item": "name_eastAsia", "expected": name_font, "actual": None})
        if role in ("body", "list_item") and rules.get("notation_abbreviation_on_first_use"):
            for match in _ABBR.finditer(text):
                token = match.group(0)
                if token in _ABBR_ALLOW or token in abbr_seen:
                    continue
                abbr_seen.add(token)
                if not re.search(r"[（(][^）)]{0,80}%s" % re.escape(token), text):
                    abbr_bad.append((index, token))
        if role == "table_cell" and UNIT_ONLY.match(text.replace(" ", "")):
            violations.append({"gate": "table_font", "where": index, "role": "table_cell", "item": "unit_column", "expected": "单位与量写在同一格或表头", "actual": text})

    seen |= _assign_front(front_slots, roles, styles, defaults, violations)

    _check_page(doc, zin, page, violations)
    _check_char_count(doc, page, violations)
    _check_header_details(zin, page, rules, violations)
    _check_tables(doc, roles.get("table_cell", {}), violations)
    _check_manuscript_rules(paragraphs, rules, styles, defaults, violations, abbr_bad)
    _check_line_breaks(columns_path, rules, violations)
    if not notation_verified:
        notes = _notation_pending(rules)
        if notes:
            violations.append({"gate": "notation", "item": "visual", "expected": notes, "actual": "未核对"})
    for req in profile.get("required_roles", []):
        if req not in seen:
            violations.append({"gate": "structure", "where": None, "role": req, "item": "missing", "expected": req, "actual": None})

    fig_pt = roles.get("figure_interior", {}).get("size_pt")
    if raster_figures and not figure_fonts_verified:
        violations.append({
            "gate": "figure_interior",
            "where": None,
            "role": "figure_interior",
            "item": "raster_font",
            "expected": fig_pt,
            "actual": f"{raster_figures} 幅图的字号不在 OOXML 中",
        })
    eq = roles.get("equation", {})
    if equation_count and not mathtype_sizes_verified:
        violations.append({
            "gate": "equation",
            "where": None,
            "role": "equation",
            "item": "mathtype_sizes",
            "expected": {
                "full": eq.get("mt_full_pt"),
                "sub": eq.get("mt_sub_pt"),
                "ssub": eq.get("mt_ssub_pt"),
                "symbol": eq.get("mt_symbol_pt"),
                "subsymbol": eq.get("mt_subsymbol_pt"),
            },
            "actual": "未核对",
        })
    _check_columns(columns_path, policy, violations)

    for v in violations:
        if v.get("role") in ("abstract_en", "abstract_zh", "keywords_en", "keywords_zh"):
            v["gate"] = "abstracts"
    gates = {}
    for name in (
        "page_setup", "doc_grid", "headers", "structure", "roles", "headings", "body",
        "abstracts", "captions", "table_font", "figure_interior", "figures", "equation",
        "references", "biography", "columns", "line_break", "notation",
    ):
        bad = [v for v in violations if v.get("gate") == name]
        gates[name] = {"passed": not bad, "count": len(bad)}
    pending = [name for name, g in gates.items() if not g["passed"]]
    return {
        "gates": gates,
        "violations": violations[:80],
        "violation_count": len(violations),
        "layout_policy": policy,
        "stop": {"allowed": not pending, "pending_gates": pending},
    }


def _assign_front(front_slots, roles, styles, defaults, violations):
    seen = set()
    for index, p, text in front_slots:
        if not text:
            continue
        fields = paragraph_fields(p, styles, defaults)
        props = _props_from_fields(fields, p)
        if re.search(r"分类号|文章编号", text):
            role = "front_meta"
        elif props.get("size_pt", 0) >= 16 or (props.get("eastAsia") or "").startswith("黑体") and props.get("size_pt", 0) >= 16:
            role = "title_zh"
        elif (props.get("eastAsia") or "").startswith("仿宋") or _near(props.get("size_pt"), 14, 0.1):
            role = "authors"
        elif (props.get("eastAsia") or "").startswith("楷体"):
            role = "affiliation_zh"
        elif _near(props.get("size_pt"), 12, 0.1):
            role = "title_en"
        elif _near(props.get("size_pt"), 10.5, 0.1):
            role = "authors_en" if "authors_en" not in seen else "affiliation_zh"
        elif _near(props.get("size_pt"), 9, 0.1):
            role = "affiliation_en"
        else:
            role = "front_common"
        seen.add(role)
        spec = roles.get(role) or roles.get("front_common")
        runs = _run_props_list(p, styles, defaults)
        if spec:
            _check_role(role, spec, props, runs, index, violations, "roles")
    return seen


def _check_page(doc, zin, page, violations):
    sects = doc.findall(f".//{W}sectPr")
    if not sects:
        violations.append({"gate": "page_setup", "item": "sectPr", "expected": "present", "actual": None})
        return
    for idx, sect in enumerate(sects):
        pg = sect.find(f"{W}pgMar")
        for key in ("top", "bottom", "left", "right", "header", "footer"):
            if key not in page or pg is None:
                continue
            actual = _attr(pg, key)
            if actual is None or abs(int(actual) - int(page[key])) > 30:
                violations.append({"gate": "page_setup", "where": idx + 1, "item": key, "expected": page[key], "actual": actual})
        cols = sect.find(f"{W}cols")
        num = int(_attr(cols, "num")) if cols is not None and _attr(cols, "num") else 1
        expect_num = page.get("section1_cols") if idx == 0 else page.get("section2_cols")
        if expect_num and num != int(expect_num) and not (idx == 0 and num == int(expect_num)):
            if idx == 0 or idx == 1:
                violations.append({"gate": "page_setup", "where": idx + 1, "item": "cols", "expected": expect_num, "actual": num})
        if idx == 1 and page.get("col_space") and cols is not None:
            space = _attr(cols, "space")
            if space is None or abs(int(space) - int(page["col_space"])) > 20:
                violations.append({"gate": "page_setup", "where": 2, "item": "col_space", "expected": page["col_space"], "actual": space})
        grid = sect.find(f"{W}docGrid")
        pitch = _attr(grid, "linePitch") if grid is not None else None
        gtype = _attr(grid, "type") if grid is not None else None
        if page.get("line_pitch_twips"):
            if pitch is None or abs(int(pitch) - int(page["line_pitch_twips"])) > 8:
                violations.append({"gate": "doc_grid", "where": idx + 1, "item": "line_pitch_twips", "expected": page["line_pitch_twips"], "actual": pitch})
        if page.get("chars_per_line") and gtype not in ("linesAndChars", "snapToChars"):
            violations.append({"gate": "doc_grid", "where": idx + 1, "item": "grid_type", "expected": "linesAndChars", "actual": gtype})
        if page.get("lines_per_page") and pitch and pg is not None:
            psz = sect.find(f"{W}pgSz")
            height = int(_attr(psz, "h") or page.get("height_twips") or 0)
            top = int(_attr(pg, "top") or 0)
            bottom = int(_attr(pg, "bottom") or 0)
            got = int(round((height - top - bottom) / int(pitch))) if int(pitch) else None
            if got != int(page["lines_per_page"]):
                violations.append({"gate": "doc_grid", "where": idx + 1, "item": "lines_per_page", "expected": page["lines_per_page"], "actual": got})
        if page.get("odd_even_header") and sect.find(f"{W}evenAndOddHeaders") is None:
            violations.append({"gate": "headers", "where": idx + 1, "item": "odd_even_header", "expected": True, "actual": False})
        if page.get("different_first_page") and sect.find(f"{W}titlePg") is None:
            violations.append({"gate": "headers", "where": idx + 1, "item": "different_first_page", "expected": True, "actual": False})
    if "word/settings.xml" in zin.namelist() and "auto_hyphenation" in page:
        settings = ET.fromstring(zin.read("word/settings.xml"))
        hy = settings.find(f"{W}autoHyphenation")
        actual = hy is not None and _attr(hy, "val") not in ("0", "false", "off")
        if bool(page["auto_hyphenation"]) != bool(actual):
            violations.append({"gate": "page_setup", "item": "auto_hyphenation", "expected": page["auto_hyphenation"], "actual": actual})


def _check_tables(doc, spec, violations):
    if not spec:
        return
    for tbl in doc.findall(f".//{W}tbl"):
        borders = tbl.find(f"{W}tblPr/{W}tblBorders")
        if borders is None:
            violations.append({"gate": "table_font", "item": "tblBorders", "expected": "three-line", "actual": None})
            continue
        for edge in ("top", "bottom"):
            el = borders.find(f"{W}{edge}")
            sz = _attr(el, "sz") if el is not None else None
            if sz is None or abs(int(sz) - int(spec.get("border_top_bottom_eighths", 6))) > 1:
                violations.append({"gate": "table_font", "item": edge, "expected": spec.get("border_top_bottom_eighths"), "actual": sz})
        inside = borders.find(f"{W}insideH")
        sz = _attr(inside, "sz") if inside is not None else None
        if sz is None or abs(int(sz) - int(spec.get("border_inside_eighths", 4))) > 1:
            violations.append({"gate": "table_font", "item": "insideH", "expected": spec.get("border_inside_eighths"), "actual": sz})
        vert = borders.find(f"{W}insideV")
        if vert is not None and (_attr(vert, "val") not in (None, "nil", "none")) and int(_attr(vert, "sz") or "0") > 0:
            violations.append({"gate": "table_font", "item": "insideV", "expected": "nil", "actual": _attr(vert, "val")})


def _check_columns(columns_path, policy, violations):
    if not columns_path:
        violations.append({"gate": "columns", "item": "measurement", "expected": "col_metrics.json", "actual": None})
        return
    if not os.path.exists(columns_path):
        violations.append({"gate": "columns", "item": "measurement", "expected": columns_path, "actual": "missing"})
        return
    with open(columns_path, encoding="utf-8-sig") as f:
        data = json.load(f)
    if data.get("pages_match") is False:
        violations.append({"gate": "columns", "item": "pages_match", "expected": True, "actual": False})
    limit = float(policy.get("column_abs_d_pt", 5.25))
    bottom = float(policy.get("column_bottom_min_pt", 745))
    checked = 0
    for page in data.get("pages", []):
        if not page.get("L_bot") or not page.get("R_bot"):
            continue
        checked += 1
        if abs(page.get("dT") or 0) > limit or abs(page.get("dB") or 0) > limit:
            violations.append({"gate": "columns", "where": page.get("page"), "item": "delta", "expected": limit, "actual": {"dT": page.get("dT"), "dB": page.get("dB")}})
        if page.get("L_bot", 0) < bottom or page.get("R_bot", 0) < bottom:
            violations.append({"gate": "columns", "where": page.get("page"), "item": "bottom", "expected": bottom, "actual": {"L": page.get("L_bot"), "R": page.get("R_bot")}})
    if checked == 0:
        violations.append({"gate": "columns", "item": "two_column_pages", "expected": ">=1", "actual": 0})


def _self_test():
    base = {
        "page": {"top": 1814, "footer": 0},
        "roles": {"biography": {"photo_min_px": [260, 330], "size_pt": 9}, "abstract_zh": {"line_exact_pt": 14}},
        "layout_policy": {"column_abs_d_pt": 5.25, "optimize_by": "prose"},
    }
    patch = {"page": {"top": 1700}, "roles": {"abstract_zh": {"line_exact_pt": 16}}}
    merged, updated, kept = merge_profile(base, patch)
    assert merged["page"]["top"] == 1700
    assert merged["page"]["footer"] == 0
    assert merged["roles"]["biography"]["photo_min_px"] == [260, 330]
    assert merged["roles"]["abstract_zh"]["line_exact_pt"] == 16
    assert merged["layout_policy"]["optimize_by"] == "prose"
    assert "page.footer" in kept and "layout_policy.optimize_by" in kept
    assert "page.top" in updated and "roles.abstract_zh.line_exact_pt" in updated

    xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>摘要：仅修改这一处的行距为固定值16pt。</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1814" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>"""
    buf = os.path.join(tempfile.gettempdir(), "journal_profile_selftest.docx")
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'></Types>")
        z.writestr("word/document.xml", xml)
        z.writestr("word/styles.xml", "<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'></w:styles>")
    got = extract_patch(buf)
    assert got["roles"]["abstract_zh"]["line_exact_pt"] == 16
    assert "biography" not in got.get("roles", {})
    assert got["page"]["top"] == 1814
    assert "footer" not in got["page"]
    merged, updated, kept = merge_profile(base, got)
    assert merged["roles"]["biography"]["size_pt"] == 9
    assert merged["page"]["footer"] == 0
    assert "roles.biography.size_pt" in kept
    os.remove(buf)
    print("self-test ok", "updated", updated)


if __name__ == "__main__":
    _self_test()
