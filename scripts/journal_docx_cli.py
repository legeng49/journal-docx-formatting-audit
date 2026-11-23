import argparse
import copy
import ctypes
import io
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

import mt_builder as mb
import template_profile as tp


def open_docx_bytes(path):
    """Read a .docx file into memory even if it is currently open/locked by Microsoft Word on Windows."""
    try:
        with open(path, "rb") as f:
            return io.BytesIO(f.read())
    except PermissionError:
        k32 = ctypes.windll.kernel32
        GENERIC_READ = 0x80000000
        FILE_SHARE_ALL = 0x00000001 | 0x00000002 | 0x00000004
        OPEN_EXISTING = 3
        h = k32.CreateFileW(os.path.abspath(path), GENERIC_READ, FILE_SHARE_ALL, None, OPEN_EXISTING, 0, None)
        if h in (-1, 0xFFFFFFFF, 0xFFFFFFFFFFFFFFFF):
            raise
        try:
            sz = k32.GetFileSize(h, None)
            buf = ctypes.create_string_buffer(sz)
            read_n = ctypes.c_ulong(0)
            k32.ReadFile(h, buf, sz, ctypes.byref(read_n), None)
            return io.BytesIO(buf.raw[: read_n.value])
        finally:
            k32.CloseHandle(h)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"
XML_NS = "http://www.w3.org/XML/1998/namespace"

W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
O = f"{{{O_NS}}}"
V = f"{{{V_NS}}}"

NS_PREFIXES = {
    "w": W_NS,
    "r": R_NS,
    "o": O_NS,
    "v": V_NS,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
for pfx, uri in NS_PREFIXES.items():
    ET.register_namespace(pfx, uri)

SPECIAL_CHARS = set("±−·≤≥∈†‡αβγδωΔητρλμσπθφΩΣ×÷≠≈∞")
FORBIDDEN_PHRASES = [
    "揭示了",
    "直接印证",
    "刚性前置约束",
    "清晰解耦",
    "精确刻画",
    "显著收缩",
    "显著偏离",
    "显著扩大",
    "显著降低",
    "显著",
    "更为关键的是",
    "综合上述",
    "可以看出",
    "可以看到",
    "源—网—储",
    "不仅体现在",
    "更取决于",
    "名义补偿比例并非决定频率调控效果的充分条件",
    "单纯削减风电并网功率缺口不足以保证峰谷频率波动同步收敛",
    "研究表明，转速恢复整定应在满足回补期限的前提下统筹",
]
UNIT_HEAD_RE = re.compile(r"^\s*(MW|kW|GW|MWh|kWh|MJ|kJ|Hz|mHz|s|ms|pu|rad/s)\b")


def el(tag, attrib=None, children=None, text=None):
    e = ET.Element(tag, attrib or {})
    if text is not None:
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            e.set(f"{{{XML_NS}}}space", "preserve")
        e.text = text
    if children:
        for c in children:
            e.append(c)
    return e


def cmd_audit(args):
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    report = {
        "input": os.path.abspath(args.input),
        "template": args.template,
        "sections": [],
        "trailing_empty_single_col_section": False,
        "leaked_special_chars": [],
        "post_ole_plain_units": [],
        "ole_diagnostics": {"total_oles": 0, "non_dsmt4": 0, "unpositioned_oles": 0},
        "displayed_equations": {"total": 0, "non_atLeast_or_bad_padding": []},
        "cross_references": {"total_ref_fields": 0, "missing_charformat": [], "hardcoded_section_refs": []},
        "figure_order": {"mentions": [], "captions": [], "out_of_order": False},
        "forbidden_phrases": [],
        "passed": True,
    }

    with zipfile.ZipFile(open_docx_bytes(args.input), "r") as zin:
        doc = ET.fromstring(zin.read("word/document.xml"))
        body = doc.find(f"{W}body")
        children = list(body)

        sects = doc.findall(f".//{W}sectPr")
        for idx, s in enumerate(sects):
            pgm = s.find(f"{W}pgMar")
            cols = s.find(f"{W}cols")
            s_info = {
                "index": idx + 1,
                "top": pgm.get(f"{W}top") if pgm is not None else None,
                "bottom": pgm.get(f"{W}bottom") if pgm is not None else None,
                "left": pgm.get(f"{W}left") if pgm is not None else None,
                "right": pgm.get(f"{W}right") if pgm is not None else None,
                "cols_num": cols.get(f"{W}num") if cols is not None else "1",
                "cols_space": cols.get(f"{W}space") if cols is not None else None,
            }
            report["sections"].append(s_info)

        if len(report["sections"]) > 2 and report["sections"][-1]["cols_num"] != "2":
            report["trailing_empty_single_col_section"] = True

        fig_mentions = []
        fig_captions = []

        for p_idx, ch in enumerate(children):
            p_txt = "".join(t.text for t in ch.findall(f".//{W}t") if t.text)
            for fp in FORBIDDEN_PHRASES:
                if fp in p_txt:
                    report["forbidden_phrases"].append({"p_index": p_idx, "phrase": fp, "snippet": p_txt[:60]})

            # Check hardcoded section refs without REF field
            instrs = [i.text or "" for i in ch.findall(f".//{W}instrText")]
            for sec_pat in ("第1节", "详见第2节", "第2.2节"):
                if sec_pat in p_txt and not any("REF " in ins for ins in instrs):
                    report["cross_references"]["hardcoded_section_refs"].append({"p_index": p_idx, "ref": sec_pat})

            for ins in instrs:
                if "REF " in ins:
                    report["cross_references"]["total_ref_fields"] += 1
                    if "CHARFORMAT" not in ins:
                        report["cross_references"]["missing_charformat"].append({"p_index": p_idx, "instr": ins.strip()})

            # Check caption vs mention order for figures
            m_cap = re.match(r"^\s*图\s*(\d+)\s*[　\s]", p_txt)
            if m_cap:
                fig_captions.append((int(m_cap.group(1)), p_idx))
            else:
                for m_f in re.finditer(r"图\s*(\d+)", p_txt):
                    fnum = int(m_f.group(1))
                    if not any(x[0] == fnum for x in fig_mentions):
                        fig_mentions.append((fnum, p_idx))

            full_instr = " ".join(instrs)
            if "SEQ" in full_instr and "公式" in full_instr:
                report["displayed_equations"]["total"] += 1

            # Inspect runs for leaked special chars, OLE objects, and post-OLE plain-text units
            runs = ch.findall(f".//{W}r")
            prev_was_ole = False
            for r_el in runs:
                oo = r_el.find(f".//{O}OLEObject")
                if oo is not None:
                    report["ole_diagnostics"]["total_oles"] += 1
                    if oo.get("ProgID") != "Equation.DSMT4":
                        report["ole_diagnostics"]["non_dsmt4"] += 1
                    rPr = r_el.find(f"{W}rPr")
                    if rPr is None or rPr.find(f"{W}position") is None:
                        report["ole_diagnostics"]["unpositioned_oles"] += 1
                    prev_was_ole = True
                else:
                    t_el = r_el.find(f"{W}t")
                    txt = t_el.text if (t_el is not None and t_el.text) else ""
                    if txt:
                        bad = [c for c in txt if c in SPECIAL_CHARS]
                        if bad:
                            report["leaked_special_chars"].append({"p_index": p_idx, "chars": "".join(set(bad)), "text": txt})
                        if prev_was_ole and UNIT_HEAD_RE.match(txt):
                            report["post_ole_plain_units"].append({"p_index": p_idx, "text": txt[:20]})
                        prev_was_ole = False

        report["figure_order"]["mentions"] = fig_mentions
        report["figure_order"]["captions"] = fig_captions
        cap_nums = [c[0] for c in fig_captions]
        if cap_nums != sorted(cap_nums):
            report["figure_order"]["out_of_order"] = True

        profile = tp.load_profile(args.profile)
        layout = tp.evaluate(
            zin,
            profile,
            columns_path=args.columns,
            figure_fonts_verified=args.figure_fonts_verified,
            mathtype_sizes_verified=args.mathtype_sizes_verified,
            notation_verified=args.notation_verified,
        )
        report["profile"] = os.path.abspath(args.profile)
        report["gates"] = layout["gates"]
        report["layout_violations"] = layout["violations"]
        report["layout_violation_count"] = layout["violation_count"]
        report["layout_policy"] = layout["layout_policy"]

    legacy = {
        "trailing_section": not report["trailing_empty_single_col_section"],
        "special_chars": not report["leaked_special_chars"] and not report["post_ole_plain_units"],
        "ole": report["ole_diagnostics"]["non_dsmt4"] == 0 and report["ole_diagnostics"]["unpositioned_oles"] == 0,
        "cross_references": not report["cross_references"]["hardcoded_section_refs"] and not report["cross_references"]["missing_charformat"],
        "figure_order": not report["figure_order"]["out_of_order"],
        "prose": not report["forbidden_phrases"],
    }
    report.setdefault("gates", {})
    for name, ok in legacy.items():
        report["gates"][name] = {"passed": ok, "count": 0 if ok else 1}
    pending = [name for name, gate in report["gates"].items() if not gate["passed"]]
    report["passed"] = not pending
    report["stop"] = {
        "allowed": not pending,
        "pending_gates": pending,
        "rule": "全部闸门通过后才可以停止。模板未改写的规定仍然有效。栏面不足时改表述，不压行距、不缩图、不缩小图表字体。",
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"Success! Audit report written to: {args.output} (passed={report['passed']})")


def fix_existing_ole_run(r_el, zin, rid_to_target, scale=0.875, is_inline=False):
    im = r_el.find(f".//{V}imagedata")
    sh = r_el.find(f".//{V}shape")
    obj = r_el.find(f".//{W}object")
    if im is None or sh is None:
        return None
    wmf_rid = im.attrib.get(f"{R}id")
    if not wmf_rid or wmf_rid not in rid_to_target:
        return None
    wmf_bytes = zin.read("word/" + rid_to_target[wmf_rid])
    w_orig, h_orig, base_orig = mb.inspect_wmf_metrics(wmf_bytes)
    w_pt = w_orig * scale
    h_pt = h_orig * scale
    crop_frac = 0.0
    if is_inline and h_pt > 15.6:
        target_h = 15.3 if h_orig < 21.0 else 15.5
        crop_frac = max(0.0, min(0.08, (1.0 - target_h / h_pt) / 2.0))
        if crop_frac > 0.005:
            cf_units = int(round(crop_frac * 65536))
            crop_frac = cf_units / 65536.0
            im.set("croptop", f"{cf_units}f")
            im.set("cropbottom", f"{cf_units}f")
            h_pt = h_pt * (1.0 - 2.0 * crop_frac)
    sh.set("style", f"width:{w_pt:.2f}pt;height:{h_pt:.2f}pt")
    if obj is not None:
        obj.set(f"{W}dxaOrig", str(int(round(w_pt * 20))))
        obj.set(f"{W}dyaOrig", str(int(round(h_pt * 20))))
    eff_base = (base_orig - crop_frac * h_orig) * scale
    pos_val = -int(round(2.0 * eff_base))
    rPr = r_el.find(f"{W}rPr")
    if rPr is None:
        rPr = el(f"{W}rPr")
        r_el.insert(0, rPr)
    pos_el = rPr.find(f"{W}position")
    if pos_el is None:
        rPr.append(el(f"{W}position", {f"{W}val": str(pos_val)}))
    else:
        pos_el.set(f"{W}val", str(pos_val))
    return w_pt, h_pt, pos_val


def cmd_build_mt(args):
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    profile = tp.load_profile(getattr(args, "profile", None))
    eq = profile.get("roles", {}).get("equation", {})
    stats = {
        "normalized_oles": 0,
        "displayed_eqs_updated": 0,
        "equation_spacing": {
            "line": eq.get("line", "single"),
            "before_twips": eq.get("before_twips", 0),
            "after_twips": eq.get("after_twips", 0),
        },
        "required_mathtype_pt": {
            "full": eq.get("mt_full_pt"),
            "sub": eq.get("mt_sub_pt"),
            "ssub": eq.get("mt_ssub_pt"),
            "symbol": eq.get("mt_symbol_pt"),
            "subsymbol": eq.get("mt_subsymbol_pt"),
        },
    }
    with zipfile.ZipFile(open_docx_bytes(args.input), "r") as zin:
        doc = ET.fromstring(zin.read("word/document.xml"))
        rels_xml = zin.read("word/_rels/document.xml.rels").decode("utf-8")
        rels_root = ET.fromstring(rels_xml)
        rid_to_target = {e.attrib["Id"]: e.attrib["Target"] for e in rels_root}

        for p in doc.findall(f".//{W}p"):
            instrs = [i.text or "" for i in p.findall(f".//{W}instrText")]
            full_instr = " ".join(instrs)
            is_disp_eq = "SEQ" in full_instr and "公式" in full_instr
            h_scaled = 17.5
            for r_el in p.findall(f".//{W}r"):
                if r_el.find(f".//{O}OLEObject") is not None:
                    res = fix_existing_ole_run(r_el, zin, rid_to_target, scale=args.scale, is_inline=(not is_disp_eq))
                    if res:
                        h_scaled = res[1]
                        stats["normalized_oles"] += 1
            if is_disp_eq:
                pPr = p.find(f"{W}pPr")
                if pPr is None:
                    pPr = el(f"{W}pPr")
                    p.insert(0, pPr)
                sp = pPr.find(f"{W}spacing")
                if sp is None:
                    sp = el(f"{W}spacing")
                    pPr.append(sp)
                if eq.get("line") == "atLeast":
                    min_twips = int(round(max(h_scaled, 17.5) * 20))
                    sp.set(f"{W}line", str(min_twips))
                    sp.set(f"{W}lineRule", "atLeast")
                elif eq.get("line_exact_pt"):
                    sp.set(f"{W}line", str(int(round(float(eq["line_exact_pt"]) * 20))))
                    sp.set(f"{W}lineRule", "exact")
                else:
                    sp.set(f"{W}line", "240")
                    sp.set(f"{W}lineRule", "auto")
                before = eq.get("before_twips", 0 if "before_pt" not in eq else int(round(float(eq["before_pt"]) * 20)))
                after = eq.get("after_twips", 0 if "after_pt" not in eq else int(round(float(eq["after_pt"]) * 20)))
                sp.set(f"{W}before", str(int(before)))
                sp.set(f"{W}after", str(int(after)))
                stats["displayed_eqs_updated"] += 1

        new_doc_xml = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with zipfile.ZipFile(args.output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    with open(args.report, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f"Success! MathType normalization written to: {args.output}, report: {args.report}")


def cmd_import_template(args):
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    report = tp.import_template(args.input, args.profile, args.spec)
    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    with open(args.report, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(
        f"Success! Profile updated: {report['updated'].__len__()} fields from the template, "
        f"{report['kept_count']} fields kept. Report: {args.report}"
    )


def cmd_fix_refs(args):
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    stats = {"updated_ref_fields": 0}
    with zipfile.ZipFile(open_docx_bytes(args.input), "r") as zin:
        doc = ET.fromstring(zin.read("word/document.xml"))
        for ins in doc.findall(f".//{W}instrText"):
            if ins.text and "REF " in ins.text and "CHARFORMAT" not in ins.text:
                ins.text = ins.text.rstrip() + " \\* CHARFORMAT "
                stats["updated_ref_fields"] += 1

        new_doc_xml = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with zipfile.ZipFile(args.output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    with open(args.report, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f"Success! Cross-references updated in: {args.output}, report: {args.report}")


def main():
    parser = argparse.ArgumentParser(description="Journal Word (.docx) Formatting Audit & MathType CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_audit = sub.add_parser("audit", help="Run full layout, MathType, cross-reference, and prose audit")
    p_audit.add_argument("--input", required=True, help="Input .docx path")
    p_audit.add_argument("--template", default="csee", help="Template profile name (default: csee)")
    p_audit.add_argument("--profile", default=tp.DEFAULT_PROFILE, help="Active profile JSON")
    p_audit.add_argument("--columns", default=None, help="col_metrics.json from measure_and_export.ps1")
    p_audit.add_argument("--output", required=True, help="Output JSON report path")
    p_audit.add_argument("--figure-fonts-verified", action="store_true", help="Set only after every raster figure's interior type is the profile figure_interior size")
    p_audit.add_argument("--mathtype-sizes-verified", action="store_true", help="Set only after MathType sizes match roles.equation mt_*_pt")
    p_audit.add_argument("--notation-verified", action="store_true", help="Set only after GB 3100, figure line weight, axis labels, and oscillograph grid labels have been checked")
    p_audit.set_defaults(func=cmd_audit)

    p_imp = sub.add_parser("import-template", help="Merge a .doc/.docx template into the profile; unspecified fields stay unchanged")
    p_imp.add_argument("--input", required=True, help="New template .doc or .docx")
    p_imp.add_argument("--profile", default=tp.DEFAULT_PROFILE, help="Profile JSON to update in place")
    p_imp.add_argument("--spec", default=tp.DEFAULT_SPEC, help="Generated spec markdown")
    p_imp.add_argument("--report", required=True, help="Merge report JSON")
    p_imp.set_defaults(func=cmd_import_template)

    p_mt = sub.add_parser("build-mt", help="Normalize MathType OLE baselines, aspect ratios, and equation line spacing")
    p_mt.add_argument("--input", required=True, help="Input .docx path")
    p_mt.add_argument("--output", required=True, help="Output .docx path")
    p_mt.add_argument("--report", required=True, help="Output JSON report path")
    p_mt.add_argument("--scale", type=float, default=0.875, help="Uniform scale factor for OLE formulas")
    p_mt.add_argument("--profile", default=tp.DEFAULT_PROFILE, help="Active profile JSON")
    p_mt.set_defaults(func=cmd_build_mt)

    p_refs = sub.add_parser("fix-refs", help="Ensure all REF fields include \\* CHARFORMAT")
    p_refs.add_argument("--input", required=True, help="Input .docx path")
    p_refs.add_argument("--output", required=True, help="Output .docx path")
    p_refs.add_argument("--report", required=True, help="Output JSON report path")
    p_refs.set_defaults(func=cmd_fix_refs)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
