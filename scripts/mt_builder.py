import base64
from collections import Counter
import ctypes
import struct
import zlib

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

SCALE = 32  # 32 logical units = 1 pt (2304 units = 1 inch)

DEFAULT_OLE_B64 = (
    "eNrtVj1v01AUPc+OSUgjOeajiIKIMyC2lgGpAiREkrpQiZSojQSb5Ra3GBw7cWwgQkz8gPwFGPgB"
    "sLEhsTABMxP8ABbLKzHXzUsUVQxJ6NI2J3l+eec+59733r3H/v5N+fn2/cIv7MMtiOjFJ3FihGO8"
    "7SEPCHzci+N4QMczHCr8oZbiZyjxvjfblmODDbj08aFCg0O9hw4mwTykYc0neiB+Ffb4T33z6r/u"
    "EandFgaach82TEyLLAQ2up5x7zsz1LQKrb+BJsWxhScT+z9F/hlf07j+k/nWcC/6ftdo93cokmn8"
    "x7yGx/WfxDrQ9R4/t0HdS9yWppbh/AxHF5SLQgbTg1GGiNl+7u2v/Xm6VA3/cb3TNNXlxauq1goM"
    "33Id5EYt2ipkGg+Miyub1fo1RNc/tDBeNqf+I/6LSN5jfnwu8BebVPhl7veNNxThx7TEkM4l0SyD"
    "7Zyvmw/VNacZ+Oo9w9kNjF0TNf2l9wqnH1hOybbLRtvarriPzBqZ2lCkutUw2+q6+VzdcBuGA0Xc"
    "7DS2XJtMFTfwLNNLjFBS1bqqvfA9A4pQNTzb9H0q60xx6UqoF7vFpTt6qRYWLpVq0Vm5hEiLLkfd"
    "sk6/8nKFurCgvguzxBWJIFa+G52LusQv6Cu6FiXfkpxjdFJMINBqGSmWSOMUabeGFgIYpPwWqY9D"
    "z4F1Pno2liqr9H+T6m9yXjcZDgzT6P9B4jD7z+Z59ZRtY/spZEqNPDfJlCyva0SJuADMsT7hIZnE"
    "MMORwF+wMmQ7"
)


def get_default_ole_bytes():
    return zlib.decompress(base64.b64decode(DEFAULT_OLE_B64))


class LOGFONTW(ctypes.Structure):
    _fields_ = [
        ("lfHeight", ctypes.c_long),
        ("lfWidth", ctypes.c_long),
        ("lfEscapement", ctypes.c_long),
        ("lfOrientation", ctypes.c_long),
        ("lfWeight", ctypes.c_long),
        ("lfItalic", ctypes.c_byte),
        ("lfUnderline", ctypes.c_byte),
        ("lfStrikeOut", ctypes.c_byte),
        ("lfCharSet", ctypes.c_byte),
        ("lfOutPrecision", ctypes.c_byte),
        ("lfClipPrecision", ctypes.c_byte),
        ("lfQuality", ctypes.c_byte),
        ("lfPitchAndFamily", ctypes.c_byte),
        ("lfFaceName", ctypes.c_wchar * 32),
    ]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("tmHeight", ctypes.c_long),
        ("tmAscent", ctypes.c_long),
        ("tmDescent", ctypes.c_long),
        ("tmInternalLeading", ctypes.c_long),
        ("tmExternalLeading", ctypes.c_long),
        ("tmAveCharWidth", ctypes.c_long),
        ("tmMaxCharWidth", ctypes.c_long),
        ("tmWeight", ctypes.c_long),
        ("tmOverhang", ctypes.c_long),
        ("tmDigitizedAspectX", ctypes.c_long),
        ("tmDigitizedAspectY", ctypes.c_long),
        ("tmFirstChar", ctypes.c_wchar),
        ("tmLastChar", ctypes.c_wchar),
        ("tmDefaultChar", ctypes.c_wchar),
        ("tmBreakChar", ctypes.c_wchar),
        ("tmItalic", ctypes.c_byte),
        ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte),
        ("tmPitchAndFamily", ctypes.c_byte),
        ("tmCharSet", ctypes.c_byte),
    ]


class ABC(ctypes.Structure):
    _fields_ = [("abcA", ctypes.c_int), ("abcB", ctypes.c_uint), ("abcC", ctypes.c_int)]


_HDC = user32.GetDC(0)
_MEMDC = gdi32.CreateCompatibleDC(_HDC)
_FONT_CACHE = {}


def _get_hfont(face, pt_size, italic=False, bold=False):
    key = (face, round(pt_size, 2), bool(italic), bool(bold))
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    lf = LOGFONTW()
    lf.lfHeight = -int(round(pt_size * SCALE))
    lf.lfWeight = 700 if bold else 400
    lf.lfItalic = 1 if italic else 0
    lf.lfCharSet = 2 if face == "Symbol" else 0
    lf.lfQuality = 4
    lf.lfFaceName = face
    hf = gdi32.CreateFontIndirectW(ctypes.byref(lf))
    _FONT_CACHE[key] = hf
    return hf


def measure_text(text, face, pt_size, italic=False, bold=False):
    hf = _get_hfont(face, pt_size, italic, bold)
    old = gdi32.SelectObject(_MEMDC, hf)
    sz = SIZE()
    gdi32.GetTextExtentPoint32W(_MEMDC, text, len(text), ctypes.byref(sz))
    tm = TEXTMETRICW()
    gdi32.GetTextMetricsW(_MEMDC, ctypes.byref(tm))
    left_extra = 0
    right_extra = 0
    if text:
        abc0 = ABC()
        if gdi32.GetCharABCWidthsW(_MEMDC, ord(text[0]), ord(text[0]), ctypes.byref(abc0)):
            if abc0.abcA < 0:
                left_extra = -abc0.abcA
        abc1 = ABC()
        if gdi32.GetCharABCWidthsW(_MEMDC, ord(text[-1]), ord(text[-1]), ctypes.byref(abc1)):
            if abc1.abcC < 0:
                right_extra = -abc1.abcC
    gdi32.SelectObject(_MEMDC, old)
    return sz.cx, tm.tmAscent, tm.tmDescent, left_extra, right_extra


SYMBOL_MAP = {
    "α": (0x61, True, 0x03B1),
    "β": (0x62, True, 0x03B2),
    "γ": (0x67, True, 0x03B3),
    "δ": (0x64, True, 0x03B4),
    "ε": (0x65, True, 0x03B5),
    "η": (0x68, True, 0x03B7),
    "θ": (0x71, True, 0x03B8),
    "λ": (0x6C, True, 0x03BB),
    "μ": (0x6D, True, 0x03BC),
    "π": (0x70, True, 0x03C0),
    "ρ": (0x72, True, 0x03C1),
    "σ": (0x73, True, 0x03C3),
    "τ": (0x74, True, 0x03C4),
    "φ": (0x66, True, 0x03C6),
    "ω": (0x77, True, 0x03C9),
    "Δ": (0x44, False, 0x0394),
    "Σ": (0x53, False, 0x03A3),
    "Ω": (0x57, False, 0x03A9),
    "±": (0xB1, False, 0x00B1),
    "∓": (0xB1, False, 0x2213),
    "×": (0xB4, False, 0x00D7),
    "÷": (0xB8, False, 0x00F7),
    "·": (0xD7, False, 0x00B7),
    "−": (0x2D, False, 0x2212),
    "≤": (0xA3, False, 0x2264),
    "≥": (0xB3, False, 0x2265),
    "≠": (0xB9, False, 0x2260),
    "≈": (0xBB, False, 0x2248),
    "∈": (0xCE, False, 0x2208),
    "∞": (0xA5, False, 0x221E),
    "→": (0xAE, False, 0x2192),
    "′": (0xA2, False, 0x2032),
    "°": (0xB0, False, 0x00B0),
}


def measure_symbol_byte(byte_val, pt_size, italic=False):
    hf = _get_hfont("Symbol", pt_size, italic=italic, bold=False)
    old = gdi32.SelectObject(_MEMDC, hf)
    sz = SIZE()
    b = bytes([byte_val])
    gdi32.GetTextExtentPoint32A(_MEMDC, b, 1, ctypes.byref(sz))
    tm = TEXTMETRICW()
    gdi32.GetTextMetricsW(_MEMDC, ctypes.byref(tm))
    left_extra = int(round(pt_size * SCALE * 0.05)) if italic else 0
    right_extra = int(round(pt_size * SCALE * 0.12)) if italic else 0
    gdi32.SelectObject(_MEMDC, old)
    return sz.cx, tm.tmAscent, tm.tmDescent, left_extra, right_extra


def layout_node(node, sizes=(10.5, 6.5, 4.5), level=0):
    pt = sizes[min(level, len(sizes) - 1)]
    kind = node[0]
    if kind == "seq":
        x = 0
        max_asc = 0
        max_desc = 0
        ops = []
        for ch in node[1]:
            w, asc, desc, c_ops = layout_node(ch, sizes, level)
            for op in c_ops:
                if op[0] == "text":
                    ops.append(("text", op[1] + x, op[2], op[3], op[4], op[5], op[6], op[7]))
                elif op[0] == "line":
                    ops.append(("line", op[1] + x, op[2], op[3] + x, op[4], op[5]))
                else:
                    ops.append(("dot", op[1] + x, op[2], op[3]))
            x += w
            max_asc = max(max_asc, asc)
            max_desc = max(max_desc, desc)
        return x, max_asc, max_desc, ops

    if kind in ("var", "rm", "num"):
        txt = node[1]
        italic = kind == "var"
        w, asc, desc, le, re = measure_text(txt, "Times New Roman", pt, italic=italic)
        return le + w + re, asc, desc, [("text", le, 0, "Times New Roman", pt, italic, False, txt.encode("cp1252"))]

    if kind in ("op", "uop"):
        txt = node[1]
        if kind == "uop":
            pad_l = 0
            pad_r = 0
        elif txt in ("=", "+", "−", "<", ">", "≤", "≥", "≠", "≈", "∈", "→", "±", "∓", "×", "÷"):
            pad_l = pad_r = int(round(pt * SCALE * 0.14))
        elif txt == "·":
            pad_l = pad_r = int(round(pt * SCALE * 0.06))
        elif txt == "," and level == 0:
            pad_l = 0
            pad_r = int(round(pt * SCALE * 0.12))
        else:
            pad_l = pad_r = 0
        if txt in SYMBOL_MAP:
            bval, is_it, _ = SYMBOL_MAP[txt]
            w, asc, desc, le, re = measure_symbol_byte(bval, pt, italic=is_it)
            return pad_l + le + w + re + pad_r, asc, desc, [("text", pad_l + le, 0, "Symbol", pt, is_it, False, bytes([bval]))]
        w, asc, desc, le, re = measure_text(txt, "Times New Roman", pt, italic=False)
        return pad_l + le + w + re + pad_r, asc, desc, [("text", pad_l + le, 0, "Times New Roman", pt, False, False, txt.encode("cp1252"))]

    if kind == "sym":
        ch = node[1]
        bval, is_it, _ = SYMBOL_MAP[ch]
        w, asc, desc, le, re = measure_symbol_byte(bval, pt, italic=is_it)
        return le + w + re, asc, desc, [("text", le, 0, "Symbol", pt, is_it, False, bytes([bval]))]

    if kind == "thin":
        w = int(round(pt * SCALE * 0.22))
        return w, 0, 0, []

    if kind == "bar":
        bw, basc, bdesc, bops = layout_node(node[1], sizes, level)
        bar_y = -int(round(pt * SCALE * 0.78))
        thick = max(8, int(round(pt * SCALE * 0.055)))
        ops = list(bops) + [("line", 14, bar_y, max(24, bw - 8), bar_y, thick)]
        return bw, max(basc, -bar_y + thick + 14), bdesc, ops

    if kind == "dot":
        bw, basc, bdesc, bops = layout_node(node[1], sizes, level)
        dot_y = -int(round(pt * SCALE * 0.82))
        rad = max(14, int(round(pt * SCALE * 0.055)))
        cx_dot = bw // 2 + int(round(pt * SCALE * 0.06))
        ops = list(bops) + [("dot", cx_dot, dot_y, rad)]
        return bw, max(basc, -dot_y + rad + 16), bdesc, ops

    if kind in ("sub", "sup", "subsup"):
        bw, basc, bdesc, bops = layout_node(node[1], sizes, level)
        ops = list(bops)
        sub_shift = int(round(pt * SCALE * 0.24))
        sup_shift = -int(round(pt * SCALE * 0.38))
        max_asc, max_desc = basc, bdesc
        extra_w = 0
        if kind in ("sub", "subsup"):
            sw, sasc, sdesc, sops = layout_node(node[2], sizes, level + 1)
            off_x = bw - 6
            for op in sops:
                if op[0] == "text":
                    ops.append(("text", op[1] + off_x, op[2] + sub_shift, op[3], op[4], op[5], op[6], op[7]))
                elif op[0] == "line":
                    ops.append(("line", op[1] + off_x, op[2] + sub_shift, op[3] + off_x, op[4] + sub_shift, op[5]))
                else:
                    ops.append(("dot", op[1] + off_x, op[2] + sub_shift, op[3]))
            extra_w = max(extra_w, sw - 6)
            max_desc = max(max_desc, sub_shift + sdesc)
        if kind in ("sup", "subsup"):
            sup_node = node[2] if kind == "sup" else node[3]
            uw, uasc, udesc, uops = layout_node(sup_node, sizes, level + 1)
            off_x = bw + (4 if node[1][0] in ("rm", "num") else -2)
            for op in uops:
                if op[0] == "text":
                    ops.append(("text", op[1] + off_x, op[2] + sup_shift, op[3], op[4], op[5], op[6], op[7]))
                elif op[0] == "line":
                    ops.append(("line", op[1] + off_x, op[2] + sup_shift, op[3] + off_x, op[4] + sup_shift, op[5]))
                else:
                    ops.append(("dot", op[1] + off_x, op[2] + sup_shift, op[3]))
            extra_w = max(extra_w, off_x + uw - bw)
            max_asc = max(max_asc, -sup_shift + uasc)
        return bw + max(0, extra_w), max_asc, max_desc, ops

    raise ValueError(f"Unknown node {node}")


def build_wmf(node, pt=10.5, sub_pt=6.5, ssub_pt=4.5):
    w, asc, desc, ops = layout_node(node, (pt, sub_pt, ssub_pt), 0)
    pad_x = 12
    pad_top = 8
    pad_bot = 8
    cx = max(48, w + 2 * pad_x)
    cy = max(48, asc + desc + pad_top + pad_bot)
    base_y = pad_top + asc
    base_from_bot_pt = (cy - base_y) / SCALE

    recs = []

    def add_rec(func, payload=b""):
        assert len(payload) % 2 == 0
        words = 3 + len(payload) // 2
        recs.append(struct.pack("<IH", words, func) + payload)

    add_rec(0x0102, struct.pack("<HH", 1, 0))
    add_rec(0x0209, struct.pack("<I", 0x00000000))
    add_rec(0x012E, struct.pack("<HH", 0x0018, 0))
    add_rec(0x020B, struct.pack("<hh", 0, 0))
    add_rec(0x020C, struct.pack("<hh", cy, cx))

    for op in ops:
        if op[0] == "text":
            _, dx, dy, face, pt_sz, italic, bold, bstr = op
            lfH = -int(round(pt_sz * SCALE))
            weight = 700 if bold else 400
            charset = 2 if face == "Symbol" else 0
            face_b = face.encode("ascii")[:31].ljust(32, b"\x00")
            lf16 = struct.pack(
                "<hhhhhBBBBBBBB",
                lfH, 0, 0, 0, weight,
                1 if italic else 0, 0, 0, charset, 0, 0, 4, 0,
            ) + face_b
            add_rec(0x02FB, lf16)
            add_rec(0x012D, struct.pack("<H", 0))
            slen = len(bstr)
            b_pad = bstr + (b"\x00" if slen % 2 else b"")
            add_rec(0x0A32, struct.pack("<hhHH", base_y + dy, pad_x + dx, slen, 0) + b_pad)
            add_rec(0x01F0, struct.pack("<H", 0))
        elif op[0] == "line":
            _, x1, y1, x2, y2, thick = op
            add_rec(0x02FA, struct.pack("<HhhI", 0, thick, 0, 0x00000000))
            add_rec(0x012D, struct.pack("<H", 0))
            add_rec(0x0214, struct.pack("<hh", base_y + y1, pad_x + x1))
            add_rec(0x0213, struct.pack("<hh", base_y + y2, pad_x + x2))
            add_rec(0x01F0, struct.pack("<H", 0))
        elif op[0] == "dot":
            _, cx_d, cy_d, rad = op
            add_rec(0x02FC, struct.pack("<HIh", 0, 0x00000000, 0))
            add_rec(0x012D, struct.pack("<H", 0))
            add_rec(0x02FA, struct.pack("<HhhI", 5, 0, 0, 0))
            add_rec(0x012D, struct.pack("<H", 1))
            add_rec(0x0418, struct.pack("<hhhh", base_y + cy_d + rad, pad_x + cx_d + rad, base_y + cy_d - rad, pad_x + cx_d - rad))
            add_rec(0x01F0, struct.pack("<H", 0))
            add_rec(0x01F0, struct.pack("<H", 1))

    add_rec(0x0000, b"")
    body = b"".join(recs)
    max_rec_words = max(struct.unpack_from("<I", r, 0)[0] for r in recs)
    total_words = 9 + len(body) // 2
    meta_hdr = struct.pack("<HHHIHIH", 1, 9, 0x0300, total_words, 2, max_rec_words, 0)
    inch = 72 * SCALE
    aldus_wo_crc = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, 0, 0, cx, cy, inch, 0)
    crc = 0
    for i in range(0, 20, 2):
        crc ^= struct.unpack_from("<H", aldus_wo_crc, i)[0]
    aldus = aldus_wo_crc + struct.pack("<H", crc)
    return aldus + meta_hdr + body, cx / SCALE, cy / SCALE, base_from_bot_pt


def ast_to_mtef_records(node):
    kind = node[0]
    if kind == "seq":
        return b"".join(ast_to_mtef_records(ch) for ch in node[1])
    if kind in ("var", "rm", "num"):
        tf = {"rm": 0x81, "var": 0x83, "num": 0x88}[kind]
        out = bytearray()
        for ch in node[1]:
            out.extend(b"\x0f\x02\x02\x00" + bytes([tf]) + struct.pack("<H", ord(ch)))
        return bytes(out)
    if kind in ("sym", "op", "uop"):
        ch = node[1]
        if ch in SYMBOL_MAP:
            _, is_it, uni = SYMBOL_MAP[ch]
            tf = 0x84 if is_it else 0x86
            return b"\x0f\x02\x02\x00" + bytes([tf]) + struct.pack("<H", uni)
        return b"\x0f\x02\x02\x00\x86" + struct.pack("<H", ord(ch))
    if kind == "thin":
        return b"\x0f\x02\x02\x00\x86\x09\xef"
    if kind == "bar":
        inner = ast_to_mtef_records(node[1])
        return b"\x0f\x01\x03\x00\x11\x00\x00\x0a\x01\x00" + inner + b"\x00\x00"
    if kind == "dot":
        inner = ast_to_mtef_records(node[1])
        if inner.startswith(b"\x0f\x02\x02\x00") and len(inner) == 7:
            return b"\x0f\x02\x02\x01" + inner[4:] + b"\x06\x00\x02\x00"
        return inner
    if kind in ("sub", "sup", "subsup"):
        base_b = ast_to_mtef_records(node[1])
        if kind == "sub":
            sub_b = ast_to_mtef_records(node[2])
            return base_b + b"\x0f\x01\x03\x00\x1b\x00\x00\x0b\x01\x00" + sub_b + b"\x00\x00"
        if kind == "sup":
            sup_b = ast_to_mtef_records(node[2])
            return base_b + b"\x0f\x01\x03\x00\x1c\x00\x00\x0b\x01\x00" + sup_b + b"\x00\x00"
        sub_b = ast_to_mtef_records(node[2])
        sup_b = ast_to_mtef_records(node[3])
        return base_b + b"\x0f\x01\x03\x00\x1d\x00\x00\x0b\x01\x00" + sub_b + b"\x00\x0b\x01\x00" + sup_b + b"\x00\x00"
    return b""


def build_mtef5_bytes(template_eq_native, tex_str, ast_node):
    cbHdr = struct.unpack_from("<H", template_eq_native, 0)[0]
    mtef_orig = template_eq_native[cbHdr:]
    hdr = b"\x05\x01\x00\x07\x0cDSMT7\x00"
    cmt_body = b"TeX Input Language\x00" + tex_str.encode("utf-8") + b"\x00"
    cmt = bytes([0x01, 0x66, len(cmt_body)]) + cmt_body
    f_start = mtef_orig.find(b"\x13")
    l_start = mtef_orig.find(b"\x0f\x01\x01\x00")
    prefs_part = mtef_orig[f_start:l_start] if (f_start != -1 and l_start != -1) else b""
    line_hdr = b"\x0f\x01\x01\x00\x10\x00\x00\x00\x00\x00\x00\x00"
    body = ast_to_mtef_records(ast_node)
    mtef = hdr + cmt + prefs_part + line_hdr + body + b"\x00\x00"
    new_hdr = struct.pack("<HIHI", 28, 0x00020000, 0xC3DA, len(mtef)) + template_eq_native[12:28]
    return new_hdr + mtef


def replace_ole_eq_native(template_ole_bytes, new_eq_native):
    buf = bytearray(template_ole_bytes)
    sec_sz = 1 << struct.unpack_from("<H", buf, 30)[0]
    mini_sz = 1 << struct.unpack_from("<H", buf, 32)[0]
    num_fat = struct.unpack_from("<I", buf, 44)[0]
    dir_start = struct.unpack_from("<I", buf, 48)[0]
    minifat_start = struct.unpack_from("<I", buf, 60)[0]
    difat = list(struct.unpack_from("<109I", buf, 76))[:num_fat]
    fat = []
    for s in difat:
        if s >= 0xFFFFFFFC:
            break
        off = 512 + s * sec_sz
        fat.extend(struct.unpack_from(f"<{sec_sz // 4}I", buf, off))

    def sec_off(s):
        return 512 + s * sec_sz

    dir_secs = []
    cur = dir_start
    while cur < 0xFFFFFFFA:
        dir_secs.append(cur)
        cur = fat[cur]

    root_start = None
    eq_ent_loc = None
    for ds in dir_secs:
        d_off = sec_off(ds)
        for slot in range(sec_sz // 128):
            e_off = d_off + slot * 128
            nlen = struct.unpack_from("<H", buf, e_off + 64)[0]
            name = buf[e_off : e_off + max(0, nlen - 2)].decode("utf-16-le", errors="ignore")
            if name == "Root Entry":
                root_start = struct.unpack_from("<I", buf, e_off + 116)[0]
            elif name == "Equation Native":
                eq_ent_loc = e_off
                eq_mini_start, _ = struct.unpack_from("<II", buf, e_off + 116)

    minifat_off = sec_off(minifat_start)
    minifat = list(struct.unpack_from(f"<{sec_sz // 4}I", buf, minifat_off))
    chain = []
    c = eq_mini_start
    while c < 0xFFFFFFFA:
        chain.append(c)
        c = minifat[c]

    root_secs = []
    c = root_start
    while c < 0xFFFFFFFA:
        root_secs.append(c)
        c = fat[c]

    max_mini_avail = (len(root_secs) * sec_sz) // mini_sz
    needed_mini = (len(new_eq_native) + mini_sz - 1) // mini_sz
    if needed_mini > len(chain):
        used = {idx for idx, v in enumerate(minifat) if v != 0xFFFFFFFF}
        next_free = max(used) + 1 if used else 0
        while len(chain) < needed_mini and next_free < max_mini_avail:
            minifat[chain[-1]] = next_free
            minifat[next_free] = 0xFFFFFFFE
            chain.append(next_free)
            next_free += 1
        while len(chain) < needed_mini:
            new_sec = (len(buf) - 512) // sec_sz
            buf.extend(b"\x00" * sec_sz)
            fat[root_secs[-1]] = new_sec
            fat[new_sec] = 0xFFFFFFFE
            root_secs.append(new_sec)
            max_mini_avail = (len(root_secs) * sec_sz) // mini_sz
            fat_sec_off = 512 + difat[0] * sec_sz
            struct.pack_into(f"<{sec_sz // 4}I", buf, fat_sec_off, *fat[: sec_sz // 4])
            for ds in dir_secs:
                for slot in range(sec_sz // 128):
                    e_off = sec_off(ds) + slot * 128
                    nlen = struct.unpack_from("<H", buf, e_off + 64)[0]
                    nm = buf[e_off : e_off + max(0, nlen - 2)].decode("utf-16-le", errors="ignore")
                    if nm == "Root Entry":
                        struct.pack_into("<I", buf, e_off + 120, len(root_secs) * sec_sz)
            while len(chain) < needed_mini and next_free < max_mini_avail:
                minifat[chain[-1]] = next_free
                minifat[next_free] = 0xFFFFFFFE
                chain.append(next_free)
                next_free += 1
        struct.pack_into(f"<{len(minifat)}I", buf, minifat_off, *minifat)

    padded = new_eq_native.ljust(len(chain) * mini_sz, b"\x00")
    for idx, m_sec in enumerate(chain):
        byte_pos = m_sec * mini_sz
        r_sec = root_secs[byte_pos // sec_sz]
        in_sec = byte_pos % sec_sz
        abs_off = sec_off(r_sec) + in_sec
        buf[abs_off : abs_off + mini_sz] = padded[idx * mini_sz : (idx + 1) * mini_sz]

    struct.pack_into("<I", buf, eq_ent_loc + 120, len(new_eq_native))
    return bytes(buf)


def extract_ole_eq_native(ole_bytes):
    sec_sz = 1 << struct.unpack_from("<H", ole_bytes, 30)[0]
    mini_sz = 1 << struct.unpack_from("<H", ole_bytes, 32)[0]
    num_fat = struct.unpack_from("<I", ole_bytes, 44)[0]
    dir_start = struct.unpack_from("<I", ole_bytes, 48)[0]
    minifat_start = struct.unpack_from("<I", ole_bytes, 60)[0]
    difat = list(struct.unpack_from("<109I", ole_bytes, 76))[:num_fat]
    fat = []
    for s in difat:
        if s >= 0xFFFFFFFC:
            break
        fat.extend(struct.unpack_from(f"<{sec_sz // 4}I", ole_bytes, 512 + s * sec_sz))
    dir_secs = []
    cur = dir_start
    while cur < 0xFFFFFFFA:
        dir_secs.append(cur)
        cur = fat[cur]
    root_start = eq_mini_start = eq_len = None
    for ds in dir_secs:
        for slot in range(sec_sz // 128):
            e_off = 512 + ds * sec_sz + slot * 128
            nlen = struct.unpack_from("<H", ole_bytes, e_off + 64)[0]
            name = ole_bytes[e_off : e_off + max(0, nlen - 2)].decode("utf-16-le", errors="ignore")
            if name == "Root Entry":
                root_start = struct.unpack_from("<I", ole_bytes, e_off + 116)[0]
            elif name == "Equation Native":
                eq_mini_start, eq_len = struct.unpack_from("<II", ole_bytes, e_off + 116)
    minifat = list(struct.unpack_from(f"<{sec_sz // 4}I", ole_bytes, 512 + minifat_start * sec_sz))
    root_secs = []
    c = root_start
    while c < 0xFFFFFFFA:
        root_secs.append(c)
        c = fat[c]
    out = bytearray()
    c = eq_mini_start
    while c < 0xFFFFFFFA:
        bp = c * mini_sz
        abs_off = 512 + root_secs[bp // sec_sz] * sec_sz + (bp % sec_sz)
        out.extend(ole_bytes[abs_off : abs_off + mini_sz])
        c = minifat[c]
    return bytes(out[:eq_len])


TEX_MACROS = {
    "alpha": ("sym", "α"),
    "beta": ("sym", "β"),
    "gamma": ("sym", "γ"),
    "delta": ("sym", "δ"),
    "epsilon": ("sym", "ε"),
    "eta": ("sym", "η"),
    "theta": ("sym", "θ"),
    "lambda": ("sym", "λ"),
    "mu": ("sym", "μ"),
    "pi": ("sym", "π"),
    "rho": ("sym", "ρ"),
    "sigma": ("sym", "σ"),
    "tau": ("sym", "τ"),
    "phi": ("sym", "φ"),
    "omega": ("sym", "ω"),
    "Delta": ("sym", "Δ"),
    "Sigma": ("sym", "Σ"),
    "Omega": ("sym", "Ω"),
    "pm": ("op", "±"),
    "mp": ("op", "∓"),
    "times": ("op", "×"),
    "div": ("op", "÷"),
    "cdot": ("op", "·"),
    "leq": ("op", "≤"),
    "le": ("op", "≤"),
    "geq": ("op", "≥"),
    "ge": ("op", "≥"),
    "neq": ("op", "≠"),
    "ne": ("op", "≠"),
    "approx": ("op", "≈"),
    "in": ("op", "∈"),
    "to": ("op", "→"),
    "rightarrow": ("op", "→"),
    "infty": ("op", "∞"),
    "circ": ("op", "°"),
    "ell": ("var", "l"),
    "dagger": ("rm", "†"),
    "ddagger": ("rm", "‡"),
    ",": ("thin",),
    " ": ("thin",),
}
RM_FUNCS = {"min", "max", "exp", "ln", "sat", "SOC", "res", "gap", "dc", "rec", "lim", "av", "req", "ideal", "farm", "load", "db"}


def _is_value_node(node):
    if node is None:
        return False
    k = node[0]
    if k in ("var", "rm", "num", "sym", "sub", "sup", "subsup", "bar", "dot", "seq"):
        return True
    if k in ("op", "uop") and node[1] in (")", "]", "}", "|", "%"):
        return True
    return False


def tex_to_ast(tex):
    i = 0
    n = len(tex)

    def parse_seq(stop_char=None, roman_mode=False):
        nonlocal i
        items = []
        prev = None
        while i < n:
            ch = tex[i]
            if stop_char and ch == stop_char:
                i += 1
                break
            if ch == " ":
                i += 1
                continue
            atom = parse_atom(roman_mode=roman_mode, prev_is_val=_is_value_node(prev))
            if atom is not None:
                items.append(atom)
                prev = atom
        return items[0] if len(items) == 1 else ("seq", items)

    def parse_group(roman_mode=False):
        nonlocal i
        if i < n and tex[i] == "{":
            i += 1
            return parse_seq("}", roman_mode=roman_mode)
        return parse_atom(roman_mode=roman_mode, allow_scripts=False)

    def parse_atom(roman_mode=False, allow_scripts=True, prev_is_val=False):
        nonlocal i
        if i >= n:
            return None
        ch = tex[i]
        if ch == "{":
            base = parse_group(roman_mode=roman_mode)
        elif ch == "\\":
            i += 1
            if i < n and not tex[i].isalpha():
                cmd = tex[i]
                i += 1
                if cmd in ("{", "}"):
                    base = ("op", cmd)
                elif cmd == "%":
                    base = ("op", "%")
                elif cmd in TEX_MACROS:
                    mac = TEX_MACROS[cmd]
                    if mac[0] == "op" and mac[1] in ("±", "∓") and not prev_is_val:
                        base = ("uop", mac[1])
                    else:
                        base = mac
                else:
                    base = ("op", cmd)
            else:
                j = i
                while j < n and tex[j].isalpha():
                    j += 1
                cmd = tex[i:j]
                i = j
                if cmd in ("mathrm", "text"):
                    base = parse_group(roman_mode=True)
                elif cmd == "bar":
                    base = ("bar", parse_group(roman_mode=roman_mode))
                elif cmd == "dot":
                    base = ("dot", parse_group(roman_mode=roman_mode))
                elif cmd in TEX_MACROS:
                    mac = TEX_MACROS[cmd]
                    if mac[0] == "op" and mac[1] in ("±", "∓") and not prev_is_val:
                        base = ("uop", mac[1])
                    else:
                        base = mac
                elif cmd in RM_FUNCS:
                    base = ("rm", cmd)
                else:
                    base = ("rm", cmd)
        elif ch.isdigit() or ch == ".":
            j = i
            while j < n and (tex[j].isdigit() or tex[j] == "."):
                j += 1
            base = ("num", tex[i:j])
            i = j
        elif ch.isalpha():
            matched_rm = None
            if not roman_mode:
                for kw in ("SOC", "min", "max", "exp", "ln", "sat", "res", "gap", "rec", "lim", "av", "req", "ideal", "farm", "load", "dc", "db"):
                    if tex.startswith(kw, i) and (i + len(kw) == n or not tex[i + len(kw)].isalpha()):
                        matched_rm = kw
                        break
            if matched_rm:
                base = ("rm", matched_rm)
                i += len(matched_rm)
            elif roman_mode:
                j = i
                while j < n and tex[j].isalpha():
                    j += 1
                base = ("rm", tex[i:j])
                i = j
            else:
                base = ("var", ch)
                i += 1
        elif ch in ("-", "−"):
            base = ("op", "−") if prev_is_val else ("uop", "−")
            i += 1
        elif ch == "+":
            base = ("op", "+") if prev_is_val else ("uop", "+")
            i += 1
        elif ch in ("=", "<", ">", "*", "/", "(", ")", "[", "]", ",", ":", "|", "·", "%"):
            base = ("op", ch)
            i += 1
        elif ch == "'":
            base = ("op", "′")
            i += 1
        elif ch in SYMBOL_MAP:
            base = ("sym", ch) if ch.isalpha() else ("op", ch)
            i += 1
        else:
            base = ("rm", ch)
            i += 1

        if not allow_scripts:
            return base

        sub_node = None
        sup_node = None
        while i < n and (tex[i] in ("_", "^") or tex[i] == "'"):
            if tex[i] == "'":
                i += 1
                sup_node = ("op", "′")
                continue
            kind = tex[i]
            i += 1
            arg = parse_group(roman_mode=False)
            if kind == "_":
                sub_node = arg
            else:
                sup_node = arg
        if sub_node is not None and sup_node is not None:
            return ("subsup", base, sub_node, sup_node)
        if sub_node is not None:
            return ("sub", base, sub_node)
        if sup_node is not None:
            return ("sup", base, sup_node)
        return base

    return parse_seq()


def inspect_wmf_metrics(wmf_bytes):
    """Read bounding box and main-text baseline (points from bottom) from an existing WMF."""
    _, _, left, top, right, bot, inch, _ = struct.unpack_from("<IHhhhhHI", wmf_bytes, 0)
    scale = inch / 72.0 if inch else 32.0
    w_pt = (right - left) / scale
    h_pt = (bot - top) / scale
    pos = 22 + 18
    text_align = 0
    objs = []
    cur_lfh = 0
    cur_face = ""
    cur_move_y = None
    main_ys = []
    any_ys = []
    while pos + 6 <= len(wmf_bytes):
        words, func = struct.unpack_from("<IH", wmf_bytes, pos)
        if func == 0 or words < 3:
            break
        rec_len = words * 2
        if func == 0x012E and rec_len >= 8:
            text_align = struct.unpack_from("<H", wmf_bytes, pos + 6)[0]
        elif func in (0x02FB, 0x02FA, 0x02FC, 0x02F7, 0x02F8, 0x02F9, 0x00F7):
            val = None
            if func == 0x02FB and rec_len >= 24:
                lfh = struct.unpack_from("<h", wmf_bytes, pos + 6)[0]
                face_raw = wmf_bytes[pos + 24 : pos + min(rec_len, 56)]
                face = face_raw.split(b"\x00", 1)[0].decode("ascii", errors="ignore")
                val = ("font", lfh, face)
            try:
                idx = objs.index(None)
                objs[idx] = val
            except ValueError:
                objs.append(val)
        elif func == 0x01F0 and rec_len >= 8:
            idx = struct.unpack_from("<H", wmf_bytes, pos + 6)[0]
            if 0 <= idx < len(objs):
                objs[idx] = None
        elif func == 0x012D and rec_len >= 8:
            idx = struct.unpack_from("<H", wmf_bytes, pos + 6)[0]
            if 0 <= idx < len(objs) and objs[idx] and objs[idx][0] == "font":
                cur_lfh, cur_face = objs[idx][1], objs[idx][2]
        elif func == 0x0214 and rec_len >= 10:
            cur_move_y = struct.unpack_from("<h", wmf_bytes, pos + 6)[0]
        elif func == 0x0A32 and rec_len >= 14:
            ey = cur_move_y if (text_align & 1) and cur_move_y is not None else struct.unpack_from("<h", wmf_bytes, pos + 6)[0]
            if cur_lfh <= -340 and cur_face != "MT Extra":
                main_ys.append(ey)
            elif cur_lfh <= -340:
                any_ys.append(ey)
        pos += rec_len
    if main_ys:
        cnt = Counter(main_ys)
        ref_y = cnt.most_common(1)[0][0]
        base_from_bot_pt = (bot - ref_y) / scale
    elif any_ys:
        ref_y = max(any_ys)
        base_from_bot_pt = (bot - ref_y) / scale
    else:
        base_from_bot_pt = h_pt * 0.25
    return w_pt, h_pt, base_from_bot_pt
