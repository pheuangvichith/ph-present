#!/usr/bin/env python3
"""Zero-dependency OOXML PowerPoint (.pptx) builder for ph-present.

Supports:
- Strict 16:9 (10 x 5.625 in) OOXML package generation
- <p159:morph option="byObject|byWord|byChar"/> transitions
- SVG path -> native PowerPoint Freeform (<a:custGeom>) with Bezier curves
- Seamless Split-Curtain polygon pairs sharing exact central seam coordinates
- Solid, linear, & radial glow gradient fills (<a:gradFill>), soft outer shadows (<a:outerShdw>)
- True rotation (<a:xfrm rot="...">) around shape centers & <a:scene3d>/<a:sp3d> (Desktop)
- Compound-Path Portal Masks (<a:custGeom> outer + inner cutout hole)
- Crop-to-Pan (<a:srcRect>) picture frames & video media frames
- Native editable clustered/stacked column & bar, line, and doughnut charts (<c:chart>) + .xlsx
- <p:timing> post-Morph stagger animations
- <p:embeddedFontLst> OOXML .odttf font embedding
- Script-aware XML font binding (<a:latin>, <a:cs>, <a:ea>) for Thai/Lao/CJK/Latin
"""

from __future__ import annotations

import io
import math
import os
import re
import uuid
import zipfile
from xml.sax.saxutils import escape

EMU_PER_INCH = 914400
SLIDE_W = int(10.0 * EMU_PER_INCH)     # 9144000
SLIDE_H = int(5.625 * EMU_PER_INCH)    # 5143500

_TOKEN_RE = re.compile(r"[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?")


def emu(inches: float) -> int:
    return int(round(inches * EMU_PER_INCH))


def rot_val(deg: float) -> int:
    """Convert degrees to OOXML 60000ths of a degree."""
    return int(round((deg % 360.0) * 60000))


def _is_thai(ch: str) -> bool:
    return "\u0E00" <= ch <= "\u0E7F"


def _is_lao(ch: str) -> bool:
    return "\u0E80" <= ch <= "\u0EFF"


def _is_cjk(ch: str) -> bool:
    code = ord(ch)
    return (0x4E00 <= code <= 0x9FFF) or (0x3040 <= code <= 0x30FF) or (0xAC00 <= code <= 0xD7AF)


def detect_script(ch: str) -> str:
    if _is_thai(ch):
        return "thai"
    if _is_lao(ch):
        return "lao"
    if _is_cjk(ch):
        return "cjk"
    return "latin"


def split_script_runs(text: str) -> list[tuple[str, str]]:
    """Split text into contiguous script runs so <a:rPr lang=...> binds accurately."""
    if not text:
        return [("latin", "")]
    runs: list[tuple[str, str]] = []
    cur_script = detect_script(text[0])
    buf = [text[0]]
    for ch in text[1:]:
        s = cur_script if ch == " " else detect_script(ch)
        if s == cur_script:
            buf.append(ch)
        else:
            runs.append((cur_script, "".join(buf)))
            cur_script = s
            buf = [ch]
    runs.append((cur_script, "".join(buf)))
    return runs


def rpr_xml(
    size_pt: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
    latin_font: str = "Arial",
    thai_font: str = "Noto Sans Thai",
    lao_font: str = "Noto Sans Lao",
    cjk_font: str = "Noto Sans CJK SC",
    mono: bool = False,
    script: str = "latin",
) -> str:
    sz = int(round(size_pt * 100))
    b_attr = ' b="1"' if bold else ' b="0"'
    i_attr = ' i="1"' if italic else ""
    lf = "DejaVu Sans Mono" if mono else latin_font
    lang_map = {"latin": "en-US", "thai": "th-TH", "lao": "lo-LA", "cjk": "zh-CN"}
    lang = lang_map.get(script, "en-US")
    cs_font = lao_font if script == "lao" else thai_font
    alt_lang = "th-TH" if script != "lao" else "lo-LA"
    return (
        f'<a:rPr lang="{lang}" altLang="{alt_lang}" sz="{sz}"{b_attr}{i_attr} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{escape(lf)}"/>'
        f'<a:ea typeface="{escape(cjk_font)}"/>'
        f'<a:cs typeface="{escape(cs_font)}"/>'
        f"</a:rPr>"
    )


def paragraph_xml(
    text: str,
    size_pt: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
    align: str = "l",
    line_spacing: float = 1.22,
    space_after_pt: float = 0.0,
    latin_font: str = "Arial",
    thai_font: str = "Noto Sans Thai",
    lao_font: str = "Noto Sans Lao",
    mono: bool = False,
) -> str:
    spc_pct = int(round(line_spacing * 100000))
    spc_aft = int(round(space_after_pt * 100))
    ppr = (
        f'<a:pPr algn="{align}">'
        f'<a:lnSpc><a:spcPct val="{spc_pct}"/></a:lnSpc>'
        f'<a:spcAft><a:spcPts val="{spc_aft}"/></a:spcAft>'
        f"<a:buNone/>"
        f"</a:pPr>"
    )
    if text == "":
        end_rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, lao_font, mono=mono)
        return f'<a:p>{ppr}{end_rpr.replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")}</a:p>'

    runs_xml = []
    for script, chunk in split_script_runs(text):
        rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, lao_font, mono=mono, script=script)
        runs_xml.append(f"<a:r>{rpr}<a:t>{escape(chunk)}</a:t></a:r>")
    end_rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, lao_font, mono=mono)
    end_rpr = end_rpr.replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")
    return f'<a:p>{ppr}{"".join(runs_xml)}{end_rpr}</a:p>'


def shadow_xml(blur_pt: float = 12.0, dist_pt: float = 3.0, alpha_pct: int = 25, angle_deg: float = 90.0) -> str:
    blur_emu = int(round(blur_pt * 12700))
    dist_emu = int(round(dist_pt * 12700))
    alpha_val = int(round(alpha_pct * 1000))
    dir_val = rot_val(angle_deg)
    return (
        "<a:effectLst>"
        f'<a:outerShdw blurRad="{blur_emu}" dist="{dist_emu}" dir="{dir_val}" algn="tl" rotWithShape="0">'
        f'<a:srgbClr val="000000"><a:alpha val="{alpha_val}"/></a:srgbClr>'
        "</a:outerShdw>"
        "</a:effectLst>"
    )


def fill_xml(
    fill_color: str | None,
    alpha_pct: int = 100,
    gradient: tuple[str, str, float] | None = None,
    radial_glow: tuple[str, int, int, int] | None = None,
) -> str:
    """Build solidFill, linear gradFill, radial circle gradFill, or noFill.

    radial_glow: (hex_color, center_alpha_pct, mid_alpha_pct, edge_alpha_pct)
    """
    if radial_glow is not None:
        col, a0, a_mid, a1 = radial_glow
        return (
            '<a:gradFill flip="none" rotWithShape="1">'
            "<a:gsLst>"
            f'<a:gs pos="0"><a:srgbClr val="{col}"><a:alpha val="{max(0, min(100, a0)) * 1000}"/></a:srgbClr></a:gs>'
            f'<a:gs pos="55000"><a:srgbClr val="{col}"><a:alpha val="{max(0, min(100, a_mid)) * 1000}"/></a:srgbClr></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{col}"><a:alpha val="{max(0, min(100, a1)) * 1000}"/></a:srgbClr></a:gs>'
            "</a:gsLst>"
            '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
            "</a:gradFill>"
        )
    if gradient is not None:
        c1, c2, ang = gradient
        ang_val = rot_val(ang)
        return (
            "<a:gradFill flip=\"none\" rotWithShape=\"1\">"
            "<a:gsLst>"
            f'<a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs>'
            "</a:gsLst>"
            f'<a:lin ang="{ang_val}" scaled="1"/>'
            "</a:gradFill>"
        )
    if fill_color is None:
        return "<a:noFill/>"
    if alpha_pct < 100:
        return (
            f'<a:solidFill><a:srgbClr val="{fill_color}">'
            f'<a:alpha val="{alpha_pct * 1000}"/></a:srgbClr></a:solidFill>'
        )
    return f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'


def scene3d_xml(
    preset: str = "perspectiveFront",
    rot_x_deg: float = 0.0,
    rot_y_deg: float = 0.0,
    rot_z_deg: float = 0.0,
    fov_deg: float = 45.0,
    light_rig: str = "threePt",
    light_dir: str = "t",
) -> str:
    """Build Desktop-compatible <a:scene3d> 3D camera + lighting rig."""
    lat = rot_val(rot_x_deg)
    lon = rot_val(rot_y_deg)
    rev = rot_val(rot_z_deg)
    fov = int(round(max(0.0, min(180.0, fov_deg)) * 60000))
    fov_attr = f' fov="{fov}"' if preset.startswith("perspective") else ""
    return (
        "<a:scene3d>"
        f'<a:camera prst="{escape(preset)}"{fov_attr}>'
        f'<a:rot lat="{lat}" lon="{lon}" rev="{rev}"/>'
        "</a:camera>"
        f'<a:lightRig rig="{escape(light_rig)}" dir="{escape(light_dir)}"/>'
        "</a:scene3d>"
    )


def sp3d_xml(
    extrude_pt: float = 0.0,
    bevel_pt: float = 0.0,
    material: str = "warmMatte",
    extrude_color: str | None = None,
) -> str:
    """Build <a:sp3d> extrusion & bevel properties (must follow <a:scene3d>)."""
    ext_emu = int(round(max(0.0, extrude_pt) * 12700))
    bev_emu = int(round(max(0.0, bevel_pt) * 12700))
    bev_xml = f'<a:bevelT w="{bev_emu}" h="{bev_emu}"/>' if bev_emu > 0 else ""
    ext_clr = f'<a:extrusionClr><a:srgbClr val="{extrude_color}"/></a:extrusionClr>' if extrude_color else ""
    ext_attr = f' extrusionH="{ext_emu}"' if ext_emu > 0 else ""
    return f'<a:sp3d{ext_attr} prstMaterial="{escape(material)}">{bev_xml}{ext_clr}</a:sp3d>'


def line_xml(line_color: str | None, line_pt: float = 1.0, dash: str | None = None, alpha_pct: int = 100) -> str:
    if line_color is None or line_pt <= 0:
        return "<a:ln><a:noFill/></a:ln>"
    w_emu = int(round(line_pt * 12700))
    if alpha_pct < 100:
        f_xml = f'<a:solidFill><a:srgbClr val="{line_color}"><a:alpha val="{alpha_pct * 1000}"/></a:srgbClr></a:solidFill>'
    else:
        f_xml = f'<a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill>'
    dash_xml = f'<a:prstDash val="{dash}"/>' if dash else '<a:prstDash val="solid"/>'
    return f'<a:ln w="{w_emu}">{f_xml}{dash_xml}</a:ln>'


def make_shape(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    prst: str = "roundRect",
    adj: int | None = None,
    rot_deg: float = 0.0,
    fill: str | None = None,
    fill_alpha: int = 100,
    gradient: tuple[str, str, float] | None = None,
    radial_glow: tuple[str, int, int, int] | None = None,
    line: str | None = None,
    line_pt: float = 1.0,
    line_dash: str | None = None,
    line_alpha: int = 100,
    shadow: bool = False,
    paragraphs: list[str] | None = None,
    pad_in: tuple[float, float, float, float] = (0.14, 0.10, 0.14, 0.10),
    anchor: str = "t",
    scene3d: str = "",
    sp3d: str = "",
) -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    rot_attr = f' rot="{rot_val(rot_deg)}"' if abs(rot_deg) > 1e-6 else ""
    av_lst = f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>' if adj is not None else "<a:avLst/>"
    eff = shadow_xml() if shadow else ""
    l_ins, t_ins, r_ins, b_ins = [emu(p) for p in pad_in]
    p_xml = "".join(paragraphs) if paragraphs else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return (
        "<p:sp>"
        "<p:nvSpPr>"
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        "<p:cNvSpPr/>"
        "<p:nvPr/>"
        "</p:nvSpPr>"
        "<p:spPr>"
        f'<a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}">{av_lst}</a:prstGeom>'
        f"{fill_xml(fill, fill_alpha, gradient, radial_glow)}"
        f"{line_xml(line, line_pt, line_dash, line_alpha)}"
        f"{eff}"
        f"{scene3d}"
        f"{sp3d}"
        "</p:spPr>"
        "<p:txBody>"
        f'<a:bodyPr wrap="square" lIns="{l_ins}" tIns="{t_ins}" rIns="{r_ins}" bIns="{b_ins}" anchor="{anchor}">'
        "<a:noAutofit/>"
        "</a:bodyPr>"
        "<a:lstStyle/>"
        f"{p_xml}"
        "</p:txBody>"
        "</p:sp>"
    )


def make_radial_glow(
    sp_id: int,
    name: str,
    cx_in: float,
    cy_in: float,
    rx_in: float,
    ry_in: float | None = None,
    color: str = "38DCFC",
    center_alpha: int = 45,
    mid_alpha: int = 16,
    edge_alpha: int = 0,
) -> str:
    """Create a native editable radial ambient glow ellipse (zero bitmap blur)."""
    ry = rx_in if ry_in is None else ry_in
    return make_shape(
        sp_id,
        name,
        cx_in - rx_in,
        cy_in - ry,
        2.0 * rx_in,
        2.0 * ry,
        prst="ellipse",
        fill=None,
        radial_glow=(color, center_alpha, mid_alpha, edge_alpha),
        line=None,
    )


def make_custgeom_shape(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    points_norm: list[tuple[int, int]],
    coord_w: int,
    coord_h: int,
    rot_deg: float = 0.0,
    fill: str | None = None,
    fill_alpha: int = 100,
    gradient: tuple[str, str, float] | None = None,
    line: str | None = None,
    line_pt: float = 1.0,
    line_dash: str | None = None,
    line_alpha: int = 100,
    shadow: bool = False,
    paragraphs: list[str] | None = None,
    pad_in: tuple[float, float, float, float] = (0.14, 0.10, 0.14, 0.10),
    anchor: str = "ctr",
    scene3d: str = "",
    sp3d: str = "",
) -> str:
    """Build a native OOXML Freeform (<a:custGeom>) polygon."""
    segs = [f'<a:moveTo><a:pt x="{points_norm[0][0]}" y="{points_norm[0][1]}"/></a:moveTo>']
    for px, py in points_norm[1:]:
        segs.append(f'<a:lnTo><a:pt x="{px}" y="{py}"/></a:lnTo>')
    segs.append("<a:close/>")
    return _custgeom_from_path_xml(
        sp_id, name, x_in, y_in, w_in, h_in, f'<a:path w="{coord_w}" h="{coord_h}">{"".join(segs)}</a:path>',
        rot_deg, fill, fill_alpha, gradient, line, line_pt, shadow,
        paragraphs=paragraphs, pad_in=pad_in, anchor=anchor,
        line_dash=line_dash, line_alpha=line_alpha, scene3d=scene3d, sp3d=sp3d,
    )


def _custgeom_from_path_xml(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    path_xml: str,
    rot_deg: float = 0.0,
    fill: str | None = None,
    fill_alpha: int = 100,
    gradient: tuple[str, str, float] | None = None,
    line: str | None = None,
    line_pt: float = 1.0,
    shadow: bool = False,
    paragraphs: list[str] | None = None,
    pad_in: tuple[float, float, float, float] = (0.14, 0.10, 0.14, 0.10),
    anchor: str = "ctr",
    line_dash: str | None = None,
    line_alpha: int = 100,
    radial_glow: tuple[str, int, int, int] | None = None,
    scene3d: str = "",
    sp3d: str = "",
) -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    rot_attr = f' rot="{rot_val(rot_deg)}"' if abs(rot_deg) > 1e-6 else ""
    eff = shadow_xml() if shadow else ""
    l_ins, t_ins, r_ins, b_ins = [emu(p) for p in pad_in]
    p_xml = "".join(paragraphs) if paragraphs else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return (
        "<p:sp>"
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        "<p:spPr>"
        f'<a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        "<a:custGeom>"
        "<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>"
        '<a:rect l="0" t="0" r="r" b="b"/>'
        f"<a:pathLst>{path_xml}</a:pathLst>"
        "</a:custGeom>"
        f"{fill_xml(fill, fill_alpha, gradient, radial_glow)}"
        f"{line_xml(line, line_pt, line_dash, line_alpha)}"
        f"{eff}"
        f"{scene3d}"
        f"{sp3d}"
        "</p:spPr>"
        f'<p:txBody><a:bodyPr wrap="square" lIns="{l_ins}" tIns="{t_ins}" rIns="{r_ins}" bIns="{b_ins}" anchor="{anchor}"><a:noAutofit/></a:bodyPr>'
        f"<a:lstStyle/>{p_xml}</p:txBody>"
        "</p:sp>"
    )


def _arc_to_cubics(x0, y0, rx, ry, phi_deg, large_arc, sweep, x1, y1):
    """Convert SVG elliptical arc to cubic Bezier segments [(c1x,c1y,c2x,c2y,ex,ey), ...]."""
    if (x0 == x1 and y0 == y1) or rx == 0 or ry == 0:
        return []
    rx, ry = abs(rx), abs(ry)
    phi = math.radians(phi_deg % 360.0)
    cp, sp = math.cos(phi), math.sin(phi)
    dx2, dy2 = (x0 - x1) / 2.0, (y0 - y1) / 2.0
    x1p = cp * dx2 + sp * dy2
    y1p = -sp * dx2 + cp * dy2
    lam = (x1p ** 2) / (rx ** 2) + (y1p ** 2) / (ry ** 2)
    if lam > 1.0:
        s = math.sqrt(lam)
        rx *= s
        ry *= s
    num = max(0.0, (rx * ry) ** 2 - (rx * y1p) ** 2 - (ry * x1p) ** 2)
    den = (rx * y1p) ** 2 + (ry * x1p) ** 2
    coef = math.sqrt(num / den) if den > 1e-12 else 0.0
    if large_arc == sweep:
        coef = -coef
    cxp = coef * (rx * y1p) / ry
    cyp = coef * -(ry * x1p) / rx
    cx = cp * cxp - sp * cyp + (x0 + x1) / 2.0
    cy = sp * cxp + cp * cyp + (y0 + y1) / 2.0

    def vec_ang(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        mag = math.hypot(ux, uy) * math.hypot(vx, vy)
        ang = math.acos(max(-1.0, min(1.0, dot / mag))) if mag > 1e-12 else 0.0
        return -ang if (ux * vy - uy * vx) < 0 else ang

    t1 = vec_ang(1.0, 0.0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = vec_ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2.0 * math.pi
    elif sweep and dt < 0:
        dt += 2.0 * math.pi

    segs = max(1, int(math.ceil(abs(dt) / (math.pi / 2.0))))
    delta = dt / segs
    alpha = (4.0 / 3.0) * math.tan(delta / 4.0)
    cubics = []
    for i in range(segs):
        a1 = t1 + i * delta
        a2 = a1 + delta
        c1, s1 = math.cos(a1), math.sin(a1)
        c2, s2 = math.cos(a2), math.sin(a2)
        p1x, p1y = rx * (c1 - alpha * s1), ry * (s1 + alpha * c1)
        p2x, p2y = rx * (c2 + alpha * s2), ry * (s2 - alpha * c2)
        epx, epy = rx * c2, ry * s2
        cubics.append((
            cx + cp * p1x - sp * p1y, cy + sp * p1x + cp * p1y,
            cx + cp * p2x - sp * p2y, cy + sp * p2x + cp * p2y,
            cx + cp * epx - sp * epy, cy + sp * epx + cp * epy,
        ))
    return cubics


def svg_path_to_ooxml_path(
    d: str,
    viewbox: tuple[float, float, float, float] | None = None,
    coord_scale: int = 10000,
) -> tuple[str, int, int]:
    """Convert an SVG path `d` string into an OOXML <a:path> preserving curves."""
    tokens = _TOKEN_RE.findall(d)
    ops: list[tuple] = []
    idx = 0
    cmd = None
    cx, cy = 0.0, 0.0
    sx, sy = 0.0, 0.0
    last_ctrl: tuple[float, float] | None = None
    prev_cmd = None

    def next_num() -> float:
        nonlocal idx
        val = float(tokens[idx])
        idx += 1
        return val

    while idx < len(tokens):
        t = tokens[idx]
        if t.isalpha():
            cmd = t
            idx += 1
            if cmd in ("Z", "z"):
                ops.append(("close",))
                cx, cy = sx, sy
                last_ctrl = None
                prev_cmd = cmd
                continue
        elif cmd is None:
            raise ValueError(f"Unexpected numeric token before command: {t}")

        rel = cmd.islower()
        ucmd = cmd.upper()

        if ucmd == "M":
            nx, ny = next_num(), next_num()
            if rel:
                nx, ny = cx + nx, cy + ny
            ops.append(("move", nx, ny))
            cx, cy = nx, ny
            sx, sy = nx, ny
            cmd = "l" if rel else "L"
            last_ctrl = None
        elif ucmd == "L":
            nx, ny = next_num(), next_num()
            if rel:
                nx, ny = cx + nx, cy + ny
            ops.append(("line", nx, ny))
            cx, cy = nx, ny
            last_ctrl = None
        elif ucmd == "H":
            nx = next_num()
            if rel:
                nx += cx
            ops.append(("line", nx, cy))
            cx = nx
            last_ctrl = None
        elif ucmd == "V":
            ny = next_num()
            if rel:
                ny += cy
            ops.append(("line", cx, ny))
            cy = ny
            last_ctrl = None
        elif ucmd == "C":
            x1, y1, x2, y2, ex, ey = [next_num() for _ in range(6)]
            if rel:
                x1, y1, x2, y2, ex, ey = cx + x1, cy + y1, cx + x2, cy + y2, cx + ex, cy + ey
            ops.append(("cubic", x1, y1, x2, y2, ex, ey))
            last_ctrl = (x2, y2)
            cx, cy = ex, ey
        elif ucmd == "S":
            x2, y2, ex, ey = [next_num() for _ in range(4)]
            if rel:
                x2, y2, ex, ey = cx + x2, cy + y2, cx + ex, cy + ey
            if prev_cmd in ("C", "c", "S", "s") and last_ctrl is not None:
                x1, y1 = 2 * cx - last_ctrl[0], 2 * cy - last_ctrl[1]
            else:
                x1, y1 = cx, cy
            ops.append(("cubic", x1, y1, x2, y2, ex, ey))
            last_ctrl = (x2, y2)
            cx, cy = ex, ey
        elif ucmd == "Q":
            x1, y1, ex, ey = [next_num() for _ in range(4)]
            if rel:
                x1, y1, ex, ey = cx + x1, cy + y1, cx + ex, cy + ey
            ops.append(("quad", x1, y1, ex, ey))
            last_ctrl = (x1, y1)
            cx, cy = ex, ey
        elif ucmd == "T":
            ex, ey = next_num(), next_num()
            if rel:
                ex, ey = cx + ex, cy + ey
            if prev_cmd in ("Q", "q", "T", "t") and last_ctrl is not None:
                x1, y1 = 2 * cx - last_ctrl[0], 2 * cy - last_ctrl[1]
            else:
                x1, y1 = cx, cy
            ops.append(("quad", x1, y1, ex, ey))
            last_ctrl = (x1, y1)
            cx, cy = ex, ey
        elif ucmd == "A":
            rx, ry, phi, la, sw, ex, ey = [next_num() for _ in range(7)]
            if rel:
                ex, ey = cx + ex, cy + ey
            for c1x, c1y, c2x, c2y, px, py in _arc_to_cubics(cx, cy, rx, ry, phi, int(la) != 0, int(sw) != 0, ex, ey):
                ops.append(("cubic", c1x, c1y, c2x, c2y, px, py))
            cx, cy = ex, ey
            last_ctrl = None
        prev_cmd = cmd

    if not ops:
        raise ValueError("Empty or invalid SVG path")

    if viewbox is None:
        xs, ys = [], []
        for op in ops:
            for i in range(1, len(op), 2):
                xs.append(op[i])
                ys.append(op[i + 1])
        if not xs or not ys:
            raise ValueError("SVG path contains no coordinates")
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        vw = max(1e-6, max_x - min_x)
        vh = max(1e-6, max_y - min_y)
    else:
        min_x, min_y, vw, vh = viewbox

    cw = coord_scale
    ch = max(1, int(round(coord_scale * (vh / vw))))

    def tx(vx: float) -> int:
        return int(round(((vx - min_x) / vw) * cw))

    def ty(vy: float) -> int:
        return int(round(((vy - min_y) / vh) * ch))

    segs = []
    for op in ops:
        kind = op[0]
        if kind == "move":
            segs.append(f'<a:moveTo><a:pt x="{tx(op[1])}" y="{ty(op[2])}"/></a:moveTo>')
        elif kind == "line":
            segs.append(f'<a:lnTo><a:pt x="{tx(op[1])}" y="{ty(op[2])}"/></a:lnTo>')
        elif kind == "cubic":
            segs.append(
                "<a:cubicBezTo>"
                f'<a:pt x="{tx(op[1])}" y="{ty(op[2])}"/>'
                f'<a:pt x="{tx(op[3])}" y="{ty(op[4])}"/>'
                f'<a:pt x="{tx(op[5])}" y="{ty(op[6])}"/>'
                "</a:cubicBezTo>"
            )
        elif kind == "quad":
            segs.append(
                "<a:quadBezTo>"
                f'<a:pt x="{tx(op[1])}" y="{ty(op[2])}"/>'
                f'<a:pt x="{tx(op[3])}" y="{ty(op[4])}"/>'
                "</a:quadBezTo>"
            )
        elif kind == "close":
            segs.append("<a:close/>")

    return f'<a:path w="{cw}" h="{ch}">{"".join(segs)}</a:path>', cw, ch


def make_svg_shape(
    sp_id: int,
    name: str,
    svg_d: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    viewbox: tuple[float, float, float, float] | None = None,
    rot_deg: float = 0.0,
    fill: str | None = None,
    fill_alpha: int = 100,
    gradient: tuple[str, str, float] | None = None,
    line: str | None = None,
    line_pt: float = 1.0,
    shadow: bool = False,
    paragraphs: list[str] | None = None,
    scene3d: str = "",
    sp3d: str = "",
) -> str:
    """Create a 100% editable PowerPoint Freeform shape directly from an SVG path `d`."""
    path_xml, _, _ = svg_path_to_ooxml_path(svg_d, viewbox=viewbox)
    return _custgeom_from_path_xml(
        sp_id, name, x_in, y_in, w_in, h_in, path_xml,
        rot_deg, fill, fill_alpha, gradient, line, line_pt, shadow, paragraphs,
        scene3d=scene3d, sp3d=sp3d,
    )


def _ccw_ellipse_svg(cx: float, cy: float, rx: float, ry: float) -> str:
    """Counter-clockwise ellipse subpath in SVG coordinates (for compound holes)."""
    k = 0.5522847498
    kx, ky = rx * k, ry * k
    return (
        f"M {cx + rx:.2f} {cy:.2f} "
        f"C {cx + rx:.2f} {cy - ky:.2f} {cx + kx:.2f} {cy - ry:.2f} {cx:.2f} {cy - ry:.2f} "
        f"C {cx - kx:.2f} {cy - ry:.2f} {cx - rx:.2f} {cy - ky:.2f} {cx - rx:.2f} {cy:.2f} "
        f"C {cx - rx:.2f} {cy + ky:.2f} {cx - kx:.2f} {cy + ry:.2f} {cx:.2f} {cy + ry:.2f} "
        f"C {cx + kx:.2f} {cy + ry:.2f} {cx + rx:.2f} {cy + ky:.2f} {cx + rx:.2f} {cy:.2f} Z"
    )


def _ccw_roundrect_svg(x: float, y: float, w: float, h: float, r: float) -> str:
    """Counter-clockwise rounded-rect subpath in SVG coordinates (for compound holes)."""
    r = max(0.0, min(r, w / 2.0, h / 2.0))
    if r < 1e-3:
        return f"M {x:.2f} {y:.2f} L {x:.2f} {y + h:.2f} L {x + w:.2f} {y + h:.2f} L {x + w:.2f} {y:.2f} Z"
    k = r * 0.5522847498
    return (
        f"M {x + r:.2f} {y:.2f} "
        f"C {x + r - k:.2f} {y:.2f} {x:.2f} {y + r - k:.2f} {x:.2f} {y + r:.2f} "
        f"L {x:.2f} {y + h - r:.2f} "
        f"C {x:.2f} {y + h - r + k:.2f} {x + r - k:.2f} {y + h:.2f} {x + r:.2f} {y + h:.2f} "
        f"L {x + w - r:.2f} {y + h:.2f} "
        f"C {x + w - r + k:.2f} {y + h:.2f} {x + w:.2f} {y + h - r + k:.2f} {x + w:.2f} {y + h - r:.2f} "
        f"L {x + w:.2f} {y + r:.2f} "
        f"C {x + w:.2f} {y + r - k:.2f} {x + w - r + k:.2f} {y:.2f} {x + w - r:.2f} {y:.2f} Z"
    )


def make_portal_mask(
    sp_id: int,
    name: str,
    hole_box: tuple[float, float, float, float] = (3.5, 1.3125, 3.0, 3.0),
    hole_type: str = "circle",
    hole_svg_d: str | None = None,
    x_in: float = 0.0,
    y_in: float = 0.0,
    w_in: float = 10.0,
    h_in: float = 5.625,
    fill: str = "0F172A",
    fill_alpha: int = 100,
    gradient: tuple[str, str, float] | None = None,
    line: str | None = None,
    line_pt: float = 0.0,
    round_radius_in: float = 0.35,
    rot_deg: float = 0.0,
) -> str:
    """Create a compound-path Portal Mask (<a:custGeom>) with outer rect + inner cutout hole.

    `hole_box`: (hx, hy, hw, hh) in inches relative to the mask's top-left (x_in, y_in).
    `hole_type`: "circle" | "roundRect" | "rect" | "svg" (uses `hole_svg_d` in same 0..w_in, 0..h_in space).
    """
    outer_d = f"M 0 0 L {w_in:.4f} 0 L {w_in:.4f} {h_in:.4f} L 0 {h_in:.4f} Z"
    hx, hy, hw, hh = hole_box
    if hole_type == "svg" and hole_svg_d:
        inner_d = hole_svg_d.strip()
    elif hole_type == "rect":
        inner_d = _ccw_roundrect_svg(hx, hy, hw, hh, 0.0)
    elif hole_type == "roundRect":
        inner_d = _ccw_roundrect_svg(hx, hy, hw, hh, round_radius_in)
    else:
        inner_d = _ccw_ellipse_svg(hx + hw / 2.0, hy + hh / 2.0, hw / 2.0, hh / 2.0)

    path_xml, _, _ = svg_path_to_ooxml_path(f"{outer_d} {inner_d}", viewbox=(0.0, 0.0, w_in, h_in))
    return _custgeom_from_path_xml(
        sp_id, name, x_in, y_in, w_in, h_in, path_xml,
        rot_deg=rot_deg, fill=fill, fill_alpha=fill_alpha, gradient=gradient,
        line=line, line_pt=line_pt,
    )


def make_video_frame(
    sp_id: int,
    name: str,
    poster_rel_id: str,
    video_rel_id: str,
    x_in: float = 0.0,
    y_in: float = 0.0,
    w_in: float = 10.0,
    h_in: float = 5.625,
) -> str:
    """Build an embedded video <p:pic> frame (e.g., !!closing_video) with poster image."""
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    return (
        "<p:pic>"
        "<p:nvPicPr>"
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr><a:videoFile r:link="{video_rel_id}"/></p:nvPr>'
        "</p:nvPicPr>"
        "<p:blipFill>"
        f'<a:blip r:embed="{poster_rel_id}"/>'
        "<a:stretch><a:fillRect/></a:stretch>"
        "</p:blipFill>"
        "<p:spPr>"
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</p:spPr>"
        "</p:pic>"
    )


def build_curtain_pair(
    seam_x_in: float = 5.0,
    offsets_in: list[float] | None = None,
    slide_w: float = 10.0,
    slide_h: float = 5.625,
) -> tuple[dict, dict]:
    """Generate left & right seamless Freeform curtain specs sharing exact central seam."""
    if offsets_in is None:
        offsets_in = [0.0, 0.26, -0.22, 0.28, -0.24, 0.26, -0.22, 0.24, 0.0]
    n = len(offsets_in)
    ys = [slide_h * (i / (n - 1)) for i in range(n)]
    seam = [(seam_x_in + dx, y) for dx, y in zip(offsets_in, ys)]

    max_sx = max(sx for sx, _ in seam)
    min_sx = min(sx for sx, _ in seam)
    left_w = max_sx
    right_origin = min_sx
    right_w = slide_w - right_origin

    left_pts = [(0, 0)] + [(int(round(sx * 1000)), int(round(sy * 1000))) for sx, sy in seam] + [(0, int(round(slide_h * 1000)))]
    right_pts = (
        [(int(round((seam[0][0] - right_origin) * 1000)), 0), (int(round(right_w * 1000)), 0), (int(round(right_w * 1000)), int(round(slide_h * 1000)))]
        + [(int(round((sx - right_origin) * 1000)), int(round(sy * 1000))) for sx, sy in reversed(seam)]
    )
    return (
        {"x": 0.0, "w": left_w, "h": slide_h, "pts": left_pts, "cw": int(round(left_w * 1000)), "ch": int(round(slide_h * 1000))},
        {"x": right_origin, "w": right_w, "h": slide_h, "pts": right_pts, "cw": int(round(right_w * 1000)), "ch": int(round(slide_h * 1000))},
    )


def make_picture(
    sp_id: int,
    name: str,
    rel_id: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    src_rect: tuple[int, int, int, int] = (0, 0, 0, 0),
    rot_deg: float = 0.0,
    line: str | None = None,
    line_pt: float = 1.5,
    shadow: bool = False,
    round_adj: int = 4000,
) -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    rot_attr = f' rot="{rot_val(rot_deg)}"' if abs(rot_deg) > 1e-6 else ""
    l, t, r, b = src_rect
    src_xml = f'<a:srcRect l="{l}" t="{t}" r="{r}" b="{b}"/>' if any(src_rect) else "<a:srcRect/>"
    eff = shadow_xml() if shadow else ""
    return (
        "<p:pic>"
        "<p:nvPicPr>"
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        "<p:nvPr/>"
        "</p:nvPicPr>"
        "<p:blipFill>"
        f'<a:blip r:embed="{rel_id}"/>'
        f"{src_xml}"
        "<a:stretch><a:fillRect/></a:stretch>"
        "</p:blipFill>"
        "<p:spPr>"
        f'<a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {round_adj}"/></a:avLst></a:prstGeom>'
        f"{line_xml(line, line_pt)}"
        f"{eff}"
        "</p:spPr>"
        "</p:pic>"
    )


def make_chart_frame(sp_id: int, name: str, rel_id: str, x_in: float, y_in: float, w_in: float, h_in: float) -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    return (
        "<p:graphicFrame>"
        "<p:nvGraphicFramePr>"
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        "<p:cNvGraphicFramePr/>"
        "<p:nvPr/>"
        "</p:nvGraphicFramePr>"
        f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>'
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rel_id}"/>'
        "</a:graphicData>"
        "</a:graphic>"
        "</p:graphicFrame>"
    )


def build_minimal_xlsx(categories: list[str], series: list[tuple[str, list[float]]]) -> bytes:
    """Create a valid embedded .xlsx workbook backing a native editable PowerPoint chart."""
    def col_letter(idx: int) -> str:
        return chr(ord("A") + idx)

    rows_xml = []
    header_cells = ['<c r="A1" t="inlineStr"><is><t>Category</t></is></c>']
    for s_idx, (s_name, _) in enumerate(series, start=1):
        header_cells.append(f'<c r="{col_letter(s_idx)}1" t="inlineStr"><is><t>{escape(s_name)}</t></is></c>')
    rows_xml.append(f'<row r="1">{"".join(header_cells)}</row>')

    for r_idx, cat in enumerate(categories, start=2):
        cells = [f'<c r="A{r_idx}" t="inlineStr"><is><t>{escape(cat)}</t></is></c>']
        for s_idx, (_, vals) in enumerate(series, start=1):
            cells.append(f'<c r="{col_letter(s_idx)}{r_idx}"><v>{vals[r_idx - 2]}</v></c>')
        rows_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        "</worksheet>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        "</styleSheet>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1)
        zf.writestr("xl/styles.xml", styles)
    return buf.getvalue()


def _chart_rpr_xml(
    size_pt: float,
    color: str,
    bold: bool = False,
    latin_font: str = "Arial",
    thai_font: str = "Noto Sans Thai",
    cjk_font: str = "Noto Sans CJK SC",
) -> str:
    sz = int(round(size_pt * 100))
    b_attr = ' b="1"' if bold else ' b="0"'
    return (
        f'<a:defRPr sz="{sz}"{b_attr}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{escape(latin_font)}"/>'
        f'<a:ea typeface="{escape(cjk_font)}"/>'
        f'<a:cs typeface="{escape(thai_font)}"/>'
        "</a:defRPr>"
    )


def _chart_title_xml(title: str, text_color: str) -> str:
    return (
        '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
        f'<a:p><a:pPr>{_chart_rpr_xml(13.0, text_color, bold=True)}</a:pPr>'
        f'<a:r><a:t>{escape(title)}</a:t></a:r></a:p>'
        '</c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:autoTitleDeleted val="0"/>'
    )


def _chart_legend_xml(text_color: str, pos: str = "b") -> str:
    return (
        f'<c:legend><c:legendPos val="{pos}"/><c:overlay val="0"/>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(10.5, text_color)}</a:pPr></a:p></c:txPr>'
        "</c:legend>"
    )


def _chart_axes_xml(
    text_color: str,
    muted_color: str,
    grid_color: str,
    fmt_code: str,
    max_val: float | None,
    horizontal: bool = False,
) -> str:
    cat_pos = "l" if horizontal else "b"
    val_pos = "b" if horizontal else "l"
    max_xml = f'<c:max val="{max_val}"/>' if max_val is not None else ""
    return (
        "<c:catAx>"
        '<c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="0"/><c:axPos val="{cat_pos}"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="{grid_color}"/></a:solidFill></a:ln></c:spPr>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(10.5, text_color, bold=True)}</a:pPr></a:p></c:txPr>'
        '<c:crossAx val="1002"/><c:crosses val="autoZero"/>'
        "</c:catAx>"
        "<c:valAx>"
        f'<c:axId val="1002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/>{max_xml}</c:scaling>'
        f'<c:delete val="0"/><c:axPos val="{val_pos}"/>'
        f'<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="{grid_color}"><a:alpha val="55000"/></a:srgbClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
        f'<c:numFmt formatCode="{fmt_code}" sourceLinked="0"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(9.5, muted_color)}</a:pPr></a:p></c:txPr>'
        '<c:crossAx val="1001"/><c:crosses val="autoZero"/>'
        "</c:valAx>"
    )


def _wrap_chart_space(title: str, plot_inner_xml: str, text_color: str, xlsx_rel_id: str, legend_pos: str = "b") -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
        "<c:chart>"
        f"{_chart_title_xml(title, text_color)}"
        f"<c:plotArea><c:layout/>{plot_inner_xml}</c:plotArea>"
        f"{_chart_legend_xml(text_color, pos=legend_pos)}"
        '<c:plotVisOnly val="1"/>'
        "</c:chart>"
        f'<c:externalData r:id="{xlsx_rel_id}"><c:autoUpdate val="0"/></c:externalData>'
        "</c:chartSpace>"
    )


def _build_bar_series_xml(
    categories: list[str],
    series: list[tuple[str, str, list[float]]],
    text_color: str,
    fmt_code: str,
    dlbl_pos: str = "outEnd",
    show_dlbls: bool = True,
) -> str:
    n_cats = len(categories)
    cat_pts = "".join(f'<c:pt idx="{i}"><c:v>{escape(c)}</c:v></c:pt>' for i, c in enumerate(categories))
    ser_xml_list = []
    for idx, (s_name, s_color, vals) in enumerate(series):
        col_letter = chr(ord("B") + idx)
        val_pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
        show_v = "1" if show_dlbls else "0"
        ser_xml_list.append(
            "<c:ser>"
            f'<c:idx val="{idx}"/><c:order val="{idx}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col_letter}$1</c:f>'
            f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape(s_name)}</c:v></c:pt></c:strCache>'
            "</c:strRef></c:tx>"
            f'<c:spPr><a:solidFill><a:srgbClr val="{s_color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
            '<c:invertIfNegative val="0"/>'
            "<c:dLbls>"
            f'<c:numFmt formatCode="{fmt_code}" sourceLinked="0"/>'
            f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(10.0, text_color, bold=True)}</a:pPr></a:p></c:txPr>'
            f'<c:dLblPos val="{dlbl_pos}"/>'
            f'<c:showLegendKey val="0"/><c:showVal val="{show_v}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            "</c:dLbls>"
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n_cats + 1}</c:f>'
            f'<c:strCache><c:ptCount val="{n_cats}"/>{cat_pts}</c:strCache>'
            "</c:strRef></c:cat>"
            f'<c:val><c:numRef><c:f>Sheet1!${col_letter}$2:${col_letter}${n_cats + 1}</c:f>'
            f'<c:numCache><c:formatCode>{fmt_code}</c:formatCode><c:ptCount val="{n_cats}"/>{val_pts}</c:numCache>'
            "</c:numRef></c:val>"
            "</c:ser>"
        )
    return "".join(ser_xml_list)


def build_clustered_bar_chart_xml(
    title: str,
    categories: list[str],
    series: list[tuple[str, str, list[float]]],
    xlsx_rel_id: str = "rId1",
    text_color: str = "F8FAFC",
    muted_color: str = "94A3B8",
    grid_color: str = "334155",
    unit_suffix: str = "%",
    max_val: float | None = 110.0,
) -> str:
    fmt_code = f'0&quot;{escape(unit_suffix)}&quot;' if unit_suffix else "0"
    ser_xml = _build_bar_series_xml(categories, series, text_color, fmt_code, dlbl_pos="outEnd")
    plot_inner = (
        "<c:barChart>"
        '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        f"{ser_xml}"
        '<c:gapWidth val="120"/><c:overlap val="-15"/>'
        '<c:axId val="1001"/><c:axId val="1002"/>'
        "</c:barChart>"
        f"{_chart_axes_xml(text_color, muted_color, grid_color, fmt_code, max_val)}"
    )
    return _wrap_chart_space(title, plot_inner, text_color, xlsx_rel_id)


def build_stacked_bar_chart_xml(
    title: str,
    categories: list[str],
    series: list[tuple[str, str, list[float]]],
    xlsx_rel_id: str = "rId1",
    text_color: str = "F8FAFC",
    muted_color: str = "94A3B8",
    grid_color: str = "334155",
    unit_suffix: str = "%",
    max_val: float | None = None,
    grouping: str = "stacked",
    bar_dir: str = "col",
) -> str:
    """Build native editable stacked / 100%-stacked column or bar chart."""
    fmt_code = f'0&quot;{escape(unit_suffix)}&quot;' if unit_suffix else "0"
    ser_xml = _build_bar_series_xml(categories, series, text_color, fmt_code, dlbl_pos="ctr")
    plot_inner = (
        "<c:barChart>"
        f'<c:barDir val="{escape(bar_dir)}"/><c:grouping val="{escape(grouping)}"/><c:varyColors val="0"/>'
        f"{ser_xml}"
        '<c:gapWidth val="120"/><c:overlap val="100"/>'
        '<c:axId val="1001"/><c:axId val="1002"/>'
        "</c:barChart>"
        f"{_chart_axes_xml(text_color, muted_color, grid_color, fmt_code, max_val, horizontal=(bar_dir == 'bar'))}"
    )
    return _wrap_chart_space(title, plot_inner, text_color, xlsx_rel_id)


def build_line_chart_xml(
    title: str,
    categories: list[str],
    series: list[tuple[str, str, list[float]]],
    xlsx_rel_id: str = "rId1",
    text_color: str = "F8FAFC",
    muted_color: str = "94A3B8",
    grid_color: str = "334155",
    unit_suffix: str = "",
    max_val: float | None = None,
    smooth: bool = True,
    show_markers: bool = True,
    line_pt: float = 2.5,
) -> str:
    """Build native editable Line Chart (<c:lineChart>) backed by embedded .xlsx."""
    n_cats = len(categories)
    cat_pts = "".join(f'<c:pt idx="{i}"><c:v>{escape(c)}</c:v></c:pt>' for i, c in enumerate(categories))
    fmt_code = f'0&quot;{escape(unit_suffix)}&quot;' if unit_suffix else "0"
    line_w = int(round(line_pt * 12700))
    smooth_val = "1" if smooth else "0"
    marker_xml = '<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>' if show_markers else '<c:marker><c:symbol val="none"/></c:marker>'

    ser_xml_list = []
    for idx, (s_name, s_color, vals) in enumerate(series):
        col_letter = chr(ord("B") + idx)
        val_pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
        ser_xml_list.append(
            "<c:ser>"
            f'<c:idx val="{idx}"/><c:order val="{idx}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col_letter}$1</c:f>'
            f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape(s_name)}</c:v></c:pt></c:strCache>'
            "</c:strRef></c:tx>"
            f'<c:spPr><a:ln w="{line_w}" cap="rnd"><a:solidFill><a:srgbClr val="{s_color}"/></a:solidFill><a:round/></a:ln></c:spPr>'
            f"{marker_xml}"
            "<c:dLbls>"
            f'<c:numFmt formatCode="{fmt_code}" sourceLinked="0"/>'
            f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(9.5, text_color, bold=True)}</a:pPr></a:p></c:txPr>'
            '<c:dLblPos val="t"/>'
            '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            "</c:dLbls>"
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n_cats + 1}</c:f>'
            f'<c:strCache><c:ptCount val="{n_cats}"/>{cat_pts}</c:strCache>'
            "</c:strRef></c:cat>"
            f'<c:val><c:numRef><c:f>Sheet1!${col_letter}$2:${col_letter}${n_cats + 1}</c:f>'
            f'<c:numCache><c:formatCode>{fmt_code}</c:formatCode><c:ptCount val="{n_cats}"/>{val_pts}</c:numCache>'
            "</c:numRef></c:val>"
            f'<c:smooth val="{smooth_val}"/>'
            "</c:ser>"
        )

    plot_inner = (
        "<c:lineChart>"
        '<c:grouping val="standard"/><c:varyColors val="0"/>'
        f'{"".join(ser_xml_list)}'
        f'<c:marker val="{"1" if show_markers else "0"}"/>'
        f'<c:smooth val="{smooth_val}"/>'
        '<c:axId val="1001"/><c:axId val="1002"/>'
        "</c:lineChart>"
        f"{_chart_axes_xml(text_color, muted_color, grid_color, fmt_code, max_val)}"
    )
    return _wrap_chart_space(title, plot_inner, text_color, xlsx_rel_id)


def build_doughnut_chart_xml(
    title: str,
    categories: list[str],
    series_name: str,
    values: list[float],
    slice_colors: list[str],
    xlsx_rel_id: str = "rId1",
    text_color: str = "F8FAFC",
    hole_size: int = 58,
    border_color: str = "0F172A",
    unit_suffix: str = "%",
) -> str:
    """Build native editable Doughnut Chart (<c:doughnutChart>) backed by embedded .xlsx."""
    n_cats = len(categories)
    cat_pts = "".join(f'<c:pt idx="{i}"><c:v>{escape(c)}</c:v></c:pt>' for i, c in enumerate(categories))
    val_pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values))
    fmt_code = f'0&quot;{escape(unit_suffix)}&quot;' if unit_suffix else "0"
    dpt_xml = "".join(
        f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/>'
        f'<c:spPr><a:solidFill><a:srgbClr val="{slice_colors[i % len(slice_colors)]}"/></a:solidFill>'
        f'<a:ln w="19050"><a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></a:ln></c:spPr>'
        "</c:dPt>"
        for i in range(n_cats)
    )
    plot_inner = (
        "<c:doughnutChart>"
        '<c:varyColors val="1"/>'
        "<c:ser>"
        '<c:idx val="0"/><c:order val="0"/>'
        '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>'
        f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape(series_name)}</c:v></c:pt></c:strCache>'
        "</c:strRef></c:tx>"
        f"{dpt_xml}"
        "<c:dLbls>"
        f'<c:numFmt formatCode="{fmt_code}" sourceLinked="0"/>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_chart_rpr_xml(10.0, text_color, bold=True)}</a:pPr></a:p></c:txPr>'
        '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
        "</c:dLbls>"
        f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n_cats + 1}</c:f>'
        f'<c:strCache><c:ptCount val="{n_cats}"/>{cat_pts}</c:strCache>'
        "</c:strRef></c:cat>"
        f'<c:val><c:numRef><c:f>Sheet1!$B$2:$B${n_cats + 1}</c:f>'
        f'<c:numCache><c:formatCode>{fmt_code}</c:formatCode><c:ptCount val="{n_cats}"/>{val_pts}</c:numCache>'
        "</c:numRef></c:val>"
        "</c:ser>"
        '<c:firstSliceAng val="0"/>'
        f'<c:holeSize val="{max(10, min(90, hole_size))}"/>'
        "</c:doughnutChart>"
    )
    return _wrap_chart_space(title, plot_inner, text_color, xlsx_rel_id, legend_pos="r")


_SCENE3D_RE = re.compile(r"<a:scene3d>.*?</a:scene3d>|<a:sp3d[^>]*>.*?</a:sp3d>|<a:sp3d[^/]*/>", re.DOTALL)


def _anim_filter(effect: str, direction: str) -> str:
    if effect == "wipe":
        dir_map = {"bottom": "down", "top": "up", "left": "left", "right": "right"}
        return f"wipe({dir_map.get(direction, 'down')})"
    if effect == "fly":
        dir_map = {"bottom": "fromBottom", "top": "fromTop", "left": "fromLeft", "right": "fromRight"}
        return f"fly({dir_map.get(direction, 'fromBottom')})"
    return "fade"


def build_timing_xml(animations: list[dict]) -> str:
    """Build PowerPoint <p:timing> mainSeq XML for post-Morph stagger animations.

    Each entry: {"sp_id": int, "effect": "fade|wipe|fly", "delay_ms": int, "dur_ms": int,
                 "trigger": "afterPrevious|withPrevious|onClick", "direction": "bottom|top|left|right"}
    """
    if not animations:
        return ""

    cid = 2

    def next_cid() -> int:
        nonlocal cid
        v = cid
        cid += 1
        return v

    # Group animations into click-batches (a batch starts on "onClick", or first entry is auto-after-Morph)
    batches: list[list[dict]] = []
    for idx, anim in enumerate(animations):
        if idx == 0 or anim.get("trigger") == "onClick":
            batches.append([anim])
        else:
            batches[-1].append(anim)

    click_pars: list[str] = []
    bld_items: list[str] = []
    seen_bld: set[int] = set()

    for b_idx, batch in enumerate(batches):
        batch_id = next_cid()
        inner_pars: list[str] = []
        for a_idx, a in enumerate(batch):
            sp_id = int(a["sp_id"])
            eff = a.get("effect", "fade")
            direction = a.get("direction", "bottom")
            delay_ms = max(0, int(a.get("delay_ms", 0)))
            dur_ms = max(50, int(a.get("dur_ms", 400)))
            trig = a.get("trigger", "withPrevious" if a_idx > 0 else "afterPrevious")
            if b_idx == 0 and a_idx == 0 and trig != "onClick":
                node_type = "afterEffect"
                st_cond = f'<p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst>'
            elif a_idx == 0:
                node_type = "clickEffect"
                st_cond = f'<p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst>'
            else:
                node_type = "afterEffect" if trig == "afterPrevious" else "withEffect"
                st_cond = f'<p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst>'

            p_id = next_cid()
            set_id = next_cid()
            eff_id = next_cid()
            flt = _anim_filter(eff, direction)

            inner_pars.append(
                f'<p:par><p:cTn id="{p_id}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="{node_type}">'
                f"{st_cond}"
                "<p:childTnLst>"
                f'<p:set><p:cBhvr><p:cTn id="{set_id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="{escape(flt)}"><p:cBhvr>'
                f'<p:cTn id="{eff_id}" dur="{dur_ms}"/>'
                f'<p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl>'
                "</p:cBhvr></p:animEffect>"
                "</p:childTnLst></p:cTn></p:par>"
            )
            if sp_id not in seen_bld:
                seen_bld.add(sp_id)
                bld_items.append(f'<p:bldP spid="{sp_id}" grpId="0" animBg="1"/>')

        grp_id = next_cid()
        is_auto = (b_idx == 0 and batch[0].get("trigger", "afterPrevious") != "onClick")
        grp_st = '<p:stCondLst><p:cond delay="0"/></p:stCondLst>' if is_auto else '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
        click_pars.append(
            f'<p:par><p:cTn id="{batch_id}" fill="hold">{grp_st}'
            "<p:childTnLst>"
            f'<p:par><p:cTn id="{grp_id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(inner_pars)}</p:childTnLst></p:cTn></p:par>'
            "</p:childTnLst></p:cTn></p:par>"
        )

    seq_id = next_cid()
    return (
        "<p:timing>"
        "<p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" restart=\"never\" nodeType=\"tmRoot\">"
        "<p:childTnLst>"
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="{seq_id}" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{"".join(click_pars)}</p:childTnLst>'
        "</p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        "</p:seq>"
        "</p:childTnLst></p:cTn></p:par></p:tnLst>"
        f'<p:bldLst>{"".join(bld_items)}</p:bldLst>'
        "</p:timing>"
    )


class SlideBuilder:
    def __init__(
        self,
        bg_color: str = "0F172A",
        morph: bool = True,
        morph_option: str = "byObject",
        speed: str = "med",
        target: str = "desktop",
    ):
        self.bg_color = bg_color
        self.morph = morph
        self.morph_option = morph_option
        self.speed = speed
        self.target = target
        self.shapes: list[str] = []
        self.animations: list[dict] = []
        self.rels: list[tuple[str, str, str]] = [
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml")
        ]
        self._next_id = 2

    def next_id(self) -> int:
        val = self._next_id
        self._next_id += 1
        return val

    def add_rel(self, rel_type: str, target: str) -> str:
        rid = f"rId{len(self.rels) + 1}"
        self.rels.append((rid, rel_type, target))
        return rid

    def add_shape_xml(self, xml_str: str):
        if self.target == "web":
            xml_str = _SCENE3D_RE.sub("", xml_str)
        self.shapes.append(xml_str)

    def add_stagger(
        self,
        sp_id: int,
        effect: str = "fade",
        delay_ms: int = 120,
        dur_ms: int = 400,
        trigger: str = "withPrevious",
        direction: str = "bottom",
    ):
        """Queue a post-Morph stagger animation targeting `sp_id`."""
        self.animations.append({
            "sp_id": sp_id,
            "effect": effect,
            "delay_ms": delay_ms,
            "dur_ms": dur_ms,
            "trigger": trigger,
            "direction": direction,
        })

    def to_slide_xml(self) -> str:
        trans = ""
        if self.morph:
            trans = (
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
                f'<p:transition spd="{self.speed}" advClick="1">'
                f'<p159:morph option="{self.morph_option}"/>'
                "</p:transition>"
                "</mc:Choice>"
                "<mc:Fallback>"
                f'<p:transition spd="{self.speed}" advClick="1"><p:fade/></p:transition>'
                "</mc:Fallback>"
                "</mc:AlternateContent>"
            )
        timing = build_timing_xml(self.animations)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            "<p:cSld>"
            f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{self.bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            "<p:spTree>"
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'{"".join(self.shapes)}'
            "</p:spTree>"
            "</p:cSld>"
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
            f"{trans}"
            f"{timing}"
            "</p:sld>"
        )

    def to_rels_xml(self) -> str:
        items = "".join(
            f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
            for rid, rtype, target in self.rels
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{items}"
            "</Relationships>"
        )


def obfuscate_odttf(ttf_bytes: bytes, guid_str: str) -> bytes:
    """ECMA-376 Part 2 font obfuscation: XOR first 32 bytes with reversed-field GUID."""
    clean = guid_str.strip("{}").replace("-", "")
    if len(clean) != 32:
        raise ValueError(f"Invalid GUID for font obfuscation: {guid_str}")
    # First 3 UUID fields (4, 2, 2 bytes) are little-endian in MS GUID byte order
    raw = bytes.fromhex(clean)
    key = raw[3::-1] + raw[5:3:-1] + raw[7:5:-1] + raw[8:]
    buf = bytearray(ttf_bytes)
    for i in range(min(32, len(buf))):
        buf[i] ^= key[i % 16]
    return bytes(buf)


def write_pptx(
    out_path: str,
    slides: list[SlideBuilder],
    media_files: dict[str, bytes] | None = None,
    charts: dict[str, tuple[str, bytes]] | None = None,
    embed_fonts: list[dict] | None = None,
    target: str = "desktop",
):
    media_files = media_files or {}
    charts = charts or {}
    embed_fonts = embed_fonts or []

    sld_ids = []
    pres_rels = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>',
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>',
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>',
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    ]
    ct_overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>',
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>',
    ]

    next_pres_rid = 6
    for i in range(1, len(slides) + 1):
        rid = f"rId{next_pres_rid}"
        next_pres_rid += 1
        sld_ids.append(f'<p:sldId id="{255 + i}" r:id="{rid}"/>')
        pres_rels.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        )
        ct_overrides.append(
            f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )

    font_blobs: dict[str, bytes] = {}
    emb_font_entries: list[str] = []
    if target != "web":
        for fspec in embed_fonts:
            typeface = fspec["typeface"]
            slot_xml: list[str] = []
            for slot_key, tag_name in (("regular", "p:regular"), ("bold", "p:bold"), ("italic", "p:italic"), ("bold_italic", "p:boldItalic")):
                src = fspec.get(slot_key)
                if not src:
                    continue
                if isinstance(src, (bytes, bytearray)):
                    ttf_data = bytes(src)
                else:
                    with open(src, "rb") as fp:
                        ttf_data = fp.read()
                guid = "{" + str(uuid.uuid4()).upper() + "}"
                font_filename = f"{guid[1:-1]}.odttf"
                font_blobs[f"ppt/fonts/{font_filename}"] = obfuscate_odttf(bytes(ttf_data), guid)
                rid = f"rId{next_pres_rid}"
                next_pres_rid += 1
                pres_rels.append(
                    f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/{font_filename}"/>'
                )
                slot_xml.append(f'<{tag_name} r:id="{rid}" fontKey="{guid}"/>')
            if slot_xml:
                emb_font_entries.append(
                    f'<p:embeddedFont><p:font typeface="{escape(typeface)}"/>{"".join(slot_xml)}</p:embeddedFont>'
                )

    embedded_font_lst_xml = f'<p:embeddedFontLst>{"".join(emb_font_entries)}</p:embeddedFontLst>' if emb_font_entries else ""

    for chart_name in charts:
        ct_overrides.append(
            f'<Override PartName="/ppt/charts/{chart_name}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        '<Default Extension="svg" ContentType="image/svg+xml"/>'
        '<Default Extension="mp4" ContentType="video/mp4"/>'
        '<Default Extension="m4a" ContentType="audio/mp4"/>'
        '<Default Extension="wav" ContentType="audio/wav"/>'
        '<Default Extension="ttf" ContentType="application/x-font-ttf"/>'
        '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        f'{"".join(ct_overrides)}'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        "</Relationships>"
    )
    emb_attr = ' embedTrueTypeFonts="1"' if emb_font_entries else ""
    presentation_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"{emb_attr}>'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{"".join(sld_ids)}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}"/>'
        f'<p:notesSz cx="{SLIDE_H}" cy="{SLIDE_W}"/>'
        f'{embedded_font_lst_xml}'
        "</p:presentation>"
    )
    pres_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(pres_rels)}'
        "</Relationships>"
    )
    slide_master = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        "<p:cSld><p:spTree>"
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        "</p:spTree></p:cSld>"
        '<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        "</p:sldMaster>"
    )
    slide_master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        "</Relationships>"
    )
    slide_layout = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
        '<p:cSld name="Blank"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
        "</p:sldLayout>"
    )
    slide_layout_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        "</Relationships>"
    )
    theme1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="PHPresent">'
        "<a:themeElements>"
        '<a:clrScheme name="PHPresent">'
        '<a:dk1><a:srgbClr val="0F172A"/></a:dk1><a:lt1><a:srgbClr val="F8FAFC"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1E293B"/></a:dk2><a:lt2><a:srgbClr val="CBD5E1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="38DCFC"/></a:accent1><a:accent2><a:srgbClr val="00A8B5"/></a:accent2>'
        '<a:accent3><a:srgbClr val="10B981"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4>'
        '<a:accent5><a:srgbClr val="60A5FA"/></a:accent5><a:accent6><a:srgbClr val="334155"/></a:accent6>'
        '<a:hlink><a:srgbClr val="38DCFC"/></a:hlink><a:folHlink><a:srgbClr val="00A8B5"/></a:folHlink>'
        "</a:clrScheme>"
        '<a:fontScheme name="StandardMultilingual">'
        '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Noto Sans Thai"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Noto Sans Thai"/></a:minorFont>'
        "</a:fontScheme>"
        '<a:fmtScheme name="PHPresent">'
        '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
        "</a:fmtScheme>"
        "</a:themeElements>"
        "</a:theme>"
    )
    pres_props = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presProps xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    view_props = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewProps xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    table_styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("ppt/presentation.xml", presentation_xml)
        zf.writestr("ppt/_rels/presentation.xml.rels", pres_rels_xml)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels)
        zf.writestr("ppt/theme/theme1.xml", theme1)
        zf.writestr("ppt/presProps.xml", pres_props)
        zf.writestr("ppt/viewProps.xml", view_props)
        zf.writestr("ppt/tableStyles.xml", table_styles)

        for idx, s in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", s.to_slide_xml())
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", s.to_rels_xml())

        for name, data in media_files.items():
            zf.writestr(f"ppt/media/{name}", data)

        for font_path, font_data in font_blobs.items():
            zf.writestr(font_path, font_data)

        for chart_idx, (chart_name, (chart_xml, xlsx_bytes)) in enumerate(charts.items(), start=1):
            xlsx_name = f"Microsoft_Excel_Worksheet{chart_idx}.xlsx"
            chart_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/{xlsx_name}"/>'
                "</Relationships>"
            )
            zf.writestr(f"ppt/charts/{chart_name}", chart_xml)
            zf.writestr(f"ppt/charts/_rels/{chart_name}.rels", chart_rels)
            zf.writestr(f"ppt/embeddings/{xlsx_name}", xlsx_bytes)
