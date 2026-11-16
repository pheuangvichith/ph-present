#!/usr/bin/env python3
"""Audit and render a .pptx deck to PNGs + contact_sheet.png.

1. Validates OOXML structure, .rels targets, and embedded .xlsx workbooks.
2. Audits ph-present rules:
   - Zero duplicate !! anchor names on the same slide (dups == 0)
   - Morph continuity across slides (flags geometry-type changes on shared !! anchors)
   - Projector font-size floor check (warns if text < 9.0 pt)
   - Post-build Text-Box Overflow check via layout_solver TTF metrics
   - <p159:morph> presence on slides 2..N
3. Renders slides via LibreOffice + pdftoppm when available, or falls back to a
   built-in Pillow OOXML renderer (using syllable-aware Thai/Lao wrapping,
   chart XML colors, custGeom Bezier curves, gradFill, rotations, and srcRect).
"""

from __future__ import annotations

import io
import math
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter

from PIL import Image, ImageChops, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import layout_solver as ls

EMU_PER_INCH = 914400.0
PX_PER_IN = 128  # 1280 x 720 for standard 10 x 5.625 in

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


def get_font(kind: str, bold: bool, pt: float):
    return ls.get_ttf_font(kind, bold, pt, dpi=float(PX_PER_IN))


def _extract_tx_specs(tx_body) -> tuple[list[ls.TextRunSpec], tuple[float, float, float, float]]:
    """Extract layout_solver TextRunSpec items and padding from <p:txBody>."""
    bpr = tx_body.find("a:bodyPr", NS)
    l_ins = int(bpr.attrib.get("lIns", 0)) / EMU_PER_INCH if bpr is not None else 0.0
    t_ins = int(bpr.attrib.get("tIns", 0)) / EMU_PER_INCH if bpr is not None else 0.0
    r_ins = int(bpr.attrib.get("rIns", 0)) / EMU_PER_INCH if bpr is not None else 0.0
    b_ins = int(bpr.attrib.get("bIns", 0)) / EMU_PER_INCH if bpr is not None else 0.0

    specs: list[ls.TextRunSpec] = []
    for p in tx_body.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        ln_spc_el = ppr.find("a:lnSpc/a:spcPct", NS) if ppr is not None else None
        ln_mult = int(ln_spc_el.attrib["val"]) / 100000.0 if ln_spc_el is not None else 1.2
        spc_aft_el = ppr.find("a:spcAft/a:spcPts", NS) if ppr is not None else None
        spc_aft_pt = int(spc_aft_el.attrib["val"]) / 100.0 if spc_aft_el is not None else 0.0

        runs = p.findall("a:r", NS)
        if not runs:
            continue
        txt = "".join((r.find("a:t", NS).text or "") for r in runs if r.find("a:t", NS) is not None)
        if not txt.strip():
            continue
        max_pt = 0.0
        any_bold = False
        any_mono = False
        for r in runs:
            rpr = r.find("a:rPr", NS)
            if rpr is not None:
                max_pt = max(max_pt, int(rpr.attrib.get("sz", 1100)) / 100.0)
                if rpr.attrib.get("b", "0") == "1":
                    any_bold = True
                lat = rpr.find("a:latin", NS)
                if lat is not None and "Mono" in lat.attrib.get("typeface", ""):
                    any_mono = True
        if max_pt <= 0.0:
            max_pt = 11.0
        specs.append(ls.TextRunSpec(
            text=txt, size_pt=max_pt, bold=any_bold, mono=any_mono,
            line_spacing=ln_mult, space_after_pt=spc_aft_pt,
        ))
    return specs, (l_ins, t_ins, r_ins, b_ins)


def audit_deck(pptx_path: str, target: str = "desktop") -> tuple[list[str], list[str]]:
    """Run structural, !! anchor, Morph continuity, timing, 3D/font target, overflow, and font audits."""
    errors: list[str] = []
    warnings: list[str] = []

    with zipfile.ZipFile(pptx_path, "r") as zf:
        names = set(zf.namelist())
        print(f"Package entries: {len(names)}")

        for name in sorted(names):
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(zf.read(name))
                except ET.ParseError as e:
                    errors.append(f"XML parse error in {name}: {e}")

        for name in sorted(names):
            if name.endswith(".rels"):
                rel_dir = posixpath.dirname(posixpath.dirname(name)) if "_rels/" in name else ""
                if name == "_rels/.rels":
                    rel_dir = ""
                root = ET.fromstring(zf.read(name))
                for rel in root:
                    rel_target = rel.attrib.get("Target", "")
                    if rel_target.startswith("http"):
                        continue
                    resolved = posixpath.normpath(posixpath.join(rel_dir, rel_target))
                    if resolved not in names:
                        errors.append(f"Broken rel in {name}: {rel_target} -> {resolved}")

        for xe in sorted(n for n in names if n.endswith(".xlsx")):
            try:
                with zipfile.ZipFile(io.BytesIO(zf.read(xe)), "r") as xzf:
                    for xn in xzf.namelist():
                        if xn.endswith(".xml") or xn.endswith(".rels"):
                            ET.fromstring(xzf.read(xn))
                print(f"Embedded XLSX OK: {xe}")
            except Exception as e:
                errors.append(f"Invalid embedded XLSX {xe}: {e}")

        if "ppt/presentation.xml" in names:
            pres_root = ET.fromstring(zf.read("ppt/presentation.xml"))
            emb_fonts = pres_root.findall(".//p:embeddedFontLst/p:embeddedFont", NS)
            if emb_fonts:
                print(f"Embedded fonts declared: {len(emb_fonts)}")
                if target == "web":
                    warnings.append("Presentation has <p:embeddedFontLst> which PowerPoint Web may substitute")

        slide_files = sorted(
            [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda s: int(s.split("slide")[-1].split(".")[0]),
        )
        print(f"Slides found: {len(slide_files)}")

        slide_anchors: list[dict[str, str]] = []
        for idx, sf in enumerate(slide_files, start=1):
            root = ET.fromstring(zf.read(sf))
            sp_tree = root.find(".//p:spTree", NS)
            bang_list: list[str] = []
            anchor_types: dict[str, str] = {}
            shape_ids: set[str] = set()
            overflow_count = 0

            if sp_tree is not None:
                for elem in _iter_tree_elements(sp_tree):
                    tag = elem.tag.split("}")[-1]
                    cnv = elem.find(".//p:cNvPr", NS)
                    name = cnv.attrib.get("name", "") if cnv is not None else ""
                    if cnv is not None and "id" in cnv.attrib:
                        shape_ids.add(cnv.attrib["id"])
                    if name.startswith("!!"):
                        bang_list.append(name)
                        if tag == "sp":
                            prst = elem.find(".//a:prstGeom", NS)
                            geom = f"prst:{prst.attrib['prst']}" if prst is not None else "custGeom"
                        else:
                            geom = tag
                        anchor_types[name] = geom

                    if tag == "sp":
                        xfrm = elem.find("p:spPr/a:xfrm", NS)
                        tx_body = elem.find("p:txBody", NS)
                        if xfrm is not None and tx_body is not None:
                            ext = xfrm.find("a:ext", NS)
                            off = xfrm.find("a:off", NS)
                            if ext is not None and off is not None:
                                bx = int(off.attrib["x"]) / EMU_PER_INCH
                                by = int(off.attrib["y"]) / EMU_PER_INCH
                                bw = int(ext.attrib["cx"]) / EMU_PER_INCH
                                bh = int(ext.attrib["cy"]) / EMU_PER_INCH
                                if bx + bw > 0 and bx < 10.0 and by + bh > 0 and by < 5.625:
                                    specs, pad = _extract_tx_specs(tx_body)
                                    if specs:
                                        m = ls.measure_text_block(specs, bw, bh, pad_in=pad, safety_in=0.02)
                                        if not m.fits and m.overflow_in > 0.04:
                                            overflow_count += 1
                                            warnings.append(
                                                f"{os.path.basename(sf)} '{name}': text height {m.required_h_in:.2f} in > box {bh:.2f} in (+{m.overflow_in:.2f} in)"
                                            )

            dups = [k for k, v in Counter(bang_list).items() if v > 1]
            if dups:
                errors.append(f"Duplicate !! names on {os.path.basename(sf)}: {dups}")

            has_morph = root.find(".//p159:morph", NS) is not None
            if idx >= 2 and not has_morph:
                warnings.append(f"{os.path.basename(sf)} has no <p159:morph> transition")

            # Audit <p:timing> target spids
            timing_tgts = root.findall(".//p:timing//p:spTgt", NS)
            for st in timing_tgts:
                spid = st.attrib.get("spid")
                if spid and spid not in shape_ids:
                    errors.append(f"{os.path.basename(sf)}: <p:timing> references missing spid={spid}")

            scene3d_count = len(root.findall(".//a:scene3d", NS))
            if scene3d_count and target == "web":
                warnings.append(f"{os.path.basename(sf)} uses <a:scene3d> ({scene3d_count}x) which PowerPoint Web flattens")

            small_fonts = []
            for rpr in root.findall(".//a:rPr", NS):
                sz = int(rpr.attrib.get("sz", 1200)) / 100.0
                if sz < 9.0:
                    small_fonts.append(sz)
            if small_fonts:
                warnings.append(f"{os.path.basename(sf)} has text < 9.0 pt: {sorted(set(small_fonts))}")

            rot_count = len([el for el in root.findall(".//a:xfrm", NS) if "rot" in el.attrib])
            cust_count = len(root.findall(".//a:custGeom", NS))
            print(
                f"  {os.path.basename(sf)}: !!anchors={len(bang_list)} (dups={len(dups)}), "
                f"morph={has_morph}, rotated={rot_count}, custGeom={cust_count}, "
                f"timing={len(timing_tgts) // 2 if timing_tgts else 0}, scene3d={scene3d_count}, overflow={overflow_count}"
            )
            slide_anchors.append(anchor_types)

        for i in range(len(slide_anchors) - 1):
            a_cur, a_nxt = slide_anchors[i], slide_anchors[i + 1]
            for name in sorted(set(a_cur) & set(a_nxt)):
                if a_cur[name] != a_nxt[name]:
                    warnings.append(
                        f"Geometry type mismatch on '{name}' between slide{i + 1} ({a_cur[name]}) and slide{i + 2} ({a_nxt[name]})"
                    )

    return errors, warnings


def _iter_tree_elements(container):
    """Yield shape/pic/graphicFrame elements recursively through <p:grpSp>."""
    for elem in list(container):
        tag = elem.tag.split("}")[-1]
        if tag == "grpSp":
            yield elem
            yield from _iter_tree_elements(elem)
        elif tag in ("sp", "pic", "graphicFrame"):
            yield elem


def _hex_rgba(h: str, alpha_pct: float = 100.0) -> tuple[int, int, int, int]:
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), int(round(255 * (alpha_pct / 100.0))))


def _parse_color(el) -> tuple[int, int, int, int] | None:
    if el is None:
        return None
    srgb = el.find(".//a:srgbClr", NS)
    if srgb is None:
        return None
    val = srgb.attrib["val"]
    alpha_el = srgb.find("a:alpha", NS)
    alpha = int(alpha_el.attrib["val"]) / 1000.0 if alpha_el is not None else 100.0
    return _hex_rgba(val, alpha)


def _rotate_pts(pts, cx, cy, deg):
    if abs(deg) < 1e-6:
        return pts
    rad = math.radians(deg)
    c, s = math.cos(rad), math.sin(rad)
    return [(cx + (x - cx) * c - (y - cy) * s, cy + (x - cx) * s + (y - cy) * c) for x, y in pts]


def _sample_custgeom_subpaths(path_el, x, y, w, h) -> list[list[tuple[float, float]]]:
    """Sample an OOXML <a:path> into one or more polygon subpaths (supports compound holes)."""
    cw = max(1, int(path_el.attrib.get("w", 10000)))
    ch = max(1, int(path_el.attrib.get("h", 10000)))

    def to_px(pt_el):
        return (x + (int(pt_el.attrib["x"]) / cw) * w, y + (int(pt_el.attrib["y"]) / ch) * h)

    subpaths: list[list[tuple[float, float]]] = []
    pts: list[tuple[float, float]] = []
    for seg in path_el:
        stag = seg.tag.split("}")[-1]
        if stag == "moveTo":
            if pts:
                subpaths.append(pts)
                pts = []
            pt = seg.find("a:pt", NS)
            if pt is not None:
                pts.append(to_px(pt))
        elif stag == "lnTo":
            pt = seg.find("a:pt", NS)
            if pt is not None:
                pts.append(to_px(pt))
        elif stag == "cubicBezTo":
            p0 = pts[-1] if pts else (x, y)
            cpts = [to_px(p) for p in seg.findall("a:pt", NS)]
            if len(cpts) == 3:
                p1, p2, p3 = cpts
                for step in range(1, 13):
                    t = step / 12.0
                    mt = 1.0 - t
                    bx = (mt ** 3) * p0[0] + 3 * (mt ** 2) * t * p1[0] + 3 * mt * (t ** 2) * p2[0] + (t ** 3) * p3[0]
                    by = (mt ** 3) * p0[1] + 3 * (mt ** 2) * t * p1[1] + 3 * mt * (t ** 2) * p2[1] + (t ** 3) * p3[1]
                    pts.append((bx, by))
        elif stag == "quadBezTo":
            p0 = pts[-1] if pts else (x, y)
            qpts = [to_px(p) for p in seg.findall("a:pt", NS)]
            if len(qpts) == 2:
                p1, p2 = qpts
                for step in range(1, 10):
                    t = step / 10.0
                    mt = 1.0 - t
                    bx = (mt ** 2) * p0[0] + 2 * mt * t * p1[0] + (t ** 2) * p2[0]
                    by = (mt ** 2) * p0[1] + 2 * mt * t * p1[1] + (t ** 2) * p2[1]
                    pts.append((bx, by))
        elif stag == "close":
            if pts:
                subpaths.append(pts)
                pts = []
    if pts:
        subpaths.append(pts)
    return subpaths


def _draw_text_body(draw: ImageDraw.ImageDraw, tx_body, x, y, w, h):
    if tx_body is None:
        return
    bpr = tx_body.find("a:bodyPr", NS)
    l_ins = int(bpr.attrib.get("lIns", 0)) / EMU_PER_INCH * PX_PER_IN
    t_ins = int(bpr.attrib.get("tIns", 0)) / EMU_PER_INCH * PX_PER_IN
    r_ins = int(bpr.attrib.get("rIns", 0)) / EMU_PER_INCH * PX_PER_IN
    b_ins = int(bpr.attrib.get("bIns", 0)) / EMU_PER_INCH * PX_PER_IN
    anchor = bpr.attrib.get("anchor", "t")
    box_w = max(10, w - l_ins - r_ins)

    rendered_lines = []
    for p in tx_body.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        algn = ppr.attrib.get("algn", "l") if ppr is not None else "l"
        ln_spc_el = ppr.find("a:lnSpc/a:spcPct", NS) if ppr is not None else None
        ln_mult = int(ln_spc_el.attrib["val"]) / 100000.0 if ln_spc_el is not None else 1.2
        spc_aft_el = ppr.find("a:spcAft/a:spcPts", NS) if ppr is not None else None
        spc_aft_px = (int(spc_aft_el.attrib["val"]) / 100.0) * (PX_PER_IN / 72.0) if spc_aft_el is not None else 0

        tokens = []
        max_pt = 0.0
        for r in p.findall("a:r", NS):
            rpr = r.find("a:rPr", NS)
            sz_pt = int(rpr.attrib.get("sz", 1100)) / 100.0 if rpr is not None else 11.0
            max_pt = max(max_pt, sz_pt)
            bold = (rpr.attrib.get("b", "0") == "1") if rpr is not None else False
            col = (_parse_color(rpr.find("a:solidFill", NS)) if rpr is not None else None) or (248, 250, 252, 255)
            lat = rpr.find("a:latin", NS) if rpr is not None else None
            is_mono = lat is not None and "Mono" in lat.attrib.get("typeface", "")
            t_el = r.find("a:t", NS)
            txt = t_el.text if (t_el is not None and t_el.text) else ""
            for tok in ls.segment_thai_lao_tokens(txt):
                tokens.append((tok, sz_pt, bold, col, is_mono))
        if max_pt <= 0.0:
            max_pt = 11.0

        lines = [[]]
        cur_w = 0.0
        for tok_txt, sz_pt, bold, col, is_mono in tokens:
            if tok_txt == "\n":
                lines.append([])
                cur_w = 0.0
                continue
            if tok_txt.strip() == "" and not lines[-1]:
                continue
            kind = ls._detect_token_kind(tok_txt, mono=is_mono)
            font = get_font(kind, bold, sz_pt)
            tw = font.getlength(tok_txt)
            if cur_w + tw > box_w:
                if tok_txt.strip() == "" and lines[-1]:
                    continue
                if tw > box_w and kind in ("thai", "lao"):
                    for cl in ls.split_clusters(tok_txt):
                        cw = font.getlength(cl)
                        if cur_w + cw > box_w and lines[-1]:
                            lines.append([])
                            cur_w = 0.0
                        lines[-1].append((cl, sz_pt, bold, col, is_mono, cw))
                        cur_w += cw
                    continue
                if lines[-1]:
                    lines.append([])
                    cur_w = 0.0
            lines[-1].append((tok_txt, sz_pt, bold, col, is_mono, tw))
            cur_w += tw

        line_h = max_pt * (PX_PER_IN / 72.0) * ln_mult
        for li, line in enumerate(lines):
            segs: list[tuple[str, object, tuple[int, int, int, int], float]] = []
            for tok_txt, sz_pt, bold, col, is_mono, _ in line:
                for kind, chunk in _split_render_runs(tok_txt, mono=is_mono):
                    font = get_font(kind, bold, sz_pt)
                    cw = font.getlength(chunk)
                    if segs and segs[-1][1] is font and segs[-1][2] == col:
                        prev_t, _, _, prev_w = segs[-1]
                        segs[-1] = (prev_t + chunk, font, col, prev_w + cw)
                    else:
                        segs.append((chunk, font, col, cw))
            extra = spc_aft_px if li == len(lines) - 1 else 0
            rendered_lines.append((segs, algn, line_h, extra))

    total_h = sum(lh + ex for _, _, lh, ex in rendered_lines)
    cy = y + t_ins + max(0.0, (h - t_ins - b_ins - total_h) / 2.0) if anchor == "ctr" else y + t_ins

    for line, algn, lh, extra in rendered_lines:
        line_w = sum(tw for _, _, _, tw in line)
        if algn == "ctr":
            cx = x + l_ins + (box_w - line_w) / 2.0
        elif algn == "r":
            cx = x + l_ins + (box_w - line_w)
        else:
            cx = x + l_ins
        for txt, font, col, _ in line:
            draw.text((cx, cy), txt, font=font, fill=col)
            cx += font.getlength(txt)
        cy += lh + extra


def _split_render_runs(text: str, mono: bool = False) -> list[tuple[str, str]]:
    """Split a wrap token into contiguous font-script runs so ASCII punctuation never hits Thai/Lao-only TTFs."""
    if not text:
        return []
    if mono:
        return [("mono", text)]

    def ch_kind(ch: str) -> str:
        if ls.is_lao(ch):
            return "lao"
        if ls.is_thai(ch):
            return "thai"
        return "latin"

    runs: list[tuple[str, str]] = []
    cur = ch_kind(text[0])
    buf = [text[0]]
    for ch in text[1:]:
        k = cur if ls.is_combining_thai_lao(ch) else ch_kind(ch)
        if k == cur:
            buf.append(ch)
        else:
            runs.append((cur, "".join(buf)))
            cur = k
            buf = [ch]
    runs.append((cur, "".join(buf)))
    return runs


def _measure_script_text(text: str, bold: bool, pt: float) -> float:
    if not text:
        return 0.0
    total = 0.0
    for kind, chunk in _split_render_runs(text):
        f = get_font(kind, bold, pt)
        total += f.getlength(chunk)
    return total


def _draw_script_text(draw: ImageDraw.ImageDraw, pos: tuple[float, float], text: str, bold: bool, pt: float, fill):
    cx, cy = pos
    for kind, chunk in _split_render_runs(text):
        f = get_font(kind, bold, pt)
        draw.text((cx, cy), chunk, font=f, fill=fill)
        cx += f.getlength(chunk)


def _render_chart(layer: Image.Image, chart_xml_bytes: bytes, x, y, w, h):
    draw = ImageDraw.Draw(layer)
    root = ET.fromstring(chart_xml_bytes)

    # Read actual colors from chart XML (works on both light and dark themes)
    title_col = _parse_color(root.find(".//c:title//a:defRPr/a:solidFill", NS)) or (248, 250, 252, 255)
    cat_col = _parse_color(root.find(".//c:catAx//a:defRPr/a:solidFill", NS)) or title_col
    val_col = _parse_color(root.find(".//c:valAx//a:defRPr/a:solidFill", NS)) or (148, 163, 184, 255)
    grid_col = _parse_color(root.find(".//c:valAx/c:majorGridlines//a:solidFill", NS)) or (51, 65, 85, 140)
    leg_col = _parse_color(root.find(".//c:legend//a:defRPr/a:solidFill", NS)) or cat_col

    title_el = root.find(".//c:title//a:t", NS)
    title = title_el.text if title_el is not None else ""
    title_w = _measure_script_text(title, True, 13.0)
    _draw_script_text(draw, (x + (w - title_w) / 2.0, y + 8), title, True, 13.0, title_col)

    bar_el = root.find(".//c:barChart", NS)
    line_el = root.find(".//c:lineChart", NS)
    doughnut_el = root.find(".//c:doughnutChart", NS)

    if doughnut_el is not None:
        ser = doughnut_el.find("c:ser", NS)
        if ser is None:
            return
        cats = [pt.find("c:v", NS).text for pt in ser.findall(".//c:cat//c:pt", NS)]
        vals = [float(pt.find("c:v", NS).text) for pt in ser.findall(".//c:val//c:pt", NS)]
        dpts = {
            int(d.find("c:idx", NS).attrib["val"]): _parse_color(d.find("c:spPr/a:solidFill", NS))
            for d in ser.findall("c:dPt", NS)
        }
        palette = [(56, 220, 252, 255), (16, 185, 129, 255), (245, 158, 11, 255), (96, 165, 250, 255)]
        hole_el = doughnut_el.find("c:holeSize", NS)
        hole_frac = (int(hole_el.attrib.get("val", 58)) / 100.0) if hole_el is not None else 0.58
        total = max(1e-6, sum(vals))
        dia = min(w * 0.54, h - 74)
        cx, cy = x + w * 0.34, y + 44 + (h - 60) / 2.0
        r_out = dia / 2.0
        r_in = r_out * hole_frac
        ang = -90.0
        for i, v in enumerate(vals):
            sweep = (v / total) * 360.0
            col = dpts.get(i) or palette[i % len(palette)]
            draw.pieslice((cx - r_out, cy - r_out, cx + r_out, cy + r_out), start=ang, end=ang + sweep, fill=col)
            ang += sweep
        # Cut center hole with transparent punch on a temp mask
        draw.ellipse((cx - r_in, cy - r_in, cx + r_in, cy + r_in), fill=(0, 0, 0, 0))
        lx, ly = cx + r_out + 24, y + 58
        for i, (cat, v) in enumerate(zip(cats, vals)):
            col = dpts.get(i) or palette[i % len(palette)]
            draw.rectangle((lx, ly + 3, lx + 12, ly + 15), fill=col)
            _draw_script_text(draw, (lx + 18, ly), f"{cat} ({v:g})", True, 10.0, leg_col)
            ly += 24
        return

    chart_node = bar_el if bar_el is not None else line_el
    if chart_node is None:
        return
    sers = chart_node.findall("c:ser", NS)
    if not sers:
        return

    cats = [pt.find("c:v", NS).text for pt in sers[0].findall(".//c:cat//c:pt", NS)]
    series = []
    for ser in sers:
        name = ser.find(".//c:tx//c:v", NS).text
        col = (
            _parse_color(ser.find("c:spPr/a:solidFill", NS))
            or _parse_color(ser.find("c:spPr/a:ln/a:solidFill", NS))
            or (56, 220, 252, 255)
        )
        dlbl_col = _parse_color(ser.find("c:dLbls//a:defRPr/a:solidFill", NS)) or title_col
        vals = [float(pt.find("c:v", NS).text) for pt in ser.findall(".//c:val//c:pt", NS)]
        series.append((name, col, dlbl_col, vals))

    plot_x, plot_y, plot_w, plot_h = x + 56, y + 48, w - 76, h - 110
    f_ax = get_font("latin", False, 9.5)

    grouping_el = chart_node.find("c:grouping", NS)
    grouping = grouping_el.attrib.get("val", "clustered") if grouping_el is not None else "clustered"
    bar_dir_el = chart_node.find("c:barDir", NS)
    bar_dir = bar_dir_el.attrib.get("val", "col") if bar_dir_el is not None else "col"

    max_el = root.find(".//c:valAx/c:scaling/c:max", NS)
    if grouping == "percentStacked":
        data_max = 100.0
    elif grouping == "stacked":
        data_max = max([sum(vs[ci] for _, _, _, vs in series) for ci in range(len(cats))] + [1.0])
    else:
        data_max = max([v for _, _, _, vs in series for v in vs] + [1.0])
    max_v = float(max_el.attrib["val"]) if max_el is not None else data_max * 1.15
    max_v = max(1e-6, max_v)

    if line_el is not None or bar_dir == "col":
        for step in range(5):
            val_tick = max_v * (step / 4.0)
            gy = plot_y + plot_h - (val_tick / max_v) * plot_h
            draw.line([(plot_x, gy), (plot_x + plot_w, gy)], fill=grid_col, width=1)
            lbl = f"{val_tick:g}"
            draw.text((plot_x - f_ax.getlength(lbl) - 8, gy - 6), lbl, font=f_ax, fill=val_col)

    group_w = plot_w / max(1, len(cats))
    if line_el is not None:
        xs = [plot_x + (ci + 0.5) * group_w for ci in range(len(cats))]
        for _, col, dlbl_col, vals in series:
            pts = [(xs[ci], plot_y + plot_h - (vals[ci] / max_v) * plot_h) for ci in range(len(cats))]
            if len(pts) >= 2:
                draw.line(pts, fill=col, width=3)
            for ci, (px, py) in enumerate(pts):
                draw.ellipse((px - 4, py - 4, px + 4, py + 4), fill=col)
                vtxt = f"{vals[ci]:g}"
                vw = _measure_script_text(vtxt, True, 9.5)
                _draw_script_text(draw, (px - vw / 2.0, py - 18), vtxt, True, 9.5, dlbl_col)
        for ci, cat in enumerate(cats):
            cw = _measure_script_text(cat, True, 10.5)
            _draw_script_text(draw, (xs[ci] - cw / 2.0, plot_y + plot_h + 10), cat, True, 10.5, cat_col)
    elif bar_dir == "bar":
        row_h = plot_h / max(1, len(cats))
        bh = min(30, row_h * 0.55)
        for ci, cat in enumerate(cats):
            cy = plot_y + ci * row_h + row_h / 2.0
            cw = _measure_script_text(cat, True, 10.0)
            _draw_script_text(draw, (plot_x - cw - 8, cy - 7), cat, True, 10.0, cat_col)
            cur_x = plot_x
            cat_sum = max(1e-6, sum(vs[ci] for _, _, _, vs in series))
            for _, col, dlbl_col, vals in series:
                v = (vals[ci] / cat_sum * 100.0) if grouping == "percentStacked" else vals[ci]
                bw_px = (v / max_v) * plot_w
                draw.rectangle((cur_x, cy - bh / 2.0, cur_x + bw_px, cy + bh / 2.0), fill=col)
                cur_x += bw_px
    elif grouping in ("stacked", "percentStacked"):
        bar_w = min(56, group_w * 0.55)
        for ci, cat in enumerate(cats):
            gx = plot_x + ci * group_w + group_w / 2.0
            bx = gx - bar_w / 2.0
            cur_y = plot_y + plot_h
            cat_sum = max(1e-6, sum(vs[ci] for _, _, _, vs in series))
            for _, col, dlbl_col, vals in series:
                v = (vals[ci] / cat_sum * 100.0) if grouping == "percentStacked" else vals[ci]
                bh = (v / max_v) * plot_h
                draw.rectangle((bx, cur_y - bh, bx + bar_w, cur_y), fill=col)
                cur_y -= bh
            cw = _measure_script_text(cat, True, 10.5)
            _draw_script_text(draw, (gx - cw / 2.0, plot_y + plot_h + 10), cat, True, 10.5, cat_col)
    else:
        bar_w = min(44, group_w * 0.28)
        gap = 8
        for ci, cat in enumerate(cats):
            gx = plot_x + ci * group_w + group_w / 2.0
            start_x = gx - (len(series) * bar_w + (len(series) - 1) * gap) / 2.0
            for si, (_, col, dlbl_col, vals) in enumerate(series):
                bx = start_x + si * (bar_w + gap)
                bh = (vals[ci] / max_v) * plot_h
                by = plot_y + plot_h - bh
                draw.rounded_rectangle((bx, by, bx + bar_w, plot_y + plot_h), radius=4, fill=col)
                vtxt = f"{vals[ci]:g}"
                vw = _measure_script_text(vtxt, True, 10.5)
                _draw_script_text(draw, (bx + (bar_w - vw) / 2.0, by - 18), vtxt, True, 10.5, dlbl_col)
            cw = _measure_script_text(cat, True, 10.5)
            _draw_script_text(draw, (gx - cw / 2.0, plot_y + plot_h + 10), cat, True, 10.5, cat_col)

    total_leg_w = sum(20 + _measure_script_text(name, True, 10.5) + 24 for name, _, _, _ in series) - 24
    lx, ly = x + (w - max(10, total_leg_w)) / 2.0, y + h - 28
    for name, col, _, _ in series:
        draw.rectangle((lx, ly + 2, lx + 14, ly + 14), fill=col)
        _draw_script_text(draw, (lx + 20, ly), name, True, 10.5, leg_col)
        lx += 20 + _measure_script_text(name, True, 10.5) + 24


def render_ooxml_slides(pptx_path: str, out_dir: str) -> list[Image.Image]:
    os.makedirs(out_dir, exist_ok=True)
    rendered: list[Image.Image] = []
    with zipfile.ZipFile(pptx_path, "r") as zf:
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        sld_sz = pres.find("p:sldSz", NS)
        w_in = int(sld_sz.attrib["cx"]) / EMU_PER_INCH if sld_sz is not None else 10.0
        h_in = int(sld_sz.attrib["cy"]) / EMU_PER_INCH if sld_sz is not None else 5.625
        w_px, h_px = int(round(w_in * PX_PER_IN)), int(round(h_in * PX_PER_IN))

        slide_files = sorted(
            [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda s: int(s.split("slide")[-1].split(".")[0]),
        )
        for idx, sf in enumerate(slide_files, start=1):
            root = ET.fromstring(zf.read(sf))
            rels_path = posixpath.join(posixpath.dirname(sf), "_rels", posixpath.basename(sf) + ".rels")
            rels = {}
            if rels_path in zf.namelist():
                for rel in ET.fromstring(zf.read(rels_path)):
                    rels[rel.attrib["Id"]] = posixpath.normpath(posixpath.join(posixpath.dirname(sf), rel.attrib["Target"]))

            bg_col = _parse_color(root.find(".//p:bg//a:solidFill", NS)) or (15, 23, 42, 255)
            canvas = Image.new("RGBA", (w_px, h_px), bg_col)
            sp_tree = root.find(".//p:spTree", NS)
            if sp_tree is not None:
                for elem in _iter_tree_elements(sp_tree):
                    tag = elem.tag.split("}")[-1]
                    if tag == "grpSp":
                        continue
                    layer = Image.new("RGBA", (w_px, h_px), (0, 0, 0, 0))
                    draw = ImageDraw.Draw(layer)

                    if tag == "sp":
                        sp_pr = elem.find("p:spPr", NS)
                        if sp_pr is None:
                            continue
                        xfrm = sp_pr.find("a:xfrm", NS)
                        if xfrm is None:
                            continue
                        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
                        if off is None or ext is None:
                            continue
                        x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
                        y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
                        w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
                        h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
                        rot_deg = int(xfrm.attrib.get("rot", 0)) / 60000.0

                        fill_col = _parse_color(sp_pr.find("a:solidFill", NS))
                        is_radial = sp_pr.find(".//a:gradFill/a:path[@path='circle']", NS) is not None
                        grad_stops = []
                        if fill_col is None:
                            for gs in sp_pr.findall(".//a:gradFill//a:gs", NS):
                                gc = _parse_color(gs)
                                if gc is not None:
                                    grad_stops.append((int(gs.attrib.get("pos", 0)) / 100000.0, gc))
                            if grad_stops:
                                fill_col = grad_stops[0][1]
                        ln_el = sp_pr.find("a:ln", NS)
                        line_col = _parse_color(ln_el.find("a:solidFill", NS)) if ln_el is not None else None
                        line_w = max(1, int(round(int(ln_el.attrib.get("w", 12700)) / 12700.0 * (PX_PER_IN / 72.0)))) if (ln_el is not None and line_col) else 0

                        cust = sp_pr.find("a:custGeom", NS)
                        prst = sp_pr.find("a:prstGeom", NS)
                        if cust is not None:
                            path_els = cust.findall(".//a:path", NS)
                            all_subpaths: list[list[tuple[float, float]]] = []
                            for p_el in path_els:
                                all_subpaths.extend(_sample_custgeom_subpaths(p_el, x, y, w, h))
                            rotated_subs = [
                                _rotate_pts(sub, x + w / 2.0, y + h / 2.0, rot_deg)
                                for sub in all_subpaths if len(sub) >= 3
                            ]
                            if len(rotated_subs) == 1:
                                draw.polygon(rotated_subs[0], fill=fill_col, outline=line_col, width=line_w if line_col else 0)
                            elif len(rotated_subs) > 1:
                                if fill_col is not None:
                                    mask = Image.new("L", (w_px, h_px), 0)
                                    for sub in rotated_subs:
                                        sub_m = Image.new("L", (w_px, h_px), 0)
                                        ImageDraw.Draw(sub_m).polygon(sub, fill=255)
                                        mask = ImageChops.difference(mask, sub_m)
                                    fill_img = Image.new("RGBA", (w_px, h_px), fill_col)
                                    layer.paste(fill_img, (0, 0), mask)
                                if line_col is not None:
                                    for sub in rotated_subs:
                                        draw.line(sub + [sub[0]], fill=line_col, width=max(1, line_w))
                        elif prst is not None:
                            stype = prst.attrib["prst"]
                            if stype == "ellipse" and is_radial and len(grad_stops) >= 2:
                                steps = 16
                                base_rgb = grad_stops[0][1][:3]
                                a_in = grad_stops[0][1][3]
                                a_out = grad_stops[-1][1][3]
                                for si in range(steps, 0, -1):
                                    frac = si / float(steps)
                                    alpha_val = int(round(a_out + (a_in - a_out) * ((1.0 - frac) ** 1.6)))
                                    rw, rh = w * frac, h * frac
                                    ox, oy = x + (w - rw) / 2.0, y + (h - rh) / 2.0
                                    draw.ellipse((ox, oy, ox + rw, oy + rh), fill=(base_rgb[0], base_rgb[1], base_rgb[2], max(0, min(255, alpha_val // 4))))
                            elif stype == "ellipse":
                                draw.ellipse((x, y, x + w, y + h), fill=fill_col, outline=line_col, width=line_w)
                            elif stype == "rightArrow":
                                pts = _rotate_pts([
                                    (x, y + h * 0.25), (x + w * 0.55, y + h * 0.25),
                                    (x + w * 0.55, y), (x + w, y + h * 0.5),
                                    (x + w * 0.55, y + h), (x + w * 0.55, y + h * 0.75),
                                    (x, y + h * 0.75),
                                ], x + w / 2.0, y + h / 2.0, rot_deg)
                                draw.polygon(pts, fill=fill_col, outline=line_col, width=line_w if line_col else 0)
                            elif abs(rot_deg) > 1e-3:
                                pts = _rotate_pts([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], x + w / 2.0, y + h / 2.0, rot_deg)
                                draw.polygon(pts, fill=fill_col, outline=line_col, width=line_w)
                            else:
                                adj_gd = prst.find(".//a:gd[@name='adj']", NS)
                                adj_val = int(adj_gd.attrib["fmla"].split()[-1]) if adj_gd is not None else 16667
                                rad = min(w, h) * (adj_val / 100000.0) if stype == "roundRect" else 0
                                draw.rounded_rectangle((x, y, x + w, y + h), radius=rad, fill=fill_col, outline=line_col, width=line_w)

                        _draw_text_body(draw, elem.find("p:txBody", NS), x, y, w, h)

                    elif tag == "pic":
                        sp_pr = elem.find("p:spPr", NS)
                        xfrm = sp_pr.find("a:xfrm", NS) if sp_pr is not None else None
                        if xfrm is None:
                            continue
                        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
                        x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
                        y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
                        w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
                        h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
                        rot_deg = int(xfrm.attrib.get("rot", 0)) / 60000.0
                        blip = elem.find(".//a:blip", NS)
                        rid = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
                        if not rid or rid not in rels or rels[rid] not in zf.namelist():
                            continue
                        img = Image.open(io.BytesIO(zf.read(rels[rid]))).convert("RGBA")
                        src = elem.find(".//a:srcRect", NS)
                        if src is not None and any(k in src.attrib for k in ("l", "t", "r", "b")):
                            iw, ih = img.size
                            l_px = max(0, min(iw - 1, int(round(int(src.attrib.get("l", 0)) / 100000.0 * iw))))
                            t_px = max(0, min(ih - 1, int(round(int(src.attrib.get("t", 0)) / 100000.0 * ih))))
                            r_px = max(l_px + 1, min(iw, int(round(iw - int(src.attrib.get("r", 0)) / 100000.0 * iw))))
                            b_px = max(t_px + 1, min(ih, int(round(ih - int(src.attrib.get("b", 0)) / 100000.0 * ih))))
                            img = img.crop((l_px, t_px, r_px, b_px))
                        img = img.resize((max(1, int(round(w))), max(1, int(round(h)))), Image.Resampling.LANCZOS)
                        if abs(rot_deg) > 1e-3:
                            rot_img = img.rotate(-rot_deg, resample=Image.Resampling.BICUBIC, expand=True)
                            px = int(round(x + w / 2.0 - rot_img.width / 2.0))
                            py = int(round(y + h / 2.0 - rot_img.height / 2.0))
                            layer.paste(rot_img, (px, py), rot_img)
                        else:
                            layer.paste(img, (int(round(x)), int(round(y))), img)
                        ln_el = sp_pr.find("a:ln", NS)
                        line_col = _parse_color(ln_el.find("a:solidFill", NS)) if ln_el is not None else None
                        if line_col:
                            draw.rounded_rectangle((x, y, x + w, y + h), radius=8, outline=line_col, width=2)

                    elif tag == "graphicFrame":
                        xfrm = elem.find("p:xfrm", NS)
                        if xfrm is None:
                            continue
                        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
                        x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
                        y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
                        w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
                        h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
                        c_el = elem.find(".//c:chart", NS)
                        if c_el is not None:
                            rid = c_el.attrib[f"{{{NS['r']}}}id"]
                            _render_chart(layer, zf.read(rels[rid]), x, y, w, h)

                    canvas = Image.alpha_composite(canvas, layer)

            rgb = canvas.convert("RGB")
            out_png = os.path.join(out_dir, f"slide_{idx}.png")
            rgb.save(out_png, "PNG")
            rendered.append(rgb)
            print(f"Rendered (OOXML fallback) {out_png}")
    return rendered


def render_libreoffice_slides(pptx_path: str, out_dir: str, soffice_bin: str) -> list[Image.Image] | None:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        res = subprocess.run(
            [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx_path],
            capture_output=True, text=True
        )
        if res.returncode != 0:
            return None
        pdfs = [f for f in os.listdir(tmp) if f.endswith(".pdf")]
        if not pdfs:
            return None
        pdf_path = os.path.join(tmp, pdfs[0])
        prefix = os.path.join(out_dir, "slide")
        subprocess.run([pdftoppm, "-png", "-r", "150", pdf_path, prefix], check=True)

        pngs = sorted(
            [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("slide-") and f.endswith(".png")]
        )
        rendered = []
        for idx, p in enumerate(pngs, start=1):
            norm_p = os.path.join(out_dir, f"slide_{idx}.png")
            os.replace(p, norm_p)
            rendered.append(Image.open(norm_p).convert("RGB"))
            print(f"Rendered (LibreOffice) {norm_p}")
        return rendered


def build_contact_sheet(images: list[Image.Image], out_path: str):
    if not images:
        return
    cols = 3
    rows = math.ceil(len(images) / cols)
    tw, th = 640, 360
    pad = 24
    sheet = Image.new("RGB", (cols * tw + (cols + 1) * pad, rows * th + (rows + 1) * pad), (10, 15, 28))
    sd = ImageDraw.Draw(sheet)
    lbl_font = get_font("latin", True, 12.0)
    for i, img in enumerate(images):
        r, c = divmod(i, cols)
        ox = pad + c * (tw + pad)
        oy = pad + r * (th + pad)
        thumb = img.resize((tw, th), Image.Resampling.LANCZOS)
        sheet.paste(thumb, (ox, oy))
        sd.rectangle((ox, oy, ox + tw, oy + th), outline=(56, 220, 252), width=2)
        sd.text((ox + 8, oy + 6), f"SLIDE {i + 1}", font=lbl_font, fill=(56, 220, 252))
    sheet.save(out_path, "PNG")
    print(f"Saved contact sheet: {out_path}")


def main() -> int:
    import argparse
    parser = argparse.ArgumentParser(description="Audit and render a .pptx deck.")
    parser.add_argument("pptx_path", help="Path to .pptx deck")
    parser.add_argument("output_dir", nargs="?", default=None, help="Output directory for PNGs")
    parser.add_argument("--target", choices=("desktop", "web"), default="desktop", help="Compatibility target")
    args = parser.parse_args()

    pptx_path = os.path.abspath(args.pptx_path)
    out_dir = os.path.abspath(args.output_dir) if args.output_dir else os.path.join(os.path.dirname(pptx_path), "rendered_audit")
    os.makedirs(out_dir, exist_ok=True)

    errors, warnings = audit_deck(pptx_path, target=args.target)
    for w in warnings:
        print(f"  WARN: {w}")
    if errors:
        print("\nAUDIT FAILED:")
        for e in errors:
            print(f"  ERROR: {e}")
        return 1

    print("\nOOXML & MORPH AUDIT PASSED (0 errors).")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    images = render_libreoffice_slides(pptx_path, out_dir, soffice) if soffice else None
    if images is None:
        if not soffice:
            print("Note: LibreOffice not installed; using built-in OOXML-to-PNG fallback renderer.")
        images = render_ooxml_slides(pptx_path, out_dir)

    build_contact_sheet(images, os.path.join(out_dir, "contact_sheet.png"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
