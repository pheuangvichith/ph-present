#!/usr/bin/env python3
"""Render .pptx slides directly from OOXML to PNGs + contact_sheet.png for visual audit."""

import io
import math
import os
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image, ImageDraw, ImageFont

EMU_PER_INCH = 914400.0
W_IN, H_IN = 10.0, 5.625
PX_PER_IN = 128  # 1280 x 720
W_PX, H_PX = int(W_IN * PX_PER_IN), int(H_IN * PX_PER_IN)

FONT_PATHS = {
    ("latin", False): "/usr/share/fonts/liberation/LiberationSans-Regular.ttf",
    ("latin", True): "/usr/share/fonts/liberation/LiberationSans-Bold.ttf",
    ("thai", False): "/usr/share/fonts/noto/NotoSansThai-Regular.ttf",
    ("thai", True): "/usr/share/fonts/noto/NotoSansThai-Bold.ttf",
    ("mono", False): "/usr/share/fonts/TTF/DejaVuSansMono.ttf",
    ("mono", True): "/usr/share/fonts/TTF/DejaVuSansMono-Bold.ttf",
}

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

_FONT_CACHE = {}


def get_font(kind: str, bold: bool, pt: float) -> ImageFont.FreeTypeFont:
    px = max(8, int(round(pt * (PX_PER_IN / 72.0))))
    key = (kind, bold, px)
    if key not in _FONT_CACHE:
        _FONT_CACHE[key] = ImageFont.truetype(FONT_PATHS[(kind, bold)], px)
    return _FONT_CACHE[key]


def hex_rgba(h: str, alpha_pct: float = 100.0):
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    a = int(round(255 * (alpha_pct / 100.0)))
    return (r, g, b, a)


def parse_color(el):
    if el is None:
        return None
    srgb = el.find(".//a:srgbClr", NS)
    if srgb is None:
        return None
    val = srgb.attrib["val"]
    alpha_el = srgb.find("a:alpha", NS)
    alpha = int(alpha_el.attrib["val"]) / 1000.0 if alpha_el is not None else 100.0
    return hex_rgba(val, alpha)


def rotate_pts(pts, cx, cy, deg):
    if abs(deg) < 1e-6:
        return pts
    rad = math.radians(deg)
    c, s = math.cos(rad), math.sin(rad)
    out = []
    for x, y in pts:
        dx, dy = x - cx, y - cy
        out.append((cx + dx * c - dy * s, cy + dx * s + dy * c))
    return out


def is_thai(ch: str) -> bool:
    return "\u0E00" <= ch <= "\u0E7F"


def is_thai_combining(ch: str) -> bool:
    code = ord(ch)
    return code in (0x0E31,) or (0x0E34 <= code <= 0x0E3A) or (0x0E47 <= code <= 0x0E4E)


def split_clusters(text: str) -> list[str]:
    """Group combining Thai marks with their base consonant so diacritics never split."""
    clusters = []
    for ch in text:
        if clusters and is_thai_combining(ch):
            clusters[-1] += ch
        else:
            clusters.append(ch)
    return clusters


def draw_text_body(draw: ImageDraw.ImageDraw, tx_body, x, y, w, h):
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
        max_pt = 11.0
        for r in p.findall("a:r", NS):
            rpr = r.find("a:rPr", NS)
            sz_pt = int(rpr.attrib.get("sz", 1100)) / 100.0
            max_pt = max(max_pt, sz_pt)
            bold = rpr.attrib.get("b", "0") == "1"
            col = parse_color(rpr.find("a:solidFill", NS)) or (248, 250, 252, 255)
            lat = rpr.find("a:latin", NS)
            is_mono = lat is not None and "Mono" in lat.attrib.get("typeface", "")
            t_el = r.find("a:t", NS)
            txt = t_el.text if (t_el is not None and t_el.text) else ""
            for subline_idx, sub in enumerate(txt.split("\n")):
                if subline_idx > 0:
                    tokens.append(("\n", sz_pt, bold, col, is_mono))
                buf = ""
                for ch in sub:
                    if ch == " ":
                        if buf:
                            tokens.append((buf, sz_pt, bold, col, is_mono))
                            buf = ""
                        tokens.append((" ", sz_pt, bold, col, is_mono))
                    else:
                        buf += ch
                if buf:
                    tokens.append((buf, sz_pt, bold, col, is_mono))

        lines = [[]]
        cur_w = 0.0
        for tok_txt, sz_pt, bold, col, is_mono in tokens:
            if tok_txt == "\n":
                lines.append([])
                cur_w = 0.0
                continue
            kind = "mono" if is_mono else ("thai" if any(is_thai(c) for c in tok_txt) else "latin")
            font = get_font(kind, bold, sz_pt)
            tw = font.getlength(tok_txt)
            if cur_w + tw > box_w:
                if tok_txt == " " and lines[-1]:
                    continue
                if tw > box_w and kind == "thai":
                    for cl in split_clusters(tok_txt):
                        cw = font.getlength(cl)
                        if cur_w + cw > box_w and lines[-1]:
                            lines.append([])
                            cur_w = 0.0
                        lines[-1].append((cl, font, col, cw))
                        cur_w += cw
                    continue
                if lines[-1]:
                    lines.append([])
                    cur_w = 0.0
            lines[-1].append((tok_txt, font, col, tw))
            cur_w += tw

        line_h = max_pt * (PX_PER_IN / 72.0) * ln_mult
        for li, line in enumerate(lines):
            # Merge adjacent spans with same (font, col) for clean Thai shaping
            merged = []
            for txt, font, col, tw in line:
                if merged and merged[-1][1] is font and merged[-1][2] == col:
                    merged[-1][0] += txt
                    merged[-1][3] += tw
                else:
                    merged.append([txt, font, col, tw])
            extra = spc_aft_px if li == len(lines) - 1 else 0
            rendered_lines.append(([(m[0], m[1], m[2], m[3]) for m in merged], algn, line_h, extra))

    total_h = sum(lh + ex for _, _, lh, ex in rendered_lines)
    if anchor == "ctr":
        cy = y + t_ins + max(0.0, (h - t_ins - b_ins - total_h) / 2.0)
    else:
        cy = y + t_ins

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


def render_chart(layer: Image.Image, chart_xml_bytes: bytes, x, y, w, h):
    draw = ImageDraw.Draw(layer)
    root = ET.fromstring(chart_xml_bytes)
    title_el = root.find(".//c:title//a:t", NS)
    title = title_el.text if title_el is not None else ""
    f_title = get_font("latin", True, 13.0)
    tw = f_title.getlength(title)
    draw.text((x + (w - tw) / 2.0, y + 8), title, font=f_title, fill=(248, 250, 252, 255))

    sers = root.findall(".//c:barChart/c:ser", NS)
    if not sers:
        return
    cats = [pt.find("c:v", NS).text for pt in sers[0].findall(".//c:cat//c:pt", NS)]
    series = []
    for ser in sers:
        name = ser.find(".//c:tx//c:v", NS).text
        col = parse_color(ser.find("c:spPr/a:solidFill", NS)) or (56, 220, 252, 255)
        vals = [float(pt.find("c:v", NS).text) for pt in ser.findall(".//c:val//c:pt", NS)]
        series.append((name, col, vals))

    plot_x = x + 56
    plot_y = y + 48
    plot_w = w - 76
    plot_h = h - 110

    f_ax = get_font("latin", False, 9.5)
    f_cat = get_font("latin", True, 10.5)
    f_val = get_font("latin", True, 10.5)

    for pct in [0, 25, 50, 75, 100]:
        gy = plot_y + plot_h - (pct / 110.0) * plot_h
        draw.line([(plot_x, gy), (plot_x + plot_w, gy)], fill=(51, 65, 85, 140), width=1)
        lbl = f"{pct}%"
        draw.text((plot_x - f_ax.getlength(lbl) - 8, gy - 6), lbl, font=f_ax, fill=(148, 163, 184, 255))

    n_cats = len(cats)
    n_ser = len(series)
    group_w = plot_w / n_cats
    bar_w = min(44, group_w * 0.28)
    gap = 8

    for ci, cat in enumerate(cats):
        gx = plot_x + ci * group_w + group_w / 2.0
        total_bars_w = n_ser * bar_w + (n_ser - 1) * gap
        start_x = gx - total_bars_w / 2.0
        for si, (_, col, vals) in enumerate(series):
            bx = start_x + si * (bar_w + gap)
            bh = (vals[ci] / 110.0) * plot_h
            by = plot_y + plot_h - bh
            draw.rounded_rectangle((bx, by, bx + bar_w, plot_y + plot_h), radius=4, fill=col)
            vtxt = f"{int(vals[ci])}%"
            vw = f_val.getlength(vtxt)
            draw.text((bx + (bar_w - vw) / 2.0, by - 18), vtxt, font=f_val, fill=(248, 250, 252, 255))
        cw = f_cat.getlength(cat)
        draw.text((gx - cw / 2.0, plot_y + plot_h + 10), cat, font=f_cat, fill=(203, 213, 225, 255))

    # Legend
    lx = x + w / 2.0 - 150
    ly = y + h - 28
    for name, col, _ in series:
        draw.rectangle((lx, ly + 2, lx + 14, ly + 14), fill=col)
        draw.text((lx + 20, ly), name, font=f_cat, fill=(203, 213, 225, 255))
        lx += 20 + f_cat.getlength(name) + 24


def render_slide(zf: zipfile.ZipFile, slide_path: str) -> Image.Image:
    root = ET.fromstring(zf.read(slide_path))
    rels_path = posixpath.join(posixpath.dirname(slide_path), "_rels", posixpath.basename(slide_path) + ".rels")
    rels = {}
    if rels_path in zf.namelist():
        r_root = ET.fromstring(zf.read(rels_path))
        for rel in r_root:
            rels[rel.attrib["Id"]] = posixpath.normpath(posixpath.join(posixpath.dirname(slide_path), rel.attrib["Target"]))

    bg_col = parse_color(root.find(".//p:bg//a:solidFill", NS)) or (15, 23, 42, 255)
    canvas = Image.new("RGBA", (W_PX, H_PX), bg_col)

    sp_tree = root.find(".//p:spTree", NS)
    for elem in list(sp_tree):
        tag = elem.tag.split("}")[-1]
        layer = Image.new("RGBA", (W_PX, H_PX), (0, 0, 0, 0))
        draw = ImageDraw.Draw(layer)

        if tag == "sp":
            sp_pr = elem.find("p:spPr", NS)
            xfrm = sp_pr.find("a:xfrm", NS)
            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
            y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
            w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
            h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
            rot_deg = int(xfrm.attrib.get("rot", 0)) / 60000.0

            fill_col = parse_color(sp_pr.find("a:solidFill", NS))
            ln_el = sp_pr.find("a:ln", NS)
            line_col = parse_color(ln_el.find("a:solidFill", NS)) if ln_el is not None else None
            line_w = max(1, int(round(int(ln_el.attrib.get("w", 12700)) / 12700.0 * (PX_PER_IN / 72.0)))) if (ln_el is not None and line_col) else 0

            cust = sp_pr.find("a:custGeom", NS)
            prst = sp_pr.find("a:prstGeom", NS)
            if cust is not None:
                path_el = cust.find(".//a:path", NS)
                cw = int(path_el.attrib["w"])
                ch = int(path_el.attrib["h"])
                pts = []
                for seg in path_el:
                    pt = seg.find("a:pt", NS)
                    if pt is not None:
                        px_pt = x + (int(pt.attrib["x"]) / cw) * w
                        py_pt = y + (int(pt.attrib["y"]) / ch) * h
                        pts.append((px_pt, py_pt))
                pts = rotate_pts(pts, x + w / 2.0, y + h / 2.0, rot_deg)
                if len(pts) >= 3:
                    draw.polygon(pts, fill=fill_col, outline=line_col, width=line_w if line_col else 0)
            elif prst is not None:
                shape_type = prst.attrib["prst"]
                if shape_type == "ellipse":
                    draw.ellipse((x, y, x + w, y + h), fill=fill_col, outline=line_col, width=line_w)
                elif shape_type == "rightArrow":
                    pts = [
                        (x, y + h * 0.25), (x + w * 0.55, y + h * 0.25),
                        (x + w * 0.55, y), (x + w, y + h * 0.5),
                        (x + w * 0.55, y + h), (x + w * 0.55, y + h * 0.75),
                        (x, y + h * 0.75),
                    ]
                    draw.polygon(pts, fill=fill_col)
                elif abs(rot_deg) > 1e-3:
                    pts = rotate_pts([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], x + w / 2.0, y + h / 2.0, rot_deg)
                    draw.polygon(pts, fill=fill_col, outline=line_col, width=line_w)
                else:
                    adj_gd = prst.find(".//a:gd[@name='adj']", NS)
                    adj_val = int(adj_gd.attrib["fmla"].split()[-1]) if adj_gd is not None else 16667
                    rad = min(w, h) * (adj_val / 100000.0) if shape_type == "roundRect" else 0
                    draw.rounded_rectangle((x, y, x + w, y + h), radius=rad, fill=fill_col, outline=line_col, width=line_w)

            draw_text_body(draw, elem.find("p:txBody", NS), x, y, w, h)

        elif tag == "pic":
            sp_pr = elem.find("p:spPr", NS)
            xfrm = sp_pr.find("a:xfrm", NS)
            off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
            x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
            y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
            w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
            h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
            blip = elem.find(".//a:blip", NS)
            rid = blip.attrib[f"{{{NS['r']}}}embed"]
            img = Image.open(io.BytesIO(zf.read(rels[rid]))).convert("RGBA")
            src = elem.find(".//a:srcRect", NS)
            if src is not None and any(k in src.attrib for k in ("l", "t", "r", "b")):
                iw, ih = img.size
                l = int(src.attrib.get("l", 0)) / 100000.0 * iw
                t = int(src.attrib.get("t", 0)) / 100000.0 * ih
                r = iw - int(src.attrib.get("r", 0)) / 100000.0 * iw
                b = ih - int(src.attrib.get("b", 0)) / 100000.0 * ih
                img = img.crop((int(l), int(t), int(r), int(b)))
            img = img.resize((int(round(w)), int(round(h))), Image.Resampling.LANCZOS)
            layer.paste(img, (int(round(x)), int(round(y))), img)
            ln_el = sp_pr.find("a:ln", NS)
            line_col = parse_color(ln_el.find("a:solidFill", NS)) if ln_el is not None else None
            if line_col:
                draw.rounded_rectangle((x, y, x + w, y + h), radius=8, outline=line_col, width=2)

        elif tag == "graphicFrame":
            xfrm = elem.find("p:xfrm", NS)
            off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
            x = int(off.attrib["x"]) / EMU_PER_INCH * PX_PER_IN
            y = int(off.attrib["y"]) / EMU_PER_INCH * PX_PER_IN
            w = int(ext.attrib["cx"]) / EMU_PER_INCH * PX_PER_IN
            h = int(ext.attrib["cy"]) / EMU_PER_INCH * PX_PER_IN
            c_el = elem.find(".//c:chart", NS)
            if c_el is not None:
                rid = c_el.attrib[f"{{{NS['r']}}}id"]
                render_chart(layer, zf.read(rels[rid]), x, y, w, h)

        canvas = Image.alpha_composite(canvas, layer)

    return canvas.convert("RGB")


def main():
    pptx_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "Inside_the_Agent_Loop.pptx")
    out_dir = os.path.join(os.path.dirname(pptx_path), "rendered_audit")
    os.makedirs(out_dir, exist_ok=True)

    rendered = []
    with zipfile.ZipFile(pptx_path, "r") as zf:
        slide_files = sorted(
            [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda s: int(s.split("slide")[-1].split(".")[0])
        )
        for idx, sf in enumerate(slide_files, start=1):
            img = render_slide(zf, sf)
            out_png = os.path.join(out_dir, f"slide_{idx}.png")
            img.save(out_png, "PNG")
            rendered.append(img)
            print(f"Rendered {out_png}")

    # Build 3-col contact sheet
    cols = 3
    rows = math.ceil(len(rendered) / cols)
    tw, th = 640, 360
    pad = 24
    sheet = Image.new("RGB", (cols * tw + (cols + 1) * pad, rows * th + (rows + 1) * pad), (10, 15, 28))
    sd = ImageDraw.Draw(sheet)
    lbl_font = get_font("latin", True, 12.0)
    for i, img in enumerate(rendered):
        r, c = divmod(i, cols)
        ox = pad + c * (tw + pad)
        oy = pad + r * (th + pad)
        thumb = img.resize((tw, th), Image.Resampling.LANCZOS)
        sheet.paste(thumb, (ox, oy))
        sd.rectangle((ox, oy, ox + tw, oy + th), outline=(56, 220, 252), width=2)
        sd.text((ox + 8, oy + 6), f"SLIDE {i + 1}", font=lbl_font, fill=(56, 220, 252))

    sheet_path = os.path.join(out_dir, "contact_sheet.png")
    sheet.save(sheet_path, "PNG")
    print(f"Saved contact sheet: {sheet_path}")


if __name__ == "__main__":
    main()
