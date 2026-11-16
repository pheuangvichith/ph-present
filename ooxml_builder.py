#!/usr/bin/env python3
"""Minimal, strict-OOXML PowerPoint (.pptx) builder with Morph, custGeom, charts & script-aware fonts."""

import io
import math
import os
import zipfile
from xml.sax.saxutils import escape

EMU_PER_INCH = 914400
SLIDE_W = int(10.0 * EMU_PER_INCH)     # 9144000
SLIDE_H = int(5.625 * EMU_PER_INCH)    # 5143500


def emu(inches: float) -> int:
    return int(round(inches * EMU_PER_INCH))


def rot_val(deg: float) -> int:
    """Convert degrees to OOXML 60000ths of a degree."""
    return int(round((deg % 360.0) * 60000))


def _is_thai(ch: str) -> bool:
    return "\u0E00" <= ch <= "\u0E7F"


def split_script_runs(text: str):
    """Split text into contiguous runs tagged with script ('thai' or 'latin')."""
    if not text:
        return [("latin", "")]
    runs = []
    cur_script = "thai" if _is_thai(text[0]) else "latin"
    buf = [text[0]]
    for ch in text[1:]:
        # Keep spaces attached to the current run to avoid awkward run breaks
        s = cur_script if ch == " " else ("thai" if _is_thai(ch) else "latin")
        if s == cur_script:
            buf.append(ch)
        else:
            runs.append((cur_script, "".join(buf)))
            cur_script = s
            buf = [ch]
    runs.append((cur_script, "".join(buf)))
    return runs


def rpr_xml(size_pt: float, color: str, bold: bool = False, italic: bool = False,
            latin_font: str = "Arial", thai_font: str = "Noto Sans Thai",
            mono: bool = False, lang: str = "en-US") -> str:
    sz = int(round(size_pt * 100))
    b_attr = ' b="1"' if bold else ' b="0"'
    i_attr = ' i="1"' if italic else ''
    lf = "DejaVu Sans Mono" if mono else latin_font
    return (
        f'<a:rPr lang="{lang}" altLang="th-TH" sz="{sz}"{b_attr}{i_attr} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{escape(lf)}"/>'
        f'<a:ea typeface="Noto Sans CJK SC"/>'
        f'<a:cs typeface="{escape(thai_font)}"/>'
        f'</a:rPr>'
    )


def paragraph_xml(text: str, size_pt: float, color: str, bold: bool = False,
                  italic: bool = False, align: str = "l", line_spacing: float = 1.22,
                  space_after_pt: float = 0.0, latin_font: str = "Arial",
                  thai_font: str = "Noto Sans Thai", mono: bool = False) -> str:
    spc_pct = int(round(line_spacing * 100000))
    spc_aft = int(round(space_after_pt * 100))
    ppr = (
        f'<a:pPr algn="{align}">'
        f'<a:lnSpc><a:spcPct val="{spc_pct}"/></a:lnSpc>'
        f'<a:spcAft><a:spcPts val="{spc_aft}"/></a:spcAft>'
        f'<a:buNone/>'
        f'</a:pPr>'
    )
    if text == "":
        end_rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, mono)
        return f'<a:p>{ppr}{end_rpr.replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")}</a:p>'

    runs_xml = []
    for script, chunk in split_script_runs(text):
        lang = "th-TH" if script == "thai" else "en-US"
        rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, mono, lang=lang)
        runs_xml.append(f'<a:r>{rpr}<a:t>{escape(chunk)}</a:t></a:r>')
    end_rpr = rpr_xml(size_pt, color, bold, italic, latin_font, thai_font, mono)
    end_rpr = end_rpr.replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")
    return f'<a:p>{ppr}{"".join(runs_xml)}{end_rpr}</a:p>'


def shadow_xml(blur_pt: float = 12.0, dist_pt: float = 3.0, alpha_pct: int = 28) -> str:
    blur_emu = int(round(blur_pt * 12700))
    dist_emu = int(round(dist_pt * 12700))
    alpha_val = int(round(alpha_pct * 1000))
    return (
        f'<a:effectLst>'
        f'<a:outerShdw blurRad="{blur_emu}" dist="{dist_emu}" dir="5400000" algn="tl" rotWithShape="0">'
        f'<a:srgbClr val="000000"><a:alpha val="{alpha_val}"/></a:srgbClr>'
        f'</a:outerShdw>'
        f'</a:effectLst>'
    )


def fill_xml(fill_color: str | None, alpha_pct: int = 100) -> str:
    if fill_color is None:
        return "<a:noFill/>"
    if alpha_pct < 100:
        return f'<a:solidFill><a:srgbClr val="{fill_color}"><a:alpha val="{alpha_pct * 1000}"/></a:srgbClr></a:solidFill>'
    return f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'


def line_xml(line_color: str | None, line_pt: float = 1.0, dash: str | None = None, alpha_pct: int = 100) -> str:
    if line_color is None or line_pt <= 0:
        return '<a:ln><a:noFill/></a:ln>'
    w_emu = int(round(line_pt * 12700))
    if alpha_pct < 100:
        f_xml = f'<a:solidFill><a:srgbClr val="{line_color}"><a:alpha val="{alpha_pct * 1000}"/></a:srgbClr></a:solidFill>'
    else:
        f_xml = f'<a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill>'
    dash_xml = f'<a:prstDash val="{dash}"/>' if dash else '<a:prstDash val="solid"/>'
    return f'<a:ln w="{w_emu}">{f_xml}{dash_xml}</a:ln>'


def make_shape(sp_id: int, name: str, x_in: float, y_in: float, w_in: float, h_in: float,
               prst: str = "roundRect", adj: int | None = None, rot_deg: float = 0.0,
               fill: str | None = None, fill_alpha: int = 100,
               line: str | None = None, line_pt: float = 1.0, line_dash: str | None = None, line_alpha: int = 100,
               shadow: bool = False, paragraphs: list[str] | None = None,
               pad_in: tuple[float, float, float, float] = (0.14, 0.10, 0.14, 0.10),
               anchor: str = "t") -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    rot_attr = f' rot="{rot_val(rot_deg)}"' if abs(rot_deg) > 1e-6 else ""
    av_lst = f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>' if adj is not None else "<a:avLst/>"
    eff = shadow_xml() if shadow else ""
    l_ins, t_ins, r_ins, b_ins = [emu(p) for p in pad_in]
    p_xml = "".join(paragraphs) if paragraphs else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvSpPr/>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}">{av_lst}</a:prstGeom>'
        f'{fill_xml(fill, fill_alpha)}'
        f'{line_xml(line, line_pt, line_dash, line_alpha)}'
        f'{eff}'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" lIns="{l_ins}" tIns="{t_ins}" rIns="{r_ins}" bIns="{b_ins}" anchor="{anchor}">'
        f'<a:noAutofit/>'
        f'</a:bodyPr>'
        f'<a:lstStyle/>'
        f'{p_xml}'
        f'</p:txBody>'
        f'</p:sp>'
    )


def make_custgeom_shape(sp_id: int, name: str, x_in: float, y_in: float, w_in: float, h_in: float,
                        points_norm: list[tuple[int, int]], coord_w: int, coord_h: int,
                        rot_deg: float = 0.0, fill: str | None = None, fill_alpha: int = 100,
                        line: str | None = None, line_pt: float = 1.0, shadow: bool = False) -> str:
    """Build a native OOXML Freeform (<a:custGeom>) polygon."""
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    rot_attr = f' rot="{rot_val(rot_deg)}"' if abs(rot_deg) > 1e-6 else ""
    eff = shadow_xml() if shadow else ""
    segs = [f'<a:moveTo><a:pt x="{points_norm[0][0]}" y="{points_norm[0][1]}"/></a:moveTo>']
    for px, py in points_norm[1:]:
        segs.append(f'<a:lnTo><a:pt x="{px}" y="{py}"/></a:lnTo>')
    segs.append("<a:close/>")
    return (
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:custGeom>'
        f'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="0" t="0" r="r" b="b"/>'
        f'<a:pathLst><a:path w="{coord_w}" h="{coord_h}">{"".join(segs)}</a:path></a:pathLst>'
        f'</a:custGeom>'
        f'{fill_xml(fill, fill_alpha)}'
        f'{line_xml(line, line_pt)}'
        f'{eff}'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
        f'</p:sp>'
    )


def make_picture(sp_id: int, name: str, rel_id: str, x_in: float, y_in: float, w_in: float, h_in: float,
                 src_rect: tuple[int, int, int, int] = (0, 0, 0, 0),
                 line: str | None = None, line_pt: float = 1.5, shadow: bool = False) -> str:
    """Build native Picture (<p:pic>) with optional <a:srcRect> Crop-to-Pan values (in 1/1000ths of pct)."""
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    l, t, r, b = src_rect
    src_xml = f'<a:srcRect l="{l}" t="{t}" r="{r}" b="{b}"/>' if any(src_rect) else '<a:srcRect/>'
    eff = shadow_xml() if shadow else ""
    return (
        f'<p:pic>'
        f'<p:nvPicPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr/>'
        f'</p:nvPicPr>'
        f'<p:blipFill>'
        f'<a:blip r:embed="{rel_id}"/>'
        f'{src_xml}'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</p:blipFill>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>'
        f'{line_xml(line, line_pt)}'
        f'{eff}'
        f'</p:spPr>'
        f'</p:pic>'
    )


def make_chart_frame(sp_id: int, name: str, rel_id: str, x_in: float, y_in: float, w_in: float, h_in: float) -> str:
    x, y, w, h = emu(x_in), emu(y_in), emu(w_in), emu(h_in)
    return (
        f'<p:graphicFrame>'
        f'<p:nvGraphicFramePr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvGraphicFramePr/>'
        f'<p:nvPr/>'
        f'</p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rel_id}"/>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</p:graphicFrame>'
    )


def build_minimal_xlsx(categories: list[str], series: list[tuple[str, list[float]]]) -> bytes:
    """Create a valid embedded .xlsx workbook backing the native editable chart."""
    def col_letter(idx: int) -> str:
        return chr(ord("A") + idx)

    rows_xml = []
    # Row 1: header
    header_cells = ['<c r="A1" t="inlineStr"><is><t>Metric</t></is></c>']
    for s_idx, (s_name, _) in enumerate(series, start=1):
        ref = f"{col_letter(s_idx)}1"
        header_cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(s_name)}</t></is></c>')
    rows_xml.append(f'<row r="1">{"".join(header_cells)}</row>')

    # Data rows
    for r_idx, cat in enumerate(categories, start=2):
        cells = [f'<c r="A{r_idx}" t="inlineStr"><is><t>{escape(cat)}</t></is></c>']
        for s_idx, (_, vals) in enumerate(series, start=1):
            ref = f"{col_letter(s_idx)}{r_idx}"
            cells.append(f'<c r="{ref}"><v>{vals[r_idx - 2]}</v></c>')
        rows_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        '</worksheet>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '</styleSheet>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
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


def build_clustered_bar_chart_xml(title: str, categories: list[str],
                                  series: list[tuple[str, str, list[float]]],
                                  xlsx_rel_id: str = "rId1") -> str:
    """Build OOXML <c:chartSpace> clustered column chart with cached values + externalData."""
    n_cats = len(categories)
    cat_pts = "".join(f'<c:pt idx="{i}"><c:v>{escape(c)}</c:v></c:pt>' for i, c in enumerate(categories))
    ser_xml_list = []

    for idx, (s_name, s_color, vals) in enumerate(series):
        col_letter = chr(ord("B") + idx)
        val_pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
        ser_xml_list.append(
            f'<c:ser>'
            f'<c:idx val="{idx}"/><c:order val="{idx}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col_letter}$1</c:f>'
            f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape(s_name)}</c:v></c:pt></c:strCache>'
            f'</c:strRef></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="{s_color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>'
            f'<c:dLbls>'
            f'<c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>'
            f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1050" b="1"><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
            f'<c:dLblPos val="outEnd"/>'
            f'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            f'</c:dLbls>'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n_cats + 1}</c:f>'
            f'<c:strCache><c:ptCount val="{n_cats}"/>{cat_pts}</c:strCache>'
            f'</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!${col_letter}$2:${col_letter}${n_cats + 1}</c:f>'
            f'<c:numCache><c:formatCode>0"%"</c:formatCode><c:ptCount val="{n_cats}"/>{val_pts}</c:numCache>'
            f'</c:numRef></c:val>'
            f'</c:ser>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:date1904 val="0"/>'
        '<c:lang val="en-US"/>'
        '<c:roundedCorners val="0"/>'
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
        '<c:chart>'
        f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1"><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:r><a:t>{escape(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:autoTitleDeleted val="0"/>'
        '<c:plotArea>'
        '<c:layout/>'
        '<c:barChart>'
        '<c:barDir val="col"/>'
        '<c:grouping val="clustered"/>'
        '<c:varyColors val="0"/>'
        f'{"".join(ser_xml_list)}'
        '<c:gapWidth val="120"/>'
        '<c:overlap val="-15"/>'
        '<c:axId val="1001"/>'
        '<c:axId val="1002"/>'
        '</c:barChart>'
        '<c:catAx>'
        '<c:axId val="1001"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:ln></c:spPr>'
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1050" b="1"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
        '<c:crossAx val="1002"/><c:crosses val="autoZero"/>'
        '</c:catAx>'
        '<c:valAx>'
        '<c:axId val="1002"/>'
        '<c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="110"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/>'
        '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="334155"><a:alpha val="55000"/></a:srgbClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
        '<c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="950"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
        '<c:crossAx val="1001"/><c:crosses val="autoZero"/>'
        '<c:majorUnit val="25"/>'
        '</c:valAx>'
        '</c:plotArea>'
        '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1050"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
        '</c:legend>'
        '<c:plotVisOnly val="1"/>'
        '</c:chart>'
        f'<c:externalData r:id="{xlsx_rel_id}"><c:autoUpdate val="0"/></c:externalData>'
        '</c:chartSpace>'
    )


class SlideBuilder:
    def __init__(self, bg_color: str = "0F172A", morph: bool = True, morph_option: str = "byObject"):
        self.bg_color = bg_color
        self.morph = morph
        self.morph_option = morph_option
        self.shapes: list[str] = []
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
        self.shapes.append(xml_str)

    def to_slide_xml(self) -> str:
        trans = ""
        if self.morph:
            trans = (
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
                '<p:transition spd="med" advClick="1">'
                f'<p159:morph option="{self.morph_option}"/>'
                '</p:transition>'
                '</mc:Choice>'
                '<mc:Fallback>'
                '<p:transition spd="med" advClick="1"><p:fade/></p:transition>'
                '</mc:Fallback>'
                '</mc:AlternateContent>'
            )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld>'
            f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{self.bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'{"".join(self.shapes)}'
            '</p:spTree>'
            '</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            f'{trans}'
            '</p:sld>'
        )

    def to_rels_xml(self) -> str:
        items = "".join(
            f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
            for rid, rtype, target in self.rels
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'{items}'
            '</Relationships>'
        )


def write_pptx(out_path: str, slides: list[SlideBuilder], media_files: dict[str, bytes],
               charts: dict[str, tuple[str, bytes]]):
    """Package all slides, theme, master, media, and charts into a valid .pptx."""
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

    for i in range(1, len(slides) + 1):
        rid = f"rId{5 + i}"
        sld_ids.append(f'<p:sldId id="{255 + i}" r:id="{rid}"/>')
        pres_rels.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        )
        ct_overrides.append(
            f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )

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
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        f'{"".join(ct_overrides)}'
        '</Types>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )
    presentation_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{"".join(sld_ids)}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}"/>'
        f'<p:notesSz cx="{SLIDE_H}" cy="{SLIDE_W}"/>'
        '</p:presentation>'
    )
    pres_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(pres_rels)}'
        '</Relationships>'
    )
    slide_master = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '</p:sldMaster>'
    )
    slide_master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>'
    )
    slide_layout = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
        '<p:cSld name="Blank"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>'
    )
    slide_layout_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>'
    )
    theme1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TealOrbit">'
        '<a:themeElements>'
        '<a:clrScheme name="TealOrbit">'
        '<a:dk1><a:srgbClr val="0F172A"/></a:dk1>'
        '<a:lt1><a:srgbClr val="F8FAFC"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1E293B"/></a:dk2>'
        '<a:lt2><a:srgbClr val="CBD5E1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="38DCFC"/></a:accent1>'
        '<a:accent2><a:srgbClr val="00A8B5"/></a:accent2>'
        '<a:accent3><a:srgbClr val="10B981"/></a:accent3>'
        '<a:accent4><a:srgbClr val="F59E0B"/></a:accent4>'
        '<a:accent5><a:srgbClr val="60A5FA"/></a:accent5>'
        '<a:accent6><a:srgbClr val="334155"/></a:accent6>'
        '<a:hlink><a:srgbClr val="38DCFC"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="00A8B5"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="StandardMultilingual">'
        '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Noto Sans Thai"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Noto Sans Thai"/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="TealOrbit">'
        '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '</a:theme>'
    )
    pres_props = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presProps xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    view_props = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewProps xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    table_styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
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

        for chart_idx, (chart_name, (chart_xml, xlsx_bytes)) in enumerate(charts.items(), start=1):
            xlsx_name = f"Microsoft_Excel_Worksheet{chart_idx}.xlsx"
            chart_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/{xlsx_name}"/>'
                '</Relationships>'
            )
            zf.writestr(f"ppt/charts/{chart_name}", chart_xml)
            zf.writestr(f"ppt/charts/_rels/{chart_name}.rels", chart_rels)
            zf.writestr(f"ppt/embeddings/{xlsx_name}", xlsx_bytes)
