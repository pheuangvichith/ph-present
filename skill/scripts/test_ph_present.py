#!/usr/bin/env python3
"""Integration & unit tests for ph-present Phase 1 fixes + Phase 2 engine."""

import os
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

import layout_solver as ls
import pptx_morph_kit as kit
import render_deck as rd


def _parse_sp(xml_frag: str) -> ET.Element:
    wrapped = (
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{xml_frag}</p:sld>'
    )
    return ET.fromstring(wrapped).find("p:sp", rd.NS)


class TestPhase2Engine(unittest.TestCase):
    def test_radial_glow_and_portal_mask_xml(self):
        glow_xml = kit.make_radial_glow(2, "!!glow_1", 5.0, 2.8, 2.5, 1.8, color="38DCFC")
        el = _parse_sp(glow_xml)
        self.assertIsNotNone(el.find(".//a:gradFill/a:path[@path='circle']", rd.NS))

        portal_xml = kit.make_portal_mask(3, "!!portal_mask", hole_box=(3.5, 1.3, 3.0, 3.0), hole_type="circle")
        pel = _parse_sp(portal_xml)
        path_el = pel.find(".//a:custGeom/a:pathLst/a:path", rd.NS)
        self.assertIsNotNone(path_el)
        # Outer rectangle + inner circle cutout => 2 moveTo commands
        self.assertEqual(len(path_el.findall("a:moveTo", rd.NS)), 2)

    def test_scene3d_and_web_target_gating(self):
        s3d = kit.scene3d_xml("perspectiveFront", rot_x_deg=12.0, rot_y_deg=-15.0)
        sp3d = kit.sp3d_xml(extrude_pt=4.0, bevel_pt=2.0)
        card = kit.make_shape(
            2, "!!hero_card", 1.0, 1.0, 4.0, 2.5,
            fill="1E293B", scene3d=s3d, sp3d=sp3d,
        )
        # Strict <p:spPr> child order check
        sp_pr = _parse_sp(card).find("p:spPr", rd.NS)
        tags = [c.tag.split("}")[-1] for c in list(sp_pr)]
        self.assertEqual(tags, ["xfrm", "prstGeom", "solidFill", "ln", "scene3d", "sp3d"])

        # Web target strips scene3d/sp3d automatically
        sb_web = kit.SlideBuilder(target="web", morph=False)
        sb_web.add_shape_xml(card)
        web_xml = sb_web.to_slide_xml()
        self.assertNotIn("scene3d", web_xml)
        self.assertNotIn("sp3d", web_xml)

    def test_timing_stagger_xml_and_audit(self):
        sb = kit.SlideBuilder(morph=True)
        id1 = sb.next_id()
        id2 = sb.next_id()
        sb.add_shape_xml(kit.make_shape(id1, "!!card_1", 0.5, 1.0, 4.0, 1.5, fill="1E293B"))
        sb.add_shape_xml(kit.make_shape(id2, "!!card_2", 5.0, 1.0, 4.0, 1.5, fill="1E293B"))
        sb.add_stagger(id1, effect="fade", delay_ms=100, dur_ms=350, trigger="afterPrevious")
        sb.add_stagger(id2, effect="wipe", delay_ms=150, dur_ms=400, trigger="withPrevious", direction="left")

        sld_root = ET.fromstring(sb.to_slide_xml())
        top_tags = [c.tag.split("}")[-1] for c in list(sld_root)]
        # <p:timing> must appear after <p:clrMapOvr> and <mc:AlternateContent>
        self.assertEqual(top_tags, ["cSld", "clrMapOvr", "AlternateContent", "timing"])

    def test_multichart_and_font_embedding_end_to_end(self):
        thai_reg = ls._resolve_font_path("thai", False)
        self.assertIsNotNone(thai_reg)

        with tempfile.TemporaryDirectory() as tmp:
            out_pptx = os.path.join(tmp, "phase2_showcase.pptx")
            out_dir = os.path.join(tmp, "rendered")

            sb1 = kit.SlideBuilder(bg_color="0F172A", morph=False, target="desktop")
            sb1.add_shape_xml(kit.make_radial_glow(sb1.next_id(), "!!bg_glow", 5.0, 2.8, 3.8, 2.5, "38DCFC"))
            gid = sb1.next_id()
            sb1.add_shape_xml(
                kit.make_shape(
                    gid, "!!hero_title", 0.7, 1.6, 8.6, 1.6,
                    fill="1E293B", line="38DCFC",
                    paragraphs=[kit.paragraph_xml("รายงานกลยุทธ์ ไทย-ลาว ประจำไตรมาส", 24.0, "F8FAFC", bold=True)],
                    scene3d=kit.scene3d_xml("perspectiveFront", rot_x_deg=8.0, rot_y_deg=-10.0),
                )
            )
            sb1.add_stagger(gid, effect="fade", delay_ms=120, dur_ms=400)

            sb2 = kit.SlideBuilder(bg_color="F7F6F2", morph=True, target="desktop")
            sb2.add_shape_xml(kit.make_portal_mask(sb2.next_id(), "!!portal_mask", hole_box=(0.6, 0.8, 4.0, 4.0), hole_type="roundRect", fill="0F172A"))
            # Stacked + Line + Doughnut charts
            cats = ["ม.ค.", "ก.พ.", "มี.ค."]
            xlsx1 = kit.build_minimal_xlsx(cats, [("ไทย", [40.0, 55.0, 72.0]), ("ลาว", [25.0, 32.0, 45.0])])
            c_line = kit.build_line_chart_xml("แนวโน้มรายได้", cats, [("ไทย", "059669", [40.0, 55.0, 72.0]), ("ลาว", "4F46E5", [25.0, 32.0, 45.0])], text_color="0F172A", muted_color="475569", grid_color="CBD5E1")
            rid_line = sb2.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart1.xml")
            sb2.add_shape_xml(kit.make_chart_frame(sb2.next_id(), "!!line_chart", rid_line, 4.9, 0.4, 4.7, 2.4))

            c_donut = kit.build_doughnut_chart_xml("สัดส่วนยอดขาย", ["ไทย", "ลาว", "อื่นๆ"], "Share", [58.0, 30.0, 12.0], ["059669", "4F46E5", "F59E0B"], text_color="0F172A", border_color="F7F6F2")
            rid_donut = sb2.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart2.xml")
            sb2.add_shape_xml(kit.make_chart_frame(sb2.next_id(), "!!donut_chart", rid_donut, 4.9, 3.0, 4.7, 2.3))

            xlsx2 = kit.build_minimal_xlsx(["ไทย", "ลาว", "อื่นๆ"], [("Share", [58.0, 30.0, 12.0])])

            sb3 = kit.SlideBuilder(bg_color="0F172A", morph=True, target="desktop")
            c_stack = kit.build_stacked_bar_chart_xml("เปรียบเทียบรายไตรมาส", cats, [("ออนไลน์", "38DCFC", [30.0, 45.0, 60.0]), ("หน้าร้าน", "10B981", [50.0, 40.0, 35.0])])
            rid_stack = sb3.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart3.xml")
            sb3.add_shape_xml(kit.make_chart_frame(sb3.next_id(), "!!stack_chart", rid_stack, 0.8, 0.7, 8.4, 4.3))
            xlsx3 = kit.build_minimal_xlsx(cats, [("ออนไลน์", [30.0, 45.0, 60.0]), ("หน้าร้าน", [50.0, 40.0, 35.0])])

            kit.write_pptx(
                out_pptx,
                [sb1, sb2, sb3],
                charts={
                    "chart1.xml": (c_line, xlsx1),
                    "chart2.xml": (c_donut, xlsx2),
                    "chart3.xml": (c_stack, xlsx3),
                },
                embed_fonts=[{"typeface": "Noto Sans Thai", "regular": thai_reg}],
                target="desktop",
            )

            with zipfile.ZipFile(out_pptx, "r") as zf:
                odttfs = [n for n in zf.namelist() if n.startswith("ppt/fonts/") and n.endswith(".odttf")]
                self.assertEqual(len(odttfs), 1)
                pres_xml = zf.read("ppt/presentation.xml").decode("utf-8")
                self.assertIn("embedTrueTypeFonts=\"1\"", pres_xml)
                self.assertIn("Noto Sans Thai", pres_xml)

            errs_desk, warns_desk = rd.audit_deck(out_pptx, target="desktop")
            self.assertEqual(errs_desk, [])
            self.assertEqual(warns_desk, [])

            # Auditing the same desktop deck with --target web must warn about scene3d and embeddedFontLst
            errs_web, warns_web = rd.audit_deck(out_pptx, target="web")
            self.assertEqual(errs_web, [])
            self.assertGreaterEqual(len(warns_web), 2)

            imgs = rd.render_ooxml_slides(out_pptx, out_dir)
            self.assertEqual(len(imgs), 3)
            cs_path = os.path.join(out_dir, "contact_sheet.png")
            rd.build_contact_sheet(imgs, cs_path)
            self.assertTrue(os.path.exists(cs_path))


if __name__ == "__main__":
    unittest.main()
