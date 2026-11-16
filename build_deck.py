#!/usr/bin/env python3
"""Build the 7-slide ph-present Morph showcase deck: Inside the Agent Loop."""

import math
import os
from ooxml_builder import (
    SlideBuilder,
    build_clustered_bar_chart_xml,
    build_minimal_xlsx,
    make_chart_frame,
    make_custgeom_shape,
    make_picture,
    make_shape,
    paragraph_xml,
    write_pptx,
)

DIR = os.path.dirname(__file__)
HERO_PNG = os.path.join(DIR, "agent_pipeline_hero.png")
OUT_PPTX = os.path.join(DIR, "Inside_the_Agent_Loop.pptx")

# Color Palette (Teal Orbit / Navy Depth)
BG = "0F172A"
CARD = "1E293B"
CARD_ALT = "162238"
CARD_SUB = "243247"
BORDER = "334155"
CYAN = "38DCFC"
TEAL = "00A8B5"
TEAL_DARK = "0E4D5C"
EMERALD = "10B981"
AMBER = "F59E0B"
SKY = "60A5FA"
TXT_PRI = "F8FAFC"
TXT_SEC = "CBD5E1"
TXT_MUT = "94A3B8"


def canonical_wedge_points(r_in_norm: float = 285.0, r_out_norm: float = 490.0,
                           half_angle_deg: float = 17.5, steps: int = 8) -> list[tuple[int, int]]:
    """Build a single radial gear wedge pointing straight UP (12 o'clock) in [0..1000, 0..1000]."""
    cx, cy = 500.0, 500.0
    pts = []
    for i in range(steps + 1):
        a = math.radians(-90.0 - half_angle_deg + (2.0 * half_angle_deg) * (i / steps))
        pts.append((int(round(cx + r_out_norm * math.cos(a))), int(round(cy + r_out_norm * math.sin(a)))))
    for i in range(steps, -1, -1):
        a = math.radians(-90.0 - half_angle_deg + (2.0 * half_angle_deg) * (i / steps))
        pts.append((int(round(cx + r_in_norm * math.cos(a))), int(round(cy + r_in_norm * math.sin(a)))))
    return pts


WEDGE_PTS = canonical_wedge_points()


def add_orbit_and_gear(sb: SlideBuilder, cx: float, cy: float, gear_r: float, base_rot: float,
                       orbit_r_out: float, orbit_r_in: float,
                       arm_len: float, arm_thick: float, arm_rot: float):
    """Add persistent !!orbit_ring_*, !!x_arm_*, !!gear_center, and !!wedge_0..7."""
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!orbit_ring_outer",
        cx - orbit_r_out, cy - orbit_r_out, 2 * orbit_r_out, 2 * orbit_r_out,
        prst="ellipse", rot_deg=base_rot * 0.4,
        fill=None, line=TEAL, line_pt=1.0, line_dash="dash", line_alpha=45
    ))
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!orbit_ring_inner",
        cx - orbit_r_in, cy - orbit_r_in, 2 * orbit_r_in, 2 * orbit_r_in,
        prst="ellipse", rot_deg=-base_rot * 0.6,
        fill=None, line=CYAN, line_pt=1.2, line_dash="dot", line_alpha=55
    ))

    # Crossed X-arm chrome centered cleanly behind the gear hub
    ax = cx - arm_len / 2.0
    ay = cy - arm_thick / 2.0
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!x_arm_1", ax, ay, arm_len, arm_thick,
        prst="roundRect", adj=50000, rot_deg=arm_rot,
        fill=TEAL_DARK, fill_alpha=55, line=TEAL, line_pt=0.75, line_alpha=60
    ))
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!x_arm_2", ax, ay, arm_len, arm_thick,
        prst="roundRect", adj=50000, rot_deg=-arm_rot,
        fill=TEAL_DARK, fill_alpha=55, line=CYAN, line_pt=0.75, line_alpha=60
    ))

    colors = [CYAN, TEAL, SKY, EMERALD, CYAN, TEAL, SKY, EMERALD]
    for i in range(8):
        angle = i * 45.0 + base_rot
        sb.add_shape_xml(make_custgeom_shape(
            sb.next_id(), f"!!wedge_{i}",
            cx - gear_r, cy - gear_r, 2 * gear_r, 2 * gear_r,
            WEDGE_PTS, 1000, 1000,
            rot_deg=angle, fill=colors[i], fill_alpha=88 if i % 2 == 0 else 65,
            line=BG, line_pt=0.5
        ))

    hub_r = gear_r * 0.34
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!gear_center",
        cx - hub_r, cy - hub_r, 2 * hub_r, 2 * hub_r,
        prst="ellipse", rot_deg=base_rot,
        fill=CARD, line=CYAN, line_pt=2.0, shadow=True
    ))


def add_progress_bar(sb: SlideBuilder, step: int, total: int = 7):
    track_x, track_y, track_w, track_h = 0.80, 5.30, 8.40, 0.12
    fill_w = max(0.75, track_w * (step / total))
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!prog_track", track_x, track_y, track_w, track_h,
        prst="roundRect", adj=50000, fill=CARD, line=BORDER, line_pt=0.75
    ))
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!prog_pill", track_x, track_y, fill_w, track_h,
        prst="roundRect", adj=50000, fill=CYAN, line=None
    ))


def add_header_card(sb: SlideBuilder, title_en: str, title_th: str, slide_num: str):
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!cover_card", 1.42, 0.30, 8.18, 0.80,
        prst="roundRect", adj=14000, fill=CARD, line=BORDER, line_pt=1.2, shadow=True,
        paragraphs=[
            paragraph_xml(f"{slide_num}  |  {title_en}", 16.5, TXT_PRI, bold=True, space_after_pt=2),
            paragraph_xml(title_th, 11.5, CYAN, bold=False),
        ],
        pad_in=(0.22, 0.12, 0.18, 0.10), anchor="ctr"
    ))


def build_curtain_geoms():
    """Return left & right Freeform polygons sharing exact central zigzag seam points."""
    ys = [0.0, 0.70, 1.40, 2.10, 2.8125, 3.52, 4.22, 4.92, 5.625]
    x_offs = [0.0, 0.26, -0.22, 0.28, -0.24, 0.26, -0.22, 0.24, 0.0]
    seam = [(5.0 + dx, y) for dx, y in zip(x_offs, ys)]

    left_pts = [(0, 0)] + [(int(round(sx * 1000)), int(round(sy * 1000))) for sx, sy in seam] + [(0, 5625)]

    right_origin = 4.76
    right_pts = (
        [(int(round((seam[0][0] - right_origin) * 1000)), 0), (5240, 0), (5240, 5625)]
        + [(int(round((sx - right_origin) * 1000)), int(round(sy * 1000))) for sx, sy in reversed(seam)]
    )
    return left_pts, right_pts


LEFT_CURTAIN_PTS, RIGHT_CURTAIN_PTS = build_curtain_geoms()


def make_slide_1() -> SlideBuilder:
    """Slide 1: Zero-Delay 1-Click Intro & Orbit Gear Mechanism."""
    sb = SlideBuilder(BG, morph=False)

    # Compact !!cover_card tucked behind center gear (bursts open on Click 1!)
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!cover_card", 4.35, 1.66, 1.30, 0.64,
        prst="roundRect", adj=18000, fill=CARD, line=TEAL, line_pt=1.0, shadow=True
    ))

    # Central orbit & 8-wedge rotary gear positioned clear above the title block (bottom at 1.98 + 1.66 = 3.64)
    add_orbit_and_gear(
        sb, cx=5.0, cy=1.98, gear_r=1.05, base_rot=0.0,
        orbit_r_out=1.66, orbit_r_in=1.30,
        arm_len=2.85, arm_thick=0.11, arm_rot=45.0
    )

    # Subtle !!tech_card anchor tucked behind gear center for smooth Slide 2 expansion
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!tech_card", 4.65, 1.83, 0.70, 0.30,
        prst="roundRect", adj=3500, fill=CARD_ALT, fill_alpha=5, line=None
    ))

    sb.add_shape_xml(make_shape(
        sb.next_id(), "cover_title_block", 0.90, 3.82, 8.20, 1.28,
        prst="rect", fill=None, line=None,
        paragraphs=[
            paragraph_xml("Inside the Agent Loop", 30.0, TXT_PRI, bold=True, align="ctr", space_after_pt=4),
            paragraph_xml("เบื้องหลังวิธีคิด การวางแผน และการใช้เครื่องมือของ AI", 15.0, CYAN, bold=False, align="ctr", space_after_pt=6),
            paragraph_xml("PH-PRESENT MORPH SHOWCASE  •  1-CLICK ROTARY EXPANSION", 10.5, TXT_MUT, bold=True, align="ctr"),
        ],
        pad_in=(0.10, 0.04, 0.10, 0.04)
    ))

    add_progress_bar(sb, 1)
    return sb


def make_slide_2() -> SlideBuilder:
    """Slide 2: Gear Spins 135 deg + !!cover_card Expands + 4-Stage Bento Grid."""
    sb = SlideBuilder(BG, morph=True)

    add_orbit_and_gear(
        sb, cx=0.72, cy=0.70, gear_r=0.40, base_rot=135.0,
        orbit_r_out=0.62, orbit_r_in=0.49,
        arm_len=1.08, arm_thick=0.05, arm_rot=45.0
    )
    add_header_card(sb, "01  The 4-Stage Agent Loop", "วงจรการทำงาน 4 ขั้นตอนหลักจากคำสั่งสู่ผลลัพธ์ที่ตรวจสอบแล้ว", "SLIDE 02")

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!tech_card", 0.44, 1.28, 9.16, 3.82,
        prst="roundRect", adj=3500, fill=CARD_ALT, line=BORDER, line_pt=1.0, shadow=True
    ))

    cards = [
        ("1. Perceive  •  รับรู้บริบท", CYAN,
         "อ่านไฟล์ โครงสร้างโปรเจกต์ ฟอนต์ที่มี และข้อจำกัดของสภาพแวดล้อมให้ครบก่อนเริ่มวางแผน"),
        ("2. Plan  •  คิดและวางแผน", TEAL,
         "แตกงานใหญ่เป็นขั้นตอนย่อย ออกแบบสถาปัตยกรรมและเลือกเครื่องมือที่ปลอดภัยและตรงจุด"),
        ("3. Act  •  ลงมือผ่านเครื่องมือ", SKY,
         "เขียนโค้ด รันคำสั่งใน Sandbox สร้างไฟล์ Vector/Chart หรือเรียก Subagent ช่วยทำงานขนาน"),
        ("4. Verify  •  ตรวจผลและปรับแก้", EMERALD,
         "รันตัวตรวจสอบ XML เรนเดอร์ภาพจริง และตรวจข้อความซ้อนทับก่อนส่งมอบงานทุกครั้ง"),
    ]

    w, h = 4.30, 1.64
    positions = [(0.66, 1.48), (5.08, 1.48), (0.66, 3.28), (5.08, 3.28)]
    for idx, ((x, y), (title, accent, body)) in enumerate(zip(positions, cards)):
        sb.add_shape_xml(make_shape(
            sb.next_id(), f"stage_card_{idx}", x, y, w, h,
            prst="roundRect", adj=6000, fill=CARD, line=BORDER, line_pt=1.0, shadow=True,
            paragraphs=[
                paragraph_xml(title, 15.0, accent, bold=True, space_after_pt=6),
                paragraph_xml(body, 11.5, TXT_SEC, bold=False, line_spacing=1.28),
            ],
            pad_in=(0.24, 0.16, 0.18, 0.14)
        ))
        sb.add_shape_xml(make_shape(
            sb.next_id(), f"stage_bar_{idx}", x + 0.08, y + 0.20, 0.07, h - 0.40,
            prst="roundRect", adj=50000, fill=accent, line=None
        ))

    add_progress_bar(sb, 2)
    return sb


def make_slide_3() -> SlideBuilder:
    """Slide 3: 3-Cluster Bento System Architecture (Dense Technical Diagram Rule)."""
    sb = SlideBuilder(BG, morph=True)

    add_orbit_and_gear(
        sb, cx=0.72, cy=0.70, gear_r=0.40, base_rot=180.0,
        orbit_r_out=0.62, orbit_r_in=0.49,
        arm_len=1.08, arm_thick=0.05, arm_rot=90.0
    )
    add_header_card(sb, "02  System Architecture (3-Cluster Bento)", "จัดกลุ่มระบบซับซ้อนเป็น Bento 3 โซน อ่านง่ายบนโปรเจกเตอร์ (10.5–11.5 pt)", "SLIDE 03")

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!tech_card", 0.44, 1.28, 9.16, 3.82,
        prst="roundRect", adj=3500, fill=CARD_ALT, line=BORDER, line_pt=1.0, shadow=True
    ))

    clusters = [
        ("CLUSTER A", "Context & Memory", CYAN, [
            ("System & Rules", "ข้อกำหนดความปลอดภัยและบทบาท"),
            ("Skill Instructions", "โหลด SKILL.md เฉพาะงานที่เรียกใช้"),
            ("Workspace Probe", "ตรวจไฟล์ ฟอนต์ และไบนารีจริง"),
        ]),
        ("CLUSTER B", "Reasoning Core", TEAL, [
            ("Task Decomposition", "จัดลำดับงาน ลดขั้นตอนซ้ำซ้อน"),
            ("Tool Schema Binding", "เลือกพารามิเตอร์ให้ตรงสัญญา API"),
            ("Recovery Loop", "อ่าน Error จริงแล้วแก้เฉพาะจุด"),
        ]),
        ("CLUSTER C", "Execution Sandbox", EMERALD, [
            ("Sandboxed Shell", "รันเทสต์และสคริปต์ในขอบเขตปลอดภัย"),
            ("Native Builders", "สร้าง OOXML, Chart, Freeform แท้"),
            ("Visual Gate", "เรนเดอร์ภาพจริงตรวจ Collision"),
        ]),
    ]

    xs = [0.64, 3.66, 6.68]
    cw, ch = 2.72, 3.42
    for c_idx, (cx, (tag, title, accent, items)) in enumerate(zip(xs, clusters)):
        sb.add_shape_xml(make_shape(
            sb.next_id(), f"bento_cluster_{c_idx}", cx, 1.48, cw, ch,
            prst="roundRect", adj=4500, fill=CARD, line=accent, line_pt=1.2, shadow=True,
            paragraphs=[
                paragraph_xml(tag, 9.5, accent, bold=True, mono=True, space_after_pt=2),
                paragraph_xml(title, 14.5, TXT_PRI, bold=True),
            ],
            pad_in=(0.16, 0.14, 0.14, 0.10)
        ))
        for i_idx, (sub_t, sub_b) in enumerate(items):
            sy = 2.18 + i_idx * 0.88
            sb.add_shape_xml(make_shape(
                sb.next_id(), f"bento_node_{c_idx}_{i_idx}", cx + 0.14, sy, cw - 0.28, 0.78,
                prst="roundRect", adj=8000, fill=CARD_SUB, line=BORDER, line_pt=0.75,
                paragraphs=[
                    paragraph_xml(sub_t, 11.0, accent, bold=True, space_after_pt=2),
                    paragraph_xml(sub_b, 10.0, TXT_SEC, line_spacing=1.18),
                ],
                pad_in=(0.12, 0.08, 0.12, 0.08)
            ))

    for ax in [3.39, 6.41]:
        sb.add_shape_xml(make_shape(
            sb.next_id(), f"cluster_arrow_{ax}", ax, 3.05, 0.24, 0.28,
            prst="rightArrow", fill=CYAN, line=None
        ))

    add_progress_bar(sb, 3)
    return sb


def make_slide_4() -> SlideBuilder:
    """Slide 4: Native Editable Chart (<c:chart> + .xlsx) + Clear KPI Callout Badges."""
    sb = SlideBuilder(BG, morph=True)

    add_orbit_and_gear(
        sb, cx=0.72, cy=0.70, gear_r=0.40, base_rot=225.0,
        orbit_r_out=0.62, orbit_r_in=0.49,
        arm_len=1.08, arm_thick=0.05, arm_rot=135.0
    )
    add_header_card(sb, "03  Measured Verification Impact", "กราฟ Native แก้ไขข้อมูลใน PowerPoint ได้จริง (Edit Data) + KPI แยกนอกแกนกราฟ", "SLIDE 04")

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!tech_card", 0.44, 1.28, 6.00, 3.82,
        prst="roundRect", adj=3500, fill=CARD, line=BORDER, line_pt=1.2, shadow=True
    ))
    chart_rid = sb.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart1.xml")
    sb.add_shape_xml(make_chart_frame(
        sb.next_id(), "impact_chart", chart_rid, 0.58, 1.40, 5.72, 3.58
    ))

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!kpi_badge_1", 6.64, 1.28, 2.96, 1.80,
        prst="roundRect", adj=6000, fill=CARD, line=CYAN, line_pt=1.5, shadow=True,
        paragraphs=[
            paragraph_xml("BUILD PASS RATE", 10.0, TXT_MUT, bold=True, mono=True, space_after_pt=4),
            paragraph_xml("+38%", 30.0, CYAN, bold=True, space_after_pt=4),
            paragraph_xml("อัตราผ่านการตรวจ XML & โค้ดตั้งแต่รอบแรก\nเมื่อเทียบกับ Prompt เดียว", 11.0, TXT_SEC, line_spacing=1.26),
        ],
        pad_in=(0.20, 0.16, 0.18, 0.14)
    ))
    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!kpi_badge_2", 6.64, 3.26, 2.96, 1.84,
        prst="roundRect", adj=6000, fill=CARD, line=EMERALD, line_pt=1.5, shadow=True,
        paragraphs=[
            paragraph_xml("COLLISION REDUCTION", 10.0, TXT_MUT, bold=True, mono=True, space_after_pt=4),
            paragraph_xml("-89%", 30.0, EMERALD, bold=True, space_after_pt=4),
            paragraph_xml("ลดปัญหาตัวอักษรซ้อนทับหรือตกขอบ\nจากการตรวจ Contact Sheet จริง", 11.0, TXT_SEC, line_spacing=1.26),
        ],
        pad_in=(0.20, 0.16, 0.18, 0.14)
    ))

    add_progress_bar(sb, 4)
    return sb


def make_slide_5() -> SlideBuilder:
    """Slide 5: Split-Curtain Closed Beat + Full Hero Visual Underneath."""
    sb = SlideBuilder(BG, morph=True)

    img_rid = sb.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "../media/agent_pipeline_hero.png")
    sb.add_shape_xml(make_picture(
        sb.next_id(), "!!hero_visual", img_rid, 0.75, 1.25, 8.50, 3.80,
        src_rect=(0, 0, 0, 0), line=BORDER, line_pt=1.0
    ))

    sb.add_shape_xml(make_custgeom_shape(
        sb.next_id(), "!!curtain_left", 0.0, 0.0, 5.28, 5.625,
        LEFT_CURTAIN_PTS, 5280, 5625, fill="132238", line=TEAL, line_pt=1.5, shadow=True
    ))
    sb.add_shape_xml(make_custgeom_shape(
        sb.next_id(), "!!curtain_right", 4.76, 0.0, 5.24, 5.625,
        RIGHT_CURTAIN_PTS, 5240, 5625, fill="172B45", line=CYAN, line_pt=1.5, shadow=True
    ))

    add_orbit_and_gear(
        sb, cx=0.72, cy=0.70, gear_r=0.40, base_rot=248.0,
        orbit_r_out=0.62, orbit_r_in=0.49,
        arm_len=1.08, arm_thick=0.05, arm_rot=158.0
    )
    add_header_card(sb, "04  The Golden Rule", "จังหวะปิดม่าน Freeform ก่อนเปิดเผย Verification Gate", "SLIDE 05")

    sb.add_shape_xml(make_shape(
        sb.next_id(), "curtain_beat_card", 1.75, 1.95, 6.50, 1.90,
        prst="roundRect", adj=7000, fill=BG, fill_alpha=96, line=CYAN, line_pt=1.5, shadow=True,
        paragraphs=[
            paragraph_xml("STAGE 04  •  VERIFICATION GATE", 11.0, EMERALD, bold=True, mono=True, align="ctr", space_after_pt=6),
            paragraph_xml("Verify Before Completion", 26.0, TXT_PRI, bold=True, align="ctr", space_after_pt=6),
            paragraph_xml("อย่าเชื่อแค่คำตอบที่ดูดี — พิสูจน์ด้วยโครงสร้างและภาพจริง", 13.5, CYAN, align="ctr"),
        ],
        pad_in=(0.20, 0.16, 0.20, 0.16), anchor="ctr"
    ))

    add_progress_bar(sb, 5)
    return sb


def make_slide_6() -> SlideBuilder:
    """Slide 6: Curtain Splits Apart + <a:srcRect> Crop-to-Pan Zoom into Verification Gate."""
    sb = SlideBuilder(BG, morph=True)

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!tech_card", 0.58, 1.32, 3.74, 3.74,
        prst="roundRect", adj=4500, fill=CARD, line=BORDER, line_pt=1.2, shadow=True,
        paragraphs=[
            paragraph_xml("AUDIT CHECKLIST", 10.0, EMERALD, bold=True, mono=True, space_after_pt=2),
            paragraph_xml("3 ด่านตรวจก่อนส่งงาน", 15.5, TXT_PRI, bold=True, space_after_pt=10),
            paragraph_xml("1. Package & OOXML Check", 12.0, CYAN, bold=True, space_after_pt=2),
            paragraph_xml("ตรวจความครบของ .rels, chart .xlsx\nและห้ามมีชื่อ !! ซ้ำในสไลด์เดียว", 10.8, TXT_SEC, space_after_pt=10, line_spacing=1.24),
            paragraph_xml("2. Crop-to-Pan & Layout", 12.0, TEAL, bold=True, space_after_pt=2),
            paragraph_xml("ซูมเจาะจุดสำคัญผ่าน <a:srcRect>\nโดยไม่ต้องครอปภาพแยกไฟล์", 10.8, TXT_SEC, space_after_pt=10, line_spacing=1.24),
            paragraph_xml("3. Evidence-First Report", 12.0, AMBER, bold=True, space_after_pt=2),
            paragraph_xml("รายงานผลจากข้อมูลที่รันและเปิดดูจริงเท่านั้น", 10.8, TXT_SEC, line_spacing=1.24),
        ],
        pad_in=(0.20, 0.16, 0.18, 0.14)
    ))

    img_rid = sb.add_rel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "../media/agent_pipeline_hero.png")
    sb.add_shape_xml(make_picture(
        sb.next_id(), "!!hero_visual", img_rid, 4.52, 1.32, 4.90, 3.74,
        src_rect=(49500, 15500, 3000, 12500), line=CYAN, line_pt=1.8, shadow=True
    ))

    # Split curtains parted clear of card borders (leaving ~0.36 in jagged edges on flanks)
    sb.add_shape_xml(make_custgeom_shape(
        sb.next_id(), "!!curtain_left", -4.92, 0.0, 5.28, 5.625,
        LEFT_CURTAIN_PTS, 5280, 5625, fill="132238", line=TEAL, line_pt=1.5, shadow=True
    ))
    sb.add_shape_xml(make_custgeom_shape(
        sb.next_id(), "!!curtain_right", 9.64, 0.0, 5.24, 5.625,
        RIGHT_CURTAIN_PTS, 5240, 5625, fill="172B45", line=CYAN, line_pt=1.5, shadow=True
    ))

    add_orbit_and_gear(
        sb, cx=0.72, cy=0.70, gear_r=0.40, base_rot=258.0,
        orbit_r_out=0.62, orbit_r_in=0.49,
        arm_len=1.08, arm_thick=0.05, arm_rot=168.0
    )
    add_header_card(sb, "04  Curtain Split + Crop-to-Pan Zoom", "ม่าน Freeform แยกซ้าย-ขวา พร้อมซูมภาพแผนภาพเข้าสู่ Verification Gate", "SLIDE 06")
    add_progress_bar(sb, 6)
    return sb


def make_slide_7() -> SlideBuilder:
    """Slide 7: Closing Takeaways — Gear Spins to 270 deg & Expands."""
    sb = SlideBuilder(BG, morph=True)

    add_orbit_and_gear(
        sb, cx=2.10, cy=2.68, gear_r=1.08, base_rot=270.0,
        orbit_r_out=1.72, orbit_r_in=1.34,
        arm_len=2.90, arm_thick=0.11, arm_rot=135.0
    )

    sb.add_shape_xml(make_shape(
        sb.next_id(), "!!cover_card", 3.95, 1.05, 5.55, 3.25,
        prst="roundRect", adj=4500, fill=CARD, line=CYAN, line_pt=1.5, shadow=True,
        paragraphs=[
            paragraph_xml("TAKEAWAY  •  PH-PRESENT", 10.5, EMERALD, bold=True, mono=True, space_after_pt=6),
            paragraph_xml("Great Agents Don't Guess —\nThey Verify.", 23.0, TXT_PRI, bold=True, space_after_pt=8, line_spacing=1.18),
            paragraph_xml("คิดเป็นระบบ • ลงมือด้วยเครื่องมือจริง • ตรวจภาพและโครงสร้างก่อนส่ง", 13.0, CYAN, space_after_pt=14),
            paragraph_xml("สร้างจาก OOXML 100% แก้ไข Shape, Text, Freeform และ Chart Data\nได้ครบทุกจุดใน Microsoft PowerPoint", 11.0, TXT_SEC, line_spacing=1.25),
        ],
        pad_in=(0.28, 0.24, 0.24, 0.20), anchor="ctr"
    ))

    pills = [
        ("Native Vectors", CYAN),
        ("True Morph Rotation", TEAL),
        ("Thai + Latin Fonts", EMERALD),
    ]
    for i, (label, col) in enumerate(pills):
        px = 3.95 + i * 1.88
        sb.add_shape_xml(make_shape(
            sb.next_id(), f"closing_pill_{i}", px, 4.48, 1.78, 0.42,
            prst="roundRect", adj=50000, fill=CARD_SUB, line=col, line_pt=1.0,
            paragraphs=[paragraph_xml(label, 10.0, col, bold=True, align="ctr")],
            pad_in=(0.06, 0.04, 0.06, 0.04), anchor="ctr"
        ))

    add_progress_bar(sb, 7)
    return sb


def main():
    with open(HERO_PNG, "rb") as f:
        hero_bytes = f.read()

    categories = ["Build Pass Rate", "Zero Collision", "Factual Accuracy"]
    series_data = [
        ("Single-Shot Prompt", "475569", [58, 44, 67]),
        ("Tool-Verified Agent", CYAN, [96, 94, 98]),
    ]
    xlsx_bytes = build_minimal_xlsx(
        categories,
        [(name, vals) for name, _, vals in series_data]
    )
    chart_xml = build_clustered_bar_chart_xml(
        "Single-Shot vs. Tool-Verified Agent (%)", categories, series_data
    )

    slides = [
        make_slide_1(),
        make_slide_2(),
        make_slide_3(),
        make_slide_4(),
        make_slide_5(),
        make_slide_6(),
        make_slide_7(),
    ]

    write_pptx(
        OUT_PPTX,
        slides=slides,
        media_files={"agent_pipeline_hero.png": hero_bytes},
        charts={"chart1.xml": (chart_xml, xlsx_bytes)},
    )
    print(f"Built {OUT_PPTX} ({len(slides)} slides)")


if __name__ == "__main__":
    main()
