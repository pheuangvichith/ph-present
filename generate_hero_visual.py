#!/usr/bin/env python3
"""Generate high-resolution Agent Loop schematic PNG with proper alpha compositing."""

from PIL import Image, ImageDraw, ImageFont
import os

W, H = 2400, 1350
OUT_PATH = os.path.join(os.path.dirname(__file__), "agent_pipeline_hero.png")

FONT_LATIN_BOLD = "/usr/share/fonts/liberation/LiberationSans-Bold.ttf"
FONT_LATIN_REG = "/usr/share/fonts/liberation/LiberationSans-Regular.ttf"
FONT_MONO = "/usr/share/fonts/TTF/DejaVuSansMono-Bold.ttf"


def main():
    base = Image.new("RGBA", (W, H), (15, 23, 42, 255))

    # Overlay layer for semi-transparent grid and ambient rings
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    for x in range(0, W, 80):
        od.line([(x, 0), (x, H)], fill=(51, 65, 85, 55), width=2)
    for y in range(0, H, 80):
        od.line([(0, y), (W, y)], fill=(51, 65, 85, 55), width=2)

    od.ellipse([1500, 120, 2320, 940], outline=(0, 168, 181, 65), width=4)
    od.ellipse([1620, 240, 2200, 820], outline=(56, 220, 252, 85), width=3)

    img = Image.alpha_composite(base, overlay)
    card_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(card_layer)

    f_title = ImageFont.truetype(FONT_LATIN_BOLD, 44)
    f_sub = ImageFont.truetype(FONT_LATIN_REG, 28)
    f_node_title = ImageFont.truetype(FONT_LATIN_BOLD, 34)
    f_node_body = ImageFont.truetype(FONT_LATIN_REG, 25)
    f_tag = ImageFont.truetype(FONT_MONO, 22)

    # Top schematic header
    draw.text((100, 72), "AGENT EXECUTION PIPELINE  //  INTERNAL TELEMETRY", font=f_title, fill=(248, 250, 252, 255))
    draw.text((100, 132), "Context Ingestion  ->  Structured Plan  ->  Sandboxed Tool Call  ->  Verification Gate", font=f_sub, fill=(56, 220, 252, 255))

    # Left half: Input & Planner nodes
    draw.rounded_rectangle((100, 240, 980, 620), radius=28, fill=(24, 34, 52, 245), outline=(51, 65, 85, 255), width=4)
    draw.text((145, 278), "01  CONTEXT & PROMPT PARSER", font=f_tag, fill=(56, 220, 252, 255))
    draw.text((145, 326), "Workspace & Skill Ingestion", font=f_node_title, fill=(248, 250, 252, 255))
    draw.text((145, 392), "• Reads SKILL.md rules & constraints\n• Inspects filesystem, fonts & tools\n• Establishes target schema & layout", font=f_node_body, fill=(203, 213, 225, 255), spacing=10)

    draw.rounded_rectangle((100, 700, 980, 1140), radius=28, fill=(24, 34, 52, 245), outline=(0, 168, 181, 255), width=4)
    draw.text((145, 738), "02  TOOL SANDBOX & EXECUTION", font=f_tag, fill=(0, 168, 181, 255))
    draw.text((145, 786), "Atomic Native Asset Synthesis", font=f_node_title, fill=(248, 250, 252, 255))
    draw.text((145, 852), "• Vector shapes & custGeom polygons\n• Embedded XLSX-backed OOXML charts\n• Script-aware Latin + Thai font runs\n• Deterministic <a:xfrm rot> angles", font=f_node_body, fill=(203, 213, 225, 255), spacing=10)

    # Connector lines
    draw.line([(980, 430), (1140, 430), (1140, 560), (1240, 560)], fill=(56, 220, 252, 255), width=6)
    draw.line([(980, 920), (1140, 920), (1140, 740), (1240, 740)], fill=(16, 185, 129, 255), width=6)

    # Right half: Verification Gate (Crop-to-Pan target)
    draw.rounded_rectangle((1240, 240, 2280, 1140), radius=34, fill=(20, 31, 50, 250), outline=(56, 220, 252, 255), width=5)
    draw.rounded_rectangle((1285, 285, 1675, 340), radius=14, fill=(16, 185, 129, 38), outline=(16, 185, 129, 255), width=3)
    draw.text((1315, 299), "03  VERIFICATION GATE", font=f_tag, fill=(52, 211, 153, 255))

    draw.text((1285, 368), "Zero-Trust Quality Audit", font=ImageFont.truetype(FONT_LATIN_BOLD, 42), fill=(248, 250, 252, 255))
    draw.text((1285, 426), "Every artifact is validated structurally & visually", font=f_sub, fill=(148, 163, 184, 255))

    checks = [
        ("XML & Package Integrity", "All .rels, [Content_Types], chart XLSX verified", (56, 220, 252)),
        ("Unique !! Anchor Check", "0 duplicate !! names per slide -> clean Morph", (0, 168, 181)),
        ("Bounding Box & Margin Guard", "0 text-on-text or border overlaps detected", (16, 185, 129)),
        ("Contact Sheet Inspection", "Full slide render reviewed before delivery", (245, 158, 11)),
    ]

    cy = 505
    for title, desc, accent in checks:
        draw.rounded_rectangle((1285, cy, 2235, cy + 132), radius=20, fill=(30, 41, 59, 245), outline=(51, 65, 85, 255), width=3)
        draw.rounded_rectangle((1310, cy + 28, 1324, cy + 104), radius=6, fill=accent + (255,))
        draw.text((1350, cy + 24), title, font=f_node_title, fill=(248, 250, 252, 255))
        draw.text((1350, cy + 72), desc, font=f_node_body, fill=(203, 213, 225, 255))
        draw.rounded_rectangle((2075, cy + 40, 2205, cy + 92), radius=14, fill=(16, 185, 129, 36), outline=(16, 185, 129, 255), width=2)
        draw.text((2106, cy + 52), "PASS", font=f_tag, fill=(52, 211, 153, 255))
        cy += 152

    final = Image.alpha_composite(img, card_layer)
    final.convert("RGB").save(OUT_PATH, "PNG")
    print(f"Saved {OUT_PATH}")


if __name__ == "__main__":
    main()
