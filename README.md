# ph-present — Presentation Architect & PowerPoint Morph OOXML Compiler

โปรเจกต์และสกิลสำหรับออกแบบ, ถอดโครงสร้าง (Reverse-Engineer), และคอมไพล์ไฟล์นำเสนอ **`.pptx` ที่แก้ไขได้ 100%** พร้อม Morph Transitions, Motion Rigs ระดับ PlacePlate, การจัดวางข้อความภาษาไทย/ลาว/อังกฤษแบบไม่ล้นกล่องด้วย Metrics จริงจาก TTF, และระบบตรวจ/เรนเดอร์ Contact Sheet

---

## โครงสร้าง Repository

```text
.
├── README.md                              # ภาพรวม จุดมุ่งหมาย วิธีใช้งาน และความคืบหน้า
├── docs/
│   ├── 2026-11-16-ph-present-vnext-design.md  # Design Spec (Phase 1 & Phase 2)
│   └── ALGORITHMS_AND_PROGRESS.md             # บันทึกเป้าหมาย วิธีทำ อัลกอริทึม และผลทดสอบ
├── skill/                                 # ตัวสกิล ph-present ฉบับสมบูรณ์ (vNext Phase 1+2)
│   ├── SKILL.md                           # Workflow 4 ขั้นตอนและกฎหลัก
│   ├── references/
│   │   ├── benchmarks.md                  # ตาราง Desktop vs Web & Benchmark Archetypes
│   │   └── morph-rigs-and-ooxml.md        # 7 Motion Rigs + OOXML Schema Cookbook
│   └── scripts/
│       ├── layout_solver.py               # Thai/Lao DP segmentation, TTF fit, AABB, Inventory
│       ├── pptx_morph_kit.py              # Zero-dependency OOXML + SVG + Chart + Timing + Font Kit
│       ├── render_deck.py                 # OOXML/Morph Audit + LibreOffice/Fallback PNG Renderer
│       ├── test_layout_solver.py          # Unit tests (Segmentation, Stack, AABB, Inventory)
│       └── test_ph_present.py             # Integration & Phase 2 Engine tests
├── build_deck.py                          # สคริปต์สร้างเด็คตัวอย่าง 7 สไลด์ (Inside the Agent Loop)
├── ooxml_builder.py / audit_ooxml.py / render_audit.py  # เครื่องมือชุดเดโมตั้งต้น
├── generate_hero_visual.py / agent_pipeline_hero.png    # ภาพประกอบหลักของเด็คตัวอย่าง
├── Inside_the_Agent_Loop.pptx             # ไฟล์เด็คตัวอย่าง
└── rendered_audit/                        # ภาพเรนเดอร์สไลด์และ contact_sheet.png ของเด็คตัวอย่าง
```

---

## ความสามารถหลัก (Phase 1 + Phase 2)

- **Smart Intake & Content-First Contract:** บังคับโครงสร้าง `SlideContentInventory` (`core_thesis`, `proof_points`, `takeaway`) และแตกสไลด์ Morph ต่อเนื่องแทนการลบเนื้อหาหรือลดขนาดฟอนต์ต่ำกว่ามาตรฐาน
- **Thai / Lao / Latin Layout Solver (`skill/scripts/layout_solver.py`):**
  - ตัดคำไทยและลาวด้วย DP + กฎพยางค์/คำควบกล้ำ/ตัวสะกด (`_dp_segment_run`, `_orthographic_syllables`)
  - วัดความสูง-กว้างข้อความจาก TTF จริง (`measure_text_block`, `solve_vertical_stack`)
  - ตรวจการชน AABB ข้ามเลเยอร์ (`check_aabb_collisions`) และคำนวณพื้นที่ว่าง (`compute_whitespace_ratio`)
- **Zero-Dependency OOXML Morph Engine (`skill/scripts/pptx_morph_kit.py`):**
  - `<p159:morph>` transitions + 7 Motion Rigs (Parallax, Off-Canvas Track, Depth Stack, Crop-to-Pan `<a:srcRect>`, Split-Curtain, Compound-Path Portal Mask, True Pivot Rotary Dial)
  - SVG Path (`M/L/H/V/C/S/Q/T/A/Z`) -> `<a:custGeom>` และ Compound-Path Cutouts (`make_portal_mask`)
  - Native Radial Glow (`make_radial_glow`)
  - Multi-Chart Editable XML + `.xlsx` (`Clustered`, `Stacked`, `Line`, `Doughnut`) พร้อม `<a:cs>`/`<a:ea>` สำหรับไทย/ลาว/CJK
  - `<p:timing>` Post-Morph Stagger (`SlideBuilder.add_stagger`)
  - `<a:scene3d>` / `<a:sp3d>` และ `.odttf` XOR Font Embedding (`<p:embeddedFontLst>`) พร้อม `target="desktop"|"web"` profile gating
- **Audit & Visual Verification (`skill/scripts/render_deck.py`):**
  - ตรวจ XML, `.rels`, embedded `.xlsx`, duplicate `!!` anchors, Morph continuity, `<p:timing>` `spid`, `--target web|desktop` compatibility, และ text overflow
  - เรนเดอร์ผ่าน LibreOffice + `pdftoppm` หรือ Built-in OOXML-to-PNG Fallback Renderer แล้วรวมเป็น `contact_sheet.png`

---

## การทดสอบ (Tests)

```bash
# รัน Unit & Phase 2 Integration Tests ทั้งหมด (Python stdlib + Pillow)
python3 -m unittest discover -s skill/scripts -v

# ตรวจและเรนเดอร์เด็คตัวอย่าง
python3 skill/scripts/render_deck.py Inside_the_Agent_Loop.pptx /tmp/audit_out --target desktop
```

รายละเอียดอัลกอริทึมและบันทึกความคืบหน้าทั้งหมดดูได้ที่ [docs/ALGORITHMS_AND_PROGRESS.md](docs/ALGORITHMS_AND_PROGRESS.md)
