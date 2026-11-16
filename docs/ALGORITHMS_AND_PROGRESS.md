# ph-present — สถาปัตยกรรม อัลกอริทึม และความคืบหน้า (vNext)

เอกสารฉบับนี้บันทึกจุดมุ่งหมาย เป้าหมาย วิธีการทำงาน อัลกอริทึมที่พัฒนาแล้ว และความคืบหน้าล่าสุดของสกิล **`ph-present`** (Senior Presentation Architect & PowerPoint Morph OOXML Compiler)

---

## 1. จุดมุ่งหมายและเป้าหมาย (Intent & Goals)

ยกระดับ `ph-present` จากชุดคำสั่ง/เช็กลิสต์ที่ยึดติดกับเทมเพลตเดิม ๆ ไปสู่ **Presentation Architecture & Zero-Dependency OOXML Compiler** ที่สร้างไฟล์ `.pptx` แบบแก้ไขได้ 100% บนมาตรฐาน PowerPoint Desktop (Offline) และ PowerPoint Web (Online) โดยมีเป้าหมายหลัก 5 ข้อ:

1. **Natural Smart Intake:** ไม่ซักถามผู้ใช้ซ้ำหลายรอบ ถามเฉพาะสิ่งที่จำเป็นจริง ๆ (หรือรับไฟล์อ้างอิง/วิดีโอแล้วเดินหน้าต่อทันที)
2. **Zero Content Loss (Structural Contract):** บังคับสรุป `SlideContentInventory` (`core_thesis`, `proof_points`, `takeaway`, `unverified_metrics`) ก่อนวางเรขาคณิต หากเนื้อหาเกิน 1 หน้าในขนาดฟอนต์ `>= 9.5–11 pt` ให้แตกเป็นสไลด์ Morph ต่อเนื่องแทนการย่อฟอนต์หรือตัดเนื้อหาทิ้ง
3. **Unrestricted Visual Range + Benchmark Elevation:** รองรับทั้ง **Benchmark Mode** (ค้นหา reference ยอดนิยมแล้วสรุป tokens/โครงสร้างลง `references/benchmarks.md`) และ **Fast Pattern Mode** โดยไม่ล็อกทุกเด็คให้อยู่ในธีมเดียว
4. **Metrics-Driven Zero-Collision Layout (Phase 1):** คำนวณความกว้าง-สูงข้อความจริงจาก TTF, ตัดคำไทย/ลาวระดับพยางค์/คำ, แยก overflow อัตโนมัติ, ตรวจ AABB collision และสัดส่วนพื้นที่ว่าง (whitespace budget)
5. **Full PowerPoint Engine (Phase 2):** รองรับ `<p:timing>` post-Morph stagger, Compound-Path Portal Masks, Radial Glow, Multi-Chart (`Clustered`, `Stacked`, `Line`, `Doughnut`), `<a:scene3d>` / `<a:sp3d>`, และการฝังฟอนต์ `.odttf` พร้อม Target Profile Gating (`desktop` vs `web`)

---

## 2. วิธีการทำงาน 4 ขั้นตอน (Workflow)

1. **Step 1 — Smart Intake Gate:** ประเมินว่าผู้ใช้ให้ Visual Reference/Style มาแล้ว (ไป Step 2 ทันที) หรือให้มาเฉพาะหัวข้อ/เนื้อหา (ถามสั้น ๆ รอบเดียวเรื่อง Landing/Morph และ Target Desktop vs. Web)
2. **Step 2 — Content Inventory & Benchmark Elevation:** สร้าง `SlideContentInventory` (ห้ามสร้างตัวเลขสถิติสมมติโดยไม่มีแหล่งอ้างอิง), เลือก 1–2 Primary Motion Rigs จาก 7 Rigs (`Parallax Depth`, `Off-Canvas Track`, `Depth Card Stack`, `Crop-to-Pan`, `Split-Curtain / SVG Mask`, `Portal Zoom`, `Rotary Dial`) และผูก `!!anchor` ข้ามสไลด์
3. **Step 3 — Solve Layout First, Then Compile Native OOXML:** รัน `layout_solver.py` ก่อนประกอบไฟล์ `.pptx` ด้วย `pptx_morph_kit.py`
4. **Step 4 — Audit & Visual Contact Sheet Verification:** รัน `render_deck.py` เพื่อตรวจโครงสร้าง XML/Relationships, Duplicate `!!`, Morph Continuity, Timing `spid`, Font/3D Target Policy, Overflow และสร้าง `contact_sheet.png` เพื่อตรวจด้วยสายตา

---

## 3. อัลกอริทึมและโมดูลที่พัฒนาแล้ว

### 3.1 Thai/Lao Word & Orthographic Syllable Segmentation (`skill/scripts/layout_solver.py`)
- **Dynamic Programming Segmentation (`_dp_segment_run`):** ผสมพจนานุกรมคำศัพท์ความถี่สูง (`_COMMON_WORDS`) กับหน่วยพยางค์เชิงอักขรวิธี (`_orthographic_syllables`) เพื่อคำนวณต้นทุนการแบ่งคำที่ดีที่สุด (`INF` fallback ปลอดภัยไม่ตัดกลางคลัสเตอร์สระ/วรรณยุกต์)
- **Orthographic Syllable Rules (`_orthographic_syllables`):**
  - รวมอักษรฐานกับสระประสม/วรรณยุกต์/ทัณฑฆาต (`split_clusters`)
  - จับคู่สระนำ (`เ แ โ ไ ใ` / `ເ ແ ໂ ໄ ໃ`) กับพยัญชนะต้น และคำควบกล้ำแท้ (`_CLUSTER_FIRSTS` + `_CLUSTER_SECONDS` เช่น `ใกล้`, `เพลง`, `ไตร`)
  - ผูกพยัญชนะตัวสะกดท้ายพยางค์เข้ากับพยางค์ที่ยังไม่มีตัวสะกด (เช่น `ไทย`, `ใคร`, `โลก`, `งาน`, `ເມືອງລາວ`)
- **Punctuation Grouping:** ผูกเครื่องหมายไม่ตัดหน้า (`_NO_BREAK_BEFORE`) และไม่ตัดหลัง (`_NO_BREAK_AFTER`) เช่น `ลดค่าใช้จ่าย`, `(บาท)`

### 3.2 TTF Metric Box-Fit, Vertical Stack & Cross-Layer Collision (`skill/scripts/layout_solver.py`)
- **Exact TTF Measuring (`measure_token_inches`, `measure_text_block`):** ใช้ Pillow `ImageFont.getlength()` บนฟอนต์จริง (`NotoSansThai`, `NotoSansLao`, `LiberationSans`, `DejaVuSansMono`) แทนการเดาความกว้างจากจำนวนตัวอักษร
- **Overflow Routing (`solve_vertical_stack`):** เรียงการ์ดแนวตั้งจนถึง `max_bottom_in` แล้วส่งคืน `overflow_ids` สำหรับแตกเป็นสไลด์ Morph ถัดไปโดยไม่ลดขนาดฟอนต์
- **AABB Collision & Whitespace Grid (`check_aabb_collisions`, `compute_whitespace_ratio`):** ตรวจการซ้อนทับในพื้นที่สไลด์ทั้งระหว่างกล่องชั้นเดียวกันและข้ามชั้น `content`/`container` ที่ไม่ได้เป็น `parent` กัน พร้อมคำนวณสัดส่วนพื้นที่ว่างบนกริด `100 x 56`

### 3.3 SVG Cubic/Quad/Arc Compiler & Compound-Path Portal Winding (`skill/scripts/pptx_morph_kit.py`)
- **`svg_path_to_ooxml_path`:** แปลงคำสั่ง SVG Path (`M/L/H/V/C/S/Q/T/A/Z`) รวมถึง Elliptical Arc (`_arc_to_cubics` -> `<a:cubicBezTo>`) เป็น `<a:custGeom>` ของ PowerPoint แบบ vector 100%
- **Compound-Path Portal Masks (`make_portal_mask`):** สร้าง `<a:path>` เดียวที่มี 2 subpaths: กรอบสี่เหลี่ยมด้านนอก (CW) + ช่องเจาะด้านในทวนเข็มนาฬิกา (`_ccw_ellipse_svg`, `_ccw_roundrect_svg`, `rect`, หรือ `svg`) เพื่อให้ PowerPoint เจาะช่องโปร่งใสได้ในรูปทรงเดียวและขยายซูมด้วย Morph ได้ลื่นไหล (Rig F)
- **Seamless Split-Curtain (`build_curtain_pair`):** คำนวณเส้นรอยต่อกลางเดียวกันสำหรับม่านซ้าย-ขวา

### 3.4 Radial Glow, `<p:timing>`, `<a:scene3d>`, Multi-Chart & `.odttf` Font Embedding (`skill/scripts/pptx_morph_kit.py`)
- **Radial Glow (`make_radial_glow`, `fill_xml`):** ใช้ `<a:gradFill><a:path path="circle"><a:fillToRect .../>` แบบ native vector แทนภาพเบลอ PNG
- **Post-Morph Stagger (`SlideBuilder.add_stagger`, `build_timing_xml`):** สร้าง `<p:timing><p:tnLst>...<p:seq nodeType="mainSeq">` พร้อม `<p:set>` (`style.visibility = visible`) และ `<p:animEffect>` (`fade`, `wipe(...)`, `fly(...)`) ตามหลัง `<p:transition>` อย่างถูกต้องตาม schema `<p:sld>`
- **3D Camera + Target Gating (`scene3d_xml`, `sp3d_xml`, `SlideBuilder(..., target=...)`):** วาง `<a:scene3d>` / `<a:sp3d>` ตามลำดับ `<p:spPr>` ที่ PowerPoint ตรวจสอบ และตัดออกอัตโนมัติเมื่อ `target="web"`
- **Multi-Chart + Embedded `.xlsx` (`build_clustered_bar_chart_xml`, `build_stacked_bar_chart_xml`, `build_line_chart_xml`, `build_doughnut_chart_xml`):** สร้างกราฟ native `<c:chart>` ผูก `<a:latin>`, `<a:ea>`, `<a:cs>` และฝังเวิร์กบุ๊ก `.xlsx` จริง (`build_minimal_xlsx`)
- **ECMA-376 `.odttf` Font Embedding (`obfuscate_odttf`, `write_pptx`):** ทำ XOR obfuscation 32 ไบต์แรกด้วย reversed-field GUID ลง `ppt/fonts/{GUID}.odttf` พร้อม `<p:embeddedFontLst>` และ `embedTrueTypeFonts="1"`

### 3.5 OOXML Audit & Script-Aware Fallback Renderer (`skill/scripts/render_deck.py`)
- **Audit Gate (`audit_deck`):** ตรวจ XML parse, `.rels` targets, embedded `.xlsx`, duplicate `!!` anchors (รวมถึงภายใน `<p:grpSp>`), Morph continuity, `<p:timing>` target `spid`, `<a:scene3d>` / `<p:embeddedFontLst>` target warnings (`--target web|desktop`), text `< 9.0 pt`, และ post-build overflow เฉพาะองค์ประกอบที่อยู่ในเฟรมสไลด์
- **Fallback OOXML-to-PNG Renderer (`render_ooxml_slides`):** แยกอักขระตามสคริปต์ (`_split_render_runs`) ก่อนเลือก TTF (แก้ปัญหาเครื่องหมายวรรคตอน/ตัวเลขในข้อความไทย/ลาวกลายเป็น tofu), รองรับ compound-path mask (`ImageChops.difference`), radial glow, `<p:pic>` rotation/`srcRect`, และกราฟ Bar/Stacked/Line/Doughnut

---

## 4. ความคืบหน้า & ผลการทดสอบ

| ระยะ | สถานะ | รายละเอียด |
|---|---|---|
| **Phase 1: Core vNext Foundation** | **เสร็จสมบูรณ์ & แก้ไขบั๊กแล้ว** | อัปเดต `SKILL.md`, `references/`, `layout_solver.py`, `pptx_morph_kit.py`, `render_deck.py`; แก้บั๊กการตัดคำไทย/ลาว (สระนำ+ตัวสะกด+วงเล็บ), cross-layer AABB, false-positive overflow (<11 pt / off-canvas แนวตั้ง), `<a:rPr>` null guard, และ tofu บนวรรคตอนผสมภาษา |
| **Phase 2: Full PowerPoint Engine** | **เสร็จสมบูรณ์ (Local Verified)** | เพิ่ม Radial Glow, Compound-Path Portal Masks, `<p:timing>` Stagger, Multi-Chart (`Clustered`, `Stacked`, `Line`, `Doughnut`), `<a:scene3d>`/`<a:sp3d>`, และ `.odttf` Font Embedding พร้อม `target="desktop"\|"web"` gating |
| **Automated Test Suite** | **ผ่าน 10/10 tests** | `python3 -m unittest discover -s skill/scripts -v` ผ่านทุกเคส (0 errors) |

> **หมายเหตุสภาพแวดล้อม:** เครื่อง Linux ที่ใช้ทดสอบไม่มี LibreOffice (`soffice`) และ MS PowerPoint การตรวจภาพใช้ตัวเรนเดอร์ OOXML-to-PNG fallback ในตัว; การเล่นจริงของ Morph animation, `<p:timing>`, `<a:scene3d>`, และการโหลด `.odttf` ควรเปิดตรวจซ้ำใน MS PowerPoint Desktop/Web
