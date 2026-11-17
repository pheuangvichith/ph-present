---
name: ph-present
description: Use when the user mentions ph.present or ph-present, asks to reverse-engineer a slide video or template into an editable .pptx, build a projector-legible PowerPoint deck with Morph transitions or Thai/Lao typography, or upgrade an existing .pptx presentation.
---

# ph-present

Senior Presentation Architect & Algorithmic Morph Pipeline for building 100% editable, projector-legible Microsoft PowerPoint (`.pptx`) decks and upgrading existing decks. Enforces a strict **4-Gate Data Flow** so the AI handles intake, perspective preservation, and visual taste while deterministic scripts handle layout math, Morph continuity, OOXML compilation, and verification.

## Core Invariants

1. **No-Drop Content Contract:** Never shrink text below readable projector scale or drop key viewpoints/proof points to force a single slide. When content exceeds safe density, split cleanly into Morph-linked continuation slides.
2. **Data Integrity:** Never invent metrics or placeholder numbers. Mark missing quantitative inputs as `[NEED DATA]` during planning.
3. **100% Editable Native Objects:** Build cards, pills, rings, gears, tables, and charts as native PowerPoint OOXML shapes/charts (`scripts/pptx_morph_kit.py` or `python-pptx`), never flattened slide screenshots.

## Bundled Toolkit (Invoke via `python3`)

- **Motion Rigs, Deconstruction & OOXML Cookbook:** `references/morph-rigs-and-ooxml.md`
- **Benchmarks & Desktop/Web Matrix:** `references/benchmarks.md`
- **Layout & Thai/Lao Constraint Solver:** `python3 scripts/layout_solver.py` (`segment_thai_lao_tokens`, `measure_text_block`, `solve_vertical_stack`, `check_aabb_collisions`, `compute_whitespace_ratio`, `validate_content_inventory`)
- **Zero-Dependency OOXML + SVG + Chart + Timing Kit:** `python3 scripts/pptx_morph_kit.py` (`SlideBuilder`, `write_pptx`, `make_svg_shape`, `make_radial_glow`, `make_portal_mask`, `build_curtain_pair`, `make_video_frame`, `split_script_runs`, `rpr_xml`, `build_*_chart_xml`, `build_minimal_xlsx`)
- **Audit + Render + Contact Sheet Gate:** `python3 scripts/render_deck.py <deck.pptx> --target desktop|web`

---

## The 4-Gate Data Flow

### Gate 1: Grill Me (Budgeted Smart Intake)

Do not jump straight into code or slides. Interview the user with up to 5 sharp questions (using smart assumptions so the user only confirms or corrects):

1. **Goal & Audience:** Who is watching and what decision/action must this deck drive?
2. **Must-Keep Perspectives & Verified Data:** What core arguments, angles, or real metrics must NOT be cut? (Flag missing numbers as `[NEED DATA]`.)
3. **Reference & Real Assets:** Does the user have an existing `.pptx`, video, link, or style to follow—plus any real logos, photos, maps, or closing videos (`!!closing_video`) to include?
4. **Language & Font Availability:** Which languages are used (Thai, Lao `Phetsarath OT` / `Noto Sans Lao`, English) and are those fonts installed on the playback machine (or needed via `.odttf` embedding)?
5. **Delivery & Environment:**
   - **Playback Target:** `desktop` (PowerPoint Offline: full 3D `<a:scene3d>`, `.odttf` fonts, `<p:timing>`), `web` (PowerPoint Online safe subset), or `read-only` (PDF/handout: each slide must read cleanly standalone without relying on multi-click Morph reveals).
   - **Room Lighting:** Bright room/projector (favor high-contrast light canvas) vs. dark/stage room (dark/crimson canvas works well).

---

### Gate 2: TRY / EXCEPT Style & Reference Resolution

- **TRY (User Has a Reference / Template):**
  - Extract exact RGB/Hex palette, typography, layout rhythm, and persistent `!!` anchors from the user's `.pptx`, video (`ffmpeg -i ref.mp4 -vf "fps=2" frames/frame_%03d.png`), or image.
  - Honor and elevate their existing visual language—never overwrite it with a generic preset.
  - Optionally save extracted tokens to `references/benchmarks.md` as a reusable Style Pack so future decks can invoke the style without creating a new skill.
- **EXCEPT NoUserReference (Multi-Source Web Research & Elevation):**
  - Search multiple design/presentation references on the web (`search_web`) matching `[Topic × Preferred Vibe × Audience]`.
  - **Anti-Dribbble Trap Filter:** Extract only **Color Palette** (60-30-10 rule, `>= 4.5:1` body contrast), **Grid Rhythm**, and **Visual Hierarchy**. Never copy tiny `6–8 pt` web-shot font sizes or tight English-only box heights; apply Thai/Lao line-height safety margins.
  - Synthesize **2–3 elevated visual directions** paired with fitting Morph Rigs (or select from `references/benchmarks.md` when speed is requested).

---

### Gate 3: Draft Plan Gate (Approve Before Compile)

Before generating `.pptx` files or build scripts, present a concise Draft Plan for user approval:

1. **Mini Style Preview:** 60-30-10 color tokens, projector type scale, and wireframe layout sketch.
2. **No-Drop Content Map:** `core_thesis`, `proof_points`, `takeaway`, and `[NEED DATA]` flags per topic—verifying zero perspectives were dropped.
3. **Slide-by-Slide Storyboard:** Layout pattern (`Hero` / `Bento` / `Split` / `Table` / `Chart`—never repeat the exact same grid 3 slides in a row), auto-split Morph continuations where needed, assigned Morph Rig (`A–G`), and `!!anchor` continuity map across `Slide N -> Slide N+1`.

**Wait for user confirmation or edits before proceeding to Gate 4.**

---

### Gate 4: Algorithmic Compile & Verification Gate

#### 1. Enforceable Design & Typography Standards (`10 x 5.625 in` 16:9 Canvas)

- **Projector Type Scale:**
  - Cover / Main Title: `22–34 pt` Bold
  - Card / Section Header: `14–18 pt` Bold (`0.46–0.54 in` height for 2-line headers)
  - Body Copy: `11–13 pt` (`line_spacing >= 1.18`, card padding `>= 0.14 in`)
  - Dense Technical / ER Diagrams: `9.5–11.5 pt` across 3–4 clean Bento clusters
  - **Hard floor:** `>= 9.0 pt` everywhere
- **Multilingual XML Binding & Pre-Build Solver (`scripts/layout_solver.py`):**
  - Call `validate_content_inventory()` before building slides.
  - Measure every Thai/Lao/Latin text block via `measure_text_block()` and `solve_vertical_stack()`, segmenting words on valid orthographic syllable/word boundaries; route `overflow_ids` to Morph continuation slides rather than shrinking text.
  - Verify `check_aabb_collisions(boxes) == []` and `compute_whitespace_ratio(boxes)`.
  - For Lao (`Phetsarath OT` / `Noto Sans Lao`), Thai (`Noto Sans Thai` / `Tahoma`), and CJK, explicitly bind `<a:cs typeface="..."/>` and `<a:ea typeface="..."/>` alongside `<a:latin>` on every `<a:rPr>` run (`split_script_runs()` / `rpr_xml()`).
- **Whitespace Budget:** `>= 28–30%` for editorial/hero slides; `>= 18%` for Bento/technical slides. One clear primary focal point per slide.

#### 2. Morph Choreography Standards (7 Rigs: A–G)

Map transitions semantically rather than moving objects randomly (skip multi-click overlapping Morph reveals when Delivery is `read-only`):

- **A. Parallax Depth:** Foreground/midground/background cutouts move at distinct speeds (`Velocity = Distance / Time`).
- **B. Off-Canvas Track:** Timeline steps, multi-slide overflow continuations, and sequential panels enter/exit from off-canvas (`x > 10 in` or `x < 0`).
- **C. Depth Card Stack:** Layered cards at progressive scale/tilt promote cleanly to focus or exit off-canvas / `1%` ghost opacity.
- **D. Crop-to-Pan (`<a:srcRect>`):** Zoom/pan into details on the exact same `!!` image anchor without duplicating cropped assets.
- **E. Seamless Split-Curtain & Vector Masks:** Left/right `<a:custGeom>` Freeform masks sharing exact seam vertices `(x_i, y_i)` (`build_curtain_pair`, `make_svg_shape`).
- **F. Portal Zoom:** Single `<a:custGeom>` compound-path mask (`make_portal_mask`) scaled 4x–8x across slides so the viewer flies through the opening.
- **G. True Pivot Rotary Dial & 0-Delay 1-Click Intro:** Construct wedges/gears at `0 deg` first, then animate `<a:xfrm rot="...">` (`shape.rotation`) across slides (never pre-rotate vertices).
- **1-Click Zero-Delay & `!!` Discipline:**
  - Assign `!!` prefixes to **5–15 hero anchors per slide** only. Never assign `!!` to micro-connectors or dense diagram nodes (morph outer `!!tech_card` and let interior nodes cross-fade).
  - Inject `<p159:morph option="byObject"/>` (use `byWord` or `byChar` only for Latin keyword/counter swaps; avoid `byChar` on Thai/Lao vowel clusters).
  - Gate Desktop-only `<a:scene3d>` / `<a:sp3d>` and `.odttf` `<p:embeddedFontLst>` with `target="desktop"|"web"` on `SlideBuilder` and `write_pptx()`, and add post-Morph stagger reveals via `SlideBuilder.add_stagger(...)`.

#### 3. Audit & Contact Sheet Before Delivery

Run the mandatory audit and render gate:

```bash
python3 scripts/render_deck.py path/to/deck.pptx --target desktop
```

Inspect `contact_sheet.png` and slide PNGs with `read_file` to confirm:
- `overflow == 0`, `collision == 0`, and zero text-on-text or text-on-border overlaps.
- `dups == 0` (zero duplicate `!!` anchor names on the same slide) and `0` XML/relationship/geometry-type mismatches.
- `--target web|desktop` policy compliance and 100% editability of all shapes, charts, and tables.
- Report truthfully whether rendering used LibreOffice + `pdftoppm` or the built-in OOXML fallback renderer (and note that live Morph playback/timing requires opening in MS PowerPoint).

---

## Gotchas

- Skill names in storage use hyphens (`ph-present`), representing `ph.present`.
- When deleting or merging slides in `python-pptx`, always drop the relationship via `prs.part.drop_rel(rId)` before removing the element from `prs.slides._sldIdLst`.
- Keep `SKILL.md` focused on workflow, gates, and standards; keep layout math, OOXML compilation, and rendering inside modular `scripts/` modules.
