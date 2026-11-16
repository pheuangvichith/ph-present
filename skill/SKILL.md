---
name: ph-present
description: Use when the user mentions ph.present or ph-present, asks to reverse-engineer a presentation video or template into an editable .pptx, build a PowerPoint deck with Morph transitions, or upgrade an existing .pptx presentation with complex scripts or dense technical diagrams.
---

# ph-present

Senior Presentation Architect & PowerPoint Morph compiler workflow for reverse-engineering reference decks/videos and designing original, PlacePlate-grade, 100% editable `.pptx` presentations.

## Core Principles

1. **Natural Flow, Zero Interrogation Overhead:** Ask only what is missing before building—never stall behind multi-round questionnaires when the user already gave enough direction.
2. **Content-First Invariant (Zero Content Loss):** Extract full substance first (`core_thesis`, `proof_points`, `takeaway`). When verified content exceeds one slide at `>= 11 pt`, split across Morph continuation slides instead of compressing or deleting points.
3. **Unrestricted Visual Craft:** Derive art direction, whitespace (`>= 30%` on editorial/hero slides, `>= 18%` on Bento), and motion rigs from the brief or reference—never lock every deck into one motif.
4. **100% Editable & Verified:** Build native shapes/charts, run `layout_solver.py` before export, and verify every rendered slide via `render_deck.py`.

## Bundled Toolkit (Invoke via `python3`)

- **Motion Rigs, Deconstruction & OOXML:** `references/morph-rigs-and-ooxml.md`
- **Benchmarks & Offline/Online Matrix:** `references/benchmarks.md`
- **Layout & Thai/Lao Constraint Solver:** `python3 scripts/layout_solver.py` (syllable-aware Thai/Lao wrap, TTF height sizing, `solve_vertical_stack`, `check_aabb_collisions`, `validate_content_inventory`)
- **Zero-Dependency OOXML + SVG Builder:** `scripts/pptx_morph_kit.py`
- **Audit + Render + Contact Sheet Gate:** `python3 scripts/render_deck.py <deck.pptx>`

## Workflow

### Step 1: Smart Intake Gate

Evaluate what the user already provided:
- **Predicate A (Visual reference/style provided):** If the prompt supplies a video, link, `.pptx`, mockup, or explicit aesthetic direction, proceed straight to Step 2 without pausing for questions.
- **Predicate B (Topic/content only, no visual direction):** Ask **one concise check** in a single turn:
  - Whether they already have a **Landing look or Morph style** in mind (or prefer AI benchmark selection)
  - Only if unstated and decision-critical: **Desktop Offline vs. PowerPoint Web** target, or **missing real numbers** where placeholders would otherwise be needed.
- **If the user tells you to pick/proceed:** Choose the best-fit art direction, record your call clearly, and move straight to Step 2.

### Step 2: Content Inventory & Benchmark Elevation

1. **Populate the Slide Content Inventory (Required Contract):**
   For every planned slide, define:
   - `core_thesis` (single main takeaway)
   - `proof_points` / verified metrics
   - `takeaway`
   - **Data Integrity Contract:** Use **only** numbers supplied by the user, source files, or cited research. Represent missing quantitative values qualitatively or mark `[NEED DATA]`—never invent unverified percentages, counts, or mockup stats.
2. **Select Elevation Mode:**
   - **Benchmark Mode (flagship / landing design):** Use `search_web` to inspect 500+ engagement design/presentation references matching the domain, sample tokens/pacing, and log reusable structural patterns to `references/benchmarks.md`.
   - **Fast Pattern Mode (urgent / clear style):** Select an archetype from `references/benchmarks.md` + `references/morph-rigs-and-ooxml.md`.
3. **Map Storyboard & Morph Continuity Matrix:**
   - Alternate layouts across adjacent slides (never repeat the exact same grid 3 slides in a row).
   - Select **1–2 Primary Motion Rigs** (*Parallax Depth*, *Off-Canvas Track*, *Depth Card Stack*, *Crop-to-Pan*, *Split-Curtain / SVG Mask*, *Portal Zoom*, or *Rotary Dial*) and map every `!!anchor` across `Slide N -> Slide N+1`.

### Step 3: Solve Layout First, Then Compile Native OOXML

1. **Pre-Build Constraint Solver (`scripts/layout_solver.py`):**
   - Call `validate_content_inventory()` before building slides.
   - Measure every Thai/Lao/Latin text block via `measure_text_block()` and `solve_vertical_stack()`. Route `overflow_ids` to a Morph continuation / Crop-to-Pan sub-slide rather than lowering fonts below `9.5–11 pt`.
   - Verify `check_aabb_collisions(boxes) == []` and whitespace ratio (`>= 0.28` editorial, `>= 0.18` Bento).
2. **4-Layer Separation & `!!` Discipline:**
   - **Vectors, Radial Glow & Compound Masks:** Build cards, pills, custom curves (`make_svg_shape`), ambient radial glows (`make_radial_glow`), and single-shape compound-path portal masks (`make_portal_mask`) as native AutoShapes or `<a:custGeom>`.
   - **Editable Multi-Charts:** Use `build_clustered_bar_chart_xml`, `build_stacked_bar_chart_xml`, `build_line_chart_xml`, or `build_doughnut_chart_xml` + `build_minimal_xlsx` (pass matching dark or light `text_color` / `grid_color`). Keep KPI callouts outside chart axes.
   - **Media Layers:** Keep cutout PNGs, photos, and embedded `!!closing_video` (`make_video_frame`, `10.0 x 5.625 in` at `(0,0)`) intact in `ppt/media/`.
   - **Selection Pane (`!!`):** Prefix only 5–15 persistent hero anchors per slide; keep interior Bento nodes/arrows unprefixed to prevent particle swarms.
3. **Morph, Stagger Timing, 3D & Multilingual Font Binding:**
   - Construct rotating shapes at canonical `0 deg` and animate via `<a:xfrm rot="...">`. Keep identical `!!name` and geometry type across consecutive slides.
   - Add post-Morph stagger reveals (`sb.add_stagger(sp_id, effect="fade|wipe|fly", delay_ms=...)`) cleanly in `<p:timing>` after `<p:transition>`.
   - Use `split_script_runs()` / `rpr_xml()` so Latin (`Arial`/`Calibri`), Thai (`Noto Sans Thai`/`Tahoma`), Lao (`Noto Sans Lao`/`Phetsarath OT`), and CJK bind cleanly across `<a:latin>`, `<a:cs>`, `<a:ea>`.
   - Set `target="desktop"|"web"` on `SlideBuilder` & `write_pptx(..., embed_fonts=..., target=...)` per `references/benchmarks.md` so Desktop-only `<a:scene3d>` / `<a:sp3d>` and `.odttf` font embedding (`<p:embeddedFontLst>`) are gated automatically for PowerPoint Web.

### Step 4: Audit & Visual Contact Sheet Verification

Run the mandatory audit and render gate:

```bash
python3 /home/pheuang02/.gemini/skills/ph-present/scripts/render_deck.py path/to/deck.pptx --target desktop
```

Inspect `contact_sheet.png` and slide PNGs with `read_file` to confirm:
- `0` XML/relationship errors, `dups == 0`, `overflow == 0`, `0` geometry-type mismatches
- Clean Thai/Lao syllable wraps and readable chart colors on the chosen canvas
- Report truthfully whether rendering used LibreOffice or the built-in OOXML fallback renderer (and note that live Morph playback/timing requires opening in MS PowerPoint).
