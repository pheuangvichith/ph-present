# Morph Rigs, Deconstruction & OOXML Cookbook

Reference playbook for `ph-present`. Use these camera/motion mechanics and OOXML patterns freely across any visual domain—never lock every deck into one fixed template.

---

## 1. Reference Deconstruction Protocol (Video / Template -> Editable PPTX)

When given a reference video, link, or template deck:

1. **Extract Distinct Keyframes:**
   ```bash
   ffmpeg -i reference.mp4 -vf "fps=2" frames/frame_%03d.png
   ```
   Group consecutive frames into discrete click-states `S1, S2, ..., Sn`. Ignore mid-transition interpolation frames except to measure motion easing/direction.
2. **Sample Design Tokens:**
   - Canvas background, primary elevated surface, secondary surface, border/stroke, primary accent, secondary accent, primary text, muted text.
   - Measure safe margins (`>= 0.40 in`), card corner radius (`adj` in `1/100000`), shadow blur/distance, and type hierarchy scale.
3. **Build the Morph Continuity Matrix:**
   Track every persistent `!!object` across adjacent slides before writing builder code:

   | `!!Object` | Layer | Slide 1 `(x, y, w, h, rot, alpha)` | Slide 2 `(x, y, w, h, rot, alpha)` | Transform Type |
   |---|---|---|---|---|
   | `!!hero_subject` | Cutout PNG | `(5.4, 0.8, 3.8, 4.6, 0, 100%)` | `(0.6, 1.1, 2.4, 2.9, 0, 100%)` | Parallax Scale |
   | `!!portal_mask` | `custGeom` | `(0, 0, 10, 5.625, 0, 100%)` | `(-12, -7, 34, 19, 0, 100%)` | Portal Zoom |
   | `!!card_next` | Vector | `(10.4, 1.4, 3.8, 3.4, 0, 100%)` | `(5.5, 1.4, 3.8, 3.4, 0, 100%)` | Off-Canvas Slide-In |

   Any object that appears on Slide `N` and disappears on Slide `N+1` should exit cleanly (slide off-canvas or fade `alpha=1%`), while persistent anchors must keep the **exact same `!!name` and geometry type** (`roundRect` stays `roundRect`, `custGeom` stays `custGeom`).

---

## 2. Seven PlacePlate-Grade Camera & Motion Rigs

Pick **1–2 primary rigs** suited to the story rather than reusing one motif everywhere:

### Rig A — Parallax Depth Planes (`Velocity = Distance / Time`)
- Separate subject cutout PNGs, hero cards, and oversized background typography onto distinct Z-layers.
- Move background display text `0.6–1.2 in`, midground cards `2.0–3.5 in`, and foreground cutouts `4.5–7.0 in` across the same slide transition so PowerPoint's fixed transition duration creates natural optical depth.

### Rig B — Off-Canvas Track & Editorial Carousel
- Stage incoming cards, timeline stops, or comparison panels just outside the `10.0 x 5.625 in` viewport (`x = 10.4 in` or `x = -4.5 in`).
- On the next slide, slide the active item off left (`x = -4.5 in`) while bringing the staged item into center focus (`x = 0.8 in`). Great for storytelling, roadmaps, and product tours.

### Rig C — Depth Card Stack / Deck Shuffle
- Layer 3 cards at `(x, y)` offsets with progressive scale (`100%`, `90%`, `80%`), subtle tilt (`0 deg`, `3 deg`, `6 deg`), and muted fill/opacity on back cards.
- On click, send the front card off-canvas (or to the back) and promote the next card to `(100%, 0 deg)` with crisp elevation shadow.

### Rig D — Camera Pan & Crop-to-Pan (`<a:srcRect>`)
- Keep the exact same `!!` picture name across slides.
- On Slide `N`, show the full overview (`srcRect = (0, 0, 0, 0)`). On Slide `N+1`, adjust `<a:srcRect l="..." t="..." r="..." b="..."/>` (units in `1/100000`, e.g., `l="45000"` crops `45%` from left) and enlarge/reposition the frame to pan smoothly into one quadrant without creating duplicate cropped image files.

### Rig E — Seamless Split-Curtain & Boolean Masks
- Build paired `<a:custGeom>` polygons sharing identical central jagged or wavy seam vertices (`build_curtain_pair` in `scripts/pptx_morph_kit.py`) so they lock flush on a beat slide and part cleanly on the reveal slide.
- For SVG waves, blobs, organic masks, or bespoke icons, use `make_svg_shape(...)` (`svg_path_to_ooxml_path`) so curves stay native `<a:cubicBezTo>` / `<a:quadBezTo>` editable vectors.

### Rig F — Text / Shape Portal Zoom
- Use `make_portal_mask(..., hole_type="circle|roundRect|rect|svg")` in `scripts/pptx_morph_kit.py` to compile a single native `<a:custGeom>` compound-path mask (outer canvas + counter-clockwise inner cutout) over a background scene or photo.
- On the next slide, scale the foreground mask 4x–8x (`w=40 in, h=22 in`) centered on its opening so the viewer "flies through" the portal into the scene behind it.

### Rig G — True Pivot Rotary Dial & 0-Delay 1-Click Intro
- Construct radial wedges (`!!wedge_0..7`) inside a symmetric square bounding box `(cx - R, cy - R, 2R, 2R)` in canonical `0 deg` (12 o'clock) orientation, then rotate via `<a:xfrm rot="...">`.
- For a 0-delay 1-click intro, tuck a compact `!!cover_card` behind the central hub (`rot=0 deg`) on Slide 1; on Slide 2, spin the hub (`rot=135 deg`), dock it to the corner, and expand `!!cover_card` to full header/hero size.

---

## 3. Art Direction Archetypes (Reference Case Studies, Not Limits)

Choose or invent an art direction that fits the audience. Examples:
- **Swiss Editorial Light:** Off-white paper `#F8FAFC` / `#F4F3EF`, stark ink `#0F172A`, crisp `1 pt` hairline rules, one vivid accent (`#EF4444` or `#2563EB`), `>= 35%` whitespace, asymmetric 12-column grid.
- **Aqua Neumorphic (e.g., PP-20 Academic/Technical):** Ice-tint `#E8F8FD` / `#F4FBFD`, aqua `#00CCF8` / `#008CA8`, white cards with soft outer shadow (`blur=12 pt, dist=3 pt, alpha=22%`), Bento clusters for dense schemas.
- **Teal Orbit / Navy Depth (e.g., PP-23 / PP-45 Startup/SaaS):** Deep slate `#0F172A` / `#20262E`, elevated cards `#1E293B` / `#2A323A`, cyan/teal `#38DCFC` / `#00A8B5`, bold KPI callouts.
- **Warm Storytelling / Cultural Campaign:** Deep espresso `#1C1917` or warm cream `#FAF5EE`, terracotta/gold `#D97706`, layered cutout photography + split curtains.

---

## 5. Pre-Build Layout & Thai/Lao Constraint Solver (`scripts/layout_solver.py`)

Before calling `write_pptx()`, validate every content inventory and calculate exact wrapped heights with installed TTF metrics so Thai/Lao copy never overflows or splits mid-syllable:

```python
import layout_solver as ls

# 1. Enforce zero-content-loss inventory (no unverified synthetic numbers)
errs = ls.validate_content_inventory(inventory)
assert not errs, errs

# 2. Measure and stack vertically; route overflow_ids to a Morph continuation slide
solved = ls.solve_vertical_stack(stack_items, x_in=0.50, w_in=5.50, start_y_in=1.32, max_bottom_in=5.00)

# 3. Verify zero sibling AABB overlap & whitespace budget
assert not ls.check_aabb_collisions(slide_boxes)
```

## 6. OOXML Cookbook (Strict Schema Patterns)

### Morph Transition (`<p:sld>` direct child after `<p:clrMapOvr>`)
```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition spd="med" advClick="1">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="med" advClick="1"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```
Use `option="byWord"` or `option="byChar"` only for keyword swaps or numeric counters on Latin text (avoid `byChar` on complex Thai/Lao vowel clusters).

### Strict `<p:spPr>` Child Element Order
PowerPoint validates `<p:spPr>` children in exact order:
`a:xfrm` -> `a:prstGeom` or `a:custGeom` -> fill (`a:noFill` / `a:solidFill` / `a:gradFill`) -> `a:ln` -> `a:effectLst` -> `a:scene3d` -> `a:sp3d`.

### Strict `<p:sld>` Child Element Order (`<p:timing>`)
`<p:cSld>` -> `<p:clrMapOvr>` -> `<mc:AlternateContent>` (`p:transition`) -> `<p:timing>`. Use `sb.add_stagger(sp_id, effect="fade|wipe|fly", delay_ms=..., trigger="afterPrevious|withPrevious|onClick")`.

### Script-Aware Multilingual `<a:rPr>`
Always split runs by script or bind both `<a:latin>` and `<a:cs>`/`<a:ea>` in strict order (`solidFill` -> `latin` -> `ea` -> `cs`):
```xml
<a:rPr lang="th-TH" altLang="en-US" sz="1200" b="0" dirty="0">
  <a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill>
  <a:latin typeface="Arial"/>
  <a:ea typeface="Noto Sans CJK SC"/>
  <a:cs typeface="Noto Sans Thai"/>
</a:rPr>
```

### Safe Slide Deletion in `python-pptx`
Always drop the relationship before removing the element from `sldIdLst`:
```python
rId = prs.slides._sldIdLst[idx].rId
prs.part.drop_rel(rId)
del prs.slides._sldIdLst[idx]
```
