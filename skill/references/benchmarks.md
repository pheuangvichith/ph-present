# Benchmark Library & Offline/Online Compatibility

Reusable high-engagement (500+ likes/saves) visual & motion patterns and PowerPoint target compatibility rules.

## 1. PowerPoint Desktop (Offline) vs. PowerPoint Web (Online)

| Feature | PowerPoint Desktop (Win/Mac 365) | PowerPoint Web (Online) | Compiler Policy |
|---|---|---|---|
| `p159:morph` (`byObject`), `custGeom`, `srcRect` | Full support | Full support | **Safe core** across all targets |
| `<p:timing>` Fade / Fly / Wipe / MotionPath | Full support | Good support | Safe for post-Morph stagger |
| `<a:scene3d>` / 3D Camera Rotation | Smooth 60fps | Limited / flattened in places | Enable only for `target=desktop` |
| `<p:embeddedFontLst>` | Loads embedded OFL/TTF | May substitute system fonts | Pair with cross-platform fallback (`Arial`, `Tahoma` / `Noto Sans Thai`) |
| Embedded MP4 / Audio (`ppt/media/`) | Smooth playback | H.264/AAC MP4 under size cap | Always embed in `ppt/media/`, keep total media compact |

## 2. Reusable High-Engagement Archetypes (Non-Locked)

Log newly discovered 500+ engagement web benchmarks here (tokens + rig + pacing—never copy text):

### Benchmark 01 — Editorial Commerce Light (Retail / Strategy / SaaS)
- **Palette:** Warm Ivory `#F7F6F2` canvas, Pure White `#FFFFFF` surface, Deep Ink `#0F172A`, Muted Slate `#475569`, 1–3 semantic accents (e.g., Emerald `#059669`, Coral `#EA580C`, Indigo `#4F46E5`).
- **Typography & Space:** Bold editorial titles (`22–30 pt`), body `11–12.5 pt`, `>= 30%` negative space, hairline `1 pt` `#E2E8F0` borders + soft `blur=12 pt` elevation.
- **Motion:** Off-Canvas Track (`Rig B`) + subtle Parallax watermark (`Rig A`) + bespoke SVG wave (`Rig E`).

### Benchmark 02 — Bento Technical Architecture (Defense / Engineering)
- **Palette:** Ice-tint `#F4FBFD` or Deep Slate `#0F172A`, crisp 3–4 Bento cluster cards, `10–11.5 pt` minimum body copy.
- **Motion:** Hero container morph (`!!tech_card`) + Crop-to-Pan (`Rig D`) zoomed sub-views for dense schemas without micro-text.

### Benchmark 03 — Storytelling Split-Curtain & Portal (Campaign / Brand)
- **Palette:** Warm Espresso `#1C1917` or Cream `#FAF5EE`, Amber/Gold `#D97706`, high-contrast serif/sans pairing.
- **Motion:** Seamless Split-Curtain (`Rig E`) or Portal Zoom (`Rig F`) + layered cutout PNG parallax (`Rig A`).
