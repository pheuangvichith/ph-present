# ph-present vNext Design Spec

## Intent & Why
Upgrade `ph-present` from a rule checklist with fixed template gravity into a natural, content-preserving, PlacePlate-grade Presentation Architecture & OOXML Compiler skill that supports unrestricted visual directions and PowerPoint Desktop (Offline) + Web (Online) compatibility.

## Goals
1. **Natural Smart Intake:** Never interrogate the user repeatedly; ask only 1 short question (*Morph/Landing in mind?*) when no visual reference or direction is supplied.
2. **Zero Content Loss (Structural Contract):** Require a slide-by-slide `Content Inventory` (`Core Thesis`, `Proof/Data`, `Takeaway`) before geometry placement, splitting dense points across Morph Sub-slides rather than shrinking or deleting content.
3. **Unrestricted Visual Range + Web Benchmark Elevation:** Support `Benchmark Mode` (500+ engagement web research -> Landing lock) and `Fast Pattern Mode`, without defaulting every deck to rotary gears or dark teal.
4. **Metrics-Driven Zero-Collision Layout (`scripts/layout_solver.py`):** Syllable-aware Thai/Lao line breaking, TTF width/height sizing, auto-overflow splitting, AABB collision (`IoU == 0`), and Whitespace budget verification.
5. **Full PowerPoint Engine (Phase 2):** `<p:timing>` post-Morph stagger, Compound-Path Portal Masks, Radial Glow, Multi-Chart (`Line`, `Doughnut`, `Stacked`), `<a:scene3d>`, and Font Embedding with honest Offline vs. Online profile gating.

## Non-Goals / Constraints
- Local Linux environment lacks MS PowerPoint and LibreOffice; Phase 1 targets deterministic layout/XML features fully verifiable locally via OOXML audit + TTF PNG contact sheets.
- Phase 2 runtime playback (60fps `<p:timing>`, `<a:scene3d>`) requires final playback validation in MS PowerPoint Desktop/Web; local verification checks strict XML schema.

## Phased Approach (Chosen over Big-Bang)
- **Phase 1:** RED baseline test -> `layout_solver.py` + unit tests -> `SKILL.md` & `references/` upgrade -> GREEN subagent + real non-gear Thai/English test deck.
- **Phase 2:** Full PowerPoint performance modules (`timing`, `compound portal`, `radial glow`, `multi-chart`, `scene3d`, `font embed`) built incrementally on verified Phase 1 foundation.
