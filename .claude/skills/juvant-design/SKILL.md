---
name: juvant-design
description: Use this skill to generate well-branded interfaces and assets for Juvant (parent company) and its products (currently: Hardys), either for production or throwaway prototypes/mocks/slides. Contains essential design guidelines, colors, type, fonts, assets, and UI kit components for prototyping.
user-invocable: true
---

Read the `README.md` file within this skill, and explore the other available files.

If creating visual artifacts (slides, mocks, throwaway prototypes, etc), copy assets out of `assets/` and create static HTML files for the user to view. Import `colors_and_type.css` for canonical CSS variables (colors, type, spacing, radii, shadows). If working on production code, you can copy assets and read the rules here to become an expert in designing with this brand.

If the user invokes this skill without any other guidance, ask them what they want to build or design, ask some questions, and act as an expert designer who outputs HTML artifacts _or_ production code, depending on the need.

Non-negotiables to remember:
- Always use the provided logo files (`assets/juvant/*`, `assets/hardys/*`) — the wordmark font DejaVu Sans Condensed Bold is base64-embedded inside the SVG/PNG, never retyped in HTML.
- The violet→cyan gradient (`#A855F7` → `#22D3EE`) is the single signature — always horizontal on text, 135° on symbols. Never reversed.
- Never use emoji. The brand is intellectually serious.
- Hardys is "a second professor" — never "assistant" or "copilot".
- Proper nouns for feature surfaces: The Room, The Thread, The Mirror, The Lens, The Bridge, The Folio.
- Tagline belongs only to Hardys: *Every student deserves a Hardy.*
