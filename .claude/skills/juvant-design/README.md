# Juvant Design System

> *Augmenting human ingenuity.*

This folder is the design system for **Juvant Srls** — a Milan-based AI company — and its first product, **Hardys**, an AI Teaching Intelligence OS for higher education. It contains the brand foundations (colors, type, motif language), reusable assets (logos, social imagery, illustrations), tokenised CSS, and UI kits for the corporate and product surfaces.

## Index / Manifest

| File / folder | What it is |
|---|---|
| `README.md` | This file. Brand context, content + visual foundations, iconography. |
| `SKILL.md` | Agent Skills entry point — read this to use the system in Claude Code. |
| `colors_and_type.css` | Canonical CSS variables for color, type, spacing, radii, shadows. |
| `assets/juvant/` | Juvant corporate logos, icon, wallpapers, social banners, email template. |
| `assets/hardys/` | Hardys product logos (static + animated SVG), product icon. |
| `fonts/` | *(not used — web stack loads via Google Fonts `@import`: DM Serif Display, DM Sans, JetBrains Mono. Brand wordmark font DejaVu Sans Condensed Bold is base64-embedded inside the logo SVGs.)* |
| `preview/` | Design System tab preview cards (HTML). |
| `ui_kits/juvant-corporate/` | UI kit for juvant.io — investor/partner-facing corporate site. |
| `ui_kits/hardys-product/` | UI kit for Hardys — marketing site + in-product screens (Room, Mirror, Lens, Folio, Thread). |
| `slides/` | Title / section / content / quote slide templates for pitch decks. |

## Sources

All source material was provided by the user via a local-mounted `Design System/` folder. Referenced files (in case you have access):

- `Design System/Juvant/juvant-brand-guidelines.md` — Version 1.0 (2026), canonical.
- `Design System/Juvant/juvant-naming-strategy.md` — Product naming philosophy.
- `Design System/Juvant/hardys-naming-strategy.md` — Hardys-specific naming + tone.
- `Design System/Hardys/hardys-naming-strategy.md` — duplicate of the above.
- `Design System/Juvant/Docs/juvant-email-template.html` — email chrome template.
- `Design System/Juvant/Logo/*.png`, `Design System/Hardys/Logo/*.{png,svg}` — logos.
- `Design System/Juvant/Icons/*.png`, `Design System/Hardys/Icons/*.png` — icons.
- `Design System/Juvant/Social/*.png` — LinkedIn + X banners.
- `Design System/Juvant/Docs/juvant-wallpaper.png`, `…-zoom-background.png` — desktop imagery.

No Figma file, no codebase, and no slide deck were provided. UI kits are therefore built from the brand guidelines + asset set, not from a source-of-truth app. **Flag:** once real product code or Figma exists, the UI kits should be regenerated against those sources for pixel fidelity.

---

## 1 — Company & product context

### Juvant Srls (parent)
- Milan-based. Latin root *iuvare* (to help, to benefit).
- Mission: **AI agents that augment human ingenuity.**
- Speaks to investors, partners, institutions. Tone: a brilliant colleague who has seen the future and explains it calmly.
- Every product is named after the **enabler behind a genius** — not the genius. "Intelligence already exists. Our agents help it emerge."
- Future pipeline names (not yet committed): Vogts, Ruskins, Mosers, Salieris, Flexners.

### Hardys (first product)
- **AI Teaching Intelligence OS** for higher education.
- Named after G.H. Hardy — the mathematician who recognised Ramanujan. The *plural* signals a platform of agents, not one assistant.
- **Tagline:** *Every student deserves a Hardy.* (lowercase after the first word; never truncate.)
- Positioning: **a second professor**, not an assistant. Interacts **directly with students**, not through the teacher.
- Named feature surfaces (proper nouns — always capitalised):
  - **The Room** — student interaction engine (virtual classroom).
  - **The Thread** — longitudinal learning memory.
  - **The Mirror** — real-time faculty analytics.
  - **The Lens** — institutional / program-director dashboard.
  - **The Bridge** — guest-speaker integration.
  - **The Folio** — each student's evolving learning portrait.
- Audience: faculty, program directors, academic leaders. Built for senior professionals — complexity is the default, not simplified.
- Domain: `hardys.ai` (registered). EUIPO Class 41 + 42 clear.

### Brand architecture
```
JUVANT (parent)
└── Hardys       → hardys.ai     ✅ active
    [Future]     → Research / Creative / Executive coaching
```
Both share: violet→cyan gradient, node/arc symbol vocabulary, enabler naming philosophy.

---

## 2 — Content fundamentals

The Juvant voice is **calm, precise, narrative-first**. It behaves more like a senior academic than a marketing department. Copy is **short, declarative, and confident** — sentences don't hedge, but they also don't sell.

### Register & tone
- **Juvant (corporate):** precise, confident, visionary-without-grandiose. A brilliant colleague. No hype, no buzzwords, no hollow AI rhetoric.
- **Hardys (product):** warm, intellectually serious, respectful of the teacher and the student. A professor who understands teaching *and* technology. Never patronising, never edtech-cheery, never assistant-framing.

### Stance & grammar
- **You, not we.** Speak to the reader as a peer: *"You teach a 200-person seminar. Hardys sees each of them."* The first person plural ("we believe…") is allowed sparingly in manifesto-style copy, never in feature description.
- **Active voice, short sentences.** Fragments are allowed for emphasis — especially as eyebrows or closers.
- **Lead with story, then data.** Every feature intro begins with a scene or a claim. Numbers come after the claim, never as the lede.
- **Title Case for named surfaces** (The Room, The Mirror) — they are proper nouns.
- **Sentence case for UI labels** (Settings, New session) — not Title Case, not ALL CAPS.
- **ALL CAPS + letter-spaced** only for eyebrows / section labels / the wordmarks. Never for button labels.
- **Emoji: never.** No emoji in UI, decks, email, or marketing copy. The brand is adult and intellectually serious; emoji undermines that.
- **Exclamation marks: avoid.** One per deck at most, if the sentence truly earns it.

### Exact words to use
- *second professor* (not "AI assistant")
- *Teaching Intelligence* (not "analytics")
- *longitudinal learning* (not just "memory")
- *in the room* (not "in the platform")
- *directly with students* (emphasises non-mediated interaction)
- *amplify* (not "empower")

### Exact words to avoid
- copilot, assistant, chatbot
- AI-powered (generic)
- personalised learning (overused in edtech)
- seamless (meaningless)
- revolutionise, disrupt, unleash, game-changing
- purple-prose adjectives: magical, delightful, effortless

### Example copy snippets
- **Hero (Hardys):** *Every student deserves a Hardy. The Teaching Intelligence OS for the AI era.*
- **Hero (Juvant):** *Augmenting human ingenuity. Milan-based. Agents that see what no one else sees.*
- **Feature intro (The Room):** *A virtual classroom where Hardys meets your students directly — precise, patient, and never tired.*
- **Feature intro (The Mirror):** *Shows you what you cannot see in real time: which students are lost, which are ahead, and why.*
- **CTA — strong:** *Talk to us.* / *See Hardys in action.* / *Read the paper.*
- **CTA — soft:** *Learn more* is forbidden. Replace with *Read the rationale* or *See the approach*.
- **Closer / manifesto:** *Intelligence already exists. Our agents help it emerge.*

---

## 3 — Visual foundations

### Palette — the single most distinctive thing
The **violet → cyan gradient** is the single most recognisable element of the brand. Every surface uses it at least once — usually as the wordmark fill, a 1px gradient hairline separator, or a CTA button. Reversed or vertical text gradients are forbidden.

- **Void** `#080B14` — primary dark background. Hero sections, title slides, covers.
- **Violet** `#A855F7` — primary accent, left end of gradient.
- **Cyan** `#22D3EE` — secondary accent, right end of gradient.
- **Violet Mid** `#7C3AED` — hover / filled states.
- **Off-White** `#E8EEFA` — body text on dark.
- **Slate** `#94A3B8` — secondary text, captions.
- **Dark Card** `#161D2C` — elevated surfaces on dark.
- **Light BG** `#F4F6FB` — light page background.
- Ink (`#0D1117`) + Muted (`#6b7280`) — text on light.

**Rules:** dark for hero / covers, light for content-heavy sections. Never use violet gradients as decorative backgrounds on white — that is the generic AI aesthetic Juvant explicitly rejects. The gradient is for *text*, *hairlines*, and *single CTA buttons*.

### Typography
- **Wordmarks** (JUVANT, HARDYS) — *DejaVu Sans Condensed Bold*, ALL CAPS, 2px letter-spacing, gradient fill. Embedded as base64 inside SVG files for cross-platform consistency.
- **Slide titles** — Georgia Bold, 26–32pt.
- **Slide section labels** — Arial Bold, 11pt, ALL CAPS, letter-spacing 4.
- **Slide body** — Arial Regular, 11–13pt. Quotes: Georgia Italic, 13–14pt. Captions: Arial 9–10pt.
- **Web editorial headlines** — DM Serif Display (or Playfair Display as fallback).
- **Web body / UI** — DM Sans (Outfit as fallback).
- **Technical / code** — JetBrains Mono.

Headlines are **serif and confident** — they carry the academic weight of the brand. UI type is humanist sans; never Inter or Roboto (too generic). Body line-height is 1.5–1.75 — never tight.

### Spacing, radii, elevation
- **Spacing scale:** 4 · 8 · 12 · 16 · 20 · 24 · 32 · 40 · 48 · 64 · 80 · 96 · 128 px.
- **Radii:** 4 (inputs), 6 (chips), 8 (buttons), 12 (cards — default), 16 (large cards), 20 (hero panels), 999 (pills). Never fully square except for hairlines.
- **Shadows on light:** single layer, low blur, tinted toward Void (`rgba(8,11,20,0.18)`). Flat and editorial.
- **Glow on dark:** the signature `feGaussianBlur` effect from the logo — 1.5–2.5 stdDev — reappears on dark surfaces behind accent nodes, CTA buttons, and key UI elements. **Used sparingly** — glow signals energy, not decoration.

### Backgrounds & imagery
- **Full-bleed Void** for title slides and hero sections.
- **Ambient dot fields** — faded violet/cyan dots scattered against Void evoke "distributed intelligence." Reuse the motif from the logo: nodes, arcs, connecting lines.
- **No textures, no grain, no photographic gradients as decoration.** The brand is geometric and diagrammatic, not photographic.
- **Photography:** real, warm, intellectually serious — university environments, focused human work, collaboration. Avoid stock "AI blue brain," circuits, robots. When used, images are desaturated-to-warm, never cool-blue.
- **Gradients** appear as text fills, hairline separators, and single-element glows. Not as page backgrounds.

### Motion
- **Easing:** `cubic-bezier(0.16, 1, 0.3, 1)` — confident, decelerating out. Never bouncy.
- **Duration:** 160ms (micro), 280ms (standard), 480ms (page-level).
- **Fades and slight translation preferred over scale.** Pulse/glow is allowed on the orbit in the Hardys animated logo — that is its home. Elsewhere, motion is calm.
- **No parallax, no scroll-jacking.**

### Interaction states
- **Hover:** on dark, gradient elements get `filter: brightness(1.08)` + brand glow. Ghost elements get a 4% off-white wash. On light, buttons darken 8%.
- **Press:** 1px translate-Y, slight brightness dip. Never scale down.
- **Focus:** 2px gradient outline at 2px offset. Never remove focus rings.
- **Disabled:** 40% opacity, no pointer.

### Borders, separators, hairlines
- Borders are **1px, low-contrast**: `rgba(232,238,250,0.08)` on dark, `rgba(13,17,23,0.06)` on light.
- The **gradient hairline** (1px, violet→cyan, 40% opacity) is the canonical separator — used in the email template header, between sections on dark, under slide titles. Do not overuse — one per screen is ideal.

### Cards
- **Dark card:** `#161D2C` background, 1px off-white-alpha border, 12px radius, no shadow (rely on background contrast for elevation).
- **Light card:** pure white, 1px ink-alpha border, 12px radius, single low shadow.
- No colored-left-border cards. No gradient borders except as explicit accent (outline button variant).

### Use of transparency + blur
- Blur only for the logo's own `feGaussianBlur` glow.
- No backdrop-filter chrome. No frosted-glass panels unless explicitly specified.
- Transparency is used for hairlines, ambient-dot fields, and protection gradients over imagery.

### Layout rules
- **Content column:** 1200–1280px max on the web.
- **Generous whitespace** — section padding 80–128px on desktop.
- **Reading line length:** 60–75ch.
- **Fixed nav:** transparent on hero, solid Void with gradient hairline below once scrolled.
- **Grid:** 12-col at desktop, 8-col tablet, 4-col mobile. 24px gutter.

---

## 4 — Iconography

Juvant does not ship a custom icon font. Icons are sourced from **[Lucide](https://lucide.dev)** (CDN) — stroke icons, 1.5–2px weight, round caps — consistent with the diagrammatic, line-based feel of the Juvant symbol itself. Lucide is the substitute for the react-icons call-out in the brand guide.

### Usage
- **On dark:** white (`#E8EEFA`) at 88% opacity, or gradient-tinted for key moments.
- **On light:** ink (`#0D1117`) or violet (`#A855F7`).
- **Size:** 16px (inline), 20px (UI), 24px (button), 48px+ (feature marker).
- **Stroke:** 1.75px default. No fill variants except for toggle/selected states.
- **Never recolor outside the palette**. No red, green, yellow except for system states (success cyan-adjacent, error violet-adjacent).

### Brand-specific symbols (not icons)
The **nodes / arcs / connecting lines** vocabulary from the logos can be used as editorial marks in slides and marketing — distinct from UI icons. They appear as SVG, often with the signature glow. These are *illustrations*, not a system.

### Emoji & unicode
- **Emoji: never.** Explicit brand rule.
- **Unicode marks:** allowed sparingly as typographic accents — em dash (—), middle dot (·), arrow (→), check (✓). Never faces or decorative unicode.

### Custom brand SVGs (provided)
- `assets/hardys/hardys-logo-animated-dark.svg` / `…-transparent.svg` — the animated logo with orbiting node and pulse.
- Logos (Juvant + Hardys) are available in dark, transparent, white, and EUIPO (legal) variants.
- No third-party illustration set is bundled. If editorial illustrations are needed, commission line-work in the same node/arc vocabulary — do not mix flat-shape illustration styles.

### Substitution flags
- **Lucide** is used in place of react-icons (brand guide says "react-icons or equivalent"). Keep Lucide unless the product team picks a specific set.
- **DejaVu Sans Condensed** is only used in the wordmarks (which are already shipped as SVG). For the rest of the system, DM Sans + DM Serif Display + JetBrains Mono (Google Fonts) cover the web stack.

---

## 5 — How to use this system

1. Import `colors_and_type.css` at the top of any HTML surface.
2. Use `--juv-*` variables for colors, `--font-*` for type, `--s-*` for spacing.
3. For UI surfaces, grab components from `ui_kits/juvant-corporate/` (marketing/corporate) or `ui_kits/hardys-product/` (product app + marketing).
4. For decks, start from `slides/index.html`.
5. For iconography, use Lucide via CDN: `<script src="https://unpkg.com/lucide@latest"></script>` then `<i data-lucide="book-open"></i>` + `lucide.createIcons()`.
6. Never invent new gradients, new typefaces, or new iconography without editorial review.

---

*Juvant Srls — juvant.io — Confidential*
