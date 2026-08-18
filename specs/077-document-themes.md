---
id: SPEC-077
title: Document themes — colour and font schemes
status: draft
area: formatting
depends_on: [SPEC-013, SPEC-019, SPEC-042, SPEC-073]
---
## Why
Two specs deferred the same thing in the same words. SPEC-019's out-of-scope
list ends with "Style sets / theme switching"; SPEC-042's begins with "Theme
colours and font schemes". The result is that every colour in every document
this product touches is a literal `#rrggbb` and every font is a literal
family name — including the colours and fonts that were *not* literals in
the source file.

That is a fidelity problem before it is a feature. In `.docx`, a run styled
"Accent 1" stores `<a:srgbClr>` nowhere — it stores
`<w:color w:themeColor="accent1"/>`, resolved through `theme1.xml`, and the
same is true of `+mj-lt` / `+mn-lt` font references in every built-in style.
ODF does the same job through its own default styles. When this product
reads such a document it resolves the reference to a literal and forgets the
link, so:

- changing the theme in Word after a round trip no longer recolours the
  document, because nothing references the theme any more;
- every heading's font is baked, so a document built on Calibri Light
  headings and Calibri body reopens as two unrelated font choices;
- SPEC-073's shape style presets and SPEC-076's table banding have no
  coordinated palette to draw from, so each invents colours;
- and there is no way to restyle a document at all — the operation every
  word processor offers as one click.

The theme is also the last piece SPEC-058's design system does *not* cover
and must not be confused with: SPEC-058 themes the **application**, this
spec themes the **document**. They are different colours, stored in
different places, and a change to one must never affect the other.

## Acceptance criteria
- [ ] A document theme is modelled per "The theme model": twelve named colours and two font choices, stored on the document
- [ ] `Run.color`, `Run.highlight`, `ParagraphFormatting.shadingFill` (SPEC-074), `TableCell.shadingFill`, `CellBorders` colours, `ShapeFill`/`ShapeOutline` colours (SPEC-073) and chart series colours (SPEC-052) may each hold a **theme reference** instead of a literal, per "Theme references"
- [ ] A theme reference resolves to a literal colour for rendering, for export and for every contrast computation, through one function, and nothing downstream of that function knows about themes
- [ ] Changing the document's theme recolours every theme-referencing value in the document at once — text, tables, shapes and charts — in one undo entry, and leaves every literal colour untouched
- [ ] Changing the theme's heading or body font changes every style and run that references it, and leaves explicitly-chosen fonts untouched
- [ ] A theme gallery (`Format → Theme…`) offers a set of built-in themes, each shown as a live preview of the current document, plus "Customise…" for editing the twelve colours and two fonts of the current theme
- [ ] The colour pickers — text colour, highlight, shading, shape fill and outline, table shading, chart palette — show the theme colours as a labelled top row with tint and shade variants beneath, and the standard colours below, so choosing a theme colour is the default path and choosing a literal is still possible
- [ ] A theme colour chosen with a tint or shade stores the reference **and** the adjustment, not the resolved literal, so it re-tints when the theme changes
- [ ] Themes round-trip through `.docx` `theme1.xml` — `a:clrScheme` (dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink) and `a:fontScheme` (`a:majorFont`, `a:minorFont`) — with references written as `@w:themeColor`/`@w:themeTint`/`@w:themeShade` and `@w:asciiTheme`, asserted on the written XML
- [ ] Themes round-trip through `.odt` per "Format mapping — .odt", which has no theme part: the mapping and its stated loss are explicit, not silent
- [ ] A `.docx` opened, saved and reopened keeps its theme part **byte-identical** where the theme was not edited, per SPEC-008's preservation rule — an unedited theme is preserved, not regenerated
- [ ] A document with no theme (a `.odt`, a Markdown import, a blank document) gets the product's default theme, and its existing literal colours stay literal — nothing is retroactively converted to a reference
- [ ] The application theme (SPEC-059) and the document theme are independent: switching the app to dark mode changes no document colour, and changing the document theme changes no app colour, asserted by a test that does both
- [ ] SPEC-073's shape style presets and SPEC-076's table style banding draw their colours from the theme, so a theme change restyles shapes and tables coherently
- [ ] The accessibility checker (SPEC-043) evaluates contrast on **resolved** colours, so a theme change can turn a passing document into a failing one and the checker says so
- [ ] Fixtures with a theme-coloured heading, a theme-tinted table, a theme-filled shape and an explicitly-coloured run open, survive a theme change with the explicit colour unmoved, save and reopen `equalStructural`

## Out of scope
Theme *effects* (`a:effectStyleLst`) and background fill styles — preserved
per SPEC-008, never modelled. Style sets — alternative formatting for the
built-in styles, a related but separate Word feature. Per-section themes.
Importing a theme from another document, or a `.thmx` file. Colour
variations beyond tint and shade. Automatically converting an existing
document's literal colours into theme references. Any change to the
application's own theming (SPEC-058, SPEC-059).

## Reference

### The theme model
```ts
// src/lib/documentModel.ts
export interface DocumentTheme {
  name: string;
  colors: {
    /** The four base colours. OOXML's names are kept, because they are
     *  what the file format uses and inventing a synonym would guarantee
     *  a mapping bug: dk1/lt1 are the text/background pair, dk2/lt2 the
     *  secondary pair. */
    dk1: string; lt1: string; dk2: string; lt2: string;
    accent1: string; accent2: string; accent3: string;
    accent4: string; accent5: string; accent6: string;
    hlink: string; folHlink: string;
  };
  fonts: {
    /** Headings. */
    major: string;
    /** Body. */
    minor: string;
  };
}

export interface DocumentModel {
  …
  /** Absent means the product default theme. */
  theme?: DocumentTheme;
}
```

Twelve colours and two fonts is not a design choice — it is the OOXML
scheme, and matching it exactly is what makes the round trip lossless.

### Theme references
```ts
// src/lib/core/colors.ts
export type ThemeColorSlot =
  | "dk1" | "lt1" | "dk2" | "lt2"
  | "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6"
  | "hlink" | "folHlink";

/** A colour anywhere in the model: a literal, or a theme reference with an
 *  optional lightening (tint) or darkening (shade) in 0..1. */
export type ModelColor =
  | string                                             // "#rrggbb" or "auto"
  | { theme: ThemeColorSlot; tint?: number; shade?: number };

/** The single resolution point. Every renderer, writer, contrast check and
 *  picker calls this; nothing else interprets a ModelColor. */
export function resolveColor(color: ModelColor, theme: DocumentTheme): string;

export function isThemeColor(color: ModelColor): color is { theme: ThemeColorSlot; tint?: number; shade?: number };
```

Every field that today holds a colour string widens to `ModelColor`. The
string form stays valid and unchanged, so no stored document migrates and
no shipped test that asserts a literal colour breaks — a widening, not a
replacement. `resolveColor` is pure, lives in the product-neutral core
layer, and is unit-tested against known tint/shade arithmetic (the
luminance-modulation OOXML specifies, not a naive RGB blend, since the two
differ visibly at high tints).

Font references are the same shape, smaller:

```ts
export type ModelFont = string | { theme: "major" | "minor" };
export function resolveFont(font: ModelFont, theme: DocumentTheme): string;
```

The discipline that makes this safe is the single resolution point. A
renderer that reads `run.color` directly and passes it to canvas is a bug
this spec must not leave anywhere, so the type change forcing every call
site through `resolveColor` is the mechanism, not an inconvenience.

### Format mapping — .docx
Read: `word/theme/theme1.xml`'s `a:clrScheme` and `a:fontScheme` populate
`DocumentTheme`; `@w:themeColor` with `@w:themeTint`/`@w:themeShade` on
runs, cells, borders and shading become references; `@w:asciiTheme="majorHAnsi"`
/ `"minorHAnsi"` become font references; DrawingML `a:schemeClr` inside
shape and chart properties likewise.

Write: the theme part is **preserved verbatim** unless the theme was
edited, per SPEC-008 — the part contains effect and background styles this
model does not carry, and regenerating it would discard them. An edited
theme rewrites only `a:clrScheme` and `a:fontScheme` within the preserved
part, leaving the rest of the XML untouched, which the byte-identity AC
tests directly.

The three colour slots that need care on the way out: `dk1`/`lt1` are
usually `a:sysClr` (`windowText`/`window`) with a `lastClr` fallback, not
`a:srgbClr`. Reading takes `lastClr`; writing preserves whichever form the
source used and emits `a:srgbClr` for a theme this product created.

### Format mapping — .odt
ODF has no theme part. The mapping is therefore lossy in one direction and
is stated rather than fudged:

- **Write** — every theme reference is resolved to a literal colour and
  font in the output, and the theme itself is stored in the document's
  own settings so a re-open by this product recovers it. One
  `DropRecord` of a new `DropKind` `"theme"` is reported per save,
  naming the theme, so the conversion report (SPEC-008) tells the user
  their document will no longer re-theme in Writer.
- **Read** — an `.odt` gets the product's default theme with every colour
  literal, exactly as today.

An `.odt` written by this product and reopened by this product keeps its
references. An `.odt` written by this product and opened in Writer looks
identical and has no theme, which is the truthful outcome.

### The picker
Every colour picker gains a theme row: ten swatches (the four base colours
and six accents, using the theme's own names), each with five tint/shade
variants beneath, followed by the standard-colour grid and the custom
option that already exists. Recent colours (`docu-edit:recent-text-colors`)
continue to store literals.

A swatch's accessible name is its slot name plus its variant — "Accent 1,
lighter 40%" — because a colour swatch with no name is unusable with a
screen reader and "the blue one" is not a name.

### Interaction with the application theme
Stated as an invariant because the two are one word apart and a mistake
here would be visible and confusing:

| | Application theme (SPEC-058/059) | Document theme (this spec) |
| --- | --- | --- |
| Scope | The app's chrome | The document's content |
| Storage | `docu-edit:preferences` | `DocumentModel.theme`, saved into the file |
| Values | CSS custom properties | Twelve colours, two fonts |
| Affects the saved file | Never | Always |
| Affects print/PDF | Never | Always |

SPEC-058's token test already forbids a document-scoped style from using an
app token; this spec adds the mirror assertion, that no app surface resolves
a document theme colour for its own chrome.

### Specs this changes
- `specs/019-paragraph-styles-and-style-picker.md` — "Style sets / theme
  switching" is half-resolved: theme switching lands here, style sets stay
  out of scope. Style resolution is unchanged; a style's colours and fonts
  may now be references.
- `specs/042-custom-styles-manager.md` — "Theme colours and font schemes"
  leaves its out-of-scope list; the style editor's colour and font fields
  offer theme values.
- `specs/013-font-family-size-and-colour.md` — the colour and highlight
  pickers gain the theme row; the run model's colour field widens to
  `ModelColor`.
- `specs/073-shape-styling-depth.md` — its style presets become
  theme-driven, which that spec anticipates.
- `specs/076-table-styles-sort-and-text-table-conversion.md` — banding
  colours become theme references.
- `specs/052-charts.md` — the chart palette becomes the theme's six
  accents, which is what both formats do.
- `specs/043-accessibility-checker.md` — contrast is computed on resolved
  colours; the check itself is unchanged.
- `specs/008-round-trip-preservation.md` — `DropKind` gains `"theme"`;
  the theme part joins the parts preserved byte-identically.
