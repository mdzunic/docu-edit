---
id: SPEC-013
title: Font family, size, text colour and highlight
status: draft
area: formatting
depends_on: [SPEC-012]
---
## Why
Direct font control is the second thing users reach for, and the first place a
document's appearance visibly diverges from Word if we get the defaults or the
units wrong.

## Acceptance criteria
- [ ] The toolbar has a font family dropdown, a font size control, a text colour picker and a highlight colour picker
- [ ] The font family list includes the web-safe families plus the common document fonts (Calibri, Cambria, Times New Roman, Arial, Georgia, Verdana, Courier New, Liberation/Carlito equivalents), each item rendered in its own typeface
- [ ] The font list includes a "Fonts in this document" group at the top listing families used by the open document, including families not installed locally
- [ ] The size control accepts values from 1 to 409 pt in 0.5 pt steps, has +/− buttons, and rejects out-of-range input by clamping with a tooltip
- [ ] Font size is stored in half-points for `.docx` and points for `.odt` with no rounding drift across a save/open round trip
- [ ] The text colour picker offers a standard palette, recent colours, and a custom hex/RGB input
- [ ] The highlight picker offers the 15 Word highlight colours plus "No colour", and maps to `w:highlight` on save where the colour matches, falling back to shading otherwise — see "Highlight mapping" below
- [ ] A highlight outside the palette (from an `.odt`, or the text colour picker's custom input reused for highlight) saves as `<w:shd w:val="clear" w:color="auto" w:fill="RRGGBB"/>` and re-opens as the same colour
- [ ] "No colour" is an explicit off, not an absence: it survives a round trip through both formats and still clears a highlight inherited from the paragraph or character style
- [ ] A `.docx` run carrying both `w:highlight` and `w:shd` keeps both across a `.docx` round trip, and reports the shading as dropped on a cross-format save
- [ ] All four controls reflect the caret's current formatting, and show a blank/indeterminate value for a mixed selection
- [ ] Subscript and superscript buttons are included, are mutually exclusive, and use Ctrl/Cmd+, and Ctrl/Cmd+.
- [ ] All of these round-trip through `.docx` and `.odt` save and re-open

## Out of scope
Font embedding and fallback rendering for missing fonts (SPEC-048). Character
spacing, kerning, small caps, all caps. Text effects and shadows.

Paragraph shading, table cell shading and page background — this spec's use of
`w:shd` is confined to `w:rPr` (run shading) and carries no UI of its own.

## Decisions
`white` stays out of the picker but not out of the palette. Word's highlight UI
offers 15 colours; `ST_HighlightColor` also admits `white` and `none`. The
read/write palette therefore has 16 entries, so an incoming
`w:highlight w:val="white"` survives `.docx` → `.odt` → `.docx` instead of
vanishing — which is what it does today, because `HIGHLIGHT_NAMES` has no
`#ffffff` and `highlightToOdfColor` drops what it cannot map.

"No colour" is `false`, not `undefined`. `RunFormatting.highlight` widens to
`string | false`, matching how `bold`/`italic` already distinguish "not set"
from "explicitly off". `false` writes `w:highlight w:val="none"` and
`fo:background-color="transparent"`; both readers map those back to `false`
rather than discarding them (`docx/properties.ts` currently skips `val="none"`,
and `odfColorToCss` rejects `transparent`). Picking "No colour" always sets
`false`. It costs a redundant element on runs that never had a highlight — the
same trade already accepted for the SPEC-012 toggles, and the only form that
reliably overrides an inherited style.

`w:highlight` wins over `w:shd` on read, and the shading is kept beside it.
Word paints highlight over shading, so the effective colour is the highlight's;
but discarding the `w:shd` would break the SPEC-008 rule, and run properties
have no element-level opaque slot. `RunFormatting` gains `shadingFill?: string`
— a preservation slot, never shown in the picker — set only when a source run
carried both. ODF has one background slot, so a cross-format save emits the
highlight and reports the shading as a `DropRecord`; `DropKind` gains
`"runShading"` with the label "Run shading".

## Reference

### Highlight mapping
The model stores the OOXML name for a palette colour and a lowercase
`#rrggbb` string for anything else. Readers normalise into that form, so the
writers never have to re-match hexes against names:

| OOXML name | Word UI name | `fo:background-color` |
| --- | --- | --- |
| `yellow` | Yellow | `#ffff00` |
| `green` | Bright Green | `#00ff00` |
| `cyan` | Turquoise | `#00ffff` |
| `magenta` | Pink | `#ff00ff` |
| `blue` | Blue | `#0000ff` |
| `red` | Red | `#ff0000` |
| `darkBlue` | Dark Blue | `#00008b` |
| `darkCyan` | Teal | `#008b8b` |
| `darkGreen` | Green | `#006400` |
| `darkMagenta` | Violet | `#8b008b` |
| `darkRed` | Dark Red | `#8b0000` |
| `darkYellow` | Dark Yellow | `#808000` |
| `lightGray` | Gray-25% | `#d3d3d3` |
| `darkGray` | Gray-50% | `#a9a9a9` |
| `black` | Black | `#000000` |
| `white` | *(not in the picker)* | `#ffffff` |

The table already exists as `HIGHLIGHT_NAMES` in `src/lib/odf/properties.ts`,
keyed by hex and inverted by the `.odt` writer. Extend that one; do not add a
second copy — `src/lib/renderFormatting.ts` holds a third for CSS and should
import from the same source.

`.odt` needs no fallback: `fo:background-color` on the run's automatic style
takes any colour. `.docx` chooses per run:

- highlight is one of the 16 names → `<w:highlight w:val="name"/>`
- highlight is `false` → `<w:highlight w:val="none"/>`
- highlight is a `#rrggbb` outside the palette →
  `<w:shd w:val="clear" w:color="auto" w:fill="RRGGBB"/>` and **no**
  `w:highlight`
- `shadingFill` is set → the same `w:shd`, alongside whatever `w:highlight`
  the rule above produced

Uppercase hex without the `#`, as `cssColorToOoxml` already produces for
`w:color`. Today `docx/writer.ts` writes `formatting.highlight` into
`w:highlight` verbatim, so an `.odt`-sourced `#ff8800` emits an invalid
`ST_HighlightColor` — `validate.ts` flags it and Word rejects the file. That
branch is this spec's work, together with reading `w:shd` back in
`docx/properties.ts`, which ignores it entirely today.

`w:shd` must also be added to `RPR_SEQUENCE` in `src/lib/docx/validate.ts`.
The CT_RPr order is `… sz, highlight, u, effect, bdr, shd, fitText, vertAlign`,
so `shd` follows `u`, and this spec's subscript/superscript `w:vertAlign`
follows `shd`.
