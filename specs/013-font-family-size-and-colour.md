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
- [ ] The highlight picker offers the 15 Word highlight colours plus "No colour", and maps to `w:highlight` on save where the colour matches, falling back to shading otherwise
- [ ] All four controls reflect the caret's current formatting, and show a blank/indeterminate value for a mixed selection
- [ ] Subscript and superscript buttons are included, are mutually exclusive, and use Ctrl/Cmd+, and Ctrl/Cmd+.
- [ ] All of these round-trip through `.docx` and `.odt` save and re-open

## Out of scope
Font embedding and fallback rendering for missing fonts (SPEC-048). Character
spacing, kerning, small caps, all caps. Text effects and shadows.
