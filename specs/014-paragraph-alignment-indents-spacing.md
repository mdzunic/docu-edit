---
id: SPEC-014
title: Paragraph alignment, indents and spacing
status: draft
area: formatting
depends_on: [SPEC-011]
---
## Why
Paragraph geometry is what makes a document look like a document. It is also
the formatting most often lost in converters, so it needs explicit unit
handling on both readers and both writers.

## Acceptance criteria
- [ ] The toolbar has alignment buttons (left, centre, right, justify), increase/decrease indent buttons, and a line spacing menu
- [ ] Alignment applies to every paragraph touched by the selection, and shortcuts Ctrl/Cmd+L/E/R/J work
- [ ] Increase indent adds 1.27 cm to the start indent; decrease removes it and stops at 0
- [ ] Tab runs the same increase-indent command as the toolbar button when the caret is collapsed at offset 0 of the paragraph — an empty paragraph, or the start of its first line, not the start of a wrapped continuation line; Shift+Tab runs decrease indent
- [ ] Tab in any other caret or selection position inserts a literal tab character
- [ ] The line spacing menu offers Single, 1.15, 1.5, Double, and "Custom spacing…"
- [ ] "Custom spacing…" opens a dialog with line spacing (multiple / exactly / at least + value) and space before/after in points, applied on OK and discarded on Cancel
- [ ] The line spacing menu carries a single space-before toggle entry: "Add space before paragraph" when the paragraph's space-before is 0 or absent, setting it to 12 pt; "Remove space before paragraph" when it is non-zero, setting an explicit 0
- [ ] The toggle is stateless and writes an explicit 0 rather than clearing the property, so removing a custom 6 pt space-before and adding it back gives 12 pt, and a style's space-before does not reappear on reload
- [ ] A horizontal ruler above the page shows start indent, first-line indent, end indent and margin markers, and dragging a marker changes the corresponding value live
- [ ] Right-to-left paragraph direction is settable; indents are held and saved as logical start/end values, and the direction decides which physical edge each one paints on — in the rendered page, in the ruler, and in both writers
- [ ] All paragraph properties round-trip through `.docx` (twips) and `.odt` (cm/in) with no unit drift greater than 0.01 cm

## Out of scope
Tab stops (defer to SPEC-032's out of scope note; custom tab stops are a later
spec). Paragraph borders and shading. Widow/orphan control and keep-with-next.
Paragraph styles (SPEC-019).

## Decisions
Question: What exactly constitutes a "first-line-caret paragraph" for the
Tab/Shift+Tab indent behaviour — is it the caret anywhere on the paragraph's
first line?
Answer: No — it is the caret collapsed at text offset 0. "First line" only
excludes the visual start of a wrapped continuation line, which looks like a
line start but sits at a non-zero offset. An empty paragraph qualifies because
offset 0 is its only position. Everywhere else — mid-paragraph, a non-collapsed
selection, the start of a wrapped line — Tab inserts a literal tab character.
Tab and Shift+Tab dispatch the toolbar's own increase/decrease indent command
(start indent ±1.27 cm, floored at 0), not Word's first-line-indent variant, so
there is one code path and the ruler stays the way to set a first-line indent.

Question: When using the "Add space before paragraph" / "Remove space before
paragraph" menu entries, what exact value (e.g. 10pt, 12pt) is added or removed
from the space-before property?
Answer: 12pt. The pair is one menu entry with two labels, toggling 0 ↔ 12 pt:
"Add" shows when space-before is 0 or absent, "Remove" when it is non-zero. The
toggle is stateless, so a custom 6 pt from the spacing dialog removes to 0 and
adds back at 12 pt, never 6. "Remove" writes an explicit 0 rather than clearing
the property — clearing it would let a style's space-before reappear on reload.

Question: For RTL paragraphs, does "swaps the meaning of left/right indent" mean
the values swap on save, or that the UI labels and ruler markers swap, or both?
Do the increase/decrease indent buttons adjust the physical left indent or the
logical start indent?
Answer: The values are logical and never swap; only their physical mapping does.
`ParagraphFormatting` holds `indentStartCm`/`indentEndCm` (renamed from
`indentLeftCm`/`indentRightCm`) plus `direction?: "ltr" | "rtl"`;
`indentFirstLineCm` is already start-relative in both formats and keeps its
name. The buttons, the Tab/Shift+Tab bindings and the "increase indent" step all
move the *start* indent regardless of direction, so there is no direction check
in the command — in an RTL paragraph that visibly moves the right edge, which is
what Word and Google Docs do. The ruler mirrors, origin at the right edge, with
each marker keeping its logical identity. Rendering sets `direction: rtl` and
uses `margin-inline-start`/`margin-inline-end` so the browser mirrors for free.

The two formats disagree, which is why the model is logical rather than
physical. `w:ind/@w:left` and `@w:right` are the logical start/end edges
(`w:start`/`w:end` under the newer schema names), so the `.docx` writer emits
the stored values unchanged and adds `w:bidi` to `w:pPr` for RTL. ODF's
`fo:margin-left`/`fo:margin-right` follow XSL and are physical, so the `.odt`
reader and writer swap start/end when `style:writing-mode` is `rl-tb`. Confirm
that swap against a LibreOffice-authored RTL fixture before the writer is
final: getting the edge backwards mirrors every RTL document silently, and the
0.01 cm drift check above passes an inverted edge without complaint.
