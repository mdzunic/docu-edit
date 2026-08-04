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
- [ ] "Touched by the selection" means the paragraph contains a collapsed caret or intersects the selection range. Every paragraph-formatting command in this spec — alignment, increase/decrease indent, the space-before toggle, and "Custom spacing…" — evaluates independently against every touched paragraph
- [ ] Increase indent adds 1.27 cm to the start indent of every touched paragraph; decrease removes it and stops at 0
- [ ] Tab runs the same increase-indent command as the toolbar button when the caret is collapsed at offset 0 of the paragraph — an empty paragraph, or the start of its first line, not the start of a wrapped continuation line; Shift+Tab runs decrease indent
- [ ] Tab in any other caret or selection position inserts a literal tab character
- [ ] The line spacing menu offers Single, 1.15, 1.5, Double, and "Custom spacing…"
- [ ] "Custom spacing…" opens a dialog with line spacing (multiple / exactly / at least + value) and space before/after in points; OK applies the values to every touched paragraph, Cancel discards them
- [ ] The line spacing menu carries one space-before toggle entry. Its label and its effect are fixed by the table below, evaluated against each paragraph's own `spaceBeforePt`:

  | Paragraph's `spaceBeforePt` | Label shown | Effect when activated |
  | --- | --- | --- |
  | `0` or absent (`undefined`) | "Add space before paragraph" | set an explicit `12` pt |
  | any non-zero number of points | "Remove space before paragraph" | set an explicit `0` pt |

- [ ] The toggle label reflects only the first touched paragraph — the paragraph containing the caret, or the first paragraph of the selection in document order when it is non-collapsed; the effect still applies independently, strictly per the table, to every touched paragraph
- [ ] The toggle is stateless: "Remove" writes an explicit `0` rather than clearing the property. Removing a custom 6 pt space-before and adding it back yields 12 pt, never 6, and a style's space-before does not reappear after a save + reload
- [ ] A horizontal ruler above the page shows start indent, first-line indent, end indent and margin markers, and dragging a marker changes the corresponding value live
- [ ] Right-to-left paragraph direction is settable; indents are held and saved as logical start/end values, and the direction decides which physical edge each one paints on — in the rendered page, in the ruler, and in both writers
- [ ] For every property above, saving and re-opening in the same format yields an `equalStructural` model, with no unit drift greater than 0.01 cm on the converted lengths — see "Equality" in the Reference

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
paragraph" menu entries, what exact value is added to or removed from the
space-before property?
Answer: 12 pt — a hard value fixed in the acceptance criteria table. The pair
is one menu entry with two labels, toggling 0 ↔ 12 pt: "Add" shows when
space-before is 0 or absent (the label source is defined in the acceptance
criteria), "Remove" when it is non-zero. The toggle is stateless, so a custom
6 pt from the spacing dialog removes to an explicit 0 and adds back at 12 pt,
never 6. "Remove" writes an explicit 0 rather than clearing the property —
clearing it would let a style's space-before reappear on reload. This matches
Word, whose spacing dropdown's "Add space before paragraph" inserts 12 pt.

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

## Reference

This spec builds on `specs/011-rich-text-editing-core.md` (which rests on
`specs/002-create-blank-document.md` and `specs/004-open-docx-text-and-paragraphs.md`).
Restated here so a reader of this file alone does not have to guess.

### Model
A `DocumentModel` holds blocks of type `"paragraph"`; each `Paragraph` has a
`formatting?: ParagraphFormatting` (`src/lib/documentModel.ts`). This spec's
fields are `alignment?: "left" | "center" | "right" | "justify"`,
`spaceBeforePt?: number`, `spaceAfterPt?: number`, `lineSpacing?: number`,
`lineSpacingPt?: number` plus the indents, which are held as *logical* start/end
values: `indentFirstLineCm` is start-relative and keeps its name, and
`indentLeftCm`/`indentRightCm` are renamed to `indentStartCm`/`indentEndCm` as
part of this spec. An absent field (`undefined`) means "not specified directly
on the paragraph" — a paragraph style may still supply a value. An explicit `0`
is a real value that travels through the readers and writers and overrides any
style value.

Paragraph direction is `direction?: "ltr" | "rtl"`; absent means `"ltr"`.

### Commands
All mutations go through the immutable command API in `src/lib/commands.ts`
(SPEC-011): each command takes the current model plus a `ModelSelection` and
returns a new `{ model, selection }`, never mutating the input, and the DOM is
never the source of truth. This spec adds paragraph-formatting commands there
(e.g. an adjust-indent command taking a signed delta floored at 0, and
set/toggle-style commands for alignment, spacing and the space-before toggle)
that apply to every paragraph the selection touches. The toolbar and the
Tab/Shift+Tab bindings call those commands and nothing else.

### Units and equality
Model lengths are centimetres and model spacing is points. `.docx` converts to
twips (1 cm = 567 twips; 1 point = 20 twentieths-of-a-point) in
`src/lib/docx/units.ts`; `.odt` uses the centimetre/inch values of its
attributes. The same-format round-trip AC asserts
`equalStructural(openedModel, reopenedModel)` using the helper in
`src/lib/equality.ts`, plus a separate tolerance check that every indent and
spacing value in points/centimetres changes by no more than 0.01 cm across the
conversion.
