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
- [ ] Increase indent adds 1.27 cm to the left indent; decrease removes it and stops at 0
- [ ] Tab at the start of an empty or first-line-caret paragraph increases indent; Shift+Tab decreases it
- [ ] The line spacing menu offers Single, 1.15, 1.5, Double, and "Custom spacing…"
- [ ] "Custom spacing…" opens a dialog with line spacing (multiple / exactly / at least + value) and space before/after in points, applied on OK and discarded on Cancel
- [ ] "Add space before paragraph" / "Remove space before paragraph" style entries appear in the line spacing menu and toggle the paragraph's space-before value
- [ ] A horizontal ruler above the page shows left indent, first-line indent, right indent and margin markers, and dragging a marker changes the corresponding value live
- [ ] Right-to-left paragraph direction is settable and swaps the meaning of left/right indent on both save formats
- [ ] All paragraph properties round-trip through `.docx` (twips) and `.odt` (cm/in) with no unit drift greater than 0.01 cm

## Out of scope
Tab stops (defer to SPEC-032's out of scope note; custom tab stops are a later
spec). Paragraph borders and shading. Widow/orphan control and keep-with-next.
Paragraph styles (SPEC-019).
