---
id: SPEC-023
title: Table formatting and cell operations
status: draft
area: structure
depends_on: [SPEC-022]
---
## Why
A table that cannot be sized, merged or bordered is a grid, not a table. This is
also where documents from Word most often look wrong if borders and widths are
not modelled properly.

## Acceptance criteria
- [ ] Dragging a column boundary resizes columns live, with a minimum width, and updates the model on release
- [ ] Dragging a row boundary sets a minimum row height; rows still grow to fit content
- [ ] Selecting a cell range and choosing "Merge cells" produces one cell whose content is the merged cells' content joined as paragraphs
- [ ] "Split cell…" asks for rows and columns and splits the selected cell accordingly
- [ ] A table properties dialog sets border style, width and colour (per table and per selected cells), cell background shading, cell padding and vertical alignment
- [ ] Border controls offer the standard presets: all borders, outer only, inner only, none
- [ ] Table alignment (left / centre / right) and table width (auto / fixed / percentage of text area) are settable
- [ ] Merged cells, borders, shading and widths round-trip through `.docx` (`w:gridSpan`, `w:vMerge`, `w:tcBorders`, `w:shd`) and `.odt` (`table:number-columns-spanned`, `table:covered-table-cell`, cell styles)
- [ ] A `.docx` fixture with vertically merged cells and mixed borders renders correctly and survives a save/re-open unchanged
- [ ] All operations are single undo steps

## Out of scope
Table styles / banded row themes. Repeat header row on page breaks. Text
wrapping around floating tables. Sorting and formulas.

## Note (SPEC-074)
No behaviour change: `BorderSide` and `BorderLineStyle` moved from
`src/lib/core/tables.ts` to `src/lib/core/borders.ts` (paragraphs and pages
carry the same edges, so the types are not a table concept) and are
re-exported from `core/tables.ts`, so every import and every test in this
spec compiles untouched. `BorderLineStyle` gained `thick`, `dashDot` and
`dashDotDot`, which cell borders may now also carry.
