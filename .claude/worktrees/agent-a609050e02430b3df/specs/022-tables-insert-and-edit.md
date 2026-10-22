---
id: SPEC-022
title: Tables — insert and edit
status: draft
area: structure
depends_on: [SPEC-011, SPEC-015]
---
## Why
Tables are the most-requested structure after lists, and the first block-level
container in the model — nested block content, cell-scoped selection, and
navigation that has to feel like Word's.

## Acceptance criteria
- [ ] An "Insert table" toolbar control opens a grid picker for up to 10×10 plus an "Insert table…" dialog for larger sizes
- [ ] Inserting places a table at the caret with equal column widths spanning the text area, and puts the caret in the first cell
- [ ] Cells accept all existing editing and formatting features: typing, character formatting, paragraph formatting, lists and styles
- [ ] Tab moves to the next cell, Shift+Tab to the previous, and Tab in the last cell appends a new row
- [ ] Arrow keys move between cells at cell boundaries; the caret never gets trapped in a cell
- [ ] Clicking and dragging across cells selects a rectangular cell range, visibly highlighted
- [ ] A row/column context menu offers insert above/below, insert left/right, and delete row/column/table
- [ ] Row and column handles appear on hover for insert and delete without opening a menu
- [ ] Deleting all rows deletes the table; the caret lands where the table was
- [ ] Tables round-trip through `.docx` (`w:tbl` with grid and cell widths) and `.odt` (`table:table` with column definitions), preserving row and column counts and cell content
- [ ] Opening a `.docx` fixture with a table renders it with the right dimensions and content, including a table with an empty cell and one containing multiple paragraphs

## Out of scope
Borders, shading, merged cells and column resizing (SPEC-023). Nested tables.
Table styles. Sorting, formulas and header-row repetition across pages.
