---
id: SPEC-030
title: Multi-column layout
status: draft
area: layout
depends_on: [SPEC-028]
---
## Why
Newsletters, leaflets and academic papers need columns, and columns are the
main reason people fall back to LibreOffice. They also stress the layout engine
in a way that will surface pagination bugs early.

## Acceptance criteria
- [ ] A "Format → Columns" dialog offers 1, 2 or 3 equal columns, plus custom column count (up to 6), individual widths and spacing, and a "Line between columns" toggle
- [ ] Columns apply to the current section, or to the selection by wrapping it in a new continuous-break section, chosen by an "Apply to" selector
- [ ] Text flows down the first column then into the next, and across page boundaries into the first column of the next page
- [ ] "Insert → Break → Column break" moves the following content to the top of the next column
- [ ] The ruler shows column boundaries, and dragging them adjusts widths and gutters
- [ ] Images, tables and lists inside a column lay out within the column width, not the page width
- [ ] Columns round-trip through `.docx` (`w:cols` with `w:num`, `w:space`, `w:sep`, and explicit `w:col` widths for unequal columns) and `.odt` (`style:columns`)
- [ ] Opening a fixture with a two-column section and a single-column heading area renders both correctly
- [ ] Caret navigation with arrow keys moves sensibly between columns — down at the end of a column continues at the top of the next

## Out of scope
Balanced columns at the end of a section. Text flowing around floating objects
across columns. Newspaper-style linked text boxes.
