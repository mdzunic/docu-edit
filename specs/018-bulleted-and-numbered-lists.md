---
id: SPEC-018
title: Bulleted and numbered lists
status: done
area: formatting
depends_on: [SPEC-014, SPEC-015]
---
## Why
Lists are the most common structured content in business documents and the
single most common source of mangled `.docx` output, because numbering lives in
a separate part with its own indirection.

## Acceptance criteria
- [ ] Toolbar buttons toggle a bulleted list and a numbered list on every paragraph in the selection
- [ ] Tab and Shift+Tab inside a list item demote and promote it, up to 9 levels, and levels use distinct markers (•, ◦, ▪ / 1., a., i.) — "inside a list item" is the caret collapsed at the item's offset 0, per SPEC-054's Tab precedence table; a Tab pressed mid-item inserts a literal tab
- [ ] Enter on a list item creates the next item at the same level; Enter on an empty list item outdents it, and on an empty top-level item converts it to a normal paragraph
- [ ] Backspace at the start of a list item removes the list formatting from that item without deleting text
- [ ] Numbering is continuous across consecutive list items and restarts when a non-list paragraph interrupts the list
- [ ] Each list button has a dropdown to pick the marker style for the current level (bullet character, or decimal / lower-alpha / upper-alpha / lower-roman / upper-roman)
- [ ] Typing "1. " or "- " at the start of an empty paragraph autoformats it into the corresponding list, and Ctrl/Cmd+Z immediately after reverts to literal text
- [ ] Lists written to `.docx` produce valid `numbering.xml` with an abstract numbering definition and a `w:num` instance referenced by the paragraphs
- [ ] Lists written to `.odt` produce correct `text:list` nesting with a matching `text:list-style`
- [ ] Lists in opened `.docx` and `.odt` fixtures render with the correct markers, levels and indents, including a fixture with a list that starts at a number other than 1

## Out of scope
Multilevel list definitions and numbering restart controls (SPEC-032). Numbered
headings. Checkbox/task lists. Images as bullets.
