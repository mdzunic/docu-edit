---
id: SPEC-026
title: Special characters and symbol picker
status: draft
area: editing
depends_on: [SPEC-011]
---
## Why
Every real document eventually needs a — an ™ or a °, and typing them is
awkward on most keyboards. Small feature, constant use, and it forces correct
handling of characters outside the Basic Multilingual Plane.

## Acceptance criteria
- [ ] An "Insert → Special characters" dialog shows a searchable grid of characters grouped by category (punctuation, currency, maths, arrows, Latin, Greek, symbols, emoji)
- [ ] Searching by name ("degree", "em dash") or by hex code point filters the grid
- [ ] Clicking a character inserts it at the caret and keeps the dialog open for another insertion
- [ ] Recently used characters appear in a "Recent" row that persists across sessions
- [ ] Non-breaking space (Ctrl/Cmd+Shift+Space), em dash, en dash and non-breaking hyphen have direct insert entries
- [ ] Autocorrect converts `--` to an en dash and straight quotes to typographic quotes as the user types, and each conversion is individually revertible with Ctrl/Cmd+Z
- [ ] An autocorrect setting in the app menu turns the smart-quote and dash substitutions off, and the setting persists
- [ ] Characters outside the BMP (e.g. 𝑥, 😀) insert, render, select and delete as single units, and survive a save/re-open in both formats
- [ ] A non-breaking space prevents a line break at that point in rendering and is written as the correct character in both formats

## Out of scope
Equation/formula editing. Handwriting or drawing input. Custom autocorrect
dictionaries.
