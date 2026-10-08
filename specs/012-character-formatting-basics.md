---
id: SPEC-012
title: Character formatting — bold, italic, underline, strikethrough
status: done
area: formatting
depends_on: [SPEC-011]
---
## Why
The four buttons everyone looks for first. They also establish the pattern for
every later toolbar control: a command that applies to the selection, a button
whose pressed state reflects the caret.

## Acceptance criteria
- [ ] The toolbar has Bold, Italic, Underline and Strikethrough buttons with tooltips showing name and shortcut
- [ ] Ctrl/Cmd+B, Ctrl/Cmd+I, Ctrl/Cmd+U and Ctrl/Cmd+Shift+X apply the matching format
- [ ] Applying a format to a selection formats exactly the selected text, splitting runs at the selection boundaries
- [ ] Applying a format with a collapsed caret sets a pending format: the next typed characters carry it, and moving the caret discards it
- [ ] Each button shows a pressed state when the caret or the entire selection has that format
- [ ] With a mixed selection (some text bold, some not) the button shows an indeterminate state, and the first click makes the whole selection bold
- [ ] Applying a format twice removes it (toggle behaviour), including on a mixed selection where the second click clears it everywhere
- [ ] Formats round-trip through both writers: applying bold, saving `.docx` and `.odt`, and re-opening preserves it
- [ ] Formatting applied over a run boundary produces the minimum number of runs — adjacent runs with identical formatting are merged after every command

## Out of scope
Font family, size and colour (SPEC-013). Subscript/superscript, small caps and
other run properties (SPEC-013 out of scope note). Character styles (SPEC-042).
