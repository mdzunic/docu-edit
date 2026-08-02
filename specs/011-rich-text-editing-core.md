---
id: SPEC-011
title: Rich text editing core
status: draft
area: editing
depends_on: [SPEC-002, SPEC-004]
---
## Why
This is the spine of the product: caret, selection, and text mutation over the
document model. Every formatting and structure feature after this one is a
command that runs against the engine defined here, so its correctness sets the
ceiling for everything else.

## Acceptance criteria
- [ ] Clicking anywhere in the page canvas places a caret at the nearest text position and the caret is visibly rendered
- [ ] Typing inserts characters at the caret and the document model updates to match what is on screen
- [ ] Backspace and Delete remove the character before/after the caret; at a paragraph boundary they merge paragraphs
- [ ] Enter splits the current paragraph at the caret, carrying paragraph formatting to the new paragraph; Shift+Enter inserts a line break within the paragraph
- [ ] Arrow keys, Home/End, Ctrl/Cmd+Home/End and PageUp/PageDown move the caret; Shift with any of them extends the selection
- [ ] Ctrl/Cmd+Left/Right moves by word; double-click selects a word; triple-click selects a paragraph; Ctrl/Cmd+A selects the whole document
- [ ] Dragging with the mouse creates a selection that is visibly highlighted, including across paragraphs
- [ ] Typing with a non-empty selection replaces the selection
- [ ] All mutations go through a documented command API (e.g. `insertText`, `deleteRange`, `splitBlock`) that returns a new model state — the DOM is never the source of truth
- [ ] IME composition (e.g. Japanese, Chinese, Korean input) commits correctly and does not corrupt the model, verified by a test that drives composition events
- [ ] Dead keys and accented input on a European keyboard layout produce single correct characters
- [ ] A screen reader reads the document content and announces caret movement — the editable region exposes a document/textbox role and its text content

## Out of scope
Formatting commands (SPEC-012 onward). Undo/redo (SPEC-015). Clipboard
(SPEC-016). Tables and images. Collaborative cursors (SPEC-044).
