---
id: SPEC-015
title: Undo and redo
status: draft
area: editing
depends_on: [SPEC-011, SPEC-012]
---
## Why
Users experiment only as far as they trust undo. Getting the grouping right —
one Ctrl+Z per intention, not per keystroke — is the difference between an
editor that feels native and one that feels like a web toy.

## Acceptance criteria
- [ ] Ctrl/Cmd+Z undoes and Ctrl/Cmd+Shift+Z (and Ctrl+Y on Windows) redoes; toolbar buttons do the same
- [ ] Both buttons are disabled when their stack is empty, with tooltips "Undo <action>" / "Redo <action>" naming the operation, e.g. "Undo typing"
- [ ] Consecutive typed characters group into one undo step, broken by a caret move, a pause of 2 seconds, or a non-typing command
- [ ] A formatting command is always its own undo step
- [ ] Undo restores the selection that existed before the undone operation
- [ ] Performing a new edit after undo clears the redo stack
- [ ] The history holds at least 200 steps and is bounded so a long session cannot exhaust memory
- [ ] Undo history is per document and is cleared when the document is closed
- [ ] Every command introduced by any later spec is undoable — a shared test asserts that applying then undoing any registered command returns the model to a deep-equal prior state

## Out of scope
Persisting undo history across reloads. Selective/"undo just this change"
history. Version history (SPEC-040).
