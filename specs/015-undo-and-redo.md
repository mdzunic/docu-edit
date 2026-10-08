---
id: SPEC-015
title: Undo and redo
status: done
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
- [ ] Every command introduced by any later spec is undoable — a shared test asserts that applying then undoing any registered command returns the model to an `equalStructural` prior state (see "Equality" below)

## Out of scope
Persisting undo history across reloads. Selective/"undo just this change"
history. Version history (SPEC-040).

## Reference

### Equality
`equalStructural` is defined normatively in
`specs/002-create-blank-document.md` ("Equality") and exported from
`src/lib/equality.ts`: deep equality after canonicalization, which drops
runtime `id` fields and default-valued props and merges adjacent runs with
equal formatting. It is the right relation for undo because a command may
legitimately rebuild nodes with new ids or split and re-merge runs while
producing the identical document.
