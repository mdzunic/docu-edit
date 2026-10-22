---
id: SPEC-020
title: Clear formatting and format painter
status: draft
area: formatting
depends_on: [SPEC-013, SPEC-019]
---
## Why
Pasted content arrives with formatting nobody wants. These two tools are how
users clean it up without retyping — and the format painter is a standing
expectation from both Word and Docs.

## Acceptance criteria
- [ ] A "Clear formatting" toolbar button removes all direct character formatting from the selection and resets its paragraphs to the Normal style
- [ ] Ctrl/Cmd+\ performs the same action
- [ ] Clear formatting does not remove content, links, list membership, comments or tracked changes
- [ ] A "Format painter" button copies the character formatting at the caret (and the paragraph formatting when a whole paragraph is selected)
- [ ] With the painter armed, the cursor changes and the next selection made receives that formatting, after which the painter disarms
- [ ] Double-clicking the painter keeps it armed for repeated application until Escape or another click on the button
- [ ] Escape at any time disarms the painter
- [ ] Both operations are single undo steps

## Out of scope
Copying table or image formatting. A persistent "formatting clipboard" with
multiple slots.
