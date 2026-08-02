---
id: SPEC-016
title: Cut, copy and paste
status: draft
area: editing
depends_on: [SPEC-011, SPEC-015]
---
## Why
Most content arrives by paste — from Word, from a browser, from an email. Paste
is where formatting fidelity is judged, and where a naive implementation
injects junk markup that poisons the saved file.

## Acceptance criteria
- [ ] Ctrl/Cmd+X, Ctrl/Cmd+C and Ctrl/Cmd+V cut, copy and paste, as do the context menu entries
- [ ] Copying puts three flavours on the clipboard: plain text, HTML, and an internal document-model payload
- [ ] Pasting content copied from this editor restores formatting, styles, lists and structure exactly
- [ ] Pasting HTML from Word Online, Google Docs or a web page maps headings, bold/italic/underline, lists, links and tables to the document model, and drops scripts, classes, positioning and unknown markup
- [ ] Ctrl/Cmd+Shift+V pastes as plain text, taking the formatting of the paragraph at the caret
- [ ] A paste options chip appears after a paste with "Keep source formatting", "Match destination formatting" and "Keep text only"; choosing one re-applies the paste as a single undo step
- [ ] Pasting into a selection replaces it; pasting multiple paragraphs into the middle of a paragraph splits it correctly
- [ ] Cut removes the content and places it on the clipboard in one undo step
- [ ] Pasted content is sanitised: no `javascript:` URLs, event handlers, `<script>`, `<style>` or embedded remote content survives into the model — covered by an explicit security test
- [ ] Pasting an image from the clipboard is accepted once SPEC-024 is merged, and until then shows "Images aren't supported yet"

## Out of scope
Drag-and-drop of selected text within the document. Clipboard history. Paste of
Excel ranges as live-linked tables.
