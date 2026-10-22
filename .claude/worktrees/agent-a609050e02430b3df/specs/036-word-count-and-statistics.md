---
id: SPEC-036
title: Word count and document statistics
status: draft
area: editing
depends_on: [SPEC-011, SPEC-028]
---
## Why
Anyone writing to a limit — students, journalists, bid writers — checks the
count constantly. It is a small feature that makes the product feel finished.

## Acceptance criteria
- [ ] The status bar shows a live word count that updates as the user types, throttled so it never delays input
- [ ] With a selection, the status bar shows "N of M words" for the selection and the document
- [ ] Clicking the count opens a statistics dialog with pages, words, characters (with and without spaces), paragraphs and sentences
- [ ] The dialog has a "Include footnotes and endnotes" toggle that changes the numbers live
- [ ] Word counting rules are documented and tested: hyphenated words count as one, numbers count as words, URLs count as one, and CJK characters are counted as characters with a separate CJK count shown
- [ ] Text inside headers, footers, comments and deleted tracked changes is excluded from the body count
- [ ] Counting a 200-page document completes in under 200 ms and is recomputed incrementally on edit, not from scratch
- [ ] The word count is reachable from the keyboard and announced by screen readers when it changes materially

## Out of scope
Readability scores. Per-section or per-heading counts. Writing goals and
progress tracking.
