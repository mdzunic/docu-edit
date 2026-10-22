---
id: SPEC-025
title: Hyperlinks and bookmarks
status: draft
area: structure
depends_on: [SPEC-016]
---
## Why
Links are the most common inline object after plain text, and internal links
plus bookmarks are what a table of contents and cross-references will later be
built on.

## Acceptance criteria
- [ ] Ctrl/Cmd+K with text selected opens a link dialog pre-filled with the selected text and an empty URL; with no selection, both fields are empty
- [ ] Saving the dialog wraps the text in a hyperlink, rendered underlined in the link colour
- [ ] Typing or pasting a bare URL followed by space or Enter autolinks it, and Ctrl/Cmd+Z immediately after reverts to plain text
- [ ] Clicking a link with the caret shows a link chip with the target, plus "Edit", "Copy link" and "Remove link"
- [ ] Ctrl/Cmd+click (or the chip's open action) opens the target in a new tab with `rel="noopener noreferrer"`
- [ ] Links to `javascript:`, `data:` and `vbscript:` URLs are rejected on entry and stripped on open/paste, covered by a security test
- [ ] The dialog offers linking to a heading in this document or to a bookmark, listing both
- [ ] An "Insert bookmark" action names a location in the document; a bookmark manager lists, jumps to and deletes bookmarks
- [ ] Bookmark names are validated (no spaces, unique) with an inline error on conflict
- [ ] Links and bookmarks round-trip through `.docx` (`w:hyperlink` with external relationships or `w:anchor`, `w:bookmarkStart/End`) and `.odt` (`text:a`, `text:bookmark`)
- [ ] Removing a link leaves the text and its other formatting intact

## Out of scope
Cross-references with automatic numbering ("see Figure 3"). Link preview cards.
Checking whether links are alive. Email/`mailto:` composition UI beyond
accepting the URL.
