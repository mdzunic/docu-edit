---
id: SPEC-017
title: Find and replace
status: draft
area: editing
depends_on: [SPEC-011, SPEC-015]
---
## Why
The first tool anyone uses on a document someone else wrote. It is also a
cheap, high-visibility win once the editing core exists.

## Acceptance criteria
- [ ] Ctrl/Cmd+F opens a find bar with a query field, match count ("3 of 17"), previous/next buttons and a close button
- [ ] Matches are highlighted in the document as the user types, with the current match highlighted distinctly and scrolled into view
- [ ] Enter and Shift+Enter step forward and backward through matches, wrapping at the end with a "Wrapped to top" hint
- [ ] Escape closes the bar, clears highlights and returns the caret to where it was
- [ ] Ctrl/Cmd+H (or an "expand" toggle) reveals a replace field with "Replace" and "Replace all"
- [ ] "Replace all" reports "Replaced N occurrences" and is a single undo step
- [ ] Options for "Match case", "Whole words only" and "Use regular expressions" are available and persist for the session
- [ ] An invalid regular expression shows an inline error and does not run
- [ ] Replacement text takes the formatting of the first run of the text it replaced
- [ ] Search covers body text, and clearly states its scope; matches inside headers, footers and footnotes are excluded until those specs exist
- [ ] Searching a 200-page document returns first results within 300 ms and does not block typing

## Out of scope
Find by formatting or style. Search across multiple documents. Regex capture
groups in the replacement string (`$1`) — noted as a follow-up.
