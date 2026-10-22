---
id: SPEC-044
title: Real-time collaborative editing
status: draft
area: collab
complexity: complex
story_points: 21
depends_on: [SPEC-011, SPEC-038, SPEC-040]
---
## Why
"Like Google Docs" ultimately means two people typing in the same document at
once. This is the largest single feature in the roadmap and the first that
requires a server, so it comes after the single-user product is solid.

## Acceptance criteria
- [ ] A "Share" action produces a link that opens the same document for another user in edit or view-only mode, chosen when the link is created
- [ ] Two users editing simultaneously see each other's changes within 500 ms on a normal connection, with no lost or duplicated characters
- [ ] Concurrent edits to the same paragraph converge to the same document for every participant — verified by a fuzz test that applies randomised concurrent operation sequences and asserts convergence
- [ ] Each participant's caret and selection are shown in their own colour with their name on a label
- [ ] An avatar row shows who is present; hovering shows the name and clicking jumps to their caret
- [ ] Comments and comment replies sync in real time, and a comment made while offline appears once reconnected
- [ ] Losing the connection switches to an offline indicator; edits continue locally and merge on reconnect without a "reload to continue" dead end
- [ ] Formatting, list, table and image operations are all collaborative, not just text insertion
- [ ] A user without edit permission cannot mutate the document through any client path, enforced server-side and covered by a test
- [ ] Document content is only readable with the share link or an explicit grant, and link permissions can be revoked
- [ ] Version history (SPEC-040) attributes each version to the users who contributed to it

## Out of scope
Presence in the file library. Chat. Suggesting mode as a collaboration mode
(track changes covers redlining). Per-paragraph permissions. Offline editing on
two devices merging across a week-long gap.
