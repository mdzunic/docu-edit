---
id: SPEC-047
title: Large document performance
status: draft
area: platform
depends_on: [SPEC-028, SPEC-033, SPEC-035]
---
## Why
Everything in this roadmap works at ten pages. The documents that make people
abandon a browser editor are the 300-page tender responses and the 500-row
tables. This spec sets the budgets and makes the product hold them.

## Acceptance criteria
- [ ] Documented performance budgets, enforced by an automated benchmark in CI against a 300-page fixture with 500 images, 200 tables and 400 footnotes
- [ ] Opening that fixture shows the first page in under 2 seconds and becomes fully interactive in under 8 seconds, with a progress indicator until it is
- [ ] Typing latency stays under 50 ms at the 95th percentile anywhere in that document, including at the end
- [ ] Only visible pages plus a small buffer are rendered to the DOM; scrolling stays above 30 fps and scroll position is stable when off-screen pages are recycled
- [ ] Parsing, pagination, spell check, word count and export run off the main thread (worker) so the UI never blocks for more than 100 ms
- [ ] Memory use for the benchmark fixture stays under a documented ceiling and does not grow during a 30-minute editing session — a leak test asserts it
- [ ] Saving that fixture to `.docx` completes in under 5 seconds and does not freeze the UI
- [ ] The benchmark runs in CI on every PR touching the editor core and fails the build on a regression beyond a stated threshold
- [ ] Behaviour when a document exceeds the supported size is graceful: a warning naming the limit, with the document still openable read-only

## Out of scope
Server-side rendering or pre-pagination. Streaming partial documents from a
server. Mobile device performance targets.
