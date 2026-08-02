---
id: SPEC-003
title: Local document library
status: draft
area: io
depends_on: [SPEC-002]
---
## Why
A document that vanishes on reload is a demo, not a product. Users need to come
back to what they were working on without having exported a file first.

## Acceptance criteria
- [ ] Documents are stored in the browser (IndexedDB) keyed by document id, and survive a page reload
- [ ] `/` shows a library list of stored documents when at least one exists, with columns: title, last modified (relative, e.g. "2 hours ago"), and format badge (DOCX / ODT / New)
- [ ] The list is sorted by last modified, newest first
- [ ] Clicking a row opens that document at `/doc/:id` with its content intact
- [ ] Each row has an overflow menu with "Rename", "Duplicate", and "Delete"
- [ ] "Delete" asks for confirmation naming the document, and removes it from storage on confirm
- [ ] "Duplicate" creates a copy titled `<title> (copy)` that opens as a separate document
- [ ] Deleting the currently open document navigates back to `/` and shows a toast "Deleted <title>"
- [ ] When storage is unavailable or the quota is exceeded, the app still works in-memory and shows a persistent warning banner "Changes will not be saved in this browser"

## Out of scope
Server-side storage or accounts. Sharing. Folders, search, or tags in the
library. Version history (SPEC-040).

## Decisions
Rename should open modal. Validation rules - empty title not allowed, duplicate title gets number appended to the end like a count: new document (2). Max lenght: 200 chars. Renaming updates the last-modified timestamp.
