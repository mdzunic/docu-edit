---
id: SPEC-009
title: Drag-and-drop and open-from-URL
status: draft
area: io
depends_on: [SPEC-004, SPEC-005]
---
## Why
The file picker is the slowest way to open a document. Dragging a file onto the
window is how people actually do it, and an openable URL is what makes the
product linkable from other tools.

## Acceptance criteria
- [ ] Dragging a file over the app window shows a full-window drop overlay reading "Drop to open"
- [ ] Dropping a `.docx` or `.odt` file opens it, exactly as the file picker would
- [ ] Dropping an unsupported file type shows the error "We can only open .docx, .odt and .md files" and does not navigate away
- [ ] Dropping multiple files opens the first and shows a toast "Opened <name> — only one file at a time"
- [ ] Dragging over the window while a document is open does not modify the open document until the drop completes successfully
- [ ] `/open?url=<encoded-url>` shows the address it is being asked to open — the host on its own line, then the full URL — and fetches it only once the reader confirms, showing a loading state while fetching
- [ ] `/open?url=` refuses any scheme other than `http`/`https` before fetching, so a link cannot smuggle a document in through `data:`/`blob:`
- [ ] A fetch that fails, is blocked by CORS, or returns a non-document content type shows "Couldn't open that link" with the underlying reason and a button to pick a local file instead
- [ ] Files larger than 50 MB are refused before parsing with "That file is too large to open (limit 50 MB)"
- [ ] The document opened from a URL is stored in the local library like any other, and is not re-fetched on reload

## Out of scope
Cloud-drive integrations (Google Drive, OneDrive, Dropbox). Opening a folder.
Uploading to a server.
