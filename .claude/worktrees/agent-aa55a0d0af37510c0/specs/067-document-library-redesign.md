---
id: SPEC-067
title: The document library — overview redesign
status: draft
area: shell
depends_on: [SPEC-003, SPEC-058, SPEC-060]
---
## Why
The library is the first screen of the product and the least designed one.
`src/views/HomeView.vue` renders a three-column table — Title, Last
modified, Format — with a per-row overflow menu holding Rename, Duplicate
and Delete, and three buttons above it (New document, New from template,
Open document). That was the right shape for SPEC-003, when a library held
the two or three documents you had just made. It does not survive fifty:
there is no search, no sort, no filter, no way to select more than one row,
no visual distinction between a fifty-page report and an empty draft, and
no way back from a delete.

Two other specs make this urgent rather than cosmetic. SPEC-060 gives the
editor a Close button, so the library stops being a place you pass through
once and becomes the place you return to all day. And SPEC-046's template
gallery already proves the product can render a preview of a document —
the library shows none.

The gap that matters most is the irreversible one. `deleteDocument(id)` in
`src/lib/documentStore.ts` removes the record, and the confirmation dialog
is the only thing between a click and permanently losing a document that
exists nowhere else, because this product's storage *is* the user's disk.
A library with multi-select and no undo would be worse than one without
multi-select, so a trash is part of the same spec, not a follow-up.

## Acceptance criteria
- [ ] The library offers two views — a list and a grid of document cards with page-one thumbnails — toggled from a control in the library header, with the choice persisted in preferences (SPEC-066)
- [ ] A card shows a rendered thumbnail of the document's first page, its title, format badge, relative modified time, and page and word counts, per "Thumbnails"
- [ ] A search field filters by title as you type, matching by subsequence, with matched characters highlighted; clearing it restores the full list
- [ ] Sorting is offered by Last modified (default, newest first), Title (A–Z) and Size, ascending or descending, persisted
- [ ] Filters are offered by format (`docx` / `odt` / `md` / new) and by starred, shown as toggle chips with live counts
- [ ] A document can be starred from its row or card and from its overflow menu; starred documents sort first within the active sort, and starring is stored per document
- [ ] Multi-select works by checkbox, Shift+click for a range and Ctrl/Cmd+click for individual rows, with Ctrl/Cmd+A selecting all matching the current filter
- [ ] With a selection, a bulk action bar offers Download, Duplicate, Star and Move to trash, showing the count and offering Clear selection
- [ ] Deleting moves documents to a trash rather than erasing them: a toast offers immediate Undo, and `Trash` (reached from the library header) lists trashed documents with Restore and Delete permanently, per "Trash"
- [ ] Trashed documents are excluded from the library, from search, and from anything that enumerates documents; they keep their id, so restoring returns them intact with their version history (SPEC-040)
- [ ] `Delete permanently` and `Empty trash` confirm once, name what will be lost, and are the only paths that call the destructive store operation
- [ ] The library shows storage usage — how much of the browser's quota the library occupies, and which documents are largest — and links to trash when usage is high, extending the SPEC-003 storage warning rather than replacing it
- [ ] The empty state is a first-run screen: what the product does, the three ways to start (blank, template, open a file), and a note that documents are stored in this browser
- [ ] A "no results" state for a search or filter that matches nothing offers to clear the filter, and is visually distinct from the empty library
- [ ] Rows and cards are a single keyboard-navigable list: arrows move, Enter opens, Space toggles selection, Delete moves to trash, F2 renames, and the overflow menu is reachable without the pointer
- [ ] Opening a document from the library is one click on the row or card — not on the title text alone — with a visible focus and hover state on the whole target
- [ ] Every existing `data-testid` in the library is preserved: `library`, `library-list`, `library-row`, `row-title`, `row-modified`, `row-format`, `row-menu`, `row-menu-list`, `menu-rename`, `menu-duplicate`, `menu-delete`, `empty-state`, `new-document`, `new-from-template`, `open-document`
- [ ] Thumbnail rendering never blocks the library: rows and cards render immediately with placeholders, and thumbnails fill in as they are produced, per "Thumbnails"
- [ ] A library of 200 documents renders its first screen without a visible delay and scrolls smoothly, with off-screen thumbnails not rendered at all

## Out of scope
Folders, tags or any nesting — starring plus search is the whole
organisation model here; folders are a bigger spec that needs a
`StoredDocumentRecord` change. Sharing, links or collaboration entry points
(SPEC-044 owns collaboration). Cloud or remote storage of any kind. A
recently-opened list separate from the library's own sort. Bulk *editing*
of documents. Renaming from the grid card's title inline (the overflow
menu's Rename dialog stays the one path). Changing the open, create,
duplicate or rename behaviours themselves — SPEC-003 owns those.

## Reference

### What ships today
`src/views/HomeView.vue` renders either `[data-testid="empty-state"]` or
`[data-testid="library"]` containing `[data-testid="library-list"]` of
`[data-testid="library-row"]` rows, each with `row-title`, `row-modified`,
`row-format` and a `row-menu` button opening `row-menu-list` with
`menu-rename`, `menu-duplicate`, `menu-delete`. The header holds
`new-document`, `new-from-template` and `open-document`, plus a hidden
`open-document-input` file input and an `open-document-error` region.

`src/lib/documentStore.ts` is the data layer:

```ts
export type DocumentFormat = "docx" | "odt" | "md" | "new";
export interface LibraryEntry { … }            // what listDocuments returns
export interface StoredDocumentRecord { … }    // the persisted shape
export function listDocuments(): LibraryEntry[];
export function getDocument(id: string): DocumentModel | undefined;
export function getStoredRecord(id: string): StoredDocumentRecord | undefined;
export function touchDocument(id: string): void;
export function renameDocument(…): …;
export function duplicateDocument(…): …;
export function deleteDocument(id: string): void;
export function isStorageAvailable(): boolean;
export function onStorageStatusChange(listener: () => void): () => void;
```

This spec adds to that module rather than around it, so the library view
stays a view.

### Record additions
```ts
// src/lib/documentStore.ts
export interface StoredDocumentRecord {
  …                                  // unchanged fields
  /** Starred by the user. Absent means not starred. */
  starred?: boolean;
  /** When the document was moved to the trash, as an epoch millisecond
   *  value. Absent means it is in the library. A record with this set is
   *  excluded from listDocuments() and from every enumeration. */
  trashedAt?: number;
}

export function listDocuments(): LibraryEntry[];          // excludes trashed
export function listTrashedDocuments(): LibraryEntry[];   // trashed only
export function trashDocument(id: string): void;          // sets trashedAt
export function restoreDocument(id: string): void;        // clears it
export function deleteDocument(id: string): void;         // unchanged: permanent
export function setStarred(id: string, starred: boolean): void;
export function libraryStorageUsage(): { totalBytes: number; byDocument: Array<{ id: string; bytes: number }> };
```

`deleteDocument` keeps its exact current meaning — permanent removal — so
no existing caller changes behaviour. What changes is *who calls it*: after
this spec only "Delete permanently" and "Empty trash" do, and the library's
delete affordance calls `trashDocument`. `useFileActions.ts`'s
`deleteOpenDocument` (the editor's File → Delete) moves to `trashDocument`
too, keeping its toast and its navigation.

Trash is not auto-emptied on a timer. A document the user cannot see and
did not delete permanently must not disappear on its own; the storage-usage
panel is how they find out it is still costing space.

### Thumbnails
The product can already render a document's first page — SPEC-034's PDF
export and SPEC-046's template gallery both do. A library thumbnail is the
same operation with a hard constraint: it must never make opening the
library slow.

- Thumbnails render **lazily and off the critical path**, via an
  `IntersectionObserver` so only visible cards render, scheduled through
  the existing idle scheduler (SPEC-047's `idleScheduler`, whose contract
  includes not keeping the process alive).
- Each is cached in the store keyed by document id and the record's
  modified timestamp, so it is produced once per edit, not once per view.
  A cached thumbnail whose timestamp no longer matches is discarded and
  re-rendered.
- The placeholder is a paper-shaped skeleton at the correct aspect ratio,
  so nothing reflows when the real one arrives.
- A document that fails to render (corrupt, or over SPEC-047's ceiling)
  shows a format-badge placeholder, never an error.
- Thumbnails are a cache: clearing them must be safe, and quota pressure
  evicts them before it evicts anything else.

### Layout
The header row holds, left to right: the library title, the search field,
the sort control, the filter chips, the view toggle (list/grid), and the
three primary actions with New document as SPEC-058's `primary` variant.
Below it, either the list (a `role="grid"` of rows, keeping today's
columns plus a star and a checkbox) or the card grid (a `role="grid"` of
cards, auto-filling at a minimum card width). Trash is a link in the
header showing its count, opening a view of the same shape with Restore
and Delete permanently as the row actions.

Selection state lives in the view; nothing about selection is persisted.

### Undo
The delete toast's Undo calls `restoreDocument` for every id in the batch.
It stays available for the toast's lifetime; after that, Trash is the undo.
This is deliberately a *shallow* undo — one action, from the toast — rather
than an undo stack for the library, because the trash is the durable
safety net and two overlapping mechanisms would be worse than one.

### Specs this changes
- `specs/003-local-document-library.md` — the library gains search, sort,
  filter, starring, multi-select, a grid view and a trash. Creating,
  opening, renaming and duplicating are unchanged, as is the storage
  layer's own contract. Its delete AC is amended: delete moves to trash,
  and permanent deletion is a second, explicit step.
- `specs/060-close-document-and-return-to-library.md` — the returning
  highlight lands on a row or card, whichever view is active.
- `specs/046-template-gallery-and-mail-merge.md` — the template gallery's
  preview rendering is the precedent this reuses; nothing in it changes.
- `specs/047-large-document-performance.md` — thumbnail work is scheduled
  through the existing idle scheduler and must not move the baseline.
