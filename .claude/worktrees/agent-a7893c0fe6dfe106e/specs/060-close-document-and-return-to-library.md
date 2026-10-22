---
id: SPEC-060
title: Close the document and return to the library
status: draft
area: shell
depends_on: [SPEC-003, SPEC-010, SPEC-049]
---
## Why
There is no way out of a document. `src/components/AppShell.vue`'s app bar
holds the product name, the title, a format badge and a slot; the editor
fills that slot with the save-state indicator and the Download menu. No
control anywhere in the editor returns to the library. The only
`router.push("/")` in the whole editing surface is in
`src/views/document/useFileActions.ts`, at the end of `deleteOpenDocument`
— so today the sanctioned way back to the document list is to delete the
document you are in.

Everything else works around that hole. `.product-name` is a `<span>`, not
a link, so the one element every web app makes a home button is inert. The
"Document not found" state has a `RouterLink to="/"` reading "Back to
home", which is the only such link in the view and proves the destination
exists — it is simply unreachable when the document *is* found. Users are
left with browser Back, which works by accident of the router and gives no
signal about whether their work is safe.

That last point is the real content of this spec. Leaving a document is
not navigation, it is a **commit boundary**: SPEC-010 autosaves on a
debounce, so the moment a user leaves is exactly the moment a pending
write may not have landed. Closing must flush, not race. Get that right
once, in one place, and every other exit — the browser's Back button, a
closed tab, an opened second document — inherits it.

## Acceptance criteria
- [ ] A Close control sits in the app bar of the editor, labelled "Close document", icon-only with a tooltip, and returns to the library at `/`
- [ ] `File → Close` performs the same command from the menu bar, registered in the `editorCommands` Map as `file.close`
- [ ] The product name in the app bar becomes a link to `/` in every view, with the same guard as the Close control, and reads as a link to assistive technology
- [ ] Closing **flushes before navigating**: any pending autosave is awaited and committed, per "Closing is a commit boundary", and the library shows the document with its just-saved content and modified time
- [ ] While the flush is in flight the Close control shows a busy state and the app bar's save indicator reads "Saving…"; navigation happens when the flush resolves, and never before
- [ ] A flush that fails does not navigate: an error toast names the failure, the document stays open with its content intact, and the user can retry or download
- [ ] A document whose content cannot be saved at all (storage unavailable, per SPEC-003) warns before closing with a dialog offering Download, Close anyway, and Cancel — the only path that can lose work, made explicit
- [ ] Browser Back out of the editor takes the same flush path via a route guard, so the accidental exit is as safe as the deliberate one
- [ ] Closing the tab or window with unflushed changes attempts the same commit through the existing unload handling, and never shows a browser "leave site?" prompt for changes that are already safely persisted
- [ ] After closing, the library scrolls to and highlights the row for the just-closed document for one `--motion-base` interval, so returning shows where you came from
- [ ] Escape does **not** close the document, at any time — Escape belongs to the innermost open surface (dialog, menu, shape selection, find bar) and never escalates to navigation
- [ ] The library's app bar shows no Close control and no document title, unchanged from today
- [ ] A recovered-draft banner (SPEC-010) blocks nothing: closing with an unresolved recovery prompt keeps the draft, so reopening still offers recovery
- [ ] The Close control is reachable by keyboard in app-bar tab order, before the Download menu, and is announced with its label
- [ ] An e2e test edits a document, closes it, and asserts the library row shows the edited title and content, with no wait or retry in the assertion

## Out of scope
A keyboard shortcut for Close — Ctrl/Cmd+W is reserved to the browser by
`RESERVED_BROWSER_KEYS` in `src/lib/shortcuts.ts` (SPEC-021 AC-11), and
claiming any other combination for "leave the document" is worse than
having none. Multiple documents open at once, tabs, or a window manager.
"Save a copy and close", "Close without saving" as a discard command —
this product autosaves, and discarding is SPEC-040's version-history
territory. Changing what the library looks like (SPEC-067) or where
documents are stored (SPEC-003). Session restore of the last-open
document.

## Reference

### The shell and the route table
`specs/001-app-shell-and-document-workspace.md` defines the shell;
`src/router.ts` holds the routes (CLAUDE.md: declared there, never
inline). As shipped:

- `/` → `src/views/HomeView.vue` — the library.
- `/document/:id` → `src/views/DocumentView.vue` — the editor.
- the open-from-URL view (SPEC-009).

`AppShell.vue` renders `header.app-bar[data-testid="app-bar"]` containing
`span.product-name[data-testid="product-name"]`, the
`h1[data-testid="document-title"]` / `input[data-testid="document-title-input"]`
pair, an optional `span.format-badge`, and `div.app-bar-actions` filled by
the `app-bar-actions` slot. This spec adds one control to the *start* of
the bar and turns `.product-name` into a `RouterLink`; it adds no region.

The Close control renders left of the product name — the position every
"leave this thing" control occupies — as
`data-testid="close-document"`, `aria-label="Close document"`, using
SPEC-058's `UiIconButton` with the `close` icon. In views with no open
document the slot is empty and the bar is unchanged.

### Closing is a commit boundary
`specs/010-autosave-and-crash-recovery.md` defines the persistence model
this must cooperate with, and `src/lib/documentStore.ts` its API:

```ts
export async function saveDraft(...): Promise<void>
export async function loadDraft(id: string): Promise<DraftRecord | undefined>
export async function deleteDraft(id: string): Promise<void>
export async function commitDocument(...): Promise<void>
export function listDocuments(): LibraryEntry[]
export function touchDocument(id: string): void
export function deleteDocument(id: string): void
```

Autosave is debounced, so at any instant there may be an edit that is in
the model and not yet in storage. The close path therefore is:

1. Cancel the pending debounce and run the write it was waiting to do —
   **await** it. There is no "fire and forget then navigate": the library
   reads storage, and a race here shows a stale row, which is the exact
   bug the memory of this project already records for one-shot reads of
   deferred work.
2. On success, `router.push("/")`.
3. On failure, stay, toast the error, leave the model untouched.

The flush lives in `src/views/document/useFileActions.ts` alongside the
existing `deleteOpenDocument`, exposed as `closeOpenDocument()`, and is the
single implementation the app-bar control, the `File → Close` command, the
product-name link and the route guard all call. Nothing calls
`router.push("/")` directly from the editor after this spec.

**The route guard.** `onBeforeRouteLeave` in the editor view runs the same
`closeOpenDocument()` flush for any navigation away from `/document/:id`
— including browser Back — and cancels the navigation when the flush
fails. Deletion is the one exception: `deleteOpenDocument` already removed
the record, so its navigation skips the flush.

**Tab close.** The existing unload path is unchanged in mechanism: a
synchronous best-effort write, no `beforeunload` prompt when there is
nothing unsaved. Adding a prompt for state that is already persisted is
what makes autosaving apps feel unsafe, and this spec explicitly does not.

### The storage-unavailable case
SPEC-003 already handles blocked storage with `isStorageAvailable()` and
the `StorageWarningBanner`. When storage is unavailable the document
exists only in memory, so closing *is* data loss. That case — and only
that case — gets a confirmation dialog, built on `BaseDialog.vue`:

> **Close without saving?** Storage is unavailable in this browser, so
> this document has not been saved. Closing it will lose your changes.
> — [Download] [Close anyway] [Cancel]

Download runs the existing download flow (SPEC-006/007) and leaves the
dialog open, so a user can save the file and then close deliberately.

### Returning to the library
`router.push("/")` carries the closed document's id, so `HomeView.vue`
can scroll its row into view and flash it. The library's shipped testids
(`library-list`, `library-row`, `row-title`, `row-modified`) are unchanged;
the highlight is a transient class, timed to `--motion-base` and
suppressed under `prefers-reduced-motion`. Passing the id as router state
rather than a query parameter keeps it out of the URL — a document id in a
shareable URL is a privacy leak the app has no reason to create.

### Specs this changes
- `specs/001-app-shell-and-document-workspace.md` — the app bar gains a
  leading Close control slot and the product name becomes a link. No
  region is added or removed.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the File menu gains
  Close, at the end of the menu after a separator, per that spec's rule
  that a new command names its menu.
- `specs/010-autosave-and-crash-recovery.md` — no mechanism change; this
  spec names close as a flush point for the debounce that spec defines.
- `specs/021-keyboard-shortcuts-and-reference.md` — unchanged. Close
  claims no binding, deliberately.
