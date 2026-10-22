---
id: SPEC-053
title: Fields — date, time, document information and cross-references
status: draft
area: structure
depends_on: [SPEC-025, SPEC-033, SPEC-041]
---
## Why
SPEC-029 called the page number "the first computed field" and SPEC-033 built
the table of contents on the same idea. Neither built the mechanism — each
hard-coded its own placeholder, its own cached value and its own update button.
A third one is where that becomes a problem, and there are eleven more field
types people actually ask for.

A field is one small idea: a placeholder that displays a computed value, keeps
the last computed result in the file so Word and Writer can show it without
recalculating, and updates when someone asks. This spec defines it once and
ships the set users reach for — the date in a letterhead that is right every
time you open the template, the filename in a footer, the client name pulled
from document properties, "see section 4.2" that still says 4.2 after you add a
section above it.

## Acceptance criteria
- [ ] "Insert → Field…" opens a dialog listing exactly the types in "Field catalogue" below, with per-type options and a live preview of the result it will insert
- [ ] "Insert → Date and time" offers the formats in "Date and time formats" below rendered in the document's language (SPEC-035), plus an "Update automatically" toggle — off inserts plain text, on inserts a `date` field
- [ ] A field is one caret atom — a run whose text is `FIELD_PLACEHOLDER_TEXT`, per "Document model" below: clicking anywhere in it selects the whole field, typing over it replaces it entirely, and one Backspace at its trailing edge deletes it whole
- [ ] Fields render with grey shading when the caret is inside one or when "View → Field shading" is on; the shading never prints or exports (see "Specs this changes")
- [ ] A selected field's context menu and toolbar chip offer "Update field", "Edit field…", "Toggle field codes", "Unlink field" (replace with its current result as plain text) and "Lock field"
- [ ] F9 updates the field at the caret and "Tools → Update all fields" updates every unlocked field in the document, including the SPEC-033 table of contents, which stops owning its own update path
- [ ] Fields update on exactly the occasions in "Recalculation" below, and opening a document is never one of them
- [ ] "Insert → Cross-reference" lists the document's headings (SPEC-019) and bookmarks (SPEC-025), inserts a reference showing the target's text, page number or above/below position, and clicking the inserted reference scrolls to its target
- [ ] A cross-reference whose target has been deleted renders as "Error! Reference source not found." and is listed by "Tools → Update all fields" rather than failing the update
- [ ] Fields work in body text, headers, footers, footnotes, table cells and shape text boxes, and a `page` field in a header shows the right value on every page
- [ ] Field results are inserted as text and never as markup, and no field type evaluates an expression — covered by an injection test using a document property whose value is markup, matching SPEC-041's untrusted-text rule
- [ ] A field type outside the catalogue keeps its code and its cached result byte-for-byte, displays the cached result, is shown read-only with "This field type isn't supported — its value is shown as it was saved", and is never recomputed
- [ ] Find (SPEC-017) matches the text a field displays; Replace does not edit inside a field, and reports the skipped matches instead of corrupting them
- [ ] Fields round-trip through `.docx` (`w:fldSimple`, or `w:fldChar`/`w:instrText` runs, with the result runs read into `cached`) and `.odt` (the `text:` elements named in the catalogue, with `text:user-field-decls` for custom properties)
- [ ] Field evaluation is deterministic under test: the evaluator reads the clock through the injectable source in "Determinism" below and never calls `Date.now()` or `new Date()` directly
- [ ] Opening the fixtures below renders one field of each catalogue type at its saved value, plus one unsupported field type, and a save with no edits leaves every cached result unchanged

## Out of scope
Captions, `SEQ` numbering, a table of figures and an index — that cluster needs
its own spec, and cross-references here target headings and bookmarks, which
already exist. Calculation fields (`= a + b`), `IF` and other conditional
fields. `INCLUDETEXT`, `LINK` and anything that reads another file.
`MACROBUTTON` and `ASK`/`FILLIN` prompts. Form fields and content controls.
Mail-merge fields — SPEC-046 owns those and they stay its own mechanism (see
"Specs this changes"). Editing the raw field code as text; codes can be viewed,
not hand-written. Nested fields. Chapter-based page numbering ("2-1"), still out
of scope from SPEC-029.

## Reference

### Document model — fields already exist
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly, **including fields —
they already exist.**
SPEC-029 built them and SPEC-031's note marks copy the convention. This spec
widens what is there; it does not introduce a parallel mechanism.

```ts
// src/lib/documentModel.ts, as it ships today
export interface Run {
  text: string                  // one character for a field run — see below
  formatting: RunFormatting
  field?: FieldSpec             // SPEC-029: renders a computed value
  note?: NoteRef                // SPEC-031: same convention
  styleId?: string
  changes?: RunTrackedChanges
}
export interface FieldSpec { kind: FieldKind; format: PageNumberFormat }
export const FIELD_PLACEHOLDER_TEXT = "1";
```

**A field is a run, not a node.** `Run.text` stays `FIELD_PLACEHOLDER_TEXT` —
one character — so the field is a single caret atom and `blockLength`,
`splitRuns`, selection, undo and both readers' offset bookkeeping are unchanged.
Every renderer substitutes the computed value at paint time instead of reading
the placeholder. That is the shipped design, and it is a better one than a
separate inline node: one Backspace removes a whole field, and no edit can
leave half a field behind.

There is no `Inline` union and no `LineBreak` node to add to; tabs and forced
line breaks are literal `"\t"` and `"\n"` in `Run.text`.

**`FieldSpec` widens into a discriminated union on `kind`.** This is the same
widening `specs/046-template-gallery-and-mail-merge.md` performs for merge
chips; the two must land on one union, not two:

```ts
export type FieldSpec =
  | { kind: "page" | "numpages"; format: PageNumberFormat }        // SPEC-029
  | { kind: "merge"; name: string }                                // SPEC-046
  | { kind: "if"; name: string } | { kind: "endif" }               // SPEC-046
  | { kind: "each"; name: string }                                 // SPEC-046
  // added by this spec:
  | { kind: "date" | "time"; format: DateTimeFormat; cached: string; locked?: true }
  | { kind: "author" | "docTitle" | "docSubject" | "docKeywords"
        | "fileName" | "wordCount"; cached: string; locked?: true }
  | { kind: "docProperty"; name: string; cached: string; locked?: true }
  | { kind: "ref" | "pageRef"; target: string; show: RefFormat
      ; cached: string; locked?: true }
  | { kind: "unsupported"; code: string; cached: string }
```

**`cached` is a string, not runs.** Formatting lives on the run that carries
the field, which is where every other run-level property already lives, so a
bolded page number is a bold run — the result needs no formatting of its own.

**`cached` is compared by `equalStructural`.** It is document content: it is
what Word and Writer display for a field they do not recalculate. Excluding it
would let a writer that emits an empty or stale result round-trip green. Two
consequences follow, and both are deliberate: field values must be
deterministic under test (see "Determinism"), and opening a document must not
update fields (see "Recalculation"), or every no-edit round trip would fail on
the date.

Page fields carry no `cached` because they are recomputed on every paint;
`format` stays required for them and appears on no other member.

**`unsupported` holds the preservation case.** Word defines around ninety field
types. Anything outside the catalogue is read into `{ kind: "unsupported",
code, cached }` — the instruction text verbatim and its last result — and is
written back from those two values unchanged. This is SPEC-008's rule applied
at run level, and it is why an unsupported field displays its value rather than
a placeholder box: we have what the file was saved with, and that is the truth.

Names here are never `id`: `src/lib/equality.ts`'s `IGNORED_KEYS` drops keys
called `id` from structural comparison, so `target` and `name` would be
invisible to a round-trip test if they were named that — the reason
`NoteRef.noteId` is spelled the way it is.

### Cross-reference targets
Bookmarks are defined in `specs/025-hyperlinks-and-bookmarks.md`: named
locations written as `w:bookmarkStart`/`w:bookmarkEnd` in `.docx` and
`text:bookmark` in `.odt`, with unique, space-free names validated on entry.
Headings are paragraphs carrying the Heading 1–6 styles from
`specs/019-paragraph-styles-and-style-picker.md`. A cross-reference stores its
target's bookmark name in `target`; referencing a heading inserts a
hidden bookmark at that heading first if it has none, which is both what Word
does and what lets `.odt`'s `text:bookmark-ref` carry the same reference.
Hidden bookmarks are not listed by SPEC-025's bookmark manager.

### Field catalogue
| Field | Shows | `.docx` code | `.odt` element |
| --- | --- | --- | --- |
| `date` | current date, formatted | `DATE \@ "…"` | `text:date` |
| `time` | current time, formatted | `TIME \@ "…"` | `text:time` |
| `author` | metadata author | `AUTHOR` | `text:author-name` |
| `docTitle` | metadata title | `TITLE` | `text:title` |
| `docSubject` | metadata subject | `SUBJECT` | `text:subject` |
| `docKeywords` | metadata keywords | `KEYWORDS` | `text:keywords` |
| `docProperty` | a custom property's value | `DOCPROPERTY "name"` | `text:user-field-get` |
| `fileName` | file name, path optional | `FILENAME` (`\p` for path) | `text:file-name` |
| `page` | this page's number | `PAGE` | `text:page-number` |
| `numpages` | total pages | `NUMPAGES` | `text:page-count` |
| `wordCount` | total words | `NUMWORDS` | `text:word-count` |
| `ref` | a target's text, number or position | `REF <bookmark> \h` | `text:bookmark-ref` |
| `pageRef` | the page a target is on | `PAGEREF <bookmark> \h` | `text:bookmark-ref`, `text:reference-format="page"` |

`author`, `docTitle`, `docSubject`, `docKeywords` and `docProperty` read
SPEC-041's metadata and custom properties; they add no storage of their own.
`wordCount` reads SPEC-036's counter. A cross-reference to a heading works by
inserting a hidden bookmark at that heading if one is not already there, which
is what Word does and what makes `.odt`'s `text:bookmark-ref` able to carry it.

### Date and time formats
A closed catalogue, not arbitrary picture strings. Each entry owns three
things: the pattern stored in `format`, the `.docx` picture string, and
the `.odt` `number:date-style` body written into `styles.xml`. Rendering uses
`Intl.DateTimeFormat` with the document's language from SPEC-035.

| `format` | Renders (en-GB, 4 March 2026) |
| --- | --- |
| `d MMMM yyyy` | 4 March 2026 |
| `d MMM yyyy` | 4 Mar 2026 |
| `dd/MM/yyyy` | 04/03/2026 |
| `MM/dd/yyyy` | 03/04/2026 |
| `yyyy-MM-dd` | 2026-03-04 |
| `EEEE, d MMMM yyyy` | Wednesday, 4 March 2026 |
| `HH:mm` | 14:30 |
| `h:mm a` | 2:30 PM |
| `HH:mm:ss` | 14:30:00 |

A picture string in an opened document that is not in this table is kept in
`format` verbatim, the field renders its `cached` value, and the field is
read-only with the same message unsupported types get. Implementing Word's
picture-string language is a spec of its own, and guessing at it would produce
dates that are wrong in a way nobody notices until a contract is signed.

### Recalculation
Four rules, exhaustive:

1. **Pagination-dependent fields** — `page`, `numpages`, `pageRef`, and
   `ref` with `show: "page"` — recompute whenever pagination changes.
   They have no stable cached value in a live editor; their `cached` string is
   written on save purely for other applications.
2. **Every other field** updates only when asked: on insert, on F9, on "Update
   all fields", or before print and PDF export when "Update fields before
   printing" is on. That setting defaults to **off**.
3. **Opening a document never updates any field.** A `Date` field in a contract
   saved last March shows last March until a person asks for today. This also
   keeps the no-edit round trip honest — see "Equality".
4. A field with `locked` set is skipped by rules 2 and 3 entirely, and by
   "Update all fields".

### Determinism
The evaluator reads the current instant through one injectable source —
`src/lib/clock.ts`, exporting `now(): Date` — and no field code calls
`Date.now()` or `new Date()`. Unit tests substitute a fixed instant; e2e tests
set it before loading the page.

This is not incidental tidiness. `cached` is compared by
`equalStructural`, so a `Date` field evaluated from the wall clock makes every
round-trip test that touches one fail whenever the suite runs across midnight —
the same shape of flakiness the scheduler tests already had, and harder to
reproduce.

### Specs this changes
- `specs/029-headers-footers-and-page-numbers.md` — **no behaviour change, and
  no migration.** Its `page` and `numpages` kinds keep their names, keep their
  required `format`, and become the first two members of the widened union. The
  insert UI, the format options and the per-section "start numbering at" are
  untouched. What this spec adds around them is the update path they never had:
  page fields are recomputed on every paint today, and rule 1 of
  "Recalculation" is a description of that, not a change to it.
- `specs/033-table-of-contents.md` — the TOC stays its own block, not a
  field run, because it holds block content. Its "Update table of
  contents" control remains, and "Update all fields" now drives the same code
  path. No user-visible change.
- `specs/034-print-and-export-pdf.md` — its list of screen-only affordances
  that must not print gains field shading. That criterion enumerates them, so
  the list changes with this spec rather than growing silently.
- `specs/046-template-gallery-and-mail-merge.md` — **one union, two specs.**
  SPEC-046 widens `FieldSpec` with `merge`, `if`, `endif` and `each`; this spec
  widens the same union with the kinds above, and whichever lands second adds
  its members rather than starting a second union. The behaviours stay
  separate: merge chips carry no `cached` value, never appear in
  "Insert → Field…", and are untouched by F9 and "Update all fields" — they are
  resolved once, by a merge run, against a record.
- `specs/041-document-properties-and-metadata.md` — no change. Custom
  properties gain a consumer, not new storage.

### Existing I/O surface
```ts
// src/lib/docx/reader.ts   parseDocx(data): Promise<ParsedDocxDocument>   — SPEC-004
// src/lib/odf/reader.ts    parseOdt(data):  Promise<ParsedOdtDocument>    — SPEC-005
// src/lib/docx/writer.ts   writeDocx(…): Promise<WriteResult>            — SPEC-006/008
// src/lib/odf/writer.ts    the .odt writer, same WriteResult             — SPEC-007/008
// src/lib/opaque.ts        WriteResult = { bytes: Uint8Array; dropped: DropRecord[] }
```
Per `CLAUDE.md`, field parsing, serialisation and evaluation live in `src/lib/`
as plain TypeScript with no Vue imports, unit-testable without mounting a
component. `.docx`'s two encodings both read into the same field run: the compact
`w:fldSimple`, and the run-split `w:fldChar begin` / `w:instrText` /
`w:fldChar separate` / result runs / `w:fldChar end` sequence. Writers emit
`w:fldSimple` for every field this spec creates, and re-emit an unsupported
field in whichever encoding it arrived in.

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic`, **which is not implemented**; this spec does not need it.
- **`equalStructural`** — for the same-format round trip. It passes without any
  clock trickery because rule 3 of "Recalculation" means opening and saving
  changes no `cached` value. A test that *does* update fields fixes the clock
  first.
- For cross-format saves, assert `equalStructural` between the source model
  masked by the writer's `dropped` records and the re-opened model. Masking is
  needed because the catalogues differ: an unsupported `.docx` field code has
  no `.odt` element to carry it, and produces a `DropRecord` naming the field
  and its location.

`equalBytes` is not used by this spec.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). Adds
`fixtures/docx/word-fields.docx` and `fixtures/odt/lo-fields.odt`: one field of
each catalogue type, one `Date` field in a header, one locked field, one
cross-reference to a heading and one to a bookmark, one cross-reference whose
target has been deleted, one field using a picture string outside the format
table, and one unsupported type (`MERGEFIELD` in the `.docx`, a
`text:conditional-text` in the `.odt`). The `.docx` includes at least one field
in the run-split `w:fldChar` encoding and one as `w:fldSimple`. Manifest entries
name each field's type, its arguments and the cached result that must survive a
no-edit round trip.
