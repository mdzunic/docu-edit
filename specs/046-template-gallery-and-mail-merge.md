---
id: SPEC-046
title: Template gallery and mail merge
status: draft
area: export
depends_on: [SPEC-045, SPEC-022, SPEC-029]
---
## Why
Most new documents are not blank — they are last quarter's report with the
numbers changed. Templates plus merge fields turn the generation engine into
something a business user can drive without an engineer.

## Acceptance criteria
- [ ] The library's "New document" flow offers a template gallery with previews, alongside "Blank document"
- [ ] The gallery ships with at least six templates (letter, memo, report, invoice, meeting minutes, CV) that render correctly and are editable like any document
- [ ] "Save as template" stores the current document as a user template, and user templates appear in the gallery with rename and delete
- [ ] `.dotx` and `.ott` files can be opened as templates: opening one creates a new untitled document from it rather than editing the template
- [ ] Merge fields can be inserted at the caret from a named field list, rendering as a chip like `{{client_name}}` — a field run and a single caret atom, per "Merge chips in the model" below; the list is populated by "The named field list" below, so a template can be written before its data exists
- [ ] Merge fields survive save and re-open in both formats, serialised exactly as "Chip serialisation" below fixes it, asserted with `equalStructural` over a document holding one chip of every kind
- [ ] Literal `{{…}}` text is never turned into a chip implicitly — see "Literal text is not a chip"
- [ ] A merge run accepts a CSV or JSON data source, previews the first record with fields filled, and reports column-to-field mapping including unmapped fields; names resolve by "Data sources and field resolution" below
- [ ] Conditional blocks are authored as paired `{{#if field}}` / `{{/if}}` chips; a condition's value is the truthiness table in "Conditional blocks" below, and a false condition removes exactly the range that section defines
- [ ] A table row repeats over an array field when it carries an `{{#each field}}` chip, inserted by "Table → Repeat row for each…"; copies resolve per "Repeating table rows" below
- [ ] Conditional blocks and repeating rows are documented for template authors, including the truthiness table and the scope rules
- [ ] Template errors — unpaired, crossed or illegally nested chips — are listed with their locations and block the run before any document is produced; data errors never fail a run
- [ ] Running the merge produces one document per record, delivered either as a single `.zip` holding one entry per record, named per "Merge output" below, or as a single document with each record starting on a new page
- [ ] A missing field in a record leaves a clearly marked gap and is listed in a merge report rather than failing the whole run
- [ ] Merge field syntax in the data is never executed or evaluated — data is inserted as text, covered by an injection test

## Out of scope
Sending merged documents by email. Connecting to a database or CRM as a data
source. Per-record PDF conversion (compose with SPEC-034). Negation,
comparison or boolean operators in a condition (`{{#if !x}}`, `{{#if a == b}}`)
— `{{#if field}}` tests one field's truthiness and nothing else. Nested
`{{#each}}` and parent-scope escapes (`../`). Arrays from a CSV source.
Conditional and repeat chips outside the document body. Hand-editing the
column-to-field mapping — mapping is by name, and reported, not remapped.

## Reference

### Specs this depends on
Intake reads one file at a time, so the contracts this spec leans on are
restated here.

- `specs/002-create-blank-document.md` — the canonical document model, its
  units and the equality relations. A merge chip is an inline atom in that
  model; in the current implementation that is a `Run` carrying a `field`
  (`src/lib/documentModel.ts:181`). **Read it before starting.**
- `specs/029-headers-footers-and-page-numbers.md` — the field run this spec
  extends: `Run.field: FieldSpec`, a one-character `FIELD_PLACEHOLDER_TEXT`
  occupying exactly one caret position, with the displayed value substituted at
  paint time. Merge chips add kinds to that union and change nothing else about
  it.
- `specs/022-tables-insert-and-edit.md` — `Table.rows: TableRow[]`, every row
  holding exactly `columnWidthsCm.length` cells, merged cells present as
  `covered: true` placeholders. Row repetition below copies whole `TableRow`s
  and never disturbs that invariant.
- `specs/045-document-generation-api.md` — the generation path a merge run
  emits through; this spec adds no serialisation code of its own.
- `specs/008-round-trip-preservation.md` — content that was not parsed is
  preserved verbatim. A chip form neither reader recognises stays an opaque
  inline rather than becoming text.
- `specs/034-print-and-export-pdf.md` — per-record PDF is composed with it, not
  built here.

**Equality.** The round-trip assertion — insert every chip kind, save, re-open,
compare — uses `equalStructural` from `src/lib/equality.ts`, in both formats.
The `.zip` delivery is asserted on its entry list and on each entry re-opened
and compared with `equalStructural` against the model the merge produced, never
with `equalBytes`: a ZIP's stored order is fixed by the writer but its
compressed bytes are not a contract.

### Boundary with SPEC-053
`specs/053-fields-date-time-and-cross-references.md` builds the general field
catalogue (date, filename, document property, cross-reference) on the same
`Run.field` union and states that mail-merge fields stay this spec's own
mechanism. The split, from this side:

- SPEC-046 lands first and adds the four kinds below to `FieldSpec`; SPEC-053
  adds its catalogue kinds to the same union rather than introducing a second
  one.
- Merge chips never appear in SPEC-053's "Insert → Field…" catalogue, are not
  touched by "Update all fields" or F9, and are not unlinkable — they have no
  computed value to refresh, only a record to be resolved against, and that
  happens once, during a merge run.
- SPEC-053's rule for a field type it does not recognise (keep the code and the
  cached result, show it read-only) is what makes a template authored here safe
  to open after SPEC-053 ships, and vice versa.

### Merge chips in the model
A merge chip is a **field run** — the mechanism SPEC-029's page fields already
use (`Run.field`, `src/lib/documentModel.ts:181`). Not literal text, and not a
new inline kind. `FieldSpec` widens into a discriminated union on `kind`:

```ts
export type FieldSpec =
  | { kind: "page" | "numpages"; format: PageNumberFormat }
  | { kind: "merge"; name: string }   // {{client_name}}
  | { kind: "if"; name: string }      // {{#if paid}}
  | { kind: "endif" }                 // {{/if}}
  | { kind: "each"; name: string };   // {{#each line_items}}
```

Every chip keeps `Run.text === FIELD_PLACEHOLDER_TEXT` — one character — so
caret arithmetic, `blockLength`, `splitRuns`, selection, undo and both readers'
offset bookkeeping are unchanged. The `{{client_name}}` the user sees is
derived from the spec at paint time, at the same substitution point where a
page field's live number is produced (`src/views/document/usePaginationView.ts:389`).

That makes a chip atomic: one Backspace removes the whole field, and no edit
can leave a half-written delimiter behind. A literal-text encoding cannot rule
that out — `{{#if paid}}` typed as text can be split across runs by a
formatting change, or half-deleted into `{{#if pai`, and the merge engine would
then have to decide what a broken delimiter means. It never has to.

`name` holds the field or array name, never `id`: `lib/equality.ts`'s
`IGNORED_KEYS` drops bare `id` keys from structural comparison, which would let
a chip pointing at the wrong field pass a round-trip test (the same reasoning
as `BookmarkAnchor`).

Resolution happens on the model, before the writers: **no merged output
document contains a merge chip.** A resolved value becomes `Run.text` on an
ordinary run, so `writeDocx`/`writeOdt` need no merge-aware path, and the AC
"data is inserted as text" holds structurally — there is no stage at which a
record's value is re-read as markup.

### Chip serialisation
One grammar, used for the chip's display label and for its on-disk form in both
formats:

| Kind | Canonical text |
| --- | --- |
| `merge` | `{{client_name}}` |
| `if` | `{{#if paid}}` |
| `endif` | `{{/if}}` |
| `each` | `{{#each line_items}}` |

A name matches `[A-Za-z_][A-Za-z0-9_.]*` (dots traverse objects — see
resolution below). Anything else is refused at insertion, and on read is left
as whatever it was rather than becoming a chip.

**.docx** — a `w:fldSimple` whose instruction is `MERGEFIELD` followed by the
canonical text in double quotes: `MERGEFIELD "{{client_name}}"`. Always
quoted, so the writer has no has-a-space branch and the reader has one pattern.
The cached run `w:fldSimple` wraps (`src/lib/docx/writer.ts:335`) carries the
canonical text rather than the model's one-character placeholder, so a consumer
that does not evaluate fields still shows `{{client_name}}`; the reader rebuilds
from the instruction and ignores the cached text, exactly as it does for `PAGE`.
Word round-trips an unresolvable `MERGEFIELD` untouched, which is what buys the
"survives save and re-open" AC for templates that make a detour through Word.

**.odt** — `<text:placeholder text:placeholder-type="text">{{client_name}}</text:placeholder>`,
extending `fieldContentXml` (`src/lib/odf/writer.ts:1050`). Chosen over
`text:user-field-get`, which requires a matching declaration in
`text:user-field-decls` — a second document-level list to keep in sync on every
insert and delete, holding a value a template has no meaning for — and over
`text:database-display`, which names a data source LibreOffice would try to
resolve. A placeholder carries its own text, needs no declarations, and
LibreOffice renders it as a grey placeholder: visually a chip, which is what it
is. Its content is the same string the .docx instruction quotes, so both
formats share one parser.

A chip form the reader does not recognise **stays what it was**: an
unrecognised `w:fldSimple` already falls through to the opaque-inline path
(`src/lib/docx/reader.ts:513`), and an unrecognised `text:placeholder` must do
the same, so SPEC-008's preserve-what-you-did-not-parse rule still holds.

### Literal text is not a chip
Typing `{{client_name}}` in the editor produces text. Opening a `.docx` whose
author typed `{{client_name}}` produces text. Conversion is one explicit,
undoable command — "Convert `{{…}}` text to merge fields" — offered when a
document is opened as a template and literal candidates are found; it scans the
body, tables, headers, footers and notes and converts every match in a single
undo step.

Implicit conversion at open is refused because it rewrites content the user did
not ask to change, and would make open → save non-lossless for any document
that legitimately contains `{{` — a code sample, a Handlebars snippet, a
template about templates.

### Conditional blocks
Delimiters are the `if` and `endif` chips above. Pairing is a stack; nesting to
any depth is allowed. Which range a pair governs depends on where its chips
sit:

**Block form** — the `{{#if}}` chip is alone in its paragraph (no other chip,
no non-whitespace text) and so is its `{{/if}}`. The region is the paragraphs
strictly between them, and the two delimiter paragraphs are removed from the
output **whatever the condition's value**: true keeps the paragraphs between
them, false removes those too. Without the carve-out, the commonest authoring
shape — markers on their own lines — would leave two empty paragraphs, i.e. two
visible blank lines, in every merged document.

**Inline form** — anything else. The region is the model range from the
`{{#if}}` chip to the `{{/if}}` chip inclusive, removed exactly as the editor's
own range deletion removes a selection spanning those two positions
(`src/lib/textEditing.ts`), paragraph join included when the chips sit in
different paragraphs. Deliberately the editor's operation and not a second one:
deleting a range across paragraphs, list items, bookmarks, comment anchors and
tracked changes is already specified and tested there, and a parallel
implementation would drift from it.

When the condition is true, only the chips themselves are removed (block form:
their paragraphs).

Truthiness of the named field's value in the current record:

| Source | False | True |
| --- | --- | --- |
| JSON | absent, `null`, `false`, `0`, `""`, `[]`, `{}` | anything else |
| CSV | absent column, empty cell, or `false` / `no` / `0` (trimmed, case-insensitive) | anything else |

`[]` and `{}` are false because a `{{#if line_items}}` guarding a repeating row
must not fire on an empty array, and plain JS truthiness gets that backwards.
The CSV word list exists because every CSV value is a string, so a cell
containing `false` would otherwise be true.

A field a condition tests and the record does not have is **false, and not a
merge-report entry**. A condition asks whether something is there; absence is
an answer, not a fault. This is the only place a missing field goes unreported
— everywhere else the missing-field AC applies.

Template errors, which block the run: an `{{/if}}` with no open `{{#if}}`; an
unclosed `{{#if}}`; a pair that starts inside a table cell and ends outside it,
or starts inside a repeating row and ends outside it; any `if`/`endif`/`each`
chip outside the document body (a header, footer or note — the region rules are
defined over the body's block list, and a header belongs to pages, not
records). These stop the run before a single document is produced, because a
half-applied conditional yields plausible-looking wrong documents, which is the
one failure a mail merge must never ship. Data errors never stop a run.

### Repeating table rows
A row repeats when one of its cells carries an `{{#each name}}` chip. There is
no closing chip: `TableRow` is already the boundary, so a `{{/each}}` would be
redundant and would raise a question with no good answer (which cell? which
paragraph?).

Authoring: with the caret in a table row, **Table → Repeat row for each…**
(also in the table context menu) offers the field list described under "The
named field list" below, narrowed to array fields, and inserts the chip at the
start of the row's first cell. On a row that already
repeats, the item removes the chip. The chip is the entire state — there is no
`TableRow` attribute — so nothing can go stale and there is nothing extra to
serialise.

Why not a `TableRow.repeat` field: `TableRow` has no round-trip channel for a
new modelled attribute. `docxTrPr` is verbatim source XML
(`src/lib/documentModel.ts:918`) and OOXML has no row property meaning "repeat
over data", so the attribute would have to travel as a custom-namespace element
that Word discards on its own save — and the AC requires merge markup to
survive a save and re-open. A chip rides the same `MERGEFIELD` /
`text:placeholder` path as every other chip, so row repetition round-trips for
free.

Expansion:

- The row is replaced by one copy per array element, in array order, each copy
  identical to the template row with the `each` chip removed.
- Inside a copy, a `{{name}}` chip resolves against the current element first
  and falls back to the record root, so `{{sku}}` (element) and `{{invoice_no}}`
  (record) both work in a line-item row without a scope-escape syntax.
- An empty array removes the row. If that leaves the table with no rows, the
  table block is removed — an invoice with no line items should print no
  line-item row and no empty grid. A header row does not carry the chip, so it
  is not part of the repeat and survives.
- A value that is missing or is not an array is a **data** error: the row is
  emitted once, its fields marked as gaps per the missing-field AC, with one
  merge-report entry. The run continues.
- Two `each` chips in one row, or an `each` row inside a row that already
  repeats (a nested table), are template errors: nested scopes need a
  parent-scope syntax, which is out of scope.

### Data sources and field resolution
**JSON** — a top-level array of objects, one per record. Dots traverse nested
objects (`{{customer.name}}`); nested arrays are what `{{#each}}` drives.
Values are inserted as `String(value)`; `null` and absent are a missing field.

**CSV** — the first row is the header, one record per row after it, every value
a string. A CSV source has no arrays, so `{{#each}}` over one is always the
not-an-array data error above.

Resolution order for a name: an exact column or property of that name first,
then dotted traversal. Exact-first matters for CSV, where a column literally
named `customer.name` is a name, not a path.

Matching is exact and case-sensitive. A column whose header is not a legal
field name — `Client Name`, with its space — can never match a chip and is
listed as unmapped in the mapping report, which is the signal to rename the
column. Unmatched chips are listed in the same report.

**The named field list.** Templates are usually written before the data
arrives, so the insertion list is not the data source's columns alone. It is,
in order: the columns or properties of the data source if one is attached to
the document; then every name already used by a chip in this document; then a
free-text entry validated against the name grammar. Array fields — the only
ones "Repeat row for each…" offers — can therefore only be listed from an
attached JSON source; without one, the array name is typed. Attaching a source
later never rewrites chips: it only changes what the mapping report says about
them.

### Merge output
**Separate files means one `.zip`.** Built with the existing `zip()`
(`src/lib/docx/zipWriter.ts:124`) — the same writer behind every `.docx`/`.odt`
save, so no new dependency — and handed to the browser by the transient-anchor
and object-URL download the editor already uses
(`src/views/DocumentView.vue:4453`).

Not N sequential downloads: Chrome prompts for "download multiple files" and
then rate-limits, Safari blocks all but the first, and 200 records would mean
200 user decisions. Not the File System Access API: absent in Firefox and
Safari, and it needs a directory grant per run.

- Archive name: `<document title>-merge.zip`, the title sanitised by the rule
  `docxFileName` already uses (`src/lib/docx/writer.ts:139`).
- Entry names: `NNN-<slug>.<ext>` — `NNN` the 1-based record number zero-padded
  to the width of the record count, `<slug>` from an optional **name field**
  chosen in the merge dialog from the source's scalar columns, `<ext>` the one
  output format chosen for the run. With no name field: `NNN.<ext>`.
- The index prefix makes entry names unique by construction, so two records
  sharing a name-field value cannot collide and there is no de-duplication rule
  to get wrong.
- Slug: keep `[A-Za-z0-9._-]`, collapse every other run of characters to a
  single `-`, trim leading and trailing `-`, truncate to 60 characters. A `/`
  would create a directory inside the archive, and Windows rejects `<>:"|?*`.
- Records are written in data-source order, so the same input produces the same
  archive — which is what makes an e2e assertion on its entry list possible.
- The archive holds exactly one document per record and nothing else. The merge
  report stays in the UI (the `ConversionReportDialog` pattern,
  `src/components/ConversionReportDialog.vue`): a stray report file would
  surprise anyone unzipping into a documents folder.

**Single document** — the records' resolved bodies are concatenated in source
order with a page break between consecutive records, keeping the template's
section setup, headers and footers. Page-number fields therefore number
continuously across the whole file: they are page fields, not per-record
fields, and restarting them per record would need a section break per record,
which is SPEC-028's feature and not implied by this AC.

Both forms build each record through the same resolve step and the same
writers; only the packaging differs. The run yields between records through
`scheduleIdleWork` (`src/lib/idleScheduler.ts:23`), as the other long passes
do, so a large merge does not freeze the editor.
