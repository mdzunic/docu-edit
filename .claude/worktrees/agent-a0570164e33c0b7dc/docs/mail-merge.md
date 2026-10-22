# Mail merge (SPEC-046)

A guide for anyone writing a template — not the engine's internals (see
`src/lib/merge.ts` for that). It covers merge fields, conditional blocks and
repeating table rows: the syntax, what a condition's value actually is, and
which content a chip governs.

## Merge fields

Insert a merge field from **Insert → Merge field…**, or type it directly —
see [Literal text](#literal-text-is-never-a-chip) below. A merge field looks
like:

```
{{client_name}}
```

A name starts with a letter or underscore and holds only letters, digits,
underscores and dots. Dots traverse nested objects in a JSON data source:
`{{customer.name}}` reads the `name` property of the `customer` object.

Once inserted, a merge field is a single unit — clicking it places the caret
before or after it, never inside, and one Backspace removes the whole thing.
It is never partially edited the way plain text is.

## Attaching data and running a merge

**File → Mail merge…** attaches a CSV or JSON file, previews the first
record with its fields filled in, and reports which fields matched a column
(and which didn't) before you run anything. A CSV's first row is the header;
every other row is one record. A JSON file is a top-level array of objects,
one per record.

A name matches a column or property of the exact same name first; only if
there is no exact match does a dotted name (`customer.name`) traverse into a
nested object — so a CSV column literally named `customer.name` (unusual,
but legal) is matched as one name, not a path.

The merge produces either:

- **A `.zip`**, one document per record — pick this when each record's
  document is a separate file (letters, certificates, individual invoices).
- **A single document**, each record starting on a new page — pick this for
  one combined file (a printed batch, a compiled report).

## Conditional blocks

Wrap content in `{{#if field}}` … `{{/if}}` to show it only when `field` is
truthy for the current record:

```
{{#if paid}}
Thank you for your payment.
{{/if}}
```

### The truthiness table

| Source | False | Everything else is true |
| --- | --- | --- |
| JSON | the field is absent, or its value is `null`, `false`, `0`, `""`, `[]`, or `{}` | |
| CSV | the column is absent, the cell is empty, or its value (trimmed, case-insensitive) is `false`, `no`, or `0` | |

An empty array or object counts as false in JSON specifically so that
`{{#if line_items}}` guarding a repeating table row does not fire when the
list is empty — a document with an empty line-items array should not show a
leftover heading for a table that will not appear.

Every CSV cell is text, so `false`/`no`/`0` are spelled out explicitly —
without that rule, a cell containing the word `false` would otherwise read
as true (a non-empty string).

**A field the condition tests but the record does not have is false — and
is not listed as a missing field.** Testing for something is not the same
as requiring it; only a `{{field}}` that is actually shown counts as
missing when it isn't there.

### What a condition governs

Where you put the `{{#if}}`/`{{/if}}` chips changes what they remove:

- **On their own line** (nothing else in the paragraph — no other text, no
  other chip): the two marker paragraphs are always removed, and if the
  condition is false, everything between them is removed too. This is the
  shape to use for a block of one or more paragraphs, or a table: put
  `{{#if field}}` alone on its own line before the block and `{{/if}}` alone
  on its own line after it.
- **Inline, mixed with other text**: only the exact span from `{{#if}}` to
  `{{/if}}` is affected. If the condition is true, just the two chips are
  removed and the text between them stays inline. If it's false, that whole
  span is removed, joining what's left on either side — same as selecting
  that span and deleting it.

Conditions can nest to any depth; each `{{#if}}`/`{{/if}}` pair is
independent of the ones around it.

### What is not allowed

A conditional pair must open and close within the same table cell, and
within the same repeating row (see below) — a chip cannot open inside one
table cell and close in a different one, or in the row outside a repeating
row's cell. A conditional or repeating chip is also only valid in the main
body of the document, never in a header, footer, footnote or endnote.
**Mail merge checks every one of these rules before writing a single
document; if any template error is found, it is reported (with its
location) and the run does not produce output.** This is deliberate: a
half-applied conditional would produce a plausible-looking but wrong
document, silently, which a mail merge must never do.

A data problem is different from a template error and never blocks a run —
see [Data errors](#data-errors-never-fail-a-run).

## Repeating table rows

A table row repeats once per element of an array field. With the caret in
the row, use **Table → Repeat row for each…** (also on the row's right-click
menu) and pick the array field — this inserts an `{{#each field_name}}`
chip into the row's first cell. There is no closing chip: the row itself is
the boundary. Choosing the same command again on an already-repeating row
turns it off.

Only a JSON data source has arrays — a CSV column is always one plain text
value, so an array field can only come from an attached JSON file, or (if
none is attached yet) be typed by name so the template can be written
before the data exists.

### Scope inside a repeating row

Inside a repeating row's cells, a `{{field}}` chip checks the current array
element first, and falls back to the whole record if the element doesn't
have that name. That's what lets a line-item row mix `{{sku}}` (only on the
element) with `{{invoice_no}}` (only on the record) with no special syntax:

```json
{ "invoice_no": "INV-100", "line_items": [{ "sku": "A1", "qty": 3 }] }
```

```
| {{sku}} | {{qty}} | {{invoice_no}} |
```

produces `A1 | 3 | INV-100` for that row.

A header row (marked as a table header) never repeats, whether or not it
carries any chips — only ordinary rows do.

### What is not allowed

Two `{{#each}}` chips in the same row is a template error (a row can only
repeat over one array) — it is reported and blocks the run, the same as a
conditional-block error above.

### Data errors never fail a run

If `{{#each}}` names a field that is missing from the record, or that isn't
an array, that's a **data** problem, not a template problem: the row is
still produced once (not repeated), its fields inside it are marked as
described below, and it's listed in the merge report. The rest of the run
continues normally.

## Missing fields

If a shown `{{field}}` has no value in the current record — the column or
property is absent, or its value is `null`/absent in JSON — the merge
leaves a clearly marked, highlighted gap (`[field_name]`) in its place
rather than silent blank text, and lists it in the merge report by record
number. This never stops the run; every other field and every other record
still merges normally.

## Data is never treated as a template

Whatever a record's value contains — even something that looks like merge
syntax, e.g. a cell whose text is literally `{{#if x}}` — is inserted as
plain text. It is never interpreted as a chip, a condition, or anything
else the merge engine understands. A resolved merge document holds no merge
chips: every field that had a value now holds that value as ordinary text.

## Literal text is never a chip

Typing `{{client_name}}` — or opening a document someone else wrote that
happens to contain that text — produces ordinary text, not a chip. Nothing
converts it automatically; a document that legitimately contains `{{` (a
code sample, a snippet about templates) must never be silently rewritten on
open.

When you do want to turn typed placeholders into real chips, use **Tools →
Convert {{…}} text to merge fields**. It scans the whole document — body,
tables, headers, footers, footnotes and endnotes — and converts every
well-formed match (`{{name}}`, `{{#if name}}`, `{{/if}}`, `{{#each name}}`)
in one step, so it's a single undo if you didn't want it. Text that isn't
quite right (a space in the name, an unmatched brace) is left alone rather
than guessed at.
