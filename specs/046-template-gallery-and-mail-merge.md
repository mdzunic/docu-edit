---
id: SPEC-046
title: Template gallery and mail merge
status: draft
area: export
depends_on: [SPEC-045, SPEC-022]
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
- [ ] Merge fields can be inserted at the caret from a named field list, rendering as a chip like `{{client_name}}`
- [ ] A merge run accepts a CSV or JSON data source, previews the first record with fields filled, and reports column-to-field mapping including unmapped fields
- [ ] Running the merge produces one document per record, downloadable as separate files or as a single document with each record starting on a new page
- [ ] Conditional blocks (`{{#if}}` / `{{/if}}`) and repeating table rows driven by an array field are supported and documented
- [ ] A missing field in a record leaves a clearly marked gap and is listed in a merge report rather than failing the whole run
- [ ] Merge field syntax in the data is never executed or evaluated — data is inserted as text, covered by an injection test
- [ ] Merge fields survive save and re-open in both formats

## Out of scope
Sending merged documents by email. Connecting to a database or CRM as a data
source. Per-record PDF conversion (compose with SPEC-034).
