---
id: SPEC-007
title: Save as .odt
status: draft
area: io
depends_on: [SPEC-005, SPEC-006]
---
## Why
Symmetry with SPEC-006, and the point at which the product becomes a converter:
open a `.docx`, save it as `.odt`, and the content survives.

## Acceptance criteria
- [ ] The "Download" menu offers "OpenDocument Text (.odt)" and downloads `<title>.odt`
- [ ] The written package contains an uncompressed `mimetype` entry stored first, plus `content.xml`, `styles.xml`, `meta.xml`, `settings.xml` and `META-INF/manifest.xml`
- [ ] The `mimetype` entry contains exactly `application/vnd.oasis.opendocument.text` and is written with ZIP STORE (no deflate), as ODF requires
- [ ] All document model content survives a save: text, run and paragraph formatting, paragraph styles, page layout
- [ ] Run and paragraph formatting is emitted as automatic styles referenced by name, not as invalid inline attributes
- [ ] `meta.xml` carries the title and a modification timestamp
- [ ] Files produced by the writer open in LibreOffice Writer without warnings — verified by a documented manual check recorded in the PR
- [ ] An automated test round-trips each `.odt` fixture (open → save → re-open) and asserts model equality
- [ ] Cross-format tests assert that `.docx` → save-as-`.odt` → open preserves text, headings and character formatting, and the same in the reverse direction

## Out of scope
Flat ODF (`.fodt`). Digital signatures. Encryption.

## Decisions
If the document has no title, go with default Document 1.
