---
id: SPEC-000
title: Product roadmap — browser word processor for .docx and .odt
status: draft
area: product
---
## Why
This file is the index for the spec set. It is not a work order — no engineer
should implement it. It exists so a business person can see the intended order
of the queue and pick which spec to label `status:ready` next.

The product: a browser-based word processor that opens `.docx` (OOXML) and
`.odt` (ODF) documents, edits them with the feature depth people expect from
Google Docs and LibreOffice Writer, saves them back in the same format without
destroying what it did not understand, and generates new documents in both
formats from scratch.

Specs are ordered so each one is shippable on its own and builds on the ones
before it. Later specs assume earlier ones are merged (`depends_on` in each
file's frontmatter).

## Acceptance criteria
- [ ] No implementation. This spec is never labelled `status:ready`.

## Out of scope
Everything. See the individual specs.

---

## Spec conventions
Intake sees **one spec file at a time**, so a spec that leans on a shared
concept must carry that concept with it. Two rules, both learned from intake
rejecting SPEC-007 as ambiguous:

1. **One normative source per shared concept.** The canonical document model,
   its units, and the equality relations are defined in
   `specs/002-create-blank-document.md` and, once merged, in
   `docs/document-model.md`. The I/O API surface is defined in
   `specs/004-open-docx-text-and-paragraphs.md` (readers) and
   `specs/006-save-as-docx.md` (writers). The shared fixture set and its
   `manifest.json` are defined in SPEC-004. Nothing redefines these.

2. **Every dependent spec ends with a `## Reference` section** that names the
   dependency spec files by path, restates the contract it relies on inline
   (model shape, function signatures, fixture layout), and names the equality
   relation each test uses. A reader of one file alone must never have to guess.

Any spec asserting a round trip must name one of the three relations —
`equalBytes`, `equalStructural`, `equalSemantic` — and use the helpers in
`src/model/equality.ts`. "Asserts model equality" is not a specification.

---

### Phase 1 — Foundation and file I/O
| Spec | Title |
| --- | --- |
| SPEC-001 | App shell and document workspace |
| SPEC-002 | Create a new blank document |
| SPEC-003 | Local document library |
| SPEC-004 | Open .docx — text and paragraphs |
| SPEC-005 | Open .odt — text and paragraphs |
| SPEC-006 | Save as .docx |
| SPEC-007 | Save as .odt |
| SPEC-008 | Round-trip preservation of unsupported content |
| SPEC-009 | Drag-and-drop and open-from-URL |
| SPEC-010 | Autosave and crash recovery — queue after SPEC-011 |

### Phase 2 — Core editing
| Spec | Title |
| --- | --- |
| SPEC-011 | Rich text editing core |
| SPEC-012 | Character formatting — bold, italic, underline, strikethrough |
| SPEC-013 | Font family, size, text colour and highlight |
| SPEC-014 | Paragraph alignment, indents and spacing |
| SPEC-015 | Undo and redo |
| SPEC-016 | Cut, copy and paste |
| SPEC-017 | Find and replace |
| SPEC-018 | Bulleted and numbered lists |
| SPEC-019 | Paragraph styles and the style picker |
| SPEC-020 | Clear formatting and format painter |
| SPEC-021 | Keyboard shortcuts and shortcut reference |

### Phase 3 — Objects and structure
| Spec | Title |
| --- | --- |
| SPEC-022 | Tables — insert and edit |
| SPEC-023 | Table formatting and cell operations |
| SPEC-024 | Images — insert, resize and wrap |
| SPEC-025 | Hyperlinks and bookmarks |
| SPEC-026 | Special characters and symbol picker |

### Phase 4 — Page layout and long documents
| Spec | Title |
| --- | --- |
| SPEC-027 | Page setup — size, margins, orientation |
| SPEC-028 | Page and section breaks |
| SPEC-029 | Headers, footers and page numbers |
| SPEC-030 | Multi-column layout |
| SPEC-031 | Footnotes and endnotes |
| SPEC-032 | Multilevel lists and numbering control |
| SPEC-033 | Table of contents |
| SPEC-034 | Print and export to PDF |

### Phase 5 — Authoring depth, review and collaboration
| Spec | Title |
| --- | --- |
| SPEC-035 | Spell check with user-supplied dictionaries, and document language |
| SPEC-036 | Word count and document statistics |
| SPEC-037 | Document outline navigator |
| SPEC-038 | Comments |
| SPEC-049 | Menu bar and toolbar consolidation — queue before SPEC-039 |
| SPEC-039 | Track changes |
| SPEC-040 | Version history and restore |
| SPEC-041 | Document properties and metadata |
| SPEC-042 | Custom styles manager |
| SPEC-043 | Accessibility checker |
| SPEC-044 | Real-time collaborative editing |

### Phase 6 — Generation and scale
| Spec | Title |
| --- | --- |
| SPEC-045 | Document generation API |
| SPEC-046 | Template gallery and mail merge |
| SPEC-047 | Large document performance |
| SPEC-048 | Font fallback and embedding |
