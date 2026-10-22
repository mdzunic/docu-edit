---
id: SPEC-032
title: Multilevel lists and numbering control
status: draft
area: formatting
depends_on: [SPEC-018, SPEC-019]
---
## Why
Contracts, standards and specifications number their headings (1., 1.1, 1.1.1),
and SPEC-018 deliberately stopped short of it. This is the numbering machinery
those documents depend on and the most common reason a `.docx` re-saved by a
web editor looks broken.

## Acceptance criteria
- [ ] A "Multilevel list" toolbar control offers gallery presets: 1./1.1/1.1.1, I./A./1., legal (1.1.1), and bullet-mixed
- [ ] "Define new multilevel list…" lets the user set, per level: number format, number style, prefix/suffix text, level to include ("1.1" vs "1"), start value, alignment, indent and text position
- [ ] A list definition can be linked to paragraph styles, so applying Heading 1–3 numbers them automatically
- [ ] "Restart numbering at 1" and "Continue previous numbering" are available on any list paragraph via the context menu
- [ ] "Set numbering value…" sets an explicit start number for a list item
- [ ] Changing a level's format updates every item at that level in the same list, and not items in other lists
- [ ] Custom tab stops are supported (left, right, centre, decimal, with leader characters), settable via the ruler and a tab dialog, since numbered lists depend on them for alignment
- [ ] Definitions round-trip through `.docx` as distinct `w:abstractNum` entries with per-level `w:lvlText` and `w:numFmt`, and through `.odt` as `text:list-level-style-number` entries
- [ ] Opening a fixture with two independent numbered lists and one restarted list preserves all three numbering sequences on save and re-open
- [ ] A numbered-heading document keeps its numbering when a heading is deleted, moved or demoted

## Out of scope
Legal-style cross-references to numbered items ("see clause 4.2"). Numbering of
captions, figures and tables. Bullet images.
